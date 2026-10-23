"""Bounded calls that can block in hidden, owned Excel COM instances."""
import threading

OPEN_TIMEOUT_SECONDS = 60


def owned_excel_pid(excel):
    import win32process
    return int(win32process.GetWindowThreadProcessId(int(excel.Hwnd))[1])


class OwnedExcelProcess:
    """Hold a kernel handle to one process instance so PID reuse cannot retarget cleanup."""

    def __init__(self, pid, handle):
        self.pid = int(pid)
        self.native_handle = handle
        self._closed = False

    def wait(self, timeout_ms):
        if self._closed:
            raise RuntimeError('Owned Excel process handle is closed.')
        import win32event
        result = win32event.WaitForSingleObject(self.native_handle, timeout_ms)
        if result == win32event.WAIT_OBJECT_0:
            return True
        if result == win32event.WAIT_TIMEOUT:
            return False
        raise RuntimeError('Could not confirm owned Excel process exit.')

    def terminate(self, timeout_ms=5000):
        if self._closed:
            raise RuntimeError('Owned Excel process handle is closed.')
        import win32api
        win32api.TerminateProcess(self.native_handle, 1)
        if not self.wait(timeout_ms):
            raise TimeoutError('Timed out waiting for the owned Excel process to exit.')
        return True

    def close(self):
        if not self._closed:
            import win32api
            try:
                win32api.CloseHandle(self.native_handle)
            finally:
                self._closed = True


def open_owned_excel_process(pid, *, hwnd=None):
    if isinstance(pid, bool) or not isinstance(pid, int) or pid <= 0:
        raise ValueError('A positive owned Excel process ID is required.')
    import win32api
    import win32con
    access = win32con.PROCESS_TERMINATE | win32con.SYNCHRONIZE
    handle = win32api.OpenProcess(access, False, pid)
    process = OwnedExcelProcess(pid, handle)
    if hwnd is not None:
        try:
            import win32process
            window_pid = int(win32process.GetWindowThreadProcessId(int(hwnd))[1])
            if window_pid != pid:
                raise RuntimeError('Excel window no longer belongs to the captured process.')
        except Exception as exc:
            process.close()
            if isinstance(exc, RuntimeError):
                raise
            raise RuntimeError('Cannot confirm that the Excel window still belongs to the captured process.') from exc
    return process


def terminate_owned_excel(pid):
    if isinstance(pid, OwnedExcelProcess):
        return pid.terminate()
    import win32api
    import win32con
    import win32event
    access = win32con.PROCESS_TERMINATE | win32con.SYNCHRONIZE
    handle = win32api.OpenProcess(access, False, pid)
    try:
        win32api.TerminateProcess(handle, 1)
        result = win32event.WaitForSingleObject(handle, 5000)
        if result != win32event.WAIT_OBJECT_0:
            raise TimeoutError('Timed out waiting for the owned Excel process to exit.')
        return True
    finally:
        win32api.CloseHandle(handle)


def wait_for_owned_excel_exit(pid, timeout_ms=5000):
    """Confirm exit of the exact Excel PID captured from our DispatchEx instance."""
    if isinstance(pid, OwnedExcelProcess):
        return pid.wait(timeout_ms)
    if isinstance(pid, bool) or not isinstance(pid, int) or pid <= 0:
        raise ValueError('A positive owned Excel process ID is required.')
    if isinstance(timeout_ms, bool) or not isinstance(timeout_ms, int) or timeout_ms < 0:
        raise ValueError('Process exit timeout must be a non-negative integer.')
    import win32api
    import win32con
    import win32event
    try:
        handle=win32api.OpenProcess(win32con.SYNCHRONIZE,False,pid)
    except Exception as exc:
        # ERROR_INVALID_PARAMETER is returned when the exact PID no longer exists.
        # Access denied and other failures are unknown, never proof of exit.
        code=getattr(exc,'winerror',None)
        if code is None and getattr(exc,'args',None):
            code=exc.args[0] if isinstance(exc.args[0],int) else None
        if code==87:
            return True
        raise
    try:
        result=win32event.WaitForSingleObject(handle,timeout_ms)
        if result==win32event.WAIT_OBJECT_0:
            return True
        if result==win32event.WAIT_TIMEOUT:
            return False
        raise RuntimeError('Could not confirm owned Excel process exit.')
    finally:
        win32api.CloseHandle(handle)


def quit_owned_excel(excel, *, quit_call=None, after_quit=None, timeout_ms=15000):
    """Quit and reap only the DispatchEx process identified by this Excel proxy.

    Excel.Application.Quit can return while its out-of-process server remains
    alive because a COM proxy is still held. Keep one process handle across the
    graceful wait and any fallback termination so PID reuse cannot redirect
    cleanup to another Excel instance.

    Returns True when termination was needed, False when Quit exited normally.
    """
    if isinstance(timeout_ms, bool) or not isinstance(timeout_ms, int) or timeout_ms < 0:
        raise ValueError('Process exit timeout must be a non-negative integer.')
    hwnd = int(excel.Hwnd)
    pid = owned_excel_pid(excel)
    process = open_owned_excel_process(pid, hwnd=hwnd)
    quit_error = None
    try:
        try:
            (quit_call if quit_call is not None else excel.Quit)()
        except Exception as exc:
            quit_error = exc
        # Release the Python proxy and the caller's COM apartment before waiting.
        # Excel may not terminate while the automation client still holds either.
        quit_call = None
        excel = None
        if after_quit is not None:
            try:
                after_quit()
            except Exception as exc:
                if quit_error is None:
                    quit_error = exc
        if process.wait(timeout_ms):
            if quit_error is not None:
                raise quit_error
            return False
        process.terminate(5000)
        if quit_error is not None:
            raise quit_error
        return True
    finally:
        process.close()


class ExcelOperationTimeout(TimeoutError):
    """A bounded Excel call expired, with the owned-process cleanup outcome attached."""

    def __init__(self, operation, timeout_seconds, *, termination_error=None, process_exited=False):
        self.operation = operation
        self.timeout_seconds = timeout_seconds
        self.termination_error = termination_error
        self.process_exited = process_exited
        if termination_error:
            cleanup = (f'the owned Excel process termination attempt failed ({termination_error}); '
                       'verify the app-owned Excel window is closed manually.')
        elif process_exited:
            cleanup = 'the owned Excel process exit was confirmed after termination.'
        else:
            cleanup = 'the owned Excel process termination status is unknown; verify it is closed.'
        super().__init__(f'{operation} exceeded {timeout_seconds:g} seconds; {cleanup}')


def call_with_process_timeout(pid, callback, timeout_seconds, *, operation,
                              terminate=terminate_owned_excel, owned_process=None):
    """Bound a call and terminate, then verify exit of, only the supplied owned PID."""
    opened_here = False
    if owned_process is None and terminate is terminate_owned_excel:
        owned_process = open_owned_excel_process(pid)
        opened_here = True
    expired = threading.Event()
    termination = {'exited': False, 'error': None}
    state_lock = threading.Lock()
    state = {'completed': False}

    def stop_owned_process():
        with state_lock:
            if state['completed']:
                return
            expired.set()
        try:
            termination['exited'] = terminate(owned_process if owned_process is not None else pid) is True
        except Exception as exc:
            # Keep only the exception type; raw OS error details can contain host data.
            termination['error'] = type(exc).__name__

    def timeout_error():
        try:
            # terminate_owned_excel waits up to five seconds for this PID. Do not
            # report an unknown cleanup result while that bounded wait is active.
            timer.join(timeout=6)
        except (AttributeError, RuntimeError):
            pass
        return ExcelOperationTimeout(operation, timeout_seconds,
            termination_error=termination['error'],
            process_exited=termination['exited'])

    timer = threading.Timer(timeout_seconds, stop_owned_process)
    timer.daemon = True
    timer.start()
    try:
        try:
            result = callback()
        except Exception as exc:
            with state_lock:
                state['completed'] = True
                timed_out = expired.is_set()
            if timed_out:
                raise timeout_error() from exc
            raise
        with state_lock:
            state['completed'] = True
            timed_out = expired.is_set()
        if timed_out:
            raise timeout_error()
        return result
    finally:
        timer.cancel()
        if opened_here:
            owned_process.close()


def call_with_timeout(excel, callback, timeout_seconds, *, enabled, operation):
    """Run a potentially modal/blocking Excel call and stop only its DispatchEx process."""
    if not enabled:
        return callback()
    try:
        pid = owned_excel_pid(excel)
    except Exception as exc:
        raise RuntimeError(f'Cannot safely bound {operation}: the owned Excel process ID is unavailable.') from exc
    process = open_owned_excel_process(pid, hwnd=int(excel.Hwnd))
    try:
        return call_with_process_timeout(pid, callback, timeout_seconds, operation=operation,
                                         terminate=terminate_owned_excel, owned_process=process)
    finally:
        process.close()
