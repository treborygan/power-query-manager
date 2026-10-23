from types import SimpleNamespace
import sys

import pytest

from pq_installer.core import excel_lifetime


def test_timeout_stops_only_the_pid_belonging_to_the_owned_excel_object(monkeypatch):
    targets=[]
    monkeypatch.setitem(sys.modules,'win32process',SimpleNamespace(
        GetWindowThreadProcessId=lambda hwnd:(77,4242)))
    monkeypatch.setattr(excel_lifetime,'open_owned_excel_process',
        lambda pid,hwnd:SimpleNamespace(pid=pid,hwnd=hwnd,close=lambda:None))
    monkeypatch.setattr(excel_lifetime,'terminate_owned_excel',lambda pid:(targets.append(pid) or True))
    class ImmediateTimer:
        def __init__(self,seconds,callback,daemon=False):self.callback=callback
        def start(self):self.callback()
        def cancel(self):pass
    monkeypatch.setattr(excel_lifetime.threading,'Timer',ImmediateTimer)
    with pytest.raises(TimeoutError,match='process exit was confirmed'):
        excel_lifetime.call_with_timeout(SimpleNamespace(Hwnd=123),lambda:'late',1,
            enabled=True,operation='workbook open')
    assert len(targets)==1 and targets[0].pid==4242


def test_injected_adapter_skips_process_termination_and_timer(monkeypatch):
    monkeypatch.setattr(excel_lifetime.threading,'Timer',lambda *args,**kwargs:pytest.fail('timer should not be created'))
    assert excel_lifetime.call_with_timeout(object(),lambda:'done',1,
        enabled=False,operation='fake adapter')=='done'


def test_timeout_does_not_claim_excel_stopped_when_termination_fails(monkeypatch):
    monkeypatch.setitem(sys.modules,'win32process',SimpleNamespace(
        GetWindowThreadProcessId=lambda hwnd:(77,4242)))
    monkeypatch.setattr(excel_lifetime,'open_owned_excel_process',
        lambda pid,hwnd:SimpleNamespace(pid=pid,hwnd=hwnd,close=lambda:None))
    def fail_termination(_pid):
        raise PermissionError('access denied')
    monkeypatch.setattr(excel_lifetime,'terminate_owned_excel',fail_termination)
    class ImmediateTimer:
        def __init__(self,seconds,callback,daemon=False):self.callback=callback
        def start(self):self.callback()
        def cancel(self):pass
        def join(self,timeout=None):pass
    monkeypatch.setattr(excel_lifetime.threading,'Timer',ImmediateTimer)
    with pytest.raises(TimeoutError,match='termination attempt failed') as caught:
        excel_lifetime.call_with_timeout(SimpleNamespace(Hwnd=123),lambda:'late',1,
            enabled=True,operation='workbook open')
    assert getattr(caught.value,'termination_error',None)=='PermissionError'


def test_timeout_waits_for_the_full_owned_process_cleanup_bound(monkeypatch):
    joined=[]
    class ImmediateTimer:
        def __init__(self,seconds,callback,daemon=False):self.callback=callback
        def start(self):self.callback()
        def join(self,timeout=None):joined.append(timeout)
        def cancel(self):pass
    monkeypatch.setattr(excel_lifetime.threading,'Timer',ImmediateTimer)
    with pytest.raises(TimeoutError,match='process exit was confirmed'):
        excel_lifetime.call_with_process_timeout(4242,lambda:'late',1,
            operation='blocked workbook open',terminate=lambda _pid:True)
    assert joined and joined[0] >= 6


def test_timer_race_cannot_terminate_excel_after_call_has_completed(monkeypatch):
    targets=[]
    class RacyTimer:
        def __init__(self,seconds,callback,daemon=False):self.callback=callback
        def start(self):pass
        def cancel(self):self.callback()
    monkeypatch.setattr(excel_lifetime.threading,'Timer',RacyTimer)

    result=excel_lifetime.call_with_process_timeout(
        4242,lambda:'completed',60,operation='completed workbook operation',
        terminate=lambda pid:(targets.append(pid) or True))

    assert result=='completed'
    assert targets==[]


def test_owned_excel_termination_waits_for_exit_and_closes_only_its_handle(monkeypatch):
    calls=[]
    api=SimpleNamespace(OpenProcess=lambda access,inherit,pid:(calls.append(('open',access,inherit,pid)) or 'owned-handle'),
                        TerminateProcess=lambda handle,code:calls.append(('terminate',handle,code)),
                        CloseHandle=lambda handle:calls.append(('close',handle)))
    con=SimpleNamespace(PROCESS_TERMINATE=1,SYNCHRONIZE=2)
    event=SimpleNamespace(WAIT_OBJECT_0=0,WaitForSingleObject=lambda handle,timeout:(calls.append(('wait',handle,timeout)) or 0))
    monkeypatch.setitem(sys.modules,'win32api',api)
    monkeypatch.setitem(sys.modules,'win32con',con)
    monkeypatch.setitem(sys.modules,'win32event',event)
    excel_lifetime.terminate_owned_excel(4242)
    assert calls==[('open',3,False,4242),('terminate','owned-handle',1),
                   ('wait','owned-handle',5000),('close','owned-handle')]


def test_owned_excel_termination_fails_if_process_does_not_exit(monkeypatch):
    api=SimpleNamespace(OpenProcess=lambda *args:'owned-handle',TerminateProcess=lambda *args:None,
                        CloseHandle=lambda handle:None)
    con=SimpleNamespace(PROCESS_TERMINATE=1,SYNCHRONIZE=2)
    event=SimpleNamespace(WAIT_OBJECT_0=0,WaitForSingleObject=lambda handle,timeout:258)
    monkeypatch.setitem(sys.modules,'win32api',api)
    monkeypatch.setitem(sys.modules,'win32con',con)
    monkeypatch.setitem(sys.modules,'win32event',event)
    with pytest.raises(TimeoutError,match='waiting for the owned Excel process to exit'):
        excel_lifetime.terminate_owned_excel(4242)


def test_timeout_termination_uses_the_captured_process_handle_not_a_reopened_pid(monkeypatch):
    process=SimpleNamespace(pid=4242, native_handle='captured-process-handle')
    targets=[]
    class ImmediateTimer:
        def __init__(self,seconds,callback,daemon=False):self.callback=callback
        def start(self):self.callback()
        def cancel(self):pass
        def join(self,timeout=None):pass
    monkeypatch.setattr(excel_lifetime.threading,'Timer',ImmediateTimer)

    with pytest.raises(TimeoutError,match='termination status is unknown'):
        excel_lifetime.call_with_process_timeout(
            4242, lambda:'late', 1, operation='blocked Excel call',
            terminate=lambda owned: (targets.append(owned) or False),
            owned_process=process)

    assert targets==[process]


def test_owned_process_handle_wait_and_termination_reuse_one_native_handle(monkeypatch):
    calls=[]
    api=SimpleNamespace(OpenProcess=lambda access,inherit,pid:(calls.append(('open',access,inherit,pid)) or 'process-handle'),
                        TerminateProcess=lambda handle,code:calls.append(('terminate',handle,code)),
                        CloseHandle=lambda handle:calls.append(('close',handle)))
    con=SimpleNamespace(PROCESS_TERMINATE=1,SYNCHRONIZE=2)
    event=SimpleNamespace(WAIT_OBJECT_0=0,WAIT_TIMEOUT=258,
                          WaitForSingleObject=lambda handle,timeout:(calls.append(('wait',handle,timeout)) or 0))
    monkeypatch.setitem(sys.modules,'win32api',api)
    monkeypatch.setitem(sys.modules,'win32con',con)
    monkeypatch.setitem(sys.modules,'win32event',event)

    process=excel_lifetime.open_owned_excel_process(4242)
    assert process.wait(5)
    assert process.terminate()
    process.close()

    assert calls==[('open',3,False,4242),('wait','process-handle',5),
                   ('terminate','process-handle',1),('wait','process-handle',5000),
                   ('close','process-handle')]


def test_process_handle_is_closed_if_excel_window_owner_changed_before_capture(monkeypatch):
    calls=[]
    api=SimpleNamespace(OpenProcess=lambda access,inherit,pid:(calls.append(('open',pid)) or 'process-handle'),
                        CloseHandle=lambda handle:calls.append(('close',handle)))
    con=SimpleNamespace(PROCESS_TERMINATE=1,SYNCHRONIZE=2)
    monkeypatch.setitem(sys.modules,'win32api',api)
    monkeypatch.setitem(sys.modules,'win32con',con)
    monkeypatch.setitem(sys.modules,'win32process',SimpleNamespace(
        GetWindowThreadProcessId=lambda hwnd:(77,9999)))

    with pytest.raises(RuntimeError,match='Excel window no longer belongs to the captured process'):
        excel_lifetime.open_owned_excel_process(4242,hwnd=123)

    assert calls==[('open',4242),('close','process-handle')]


def test_normal_excel_quit_waits_then_terminates_only_the_captured_dispatch_process(monkeypatch):
    calls=[]
    waits=iter((False,True))
    def wait(timeout):
        calls.append(('wait',timeout))
        return next(waits)
    def terminate(timeout):
        calls.append(('terminate',timeout))
        return wait(timeout)
    process=SimpleNamespace(
        wait=wait,
        terminate=terminate,
        close=lambda:calls.append(('close',)))
    monkeypatch.setattr(excel_lifetime,'owned_excel_pid',lambda excel:4242)
    monkeypatch.setattr(excel_lifetime,'open_owned_excel_process',
        lambda pid,hwnd: (calls.append(('open',pid,hwnd)) or process))
    excel=SimpleNamespace(Hwnd=123)

    terminated=excel_lifetime.quit_owned_excel(
        excel,quit_call=lambda:calls.append(('quit',)),timeout_ms=4000)

    assert terminated is True
    assert calls==[('open',4242,123),('quit',),('wait',4000),
                   ('terminate',5000),('wait',5000),('close',)]


def test_normal_excel_quit_does_not_terminate_when_the_owned_process_exits(monkeypatch):
    calls=[]
    process=SimpleNamespace(wait=lambda timeout:(calls.append(('wait',timeout)) or True),
        terminate=lambda timeout:pytest.fail('a cleanly closed Excel process must not be terminated'),
        close=lambda:calls.append(('close',)))
    monkeypatch.setattr(excel_lifetime,'owned_excel_pid',lambda excel:4242)
    monkeypatch.setattr(excel_lifetime,'open_owned_excel_process',
        lambda pid,hwnd:(calls.append(('open',pid,hwnd)) or process))

    terminated=excel_lifetime.quit_owned_excel(
        SimpleNamespace(Hwnd=123),quit_call=lambda:calls.append(('quit',)),timeout_ms=2500)

    assert terminated is False
    assert calls==[('open',4242,123),('quit',),('wait',2500),('close',)]


def test_normal_excel_quit_allows_bounded_office_shutdown_grace_by_default(monkeypatch):
    calls=[]
    process=SimpleNamespace(
        wait=lambda timeout:(calls.append(timeout) or True),
        terminate=lambda timeout:pytest.fail('a cleanly closed Excel process must not be terminated'),
        close=lambda:None)
    monkeypatch.setattr(excel_lifetime,'owned_excel_pid',lambda excel:4242)
    monkeypatch.setattr(excel_lifetime,'open_owned_excel_process',lambda pid,hwnd:process)
    assert excel_lifetime.quit_owned_excel(
        SimpleNamespace(Hwnd=123,Quit=lambda:None)) is False
    assert calls==[15000]


def test_normal_excel_quit_releases_com_apartment_before_waiting_for_process(monkeypatch):
    calls=[]
    process=SimpleNamespace(
        wait=lambda timeout:(calls.append(('wait',timeout)) or True),
        terminate=lambda timeout:pytest.fail('clean quit should not terminate Excel'),
        close=lambda:calls.append(('close',)))
    monkeypatch.setattr(excel_lifetime,'owned_excel_pid',lambda excel:4242)
    monkeypatch.setattr(excel_lifetime,'open_owned_excel_process',
        lambda pid,hwnd:(calls.append(('open',pid,hwnd)) or process))
    excel=SimpleNamespace(Hwnd=123,Quit=lambda:calls.append(('quit',)))
    excel_lifetime.quit_owned_excel(excel,
        after_quit=lambda:calls.append(('release_com_apartment',)))
    assert calls==[('open',4242,123),('quit',),('release_com_apartment',),
                   ('wait',15000),('close',)]


def test_normal_excel_quit_cleans_owned_process_even_when_quit_raises(monkeypatch):
    calls=[]
    waits=iter((False,True))
    def wait(timeout):
        calls.append(('wait',timeout))
        return next(waits)
    process=SimpleNamespace(wait=wait,
        terminate=lambda timeout:(calls.append(('terminate',timeout)) or wait(timeout)),
        close=lambda:calls.append(('close',)))
    monkeypatch.setattr(excel_lifetime,'owned_excel_pid',lambda excel:4242)
    monkeypatch.setattr(excel_lifetime,'open_owned_excel_process',lambda pid,hwnd:process)

    with pytest.raises(RuntimeError,match='Quit failed'):
        excel_lifetime.quit_owned_excel(SimpleNamespace(Hwnd=123),
            quit_call=lambda:(_ for _ in ()).throw(RuntimeError('Quit failed')),
            timeout_ms=1000)

    assert calls==[('wait',1000),('terminate',5000),('wait',5000),('close',)]
