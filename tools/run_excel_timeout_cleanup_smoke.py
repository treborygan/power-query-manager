"""Disposable Excel-process timeout smoke; never opens a workbook or touches user Excel."""
import platform
from pathlib import Path
import sys
import time

import pythoncom
import win32api
import win32con
import win32event
import win32process
import win32com.client

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'app' / 'src'))
from pq_installer.core.excel_lifetime import call_with_timeout


def main():
    pythoncom.CoInitialize()
    excel = None
    process_handle = None
    try:
        excel = win32com.client.DispatchEx('Excel.Application')
        excel.Visible = False
        excel.DisplayAlerts = False
        version = str(excel.Version)
        build = str(excel.Build)
        pid = int(win32process.GetWindowThreadProcessId(int(excel.Hwnd))[1])
        process_handle = win32api.OpenProcess(win32con.SYNCHRONIZE, False, pid)
        try:
            call_with_timeout(excel, lambda: time.sleep(2), 1,
                              enabled=True, operation='disposable Excel timeout smoke')
        except TimeoutError as exc:
            if not getattr(exc, 'process_exited', False):
                raise RuntimeError(f'Owned Excel cleanup was not confirmed: {exc}') from exc
        else:
            raise RuntimeError('Timeout smoke unexpectedly completed without timing out.')
        if win32event.WaitForSingleObject(process_handle, 0) != win32event.WAIT_OBJECT_0:
            raise RuntimeError('The exact DispatchEx process handle did not signal exit.')
        print(f'PASS: owned Excel timeout cleanup; Excel {version} build {build}; Python {platform.python_version()}')
    finally:
        if process_handle is not None:
            win32api.CloseHandle(process_handle)
        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()


if __name__ == '__main__':
    main()
