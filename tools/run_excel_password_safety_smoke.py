"""Confirm password-protected disposable workbooks fail closed without changing bytes."""
from __future__ import annotations

import hashlib
from pathlib import Path
import tempfile
import time

import pythoncom
import win32api
import win32con
import win32event
import win32process
import win32com.client

from pq_installer.core.excel_inspector import ExcelQueryInspector
from pq_installer.core.excel_queries import ExcelQueryInstaller
from pq_installer.core.excel_lifetime import ExcelOperationTimeout
from pq_installer.core.models import ExistingQueryMode, QuerySource
import pq_installer.core.excel_queries as installer_module
import pq_installer.core.excel_inspector as inspector_module

# Exercise timeout cleanup quickly; production defaults remain 60 seconds.
installer_module.OPEN_TIMEOUT_SECONDS=10
inspector_module.OPEN_TIMEOUT_SECONDS=10


def running(pid):
    try:
        handle = win32api.OpenProcess(win32con.SYNCHRONIZE, False, pid)
    except Exception:
        return False
    try:
        return win32event.WaitForSingleObject(handle, 0) == win32event.WAIT_TIMEOUT
    finally:
        win32api.CloseHandle(handle)


def main():
    pids=[]
    original_dispatch=win32com.client.DispatchEx
    def tracked_dispatch(*args,**kwargs):
        instance=original_dispatch(*args,**kwargs)
        pids.append(int(win32process.GetWindowThreadProcessId(int(instance.Hwnd))[1]))
        return instance
    with tempfile.TemporaryDirectory(prefix='pqm-password-safety-') as temp:
        root=Path(temp);book_path=root/'password.xlsx';source_path=root/'query.m'
        source_path.write_text('let Source = 1 in Source',encoding='utf-8')
        pythoncom.CoInitialize()
        try:
            excel=original_dispatch('Excel.Application');excel.Visible=False;excel.DisplayAlerts=False
            workbook=excel.Workbooks.Add()
            workbook.SaveAs(str(book_path),51,Password='pqm-test-password')
            workbook.Close(False);workbook=None
            excel.Quit();excel=None
        finally:
            pythoncom.CoUninitialize()
        original_hash=hashlib.sha256(book_path.read_bytes()).hexdigest()
        win32com.client.DispatchEx=tracked_dispatch
        try:
            try:
                ExcelQueryInspector().inspect(book_path)
            except (RuntimeError, ExcelOperationTimeout) as exc:
                print(f'Inspection blocked as expected: {exc}')
            else:
                raise AssertionError('Read-only inspection unexpectedly opened a password-protected workbook.')
            try:
                ExcelQueryInstaller().install(book_path,[QuerySource(source_path,'PasswordProbe')],ExistingQueryMode.REPLACE,
                    expected_workbook_sha256=original_hash)
            except (RuntimeError, ExcelOperationTimeout) as exc:
                print(f'Import blocked as expected: {exc}')
            else:
                raise AssertionError('Staged import unexpectedly opened a password-protected workbook.')
        finally:
            win32com.client.DispatchEx=original_dispatch
        if hashlib.sha256(book_path.read_bytes()).hexdigest()!=original_hash:
            raise AssertionError('Password-protected workbook bytes changed after blocked operations.')
        if list(root.glob('*.pqstage-*')):
            raise AssertionError('A failed password-protected stage was not cleaned up.')
    for pid in pids:
        deadline=time.monotonic()+20
        while running(pid) and time.monotonic()<deadline:time.sleep(0.1)
        if running(pid):
            handle=win32api.OpenProcess(win32con.PROCESS_TERMINATE,False,pid)
            try:win32api.TerminateProcess(handle,1)
            finally:win32api.CloseHandle(handle)
            raise RuntimeError(f'Owned Excel process {pid} remained after password smoke; stopped that owned process.')
    print('PASS: password-protected disposable workbook remained byte-identical; inspection/import stopped; owned Excel processes exited.')
    return 0


if __name__=='__main__':
    raise SystemExit(main())
