"""Verify internet-zone workbook marks fail closed before any Excel automation."""
from __future__ import annotations

import hashlib
import os
from pathlib import Path
import tempfile
import time
import gc

import pythoncom
import win32api
import win32con
import win32process
import win32event
import win32com.client

from pq_installer.core.excel_inspector import ExcelQueryInspector
from pq_installer.core.excel_queries import ExcelQueryInstaller
from pq_installer.core.models import ExistingQueryMode, QuerySource


def _sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def _wait_owned_excel(pid: int, timeout: float = 15) -> None:
    handle = win32api.OpenProcess(win32con.SYNCHRONIZE, False, pid)
    try:
        result = win32event.WaitForSingleObject(handle, int(timeout * 1000))
        if result != win32event.WAIT_OBJECT_0:
            raise RuntimeError(f'Smoke-owned Excel PID {pid} did not exit after Quit().')
    finally:
        win32api.CloseHandle(handle)


def main() -> int:
    with tempfile.TemporaryDirectory(prefix='pqm-zone-mark-') as temporary:
        root = Path(temporary)
        workbook_path = root / 'internet.xlsx'
        source_path = root / 'query.m'
        source_path.write_text('let Source = 1 in Source', encoding='utf-8')

        pythoncom.CoInitialize()
        excel = book = None
        owned_pid = None
        try:
            excel = win32com.client.DispatchEx('Excel.Application')
            owned_pid = int(win32process.GetWindowThreadProcessId(int(excel.Hwnd))[1])
            excel.Visible = False
            excel.DisplayAlerts = False
            excel.AutomationSecurity = 3
            book = excel.Workbooks.Add()
            book.SaveAs(str(workbook_path), 51)
            book.Close(False)
            book = None
            excel.Quit()
            excel = None
            gc.collect()
            _wait_owned_excel(owned_pid)
        finally:
            if book is not None:
                try: book.Close(False)
                except Exception: pass
            if excel is not None:
                try: excel.Quit()
                except Exception: pass
            book = excel = None
            gc.collect()
            pythoncom.CoUninitialize()

        original_hash = _sha256(workbook_path)
        zone_stream = Path(str(workbook_path) + ':Zone.Identifier')
        zone_bytes = b'[ZoneTransfer]\r\nZoneId=3\r\nHostUrl=https://example.invalid/disposable.xlsx\r\n'
        with zone_stream.open('wb') as stream:
            stream.write(zone_bytes)

        factories = []
        inspector = ExcelQueryInspector(lambda: factories.append('inspector'))
        installer = ExcelQueryInstaller(lambda: factories.append('installer'))
        for operation in (
            lambda: inspector.inspect(workbook_path),
            lambda: installer.install(workbook_path, [QuerySource(source_path, 'ZoneProbe')],
                                       ExistingQueryMode.REPLACE),
        ):
            try:
                operation()
            except ValueError as exc:
                if 'Internet or Restricted zone' not in str(exc):
                    raise
            else:
                raise AssertionError('Internet-zone workbook was unexpectedly opened by the automation path.')
        if factories:
            raise AssertionError(f'Excel application factory was called before zone-mark refusal: {factories!r}.')
        if _sha256(workbook_path) != original_hash:
            raise AssertionError('Internet-zone workbook bytes changed after refused automation.')
        if zone_stream.read_bytes() != zone_bytes:
            raise AssertionError('Internet-zone metadata changed after refused automation.')
        if list(root.glob('*.pqstage-*')):
            raise AssertionError('Zone-mark refusal created a staged workbook.')

    print(f'PASS: real Excel-created disposable .xlsx with ZoneId=3 was refused by inspect/import before another Excel instance was created; workbook hash and Zone.Identifier remained unchanged; smoke-owned seed Excel PID {owned_pid} exited.')
    print('Scope: confirms fail-closed handling of the Windows Internet zone mark; it does not exercise or certify Excel Protected View UI behavior.')
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
