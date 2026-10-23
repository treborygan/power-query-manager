"""Opt-in, disposable real-Excel refresh smoke using a static M table (no external source)."""
from __future__ import annotations

import argparse
import hashlib
import os
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
from pq_installer.core.excel_refresh import ExcelQueryRefresher


def process_is_running(pid: int) -> bool:
    try:
        handle = win32api.OpenProcess(win32con.SYNCHRONIZE, False, pid)
    except Exception:
        return False
    try:
        return win32event.WaitForSingleObject(handle, 0) == win32event.WAIT_TIMEOUT
    finally:
        win32api.CloseHandle(handle)


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--mode', choices=('background', 'synchronous'), default='background')
    mode = parser.parse_args().mode
    pythoncom.CoInitialize()
    excel = None
    try:
        with tempfile.TemporaryDirectory(prefix='pqm-refresh-smoke-') as temp:
            workbook_path = Path(temp) / 'static-refresh.xlsx'
            excel = win32com.client.DispatchEx('Excel.Application')
            excel.Visible = False
            excel.DisplayAlerts = False
            excel.AutomationSecurity = 3
            version, build = str(excel.Version), str(excel.Build)
            workbook = excel.Workbooks.Add()
            workbook.Queries.Add('PQM_Refresh_Smoke', 'let Source = #table({"Value"}, {{1}}) in Source')
            connection_string = ('OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;'
                                 'Location=PQM_Refresh_Smoke;Extended Properties=""')
            workbook.Connections.Add2('Query - PQM_Refresh_Smoke', 'Static refresh smoke',
                                      connection_string, 'SELECT * FROM [PQM_Refresh_Smoke]',
                                      2, False, False)
            workbook.SaveAs(str(workbook_path), 51)
            workbook.Close(False)
            workbook = None
            excel.Quit()
            excel = None

            original_hash = hashlib.sha256(workbook_path.read_bytes()).hexdigest()
            refresh_pid = None

            def create_owned_excel():
                nonlocal refresh_pid
                instance = win32com.client.DispatchEx('Excel.Application')
                refresh_pid = int(win32process.GetWindowThreadProcessId(int(instance.Hwnd))[1])
                return instance

            results = ExcelQueryRefresher(excel_factory=create_owned_excel).refresh(
                workbook_path, ('PQM_Refresh_Smoke',), timeout_seconds=60,
                asynchronous=(mode == 'background'), make_backup=True,
            )
            assert len(results) == 1 and results[0].status == 'refreshed', results
            backup = next(Path(temp).glob('static-refresh_backup_*.xlsx'))
            assert hashlib.sha256(backup.read_bytes()).hexdigest() == original_hash
            inventory = ExcelQueryInspector().inspect(workbook_path)
            query = next(item for item in inventory if item.name == 'PQM_Refresh_Smoke')
            assert query.formula == 'let Source = #table({"Value"}, {{1}}) in Source'
            assert query.load_state == 'connection_only', query.load_state

            deadline = time.monotonic() + 10
            while refresh_pid and process_is_running(refresh_pid) and time.monotonic() < deadline:
                time.sleep(0.1)
            if refresh_pid and process_is_running(refresh_pid):
                raise RuntimeError(f'Refresh-owned Excel process {refresh_pid} remained after cleanup.')
            print(f'PASS: Excel {version} build {build}; {mode} static M refresh succeeded; backup hash preserved; query reopened; owned Excel process exited.')
            return 0
    finally:
        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()


if __name__ == '__main__':
    raise SystemExit(main())
