"""Real Excel staged-import smoke on a disposable workbook and static M source."""
from __future__ import annotations

import argparse
from pathlib import Path
import tempfile
import time

import pythoncom
import win32api
import win32con
import win32event
import win32process
import win32com.client

from pq_installer.core.excel_capabilities import probe_excel_capabilities
from pq_installer.core.excel_inspector import ExcelQueryInspector
from pq_installer.core.excel_queries import ExcelQueryInstaller
from pq_installer.core.models import ExistingQueryMode, LoadMode, QuerySource


def _running(pid: int) -> bool:
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
    parser.add_argument('--xlsm', action='store_true',
                        help='use a disposable macro-enabled workbook container instead of .xlsx')
    args = parser.parse_args()
    pythoncom.CoInitialize()
    owned_pid = None
    try:
        with tempfile.TemporaryDirectory(prefix='pqm-import-smoke-') as temp:
            root = Path(temp)
            workbook_path = root / ('import-smoke.xlsm' if args.xlsm else 'import-smoke.xlsx')
            source_path = root / 'probe.m'
            source_path.write_text('let Source = #table({"Value"}, {{1}, {2}}) in Source', encoding='utf-8')
            excel = win32com.client.DispatchEx('Excel.Application')
            owned_pid = int(win32process.GetWindowThreadProcessId(int(excel.Hwnd))[1])
            excel.Visible = False
            excel.DisplayAlerts = False
            excel.AutomationSecurity = 3
            version, build = str(excel.Version), str(excel.Build)
            workbook = excel.Workbooks.Add()
            workbook.SaveAs(str(workbook_path), 52 if args.xlsm else 51)
            workbook.Close(False)
            excel.Quit()
            excel = None

            capability = probe_excel_capabilities(staging_parent=root)
            if not capability.supports('connection_only'):
                raise RuntimeError(f'Connection-only capability probe did not pass: {capability.reasons}')
            source = QuerySource(source_path, 'PQM_Import_Smoke', load_mode=LoadMode.CONNECTION_ONLY)
            owned_pids = [owned_pid]
            original_dispatch = win32com.client.DispatchEx
            def tracked_dispatch(*args, **kwargs):
                instance = original_dispatch(*args, **kwargs)
                owned_pids.append(int(win32process.GetWindowThreadProcessId(int(instance.Hwnd))[1]))
                return instance
            win32com.client.DispatchEx = tracked_dispatch
            try:
                result = ExcelQueryInstaller().install(workbook_path, [source], ExistingQueryMode.REPLACE,
                                                       capabilities=capability)
                assert result[0].status.value == 'Created', result
                result = ExcelQueryInstaller().install(workbook_path, [source], ExistingQueryMode.REPLACE,
                                                       capabilities=capability)
                assert result[0].status.value == 'Replaced', result
            finally:
                win32com.client.DispatchEx = original_dispatch
            query = ExcelQueryInspector().inspect(workbook_path)
            observed = next(item for item in query if item.name == 'PQM_Import_Smoke')
            assert observed.formula == source_path.read_text(encoding='utf-8')
            assert observed.load_state == 'connection_only', observed.load_state
            deadline = time.monotonic() + 10
            while owned_pid and _running(owned_pid) and time.monotonic() < deadline:
                time.sleep(0.1)
            for pid in owned_pids:
                deadline = time.monotonic() + 10
                while _running(pid) and time.monotonic() < deadline:
                    time.sleep(0.1)
                if _running(pid):
                    raise RuntimeError(f'Import-owned Excel process {pid} remained after cleanup.')
            print(f'PASS: Excel {version} build {build}; {workbook_path.suffix} staged create and repeat-replace, read-only reopen verification, connection-only preservation, and owned Excel cleanup succeeded.')
            return 0
    finally:
        pythoncom.CoUninitialize()


if __name__ == '__main__':
    raise SystemExit(main())
