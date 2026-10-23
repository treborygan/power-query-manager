"""Test staged refresh on disposable VBA-bearing workbooks with no external sources."""
from __future__ import annotations

import hashlib
from pathlib import Path
import tempfile
import time
import zipfile

import pythoncom
import win32api
import win32con
import win32event
import win32process
import win32com.client

from pq_installer.core.excel_inspector import ExcelQueryInspector
from pq_installer.core.excel_refresh import ExcelQueryRefresher

QUERY_NAME = 'PQM_VbaRefreshSmoke'
FORMULA = 'let Source = #table({"Value"}, {{5}, {8}}) in Source'
EVENT_SOURCE = ('Private Sub Workbook_Open()\r\n'
                '    Worksheets(1).Range("A1").Value = "MACRO_RAN"\r\n'
                'End Sub')


def _alive(pid: int) -> bool:
    try:
        handle = win32api.OpenProcess(win32con.SYNCHRONIZE, False, pid)
    except Exception:
        return False
    try:
        return win32event.WaitForSingleObject(handle, 0) == win32event.WAIT_TIMEOUT
    finally:
        win32api.CloseHandle(handle)


def _norm(value: str) -> list[str]:
    return [line.strip() for line in value.replace('\r', '').split('\n') if line.strip()]


def _source(workbook) -> str:
    module = workbook.VBProject.VBComponents.Item('ThisWorkbook').CodeModule
    return module.Lines(1, int(module.CountOfLines))


def main() -> int:
    original_dispatch = win32com.client.DispatchEx
    owned_pids: list[int] = []

    def tracked_dispatch(*args, **kwargs):
        app = original_dispatch(*args, **kwargs)
        owned_pids.append(int(win32process.GetWindowThreadProcessId(int(app.Hwnd))[1]))
        return app

    win32com.client.DispatchEx = tracked_dispatch
    pythoncom.CoInitialize()
    excel_version = excel_build = 'unknown'
    try:
        with tempfile.TemporaryDirectory(prefix='pqm-vba-refresh-') as temporary:
            root = Path(temporary)
            for asynchronous in (False, True):
                mode = 'background' if asynchronous else 'synchronous'
                path = root / f'vba-refresh-{mode}.xlsm'
                app = win32com.client.DispatchEx('Excel.Application')
                app.Visible = False
                app.DisplayAlerts = False
                app.AutomationSecurity = 3
                excel_version, excel_build = str(app.Version), str(app.Build)
                book = app.Workbooks.Add()
                book.Worksheets.Item(1).Range('A1').Value = 'VBA_NOT_RUN'
                book.Queries.Add(QUERY_NAME, FORMULA)
                connection_string = ('OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;'
                                     f'Location={QUERY_NAME};Extended Properties=""')
                book.Connections.Add2(f'Query - {QUERY_NAME}', 'Disposable VBA refresh probe',
                                      connection_string, f'SELECT * FROM [{QUERY_NAME}]', 2, False, False)
                book.SaveAs(str(path), 52)
                module = book.VBProject.VBComponents.Item('ThisWorkbook').CodeModule
                module.AddFromString(EVENT_SOURCE)
                book.Save()
                before_source = _source(book)
                book.Close(False)
                book = None
                app.Quit()
                app = None
                if _norm(before_source) != _norm(EVENT_SOURCE):
                    raise AssertionError('Workbook_Open fixture source changed before refresh.')

                original_hash = hashlib.sha256(path.read_bytes()).hexdigest()
                results = ExcelQueryRefresher().refresh(
                    path, (QUERY_NAME,), timeout_seconds=60,
                    asynchronous=asynchronous, make_backup=True)
                if len(results) != 1 or results[0].status != 'refreshed':
                    raise AssertionError(f'{mode} refresh failed: {results!r}')
                backup = next(root.glob(f'{path.stem}_backup_*{path.suffix}'))
                if hashlib.sha256(backup.read_bytes()).hexdigest() != original_hash:
                    raise AssertionError(f'{mode} backup did not match original workbook bytes.')
                inventory = ExcelQueryInspector().inspect(path)
                query = next(item for item in inventory if item.name == QUERY_NAME)
                if query.formula != FORMULA or query.load_state != 'connection_only':
                    raise AssertionError(f'{mode} query did not reopen correctly: {query!r}')
                with zipfile.ZipFile(path) as archive:
                    if 'xl/vbaProject.bin' not in archive.namelist():
                        raise AssertionError(f'{mode} refresh dropped the VBA project.')

                verify = win32com.client.DispatchEx('Excel.Application')
                verify.Visible = False
                verify.DisplayAlerts = False
                verify.AutomationSecurity = 3
                check = verify.Workbooks.Open(str(path), UpdateLinks=0, ReadOnly=True,
                    IgnoreReadOnlyRecommended=True, AddToMru=False)
                try:
                    marker = check.Worksheets.Item(1).Range('A1').Value2
                    after_source = _source(check)
                    if marker != 'VBA_NOT_RUN':
                        raise AssertionError(f'{mode} Workbook_Open ran during refresh/open: {marker!r}')
                    if _norm(after_source) != _norm(before_source):
                        raise AssertionError(f'{mode} Workbook_Open source changed during refresh.')
                finally:
                    check.Close(False)
                    check = None
                    verify.Quit()
                    verify = None

            for pid in set(owned_pids):
                deadline = time.monotonic() + 15
                while _alive(pid) and time.monotonic() < deadline:
                    time.sleep(0.1)
                if _alive(pid):
                    raise RuntimeError(f'Smoke-owned Excel process {pid} did not exit.')
            print(f'PASS Excel {excel_version} build {excel_build}; staged synchronous and background refreshes '
                  'preserved Workbook_Open source, did not run the event, verified formulas/backups, and exited '
                  'all owned Excel processes; static inline M only.')
            return 0
    finally:
        win32com.client.DispatchEx = original_dispatch
        pythoncom.CoUninitialize()


if __name__ == '__main__':
    raise SystemExit(main())
