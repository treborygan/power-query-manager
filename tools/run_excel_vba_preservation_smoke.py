"""Verify staged import preserves a harmless VBA project without running its Open event."""
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
from pq_installer.core.excel_queries import ExcelQueryInstaller
from pq_installer.core.models import ExistingQueryMode, QuerySource


def _hash(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def _event_source(path: Path) -> str:
    app = win32com.client.DispatchEx('Excel.Application')
    app.Visible = False
    app.DisplayAlerts = False
    app.AutomationSecurity = 3
    book = app.Workbooks.Open(str(path), UpdateLinks=0, ReadOnly=True,
                              IgnoreReadOnlyRecommended=True, AddToMru=False)
    try:
        module = book.VBProject.VBComponents.Item('ThisWorkbook').CodeModule
        return module.Lines(1, int(module.CountOfLines))
    finally:
        book.Close(False)
        book = None
        app.Quit()
        app = None


def _alive(pid: int) -> bool:
    try:
        handle = win32api.OpenProcess(win32con.SYNCHRONIZE, False, pid)
    except Exception:
        return False
    try:
        return win32event.WaitForSingleObject(handle, 0) == win32event.WAIT_TIMEOUT
    finally:
        win32api.CloseHandle(handle)


def main() -> int:
    original_dispatch = win32com.client.DispatchEx
    owned_pids: list[int] = []

    def tracked_dispatch(*args, **kwargs):
        app = original_dispatch(*args, **kwargs)
        owned_pids.append(int(win32process.GetWindowThreadProcessId(int(app.Hwnd))[1]))
        return app

    win32com.client.DispatchEx = tracked_dispatch
    pythoncom.CoInitialize()
    try:
        with tempfile.TemporaryDirectory(prefix='pqm-vba-preservation-') as temporary:
            root = Path(temporary)
            workbook_path = root / 'vba-preservation.xlsm'
            source_path = root / 'probe.m'
            formula = 'let Source = #table({"Value"}, {{23}, {41}}) in Source'
            source_path.write_text(formula, encoding='utf-8')

            excel = win32com.client.DispatchEx('Excel.Application')
            excel.Visible = False
            excel.DisplayAlerts = False
            excel.AutomationSecurity = 3
            version, build = str(excel.Version), str(excel.Build)
            book = excel.Workbooks.Add()
            book.Worksheets.Item(1).Range('A1').Value = 'VBA_NOT_RUN'
            book.SaveAs(str(workbook_path), 52)
            event_code = ('Private Sub Workbook_Open()\r\n'
                          '    Worksheets(1).Range("A1").Value = "MACRO_RAN"\r\n'
                          'End Sub')
            code_module = book.VBProject.VBComponents.Item('ThisWorkbook').CodeModule
            code_module.AddFromString(event_code)
            book.Save()
            book.Close(False)
            book = None
            excel.Quit()
            excel = None

            with zipfile.ZipFile(workbook_path) as archive:
                if 'xl/vbaProject.bin' not in archive.namelist():
                    raise AssertionError('Disposable VBA project was not embedded in the .xlsm.')
            before_event_source = _event_source(workbook_path)
            if 'Workbook_Open' not in before_event_source or 'MACRO_RAN' not in before_event_source:
                raise AssertionError('The harmless Workbook_Open fixture was not saved correctly.')

            result = ExcelQueryInstaller().install(
                workbook_path,
                [QuerySource(source_path, 'PQM_VbaSafetyProbe')],
                ExistingQueryMode.REPLACE,
            )
            if len(result) != 1 or result[0].status.value != 'Created':
                raise AssertionError(f'Unexpected staged import outcome: {result!r}')

            queries = ExcelQueryInspector().inspect(workbook_path)
            match = [query for query in queries if query.name == 'PQM_VbaSafetyProbe']
            if len(match) != 1 or match[0].formula != formula:
                raise AssertionError('Read-only inspection did not recover the exact imported M formula.')
            with zipfile.ZipFile(workbook_path) as archive:
                if 'xl/vbaProject.bin' not in archive.namelist():
                    raise AssertionError('The VBA project was dropped during staged query import.')
            after_event_source = _event_source(workbook_path)
            normalize = lambda code: [line.strip() for line in code.replace('\r', '').split('\n') if line.strip()]
            if normalize(after_event_source) != normalize(before_event_source):
                raise AssertionError('The ThisWorkbook event source changed during staged query import.')

            verify = win32com.client.DispatchEx('Excel.Application')
            verify.Visible = False
            verify.DisplayAlerts = False
            verify.AutomationSecurity = 3
            verify_book = verify.Workbooks.Open(str(workbook_path), UpdateLinks=0, ReadOnly=True,
                                                IgnoreReadOnlyRecommended=True, AddToMru=False)
            try:
                marker = verify_book.Worksheets.Item(1).Range('A1').Value2
                if marker != 'VBA_NOT_RUN':
                    raise AssertionError(f'The Workbook_Open event ran during import/open: {marker!r}')
                code_module = verify_book.VBProject.VBComponents.Item('ThisWorkbook').CodeModule
                code = code_module.Lines(1, int(code_module.CountOfLines))
                if 'MACRO_RAN' not in code or 'Workbook_Open' not in code:
                    raise AssertionError('The harmless Workbook_Open VBA source was not preserved.')
            finally:
                verify_book.Close(False)
                verify_book = None
                verify.Quit()
                verify = None

            for pid in set(owned_pids):
                deadline = time.monotonic() + 15
                while _alive(pid) and time.monotonic() < deadline:
                    time.sleep(0.1)
                if _alive(pid):
                    raise RuntimeError(f'Smoke-owned Excel process {pid} did not exit.')
            print(f'PASS Excel {version} build {build}; VBA-bearing disposable .xlsm import preserved '
                  'Workbook_Open source; marker stayed unchanged; query formula verified; '
                  'all smoke-owned Excel processes exited.')
            return 0
    finally:
        win32com.client.DispatchEx = original_dispatch
        pythoncom.CoUninitialize()


if __name__ == '__main__':
    raise SystemExit(main())
