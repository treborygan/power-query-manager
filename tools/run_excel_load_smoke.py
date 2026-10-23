"""Exercise app-level static-M worksheet and Data Model imports on disposable workbooks."""
from __future__ import annotations

from pathlib import Path
import tempfile
import time
from contextlib import contextmanager

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


@contextmanager
def track_excel_pids(pids, label='Excel operation'):
    original = win32com.client.DispatchEx
    def tracked(*args, **kwargs):
        instance = original(*args, **kwargs)
        pids.append((label, int(win32process.GetWindowThreadProcessId(int(instance.Hwnd))[1])))
        return instance
    win32com.client.DispatchEx = tracked
    try:
        yield
    finally:
        win32com.client.DispatchEx = original


def running(pid):
    try:
        handle = win32api.OpenProcess(win32con.SYNCHRONIZE, False, pid)
    except Exception:
        return False
    try:
        return win32event.WaitForSingleObject(handle, 0) == win32event.WAIT_TIMEOUT
    finally:
        win32api.CloseHandle(handle)


def stop_owned(pid):
    handle = win32api.OpenProcess(win32con.PROCESS_TERMINATE, False, pid)
    try:
        win32api.TerminateProcess(handle, 1)
    finally:
        win32api.CloseHandle(handle)


def assert_owned_processes_exited(pids, label, timeout=30):
    for process_label, pid in pids:
        deadline = time.monotonic() + timeout
        while running(pid) and time.monotonic() < deadline:
            time.sleep(0.1)
        if running(pid):
            stop_owned(pid)
            raise RuntimeError(f'{label}/{process_label}: owned Excel process {pid} remained after {timeout}s; smoke stopped that owned process.')


def verify_loaded_values(path, mode, query_name):
    import gc
    pythoncom.CoInitialize()
    excel = book = table = sheet = candidate = model_tables = model_table = None
    pid = None
    try:
        excel = win32com.client.DispatchEx('Excel.Application')
        pid = int(win32process.GetWindowThreadProcessId(int(excel.Hwnd))[1])
        excel.Visible = False
        excel.DisplayAlerts = False
        book = excel.Workbooks.Open(str(path), UpdateLinks=0, ReadOnly=True,
                                    IgnoreReadOnlyRecommended=True, AddToMru=False)
        if not bool(book.ReadOnly):
            raise AssertionError('Read-only data verification open was not read-only.')
        if book.Worksheets.Item('PQM_Marker').Range('A1').Value != 'KEEP-UNRELATED':
            raise AssertionError(f'{mode.value} changed an unrelated worksheet cell.')
        if mode in (LoadMode.NEW_SHEET, LoadMode.EXISTING_SHEET):
            for sheet_index in range(1, int(book.Worksheets.Count) + 1):
                sheet = book.Worksheets.Item(sheet_index)
                for table_index in range(1, int(sheet.ListObjects.Count) + 1):
                    candidate = sheet.ListObjects.Item(table_index)
                    if str(candidate.QueryTable.WorkbookConnection.Name).casefold() == f'query - {query_name}'.casefold():
                        table = candidate
            if table is None:
                raise AssertionError(f'{mode.value} table is missing after reopen.')
            values = tuple(row[0] for row in table.DataBodyRange.Value)
            if values != (3.0, 4.0):
                raise AssertionError(f'{mode.value} output rows differ from expected static M result: {values!r}')
        elif mode == LoadMode.CONNECTION_ONLY:
            connection = book.Connections.Item(f'Query - {query_name}')
            if str(connection.Name).casefold() != f'query - {query_name}'.casefold():
                raise AssertionError(f'Connection-only import created an unexpected connection for {query_name}.')
            for sheet_index in range(1, int(book.Worksheets.Count) + 1):
                sheet = book.Worksheets.Item(sheet_index)
                for table_index in range(1, int(sheet.ListObjects.Count) + 1):
                    candidate = sheet.ListObjects.Item(table_index)
                    if str(candidate.QueryTable.WorkbookConnection.Name).casefold() == f'query - {query_name}'.casefold():
                        raise AssertionError(f'Connection-only import created a worksheet output table for {query_name}.')
            model_tables = tuple(book.Model.ModelTables.Item(index)
                for index in range(1, int(book.Model.ModelTables.Count) + 1))
            if any(str(item.SourceWorkbookConnection.Name).casefold() == f'query - {query_name}'.casefold()
                   for item in model_tables):
                raise AssertionError(f'Connection-only import created a Data Model table for {query_name}.')
        else:
            model_tables = tuple(book.Model.ModelTables.Item(index)
                for index in range(1, int(book.Model.ModelTables.Count) + 1))
            model_table = next((item for item in model_tables if str(item.SourceWorkbookConnection.Name).casefold() == f'query - {query_name}'.casefold()), None)
            if model_table is None or int(model_table.RecordCount) != 2:
                raise AssertionError(f'Data Model output table/row count is incorrect for {query_name}.')
        return pid
    finally:
        if book is not None:
            try: book.Close(False)
            except Exception: pass
        book = table = sheet = candidate = model_tables = model_table = None
        gc.collect()
        if excel is not None:
            try: excel.Quit()
            except Exception: pass
        excel = None
        gc.collect()
        pythoncom.CoUninitialize()


def main():
    pids = []
    with track_excel_pids(pids, 'capability probe'):
        capability = probe_excel_capabilities()
    print(f'Excel {capability.office_version} build {capability.build}; modes: {capability.supported_load_modes}')
    modes = (LoadMode.CONNECTION_ONLY, LoadMode.NEW_SHEET, LoadMode.EXISTING_SHEET, LoadMode.DATA_MODEL)
    unsupported = [mode.value for mode in modes if not capability.supports(mode.value)]
    if unsupported:
        raise RuntimeError(f'Office capability probe did not certify modes: {unsupported}; reasons: {capability.reasons}')
    assert_owned_processes_exited(pids, 'Capability probe', timeout=45)
    pids.clear()
    with tempfile.TemporaryDirectory(prefix='pqm-load-smoke-') as temp:
        root = Path(temp)
        for mode, query_name in zip(modes, ('PQM_Connéxion - Load_1', 'PQM_New_Load', 'PQM_Existing_Load', 'PQM_Model_Load')):
            pids.clear()
            book_path = root / f'{mode.value}.xlsx'
            source_path = root / f'{query_name}.m'
            formula = 'let Source = #table({"Value"}, {{1}, {2}}) in Source'
            source_path.write_text(formula, encoding='utf-8')
            pythoncom.CoInitialize()
            try:
                seed_excel = win32com.client.DispatchEx('Excel.Application')
                seed_pid = int(win32process.GetWindowThreadProcessId(int(seed_excel.Hwnd))[1])
                pids.append(('seed workbook', seed_pid))
                seed_excel.Visible = False
                seed_excel.DisplayAlerts = False
                workbook = seed_excel.Workbooks.Add()
                if mode == LoadMode.EXISTING_SHEET:
                    workbook.Worksheets.Item(1).Name = 'Output'
                marker = workbook.Worksheets.Add(After=workbook.Worksheets.Item(workbook.Worksheets.Count))
                marker.Name = 'PQM_Marker'
                marker.Range('A1').Value = 'KEEP-UNRELATED'
                workbook.SaveAs(str(book_path), 51)
                workbook.Close(False)
                workbook = marker = None
                import gc
                gc.collect()
                seed_excel.Quit()
                seed_excel = None
                gc.collect()
            finally:
                pythoncom.CoUninitialize()

            source = QuerySource(source_path, query_name, load_mode=mode,
                sheet_name='Output' if mode == LoadMode.EXISTING_SHEET else None,
                cell='B3' if mode == LoadMode.EXISTING_SHEET else None)
            assert_owned_processes_exited(pids, f'{mode.value} seed', timeout=45)
            pids.clear()
            with track_excel_pids(pids, f'{mode.value} import 1'):
                first = ExcelQueryInstaller().install(book_path, [source], ExistingQueryMode.REPLACE,
                                                      capabilities=capability)
            assert_owned_processes_exited(pids, f'{mode.value} import 1', timeout=45)
            pids.clear()
            with track_excel_pids(pids, f'{mode.value} import 2'):
                second = ExcelQueryInstaller().install(book_path, [source], ExistingQueryMode.REPLACE,
                                                       capabilities=capability)
            assert_owned_processes_exited(pids, f'{mode.value} import 2', timeout=45)
            pids.clear()
            if first[0].status.value != 'Created' or second[0].status.value != 'Replaced':
                raise AssertionError((mode.value, first, second))
            changed_formula = 'let Source = #table({"Value"}, {{3}, {4}}) in Source'
            source_path.write_text(changed_formula, encoding='utf-8')
            with track_excel_pids(pids, f'{mode.value} changed M import'):
                third = ExcelQueryInstaller().install(book_path, [source], ExistingQueryMode.REPLACE,
                                                      capabilities=capability)
            assert_owned_processes_exited(pids, f'{mode.value} changed M import', timeout=45)
            pids.clear()
            if third[0].status.value != 'Replaced':
                raise AssertionError((mode.value, third))
            with track_excel_pids(pids, f'{mode.value} inspector'):
                observed = {q.name: q for q in ExcelQueryInspector().inspect(book_path)}[query_name]
            assert_owned_processes_exited(pids, f'{mode.value} inspector', timeout=45)
            pids.clear()
            if observed.formula != changed_formula:
                raise AssertionError(f'{mode.value} M formula failed read-only reopen verification: {observed.formula!r}')
            if mode == LoadMode.DATA_MODEL and observed.load_state != 'data_model':
                raise AssertionError((mode.value, observed.load_state))
            if mode == LoadMode.CONNECTION_ONLY and observed.load_state != 'connection_only':
                raise AssertionError((mode.value, observed.load_state))
            if mode in (LoadMode.NEW_SHEET, LoadMode.EXISTING_SHEET) and 'worksheet:' not in observed.load_state:
                raise AssertionError((mode.value, observed.load_state))
            with track_excel_pids(pids, f'{mode.value} values verifier'):
                verify_loaded_values(book_path, mode, query_name)
            assert_owned_processes_exited(pids, f'{mode.value} values verifier', timeout=45)
            pids.clear()
            print(f'PASS: {mode.value} created, repeat import and changed-formula reimport preserved/refreshed the existing mapping, values and unrelated cell verified after read-only reopen, all owned Excel processes exited; state={observed.load_state}')
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
