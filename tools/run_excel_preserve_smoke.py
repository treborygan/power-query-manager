"""Replace query definitions and verify unrelated workbook objects survive."""
from __future__ import annotations

import hashlib
import gc
from pathlib import Path
import tempfile
import time
import sys

import pythoncom
import win32api
import win32con
import win32event
import win32process
import win32com.client

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / 'app' / 'src'))
from pq_installer.core.excel_capabilities import probe_excel_capabilities
from pq_installer.core.excel_inspector import ExcelQueryInspector
from pq_installer.core.excel_queries import ExcelQueryInstaller
from pq_installer.core.mashup_privacy import WorkbookPrivacy, read_workbook_privacy
from pq_installer.core.models import ExistingQueryMode, LoadMode, QuerySource


def _owned_dispatch(pids):
    original = win32com.client.DispatchEx
    def tracked(*args, **kwargs):
        excel = original(*args, **kwargs)
        pids.append(int(win32process.GetWindowThreadProcessId(int(excel.Hwnd))[1]))
        return excel
    return original, tracked


def _running(pid):
    try:
        handle = win32api.OpenProcess(win32con.SYNCHRONIZE, False, pid)
    except Exception:
        return False
    try:
        return win32event.WaitForSingleObject(handle, 0) == win32event.WAIT_TIMEOUT
    finally:
        win32api.CloseHandle(handle)


def _wait_owned(pids):
    for pid in pids:
        deadline = time.monotonic() + 20
        while _running(pid) and time.monotonic() < deadline:
            time.sleep(.1)
        if _running(pid):
            # All workbooks are closed and Quit was called. Confirm this is still
            # the exact Excel process returned by DispatchEx before cleanup.
            import subprocess
            result = subprocess.run(['powershell.exe', '-NoProfile', '-Command',
                f"$p=Get-CimInstance Win32_Process -Filter 'ProcessId = {pid}'; "
                "if ($p) { \"$($p.Name)|$($p.ExecutablePath)\" }"],
                capture_output=True, text=True, timeout=10)
            identity = result.stdout.strip().casefold()
            if identity != 'excel.exe|c:\\program files\\microsoft office\\root\\office16\\excel.exe':
                raise AssertionError(f'Excel PID {pid} identity changed; refusing cleanup action.')
            handle = win32api.OpenProcess(win32con.PROCESS_TERMINATE | win32con.SYNCHRONIZE, False, pid)
            try:
                win32api.TerminateProcess(handle, 1)
                if win32event.WaitForSingleObject(handle, 5000) != win32event.WAIT_OBJECT_0:
                    raise AssertionError(f'Excel process created by this smoke did not exit: {pid}')
            finally:
                win32api.CloseHandle(handle)


def _sha256(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def main():
    pythoncom.CoInitialize()
    owned_pids = []
    original_dispatch, tracked_dispatch = _owned_dispatch(owned_pids)
    win32com.client.DispatchEx = tracked_dispatch
    try:
        with tempfile.TemporaryDirectory(prefix='pqm-preserve-smoke-') as temp:
            root = Path(temp)
            workbook_path = root / 'preserve.xlsx'
            source_path = root / 'preserve.pq'
            report_source_path = root / 'report.pq'
            query_name = 'PQM_PreserveSmoke'
            report_query_name = 'PQM_ReportSmoke'
            first_formula = 'let Source = #table({"Value"}, {{1}, {2}}) in Source'
            changed_formula = 'let Source = #table({"Value"}, {{3}, {4}}) in Source'
            report_first_formula = 'let Source = #table({"Value"}, {{10}, {20}}) in Source'
            report_changed_formula = 'let Source = #table({"Value"}, {{30}, {40}}) in Source'
            excel = win32com.client.DispatchEx('Excel.Application')
            excel.Visible = False
            excel.DisplayAlerts = False
            version, build = str(excel.Version), str(excel.Build)
            book = excel.Workbooks.Add()
            sheet = book.Worksheets.Item(1)
            sheet.Range('A1:B4').Value = (('Category', 'Value'), ('Keep', 42), ('Drop', 7), ('Keep', 12))
            sheet.Range('D1').Formula = '=SUM(B2:B4)'
            sheet.Range('J1:J2').Value = (('One',), ('Two',))
            sheet.Range('F2').Validation.Add(3, 1, 1, '=$J$1:$J$2')
            condition = sheet.Range('G2:G4').FormatConditions.Add(1, 3, '=1')
            condition.Interior.Color = 65535
            hyperlink = sheet.Hyperlinks.Add(sheet.Range('H2'),
                Address='https://example.invalid/pqm-preserve', TextToDisplay='PQM preserved hyperlink')
            book.Names.Add(Name='PQM_KeepMarker',RefersTo='=Sheet1!$B$2')
            book.ForceFullCalculation = True
            plain_table = sheet.ListObjects.Add(1, sheet.Range('A1:B4'), None, 1)
            plain_table.Name = 'tblKeep'
            plain_table.TableStyle = 'TableStyleMedium2'
            plain_table.Range.AutoFilter(1,'Keep')
            chart_shape = sheet.Shapes.AddChart2(201, 51, 360, 20, 350, 220)
            chart_shape.Name = 'PQM_KeepChart'
            chart_shape.Chart.SetSourceData(plain_table.Range)
            chart_shape.Chart.HasTitle = True
            chart_shape.Chart.ChartTitle.Text = 'Preserve non-query chart'
            marker_shape = sheet.Shapes.AddShape(1, 360, 260, 180, 40)
            marker_shape.Name = 'PQM_KeepShape'
            marker_shape.AlternativeText = 'PQM unrelated shape marker'
            pivot_sheet = book.Worksheets.Add()
            pivot_sheet.Name = 'PQM_Pivot'
            pivot_source = f"'{sheet.Name}'!{plain_table.Range.Address}"
            pivot_cache = book.PivotCaches().Create(1, pivot_source)
            pivot_table = pivot_cache.CreatePivotTable(pivot_sheet.Range('A1'), 'PQM_KeepPivot')
            pivot_table.PivotFields('Category').Orientation = 1
            pivot_data_field = pivot_table.AddDataField(
                pivot_table.PivotFields('Value'), 'PQM Sum of Value', -4157)
            book.SaveAs(str(workbook_path), 51)
            plain_table = sheet = pivot_sheet = pivot_cache = pivot_table = pivot_data_field = None
            condition = hyperlink = None
            gc.collect()
            book.Close(False)
            book = None
            excel.Quit()
            excel = None

            capabilities = probe_excel_capabilities(staging_parent=root)
            if not capabilities.supports('connection_only') or not capabilities.supports('new_sheet'):
                raise AssertionError(f'Connection-only or worksheet-load probe failed on this build: {capabilities.reasons}')
            source = QuerySource(source_path, query_name, load_mode=LoadMode.CONNECTION_ONLY)
            report_source = QuerySource(report_source_path, report_query_name,
                load_mode=LoadMode.NEW_SHEET, sheet_name='PQM_Output', table_name='tblPqmReport')
            installer = ExcelQueryInstaller()
            source_path.write_text(first_formula, encoding='utf-8')
            report_source_path.write_text(report_first_formula, encoding='utf-8')
            installer.install(workbook_path, [source, report_source], ExistingQueryMode.REPLACE, capabilities=capabilities)

            # Add workbook formulas and a name that consume the query-backed
            # output. Replacing the M definition must leave these report links
            # attached to the same table and calculate against the new rows.
            report_excel = win32com.client.DispatchEx('Excel.Application')
            report_excel.Visible = False
            report_excel.DisplayAlerts = False
            report_book = report_excel.Workbooks.Open(str(workbook_path), UpdateLinks=0, ReadOnly=False,
                IgnoreReadOnlyRecommended=True, AddToMru=False)
            try:
                report = report_book.Worksheets.Add()
                report.Name = 'Report'
                report_book.Names.Add(Name='PQM_ReportTotal', RefersTo='=SUM(tblPqmReport[Value])')
                report.Range('A1').Formula = '=SUM(tblPqmReport[Value])'
                report.Range('A2').Formula = '=PQM_ReportTotal'
                report_book.Save()
            finally:
                report = None
                report_book.Close(False)
                report_book = None
                gc.collect()
                report_excel.Quit()
                report_excel = None

            source_path.write_text(changed_formula, encoding='utf-8')
            report_source_path.write_text(report_changed_formula, encoding='utf-8')
            privacy_before = read_workbook_privacy(workbook_path)
            if privacy_before != WorkbookPrivacy(True, None):
                raise AssertionError(f'Unexpected default disposable privacy settings: {privacy_before!r}')
            installer.install(workbook_path, [source, report_source], ExistingQueryMode.REPLACE, capabilities=capabilities)
            privacy_after = read_workbook_privacy(workbook_path)
            if privacy_after != privacy_before:
                raise AssertionError(f'Workbook Power Query privacy permissions changed: {privacy_before!r} -> {privacy_after!r}')

            inventory = ExcelQueryInspector().inspect(workbook_path)
            observed = {item.name: item for item in inventory}
            assert observed[query_name].formula == changed_formula, observed[query_name].formula
            assert observed[report_query_name].formula == report_changed_formula, observed[report_query_name].formula
            verify = win32com.client.DispatchEx('Excel.Application')
            verify.Visible = False
            verify.DisplayAlerts = False
            check = verify.Workbooks.Open(str(workbook_path), UpdateLinks=0, ReadOnly=True,
                                          IgnoreReadOnlyRecommended=True, AddToMru=False)
            try:
                saved_table = None
                for sheet_index in range(1, int(check.Worksheets.Count) + 1):
                    candidate_sheet = check.Worksheets.Item(sheet_index)
                    try:
                        saved_table = candidate_sheet.ListObjects.Item('tblKeep')
                        break
                    except Exception:
                        continue
                if saved_table is None:
                    raise AssertionError('Unrelated plain-range table tblKeep was not preserved.')
                assert int(saved_table.SourceType) == 1
                assert saved_table.Range.Address == '$A$1:$B$4'
                assert saved_table.DataBodyRange.Cells(1, 1).Value == 'Keep'
                assert str(saved_table.TableStyle) == 'TableStyleMedium2'
                assert bool(saved_table.AutoFilter.Filters.Item(1).On)
                assert str(saved_table.AutoFilter.Filters.Item(1).Criteria1) == '=Keep'
                keep_sheet = saved_table.Parent
                saved_chart = keep_sheet.ChartObjects().Item('PQM_KeepChart').Chart
                assert bool(saved_chart.HasTitle)
                assert str(saved_chart.ChartTitle.Text) == 'Preserve non-query chart'
                assert str(keep_sheet.Shapes.Item('PQM_KeepShape').AlternativeText) == 'PQM unrelated shape marker'
                saved_validation = keep_sheet.Range('F2').Validation
                assert int(saved_validation.Type) == 3
                assert str(saved_validation.Formula1) == '=$J$1:$J$2'
                assert bool(saved_validation.InCellDropdown)
                saved_conditions = keep_sheet.Range('G2:G4').FormatConditions
                assert int(saved_conditions.Count) == 1
                saved_condition = saved_conditions.Item(1)
                assert int(saved_condition.Type) == 1
                assert int(saved_condition.Operator) == 3
                assert str(saved_condition.Formula1) == '=1'
                assert int(saved_condition.Interior.Color) == 65535
                saved_link = keep_sheet.Hyperlinks.Item(1)
                assert str(saved_link.Address) == 'https://example.invalid/pqm-preserve'
                assert str(saved_link.TextToDisplay) == 'PQM preserved hyperlink'
                saved_pivot = check.Worksheets.Item('PQM_Pivot').PivotTables().Item('PQM_KeepPivot')
                assert str(saved_pivot.Name) == 'PQM_KeepPivot'
                assert int(saved_pivot.PivotFields('Category').Orientation) == 1
                assert str(saved_pivot.DataFields.Item(1).Name) == 'PQM Sum of Value'
                pivot_values = saved_pivot.DataBodyRange.Value
                if not isinstance(pivot_values, tuple):
                    pivot_values = ((pivot_values,),)
                elif pivot_values and not isinstance(pivot_values[0], tuple):
                    pivot_values = tuple((value,) for value in pivot_values)
                pivot_totals = [float(value) for row in pivot_values for value in row if value is not None]
                assert (len(pivot_totals) == 3 and pivot_totals[-1] == 61.0
                        and sum(pivot_totals[:-1]) == pivot_totals[-1]
                        and sorted(pivot_totals[:-1]) == [7.0, 54.0]), (
                    f'PivotTable data values were not preserved: {pivot_values!r}')
                assert str(check.Names.Item('PQM_KeepMarker').RefersTo).endswith('!$B$2')
                assert bool(check.ForceFullCalculation)
                assert str(saved_table.Parent.Range('D1').Formula) == '=SUM(B2:B4)'
                connection = check.Connections.Item(f'Query - {query_name}')
                assert not bool(connection.InModel)
                assert observed[query_name].load_state == 'connection_only', observed[query_name].load_state
                assert observed[report_query_name].load_state.startswith('worksheet:'), observed[report_query_name].load_state
                report_table = check.Worksheets.Item('PQM_Output').ListObjects.Item('tblPqmReport')
                assert int(report_table.DataBodyRange.Cells(1, 1).Value) == 30
                assert int(report_table.DataBodyRange.Cells(2, 1).Value) == 40
                report_sheet = check.Worksheets.Item('Report')
                assert str(report_sheet.Range('A1').Formula) == '=SUM(tblPqmReport[Value])'
                assert str(report_sheet.Range('A2').Formula) == '=PQM_ReportTotal'
                assert str(check.Names.Item('PQM_ReportTotal').RefersTo) == '=SUM(tblPqmReport[Value])'
                assert float(report_sheet.Range('A1').Value) == 70.0
                assert float(report_sheet.Range('A2').Value) == 70.0
            finally:
                saved_table = connection = report_table = report_sheet = keep_sheet = saved_chart = saved_pivot = None
                saved_validation = saved_conditions = saved_condition = saved_link = None
                check.Close(False)
                check = None
                gc.collect()
                verify.Quit()
                verify = None
            result = (f'PASS Excel {version} build {build}; formula replacement preserved ordinary table/style/filter, '
                      f'defined names, cell formulas, data validation, conditional formatting, hyperlink, '
                      f'full-calculation setting, chart and shape objects, '
                      f'PivotTable values 7/54/61, query-backed report table/formulas; '
                      f'connection_only and worksheet states verified; workbook privacy permissions preserved '
                      f'(FirewallEnabled={privacy_after.firewall_enabled}, WorkbookGroupType={privacy_after.workbook_group_type}); '
                      f'owned Excel processes exited')
    finally:
        win32com.client.DispatchEx = original_dispatch
        _wait_owned(owned_pids)
        pythoncom.CoUninitialize()
    print(result)


if __name__ == '__main__':
    main()
