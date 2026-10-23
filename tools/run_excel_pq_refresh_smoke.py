"""Exercise successful packaged Power Query refreshes against a loopback CSV source."""
from __future__ import annotations

import csv
import hashlib
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
import json
import os
from pathlib import Path
import subprocess
import sys
import tempfile
import threading
import time

import pythoncom
import win32api
import win32con
import win32event
import win32process
import win32com.client

ROOT = Path(__file__).resolve().parents[1]
PACKAGED_CLI = ROOT / 'dist' / 'PowerQueryManager' / 'PowerQueryManagerCLI.exe'
sys.path.insert(0, str(ROOT / 'app' / 'src'))
from pq_installer.core.excel_lifetime import open_owned_excel_process  # noqa: E402


def _sha(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def _excel_pids():
    result = subprocess.run(
        ['tasklist.exe', '/FI', 'IMAGENAME eq EXCEL.EXE', '/FO', 'CSV', '/NH'],
        check=True, capture_output=True, text=True, encoding='mbcs', errors='replace')
    return {int(row[1]) for row in csv.reader(result.stdout.splitlines())
            if len(row) >= 2 and row[1].isdigit()}


def _wait_for_exit(process, timeout=15):
    if not process.wait(timeout * 1000):
        process.terminate()
        if not process.wait(5000):
            raise RuntimeError(f'Owned Excel PID {process.pid} did not exit after termination.')


def _read_workbook(path, expected_formula, baseline_sheet_count, *, worksheet_loaded=False):
    pythoncom.CoInitialize()
    excel = None
    process = None
    try:
        excel = win32com.client.DispatchEx('Excel.Application')
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.AutomationSecurity = 3
        pid = int(win32process.GetWindowThreadProcessId(int(excel.Hwnd))[1])
        process = open_owned_excel_process(pid, hwnd=int(excel.Hwnd))
        workbook = excel.Workbooks.Open(
            str(path), UpdateLinks=0, ReadOnly=True, Password='',
            WriteResPassword='', IgnoreReadOnlyRecommended=True, AddToMru=False)
        try:
            formula = str(workbook.Queries.Item('PQM_Loopback_Success').Formula)
            connections = [str(workbook.Connections.Item(i).Name)
                           for i in range(1, workbook.Connections.Count + 1)]
            sheet_count = int(workbook.Worksheets.Count)
            if formula != expected_formula:
                raise RuntimeError('Read-only reopen found a changed M formula.')
            if connections != ['Query - PQM_Loopback_Success']:
                raise RuntimeError(f'Temporary refresh connection was not removed: {connections}')
            if sheet_count != baseline_sheet_count:
                raise RuntimeError(f'Temporary refresh sheet was not removed: {sheet_count}')
            tables = []
            for index in range(1, sheet_count + 1):
                sheet = workbook.Worksheets.Item(index)
                for table_index in range(1, int(sheet.ListObjects.Count) + 1):
                    tables.append(sheet.ListObjects.Item(table_index))
            if worksheet_loaded:
                if len(tables) != 1:
                    raise RuntimeError(f'Worksheet load table count changed: {len(tables)}')
                table = tables[0]
                if str(table.QueryTable.WorkbookConnection.Name) != 'Query - PQM_Loopback_Success':
                    raise RuntimeError('Worksheet table no longer maps to its original Power Query connection.')
                refreshed_value = str(table.Range.Cells(2, 1).Value).strip()
                if refreshed_value != '42':
                    raise RuntimeError(f'Worksheet refresh did not load the loopback result: {refreshed_value!r}')
            elif tables:
                raise RuntimeError(f'Connection-only refresh created an unexpected table: {len(tables)}')
        finally:
            workbook.Close(False)
        excel.Quit()
        excel = None
        return pid, connections, sheet_count
    finally:
        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                pass
        if process is not None:
            _wait_for_exit(process)
            process.close()
        pythoncom.CoUninitialize()


def _seed_worksheet_workbook(path, formula):
    excel = None
    workbook = None
    process = None
    try:
        excel = win32com.client.DispatchEx('Excel.Application')
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.AutomationSecurity = 3
        pid = int(win32process.GetWindowThreadProcessId(int(excel.Hwnd))[1])
        process = open_owned_excel_process(pid, hwnd=int(excel.Hwnd))
        workbook = excel.Workbooks.Add()
        sheet_count = int(workbook.Worksheets.Count)
        name = 'PQM_Loopback_Success'
        workbook.Queries.Add(name, 'let Source = #table({"Value"}, {{1}}) in Source')
        connection_string = (
            'OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;'
            f'Location={name};Extended Properties=""')
        table = workbook.Worksheets.Item(1).ListObjects.Add(
            3, connection_string, True, 1, workbook.Worksheets.Item(1).Range('$D$4'))
        table.Name = 'PQM_Loopback_Result'
        table.QueryTable.CommandType = 2
        table.QueryTable.CommandText = f'SELECT * FROM [{name}]'
        table.QueryTable.BackgroundQuery = False
        if table.QueryTable.Refresh(False) is not True:
            raise RuntimeError('Could not create the static worksheet-load seed.')
        table.QueryTable.WorkbookConnection.Name = f'Query - {name}'
        workbook.Queries.Item(name).Formula = formula
        workbook.SaveAs(str(path), 51)
        workbook.Close(False)
        workbook = None
        excel.Quit()
        excel = None
        _wait_for_exit(process)
        return sheet_count
    finally:
        if workbook is not None:
            try:
                workbook.Close(False)
            except Exception:
                pass
        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                pass
        if process is not None:
            if not process.wait(0):
                _wait_for_exit(process)
            process.close()


def main():
    import argparse

    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--cli-exe', type=Path, default=PACKAGED_CLI,
                        help='Packaged CLI executable (defaults to dist).')
    args = parser.parse_args()
    if not args.cli_exe.is_file():
        parser.error(f'Packaged CLI does not exist: {args.cli_exe}')

    hits = []
    hits_lock = threading.Lock()

    class CsvHandler(BaseHTTPRequestHandler):
        def do_GET(self):
            with hits_lock:
                hits.append(self.path)
            time.sleep(0.5)
            value = '42' if self.path == '/worksheet.csv' else '1'
            payload = f'Value\r\n{value}\r\n'.encode('ascii')
            self.send_response(200)
            self.send_header('Content-Type', 'text/csv; charset=utf-8')
            self.send_header('Content-Length', str(len(payload)))
            self.end_headers()
            try:
                self.wfile.write(payload)
            except (BrokenPipeError, ConnectionResetError):
                pass

        def log_message(self, *_args):
            pass

    server = ThreadingHTTPServer(('127.0.0.1', 0), CsvHandler)
    server.daemon_threads = True
    threading.Thread(target=server.serve_forever, daemon=True).start()
    baseline_pids = _excel_pids()
    seed_excel = None
    seed_process = None
    pythoncom.CoInitialize()
    try:
        with tempfile.TemporaryDirectory(prefix='pqm-pq-refresh-') as folder:
            root = Path(folder)
            target = root / 'loopback-refresh.xlsx'
            app_data = root / 'app-data'
            app_data.mkdir()
            seed_excel = win32com.client.DispatchEx('Excel.Application')
            seed_excel.Visible = False
            seed_excel.DisplayAlerts = False
            seed_excel.AutomationSecurity = 3
            office_version, office_build = str(seed_excel.Version), str(seed_excel.Build)
            seed_pid = int(win32process.GetWindowThreadProcessId(int(seed_excel.Hwnd))[1])
            seed_process = open_owned_excel_process(seed_pid, hwnd=int(seed_excel.Hwnd))
            workbook = seed_excel.Workbooks.Add()
            sheet_count = int(workbook.Worksheets.Count)
            query_name = 'PQM_Loopback_Success'
            url = f'http://127.0.0.1:{server.server_port}/success.csv'
            formula = (
                'let Source = Csv.Document(Web.Contents('
                f'"{url}"), [Delimiter=",", Encoding=65001, QuoteStyle=QuoteStyle.None]), '
                'Headers = Table.PromoteHeaders(Source) in Headers')
            workbook.Queries.Add(query_name, formula)
            connection_string = (
                'OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;'
                f'Location={query_name};Extended Properties=""')
            workbook.Connections.Add2(
                f'Query - {query_name}', 'Disposable loopback refresh smoke',
                connection_string, f'SELECT * FROM [{query_name}]', 2, False, False)
            workbook.SaveAs(str(target), 51)
            workbook.Close(False)
            workbook = None
            seed_excel.Quit()
            seed_excel = None
            _wait_for_exit(seed_process)
            seed_process.close()
            seed_process = None
            if _excel_pids() != baseline_pids:
                raise RuntimeError('Excel process set did not return to baseline after workbook seeding.')

            reports = []
            verifier_pids = []
            for mode in ('background', 'synchronous'):
                pre_refresh_hash = _sha(target)
                previous_backups = {item.resolve() for item in root.glob('loopback-refresh_backup_*.xlsx')}
                command = [
                    str(args.cli_exe), 'refresh', '--apply', '--query', query_name,
                    '--workbook', str(target), '--timeout', '30', '--confirm', 'REFRESH',
                ]
                if mode == 'synchronous':
                    command.append('--synchronous')
                env = os.environ.copy()
                env['LOCALAPPDATA'] = str(app_data)
                env['APPDATA'] = str(app_data)
                completed = subprocess.run(
                    command, capture_output=True, text=True, encoding='utf-8',
                    errors='replace', env=env, timeout=90)
                if completed.returncode != 0:
                    raise RuntimeError(
                        f'{mode} refresh returned {completed.returncode}:\n'
                        f'{completed.stdout}\n{completed.stderr}')
                json_start = completed.stdout.find('{')
                if json_start < 0:
                    raise RuntimeError(f'{mode} refresh returned no JSON report: {completed.stdout}')
                report = json.loads(completed.stdout[json_start:])
                if len(report.get('results', [])) != 1 or report['results'][0].get('status') != 'refreshed':
                    raise RuntimeError(f'{mode} refresh did not confirm success: {report}')
                new_backups = [item for item in root.glob('loopback-refresh_backup_*.xlsx')
                               if item.resolve() not in previous_backups]
                if len(new_backups) != 1 or _sha(new_backups[0]) != pre_refresh_hash:
                    raise RuntimeError(f'{mode} refresh did not preserve exactly one pre-refresh-hash backup.')
                if _sha(target) == pre_refresh_hash:
                    raise RuntimeError(f'{mode} refresh did not commit the staged workbook.')
                if list(root.glob('.loopback-refresh.pqstage-*')):
                    raise RuntimeError(f'{mode} refresh left a stage behind.')
                pid, connection_names, final_sheet_count = _read_workbook(
                    target, formula, sheet_count)
                verifier_pids.append(pid)
                reports.append({'mode': mode, 'status': report['results'][0]['status'],
                                'load_mode': 'connection_only',
                                'backup_sha256': _sha(new_backups[0]),
                                'connections_after_refresh': connection_names,
                                'worksheets_after_refresh': final_sheet_count})

            worksheet_target = root / 'worksheet-refresh.xlsx'
            worksheet_url = f'http://127.0.0.1:{server.server_port}/worksheet.csv'
            worksheet_formula = (
                'let Source = Csv.Document(Web.Contents('
                f'"{worksheet_url}"), [Delimiter=",", Encoding=65001, QuoteStyle=QuoteStyle.None]), '
                'Headers = Table.PromoteHeaders(Source) in Headers')
            worksheet_count = _seed_worksheet_workbook(worksheet_target, worksheet_formula)
            worksheet_hashes = []
            for mode in ('background', 'synchronous'):
                pre_refresh_hash = _sha(worksheet_target)
                previous_backups = {item.resolve() for item in root.glob('worksheet-refresh_backup_*.xlsx')}
                command = [
                    str(args.cli_exe), 'refresh', '--apply', '--query', 'PQM_Loopback_Success',
                    '--workbook', str(worksheet_target), '--timeout', '30', '--confirm', 'REFRESH',
                ]
                if mode == 'synchronous':
                    command.append('--synchronous')
                env = os.environ.copy()
                env['LOCALAPPDATA'] = str(app_data)
                env['APPDATA'] = str(app_data)
                completed = subprocess.run(command, capture_output=True, text=True,
                    encoding='utf-8', errors='replace', env=env, timeout=90)
                if completed.returncode != 0:
                    raise RuntimeError(f'Worksheet-loaded {mode} refresh failed: {completed.stdout}\n{completed.stderr}')
                json_start = completed.stdout.find('{')
                if json_start < 0:
                    raise RuntimeError(f'Worksheet-loaded {mode} refresh returned no JSON report.')
                report = json.loads(completed.stdout[json_start:])
                if len(report.get('results', [])) != 1 or report['results'][0].get('status') != 'refreshed':
                    raise RuntimeError(f'Worksheet-loaded {mode} refresh was not confirmed: {report}')
                new_backups = [item for item in root.glob('worksheet-refresh_backup_*.xlsx')
                               if item.resolve() not in previous_backups]
                if len(new_backups) != 1 or _sha(new_backups[0]) != pre_refresh_hash:
                    raise RuntimeError(f'Worksheet-loaded {mode} backup did not match the pre-refresh hash.')
                if list(root.glob('.worksheet-refresh.pqstage-*')):
                    raise RuntimeError(f'Worksheet-loaded {mode} refresh left a stage behind.')
                pid, connection_names, final_sheet_count = _read_workbook(
                    worksheet_target, worksheet_formula, worksheet_count, worksheet_loaded=True)
                verifier_pids.append(pid)
                worksheet_hashes.append(_sha(new_backups[0]))
                reports.append({'mode': mode, 'load_mode': 'worksheet',
                                'status': report['results'][0]['status'],
                                'backup_sha256': worksheet_hashes[-1],
                                'connections_after_refresh': connection_names,
                                'worksheets_after_refresh': final_sheet_count,
                                'result_value': 42})

            if hits.count('/success.csv') < 2 or hits.count('/worksheet.csv') < 2:
                raise RuntimeError(f'Expected local refresh requests in both load shapes/modes, observed: {hits}')
            if _excel_pids() != baseline_pids:
                raise RuntimeError(
                    f'Excel process set changed: before={sorted(baseline_pids)}, '
                    f'after={sorted(_excel_pids())}')
            print(json.dumps({
                'status': 'PASS',
                'method': 'packaged Power Query refresh using temporary and existing worksheet QueryTables',
                'office_version': office_version,
                'build': office_build,
                'modes': reports,
                'loopback_requests': len(hits),
                'source_formula_preserved': True,
                'temporary_connections_and_sheets_removed': True,
                'worksheet_load_value_updated': True,
                'excel_process_ids_before_and_after': sorted(baseline_pids),
                'read_only_verifier_pids': verifier_pids,
                'external_network_used': False,
            }, indent=2))
            return 0
    finally:
        if seed_excel is not None:
            try:
                seed_excel.Quit()
            except Exception:
                pass
        if seed_process is not None:
            try:
                _wait_for_exit(seed_process)
            finally:
                seed_process.close()
        pythoncom.CoUninitialize()
        server.shutdown()
        server.server_close()


if __name__ == '__main__':
    raise SystemExit(main())
