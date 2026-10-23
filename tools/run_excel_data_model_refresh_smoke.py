"""Verify packaged selected-query Data Model refresh against a loopback-only source."""
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
import win32process
import win32com.client

ROOT = Path(__file__).resolve().parents[1]
DEFAULT_CLI = ROOT / 'dist' / 'PowerQueryManager' / 'PowerQueryManagerCLI.exe'
sys.path.insert(0, str(ROOT / 'app' / 'src'))

from pq_installer.core.excel_capabilities import probe_excel_capabilities
from pq_installer.core.excel_lifetime import open_owned_excel_process, owned_excel_pid, quit_owned_excel
from pq_installer.core.excel_queries import ExcelQueryInstaller
from pq_installer.core.models import ExistingQueryMode, LoadMode, QuerySource


def sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def excel_pids() -> set[int]:
    output = subprocess.run(
        ['tasklist.exe', '/FI', 'IMAGENAME eq EXCEL.EXE', '/FO', 'CSV', '/NH'],
        check=True, capture_output=True, text=True, encoding='mbcs', errors='replace').stdout
    return {int(row[1]) for row in csv.reader(output.splitlines())
            if len(row) > 1 and row[1].isdigit()}


def with_excel(callback):
    pythoncom.CoInitialize()
    excel = process = None
    try:
        excel = win32com.client.DispatchEx('Excel.Application')
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.AutomationSecurity = 3
        process = open_owned_excel_process(owned_excel_pid(excel), hwnd=int(excel.Hwnd))
        return callback(excel)
    finally:
        if excel is not None and process is not None:
            try:
                quit_owned_excel(excel, after_quit=pythoncom.CoUninitialize)
                excel = None
                pythoncom.CoInitialize()
            except Exception:
                pass
        elif excel is not None:
            try:
                excel.Quit()
            except Exception:
                pass
        excel = None
        if process is not None:
            process.close()
        pythoncom.CoUninitialize()


def main() -> int:
    import argparse

    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--cli-exe', type=Path, default=DEFAULT_CLI)
    parser.add_argument('--mode', choices=('background', 'synchronous'),
        help='Run one refresh mode in a fresh workbook instead of both sequentially.')
    args = parser.parse_args()
    if not args.cli_exe.is_file():
        parser.error(f'Packaged CLI does not exist: {args.cli_exe}')

    hits: list[str] = []
    class CsvHandler(BaseHTTPRequestHandler):
        def do_GET(self):
            hits.append(self.path)
            time.sleep(0.5)
            body = b'Value\r\n10\r\n20\r\n'
            try:
                self.send_response(200)
                self.send_header('Content-Type', 'text/csv; charset=utf-8')
                self.send_header('Content-Length', str(len(body)))
                self.end_headers()
                self.wfile.write(body)
            except (BrokenPipeError, ConnectionResetError):
                pass

        def log_message(self, *_args):
            pass

    server = ThreadingHTTPServer(('127.0.0.1', 0), CsvHandler)
    server.daemon_threads = True
    threading.Thread(target=server.serve_forever, daemon=True).start()
    selected_url = f'http://127.0.0.1:{server.server_port}/selected.csv'
    helper_url = f'http://127.0.0.1:{server.server_port}/helper.csv'
    selected_formula = (
        'let Source = Csv.Document(Web.Contents("' + selected_url +
        '"), [Delimiter=",", Encoding=65001, QuoteStyle=QuoteStyle.None]), '
        'Headers = Table.PromoteHeaders(Source) in Headers')
    helper_formula = (
        'let Source = Csv.Document(Web.Contents("' + helper_url +
        '"), [Delimiter=",", Encoding=65001, QuoteStyle=QuoteStyle.None]), '
        'Headers = Table.PromoteHeaders(Source) in Headers')
    baseline_pids = excel_pids()
    capability = probe_excel_capabilities()
    if not capability.supports('data_model'):
        raise RuntimeError(f'Data Model capability probe did not pass: {capability.reasons}')

    try:
        with tempfile.TemporaryDirectory(prefix='pqm-model-refresh-acceptance-') as folder:
            root = Path(folder)
            target = root / 'model-refresh.xlsx'
            source_path = root / 'model-refresh.m'
            helper_path = root / 'model-helper.m'
            source_path.write_text('let Source = #table({"Value"}, {{1}}) in Source', encoding='utf-8')
            helper_path.write_text('let Source = #table({"Value"}, {{99}}) in Source', encoding='utf-8')

            def seed(excel):
                if str(excel.Version) != capability.office_version:
                    raise RuntimeError('The disposable workbook was seeded under a different Excel version.')
                book = excel.Workbooks.Add()
                worksheet_count = int(book.Worksheets.Count)
                book.SaveAs(str(target), 51)
                book.Close(False)
                return worksheet_count

            baseline_worksheet_count = with_excel(seed)
            sources = [
                QuerySource(source_path, 'PQM_Model_Refresh', load_mode=LoadMode.DATA_MODEL),
                QuerySource(helper_path, 'PQM_Model_Helper', load_mode=LoadMode.DATA_MODEL),
            ]
            installed = ExcelQueryInstaller().install(
                target, sources, ExistingQueryMode.REPLACE, capabilities=capability)
            if len(installed) != 2 or any(item.status.value != 'Created' for item in installed):
                raise RuntimeError(f'Could not seed Data Model query: {installed!r}')

            def point_query_at_loopback(excel):
                book = excel.Workbooks.Open(str(target), UpdateLinks=0, ReadOnly=False,
                    IgnoreReadOnlyRecommended=True, AddToMru=False)
                try:
                    before = {str(book.Model.ModelTables.Item(index).SourceWorkbookConnection.Name):
                        int(book.Model.ModelTables.Item(index).RecordCount)
                        for index in range(1, int(book.Model.ModelTables.Count) + 1)}
                    book.Queries.Item('PQM_Model_Refresh').Formula = selected_formula
                    book.Queries.Item('PQM_Model_Helper').Formula = helper_formula
                    book.Save()
                    return before
                finally:
                    book.Close(False)

            seeded_count = with_excel(point_query_at_loopback)
            if seeded_count != {'Query - PQM_Model_Refresh': 1, 'Query - PQM_Model_Helper': 1}:
                raise RuntimeError(f'Initial Data Model row counts differ from one per query: {seeded_count!r}')

            reports = []
            modes = (args.mode,) if args.mode else ('background', 'synchronous')
            for mode in modes:
                before_hash = sha256(target)
                old_backups = {path.resolve() for path in root.glob('model-refresh_backup_*.xlsx')}
                old_hits = len(hits)
                command = [str(args.cli_exe), 'refresh', '--apply', '--all',
                    '--workbook', str(target), '--timeout', '30', '--confirm', 'REFRESH']
                if mode == 'synchronous':
                    command.append('--synchronous')
                env = os.environ.copy()
                env['LOCALAPPDATA'] = str(root / 'app-data')
                env['APPDATA'] = str(root / 'app-data')
                (root / 'app-data').mkdir(exist_ok=True)
                completed = subprocess.run(command, capture_output=True, text=True,
                    encoding='utf-8', errors='replace', env=env, timeout=120)
                if completed.returncode != 0:
                    raise RuntimeError(f'{mode} packaged refresh failed ({completed.returncode}):\n'
                        f'{completed.stdout}\n{completed.stderr}')
                start = completed.stdout.find('{')
                if start < 0:
                    raise RuntimeError(f'{mode} refresh returned no JSON report: {completed.stdout}')
                report = json.loads(completed.stdout[start:])
                if (len(report.get('results', ())) != 2 or any(
                        item.get('status') != 'refreshed' for item in report['results'])):
                    raise RuntimeError(f'{mode} refresh did not report both explicitly selected model queries refreshed: {report}')
                for item in report['results']:
                    message = item.get('message', '')
                    if ('synchronously' not in message.casefold() or
                            'background option' not in message.casefold()):
                        raise RuntimeError(f'{mode} report did not explain the Data Model API mode: {report}')
                new_backups = [path for path in root.glob('model-refresh_backup_*.xlsx')
                    if path.resolve() not in old_backups]
                if len(new_backups) != 1 or sha256(new_backups[0]) != before_hash:
                    raise RuntimeError(f'{mode} refresh did not create an exact pre-refresh backup.')
                if sha256(target) == before_hash:
                    raise RuntimeError(f'{mode} refresh did not commit the staged workbook.')
                if list(root.glob('.model-refresh.pqstage-*')):
                    raise RuntimeError(f'{mode} refresh left a staged workbook behind.')
                new_hits = hits[old_hits:]
                if (not any(item.startswith('/selected.csv') for item in new_hits) or
                        not any(item.startswith('/helper.csv') for item in new_hits) or
                        any(not item.startswith(('/selected.csv', '/helper.csv')) for item in new_hits)):
                    raise RuntimeError(f'{mode} refresh did not contact exactly the selected model sources: {new_hits!r}')

                def verify(excel):
                    book = excel.Workbooks.Open(str(target), UpdateLinks=0, ReadOnly=True,
                        IgnoreReadOnlyRecommended=True, AddToMru=False)
                    try:
                        if not bool(book.ReadOnly):
                            raise RuntimeError('Final verification workbook was not read-only.')
                        if str(book.Queries.Item('PQM_Model_Refresh').Formula) != selected_formula:
                            raise RuntimeError('M formula changed during Data Model refresh.')
                        if str(book.Queries.Item('PQM_Model_Helper').Formula) != helper_formula:
                            raise RuntimeError('Selected helper Data Model query formula changed during refresh.')
                        connection = book.Connections.Item('Query - PQM_Model_Refresh')
                        if not bool(connection.InModel):
                            raise RuntimeError('Selected query is no longer loaded to the Data Model.')
                        helper_connection = book.Connections.Item('Query - PQM_Model_Helper')
                        if not bool(helper_connection.InModel):
                            raise RuntimeError('Selected helper query is no longer loaded to the Data Model.')
                        matches = [book.Model.ModelTables.Item(index)
                            for index in range(1, int(book.Model.ModelTables.Count) + 1)
                            if str(book.Model.ModelTables.Item(index).SourceWorkbookConnection.Name).casefold()
                            == 'query - pqm_model_refresh']
                        if len(matches) != 1 or int(matches[0].RecordCount) != 2:
                            raise RuntimeError('Data Model table association or refreshed row count failed readback.')
                        helper_matches = [book.Model.ModelTables.Item(index)
                            for index in range(1, int(book.Model.ModelTables.Count) + 1)
                            if str(book.Model.ModelTables.Item(index).SourceWorkbookConnection.Name).casefold()
                            == 'query - pqm_model_helper']
                        if len(helper_matches) != 1 or int(helper_matches[0].RecordCount) != 2:
                            raise RuntimeError('Explicitly selected helper Data Model table did not refresh.')
                        if int(book.Worksheets.Count) != baseline_worksheet_count:
                            raise RuntimeError('Refresh unexpectedly added/removed workbook worksheets.')
                        return int(matches[0].RecordCount)
                    finally:
                        book.Close(False)

                count = with_excel(verify)
                if excel_pids() != baseline_pids:
                    raise RuntimeError(f'Excel process inventory differs from baseline: {excel_pids()} vs {baseline_pids}')
                reports.append({'mode': mode, 'status': 'refreshed', 'selected_model_rows': count,
                    'helper_model_rows': 2, 'new_loopback_requests': len(hits) - old_hits,
                    'backup_sha256': sha256(new_backups[0])})

            before_subset_hash = sha256(target)
            before_subset_hits = len(hits)
            subset_command = [str(args.cli_exe), 'refresh', '--apply', '--query',
                'PQM_Model_Refresh', '--workbook', str(target), '--timeout', '30',
                '--confirm', 'REFRESH']
            subset_env = os.environ.copy()
            subset_env['LOCALAPPDATA'] = str(root / 'app-data')
            subset_env['APPDATA'] = str(root / 'app-data')
            subset_result = subprocess.run(subset_command, capture_output=True, text=True,
                encoding='utf-8', errors='replace', env=subset_env, timeout=120)
            if subset_result.returncode == 0 or 'select all Data Model queries' not in (
                    subset_result.stdout + subset_result.stderr):
                raise RuntimeError(f'Partial Data Model selection was not refused clearly: {subset_result.stdout}{subset_result.stderr}')
            if len(hits) != before_subset_hits or sha256(target) != before_subset_hash:
                raise RuntimeError('Refused partial Data Model refresh contacted a source or changed the workbook.')
            if list(root.glob('.model-refresh.pqstage-*')):
                raise RuntimeError('Refused partial Data Model refresh left a stage behind.')
            if excel_pids() != baseline_pids:
                raise RuntimeError('Excel process inventory changed after partial Data Model refresh refusal.')

            print(json.dumps({'status': 'PASS', 'office_version': capability.office_version,
                'office_build': capability.build, 'reports': reports, 'request_paths': hits,
                'excel_pids_after': sorted(excel_pids()), 'excel_pids_before': sorted(baseline_pids)}, indent=2))
            return 0
    finally:
        server.shutdown()
        server.server_close()


if __name__ == '__main__':
    raise SystemExit(main())
