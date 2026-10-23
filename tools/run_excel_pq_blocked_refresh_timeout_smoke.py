"""Exercise the production Power Query refresh timeout with a loopback-only source."""
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
import win32com.client
import win32process

ROOT = Path(__file__).resolve().parents[1]
PACKAGED_CLI = ROOT / 'dist' / 'PowerQueryManager' / 'PowerQueryManagerCLI.exe'


def _sha(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def _excel_pids():
    result = subprocess.run(
        ['tasklist.exe', '/FI', 'IMAGENAME eq EXCEL.EXE', '/FO', 'CSV', '/NH'],
        check=True, capture_output=True, text=True, encoding='mbcs', errors='replace')
    return {int(row[1]) for row in csv.reader(result.stdout.splitlines())
            if len(row) >= 2 and row[1].isdigit()}


def _wait_for_pid_exit(pid, timeout=15):
    deadline = time.monotonic() + timeout
    while pid in _excel_pids() and time.monotonic() < deadline:
        time.sleep(0.1)
    return pid not in _excel_pids()


def main():
    import argparse

    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--cli-exe', type=Path, default=PACKAGED_CLI,
                        help='Packaged CLI executable to exercise (defaults to dist).')
    args = parser.parse_args()
    if not args.cli_exe.is_file():
        parser.error(f'Packaged CLI does not exist: {args.cli_exe}')

    request_started = threading.Event()
    release_request = threading.Event()

    class DelayedCsvHandler(BaseHTTPRequestHandler):
        def do_GET(self):
            request_started.set()
            release_request.wait(30)
            payload = b'Value\r\n1\r\n'
            try:
                self.send_response(200)
                self.send_header('Content-Type', 'text/csv; charset=utf-8')
                self.send_header('Content-Length', str(len(payload)))
                self.end_headers()
                self.wfile.write(payload)
            except (BrokenPipeError, ConnectionResetError):
                pass

        def log_message(self, _format, *_args):
            pass

    server = ThreadingHTTPServer(('127.0.0.1', 0), DelayedCsvHandler)
    server.daemon_threads = True
    server_thread = threading.Thread(target=server.serve_forever, daemon=True)
    server_thread.start()
    baseline_pids = _excel_pids()
    com_initialized = False
    seed_excel = None
    seed_pid = None
    try:
        pythoncom.CoInitialize()
        com_initialized = True
        with tempfile.TemporaryDirectory(prefix='pqm-pq-timeout-') as folder:
            root = Path(folder)
            target = root / 'loopback-refresh.xlsx'
            local_app_data = root / 'local-app-data'
            local_app_data.mkdir()

            seed_excel = win32com.client.DispatchEx('Excel.Application')
            seed_excel.Visible = False
            seed_excel.DisplayAlerts = False
            seed_excel.AutomationSecurity = 3
            office_version, office_build = str(seed_excel.Version), str(seed_excel.Build)
            seed_pid = int(win32process.GetWindowThreadProcessId(int(seed_excel.Hwnd))[1])
            workbook = seed_excel.Workbooks.Add()
            query_name = 'PQM_Loopback_Timeout'
            url = f'http://127.0.0.1:{server.server_port}/delayed.csv'
            formula = (
                'let Source = Csv.Document(Web.Contents('
                f'"{url}"), [Delimiter=",", Encoding=65001, QuoteStyle=QuoteStyle.None]), '
                'Headers = Table.PromoteHeaders(Source) in Headers')
            workbook.Queries.Add(query_name, formula)
            connection_string = (
                'OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;'
                f'Location={query_name};Extended Properties=""')
            workbook.Connections.Add2(
                f'Query - {query_name}', 'Disposable loopback timeout smoke',
                connection_string, f'SELECT * FROM [{query_name}]', 2, False, False)
            workbook.SaveAs(str(target), 51)
            workbook.Close(False)
            workbook = None
            seed_excel.Quit()
            seed_excel = None

            if not _wait_for_pid_exit(seed_pid):
                raise RuntimeError(f'Disposable workbook seed Excel process {seed_pid} remained alive.')
            if _excel_pids() != baseline_pids:
                raise RuntimeError('Excel process set did not return to baseline after seeding the workbook.')

            original_hash = _sha(target)
            command = [
                str(args.cli_exe), 'refresh', '--apply', '--synchronous',
                '--query', query_name, '--workbook', str(target),
                '--timeout', '3', '--confirm', 'REFRESH',
            ]
            env = os.environ.copy()
            env['LOCALAPPDATA'] = str(local_app_data)
            env['APPDATA'] = str(local_app_data)
            completed = subprocess.run(
                command, capture_output=True, text=True, encoding='utf-8',
                errors='replace', env=env, timeout=90)

            if completed.returncode != 5:
                raise RuntimeError(
                    f'Expected an explicit failed-refresh exit code 5, got {completed.returncode}; '
                    f'loopback_request_received={request_started.is_set()}.\n'
                    f'stdout:\n{completed.stdout}\nstderr:\n{completed.stderr}')
            if not request_started.wait(2):
                raise RuntimeError(
                    'Power Query did not reach the delayed loopback source before the CLI returned.\n'
                    f'stdout:\n{completed.stdout}\nstderr:\n{completed.stderr}')
            if _sha(target) != original_hash:
                raise RuntimeError('A timed-out refresh changed the original workbook bytes.')
            backups = list(root.glob('loopback-refresh_backup_*.xlsx'))
            if len(backups) != 1 or _sha(backups[0]) != original_hash:
                raise RuntimeError('Timed-out refresh did not retain exactly one original-hash backup.')
            stages = list(root.glob('.loopback-refresh.pqstage-*'))
            if stages:
                raise RuntimeError(f'Timed-out refresh left a staged workbook behind: {stages}')
            remaining_pids = _excel_pids()
            if remaining_pids != baseline_pids:
                raise RuntimeError(
                    f'Excel process set changed outside the test: '
                    f'before={sorted(baseline_pids)}, after={sorted(remaining_pids)}')

            output_start = completed.stdout.find('{')
            if output_start < 0:
                raise RuntimeError(f'CLI did not report structured timeout details: {completed.stdout}')
            report = json.loads(completed.stdout[output_start:])
            if 'timeout' not in report.get('error', '').casefold():
                raise RuntimeError(f'CLI did not report a timeout: {report}')

            print(json.dumps({
                'status': 'PASS',
                'method': 'packaged Power Query synchronous refresh blocked on delayed loopback Web.Contents',
                'office_version': office_version,
                'build': office_build,
                'timeout_seconds': 3,
                'cli_exit_code': completed.returncode,
                'loopback_request_received': request_started.is_set(),
                'timeout_reported': True,
                'original_hash_preserved': True,
                'backup_hash_verified': True,
                'stage_removed': True,
                'excel_process_ids_before_and_after': sorted(baseline_pids),
                'external_network_used': False,
            }, indent=2))
            return 0
    finally:
        if seed_excel is not None:
            try:
                seed_excel.Quit()
            except Exception:
                pass
        release_request.set()
        server.shutdown()
        server.server_close()
        server_thread.join(timeout=2)
        if com_initialized:
            pythoncom.CoUninitialize()


if __name__ == '__main__':
    raise SystemExit(main())
