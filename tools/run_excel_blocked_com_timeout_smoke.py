"""Verify timeout cleanup while an owned Excel COM refresh is blocked on loopback."""
from __future__ import annotations

import csv
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
import platform
import subprocess
import threading
import time
from pathlib import Path
import sys

import pythoncom
import win32api
import win32con
import win32event
import win32process
import win32com.client

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / 'app' / 'src'))
from pq_installer.core.excel_lifetime import (  # noqa: E402
    ExcelOperationTimeout,
    call_with_timeout,
    open_owned_excel_process,
)


def _excel_pids():
    result = subprocess.run(
        ['tasklist.exe', '/FI', 'IMAGENAME eq EXCEL.EXE', '/FO', 'CSV', '/NH'],
        check=True, capture_output=True, text=True, encoding='mbcs', errors='replace')
    return {int(row[1]) for row in csv.reader(result.stdout.splitlines())
            if len(row) >= 2 and row[1].isdigit()}


def main():
    request_started = threading.Event()
    release_request = threading.Event()

    class DelayedCsvHandler(BaseHTTPRequestHandler):
        def do_GET(self):
            request_started.set()
            release_request.wait(20)
            payload = b'Value\r\n1\r\n'
            try:
                self.send_response(200)
                self.send_header('Content-Type', 'text/csv; charset=utf-8')
                self.send_header('Content-Length', str(len(payload)))
                self.end_headers()
                self.wfile.write(payload)
            except (BrokenPipeError, ConnectionResetError):
                # Excel is expected to be terminated before the delayed response.
                pass

        def log_message(self, _format, *_args):
            pass

    server = ThreadingHTTPServer(('127.0.0.1', 0), DelayedCsvHandler)
    server.daemon_threads = True
    server_thread = threading.Thread(target=server.serve_forever, daemon=True)
    server_thread.start()

    baseline_pids = _excel_pids()
    com_initialized = False
    excel = None
    workbook = None
    sheet = None
    query_table = None
    owned_process = None
    version = build = 'unknown'
    pid = None
    outcome = None
    try:
        pythoncom.CoInitialize()
        com_initialized = True
        excel = win32com.client.DispatchEx('Excel.Application')
        excel.Visible = False
        excel.DisplayAlerts = False
        version, build = str(excel.Version), str(excel.Build)
        hwnd = int(excel.Hwnd)
        _thread_id, pid = win32process.GetWindowThreadProcessId(hwnd)
        pid = int(pid)
        owned_process = open_owned_excel_process(pid, hwnd=hwnd)

        workbook = excel.Workbooks.Add()
        sheet = workbook.Worksheets(1)
        url = f'http://127.0.0.1:{server.server_port}/delayed.csv'
        query_table = sheet.QueryTables.Add(
            Connection=f'URL;{url}', Destination=sheet.Range('A1'))
        query_table.BackgroundQuery = False

        try:
            call_with_timeout(
                excel, lambda: query_table.Refresh(False), 3,
                enabled=True, operation='disposable loopback Excel COM refresh')
        except ExcelOperationTimeout as exc:
            outcome = exc
            if not exc.process_exited:
                raise RuntimeError(f'Owned Excel process exit was not confirmed: {exc}') from exc
        else:
            raise RuntimeError('The delayed loopback refresh returned before the timeout.')

        if not request_started.wait(2):
            raise RuntimeError('Excel never reached the local delayed HTTP fixture.')
        if win32event.WaitForSingleObject(owned_process.native_handle, 0) != win32event.WAIT_OBJECT_0:
            raise RuntimeError('The captured DispatchEx process handle did not signal exit.')
        remaining_pids = _excel_pids()
        if remaining_pids != baseline_pids:
            raise RuntimeError(
                f'Excel process set changed outside the captured process: '
                f'before={sorted(baseline_pids)}, after={sorted(remaining_pids)}')

        print({
            'status': 'PASS',
            'method': 'Excel QueryTable.Refresh(False) blocked on a delayed loopback CSV response',
            'office_version': version,
            'build': build,
            'python': platform.python_version(),
            'excel_pid': pid,
            'timeout_seconds': outcome.timeout_seconds,
            'process_exit_confirmed': outcome.process_exited,
            'loopback_request_received': request_started.is_set(),
            'excel_processes_before_and_after': sorted(baseline_pids),
            'external_network_used': False,
        })
        return 0
    finally:
        # On a setup failure before the timed call, close only our own instance.
        if excel is not None and outcome is None:
            try:
                excel.Quit()
            except Exception:
                pass
        query_table = None
        sheet = None
        workbook = None
        excel = None
        if owned_process is not None:
            try:
                if not owned_process.wait(5):
                    owned_process.terminate()
                    owned_process.wait(5)
            finally:
                owned_process.close()
        release_request.set()
        server.shutdown()
        server.server_close()
        server_thread.join(timeout=2)
        if com_initialized:
            pythoncom.CoUninitialize()


if __name__ == '__main__':
    raise SystemExit(main())
