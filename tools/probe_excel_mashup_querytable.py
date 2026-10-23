"""Probe synchronous QueryTable refresh against a disposable Power Query Mashup connection."""
from __future__ import annotations

import json
import threading
import time
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

import pythoncom
import win32com.client
import win32process
import pywintypes

from pq_installer.core.excel_lifetime import (
    ExcelOperationTimeout,
    call_with_process_timeout,
    open_owned_excel_process,
    terminate_owned_excel,
)


def main() -> int:
    request_seen = threading.Event()
    release_request = threading.Event()

    class Handler(BaseHTTPRequestHandler):
        def do_GET(self):
            request_seen.set()
            release_request.wait(3)
            payload = b"Value\r\n1\r\n"
            try:
                self.send_response(200)
                self.send_header("Content-Length", str(len(payload)))
                self.end_headers()
                self.wfile.write(payload)
            except (BrokenPipeError, ConnectionResetError):
                pass

        def log_message(self, *_args):
            pass

    server = ThreadingHTTPServer(("127.0.0.1", 0), Handler)
    server.daemon_threads = True
    threading.Thread(target=server.serve_forever, daemon=True).start()
    excel = workbook = None
    process_id = None
    com_initialized = False
    timed_out = False
    try:
        pythoncom.CoInitialize()
        com_initialized = True
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.AutomationSecurity = 3
        process_id = int(win32process.GetWindowThreadProcessId(int(excel.Hwnd))[1])
        workbook = excel.Workbooks.Add()
        name = "PQM_QueryTable_Probe"
        url = f"http://127.0.0.1:{server.server_port}/delayed.csv"
        formula = (
            "let Source = Csv.Document(Web.Contents("
            f'"{url}"), [Delimiter=",", Encoding=65001, QuoteStyle=QuoteStyle.None]), '
            "Headers = Table.PromoteHeaders(Source) in Headers"
        )
        workbook.Queries.Add(name, formula)
        connection_string = (
            "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;"
            f'Location={name};Extended Properties=""'
        )
        workbook.Connections.Add2(
            f"Query - {name}", "Disposable Mashup QueryTable probe",
            connection_string, f"SELECT * FROM [{name}]", 2, False, False,
        )
        existing_connection = workbook.Connections.Item(f"Query - {name}")
        existing_connection_string = str(existing_connection.OLEDBConnection.Connection)
        existing_command_text = str(existing_connection.OLEDBConnection.CommandText)
        initial_connection_names = {
            str(workbook.Connections.Item(i).Name).casefold()
            for i in range(1, workbook.Connections.Count + 1)
        }
        original_active_sheet = workbook.ActiveSheet
        sheet = workbook.Worksheets.Add(
            After=workbook.Worksheets.Item(workbook.Worksheets.Count)
        )
        sheet.Visible = 2
        query_table = sheet.QueryTables.Add(
            connection_string, sheet.Range("$A$1"), f"SELECT * FROM [{name}]"
        )
        query_table.BackgroundQuery = True
        owned = open_owned_excel_process(process_id, hwnd=int(excel.Hwnd))
        started = time.monotonic()
        try:
            result = call_with_process_timeout(
                process_id,
                lambda: query_table.Refresh(True),
                10,
                operation="asynchronous Mashup QueryTable refresh probe",
                terminate=terminate_owned_excel,
                owned_process=owned,
            )
        except ExcelOperationTimeout as exc:
            timed_out = True
            result = f"{type(exc).__name__}: {exc}"
        finally:
            owned.close()
        refreshing_at_return = None if timed_out else bool(query_table.Refreshing)
        deadline = time.monotonic() + 10
        busy_retries = 0
        while not timed_out and time.monotonic() < deadline:
            try:
                refreshing = bool(query_table.Refreshing)
            except pywintypes.com_error as exc:
                if int(exc.args[0]) & 0xFFFFFFFF not in (0x80010001, 0x8001010A):
                    raise
                busy_retries += 1
                pythoncom.PumpWaitingMessages()
                time.sleep(0.05)
                continue
            if not refreshing:
                break
            time.sleep(0.05)
        elapsed = time.monotonic() - started
        refreshing_after = None if timed_out else bool(query_table.Refreshing)
        data_rows = None if timed_out else int(sheet.UsedRange.Rows.Count)
        querytable_connection_name = None if timed_out else str(query_table.WorkbookConnection.Name)
        workbook_connection_names = None if timed_out else [
            str(workbook.Connections.Item(i).Name)
            for i in range(1, workbook.Connections.Count + 1)
        ]
        post_cleanup_connection_names = None
        if not timed_out:
            query_table.Delete()
            sheet.Visible = -1
            original_active_sheet.Activate()
            sheet.Delete()
            if querytable_connection_name.casefold() not in initial_connection_names:
                workbook.Connections.Item(querytable_connection_name).Delete()
            post_cleanup_connection_names = [
                str(workbook.Connections.Item(i).Name)
                for i in range(1, workbook.Connections.Count + 1)
            ]
        print(json.dumps({
            "office_version": str(excel.Version) if not timed_out else None,
            "build": str(excel.Build) if not timed_out else None,
            "excel_pid": process_id,
            "querytable_refresh_result": repr(result),
            "elapsed_seconds": round(elapsed, 3),
            "loopback_request_received": request_seen.is_set(),
            "timed_out": timed_out,
            "refreshing_at_return": refreshing_at_return,
            "busy_retries": busy_retries,
            "refreshing_after": refreshing_after,
            "data_rows": data_rows,
            "querytable_connection_name": querytable_connection_name,
            "workbook_connection_names": workbook_connection_names,
            "post_cleanup_connection_names": post_cleanup_connection_names,
            "connection_string_roundtrip": existing_connection_string,
            "command_text_roundtrip": existing_command_text,
            "external_network_used": False,
        }, indent=2))
        return 0 if request_seen.is_set() and not timed_out else 2
    finally:
        release_request.set()
        if excel is not None and not timed_out:
            try:
                if workbook is not None:
                    workbook.Close(False)
                excel.Quit()
            except Exception:
                pass
        if com_initialized:
            pythoncom.CoUninitialize()
        server.shutdown()
        server.server_close()


if __name__ == "__main__":
    raise SystemExit(main())
