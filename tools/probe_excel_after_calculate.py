"""Probe whether Excel AfterCalculate follows completion of a delayed local Mashup request."""
from __future__ import annotations

import json
import threading
import time
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

import pythoncom
import win32com.client
import win32process


def main() -> int:
    hit = threading.Event()
    release = threading.Event()
    after_calculate = threading.Event()

    class Handler(BaseHTTPRequestHandler):
        def do_GET(self):
            hit.set()
            release.wait(2)
            payload = b"Value\r\n1\r\n"
            self.send_response(200)
            self.send_header("Content-Length", str(len(payload)))
            self.end_headers()
            self.wfile.write(payload)

        def log_message(self, *_args):
            pass

    class Events:
        def OnAfterCalculate(self):
            after_calculate.set()

    server = ThreadingHTTPServer(("127.0.0.1", 0), Handler)
    server.daemon_threads = True
    threading.Thread(target=server.serve_forever, daemon=True).start()
    excel = None
    pythoncom.CoInitialize()
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.AutomationSecurity = 3
        process_id = int(win32process.GetWindowThreadProcessId(int(excel.Hwnd))[1])
        _sink = win32com.client.WithEvents(excel, Events)
        workbook = excel.Workbooks.Add()
        name = "PQM_AfterCalculate"
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
            f"Query - {name}", "Local delayed source", connection_string,
            f"SELECT * FROM [{name}]", 2, False, False,
        )
        connection = workbook.Connections.Item(f"Query - {name}")
        connection.OLEDBConnection.BackgroundQuery = False
        try:
            refresh_date_before = str(connection.OLEDBConnection.RefreshDate)
        except Exception as exc:
            refresh_date_before = f"{type(exc).__name__}: {exc}"
        started = time.monotonic()
        refresh_result = connection.Refresh()
        returned = time.monotonic()
        saw_request = hit.wait(3)
        deadline = time.monotonic() + 10
        while not after_calculate.is_set() and time.monotonic() < deadline:
            pythoncom.PumpWaitingMessages()
            time.sleep(0.02)
        try:
            refresh_date_after = str(connection.OLEDBConnection.RefreshDate)
        except Exception as exc:
            refresh_date_after = f"{type(exc).__name__}: {exc}"
        try:
            refreshing_after = bool(connection.OLEDBConnection.Refreshing)
        except Exception as exc:
            refreshing_after = f"{type(exc).__name__}: {exc}"
        print(json.dumps({
            "pid": process_id,
            "refresh_return_seconds": round(returned - started, 3),
            "refresh_return_value": repr(refresh_result),
            "loopback_hit": saw_request,
            "after_calculate": after_calculate.is_set(),
            "after_calculate_seconds": round(time.monotonic() - started, 3),
            "server_still_waiting_at_after_calculate": hit.is_set() and not release.is_set(),
            "refresh_date_before": refresh_date_before,
            "refresh_date_after": refresh_date_after,
            "refreshing_after": refreshing_after,
        }, indent=2))
        return 0 if saw_request and after_calculate.is_set() else 2
    finally:
        release.set()
        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()
        server.shutdown()
        server.server_close()


if __name__ == "__main__":
    raise SystemExit(main())
