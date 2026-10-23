"""Probe real Excel Data Model refresh semantics using a disposable static-M workbook."""
from __future__ import annotations

import gc
import json
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
import tempfile
import threading
import time

import pythoncom
import win32process
import win32com.client

ROOT = Path(__file__).resolve().parents[1]
import sys
sys.path.insert(0, str(ROOT / 'app' / 'src'))

from pq_installer.core.excel_capabilities import probe_excel_capabilities
from pq_installer.core.excel_lifetime import (
    call_with_process_timeout,
    open_owned_excel_process,
    owned_excel_pid,
    quit_owned_excel,
)
from pq_installer.core.excel_queries import ExcelQueryInstaller
from pq_installer.core.models import ExistingQueryMode, LoadMode, QuerySource


def _read_record_count(model_table, pid, process):
    return call_with_process_timeout(
        pid, lambda: int(model_table.RecordCount), 10,
        operation='read disposable Data Model row count', owned_process=process)


def _seed_empty_workbook(path):
    pythoncom.CoInitialize()
    excel = workbook = process = None
    try:
        excel = win32com.client.DispatchEx('Excel.Application')
        excel.Visible = False
        excel.DisplayAlerts = False
        pid = owned_excel_pid(excel)
        process = open_owned_excel_process(pid, hwnd=int(excel.Hwnd))
        workbook = excel.Workbooks.Add()
        workbook.SaveAs(str(path), 51)
        workbook.Close(False)
        workbook = None
        return pid
    finally:
        if workbook is not None:
            try:
                workbook.Close(False)
            except Exception:
                pass
        workbook = None
        gc.collect()
        if excel is not None:
            try:
                quit_owned_excel(excel, after_quit=pythoncom.CoUninitialize)
                excel = None
                pythoncom.CoInitialize()
            except Exception:
                pass
        excel = None
        gc.collect()
        if process is not None:
            process.close()
        pythoncom.CoUninitialize()


def main() -> int:
    baseline = probe_excel_capabilities()
    if not baseline.supports('data_model'):
        raise RuntimeError(f'Data Model load capability probe failed: {baseline.reasons}')
    query_name = 'PQM_Model_Refresh_Spike'
    one_row = 'let Source = #table({"Value"}, {{1}}) in Source'
    hits: list[str] = []
    class CsvHandler(BaseHTTPRequestHandler):
        def do_GET(self):
            hits.append(self.path)
            time.sleep(0.5)
            payload = b'Value\r\n10\r\n20\r\n'
            self.send_response(200)
            self.send_header('Content-Type', 'text/csv; charset=utf-8')
            self.send_header('Content-Length', str(len(payload)))
            self.end_headers()
            self.wfile.write(payload)

        def log_message(self, *_args):
            pass

    server = ThreadingHTTPServer(('127.0.0.1', 0), CsvHandler)
    server.daemon_threads = True
    threading.Thread(target=server.serve_forever, daemon=True).start()
    url = f'http://127.0.0.1:{server.server_port}/refresh.csv'
    two_rows = (
        'let Source = Csv.Document(Web.Contents("' + url +
        '"), [Delimiter=",", Encoding=65001, QuoteStyle=QuoteStyle.None]), '
        'Headers = Table.PromoteHeaders(Source) in Headers')
    result: dict[str, object] = {
        'office_version': baseline.office_version,
        'office_build': baseline.build,
        'refresh_method': None,
        'initial_record_count': None,
        'final_record_count': None,
        'refresh_return': None,
        'refresh_error': None,
        'requests': 0,
    }
    with tempfile.TemporaryDirectory(prefix='pqm-model-refresh-spike-') as folder:
        root = Path(folder)
        workbook_path = root / 'model-refresh.xlsx'
        source_path = root / 'model-refresh.m'
        _seed_empty_workbook(workbook_path)
        source_path.write_text(one_row, encoding='utf-8')
        source = QuerySource(source_path, query_name, load_mode=LoadMode.DATA_MODEL)
        installed = ExcelQueryInstaller().install(
            workbook_path, [source], ExistingQueryMode.REPLACE, capabilities=baseline)
        if len(installed) != 1 or installed[0].status.value != 'Created':
            raise RuntimeError(f'Could not create disposable Data Model fixture: {installed!r}')

        pythoncom.CoInitialize()
        excel = workbook = process = connection = model_table = None
        try:
            excel = win32com.client.DispatchEx('Excel.Application')
            excel.Visible = False
            excel.DisplayAlerts = False
            excel.AutomationSecurity = 3
            hwnd = int(excel.Hwnd)
            pid = owned_excel_pid(excel)
            process = open_owned_excel_process(pid, hwnd=hwnd)
            workbook = excel.Workbooks.Open(
                str(workbook_path), UpdateLinks=0, ReadOnly=False,
                IgnoreReadOnlyRecommended=True, AddToMru=False)
            connection = workbook.Connections.Item(f'Query - {query_name}')
            model_table = next(
                workbook.Model.ModelTables.Item(index)
                for index in range(1, int(workbook.Model.ModelTables.Count) + 1)
                if str(workbook.Model.ModelTables.Item(index).SourceWorkbookConnection.Name).casefold()
                == f'query - {query_name}'.casefold()
            )
            initial = _read_record_count(model_table, pid, process)
            result['initial_record_count'] = initial
            if initial != 1:
                raise RuntimeError(f'Data Model fixture has {initial} rows, expected one.')

            workbook.Queries.Item(query_name).Formula = two_rows
            workbook.Save()
            result['formula_updated_without_refresh'] = str(
                workbook.Queries.Item(query_name).Formula) == two_rows

            methods = (
                ('ModelTable.Refresh', lambda: model_table.Refresh()),
            )
            for method_name, method in methods:
                try:
                    returned = call_with_process_timeout(
                        pid, method, 30, operation=f'probe {method_name}',
                        owned_process=process)
                    result['refresh_method'] = method_name
                    result['refresh_return'] = repr(returned)
                    deadline = time.monotonic() + 10
                    observed = initial
                    while time.monotonic() < deadline:
                        observed = _read_record_count(model_table, pid, process)
                        if observed == 2:
                            break
                        time.sleep(0.25)
                    result['final_record_count'] = observed
                    if observed == 2:
                        result['status'] = 'PASS'
                        break
                    result['refresh_error'] = 'refresh returned without updating the Data Model row count'
                    result['status'] = 'NO_VERIFIED_COMPLETION'
                except Exception as exc:
                    result['refresh_method'] = method_name
                    result['refresh_error'] = f'{type(exc).__name__}: {str(exc)[:400]}'
                    result['status'] = 'METHOD_FAILED'
                    if 'timed out' in str(exc).casefold():
                        # The exact DispatchEx process was terminated by the bounded-call helper.
                        raise
                if result.get('status') == 'PASS':
                    break
            if result.get('status') != 'PASS':
                raise RuntimeError(json.dumps(result, indent=2))
            result['requests'] = len(hits)
            result['request_paths'] = list(hits)
            if not hits or any(not path.startswith('/refresh.csv') for path in hits):
                raise RuntimeError(f'Data Model refresh did not use only the local fixture: {hits!r}')
            print(json.dumps(result, indent=2))
            return 0
        finally:
            # All operations use this script's DispatchEx instance; no other Excel PID is targeted.
            model_table = connection = None
            if workbook is not None:
                try:
                    workbook.Close(False)
                except Exception:
                    pass
            workbook = None
            gc.collect()
            if excel is not None and process is not None:
                try:
                    quit_owned_excel(excel, after_quit=pythoncom.CoUninitialize)
                    pythoncom.CoInitialize()
                except Exception:
                    pass
            elif excel is not None:
                try:
                    excel.Quit()
                except Exception:
                    pass
            excel = None
            gc.collect()
            if process is not None:
                process.close()
            pythoncom.CoUninitialize()
            server.shutdown()
            server.server_close()


if __name__ == '__main__':
    raise SystemExit(main())
