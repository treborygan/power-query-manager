"""Run the CLI refresh command on a disposable workbook containing inline static M only."""
from __future__ import annotations

import contextlib
import csv
import hashlib
import io
import json
from pathlib import Path
import subprocess
import sys
import tempfile
import time

import pythoncom
import win32api
import win32com.client
import win32con
import win32event
import win32process

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / 'app' / 'src'))
from pq_installer import cli
from pq_installer.core.excel_inspector import ExcelQueryInspector


def _hash(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def _excel_pids():
    result = subprocess.run(['tasklist', '/FI', 'IMAGENAME eq EXCEL.EXE', '/FO', 'CSV', '/NH'],
                           check=True, capture_output=True, text=True)
    pids = set()
    for row in csv.reader(result.stdout.splitlines()):
        if len(row) >= 2 and row[1].isdigit():
            pids.add(int(row[1]))
    return pids


def _running(pid):
    try:
        handle = win32api.OpenProcess(win32con.SYNCHRONIZE, False, pid)
    except Exception:
        return False
    try:
        return win32event.WaitForSingleObject(handle, 0) == win32event.WAIT_TIMEOUT
    finally:
        win32api.CloseHandle(handle)


def _run_cli_refresh(workbook_path, mode, cli_exe=None):
    command=(['refresh','--apply','--query','PQM_CLI_Refresh',
              '--workbook',str(workbook_path),'--timeout','60'] +
             (['--synchronous'] if mode=='synchronous' else []))
    if cli_exe is not None:
        completed=subprocess.run([str(cli_exe),*command,'--confirm','REFRESH'],
            capture_output=True,text=True,encoding='utf-8',errors='replace',timeout=180)
        code=completed.returncode
        rendered=completed.stdout
    else:
        output=io.StringIO()
        with contextlib.redirect_stdout(output):
            code=cli.main(command,input_func=lambda _prompt:'REFRESH')
        rendered=output.getvalue()
    if code != 0:
        detail=completed.stderr if cli_exe is not None else rendered
        raise AssertionError(f'CLI refresh {mode} failed with {code}: {detail}\n{rendered}')
    start = rendered.find('{')
    if start < 0:
        raise AssertionError(f'CLI emitted no JSON report: {rendered}')
    report = json.loads(rendered[start:])
    if len(report['results']) != 1 or report['results'][0]['status'] != 'refreshed':
        raise AssertionError(f'Unexpected CLI refresh report: {report}')
    return report


def main():
    import argparse
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--cli-exe',type=Path,help='Also run refresh through this packaged CLI executable.')
    args=parser.parse_args()
    if args.cli_exe and not args.cli_exe.is_file():
        parser.error(f'CLI executable does not exist: {args.cli_exe}')
    pythoncom.CoInitialize()
    seed_excel = None
    try:
        with tempfile.TemporaryDirectory(prefix='pqm-cli-refresh-smoke-') as temp:
            root = Path(temp)
            target = root / 'static-refresh.xlsx'
            seed_excel = win32com.client.DispatchEx('Excel.Application')
            version, build = str(seed_excel.Version), str(seed_excel.Build)
            seed_excel.Visible = False
            seed_excel.DisplayAlerts = False
            seed_excel.AutomationSecurity = 3
            workbook = seed_excel.Workbooks.Add()
            seed_pid = int(win32process.GetWindowThreadProcessId(int(seed_excel.Hwnd))[1])
            formula = 'let Source = #table({"Value"}, {{1}}) in Source'
            workbook.Queries.Add('PQM_CLI_Refresh', formula)
            connection_string = ('OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;'
                                 'Location=PQM_CLI_Refresh;Extended Properties=""')
            workbook.Connections.Add2('Query - PQM_CLI_Refresh', 'Static CLI refresh smoke',
                                      connection_string, 'SELECT * FROM [PQM_CLI_Refresh]',
                                      2, False, False)
            workbook.SaveAs(str(target), 51)
            workbook.Close(False)
            workbook = None
            seed_excel.Quit()
            seed_excel = None
            deadline = time.monotonic() + 10
            while _running(seed_pid) and time.monotonic() < deadline:
                time.sleep(0.1)
            if _running(seed_pid):
                raise RuntimeError('Seed Excel process remained after workbook creation.')
            baseline_pids = _excel_pids()

            reports = []
            backup_hashes = []
            for mode in ('background', 'synchronous'):
                before_hash = _hash(target)
                backups_before = set(root.glob('static-refresh_backup_*.xlsx'))
                report = _run_cli_refresh(target, mode,args.cli_exe)
                reports.append(report)
                new_backups = set(root.glob('static-refresh_backup_*.xlsx')) - backups_before
                if len(new_backups) != 1:
                    raise AssertionError('Expected exactly one new verified-backup candidate per refresh.')
                if _hash(next(iter(new_backups))) != before_hash:
                    raise AssertionError('A refresh backup did not match the exact pre-refresh workbook hash.')
                backup_hashes.append(before_hash)
                inventory = ExcelQueryInspector().inspect(target)
                query = next(row for row in inventory if row.name == 'PQM_CLI_Refresh')
                if query.formula != formula or query.load_state != 'connection_only':
                    raise AssertionError(f'Post-refresh read-only verification mismatch: {query}')
                if not target.is_file() or target.stat().st_size == 0:
                    raise AssertionError('Committed workbook is empty or missing.')
                deadline = time.monotonic() + 10
                while (_excel_pids() - baseline_pids) and time.monotonic() < deadline:
                    time.sleep(0.1)
                remaining = _excel_pids() - baseline_pids
                if remaining:
                    raise RuntimeError(f'Excel process(es) remained after CLI refresh: {sorted(remaining)}')

            print(json.dumps({'status': 'PASS', 'office_version': version, 'build': build,
                'modes': [row['mode'] for row in reports], 'refresh_statuses': [row['results'][0]['status'] for row in reports],
                'packaged_cli':args.cli_exe is not None,
                'verified_backup_hashes': backup_hashes, 'no_external_source': True,
                'owned_excel_processes_exited': True}, indent=2))
            return 0
    finally:
        if seed_excel is not None:
            try:
                seed_excel.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()


if __name__ == '__main__':
    raise SystemExit(main())
