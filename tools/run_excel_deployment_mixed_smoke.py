"""Verify isolated batch deployment when one disposable target is read-only."""
from __future__ import annotations

from concurrent.futures import ThreadPoolExecutor
import argparse
import hashlib
import json
import os
from pathlib import Path
import stat
import subprocess
import tempfile
import time

import pythoncom
import win32api
import win32com.client
import win32con
import win32event
import win32process

ROOT = Path(__file__).resolve().parents[1]
import sys
sys.path.insert(0, str(ROOT / 'app' / 'src'))
from pq_installer.core.deployment import deploy_multi_workbook
from pq_installer.core.excel_inspector import ExcelQueryInspector
from pq_installer.core.excel_queries import ExcelQueryInstaller
from pq_installer.core.history import HistoryStore


def _sha(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def _running(pid):
    try:
        handle = win32api.OpenProcess(win32con.SYNCHRONIZE, False, pid)
    except Exception:
        return False
    try:
        return win32event.WaitForSingleObject(handle, 0) == win32event.WAIT_TIMEOUT
    finally:
        win32api.CloseHandle(handle)


def _excel_pids():
    result = subprocess.run(['tasklist', '/FI', 'IMAGENAME eq EXCEL.EXE', '/FO', 'CSV', '/NH'],
                            check=True, capture_output=True, text=True)
    import csv
    return {int(row[1]) for row in csv.reader(result.stdout.splitlines())
            if len(row) >= 2 and row[1].isdigit()}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--cli-exe', type=Path,
                        help='exercise the packaged CLI instead of the in-process deployment service')
    args = parser.parse_args()
    if args.cli_exe is not None and not args.cli_exe.is_file():
        parser.error(f'Packaged CLI does not exist: {args.cli_exe}')
    pythoncom.CoInitialize()
    original_dispatch = win32com.client.DispatchEx
    owned_pids = []
    readonly_target = None
    try:
        with tempfile.TemporaryDirectory(prefix='pqm-batch-mixed-') as folder:
            root = Path(folder)
            targets = [root / 'good.xlsx', root / 'readonly.xlsx']
            baseline_pids = _excel_pids()
            excel = win32com.client.DispatchEx('Excel.Application')
            version, build = str(excel.Version), str(excel.Build)
            excel.Visible = False
            excel.DisplayAlerts = False
            excel.AutomationSecurity = 3
            seed_pid = int(win32process.GetWindowThreadProcessId(int(excel.Hwnd))[1])
            for target in targets:
                book = excel.Workbooks.Add()
                book.SaveAs(str(target), 51)
                book.Close(False)
            excel.Quit()
            excel = None
            deadline = time.monotonic() + 15
            while _running(seed_pid) and time.monotonic() < deadline:
                time.sleep(.1)
            if _running(seed_pid):
                raise RuntimeError(f'Seed Excel process {seed_pid} did not exit after Quit.')
            if _excel_pids() != baseline_pids:
                raise RuntimeError(f'Seed workbook setup did not return to the initial Excel process set: '
                                   f'{sorted(_excel_pids())} != {sorted(baseline_pids)}')
            original_hashes = [_sha(path) for path in targets]
            readonly_target = targets[1]
            readonly_target.chmod(stat.S_IREAD)

            source = root / 'static.pq'
            formula = 'let Source = #table({"Value"}, {{19}}) in Source'
            source.write_text(formula, encoding='utf-8')
            profile = {
                'schema_version': 1,
                'profile_name': 'Disposable mixed deployment smoke',
                'workbook': {'path': str(targets[0])},
                'options': {'existing_query_policy': 'replace', 'backup': True,
                            'dry_run': False, 'refresh': 'none'},
                'queries': [{'id': 'mixed-query', 'source_path': str(source),
                             'query_name': 'PQM_Mixed_Smoke', 'enabled': True,
                             'order': 0, 'load': {'mode': 'definition_only'}}],
            }
            profile_path = root / 'profile.json'

            if args.cli_exe is not None:
                profile_path.write_text(json.dumps(profile, ensure_ascii=False, indent=2), encoding='utf-8')
                cli_env = os.environ.copy()
                cli_env['LOCALAPPDATA'] = str(root / 'cli-localappdata')
                completed = subprocess.run(
                    [str(args.cli_exe.resolve()), 'deploy', '--profile', str(profile_path),
                     '--workbook', str(targets[0]), '--workbook', str(targets[1]),
                     '--apply', '--confirm', 'DEPLOY'],
                    capture_output=True, text=True, encoding='utf-8', errors='replace', timeout=180,
                    env=cli_env)
                if completed.returncode != 5:
                    raise RuntimeError(f'Expected partial-deployment exit 5; got {completed.returncode}: '
                                       f'{completed.stderr}\n{completed.stdout}')
                try:
                    report_start = completed.stdout.index('{')
                    bundle, _ = json.JSONDecoder().raw_decode(completed.stdout, report_start)
                except json.JSONDecodeError as exc:
                    raise RuntimeError(f'Packaged CLI did not return a JSON deployment report: '
                                       f'{completed.stdout}\n{completed.stderr}') from exc
                except ValueError as exc:
                    raise RuntimeError(f'Packaged CLI did not include a JSON deployment report: '
                                       f'{completed.stdout}\n{completed.stderr}') from exc
                reports = bundle['results']
                states = [row['status'] for row in reports]
                inspection = subprocess.run(
                    [str(args.cli_exe.resolve()), 'inspect', '--workbook', str(targets[0])],
                    capture_output=True, text=True, encoding='utf-8', errors='replace', timeout=90,
                    env=cli_env)
                if inspection.returncode != 0:
                    raise RuntimeError(f'Packaged CLI could not inspect its successful target: '
                                       f'{inspection.stderr}\n{inspection.stdout}')
                observed = next(row for row in json.loads(inspection.stdout)['queries']
                                if row['name'] == 'PQM_Mixed_Smoke')
                if observed['formula'] != formula:
                    raise RuntimeError('Packaged CLI deployment formula did not verify after reopen.')
                events = HistoryStore(cli_env['LOCALAPPDATA'] + '\\PowerQueryManager\\history.sqlite3').list()
                if len(events) != 1 or events[0]['event'] != 'cli_deployment_completed':
                    raise AssertionError(f'Packaged CLI did not record exactly one deployment history event: {events!r}')
                expected_history = {'action': 'apply', 'created_count': 1, 'query_count': 1,
                    'replaced_count': 0, 'skipped_count': 0, 'status': 'partial', 'workbook_count': 2}
                if events[0]['details'] != expected_history:
                    raise AssertionError(f'Packaged CLI deployment history mismatch: {events[0]!r}')
                audit_text = json.dumps(events, ensure_ascii=False)
                if (any(str(target) in audit_text for target in targets)
                        or str(source) in audit_text or 'PQM_Mixed_Smoke' in audit_text or formula in audit_text):
                    raise AssertionError('Packaged CLI history stored a source/workbook path, query name, or formula.')
                if _excel_pids() != baseline_pids:
                    raise RuntimeError(f'Packaged CLI changed the Excel process set: '
                                       f'{sorted(_excel_pids())} != {sorted(baseline_pids)}')
            else:
                def tracked_dispatch(*dispatch_args, **kwargs):
                    instance = original_dispatch(*dispatch_args, **kwargs)
                    owned_pids.append(int(win32process.GetWindowThreadProcessId(int(instance.Hwnd))[1]))
                    return instance

                win32com.client.DispatchEx = tracked_dispatch
                try:
                    reports = ThreadPoolExecutor(max_workers=1, thread_name_prefix='pqm-mixed-deploy').submit(
                        deploy_multi_workbook, profile, profile_path, targets,
                        inventory_provider=lambda path: ExcelQueryInspector().inspect(path),
                        installer_factory=lambda _path: ExcelQueryInstaller(),
                        confirm_apply=lambda plans: True,
                    ).result()
                finally:
                    win32com.client.DispatchEx = original_dispatch
                states = [row.status for row in reports]

            if states != ['applied', 'blocked_preflight']:
                raise RuntimeError(f'Expected isolated applied/blocked results; got {states}.')
            backup_path = reports[1].get('backup_path') if args.cli_exe is not None else reports[1].backup_path
            if backup_path is not None:
                raise RuntimeError('Read-only target received a backup before write preflight.')
            if _sha(targets[1]) != original_hashes[1]:
                raise RuntimeError('Read-only target changed despite blocked preflight.')
            successful_backup = reports[0].get('backup_path') if args.cli_exe is not None else reports[0].backup_path
            if not successful_backup or _sha(successful_backup) != original_hashes[0]:
                raise RuntimeError('Successful target backup does not match its original hash.')
            if args.cli_exe is None:
                observed = next(row for row in ExcelQueryInspector().inspect(targets[0])
                                if row.name == 'PQM_Mixed_Smoke')
                if observed.formula != formula:
                    raise RuntimeError('Successful target formula did not verify after reopen.')

            readonly_target.chmod(stat.S_IREAD | stat.S_IWRITE)
            deadline = time.monotonic() + 15
            while any(_running(pid) for pid in owned_pids) and time.monotonic() < deadline:
                time.sleep(.1)
            lingering = [pid for pid in owned_pids if _running(pid)]
            if lingering:
                raise RuntimeError(f'Mixed deployment left owned Excel processes running: {lingering}')
            print(json.dumps({'status': 'PASS', 'office_version': version, 'build': build,
                              'workbook_statuses': states, 'successful_backup_verified': True,
                              'blocked_target_hash_unchanged': True,
                              'blocked_target_backup': backup_path,
                              'packaged_cli': str(args.cli_exe.resolve()) if args.cli_exe else None,
                              'owned_excel_pids': owned_pids}, indent=2))
            return 0
    finally:
        win32com.client.DispatchEx = original_dispatch
        if readonly_target is not None and readonly_target.exists():
            readonly_target.chmod(stat.S_IREAD | stat.S_IWRITE)
        pythoncom.CoUninitialize()


if __name__ == '__main__':
    raise SystemExit(main())
