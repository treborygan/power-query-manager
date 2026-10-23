"""Run guarded CLI package import against a disposable blank Excel workbook,
including a Unicode and punctuation-heavy query name.
"""
from __future__ import annotations

import csv
import hashlib
import json
from pathlib import Path
import subprocess
import sys
import tempfile
import time

import pythoncom
from concurrent.futures import ThreadPoolExecutor
import win32api
import win32com.client
import win32con
import win32event
import win32process

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / 'app' / 'src'))
from pq_installer.core.excel_inspector import ExcelQueryInspector
from pq_installer.core.manager import export_query_package
from pq_installer.core.models import ExistingQueryMode, LoadMode, QuerySource
from pq_installer.core.excel_queries import ExcelQueryInstaller
PACKAGED_CLI = ROOT / 'dist' / 'PowerQueryManager' / 'PowerQueryManagerCLI.exe'


def _hash(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def _excel_pids():
    result = subprocess.run(['tasklist', '/FI', 'IMAGENAME eq EXCEL.EXE', '/FO', 'CSV', '/NH'],
                           check=True, capture_output=True, text=True)
    return {int(row[1]) for row in csv.reader(result.stdout.splitlines())
            if len(row) >= 2 and row[1].isdigit()}


def _running(pid):
    try:
        handle = win32api.OpenProcess(win32con.SYNCHRONIZE, False, pid)
    except Exception:
        return False
    try:
        return win32event.WaitForSingleObject(handle, 0) == win32event.WAIT_TIMEOUT
    finally:
        win32api.CloseHandle(handle)


def main():
    if {'-h','--help'} & set(sys.argv[1:]):
        print('Usage: python tools/run_excel_cli_package_import_smoke.py\n'
              'Uses an explicit --confirm IMPORT approval with the packaged CLI.')
        return 0
    if sys.argv[1:]:
        raise SystemExit('This smoke accepts no arguments; use --help for usage.')
    if not PACKAGED_CLI.is_file():
        raise FileNotFoundError(f'Packaged CLI not found: {PACKAGED_CLI}')
    pythoncom.CoInitialize()
    excel = None
    try:
        with tempfile.TemporaryDirectory(prefix='pqm-cli-package-import-') as temp:
            root = Path(temp)
            source_target = root / 'package-source.xlsx'
            target = root / 'package-target.xlsx'
            package = root / 'queries.pqm.zip'
            formulas = {'Sales': 'let Source = 11 in Source',
                        'Inventory': 'let Source = 29 in Source',
                        'Équipe – Résultats 2026 (Q1)': 'let Source = 37 in Source'}
            excel = win32com.client.DispatchEx('Excel.Application')
            version, build = str(excel.Version), str(excel.Build)
            excel.Visible = False
            excel.DisplayAlerts = False
            excel.AutomationSecurity = 3
            seed_pid = int(win32process.GetWindowThreadProcessId(int(excel.Hwnd))[1])
            for workbook_path in (source_target,target):
                workbook = excel.Workbooks.Add()
                workbook.SaveAs(str(workbook_path), 51)
                workbook.Close(False)
                workbook = None
            excel.Quit()
            excel = None
            deadline = time.monotonic() + 10
            while _running(seed_pid) and time.monotonic() < deadline:
                time.sleep(0.1)
            if _running(seed_pid):
                raise RuntimeError('Seed Excel process remained after workbook creation.')
            baseline_pids = _excel_pids()

            sources=[]
            for order,(name,formula) in enumerate(formulas.items()):
                source=root/f'export-{order:02d}.pq'
                source.write_text(formula,encoding='utf-8')
                sources.append(QuerySource(source,name,order,True,LoadMode.DEFINITION_ONLY))
            installed=ThreadPoolExecutor(max_workers=1,thread_name_prefix='pqm-package-source').submit(
                ExcelQueryInstaller().install,source_target,sources,ExistingQueryMode.REPLACE).result()
            if any(row.status.value!='Created' for row in installed):
                raise RuntimeError(f'Live source workbook setup failed: {[row.status.value for row in installed]}')
            source_before_export=_hash(source_target)
            source_inventory=ExcelQueryInspector().inspect(source_target)
            if {row.name:row.formula for row in source_inventory} != formulas:
                raise RuntimeError('Source workbook did not read back exact query names/formulas.')
            export_query_package(source_inventory,package)
            if _hash(source_target)!=source_before_export:
                raise RuntimeError('Package export changed the source workbook.')
            deadline=time.monotonic()+15
            while (_excel_pids()-baseline_pids) and time.monotonic()<deadline:
                time.sleep(0.1)
            source_setup_processes=_excel_pids()-baseline_pids
            if source_setup_processes:
                raise RuntimeError(f'Excel process(es) remained after source workbook export: {sorted(source_setup_processes)}')
            original_hash = _hash(target)

            completed = subprocess.run([str(PACKAGED_CLI), 'import', '--apply', '--package', str(package),
                '--workbook', str(target), '--confirm', 'IMPORT'], timeout=180,
                capture_output=True,text=True,encoding='utf-8',errors='replace')
            if completed.returncode != 0:
                raise AssertionError(f'Packaged CLI package import failed ({completed.returncode}): {completed.stderr}\n{completed.stdout}')
            backups=list(root.glob('package-target_backup_*.xlsx'))
            if len(backups)!=1:
                raise AssertionError(f'Expected exactly one import backup, found {len(backups)}.')
            backup=backups[0]
            if _hash(backup) != original_hash:
                raise AssertionError('Backup hash differs from pre-import workbook.')
            inventory = ExcelQueryInspector().inspect(target)
            actual = {row.name: row.formula for row in inventory}
            if actual != formulas:
                raise AssertionError(f'Imported query names/formulas mismatch: {actual}')
            deadline = time.monotonic() + 10
            while (_excel_pids() - baseline_pids) and time.monotonic() < deadline:
                time.sleep(0.1)
            remaining = _excel_pids() - baseline_pids
            if remaining:
                raise RuntimeError(f'Excel process(es) remained after CLI import: {sorted(remaining)}')
            print(json.dumps({'status': 'PASS', 'office_version': version, 'build': build,
                'query_count': len(actual), 'names_and_formulas_verified': True,
                'live_workbook_export_verified': True,
                'unicode_query_name': 'Équipe – Résultats 2026 (Q1)',
                'source_workbook_hash_unchanged_on_export': True,
                'backup_sha256': _hash(backup), 'owned_excel_processes_exited': True,
                'external_source_evaluation': False}, indent=2))
            return 0
    finally:
        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()


if __name__ == '__main__':
    raise SystemExit(main())
