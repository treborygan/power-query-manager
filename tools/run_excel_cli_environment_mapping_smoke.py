"""Exercise approved CLI M-code mapping on a disposable Excel workbook.

Uses only definition-only queries, so Excel stores the mapped formula without
evaluating File.Contents or contacting the mapped path.
"""
from __future__ import annotations

import hashlib
import csv
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
import sys
sys.path.insert(0, str(ROOT / 'app' / 'src'))
from pq_installer import cli
from pq_installer.core.excel_inspector import ExcelQueryInspector

PACKAGED_CLI = ROOT / 'dist' / 'PowerQueryManager' / 'PowerQueryManagerCLI.exe'


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
    result=subprocess.run(['tasklist','/FI','IMAGENAME eq EXCEL.EXE','/FO','CSV','/NH'],
                          check=True,capture_output=True,text=True)
    return {int(row[1]) for row in csv.reader(result.stdout.splitlines())
            if len(row)>=2 and row[1].isdigit()}


def main():
    use_packaged='--packaged' in sys.argv[1:]
    if use_packaged and not PACKAGED_CLI.is_file():
        raise FileNotFoundError(f'Packaged CLI not found: {PACKAGED_CLI}')
    pythoncom.CoInitialize()
    excel = None
    try:
        with tempfile.TemporaryDirectory(prefix='pqm-cli-env-map-') as folder:
            root = Path(folder)
            target = root / 'mapping-smoke.xlsx'
            source_dir = root / 'sources'
            mapped_dir = root / 'mapped'
            source_dir.mkdir()
            mapped_dir.mkdir()
            excel = win32com.client.DispatchEx('Excel.Application')
            version, build = str(excel.Version), str(excel.Build)
            excel.Visible = False
            excel.DisplayAlerts = False
            excel.AutomationSecurity = 3
            seed_pid = int(win32process.GetWindowThreadProcessId(int(excel.Hwnd))[1])
            book = excel.Workbooks.Add()
            book.SaveAs(str(target), 51)
            book.Close(False)
            excel.Quit()
            excel = None

            original = ('let Source = File.Contents("C:\\old\\ignored.csv"), '
                        'Value = #table({"Value"}, {{1}}) in Value')
            mapped = original.replace('C:\\old', 'D:\\new')
            untouched = 'let Value = #table({"Value"}, {{2}}) in Value'
            original_first = source_dir / 'first.pq'
            original_second = source_dir / 'second.pq'
            first = mapped_dir / 'first.pq'
            second = mapped_dir / 'second.pq'
            original_first.write_text(untouched, encoding='utf-8')
            original_second.write_text(untouched, encoding='utf-8')
            first.write_text(original, encoding='utf-8')
            second.write_text(untouched, encoding='utf-8')
            source_hashes = tuple(_sha(path) for path in (original_first,original_second,first,second))
            profile = root / 'profile.json'
            report = root / 'report.json'
            profile.write_text(json.dumps({
                'schema_version': 1,
                'profile_name': 'Disposable environment mapping smoke',
                'workbook': {'path': str(target)},
                'options': {'existing_query_policy': 'replace', 'backup': True,
                            'dry_run': False, 'refresh': 'none'},
                'path_mappings': [{'key': 'source-root', 'from': 'sources',
                                   'to': 'mapped', 'enabled': True}],
                'environment_mappings': [{'key': 'disk', 'from': 'C:\\old',
                                          'to': 'D:\\new', 'enabled': True}],
                'queries': [
                    {'id': 'first', 'source_path': 'sources/first.pq', 'query_name': 'PQM_EnvMap_One',
                     'enabled': True, 'order': 0, 'load': {'mode': 'definition_only'}},
                    {'id': 'second', 'source_path': 'sources/second.pq', 'query_name': 'PQM_EnvMap_Two',
                     'enabled': True, 'order': 1, 'load': {'mode': 'definition_only'}}],
            }), encoding='utf-8')

            owned_pids = []
            original_dispatch = win32com.client.DispatchEx
            baseline_excel_pids=_excel_pids()
            def tracked_dispatch(*args, **kwargs):
                instance = original_dispatch(*args, **kwargs)
                pid = int(win32process.GetWindowThreadProcessId(int(instance.Hwnd))[1])
                if pid != seed_pid:
                    owned_pids.append(pid)
                return instance
            win32com.client.DispatchEx = tracked_dispatch
            try:
                answers = iter(('APPLY PATH MAPPINGS', 'APPLY M-CODE MAPPINGS', 'DEPLOY'))
                arguments=['deploy', '--apply', '--profile', str(profile),
                           '--workbook', str(target), '--apply-path-mappings',
                           '--apply-environment-mappings','--output',str(report)]
                if use_packaged:
                    # The packaged process cannot inherit the in-process test callback.
                    # Supply each exact review phrase through the CLI's explicit approval
                    # channel so the unattended smoke exercises the same guarded flow.
                    completed=subprocess.run(
                        [str(PACKAGED_CLI), *arguments,
                         '--confirm', 'APPLY PATH MAPPINGS',
                         '--confirm', 'APPLY M-CODE MAPPINGS',
                         '--confirm', 'DEPLOY'],
                        check=False, timeout=240)
                    code=completed.returncode
                else:
                    code = cli.main(arguments,inspector_factory=ExcelQueryInspector,
                                    input_func=lambda _prompt: next(answers))
            finally:
                win32com.client.DispatchEx = original_dispatch
            if code != 0:
                raise RuntimeError(f'CLI deployment returned exit code {code}.')
            if tuple(_sha(path) for path in (original_first,original_second,first,second)) != source_hashes:
                raise RuntimeError('CLI mapping modified a source file.')
            report_data=json.loads(report.read_text(encoding='utf-8'))
            if report_data['path_mapping_preview'][0]['mapped_path']!=str(first.resolve()):
                raise RuntimeError('CLI path preview did not resolve the mapped source file.')
            if report_data['environment_mapping_preview'][0]['source_path']!=str(first.resolve()):
                raise RuntimeError('CLI M-code diff was not previewed against the path-mapped source.')
            inventory = {row.name: row for row in ExcelQueryInspector().inspect(target)}
            if inventory['PQM_EnvMap_One'].formula != mapped:
                raise RuntimeError('Excel did not retain the reviewed mapped M formula.')
            if inventory['PQM_EnvMap_Two'].formula != untouched:
                raise RuntimeError('Excel did not retain the untouched query formula.')
            if any(row.load_state != 'definition_only' for row in inventory.values()
                   if row.name.startswith('PQM_EnvMap_')):
                raise RuntimeError('Smoke query did not remain definition-only.')
            deadline = time.monotonic() + 15
            while any(_running(pid) for pid in owned_pids) and time.monotonic() < deadline:
                time.sleep(.1)
            lingering = [pid for pid in owned_pids if _running(pid)]
            if use_packaged:
                while (_excel_pids()-baseline_excel_pids) and time.monotonic()<deadline:
                    time.sleep(.1)
                lingering.extend(sorted(_excel_pids()-baseline_excel_pids))
            if lingering:
                raise RuntimeError(f'CLI mapping smoke left owned Excel processes running: {lingering}')
            print(json.dumps({'status': 'PASS', 'office_version': version, 'build': build,
                              'queries_verified': 2, 'path_and_m_code_mappings_verified': True,
                              'packaged_cli': use_packaged,
                              'original_sources_unchanged': True,
                              'owned_excel_processes_exited': not lingering,
                              'owned_excel_pids': owned_pids}, indent=2))
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
