"""Build and deploy a portable release bundle to disposable Excel workbooks."""
from __future__ import annotations

import hashlib
import json
from pathlib import Path
import subprocess
import sys
import tempfile
import time

import pythoncom
import win32api
import win32con
import win32event
import win32process
import win32com.client

ROOT=Path(__file__).resolve().parents[1]
sys.path.insert(0,str(ROOT/'app'/'src'))
from pq_installer.core.excel_inspector import ExcelQueryInspector
from pq_installer.core.manager import export_deployment_bundle
from pq_installer.cli import main as cli_main


def _excel_process_ids():
    result=subprocess.run(['tasklist.exe','/FI','IMAGENAME eq EXCEL.EXE','/FO','CSV','/NH'],
        capture_output=True,text=True,encoding='mbcs',errors='replace',check=True)
    import csv,io
    return {int(row[1]) for row in csv.reader(io.StringIO(result.stdout))
        if len(row)>1 and row[0].casefold()=='excel.exe' and row[1].isdigit()}


def _deploy(cli_exe,args,output_path):
    if cli_exe is None:
        return cli_main(args,input_func=lambda _prompt:'DEPLOY')
    before=_excel_process_ids()
    result=subprocess.run([str(cli_exe),*args,'--confirm','DEPLOY'],input='',capture_output=True,
        text=True,encoding='utf-8',errors='replace',timeout=180)
    if result.returncode:
        raise RuntimeError(f'Packaged CLI failed ({result.returncode}): {result.stderr}\n{result.stdout}')
    deadline=time.monotonic()+20
    while time.monotonic()<deadline:
        current=_excel_process_ids()
        if current==before:break
        if before-current:
            raise RuntimeError(f'Pre-existing Excel PIDs disappeared during the packaged CLI run: {sorted(before-current)}')
        time.sleep(.2)
    else:
        raise RuntimeError(f'Packaged CLI left Excel process IDs running: {sorted(_excel_process_ids()-before)}')
    if not Path(output_path).is_file():
        raise RuntimeError(f'Packaged CLI did not create its report: {output_path}')
    return 0


def _sha256(path):return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def _running(pid):
    try:handle=win32api.OpenProcess(win32con.SYNCHRONIZE,False,pid)
    except Exception:return False
    try:return win32event.WaitForSingleObject(handle,0)==win32event.WAIT_TIMEOUT
    finally:win32api.CloseHandle(handle)


def main():
    import argparse
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--cli-exe',type=Path,help='Also run deployment through this packaged CLI executable.')
    args=parser.parse_args()
    if args.cli_exe and not args.cli_exe.is_file():
        parser.error(f'CLI executable does not exist: {args.cli_exe}')
    pythoncom.CoInitialize()
    original_dispatch=win32com.client.DispatchEx
    owned_pids=[]
    def tracked_dispatch(*args,**kwargs):
        app=original_dispatch(*args,**kwargs)
        owned_pids.append(int(win32process.GetWindowThreadProcessId(int(app.Hwnd))[1]))
        return app
    win32com.client.DispatchEx=tracked_dispatch
    try:
        with tempfile.TemporaryDirectory(prefix='pqm-deployment-bundle-smoke-') as temporary:
            root=Path(temporary)
            targets=[root/'release-one.xlsx',root/'release-two.xlsx']
            excel=win32com.client.DispatchEx('Excel.Application')
            version,build=str(excel.Version),str(excel.Build)
            excel.Visible=False;excel.DisplayAlerts=False;excel.AutomationSecurity=3
            for target in targets:
                book=excel.Workbooks.Add();book.SaveAs(str(target),51);book.Close(False);book=None
            excel.Quit();excel=None
            for pid in owned_pids:
                deadline=time.monotonic()+15
                while _running(pid) and time.monotonic()<deadline:time.sleep(.1)
                if _running(pid):raise RuntimeError(f'Smoke-owned Excel PID {pid} did not exit after seeding.')

            source=root/'source.pq'
            formula='let Source = #table({"Value"}, {{41}, {59}}) in Source'
            source.write_text(formula,encoding='utf-8')
            profile={'schema_version':1,'profile_name':'Portable Excel deployment smoke',
              'workbook':{'path':str(targets[0])},
              'options':{'existing_query_policy':'replace','backup':True,'dry_run':False,'refresh':'none'},
              'queries':[{'id':'enabled','source_path':source.name,'query_name':'PQM_Release_Smoke',
                'enabled':True,'order':0,'load':{'mode':'definition_only'}},
                {'id':'disabled-missing','source_path':'never-created.pq','query_name':'PQM_Disabled_Missing',
                'enabled':False,'order':1,'load':{'mode':'definition_only'}}]}
            profile_path=root/'profile.json';profile_path.write_text(json.dumps(profile),encoding='utf-8')
            bundle=root/'release.pqmdeploy'
            export_deployment_bundle(profile_path,bundle)
            originals={str(path.resolve()):_sha256(path) for path in targets}

            first_report=root/'first-report.json'
            deploy_args=['deploy','--apply','--profile',str(bundle),
                '--workbook',str(targets[0]),'--workbook',str(targets[1]),
                '--output',str(first_report)]
            code=_deploy(args.cli_exe,deploy_args,first_report)
            if code!=0:raise RuntimeError(f'Packaged-profile CLI deployment returned {code}.')
            first=json.loads(first_report.read_text(encoding='utf-8'))
            reports=first['results']
            if [row['status'] for row in reports]!=['applied','applied']:
                raise RuntimeError(f'Portable bundle did not apply to both workbooks: {[row["status"] for row in reports]}')
            for report in reports:
                target=Path(report['workbook_path'])
                inventory=ExcelQueryInspector().inspect(target)
                observed=next(item for item in inventory if item.name=='PQM_Release_Smoke')
                if observed.formula!=formula or observed.load_state!='definition_only':
                    raise RuntimeError(f'Post-deployment verification failed for {target.name}.')
                if _sha256(report['backup_path'])!=originals[report['workbook_path']]:
                    raise RuntimeError(f'Pre-deployment backup hash mismatch for {target.name}.')
            for pid in set(owned_pids):
                deadline=time.monotonic()+15
                while _running(pid) and time.monotonic()<deadline:time.sleep(.1)
                if _running(pid):raise RuntimeError(f'Smoke-owned Excel PID {pid} did not exit after readback.')

            repeat_report=root/'repeat-report.json'
            deploy_args=['deploy','--apply','--profile',str(bundle),
                '--workbook',str(targets[0]),'--workbook',str(targets[1]),
                '--output',str(repeat_report)]
            code=_deploy(args.cli_exe,deploy_args,repeat_report)
            if code!=0:raise RuntimeError(f'Repeated bundle deployment returned {code}.')
            repeat=json.loads(repeat_report.read_text(encoding='utf-8'))
            if [row['status'] for row in repeat['results']]!=['no_changes','no_changes']:
                raise RuntimeError('Repeated bundle deployment was not a no-op for both targets.')
            owned=set(owned_pids)
            deadline=time.monotonic()+20
            while any(_running(pid) for pid in owned) and time.monotonic()<deadline:time.sleep(.1)
            lingering=[pid for pid in owned if _running(pid)]
            if lingering:raise RuntimeError(f'Bundle smoke left owned Excel processes running: {lingering}')
            print(json.dumps({'status':'PASS','office_version':version,'build':build,
                'bundle':'release.pqmdeploy','workbooks':len(targets),
                'first_batch':[row['status'] for row in reports],
                'repeat_batch':[row['status'] for row in repeat['results']],
                'exact_formula_readback':True,'verified_original_hash_backups':len(reports),
                'disabled_missing_source_ignored':True,'owned_excel_pids':sorted(owned)},indent=2))
            return 0
    finally:
        win32com.client.DispatchEx=original_dispatch
        pythoncom.CoUninitialize()


if __name__=='__main__':raise SystemExit(main())
