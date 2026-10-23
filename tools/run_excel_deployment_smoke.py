"""Real Excel batch-deployment smoke using only disposable static-M workbooks."""
from __future__ import annotations

from concurrent.futures import ThreadPoolExecutor
import hashlib
import json
from pathlib import Path
import tempfile
import time

import pythoncom
import win32api
import win32com.client
import win32con
import win32event
import win32process

from pq_installer.core.deployment import deploy_multi_workbook
from pq_installer.core.excel_inspector import ExcelQueryInspector
from pq_installer.core.excel_queries import ExcelQueryInstaller


def _hash(path):return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def _running(pid):
    try:handle=win32api.OpenProcess(win32con.SYNCHRONIZE,False,pid)
    except Exception:return False
    try:return win32event.WaitForSingleObject(handle,0)==win32event.WAIT_TIMEOUT
    finally:win32api.CloseHandle(handle)


def main():
    pythoncom.CoInitialize();owned_pids=[];original_dispatch=win32com.client.DispatchEx
    def tracked_dispatch(*args,**kwargs):
        instance=original_dispatch(*args,**kwargs)
        owned_pids.append(int(win32process.GetWindowThreadProcessId(int(instance.Hwnd))[1]))
        return instance
    win32com.client.DispatchEx=tracked_dispatch
    try:
        with tempfile.TemporaryDirectory(prefix='pqm-deployment-smoke-') as temporary:
            root=Path(temporary);targets=[root/'batch-one.xlsx',root/'batch-two.xlsx']
            excel=win32com.client.DispatchEx('Excel.Application')
            version,build=str(excel.Version),str(excel.Build)
            excel.Visible=False;excel.DisplayAlerts=False;excel.AutomationSecurity=3
            for target in targets:
                workbook=excel.Workbooks.Add();workbook.SaveAs(str(target),51);workbook.Close(False)
            excel.Quit();excel=None;workbook=None
            original_hashes={str(path.resolve()):_hash(path) for path in targets}
            source=root/'static.pq'
            formula='let Source = #table({"Value"}, {{1}, {2}}) in Source'
            source.write_text(formula,encoding='utf-8')
            profile={'schema_version':1,'profile_name':'Disposable deployment smoke',
              'workbook':{'path':str(targets[0])},
              'options':{'existing_query_policy':'replace','backup':True,'dry_run':False,'refresh':'none'},
              'queries':[{'id':'smoke-query','source_path':str(source),'query_name':'PQM_Batch_Smoke',
                'enabled':True,'order':0,'load':{'mode':'definition_only'}},
                {'id':'disabled-missing','source_path':str(root/'disabled-source.pq'),
                'query_name':'PQM_Disabled_Missing','enabled':False,'order':1,
                'load':{'mode':'definition_only'}}]}
            profile_path=root/'profile.json'
            reports=ThreadPoolExecutor(max_workers=1,thread_name_prefix='pqm-deploy-smoke').submit(
                deploy_multi_workbook,profile,profile_path,targets,
                inventory_provider=lambda path:ExcelQueryInspector().inspect(path),
                installer_factory=lambda path:ExcelQueryInstaller(),
                confirm_apply=lambda plans:True).result()
            if [row.status for row in reports]!=['applied','applied']:
                raise RuntimeError(f'Batch deployment did not apply to both workbooks: {[row.status for row in reports]}')
            for report in reports:
                target=Path(report.workbook_path)
                observed=next(row for row in ExcelQueryInspector().inspect(target) if row.name=='PQM_Batch_Smoke')
                if observed.formula!=formula or observed.load_state!='definition_only':
                    raise RuntimeError(f'Post-deployment read-only verification failed for {target.name}.')
                if _hash(report.backup_path)!=original_hashes[report.workbook_path]:
                    raise RuntimeError(f'Backup does not match the pre-deployment workbook: {target.name}.')
            backup_paths=[row.backup_path for row in reports]
            repeat=ThreadPoolExecutor(max_workers=1,thread_name_prefix='pqm-deploy-repeat').submit(
                deploy_multi_workbook,profile,profile_path,targets,
                inventory_provider=lambda path:ExcelQueryInspector().inspect(path),
                installer_factory=lambda path:ExcelQueryInstaller()).result()
            if [row.status for row in repeat]!=['no_changes','no_changes']:
                raise RuntimeError(f'Repeated batch deployment was not a no-op: {[row.status for row in repeat]}')
            deadline=time.monotonic()+15
            while any(_running(pid) for pid in owned_pids) and time.monotonic()<deadline:time.sleep(.1)
            lingering=[pid for pid in owned_pids if _running(pid)]
            if lingering:raise RuntimeError(f'Deployment smoke left owned Excel processes running: {lingering}')
            print(json.dumps({'status':'PASS','office_version':version,'build':build,
              'workbooks':len(targets),'first_batch':[row.status for row in reports],
              'repeat_batch':[row.status for row in repeat],'verified_backups':len(backup_paths),
              'ignored_missing_disabled_source':True,
              'owned_excel_pids':owned_pids},indent=2))
            return 0
    finally:
        win32com.client.DispatchEx=original_dispatch
        pythoncom.CoUninitialize()


if __name__=='__main__':raise SystemExit(main())
