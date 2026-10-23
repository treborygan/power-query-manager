"""Verify staged definition-only import preserves a local PivotTable slicer."""
from __future__ import annotations

import csv
import gc
import hashlib
import json
from pathlib import Path
import subprocess
import sys
import tempfile
import time
from concurrent.futures import ThreadPoolExecutor

import pythoncom
import win32api
import win32com.client
import win32con
import win32event
import win32process

ROOT=Path(__file__).resolve().parents[1]
sys.path.insert(0,str(ROOT/'app'/'src'))
from pq_installer.core.excel_inspector import ExcelQueryInspector
from pq_installer.core.excel_queries import ExcelQueryInstaller
from pq_installer.core.models import ExistingQueryMode,LoadMode,QuerySource


def _excel_pids():
    result=subprocess.run(['tasklist.exe','/FI','IMAGENAME eq EXCEL.EXE','/FO','CSV','/NH'],
        capture_output=True,text=True,encoding='mbcs',errors='replace',check=True)
    return {int(row[1]) for row in csv.reader(result.stdout.splitlines())
        if len(row)>1 and row[0].casefold()=='excel.exe' and row[1].isdigit()}


def _running(pid):
    try:handle=win32api.OpenProcess(win32con.SYNCHRONIZE,False,pid)
    except Exception:return False
    try:return win32event.WaitForSingleObject(handle,0)==win32event.WAIT_TIMEOUT
    finally:win32api.CloseHandle(handle)


def _sha256(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def _wait_owned(pids,timeout=20):
    for pid in pids:
        deadline=time.monotonic()+timeout
        while _running(pid) and time.monotonic()<deadline:time.sleep(.1)
        if _running(pid):
            result=subprocess.run(['powershell.exe','-NoProfile','-Command',
                f"$p=Get-CimInstance Win32_Process -Filter 'ProcessId = {int(pid)}'; if ($p) {{ \"$($p.Name)|$($p.ExecutablePath)\" }}"],
                capture_output=True,text=True,timeout=10)
            if result.stdout.strip().casefold()!='excel.exe|c:\\program files\\microsoft office\\root\\office16\\excel.exe':
                raise RuntimeError(f'Owned PID {pid} identity changed; refusing termination.')
            handle=win32api.OpenProcess(win32con.PROCESS_TERMINATE|win32con.SYNCHRONIZE,False,pid)
            try:
                win32api.TerminateProcess(handle,1)
                if win32event.WaitForSingleObject(handle,5000)!=win32event.WAIT_OBJECT_0:
                    raise RuntimeError(f'Smoke-owned Excel PID {pid} did not exit.')
            finally:win32api.CloseHandle(handle)


def main():
    pythoncom.CoInitialize()
    original_dispatch=win32com.client.DispatchEx
    owned=[]
    baseline=_excel_pids()
    def dispatch(*args,**kwargs):
        app=original_dispatch(*args,**kwargs)
        owned.append(int(win32process.GetWindowThreadProcessId(int(app.Hwnd))[1]))
        return app
    win32com.client.DispatchEx=dispatch
    try:
        with tempfile.TemporaryDirectory(prefix='pqm-slicer-preserve-') as temporary:
            root=Path(temporary)
            workbook_path=root/'slicer-preserve.xlsx'
            source_path=root/'new-query.pq'
            formula='let Source = #table({"Value"}, {{17}, {29}}) in Source'
            source_path.write_text(formula,encoding='utf-8')

            excel=win32com.client.DispatchEx('Excel.Application')
            version,build=str(excel.Version),str(excel.Build)
            excel.Visible=False;excel.DisplayAlerts=False;excel.AutomationSecurity=3
            book=excel.Workbooks.Add()
            source_sheet=book.Worksheets.Item(1)
            source_sheet.Range('A1:B4').Value=(('Category','Value'),('Keep',42),('Drop',7),('Keep',12))
            table=source_sheet.ListObjects.Add(1,source_sheet.Range('A1:B4'),None,1)
            table.Name='PQM_Slicer_Source'
            pivot_sheet=book.Worksheets.Add();pivot_sheet.Name='PQM_Slicer_Pivot'
            cache=book.PivotCaches().Create(1,f"'{source_sheet.Name}'!{table.Range.Address}")
            pivot=cache.CreatePivotTable(pivot_sheet.Range('A1'),'PQM_Slicer_PivotTable')
            pivot.PivotFields('Category').Orientation=1
            pivot.AddDataField(pivot.PivotFields('Value'),'PQM Slicer Sum',-4157)
            slicer_cache=book.SlicerCaches.Add(pivot,'Category','PQM_Slicer_Cache')
            # Early-binding this already-owned collection makes pywin32 pass
            # all optional middle arguments to Excel's Slicers.Add correctly.
            slicers=win32com.client.gencache.EnsureDispatch(slicer_cache.Slicers)
            slicer=slicers.Add(SlicerDestination=pivot_sheet,Name='PQM_Slicer',
                Caption='Keep category filter',Top=10,Left=250,Width=144,Height=216)
            slicer_cache.SlicerItems.Item('Drop').Selected=False
            book.SaveAs(str(workbook_path),51)
            source_sheet=table=pivot_sheet=cache=pivot=slicer_cache=slicers=slicer=None
            gc.collect();book.Close(False);book=None;excel.Quit();excel=None;gc.collect()
            _wait_owned(owned)

            source=QuerySource(source_path,'PQM_Slicer_Import',load_mode=LoadMode.DEFINITION_ONLY)
            with ThreadPoolExecutor(max_workers=1,thread_name_prefix='pqm-slicer-install') as executor:
                results=executor.submit(ExcelQueryInstaller().install,workbook_path,[source],
                    ExistingQueryMode.REPLACE).result()
            if len(results)!=1 or results[0].status.value!='Created':
                raise AssertionError(f'Definition-only import failed: {results!r}')
            with ThreadPoolExecutor(max_workers=1,thread_name_prefix='pqm-slicer-inspect') as executor:
                inventory=executor.submit(ExcelQueryInspector().inspect,workbook_path).result()
            imported=next(row for row in inventory if row.name=='PQM_Slicer_Import')
            if imported.formula!=formula or imported.load_state!='definition_only':
                raise AssertionError(f'Imported query verification mismatch: {imported!r}')

            verify=win32com.client.DispatchEx('Excel.Application')
            verify.Visible=False;verify.DisplayAlerts=False;verify.AutomationSecurity=3
            check=verify.Workbooks.Open(str(workbook_path),UpdateLinks=0,ReadOnly=True,
                IgnoreReadOnlyRecommended=True,AddToMru=False)
            try:
                saved_cache=check.SlicerCaches.Item('PQM_Slicer_Cache')
                if int(saved_cache.Slicers.Count)!=1:
                    raise AssertionError(f'Expected one saved slicer, observed {saved_cache.Slicers.Count}.')
                saved_slicer=saved_cache.Slicers.Item(1)
                if (str(saved_slicer.Name)!='PQM_Slicer'
                    or str(saved_slicer.Caption)!='Keep category filter'):
                    raise AssertionError(f'Saved slicer identity changed: {saved_slicer.Name!r}, {saved_slicer.Caption!r}.')
                if not bool(saved_cache.SlicerItems.Item('Keep').Selected):
                    raise AssertionError('The selected Keep slicer item was not preserved.')
                if bool(saved_cache.SlicerItems.Item('Drop').Selected):
                    raise AssertionError('The excluded Drop slicer item became selected.')
                saved_pivot=check.Worksheets.Item('PQM_Slicer_Pivot').PivotTables().Item('PQM_Slicer_PivotTable')
                if str(saved_pivot.DataFields.Item(1).Name)!='PQM Slicer Sum':
                    raise AssertionError('PivotTable data field changed during staged import.')
                values=saved_pivot.DataBodyRange.Value
                if not isinstance(values,tuple):values=((values,),)
                elif values and not isinstance(values[0],tuple):values=tuple((value,) for value in values)
                numbers=[float(value) for row in values for value in row if value is not None]
                if len(numbers)!=2 or numbers[0]!=54.0 or numbers[-1]!=54.0:
                    raise AssertionError(f'Slicer-filtered PivotTable values changed: {values!r}')
                check.Close(False);check=None
            finally:
                saved_pivot=saved_slicer=saved_cache=None
                if check is not None:check.Close(False)
                verify.Quit();verify=None;gc.collect()
            _wait_owned(owned)
            current=_excel_pids()
            if current!=baseline:
                raise RuntimeError(f'Excel process inventory changed: before={sorted(baseline)}, after={sorted(current)}')
            print(json.dumps({'status':'PASS','office_version':version,'build':build,
                'pivot_table':'PQM_Slicer_PivotTable','slicer':'PQM_Slicer',
                'selected_item':'Keep','excluded_item':'Drop','filtered_value':54,
                'imported_formula_verified':True,
                'owned_excel_processes_exited':True,'no_external_source':True},indent=2))
            return 0
    finally:
        win32com.client.DispatchEx=original_dispatch
        try:_wait_owned(owned)
        finally:pythoncom.CoUninitialize()


if __name__=='__main__':raise SystemExit(main())
