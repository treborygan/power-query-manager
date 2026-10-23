"""Verify abandoned-stage recovery using a disposable real Excel workbook."""
from pathlib import Path
from tempfile import TemporaryDirectory
import argparse
import hashlib
import shutil
import sys
import uuid
ROOT=Path(__file__).resolve().parents[1]
sys.path.insert(0,str(ROOT/'app'/'src'))

from PySide6.QtWidgets import QApplication
import pythoncom
import win32api
import win32con
import win32event
import win32process
import win32com.client

from pq_installer.core.excel_inspector import ExcelQueryInspector
from pq_installer.ui.recovery_worker import StagedRecoveryWorker

QUERY_NAME='RecoveryQuery'
FORMULA='let Source = #table({"Value"}, {{7}, {8}}) in Source'


def sha256(path):
    digest=hashlib.sha256()
    with Path(path).open('rb') as stream:
        for chunk in iter(lambda:stream.read(1024*1024),b''):
            digest.update(chunk)
    return digest.hexdigest()


def main():
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--xlsm',action='store_true',help='use a disposable macro-enabled workbook container')
    args=parser.parse_args()
    QApplication.instance() or QApplication([])
    pythoncom.CoInitialize()
    excel=None;book=None;handle=None
    try:
        excel=win32com.client.DispatchEx('Excel.Application')
        excel.Visible=False;excel.DisplayAlerts=False
        version=str(excel.Version);build=str(excel.Build)
        _thread_id,excel_pid=win32process.GetWindowThreadProcessId(int(excel.Hwnd))
        handle=win32api.OpenProcess(win32con.SYNCHRONIZE,False,excel_pid)
        with TemporaryDirectory(prefix='pqm-excel-recovery-smoke-') as directory:
            root=Path(directory)
            workbook_path=root/('recovery.xlsm' if args.xlsm else 'recovery.xlsx')
            stage=root/f'.recovery.pqstage-{uuid.uuid4().hex}{workbook_path.suffix}'
            output=root/('recovery_recovered.xlsm' if args.xlsm else 'recovery_recovered.xlsx')
            book=excel.Workbooks.Add()
            book.Worksheets.Item(1).Cells(1,1).Value='disposable staged recovery'
            book.Queries.Add(QUERY_NAME,FORMULA,'')
            book.SaveAs(str(workbook_path),FileFormat=52 if args.xlsm else 51)
            book.Close(False);book=None
            excel.Quit();excel=None
            if win32event.WaitForSingleObject(handle,5000)!=win32event.WAIT_OBJECT_0:
                raise RuntimeError('Smoke-owned seed Excel process did not exit.')
            original_hash=sha256(workbook_path)
            shutil.copy2(workbook_path,stage)
            stage_hash=sha256(stage)
            outcomes=[];failures=[];owned_processes=[]
            original_dispatch=win32com.client.DispatchEx
            def tracked_dispatch(*args,**kwargs):
                instance=original_dispatch(*args,**kwargs)
                pid=int(win32process.GetWindowThreadProcessId(int(instance.Hwnd))[1])
                handle=win32api.OpenProcess(win32con.SYNCHRONIZE,False,pid)
                owned_processes.append((pid,handle))
                return instance
            win32com.client.DispatchEx=tracked_dispatch
            worker=StagedRecoveryWorker(workbook_path,stage,output)
            try:
                worker.completed.connect(outcomes.append);worker.failed.connect(failures.append)
                worker.run()
                recovered_inventory=ExcelQueryInspector().inspect(output)
            finally:
                win32com.client.DispatchEx=original_dispatch
            if failures or len(outcomes)!=1:
                raise RuntimeError(f'Recovery worker failed: {failures!r}')
            if sha256(workbook_path)!=original_hash:
                raise RuntimeError('Selected workbook changed during recovery.')
            if (not stage.is_file() or sha256(stage)!=stage_hash or sha256(output)!=stage_hash
                    or outcomes[0]['sha256']!=stage_hash):
                raise RuntimeError('Recovered copy or preserved stage did not match expected bytes.')
            if outcomes[0]['query_count']!=1 or len(recovered_inventory)!=1:
                raise RuntimeError(f'Recovery inspection did not find exactly one query: {recovered_inventory!r}')
            recovered_query=recovered_inventory[0]
            if recovered_query.name!=QUERY_NAME or recovered_query.formula!=FORMULA:
                raise RuntimeError(f'Recovered query formula/name did not verify: {recovered_query!r}')
            try:
                for pid,owned_handle in owned_processes:
                    if win32event.WaitForSingleObject(owned_handle,10000)!=win32event.WAIT_OBJECT_0:
                        raise RuntimeError(f'Recovery-owned Excel process did not exit: {pid}')
            finally:
                for _pid,process_handle in owned_processes:
                    win32api.CloseHandle(process_handle)
            print(f'PASS: staged copy recovery; container={workbook_path.suffix}; Excel {version} build {build}; '
                  f'queries={outcomes[0]["query_count"]}; exact query formula reopened read-only; '
                  f'original unchanged; stage preserved; output hash verified; '
                  f'{len(owned_processes)} recovery-owned Excel processes exited; seed PID {excel_pid} exited.')
    finally:
        if book is not None:
            try:book.Close(False)
            except Exception:pass
        if excel is not None:
            try:excel.Quit()
            except Exception:pass
        if handle is not None:win32api.CloseHandle(handle)
        pythoncom.CoUninitialize()


if __name__=='__main__':
    raise SystemExit(main())
