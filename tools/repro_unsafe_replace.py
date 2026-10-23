#!/usr/bin/env python3
"""Baseline unsafe replacement reproducer; no real Excel or workbook is touched.

Run BEFORE fix: `python tools/repro_unsafe_replace.py` (expect 'BUG REPRODUCED').
Run AFTER fix:  `python tools/repro_unsafe_replace.py --expect-fixed`
Expected fixed behavior: existing query is not deleted when old Formula is None.
"""
import argparse
import sys
import tempfile
from pathlib import Path
ROOT=Path(__file__).resolve().parent.parent
sys.path.insert(0,str(ROOT/'app/src'))
from pq_installer.core.models import QuerySource, ExistingQueryMode
from pq_installer.core.excel_queries import ExcelQueryInstaller

class Existing:
    name='KeepMe'
    def __init__(self): self.deleted=False
    @property
    def Formula(self): return None  # COM returned no recoverable old formula
    def Delete(self): self.deleted=True
class Queries:
    def __init__(self): self.existing=Existing()
    def Item(self,name): return self.existing if name=='KeepMe' else None
    def Add(self,name,formula): raise RuntimeError('simulated COM Add failure')
class Workbook:
    def __init__(self): self.Queries=Queries()
    def Close(self,save): pass
class Workbooks:
    def __init__(self,workbook): self.workbook=workbook
    def Open(self,*args,**kwargs): return self.workbook
class Excel:
    def __init__(self,workbook): self.Workbooks=Workbooks(workbook)
    def Quit(self): pass

def main():
    parser=argparse.ArgumentParser();parser.add_argument('--expect-fixed',action='store_true');args=parser.parse_args()
    wb=Workbook()
    with tempfile.TemporaryDirectory() as temp:
        p=Path(temp)/'KeepMe.pq';p.write_text('let x=1 in x',encoding='utf-8')
        engine=ExcelQueryInstaller(excel_factory=lambda:Excel(wb))
        try: engine.install(Path(temp)/'unused.xlsx',[QuerySource(p,'KeepMe')],ExistingQueryMode.REPLACE)
        except Exception as e: print('Observed simulated failure:',type(e).__name__,str(e)[:180])
        else: print('Unexpectedly completed even though replacement Add would fail')
    deleted=wb.Queries.existing.deleted
    print('Original query Delete() called:',deleted)
    if args.expect_fixed:
        print('RESULT:', 'FIX VERIFIED IN SIMULATION' if not deleted else 'DEFECT REMAINS')
        return 0 if not deleted else 1
    print('RESULT:', 'BUG REPRODUCED' if deleted else 'BEHAVIOR APPEARS FIXED')
    return 0 if deleted else 1
if __name__=='__main__':raise SystemExit(main())
