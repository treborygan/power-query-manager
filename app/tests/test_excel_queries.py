from pathlib import Path
from pq_installer.core.models import QuerySource, ExistingQueryMode, QueryStatus
from pq_installer.core.excel_queries import ExcelQueryInstaller
import pq_installer.core.excel_queries as module

def fake_workbook_path(tmp_path, name):
    path=tmp_path/name
    if not path.exists():path.write_bytes(b'disposable workbook')
    return path

class FakeQuery:
    def __init__(self, coll, name, formula='old'): self.coll,self.name,self.Formula=coll,name,formula
    def Delete(self): self.coll.items.pop(self.name,None)

class FakeMissingQueryError(Exception):
    def __init__(self,name):
        super().__init__(-2147352567,'Exception occurred.',
            (0,'Microsoft.Mashup.Client.Excel',f"Cannot find a query by the name '{name}'.",None,0,-2147319765),None)

class FakeMissingConnectionError(Exception):
    def __init__(self):
        super().__init__(-2147352567,'Exception occurred.',(0,None,None,None,0,-2147352565),None)

class FakeQueries:
    def __init__(self,names=()): self.items={n:FakeQuery(self,n) for n in names}; self.added=[]; self.fail_name=None
    def Item(self,name):
        if name not in self.items: raise FakeMissingQueryError(name)
        return self.items[name]
    def Add(self,name,formula):
        if name==self.fail_name: raise RuntimeError('simulated COM Add failure')
        self.added.append((name,formula)); q=FakeQuery(self,name,formula); self.items[name]=q; return q
class FakeWorkbook:
    def __init__(self,names=()): self.Queries=FakeQueries(names); self.Connections=FakeConnections(); self.Worksheets=type('Sheets',(),{'Count':0})(); self.saved=False; self.closed=None
    def Save(self): self.saved=True
    def Close(self,save): self.closed=save
class FakeWorkbooks:
    def __init__(self,wb): self.wb=wb
    def Open(self,path,**kwargs): return self.wb
class FakeExcel:
    def __init__(self,wb): self.Workbooks=FakeWorkbooks(wb); self.DisplayAlerts=True; self.Visible=True; self.AutomationSecurity=1; self.Hwnd=99; self.quit=False
    def Quit(self): self.quit=True
class FakeFactory:
    def __init__(self,excel):self.excel=excel
    def __call__(self):return self.excel

class FakeConnection:
    def __init__(self,coll,name):self.coll,self.Name,self.InModel=coll,name,False
    def Delete(self):self.coll.items.pop(self.Name,None)
class FakeConnections:
    def __init__(self):self.items={};self.fail=False;self.calls=[]
    @property
    def Count(self):return len(self.items)
    def Add2(self,*args):
        self.calls.append(args)
        if self.fail:raise RuntimeError('connection add failed')
        name=args[0]
        if name in self.items:raise RuntimeError('duplicate connection')
        connection=FakeConnection(self,name);self.items[name]=connection;return connection
    def Item(self,name):
        if isinstance(name,str):
            if name not in self.items:raise FakeMissingConnectionError()
            return self.items[name]
        return list(self.items.values())[name-1]

def test_create_and_save_query_only(tmp_path):
    m=tmp_path/'q.m'; m.write_text('let x=1 in x')
    wb=FakeWorkbook(); excel=FakeExcel(wb)
    r=ExcelQueryInstaller(FakeFactory(excel)).install(fake_workbook_path(tmp_path, 'book.xlsx'),[QuerySource(m,'q')],ExistingQueryMode.REPLACE)
    assert r[0].status==QueryStatus.CREATED
    assert wb.Queries.added==[('q','let x=1 in x')]
    assert wb.saved and wb.closed is False and excel.quit


def test_installer_uses_injected_workbook_lifecycle_operations(tmp_path):
    source=tmp_path/'q.m';source.write_text('let x=1 in x')
    workbook=FakeWorkbook();excel=FakeExcel(workbook);calls=[]
    class WorkbookOperations:
        def add_workbook(self, application):
            return application.Workbooks.wb
        def open_workbook(self, application, path, **options):
            calls.append(('open', Path(path), options.get('ReadOnly')))
            return application.Workbooks.Open(str(path), **options)
        def save_workbook(self, workbook):
            calls.append(('save', workbook));workbook.Save()
        def save_workbook_as(self, workbook, path, file_format):
            calls.append(('save_as', Path(path), file_format));workbook.SaveAs(str(path), file_format)
        def close_workbook(self, workbook, save_changes=False):
            calls.append(('close', workbook, save_changes));workbook.Close(save_changes)

    result=ExcelQueryInstaller(FakeFactory(excel),workbook_operations=WorkbookOperations()).install(
        fake_workbook_path(tmp_path,'lifecycle.xlsx'),[QuerySource(source,'q')],ExistingQueryMode.REPLACE)

    assert result[0].status==QueryStatus.CREATED
    assert [row[0] for row in calls]==['open','save','close','open','close']
    assert calls[0][2] is False and calls[3][2] is True

def test_excel_trailing_newline_canonicalization_does_not_reject_source(tmp_path):
    source=tmp_path/'q.m'
    original=b'\xef\xbb\xbf// header\r\nlet x=1 in x\r\n'
    source.write_bytes(original)
    workbook=FakeWorkbook()
    class ExcelCanonicalizingQueries(FakeQueries):
        def Add(self,name,formula):
            self.added.append((name,formula))
            query=FakeQuery(self,name,formula.rstrip('\r\n'))
            self.items[name]=query
            return query
    workbook.Queries=ExcelCanonicalizingQueries()
    result=ExcelQueryInstaller(FakeFactory(FakeExcel(workbook))).install(
        fake_workbook_path(tmp_path,'book.xlsx'),[QuerySource(source,'q')],ExistingQueryMode.REPLACE)
    assert result[0].status==QueryStatus.CREATED
    assert workbook.Queries.added==[('q','// header\r\nlet x=1 in x\r\n')]
    assert workbook.Queries.items['q'].Formula=='// header\r\nlet x=1 in x'
    assert source.read_bytes()==original

def test_excel_formula_verification_still_rejects_internal_content_changes():
    expected='// header\r\nlet x=1 in x\r\n'
    assert module._formula_matches_excel_roundtrip(expected,'// header\r\nlet x=1 in x')
    assert not module._formula_matches_excel_roundtrip(expected,'// header\nlet x=1 in x')
    assert not module._formula_matches_excel_roundtrip(expected,'// header\r\nlet x=2 in x')

def test_staged_verification_open_uses_bounded_excel_call(tmp_path,monkeypatch):
    calls=[]
    real_call=module.call_with_timeout
    def observe(excel,callback,timeout_seconds,*,enabled,operation):
        calls.append(operation)
        return real_call(excel,callback,timeout_seconds,enabled=enabled,operation=operation)
    monkeypatch.setattr(module,'call_with_timeout',observe)
    source=tmp_path/'q.m';source.write_text('let x=1 in x')
    ExcelQueryInstaller(FakeFactory(FakeExcel(FakeWorkbook()))).install(
        fake_workbook_path(tmp_path,'book.xlsx'),[QuerySource(source,'q')],ExistingQueryMode.REPLACE)
    assert 'configure Excel and read import macro security' in calls
    assert 'force-disable macros for import' in calls
    assert 'staged workbook open' in calls
    assert 'add query q' in calls
    assert 'staged workbook save' in calls
    assert 'staged workbook verification open' in calls
    assert 'verify staged workbook query formulas and loads' in calls
    assert 'close staged workbook verification' in calls
    assert 'quit owned Excel instance after import' in calls


def test_importer_bounded_call_uses_runtime_open_timeout(monkeypatch):
    timeouts=[]
    monkeypatch.setattr(module,'OPEN_TIMEOUT_SECONDS',7)
    def observe(excel,callback,timeout_seconds,*,enabled,operation):
        timeouts.append(timeout_seconds)
        return callback()
    monkeypatch.setattr(module,'call_with_timeout',observe)
    installer=ExcelQueryInstaller(FakeFactory(FakeExcel(FakeWorkbook())))
    assert installer._bounded_excel_call(object(),lambda:'ok','runtime timeout')=='ok'
    assert timeouts==[7]


def test_query_lookup_failure_is_not_mistaken_for_absence():
    class UnavailableQueryCollection:
        def Item(self,_name):
            raise RuntimeError('COM bridge unavailable')
    workbook=type('Workbook',(),{'Queries':UnavailableQueryCollection()})()
    with __import__('pytest').raises(RuntimeError,match='could not be checked'):
        ExcelQueryInstaller._get_existing_query(workbook,'Q')


def test_connection_lookup_failure_is_not_mistaken_for_absence():
    class UnavailableConnectionCollection:
        def Item(self,_name):
            raise RuntimeError('COM bridge unavailable')
    workbook=type('Workbook',(),{'Connections':UnavailableConnectionCollection()})()
    with __import__('pytest').raises(RuntimeError,match='lookup could not be checked'):
        ExcelQueryInstaller._get_existing_connection(workbook,'Query - Q')


def test_import_timeout_after_staged_query_mutation_discards_stage_and_suppresses_com(tmp_path,monkeypatch):
    from pq_installer.core.excel_lifetime import ExcelOperationTimeout
    calls=[]
    def timeout_after_add(excel,callback,timeout_seconds,*,enabled,operation):
        calls.append(operation)
        if operation=='add query q':
            callback()  # Excel may apply the mutation before the timeout is observed.
            raise ExcelOperationTimeout(operation,timeout_seconds,termination_error='AccessDenied')
        return callback()
    monkeypatch.setattr(module,'call_with_timeout',timeout_after_add)
    source=tmp_path/'q.m';source.write_text('let x=1 in x')
    workbook_path=fake_workbook_path(tmp_path,'book.xlsx')
    original=workbook_path.read_bytes()
    workbook=FakeWorkbook()
    excel=FakeExcel(workbook)
    with __import__('pytest').raises(module.InstallOperationError,match='add query q'):
        ExcelQueryInstaller(FakeFactory(excel)).install(
            workbook_path,[QuerySource(source,'q')],ExistingQueryMode.REPLACE)
    assert calls[-1]=='add query q'
    assert workbook.Queries.items['q'].Formula=='let x=1 in x'
    assert workbook_path.read_bytes()==original
    assert not excel.quit
    assert not list(tmp_path.glob('*.pqstage-*'))

def test_failed_stage_cleanup_is_reported_without_masking_import_error(tmp_path):
    from pq_installer.core.excel_transaction import LocalWorkbookTransactionAdapter
    class RetainedStageAdapter(LocalWorkbookTransactionAdapter):
        def cleanup_stage(self,_stage):return False
    source=tmp_path/'q.m';source.write_text('let x=1 in x')
    workbook=FakeWorkbook()
    def fail_save():raise RuntimeError('primary save failure')
    workbook.Save=fail_save
    logs=[]
    with __import__('pytest').raises(__import__('pq_installer.core.excel_queries',fromlist=['InstallOperationError']).InstallOperationError,
                                     match='primary save failure'):
        ExcelQueryInstaller(FakeFactory(FakeExcel(workbook)),
            workbook_transaction_adapter=RetainedStageAdapter(check_active=lambda _path:None)).install(
                fake_workbook_path(tmp_path,'book.xlsx'),[QuerySource(source,'q')],
                ExistingQueryMode.REPLACE,logger=logs.append)
    assert any('temporary staged workbook remains' in str(item).lower() and '.pqstage-' in str(item) for item in logs)


def test_final_commit_failure_never_reports_staged_queries_as_installed(tmp_path):
    from pq_installer.core.excel_transaction import LocalWorkbookTransactionAdapter

    class CommitFailureAdapter(LocalWorkbookTransactionAdapter):
        def commit_stage(self, *_args, **_kwargs):
            raise RuntimeError('injected final commit failure')

    from pq_installer.core.excel_queries import InstallOperationError
    first=tmp_path/'first.m';first.write_text('let x=1 in x')
    second=tmp_path/'second.m';second.write_text('let y=2 in y')
    workbook_path=fake_workbook_path(tmp_path,'commit-failure.xlsx')
    original=workbook_path.read_bytes()
    installer=ExcelQueryInstaller(FakeFactory(FakeExcel(FakeWorkbook())),
        workbook_transaction_adapter=CommitFailureAdapter(check_active=lambda _path:None))

    with __import__('pytest').raises(InstallOperationError,match='injected final commit failure') as caught:
        installer.install(workbook_path,[QuerySource(first,'First'),QuerySource(second,'Second')],
            ExistingQueryMode.REPLACE)

    assert {row.status for row in caught.value.results}=={QueryStatus.ROLLED_BACK}
    assert all(row.status not in (QueryStatus.CREATED,QueryStatus.REPLACED) for row in caught.value.results)
    assert workbook_path.read_bytes()==original
    assert not list(tmp_path.glob('*.pqstage-*'))

def test_verification_workbook_close_failure_prevents_stage_commit(tmp_path):
    from pq_installer.core.excel_queries import InstallOperationError

    class VerificationCloseFailure(FakeWorkbook):
        def __init__(self):
            super().__init__()
            self.close_calls=0
        def Close(self,save):
            self.close_calls+=1
            if self.close_calls==2:
                raise RuntimeError('verification close denied')
            super().Close(save)

    source=tmp_path/'q.m';source.write_text('let x=1 in x')
    workbook_path=fake_workbook_path(tmp_path,'verification-close.xlsx')
    original=workbook_path.read_bytes()
    workbook=VerificationCloseFailure()

    with __import__('pytest').raises(InstallOperationError,match='verification close denied') as caught:
        ExcelQueryInstaller(FakeFactory(FakeExcel(workbook))).install(
            workbook_path,[QuerySource(source,'q')],ExistingQueryMode.REPLACE)

    assert caught.value.results[0].status==QueryStatus.ROLLED_BACK
    assert workbook_path.read_bytes()==original
    assert not list(tmp_path.glob('*.pqstage-*'))

def test_existing_formula_update_preserves_query_object(tmp_path):
    p=tmp_path/'q.pq'; p.write_text('new formula')
    wb=FakeWorkbook(['q']); old_object=wb.Queries.items['q']
    r=ExcelQueryInstaller(FakeFactory(FakeExcel(wb))).install(fake_workbook_path(tmp_path, 'b.xlsx'),[QuerySource(p,'q')],ExistingQueryMode.REPLACE)
    assert r[0].status==QueryStatus.REPLACED
    assert wb.Queries.items['q'] is old_object and old_object.Formula=='new formula'
    assert 'associations were left unchanged' in r[0].message
    assert 'definition only' not in r[0].message


def test_all_skipped_import_does_not_save_or_commit_workbook(tmp_path):
    from pq_installer.core.excel_transaction import LocalWorkbookTransactionAdapter

    source = tmp_path / 'q.m'
    source.write_text('new formula')
    workbook_path = fake_workbook_path(tmp_path, 'book.xlsx')
    workbook_path.write_bytes(b'original workbook bytes')
    workbook = FakeWorkbook(['q'])
    workbook.Queries.items['q'].Formula = 'existing formula'

    class RecordingTransactionAdapter(LocalWorkbookTransactionAdapter):
        def __init__(self):
            super().__init__(check_active=lambda _path: None)
            self.commits = 0

        def commit_stage(self, *args, **kwargs):
            self.commits += 1
            return super().commit_stage(*args, **kwargs)

    transaction = RecordingTransactionAdapter()
    logged = []
    result = ExcelQueryInstaller(FakeFactory(FakeExcel(workbook)),
        workbook_transaction_adapter=transaction).install(
            workbook_path, [QuerySource(source, 'q')], ExistingQueryMode.SKIP, logger=logged.append)

    assert result[0].status == QueryStatus.SKIPPED
    assert workbook.Queries.items['q'].Formula == 'existing formula'
    assert not workbook.saved
    assert transaction.commits == 0
    assert workbook_path.read_bytes() == b'original workbook bytes'
    assert not list(tmp_path.glob('*.pqstage-*'))
    assert logged == result


def test_replacing_loaded_query_reports_that_existing_load_is_preserved(tmp_path):
    p=tmp_path/'q.m';p.write_text('new formula')
    wb=FakeWorkbook(['q'])
    existing_connection=FakeConnection(wb.Connections,'Query - q')
    wb.Connections.items['Query - q']=existing_connection
    old_query=wb.Queries.items['q']
    result=ExcelQueryInstaller(FakeFactory(FakeExcel(wb))).install(
        fake_workbook_path(tmp_path, 'b.xlsx'),[QuerySource(p,'q')],ExistingQueryMode.REPLACE
    )
    assert result[0].status==QueryStatus.REPLACED
    assert wb.Queries.items['q'] is old_query
    assert wb.Connections.items['Query - q'] is existing_connection
    assert 'worksheet' in result[0].message and 'Data Model' in result[0].message
    assert 'No refresh was requested' in result[0].message


def test_load_signature_fails_closed_on_duplicate_power_query_connections():
    from types import SimpleNamespace
    class Indexed:
        def __init__(self,values):self.values=tuple(values);self.Count=len(self.values)
        def Item(self,index):return self.values[index-1]
    workbook=SimpleNamespace(
        Connections=Indexed((SimpleNamespace(Name='Query - Sales',InModel=False),
                             SimpleNamespace(Name='query - sales',InModel=True))),
        Worksheets=Indexed(()))
    with __import__('pytest').raises(RuntimeError,match='Cannot verify existing load associations'):
        ExcelQueryInstaller._query_load_signature(workbook,'Sales')


def test_load_signature_fails_closed_when_a_table_connection_cannot_be_read():
    from types import SimpleNamespace
    class Indexed:
        def __init__(self,values):self.values=tuple(values);self.Count=len(self.values)
        def Item(self,index):return self.values[index-1]
    class BrokenTable:
        @property
        def QueryTable(self):raise RuntimeError('connection link unavailable')
    connection=SimpleNamespace(Name='Query - Sales',InModel=False)
    sheet=SimpleNamespace(Name='Output',ListObjects=Indexed((BrokenTable(),)))
    workbook=SimpleNamespace(Connections=Indexed((connection,)),Worksheets=Indexed((sheet,)))
    with __import__('pytest').raises(RuntimeError,match='Cannot verify existing load associations'):
        ExcelQueryInstaller._query_load_signature(workbook,'Sales')


def test_load_signature_ignores_explicit_plain_range_tables():
    from types import SimpleNamespace
    class Indexed:
        def __init__(self,values):self.values=tuple(values);self.Count=len(self.values)
        def Item(self,index):return self.values[index-1]
    class PlainRangeTable:
        SourceType=1  # Excel xlSrcRange; QueryTable is not available for this ListObject.
        @property
        def QueryTable(self):raise AssertionError('ordinary Excel table must not be treated as a query table')
    connection=SimpleNamespace(Name='Query - Sales',InModel=False)
    sheet=SimpleNamespace(Name='Output',ListObjects=Indexed((PlainRangeTable(),)))
    workbook=SimpleNamespace(Connections=Indexed((connection,)),Worksheets=Indexed((sheet,)))
    assert ExcelQueryInstaller._query_load_signature(workbook,'Sales')==('query - sales',False,())


def test_worksheet_load_mapping_ignores_explicit_plain_range_tables():
    from types import SimpleNamespace
    class Indexed:
        def __init__(self,values):self.values=tuple(values);self.Count=len(self.values)
        def Item(self,index):return self.values[index-1]
    class PlainRangeTable:
        SourceType=1
        @property
        def QueryTable(self):raise AssertionError('plain range has no QueryTable')
    query_table=SimpleNamespace(WorkbookConnection=SimpleNamespace(Name='Query - Sales'))
    worksheet_query=SimpleNamespace(SourceType=3,QueryTable=query_table)
    sheet=SimpleNamespace(Name='Output',ListObjects=Indexed((PlainRangeTable(),worksheet_query)))
    workbook=SimpleNamespace(Worksheets=Indexed((sheet,)))
    assert ExcelQueryInstaller._query_worksheet_tables(workbook,'Sales')==[(sheet,worksheet_query)]


def test_duplicate_query_connections_block_replacement_before_formula_mutation(tmp_path):
    p=tmp_path/'q.m';p.write_text('new formula')
    workbook=FakeWorkbook(['q'])
    old_query=workbook.Queries.items['q']
    for name in ('Query - q','query - Q'):
        workbook.Connections.items[name]=FakeConnection(workbook.Connections,name)
    with __import__('pytest').raises(RuntimeError,match='Cannot verify existing load associations'):
        ExcelQueryInstaller(FakeFactory(FakeExcel(workbook))).install(
            fake_workbook_path(tmp_path, 'book.xlsx'),[QuerySource(p,'q')],ExistingQueryMode.REPLACE)
    assert old_query.Formula=='old'
    assert not workbook.saved

def test_skip_existing(tmp_path):
    p=tmp_path/'q.m'; p.write_text('new')
    wb=FakeWorkbook(['q'])
    r=ExcelQueryInstaller(FakeFactory(FakeExcel(wb))).install(fake_workbook_path(tmp_path, 'b.xlsx'),[QuerySource(p,'q')],ExistingQueryMode.SKIP)
    assert r[0].status==QueryStatus.SKIPPED and wb.Queries.added==[]

def test_definitions_allow_forward_reference_order(tmp_path):
    a=tmp_path/'1_first.m'; b=tmp_path/'2_later.pq'
    a.write_text('let Source = later in Source'); b.write_text('let later=1 in later')
    from pq_installer.core.query_loader import discover_query_files
    sources=discover_query_files([a,b]); wb=FakeWorkbook()
    ExcelQueryInstaller(FakeFactory(FakeExcel(wb))).install(fake_workbook_path(tmp_path, 'b.xlsx'),sources,ExistingQueryMode.REPLACE)
    assert [x[0] for x in wb.Queries.added]==['first','later']
    assert wb.Queries.items['first'].Formula=='let Source = later in Source'

def test_mid_run_failure_rolls_back_only_new_definitions(tmp_path):
    a=tmp_path/'1_a.m'; b=tmp_path/'2_b.m'; c=tmp_path/'3_c.m'; a.write_text('a');b.write_text('b');c.write_text('c')
    wb=FakeWorkbook(['preexisting']); wb.Queries.fail_name='b'
    with __import__('pytest').raises(RuntimeError,match='rolled back 1 change') as caught:
        ExcelQueryInstaller(FakeFactory(FakeExcel(wb))).install(fake_workbook_path(tmp_path, 'b.xlsx'),[__import__('pq_installer.core.models',fromlist=['QuerySource']).QuerySource(a,'a',1),__import__('pq_installer.core.models',fromlist=['QuerySource']).QuerySource(b,'b',2),__import__('pq_installer.core.models',fromlist=['QuerySource']).QuerySource(c,'c',3)],ExistingQueryMode.REPLACE)
    assert 'a' not in wb.Queries.items and 'preexisting' in wb.Queries.items and not wb.saved
    assert [(row.query_name,row.status.value) for row in caught.value.results]==[
        ('a','Rolled back'),('b','Failed'),('c','Not attempted')]
    from pq_installer.core.error_log import format_failure
    report=format_failure(caught.value)
    assert 'a: Rolled back' in report and 'b: Failed' in report and 'c: Not attempted' in report

def test_existing_query_formula_restored_on_later_failure(tmp_path):
    a=tmp_path/'1_a.m'; b=tmp_path/'2_b.m'; a.write_text('changed'); b.write_text('b')
    wb=FakeWorkbook(['a']); wb.Queries.items['a'].Formula='original'; wb.Queries.fail_name='b'
    from pq_installer.core.models import QuerySource
    with __import__('pytest').raises(RuntimeError,match='rolled back 1 change'):
        ExcelQueryInstaller(FakeFactory(FakeExcel(wb))).install(fake_workbook_path(tmp_path, 'b.xlsx'),[QuerySource(a,'a',1),QuerySource(b,'b',2)],ExistingQueryMode.REPLACE)
    assert wb.Queries.items['a'].Formula=='original'

def test_formula_setter_that_mutates_then_raises_is_rolled_back(tmp_path):
    class PartialWriteQuery:
        def __init__(self,coll,name):self.coll,self.name,self.value=coll,name,'original';self.writes=0
        @property
        def Formula(self):return self.value
        @Formula.setter
        def Formula(self,value):
            self.writes+=1;self.value=value
            if self.writes==1:raise RuntimeError('COM setter failed after partial write')
        def Delete(self):self.coll.items.pop(self.name,None)
    source_path=tmp_path/'q.m';source_path.write_text('replacement')
    original_workbook=fake_workbook_path(tmp_path, 'book.xlsx');original_workbook.write_bytes(b'original workbook bytes')
    workbook=FakeWorkbook();query=PartialWriteQuery(workbook.Queries,'q');workbook.Queries.items['q']=query
    with __import__('pytest').raises(RuntimeError,match='COM setter failed after partial write'):
        ExcelQueryInstaller(FakeFactory(FakeExcel(workbook))).install(
            original_workbook,[QuerySource(source_path,'q')],ExistingQueryMode.REPLACE)
    assert workbook.Queries.items['q'] is query
    assert query.Formula=='original'
    assert original_workbook.read_bytes()==b'original workbook bytes'

def test_unverified_formula_rollback_is_reported_as_rollback_error(tmp_path):
    class UnrestorableQuery:
        def __init__(self,coll,name):self.coll,self.name,self.value=coll,name,'original';self.writes=0
        @property
        def Formula(self):return self.value
        @Formula.setter
        def Formula(self,value):
            self.writes+=1
            if self.writes>1:raise RuntimeError('rollback setter denied')
            self.value=value
    source_path=tmp_path/'q.m';source_path.write_text('replacement')
    failing_path=tmp_path/'b.m';failing_path.write_text('new query')
    workbook=FakeWorkbook();workbook.Queries.fail_name='b'
    query=UnrestorableQuery(workbook.Queries,'q');workbook.Queries.items['q']=query
    with __import__('pytest').raises(RuntimeError,match='Rollback errors: q: RuntimeError: rollback setter denied'):
        ExcelQueryInstaller(FakeFactory(FakeExcel(workbook))).install(
            fake_workbook_path(tmp_path, 'book.xlsx'),[QuerySource(source_path,'q',0),QuerySource(failing_path,'b',1)],ExistingQueryMode.REPLACE)
    assert query.Formula=='replacement'

def test_default_installer_initializes_com_for_worker_thread(tmp_path,monkeypatch):
    import sys
    from types import SimpleNamespace
    import pq_installer.core.excel_queries as installer_module
    import pq_installer.core.excel_lifetime as lifetime_module
    monkeypatch.setattr(installer_module,'validate_workbook_for_write',lambda *_args,**_kwargs:None)
    monkeypatch.setattr(installer_module,'ensure_no_active_file_users',lambda *_args,**_kwargs:None)
    calls=[]; monkeypatch.setitem(sys.modules,'pythoncom',SimpleNamespace(CoInitialize=lambda:calls.append('init'),CoUninitialize=lambda:calls.append('uninit')))
    monkeypatch.setitem(sys.modules,'win32process',SimpleNamespace(GetWindowThreadProcessId=lambda _hwnd:(1,999)))
    monkeypatch.setattr(lifetime_module,'open_owned_excel_process',lambda pid,**kwargs:SimpleNamespace(pid=pid,close=lambda:None))
    p=tmp_path/'q.m';p.write_text('let x=1 in x'); wb=FakeWorkbook(); installer=ExcelQueryInstaller(); installer._excel_factory=FakeFactory(FakeExcel(wb))
    installer.install(fake_workbook_path(tmp_path, 'book.xlsx'),[QuerySource(p,'q')],ExistingQueryMode.REPLACE)
    assert calls==['init','uninit']

def test_unreadable_formula_blocks_replace_without_deleting(tmp_path):
    class Unreadable:
        def __init__(self, coll, name): self.coll,self.name=coll,name
        @property
        def Formula(self): raise RuntimeError('access denied')
        def Delete(self): self.coll.items.pop(self.name,None)
    p=tmp_path/'q.m'; p.write_text('replacement')
    wb=FakeWorkbook(['q']); old=Unreadable(wb.Queries,'q'); wb.Queries.items['q']=old
    with __import__('pytest').raises(RuntimeError,match='cannot safely replace'):
        ExcelQueryInstaller(FakeFactory(FakeExcel(wb))).install(fake_workbook_path(tmp_path, 'b.xlsx'),[QuerySource(p,'q')],ExistingQueryMode.REPLACE)
    assert wb.Queries.items['q'] is old and wb.Queries.added==[]

def test_excel_owner_lock_blocks_install_before_excel_is_created(tmp_path):
    source_path=tmp_path/'q.m';source_path.write_text('let x=1 in x')
    book=fake_workbook_path(tmp_path,'locked.xlsx')
    (tmp_path/'~$locked.xlsx').write_bytes(b'Excel owner lock')
    class UnexpectedFactory:
        def __call__(self):raise AssertionError('Excel must not start while the workbook lock exists')
    with __import__('pytest').raises(ValueError,match='appears to be open in Excel'):
        ExcelQueryInstaller(UnexpectedFactory()).install(book,[QuerySource(source_path,'q')],ExistingQueryMode.REPLACE)

def test_backup_names_do_not_collide_within_one_second(tmp_path):
    from pq_installer.core.backup import create_timestamped_backup
    from datetime import datetime
    book=fake_workbook_path(tmp_path, 'book.xlsx'); book.write_bytes(b'original')
    fixed=datetime(2026,1,1,0,0,0)
    a=create_timestamped_backup(book,now=fixed); b=create_timestamped_backup(book,now=fixed)
    assert a != b and a.read_bytes()==b'original' and b.read_bytes()==b'original'

def test_backup_verification_failure_never_publishes_a_partial_backup(tmp_path,monkeypatch):
    import pq_installer.core.backup as backup_module
    book=fake_workbook_path(tmp_path, 'book.xlsx');book.write_bytes(b'original')
    original_copy=backup_module.shutil.copy2
    def corrupt_copy(source,destination):
        result=original_copy(source,destination)
        __import__('pathlib').Path(destination).write_bytes(b'partial/corrupt')
        return result
    monkeypatch.setattr(backup_module.shutil,'copy2',corrupt_copy)
    with __import__('pytest').raises(RuntimeError,match='Backup verification failed'):
        backup_module.create_timestamped_backup(book)
    assert book.read_bytes()==b'original'
    assert not list(tmp_path.glob('book_backup_*.xlsx'))
    assert not list(tmp_path.glob('*.backup-*.tmp'))

def test_backup_name_race_never_overwrites_concurrent_backup(tmp_path,monkeypatch):
    from pathlib import Path
    import pq_installer.core.backup as backup_module
    book=fake_workbook_path(tmp_path,'book.xlsx');book.write_bytes(b'original')
    now=__import__('datetime').datetime(2026,1,1,0,0,0)
    raced=book.with_name('book_backup_20260101_000000.xlsx')
    real_link=backup_module.os.link
    calls=[]
    def collide_once(source,destination):
        calls.append(Path(destination))
        if len(calls)==1:
            Path(destination).write_bytes(b'other backup created concurrently')
            raise FileExistsError(destination)
        return real_link(source,destination)
    monkeypatch.setattr(backup_module.os,'link',collide_once)
    result=backup_module.create_timestamped_backup(book,now=now)
    assert raced.read_bytes()==b'other backup created concurrently'
    assert result.name=='book_backup_20260101_000000_1.xlsx'
    assert result.read_bytes()==b'original'
    assert len(calls)==2

def test_staged_save_failure_keeps_original_workbook_bytes(tmp_path):
    p=tmp_path/'q.m'; p.write_text('formula')
    book=fake_workbook_path(tmp_path, 'book.xlsx'); book.write_bytes(b'original workbook bytes')
    class FailingWorkbook(FakeWorkbook):
        def Save(self): raise RuntimeError('simulated save failure')
    wb=FailingWorkbook()
    with __import__('pytest').raises(RuntimeError,match='simulated save failure'):
        ExcelQueryInstaller(FakeFactory(FakeExcel(wb))).install(book,[QuerySource(p,'q')],ExistingQueryMode.REPLACE)
    assert book.read_bytes()==b'original workbook bytes'
    assert list(tmp_path.glob('*.pqstage-*'))==[]

def test_workbook_opened_by_another_app_before_commit_keeps_original_bytes(tmp_path,monkeypatch):
    p=tmp_path/'q.m';p.write_text('let x=1 in x')
    book=fake_workbook_path(tmp_path,'book.xlsx');book.write_bytes(b'original workbook bytes')
    calls=[]
    def check_users(path):
        calls.append(path)
        if len(calls)==1:raise RuntimeError('The workbook is currently open or in use by another application. Close it before continuing.')
    monkeypatch.setattr(module,'ensure_no_active_file_users',check_users)
    with __import__('pytest').raises(RuntimeError,match='currently open or in use'):
        ExcelQueryInstaller(FakeFactory(FakeExcel(FakeWorkbook()))).install(
            book,[QuerySource(p,'q')],ExistingQueryMode.REPLACE)
    assert book.read_bytes()==b'original workbook bytes'
    assert len(calls)==1
    assert not list(tmp_path.glob('*.pqstage-*'))

def test_source_hash_mismatch_fails_before_excel_factory(tmp_path):
    p=tmp_path/'q.m'; p.write_text('new')
    source=QuerySource(p,'q',source_sha256='0'*64)
    class UnexpectedFactory:
        def __call__(self): raise AssertionError('Excel must not open')
    with __import__('pytest').raises(RuntimeError,match='Source changed after review'):
        ExcelQueryInstaller(UnexpectedFactory()).install(fake_workbook_path(tmp_path, 'b.xlsx'),[source],ExistingQueryMode.REPLACE)

def test_workbook_hash_mismatch_fails_before_excel_open(tmp_path):
    p=tmp_path/'q.m';p.write_text('formula')
    book=fake_workbook_path(tmp_path, 'book.xlsx');book.write_bytes(b'changed workbook')
    class UnexpectedFactory:
        def __call__(self): raise AssertionError('Excel must not open')
    with __import__('pytest').raises(RuntimeError,match='Workbook changed after import review'):
        ExcelQueryInstaller(UnexpectedFactory()).install(book,[QuerySource(p,'q')],ExistingQueryMode.REPLACE,
                                                           expected_workbook_sha256='0'*64)

def test_macro_security_failure_aborts_before_workbook_open(tmp_path):
    p=tmp_path/'q.m';p.write_text('formula')
    class NoMacroSecurity(FakeExcel):
        def __getattribute__(self,name):
            if name == 'AutomationSecurity':raise RuntimeError('policy denied')
            return super().__getattribute__(name)
    wb=FakeWorkbook();excel=NoMacroSecurity(wb)
    with __import__('pytest').raises(RuntimeError,match='macro security could not be set'):
        ExcelQueryInstaller(FakeFactory(excel)).install(fake_workbook_path(tmp_path, 'b.xlsx'),[QuerySource(p,'q')],ExistingQueryMode.REPLACE)
    assert not wb.Queries.added and excel.quit

def test_connection_only_adds_distinct_connection_for_new_query(tmp_path):
    from pq_installer.core.models import LoadMode
    from pq_installer.core.excel_capabilities import ExcelCapabilities
    p=tmp_path/'q.m';p.write_text('let x=1 in x');wb=FakeWorkbook()
    source=QuerySource(p,'q',load_mode=LoadMode.CONNECTION_ONLY)
    installer=ExcelQueryInstaller(FakeFactory(FakeExcel(wb)))
    with __import__('pytest').raises(RuntimeError,match='requires a successful Excel capability probe'):
        installer.install(fake_workbook_path(tmp_path, 'book.xlsx'),[source],ExistingQueryMode.REPLACE)
    cap=ExcelCapabilities('16.0','20326.0',('definition_only','connection_only'),{})
    result=installer.install(fake_workbook_path(tmp_path, 'book.xlsx'),[source],ExistingQueryMode.REPLACE,capabilities=cap)
    assert result[0].status==QueryStatus.CREATED
    assert 'Query - q' in wb.Connections.items
    call=wb.Connections.calls[0]
    assert call[0]=='Query - q' and call[5] is False and call[6] is False
    assert 'Location=q' in call[2] and call[3]=='SELECT * FROM [q]'

def test_connection_only_rejects_connection_string_or_command_delimiters_before_excel(tmp_path):
    from pq_installer.core.models import LoadMode
    p=tmp_path/'q.m';p.write_text('let x=1 in x');wb=FakeWorkbook()
    book=fake_workbook_path(tmp_path, 'book.xlsx')
    class UnexpectedFactory:
        def __call__(self): raise AssertionError('Excel must not open for unsafe connection names')
    installer=ExcelQueryInstaller(UnexpectedFactory())
    for name in ('North;Location=Injected', 'North] WHERE 1=1 --', 'North";Provider=Untrusted'):
        source=QuerySource(p,name,load_mode=LoadMode.CONNECTION_ONLY)
        with __import__('pytest').raises(ValueError,match='unsupported'):
            installer.install(book,[source],ExistingQueryMode.REPLACE)
    assert not wb.Connections.calls
    assert not wb.Queries.added

def test_connection_add_failure_rolls_back_query_definition(tmp_path):
    from pq_installer.core.models import LoadMode
    from pq_installer.core.excel_capabilities import ExcelCapabilities
    p=tmp_path/'q.m';p.write_text('let x=1 in x');wb=FakeWorkbook();wb.Connections.fail=True
    source=QuerySource(p,'q',load_mode=LoadMode.CONNECTION_ONLY)
    cap=ExcelCapabilities('16.0','20326.0',('definition_only','connection_only'),{})
    with __import__('pytest').raises(RuntimeError,match='connection add failed'):
        ExcelQueryInstaller(FakeFactory(FakeExcel(wb))).install(fake_workbook_path(tmp_path, 'book.xlsx'),[source],ExistingQueryMode.REPLACE,capabilities=cap)
    assert 'q' not in wb.Queries.items

def test_existing_load_mode_conversion_fails_closed_without_changing_definition(tmp_path):
    from pq_installer.core.models import LoadMode
    from pq_installer.core.excel_capabilities import ExcelCapabilities
    from pq_installer.core.excel_queries import InstallOperationError
    source_path = tmp_path / 'q.m'
    source_path.write_text('let Value = 2 in Value')
    workbook = FakeWorkbook(['q'])
    existing_query = workbook.Queries.items['q']
    connection = FakeConnection(workbook.Connections, 'Query - q')
    workbook.Connections.items[connection.Name] = connection
    excel = FakeExcel(workbook)
    capabilities = ExcelCapabilities('16.0', '20430.0', ('definition_only', 'new_sheet'), {})

    with __import__('pytest').raises(InstallOperationError, match='refusing to alter it implicitly'):
        ExcelQueryInstaller(FakeFactory(excel)).install(
            fake_workbook_path(tmp_path, 'book.xlsx'),
            [QuerySource(source_path, 'q', load_mode=LoadMode.NEW_SHEET,
                         sheet_name='Output', table_name='tblQ')],
            ExistingQueryMode.REPLACE, capabilities=capabilities,
        )

    assert existing_query.Formula == 'old'
    assert workbook.Queries.items['q'] is existing_query
    assert workbook.Connections.items[connection.Name] is connection
    assert workbook.Queries.added == []
    assert excel.quit

def test_new_worksheet_refuses_to_assume_destination_absent_when_inventory_fails(tmp_path):
    from pq_installer.core.models import LoadMode

    class BrokenNameSheet:
        @property
        def Name(self):
            raise RuntimeError('worksheet name read failed')

    class Worksheets:
        Count=1
        added=False
        def Item(self,key):
            if isinstance(key,int):
                return BrokenNameSheet()
            raise RuntimeError('named worksheet lookup unavailable')
        def Add(self,**_kwargs):
            self.added=True
            raise AssertionError('must not create when worksheet inventory is unknown')

    source=QuerySource(tmp_path/'q.m','q',load_mode=LoadMode.NEW_SHEET,sheet_name='Target')
    worksheets=Worksheets()
    workbook=type('Workbook',(),{'Worksheets':worksheets})()

    with __import__('pytest').raises(RuntimeError,match='worksheet inventory could not be checked'):
        ExcelQueryInstaller._add_worksheet_load(workbook,source)

    assert not worksheets.added

def test_staged_import_refuses_change_to_workbook_privacy_permissions(tmp_path, monkeypatch):
    from pq_installer.core.mashup_privacy import WorkbookPrivacy
    from pq_installer.core.excel_queries import InstallOperationError
    source_path = tmp_path / 'q.m'
    source_path.write_text('let Value = 1 in Value')
    workbook_path = fake_workbook_path(tmp_path, 'book.xlsx')
    original_bytes = workbook_path.read_bytes()
    observations = iter((WorkbookPrivacy(True, None), WorkbookPrivacy(False, None)))
    reads = []

    def read_privacy(path):
        reads.append(path)
        return next(observations)

    monkeypatch.setattr(module, 'read_workbook_privacy', read_privacy)
    workbook = FakeWorkbook()
    excel = FakeExcel(workbook)

    with __import__('pytest').raises(InstallOperationError, match='Power Query privacy settings changed'):
        ExcelQueryInstaller(FakeFactory(excel)).install(
            workbook_path, [QuerySource(source_path, 'q')], ExistingQueryMode.REPLACE,
        )

    assert workbook_path.read_bytes() == original_bytes
    assert len(reads) == 2
    assert excel.quit

def test_sheet_and_model_loads_require_build_capability_probe(tmp_path):
    from pq_installer.core.models import LoadMode
    p=tmp_path/'q.m';p.write_text('let x=1 in x')
    for mode in (LoadMode.NEW_SHEET, LoadMode.EXISTING_SHEET, LoadMode.DATA_MODEL):
        with __import__('pytest').raises(RuntimeError, match='capability probe'):
            ExcelQueryInstaller(FakeFactory(FakeExcel(FakeWorkbook()))).install(
                fake_workbook_path(tmp_path, 'book.xlsx'),[QuerySource(p,'q',load_mode=mode,
                    sheet_name='Dest' if mode == LoadMode.EXISTING_SHEET else None,
                    cell='A1' if mode == LoadMode.EXISTING_SHEET else None)],ExistingQueryMode.REPLACE
            )

def test_staged_install_reopens_read_only_and_verifies_before_commit(tmp_path):
    p=tmp_path/'q.m';p.write_text('let x=1 in x')
    book=fake_workbook_path(tmp_path, 'book.xlsx');book.write_bytes(b'original workbook')
    wb=FakeWorkbook(); excel=FakeExcel(wb)
    class VerifyWorkbook(FakeWorkbook):
        ReadOnly=True
        def __init__(self):super().__init__();self.Queries.items['q']=FakeQuery(self.Queries,'q','wrong formula')
    class Workbooks(FakeWorkbooks):
        def __init__(self,initial):super().__init__(initial);self.opens=0
        def Open(self,path,**kwargs):
            self.opens+=1
            return wb if self.opens==1 else VerifyWorkbook()
    excel.Workbooks=Workbooks(wb)
    logs=[]
    with __import__('pytest').raises(RuntimeError,match='verification'):
        ExcelQueryInstaller(FakeFactory(excel)).install(book,[QuerySource(p,'q')],ExistingQueryMode.REPLACE,logger=logs.append)
    assert book.read_bytes()==b'original workbook'
    assert not any(getattr(row,'status',None)==QueryStatus.CREATED for row in logs)

def test_final_stage_commit_failure_preserves_original_and_verified_backup(tmp_path, monkeypatch):
    import hashlib
    from pathlib import Path
    from pq_installer.core.backup import create_timestamped_backup
    import pq_installer.core.excel_transaction as transaction
    from pq_installer.core.excel_queries import InstallOperationError

    source = tmp_path / 'replacement.m'
    source.write_text('let Value = 2 in Value', encoding='utf-8')
    workbook = fake_workbook_path(tmp_path, 'book.xlsx')
    original = b'original workbook bytes before commit'
    workbook.write_bytes(original)
    backup = create_timestamped_backup(workbook)

    def fail_commit(stage, target, _expected_sha256):
        assert Path(stage).name.startswith('.book.pqstage-')
        assert Path(target) == workbook
        raise PermissionError('simulated final replacement denial')

    monkeypatch.setattr(transaction, '_replace_existing_file', fail_commit)
    with __import__('pytest').raises(InstallOperationError, match='final replacement denial') as failure:
        ExcelQueryInstaller(FakeFactory(FakeExcel(FakeWorkbook()))).install(
            workbook, [QuerySource(source, 'Q')], ExistingQueryMode.REPLACE)
    assert failure.value.results[0].status == QueryStatus.ROLLED_BACK
    assert 'No change was committed' not in failure.value.results[0].message

    assert workbook.read_bytes() == original
    assert hashlib.sha256(backup.read_bytes()).digest() == hashlib.sha256(original).digest()
    assert not list(tmp_path.glob('*.pqstage-*'))
