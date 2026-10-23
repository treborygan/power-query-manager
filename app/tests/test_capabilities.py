from pq_installer.core.excel_capabilities import probe_excel_capabilities


class Query:
    def __init__(self,name,formula):self.Name=name;self.Formula=formula
class Queries:
    def __init__(self):self.items={}
    def Add(self,name,formula):self.items[name]=Query(name,formula)
    @property
    def Count(self):return len(self.items)
    def Item(self,index):return self.items[index] if isinstance(index,str) else list(self.items.values())[index-1]
class Connection:
    def __init__(self,name):self.Name=name;self.InModel=False
class Connections:
    def __init__(self):self.items={};self.fail=False
    def Add2(self,name,*args):
        if self.fail:raise RuntimeError('Add2 unavailable')
        self.items[name]=Connection(name);return self.items[name]
    @property
    def Count(self):return len(self.items)
    def Item(self,index):return self.items[index] if isinstance(index,str) else list(self.items.values())[index-1]
class Sheet:
    class ListObjects:Count=0
class Sheets:
    Count=1
    def Item(self,index):return Sheet()
class Workbook:
    def __init__(self):self.Queries=Queries();self.Connections=Connections();self.Worksheets=Sheets();self.ReadOnly=True;self.saved=False
    def SaveAs(self,*args,**kwargs):self.path=args[0]
    def Save(self):self.saved=True
    def Close(self,*args):pass
class Workbooks:
    def __init__(self):self.book=None;self.fail_connection=False
    def Add(self):
        self.book=Workbook();self.book.Connections.fail=self.fail_connection;return self.book
    def Open(self,*args,**kwargs):return self.book
class Excel:
    Version='16.0';Build='20326.0';AutomationSecurity=1;Hwnd=123
    def __init__(self):self.Workbooks=Workbooks();self.Visible=True;self.DisplayAlerts=True;self.quit=False
    def Quit(self):self.quit=True


def test_connection_mode_is_enabled_only_after_distinct_connection_reopen_probe(tmp_path):
    excel=Excel();cap=probe_excel_capabilities(excel_factory=lambda:excel,staging_parent=tmp_path)
    assert cap.office_version=='16.0' and cap.build=='20326.0'
    assert cap.supports('definition_only') and cap.supports('connection_only')
    assert cap.supports('preserve_existing')
    assert not cap.supports('new_sheet') and not cap.supports('existing_sheet') and not cap.supports('data_model')
    assert all(cap.reasons.get(mode) for mode in ('new_sheet','existing_sheet','data_model'))
    assert excel.quit and list(tmp_path.iterdir())==[]


def test_failed_connection_probe_keeps_definition_mode_and_reports_reason(tmp_path):
    excel=Excel()
    excel.Workbooks.fail_connection=True
    cap=probe_excel_capabilities(excel_factory=lambda:excel,staging_parent=tmp_path)
    assert cap.supports('definition_only') and not cap.supports('connection_only')
    assert cap.reasons['connection_only'] and excel.quit


def test_capability_timeout_stops_probe_and_disables_unverified_modes(tmp_path,monkeypatch):
    from pq_installer.core import excel_capabilities as module
    from pq_installer.core.excel_lifetime import ExcelOperationTimeout
    excel=Excel()
    calls=[]
    def connection_probe(workbook,name,*,model=False):
        calls.append(name)
        if model:
            raise ExcelOperationTimeout('Data Model capability probe',60,process_exited=True)
        return Connection(f'Query - {name}')
    monkeypatch.setattr(module,'_query_connection',connection_probe)
    cap=probe_excel_capabilities(excel_factory=lambda:excel,staging_parent=tmp_path)
    assert not any(cap.supports(mode) for mode in
        ('definition_only','connection_only','new_sheet','existing_sheet','data_model'))
    assert all('timed out' in cap.reasons[mode].lower() for mode in
        ('definition_only','connection_only','new_sheet','existing_sheet','data_model'))
    assert not excel.quit and list(tmp_path.iterdir())==[]


def test_real_com_capability_probe_bounds_load_save_and_cleanup_calls(monkeypatch):
    import sys
    from types import SimpleNamespace
    from pq_installer.core import excel_capabilities as module
    class OwnedAdapter:
        requires_com_initialization=True
        def create_application(self):return excel
    class RefreshableConnection(Connection):
        def Refresh(self):return None
    excel=Excel()
    calls=[]
    monkeypatch.setitem(sys.modules,'pythoncom',SimpleNamespace(CoInitialize=lambda:None,CoUninitialize=lambda:None))
    monkeypatch.setattr(module,'owned_excel_pid',lambda _excel:43921)
    process=SimpleNamespace(pid=43921,close=lambda:None)
    monkeypatch.setattr(module,'open_owned_excel_process',lambda pid,**kwargs:process)
    def bounded(pid,callback,timeout,*,operation,**kwargs):
        calls.append((pid,timeout,operation))
        return callback()
    monkeypatch.setattr(module,'call_with_process_timeout',bounded)
    def connection(workbook,name,*,model=False):
        return RefreshableConnection(f'Query - {name}')
    monkeypatch.setattr(module,'_query_connection',connection)
    cap=probe_excel_capabilities(excel_adapter=OwnedAdapter())
    assert cap.office_version=='16.0'
    assert all(pid==43921 and timeout==module.PROBE_CALL_TIMEOUT_SECONDS for pid,timeout,_ in calls)
    operations={operation for _,_,operation in calls}
    assert 'Excel capability workbook save' in operations
    assert 'Excel capability verification open' in operations
    assert 'Data Model capability refresh' in operations
    assert 'Excel capability process shutdown' in operations and excel.quit


def test_quit_com_error_is_accepted_when_exact_owned_excel_pid_has_exited(monkeypatch):
    import sys
    from types import SimpleNamespace
    from pq_installer.core import excel_capabilities as module
    class OwnedAdapter:
        requires_com_initialization=True
        def create_application(self):return excel
    excel=Excel()
    def fail_quit():raise RuntimeError('COM proxy disconnected during Quit')
    excel.Quit=fail_quit
    monkeypatch.setitem(sys.modules,'pythoncom',SimpleNamespace(CoInitialize=lambda:None,CoUninitialize=lambda:None))
    monkeypatch.setattr(module,'owned_excel_pid',lambda _excel:43921)
    process=SimpleNamespace(pid=43921,close=lambda:None)
    monkeypatch.setattr(module,'open_owned_excel_process',lambda pid,**kwargs:process)
    monkeypatch.setattr(module,'call_with_process_timeout',lambda _pid,callback,_timeout,**_kwargs:callback())
    monkeypatch.setattr(module,'wait_for_owned_excel_exit',lambda owned,timeout_ms:owned is process)

    cap=probe_excel_capabilities(excel_adapter=OwnedAdapter())

    assert cap.supports('definition_only')
    assert cap.supports('preserve_existing')
    assert 'owned-process cleanup' not in cap.reasons.get('definition_only','')


def test_quit_com_error_terminates_only_owned_excel_if_it_remains_alive(monkeypatch):
    import sys
    from types import SimpleNamespace
    from pq_installer.core import excel_capabilities as module
    class OwnedAdapter:
        requires_com_initialization=True
        def create_application(self):return excel
    excel=Excel()
    def fail_quit():raise RuntimeError('COM proxy disconnected during Quit')
    excel.Quit=fail_quit
    monkeypatch.setitem(sys.modules,'pythoncom',SimpleNamespace(CoInitialize=lambda:None,CoUninitialize=lambda:None))
    monkeypatch.setattr(module,'owned_excel_pid',lambda _excel:43921)
    process=SimpleNamespace(pid=43921,close=lambda:None)
    monkeypatch.setattr(module,'open_owned_excel_process',lambda pid,**kwargs:process)
    monkeypatch.setattr(module,'call_with_process_timeout',lambda _pid,callback,_timeout,**_kwargs:callback())
    waits=iter((False,True));terminated=[]
    monkeypatch.setattr(module,'wait_for_owned_excel_exit',lambda owned,timeout_ms:next(waits))
    monkeypatch.setattr(module,'terminate_owned_excel',lambda owned:(terminated.append(owned) or True))

    cap=probe_excel_capabilities(excel_adapter=OwnedAdapter())

    assert terminated==[process]
    assert cap.supports('definition_only')
    assert 'owned-process cleanup' not in cap.reasons.get('definition_only','')
