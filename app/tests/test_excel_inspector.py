from types import SimpleNamespace
import pytest
from pq_installer.core.excel_inspector import ExcelQueryInspector


class Queries:
    Count=2
    def Item(self, index):
        if index==1:return SimpleNamespace(Name='Alpha',Formula='let x=1 in x',Description='desc')
        if index==2:
            class Broken:
                Name='Beta'; Description=''
                @property
                def Formula(self): raise RuntimeError('unavailable')
            return Broken()
        raise IndexError(index)


class ReadOnlyWorkbook:
    Queries=Queries()
    ReadOnly=True
    def Close(self,save): assert save is False


class Workbooks:
    def __init__(self):self.args=None
    def Open(self,*args,**kwargs):self.args=(args,kwargs);return ReadOnlyWorkbook()


class Excel:
    def __init__(self):self.Workbooks=Workbooks();self.DisplayAlerts=True;self.Visible=True;self.AutomationSecurity=2;self.closed=False
    def Quit(self):self.closed=True


def test_read_only_inspection_reports_formulas_and_unknown_loads_without_save_or_refresh(tmp_path):
    excel=Excel(); path=tmp_path/'book.xlsx';path.write_bytes(b'disposable')
    result=ExcelQueryInspector(excel_factory=lambda:excel).inspect(path)
    assert result[0].name=='Alpha' and result[0].formula=='let x=1 in x'
    assert result[0].load_state=='unknown' and result[1].formula is None
    args,kwargs=excel.Workbooks.args
    assert kwargs.get('ReadOnly') is True and kwargs.get('UpdateLinks')==0
    assert kwargs.get('Password')=='' and kwargs.get('WriteResPassword')=='' and kwargs.get('AddToMru') is False
    assert excel.closed and excel.AutomationSecurity==2


class ComCollection:
    def __init__(self, items): self._items = tuple(items)
    @property
    def Count(self): return len(self._items)
    def Item(self, index): return self._items[index - 1]


def test_load_state_inspection_distinguishes_definition_connection_sheet_and_model():
    def connection(name, in_model=False):
        return SimpleNamespace(Name=name, InModel=in_model)
    connections = ComCollection((
        connection('Query - Connection'),
        connection('Query - Sheet'),
        connection('Query - Model', True),
    ))
    def table(name, address):
        return SimpleNamespace(
            SourceType=3,
            QueryTable=SimpleNamespace(WorkbookConnection=SimpleNamespace(Name=name)),
            Range=SimpleNamespace(Address=address),
        )
    sheets = ComCollection((SimpleNamespace(
        Name='Results',
        ListObjects=ComCollection((table('Query - Sheet', '$B$3:$D$8'),)),
    ),))
    workbook = SimpleNamespace(Connections=connections, Worksheets=sheets)
    states = ExcelQueryInspector._load_states(workbook, ('Definition', 'Connection', 'Sheet', 'Model'))
    assert states == {
        'Definition': 'definition_only',
        'Connection': 'connection_only',
        'Sheet': 'worksheet: Results!$B$3:$D$8',
        'Model': 'data_model',
    }


def test_load_state_stays_unknown_when_workbook_inspection_is_incomplete():
    workbook = SimpleNamespace(Connections=ComCollection(()))
    states = ExcelQueryInspector._load_states(workbook, ('Query',))
    assert states == {'Query': 'unknown'}


def test_load_state_does_not_infer_connection_only_when_table_mapping_fails():
    connection=SimpleNamespace(Name='Query - Sales',InModel=False)
    class UnreadableTable:
        @property
        def QueryTable(self):raise RuntimeError('QueryTable unavailable')
    workbook=SimpleNamespace(
        Connections=ComCollection((connection,)),
        Worksheets=ComCollection((SimpleNamespace(Name='Output',ListObjects=ComCollection((UnreadableTable(),))),)))
    assert ExcelQueryInspector._load_states(workbook,('Sales',))=={'Sales':'unknown'}


def test_load_state_ignores_plain_range_table_when_mapping_power_query_loads():
    connection=SimpleNamespace(Name='Query - Sales',InModel=False)
    class PlainRangeTable:
        SourceType=1
        @property
        def QueryTable(self):raise AssertionError('xlSrcRange has no query mapping')
    workbook=SimpleNamespace(
        Connections=ComCollection((connection,)),
        Worksheets=ComCollection((SimpleNamespace(Name='Output',ListObjects=ComCollection((PlainRangeTable(),))),)))
    assert ExcelQueryInspector._load_states(workbook,('Sales',))=={'Sales':'connection_only'}


def test_load_state_does_not_overwrite_duplicate_connection_names():
    connections=ComCollection((SimpleNamespace(Name='Query - Q',InModel=False),
                               SimpleNamespace(Name='query - q',InModel=True)))
    workbook=SimpleNamespace(Connections=connections,Worksheets=ComCollection(()))
    assert ExcelQueryInspector._load_states(workbook,('Q',))=={'Q':'unknown'}


def test_read_only_inspector_includes_observed_worksheet_destination(tmp_path):
    query = SimpleNamespace(Name='Sales', Formula='let x=1 in x', Description='')
    connection = SimpleNamespace(Name='Query - Sales', InModel=False)
    table = SimpleNamespace(
        SourceType=3,
        QueryTable=SimpleNamespace(WorkbookConnection=SimpleNamespace(Name='Query - Sales')),
        Range=SimpleNamespace(Address='$A$1:$C$4'),
    )
    workbook = SimpleNamespace(
        Queries=ComCollection((query,)),
        Connections=ComCollection((connection,)),
        Worksheets=ComCollection((SimpleNamespace(Name='Output', ListObjects=ComCollection((table,))),)),
        ReadOnly=True,
        Close=lambda save: None,
    )
    excel = SimpleNamespace(
        Workbooks=SimpleNamespace(Open=lambda *args, **kwargs: workbook),
        DisplayAlerts=True,
        Visible=True,
        AutomationSecurity=2,
        Quit=lambda: None,
    )
    path = tmp_path / 'loads.xlsx'
    path.write_bytes(b'disposable')
    result = ExcelQueryInspector(excel_factory=lambda: excel).inspect(path)
    assert result[0].load_state == 'worksheet: Output!$A$1:$C$4'


def test_production_inspector_bounds_query_collection_and_propagates_timeout(tmp_path, monkeypatch):
    import sys
    from types import SimpleNamespace
    import pq_installer.core.excel_inspector as inspector_module
    from pq_installer.core.excel_lifetime import ExcelOperationTimeout

    class QueryCollection:
        @property
        def Count(self):
            return 1
        def Item(self, _index):
            return SimpleNamespace(Name='Q', Formula='let x=1 in x', Description='')

    workbook = SimpleNamespace(Queries=QueryCollection(), ReadOnly=True, Close=lambda _save: None)
    excel = SimpleNamespace(
        Hwnd=42, Visible=True, DisplayAlerts=True, AutomationSecurity=2,
        Workbooks=SimpleNamespace(Open=lambda *_args, **_kwargs: workbook),
        Quit=lambda: None,
    )
    adapter = SimpleNamespace(requires_com_initialization=True,
                              create_application=lambda: excel)
    monkeypatch.setitem(sys.modules, 'pythoncom', SimpleNamespace(CoInitialize=lambda: None,
                                                                  CoUninitialize=lambda: None))
    operations = []

    def bounded(_excel, callback, _timeout, *, enabled, operation):
        assert enabled is True
        operations.append(operation)
        if operation.endswith('workbook query collection count'):
            raise ExcelOperationTimeout(operation, 60, process_exited=True)
        return callback()

    monkeypatch.setattr(inspector_module, 'call_with_timeout', bounded)
    monkeypatch.setattr(inspector_module, 'quit_owned_excel',
        lambda _excel, *, quit_call, after_quit=None, timeout_ms=15000:
            (quit_call(), after_quit() if after_quit else None, False)[-1])
    path = tmp_path / 'blocked-inspection.xlsx'
    path.write_bytes(b'disposable')
    with pytest.raises(ExcelOperationTimeout, match='query collection count'):
        ExcelQueryInspector(excel_adapter=adapter).inspect(path)
    assert any(operation.endswith('workbook query collection count') for operation in operations)
