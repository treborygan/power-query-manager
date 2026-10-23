from pathlib import Path

import pytest

from pq_installer.core.excel_refresh import ExcelQueryRefresher, RefreshOperationError


class Collection:
    def __init__(self, items):
        self.items = list(items)

    @property
    def Count(self):
        return len(self.items)

    def Item(self, key):
        if isinstance(key, int):
            return self.items[key - 1]
        for item in self.items:
            if item.Name.casefold() == str(key).casefold():
                return item
        raise KeyError(key)


class FakeConnection:
    def __init__(self, name, *, stuck=False):
        self.Name = name
        self.InModel = False
        self.stuck = stuck
        self.RefreshCount = 0
        self.CancelCount = 0
        self.OLEDBConnection = self
        self.BackgroundQuery = False
        self.Refreshing = False
        self.RefreshModes = []

    def Refresh(self, background_query=None):
        self.RefreshCount += 1
        self.RefreshModes.append(background_query)
        self.Refreshing = self.stuck
        return True

    def CancelRefresh(self):
        self.CancelCount += 1
        self.Refreshing = False


class Workbook:
    def __init__(self, path, queries, connections):
        self.path = Path(path)
        self.Queries = Collection([type('Query', (), {'Name': name, 'Formula': 'let Value = 1 in Value'})() for name in queries])
        self.Connections = Collection(connections)
        self.ReadOnly = False
        self.ProtectStructure = False
        self.closed = None
        self.saved = False

    def Save(self):
        self.saved = True
        self.path.write_bytes(b'refreshed workbook')

    def Close(self, save):
        self.closed = save


class Workbooks:
    def __init__(self, factory):
        self.factory = factory
        self.opened = None

    def Open(self, path, **kwargs):
        self.opened = Path(path)
        workbook=self.factory(self.opened)
        workbook.ReadOnly=bool(kwargs.get('ReadOnly',False))
        return workbook


class Excel:
    def __init__(self, workbooks):
        self.Workbooks = workbooks
        self.Visible = True
        self.DisplayAlerts = True
        self.AutomationSecurity = 2
        self.quit = False

    def Quit(self):
        self.quit = True


def make_refresh(tmp_path, *, queries=('Sales',), connections=None):
    workbook_path = tmp_path / 'book.xlsx'
    workbook_path.write_bytes(b'original workbook')
    connections = connections if connections is not None else [FakeConnection(f'Query - {name}') for name in queries]
    wb_ref = {}
    books = Workbooks(lambda path: wb_ref.setdefault('workbook', Workbook(path, queries, connections)))
    excel = Excel(books)
    return workbook_path, wb_ref, ExcelQueryRefresher(excel_factory=lambda: excel), excel


def test_refresh_selected_commits_staged_copy_and_creates_backup(tmp_path):
    path, wb_ref, refresher, excel = make_refresh(tmp_path, queries=('Sales', 'Helper'))
    results = refresher.refresh(path, ('Sales',), timeout_seconds=2, asynchronous=False, make_backup=True)
    assert [(item.query_name, item.status) for item in results] == [('Sales', 'refreshed')]
    assert path.read_bytes() == b'refreshed workbook'
    assert list(tmp_path.glob('book_backup_*.xlsx'))[0].read_bytes() == b'original workbook'
    assert wb_ref['workbook'].path != path
    assert wb_ref['workbook'].Connections.Item('Query - Sales').RefreshCount == 1
    assert wb_ref['workbook'].Connections.Item('Query - Helper').RefreshCount == 0
    assert excel.quit and wb_ref['workbook'].closed is False


def test_staged_verification_open_uses_bounded_excel_call(tmp_path, monkeypatch):
    import pq_installer.core.excel_refresh as module
    calls=[]
    real_call=module.call_with_timeout
    def observe(excel,callback,timeout_seconds,*,enabled,operation):
        calls.append(operation)
        return real_call(excel,callback,timeout_seconds,enabled=enabled,operation=operation)
    monkeypatch.setattr(module,'call_with_timeout',observe)
    path, _, refresher, _=make_refresh(tmp_path,queries=('Sales',))
    refresher.refresh(path,('Sales',),timeout_seconds=2,asynchronous=False,make_backup=False)
    assert 'staged workbook refresh open' in calls
    assert 'staged workbook refresh query inventory' in calls
    assert 'staged workbook refresh save' in calls
    assert 'staged workbook refresh verification open' in calls
    assert 'verify staged workbook refresh results' in calls
    assert 'close staged refresh verification workbook' in calls


def test_refresh_bounded_call_uses_runtime_open_timeout(monkeypatch):
    import pq_installer.core.excel_refresh as module
    timeouts=[]
    monkeypatch.setattr(module,'OPEN_TIMEOUT_SECONDS',9)
    def observe(excel,callback,timeout_seconds,*,enabled,operation):
        timeouts.append(timeout_seconds)
        return callback()
    monkeypatch.setattr(module,'call_with_timeout',observe)
    refresher=ExcelQueryRefresher(excel_factory=lambda:None)
    assert refresher._bounded_excel_call(object(),lambda:'ok','runtime timeout')=='ok'
    assert timeouts==[9]


def test_refresh_executes_and_waits_on_mashup_querytable_not_workbook_connection():
    class MashupQueryTable:
        def __init__(self):
            self.BackgroundQuery = False
            self.Refreshing = False
            self.calls = []

        def Refresh(self, background_query):
            self.calls.append(background_query)
            return True

        def CancelRefresh(self):
            self.Refreshing = False

    connection = FakeConnection('Query - Sales')
    connection.Refresh = lambda: (_ for _ in ()).throw(
        AssertionError('WorkbookConnection.Refresh cannot verify Mashup completion'))
    query_table = MashupQueryTable()
    refresher = ExcelQueryRefresher()
    refresher._default = False

    result = refresher._refresh_connection(
        'Sales', connection, query_table=query_table, timeout_seconds=3,
        asynchronous=False, excel_pid=None,
    )

    assert result.status == 'refreshed'
    assert query_table.calls == [False]


def test_refresh_refuses_unconfirmed_com_querytable_refresh(monkeypatch):
    import pq_installer.core.excel_refresh as module
    from types import SimpleNamespace

    monkeypatch.setattr(module, 'call_with_process_timeout',
        lambda _pid, callback, _timeout, **_kwargs: callback())
    monkeypatch.setattr(module, 'open_owned_excel_process',
        lambda _pid, **_kwargs: SimpleNamespace(close=lambda: None))

    class UnconfirmedQueryTable:
        BackgroundQuery = False
        Refreshing = False

        def Refresh(self, _background_query):
            return None

        def CancelRefresh(self):
            pass

    refresher = ExcelQueryRefresher()
    refresher._default = True
    with pytest.raises(RuntimeError, match='did not confirm'):
        refresher._refresh_connection(
            'Sales', FakeConnection('Query - Sales'), query_table=UnconfirmedQueryTable(),
            timeout_seconds=3, asynchronous=False, excel_pid=4242,
        )


def test_data_model_refresh_uses_only_the_matching_model_table():
    class ModelTable:
        def __init__(self, name, connection_name, record_count=1):
            self.Name = name
            self.SourceWorkbookConnection = type('SourceConnection', (), {'Name': connection_name})()
            self.RecordCount = record_count
            self.refresh_calls = 0

        def Refresh(self):
            self.refresh_calls += 1
            self.RecordCount += 1

    selected = ModelTable('SalesTable', 'Query - Sales')
    unrelated = ModelTable('HelperTable', 'Query - Helper')
    workbook = type('Workbook', (), {
        'Model': type('Model', (), {'ModelTables': Collection([unrelated, selected])})(),
    })()
    connection = FakeConnection('Query - Sales')
    connection.InModel = True
    refresher = ExcelQueryRefresher()
    refresher._default = False

    result = refresher._refresh_query(
        workbook, 'Sales', connection, timeout_seconds=3,
        asynchronous=True, excel_pid=None,
    )

    assert result.status == 'refreshed'
    assert selected.refresh_calls == 1
    assert selected.RecordCount == 2
    assert unrelated.refresh_calls == 0
    assert connection.RefreshCount == 0


def test_committed_data_model_refresh_keeps_synchronous_api_note(tmp_path):
    class ModelTable:
        Name = 'SalesTable'
        RecordCount = 2

        def __init__(self):
            self.SourceWorkbookConnection = type('SourceConnection', (), {'Name': 'Query - Sales'})()

        def Refresh(self):
            self.RecordCount = 2

    table = ModelTable()
    path, workbook_ref, refresher, excel = make_refresh(tmp_path, queries=('Sales',))
    connection = FakeConnection('Query - Sales')
    connection.InModel = True

    def factory(open_path):
        workbook = Workbook(open_path, ('Sales',), [connection])
        workbook.Model = type('Model', (), {'ModelTables': Collection([table])})()
        workbook_ref['workbook'] = workbook
        return workbook

    excel.Workbooks.factory = factory

    result = refresher.refresh(path, ('Sales',), timeout_seconds=2,
        asynchronous=True, make_backup=False)[0]

    assert result.status == 'refreshed'
    assert 'synchronously' in result.message
    assert 'background option' in result.message
    assert 'committed' in result.message


@pytest.mark.parametrize('connections', [(), ('Query - Sales', 'Query - Sales')])
def test_data_model_refresh_refuses_missing_or_ambiguous_model_table(connections):
    class ModelTable:
        def __init__(self, name):
            self.Name = name
            self.SourceWorkbookConnection = type('SourceConnection', (), {'Name': name})()
            self.RecordCount = 1
            self.refresh_calls = 0

        def Refresh(self):
            self.refresh_calls += 1

    tables = [ModelTable(name) for name in connections]
    workbook = type('Workbook', (), {
        'Model': type('Model', (), {'ModelTables': Collection(tables)})(),
    })()
    connection = FakeConnection('Query - Sales')
    connection.InModel = True
    refresher = ExcelQueryRefresher()
    refresher._default = False

    with pytest.raises(RuntimeError, match='Could not identify exactly one Data Model table'):
        refresher._refresh_query(
            workbook, 'Sales', connection, timeout_seconds=3,
            asynchronous=False, excel_pid=None,
        )

    assert all(table.refresh_calls == 0 for table in tables)


def test_multiple_data_model_tables_require_all_model_sources_to_be_selected():
    class ModelTable:
        def __init__(self, connection_name):
            self.SourceWorkbookConnection = type('SourceConnection', (), {'Name': connection_name})()

    tables = [ModelTable('Query - Sales'), ModelTable('Query - Helper')]

    with pytest.raises(RuntimeError, match='select all Data Model queries'):
        ExcelQueryRefresher._plan_data_model_refresh(('Query - Sales',), tables)


def test_all_data_model_tables_can_be_refreshed_as_one_explicit_group():
    class ModelTable:
        def __init__(self, connection_name):
            self.SourceWorkbookConnection = type('SourceConnection', (), {'Name': connection_name})()

    tables = [ModelTable('Query - Sales'), ModelTable('Query - Helper')]

    strategy, mapped = ExcelQueryRefresher._plan_data_model_refresh(
        ('Query - Sales', 'Query - Helper'), tables)

    assert strategy == 'all'
    assert len(mapped) == 2


def test_single_data_model_table_supports_single_query_refresh():
    class ModelTable:
        SourceWorkbookConnection = type('SourceConnection', (), {'Name': 'Query - Sales'})()

    strategy, mapped = ExcelQueryRefresher._plan_data_model_refresh(
        ('Query - Sales',), (ModelTable(),))

    assert strategy == 'single'
    assert len(mapped) == 1


def test_temporary_refresh_target_cleanup_unhides_activates_and_removes_only_owned_objects():
    calls = []

    class Sheet:
        Visible = 2

        def Delete(self):
            calls.append('delete-sheet')

    class ActiveSheet:
        def Activate(self):
            calls.append('activate-original-sheet')

    class QueryTable:
        def Delete(self):
            calls.append('delete-query-table')

    class Connection:
        def Delete(self):
            calls.append('delete-generated-connection')

    refresher = ExcelQueryRefresher()
    refresher._default = False
    scratch_sheet = Sheet()
    refresher._cleanup_temporary_refresh_target(
        QueryTable(), (scratch_sheet, Connection(), ActiveSheet()), excel_pid=None,
        query_name='Sales',
    )

    assert calls == [
        'delete-query-table', 'activate-original-sheet', 'delete-sheet',
        'delete-generated-connection',
    ]
    assert scratch_sheet.Visible == -1


def test_refresh_connection_bounds_background_mode_and_completion_properties(monkeypatch):
    import pq_installer.core.excel_refresh as module
    from types import SimpleNamespace

    calls=[]
    def bounded(pid, callback, timeout, *, operation, terminate, **kwargs):
        calls.append((pid, timeout, operation))
        return callback()
    monkeypatch.setattr(module, 'call_with_process_timeout', bounded)
    monkeypatch.setattr(module, 'open_owned_excel_process',lambda pid,**kwargs:SimpleNamespace(pid=pid,close=lambda:None))
    class CompletesAsync(FakeConnection):
        def __init__(self):
            super().__init__('Query - Sales')
            self.refresh_reads = 0
            self.ResultRange = type('ResultRange', (), {'Rows': type('Rows', (), {'Count': 2})()})()

        @property
        def Refreshing(self):
            self.refresh_reads += 1
            return self.refresh_reads == 1

        @Refreshing.setter
        def Refreshing(self, value):
            self._refreshing = value

        def Refresh(self, background_query=None):
            self.RefreshModes.append(background_query)
            self.refresh_reads = 0
            return True

    connection=CompletesAsync()
    refresher=ExcelQueryRefresher()
    refresher._default=True

    outcome=refresher._refresh_connection('Sales', connection, timeout_seconds=3,
        asynchronous=True, excel_pid=4242)

    assert outcome.status=='refreshed'
    operations=[item[2] for item in calls]
    assert 'read Sales background refresh setting' in operations
    assert 'set Sales background refresh setting' in operations
    assert 'refresh for Sales' in operations
    assert 'poll Sales refresh completion' in operations
    assert 'restore Sales background refresh setting' in operations
    assert all(item[0]==4242 for item in calls)


def test_refresh_property_timeout_propagates_and_does_not_touch_dead_excel_again(monkeypatch):
    import pq_installer.core.excel_refresh as module
    from types import SimpleNamespace
    from pq_installer.core.excel_lifetime import ExcelOperationTimeout

    calls=[]
    def bounded(pid, callback, timeout, *, operation, terminate, **kwargs):
        calls.append(operation)
        if operation == 'poll Sales refresh completion':
            raise ExcelOperationTimeout(operation, timeout, process_exited=True)
        return callback()
    monkeypatch.setattr(module, 'call_with_process_timeout', bounded)
    monkeypatch.setattr(module, 'open_owned_excel_process',lambda pid,**kwargs:SimpleNamespace(pid=pid,close=lambda:None))
    connection=FakeConnection('Query - Sales', stuck=True)
    refresher=ExcelQueryRefresher()
    refresher._default=True

    with pytest.raises(ExcelOperationTimeout, match='poll Sales refresh completion'):
        refresher._refresh_connection('Sales', connection, timeout_seconds=30,
            asynchronous=True, excel_pid=4242)

    assert 'restore Sales background refresh setting' not in calls


def test_refresh_inventory_is_bounded_and_timeout_discards_stage(tmp_path, monkeypatch):
    import pq_installer.core.excel_refresh as module
    from pq_installer.core.excel_lifetime import ExcelOperationTimeout

    calls=[]
    real_call=module.call_with_timeout
    def bounded(excel, callback, timeout, *, enabled, operation):
        calls.append(operation)
        if operation == 'staged workbook refresh query inventory':
            raise ExcelOperationTimeout(operation, timeout, termination_error='AccessDenied')
        return real_call(excel, callback, timeout, enabled=enabled, operation=operation)
    monkeypatch.setattr(module, 'call_with_timeout', bounded)
    path, workbook_ref, refresher, excel=make_refresh(tmp_path)
    original=path.read_bytes()

    with pytest.raises(RefreshOperationError, match='query inventory'):
        refresher.refresh(path, ('Sales',), timeout_seconds=2, asynchronous=False,
                          make_backup=False)

    assert 'staged workbook refresh query inventory' in calls
    assert calls[-1]=='staged workbook refresh query inventory'
    assert path.read_bytes()==original
    assert not list(tmp_path.glob('*.pqstage-*'))
    assert workbook_ref['workbook'].saved is False


def test_failed_stage_cleanup_is_reported_without_masking_refresh_error(tmp_path):
    from pq_installer.core.excel_transaction import LocalWorkbookTransactionAdapter
    class RetainedStageAdapter(LocalWorkbookTransactionAdapter):
        def cleanup_stage(self,_stage):return False
    connection=FakeConnection('Query - Sales')
    connection.Refresh=lambda *_args:(_ for _ in ()).throw(RuntimeError('primary refresh failure'))
    path, _, refresher, _=make_refresh(tmp_path,connections=[connection])
    refresher._workbook_transactions=RetainedStageAdapter(check_active=lambda _path:None)
    logs=[]
    with pytest.raises(RefreshOperationError,match='primary refresh failure'):
        refresher.refresh(path,('Sales',),timeout_seconds=2,asynchronous=False,
                          make_backup=False,logger=logs.append)
    assert any('temporary staged workbook remains' in str(item).lower() and '.pqstage-' in str(item) for item in logs)


def test_async_refresh_timeout_cancels_and_preserves_original_bytes(tmp_path):
    stuck = FakeConnection('Query - Sales', stuck=True)
    path, wb_ref, refresher, _ = make_refresh(tmp_path, connections=[stuck])
    original = path.read_bytes()
    ticks = iter([0.0, 0.0, 0.0, 0.0, 2.0])
    refresher._monotonic = lambda: next(ticks, 2.0)
    refresher._sleep = lambda _seconds: None
    with pytest.raises(RefreshOperationError, match='timed out'):
        refresher.refresh(path, ('Sales',), timeout_seconds=1, asynchronous=True, make_backup=False)
    assert path.read_bytes() == original
    assert stuck.CancelCount == 1
    assert not list(tmp_path.glob('*.pqstage-*'))
    assert wb_ref['workbook'].saved is False


def test_synchronous_timeout_targets_only_the_owned_excel_pid(monkeypatch):
    import pq_installer.core.excel_refresh as module
    from types import SimpleNamespace
    import time
    connection=FakeConnection('Query - Sales')
    connection.Refresh=lambda *_args:time.sleep(0.05)
    refresher=ExcelQueryRefresher()
    terminated=[]
    monkeypatch.setattr(module,'open_owned_excel_process',lambda pid,**kwargs:SimpleNamespace(pid=pid,close=lambda:None))
    monkeypatch.setattr(refresher,'_terminate_owned_excel',lambda pid:(terminated.append(pid) or True))
    with pytest.raises(TimeoutError,match='process exit was confirmed'):
        refresher._refresh_connection('Sales',connection,timeout_seconds=0.01,
                                      asynchronous=False,excel_pid=47291)
    assert len(terminated)==1 and terminated[0].pid==47291


def test_synchronous_timeout_reports_owned_process_termination_failure(monkeypatch):
    import pq_installer.core.excel_refresh as module
    from types import SimpleNamespace
    import time
    connection=FakeConnection('Query - Sales')
    connection.Refresh=lambda *_args:time.sleep(0.05)
    refresher=ExcelQueryRefresher()
    monkeypatch.setattr(module,'open_owned_excel_process',lambda pid,**kwargs:SimpleNamespace(pid=pid,close=lambda:None))
    def fail_termination(_pid):raise PermissionError('protected process')
    monkeypatch.setattr(refresher,'_terminate_owned_excel',fail_termination)
    with pytest.raises(TimeoutError,match='termination attempt failed') as caught:
        refresher._refresh_connection('Sales',connection,timeout_seconds=0.01,
                                      asynchronous=False,excel_pid=47291)
    assert getattr(caught.value,'termination_error',None)=='PermissionError'


def test_definition_only_query_is_reported_skipped_without_workbook_write(tmp_path):
    path, wb_ref, refresher, _ = make_refresh(tmp_path, queries=('Sales',), connections=[])
    original = path.read_bytes()
    results = refresher.refresh(path, ('Sales',), timeout_seconds=2, asynchronous=True, make_backup=True)
    assert results[0].status == 'skipped'
    assert 'No refreshable Power Query connection' in results[0].message
    assert path.read_bytes() == original
    assert wb_ref['workbook'].saved is False
    assert not list(tmp_path.glob('book_backup_*.xlsx'))


def test_refresh_all_reports_refreshable_and_definition_only_queries(tmp_path):
    path, _, refresher, _ = make_refresh(
        tmp_path, queries=('Sales', 'Helper'), connections=[FakeConnection('Query - Sales')]
    )
    results = refresher.refresh(path, timeout_seconds=2, asynchronous=True, make_backup=False)
    assert [(result.query_name, result.status) for result in results] == [
        ('Sales', 'refreshed'), ('Helper', 'skipped')
    ]
    assert path.read_bytes() == b'refreshed workbook'


def test_unknown_selected_query_fails_without_mutating_workbook(tmp_path):
    path, _, refresher, _ = make_refresh(tmp_path, queries=('Sales',))
    original = path.read_bytes()
    with pytest.raises(RefreshOperationError, match='not found in workbook inventory'):
        refresher.refresh(path, ('Missing',), timeout_seconds=2)
    assert path.read_bytes() == original


def test_refresh_failure_rolls_back_staged_changes(tmp_path):
    class Broken(FakeConnection):
        def Refresh(self, _background_query=None):
            raise RuntimeError('provider rejected refresh; PWD=topsecret')
    connection = Broken('Query - Sales')
    path, wb_ref, refresher, _ = make_refresh(tmp_path, connections=[connection])
    original = path.read_bytes()
    with pytest.raises(RefreshOperationError, match='provider rejected refresh') as error:
        refresher.refresh(path, ('Sales',), timeout_seconds=2, asynchronous=False, make_backup=False)
    assert 'topsecret' not in str(error.value)
    assert path.read_bytes() == original
    assert wb_ref['workbook'].saved is False
    assert not list(tmp_path.glob('*.pqstage-*'))


def test_success_is_not_reported_until_staged_save_and_commit_succeed(tmp_path):
    path, wb_ref, refresher, _ = make_refresh(tmp_path, queries=('Sales',))
    original = path.read_bytes()
    wb_ref.clear()
    def fail_save(_path):
        workbook = Workbook(_path, ('Sales',), [FakeConnection('Query - Sales')])
        def fail(): raise RuntimeError('save failed')
        workbook.Save = fail
        return workbook
    refresher._factory = lambda: Excel(Workbooks(fail_save))
    events=[]
    with pytest.raises(RefreshOperationError, match='save failed') as error:
        refresher.refresh(path, ('Sales',), timeout_seconds=2, asynchronous=False,
                          make_backup=False, logger=events.append)
    assert path.read_bytes() == original
    assert any(item.status=='in_progress' for item in events)
    assert not any(item.status=='refreshed' for item in events)
    assert any(item.status=='rolled_back' for item in error.value.results)


def test_concurrent_workbook_change_is_never_overwritten_by_refresh(tmp_path):
    path, _, refresher, _ = make_refresh(tmp_path, queries=('Sales',))
    original = path.read_bytes()
    newer = b'concurrent workbook edit'
    def factory(stage_path):
        workbook=Workbook(stage_path,('Sales',),[FakeConnection('Query - Sales')])
        def save():
            workbook.saved=True
            Path(stage_path).write_bytes(b'refreshed stage')
            path.write_bytes(newer)
        workbook.Save=save
        return workbook
    refresher._factory=lambda:Excel(Workbooks(factory))
    with pytest.raises(RefreshOperationError,match='changed during refresh'):
        refresher.refresh(path,('Sales',),timeout_seconds=2,asynchronous=False,make_backup=False)
    assert path.read_bytes()==newer and path.read_bytes()!=original
    assert not list(tmp_path.glob('*.pqstage-*'))


def test_refresh_rejects_symbolic_link_target_before_excel_starts(tmp_path):
    real=tmp_path/'real.xlsx';real.write_bytes(b'real workbook bytes')
    link=tmp_path/'selected.xlsx'
    try:
        link.symlink_to(real)
    except OSError as exc:
        pytest.skip(f'Windows symbolic-link creation unavailable: {exc}')
    class UnexpectedFactory:
        def __call__(self):raise AssertionError('Excel must not start for a symbolic-link workbook')
    with pytest.raises(ValueError,match='symbolic link'):
        ExcelQueryRefresher(UnexpectedFactory()).refresh(link,('Sales',),make_backup=False)
    assert real.read_bytes()==b'real workbook bytes' and link.is_symlink()
