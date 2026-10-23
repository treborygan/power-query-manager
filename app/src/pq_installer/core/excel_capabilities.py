"""Probe Excel load behavior in a disposable workbook using only static inline M data."""
from dataclasses import dataclass
from pathlib import Path
import tempfile
import gc
from .redaction import safe_error
from .excel_adapter import resolve_excel_application_adapter, resolve_excel_workbook_operations
from .excel_lifetime import (call_with_process_timeout, owned_excel_pid,
                             open_owned_excel_process, terminate_owned_excel, wait_for_owned_excel_exit,
                             ExcelOperationTimeout)

PROBE_CALL_TIMEOUT_SECONDS = 60


@dataclass(frozen=True)
class ExcelCapabilities:
    office_version: str
    build: str
    supported_load_modes: tuple[str, ...]
    reasons: dict

    def supports(self, mode):
        return mode in self.supported_load_modes


def _query_connection(workbook, name, *, model=False):
    connection_string = (
        'OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;'
        f'Location={name};Extended Properties=""'
    )
    return workbook.Connections.Add2(
        f'Query - {name}', f'Power Query capability probe ({name})', connection_string,
        f'SELECT * FROM [{name}]', 2, bool(model), False,
    )


def _sheet_listobjects(workbook):
    for index in range(1, int(workbook.Worksheets.Count) + 1):
        sheet = workbook.Worksheets.Item(index)
        for table_index in range(1, int(sheet.ListObjects.Count) + 1):
            yield sheet, sheet.ListObjects.Item(table_index)


def probe_excel_capabilities(excel_factory=None, *, excel_adapter=None, workbook_operations=None,
                             staging_parent=None):
    """Return modes observed after a save/close/read-only-reopen probe.

    Worksheet and Data Model probes evaluate only a built-in static `#table` expression;
    no user query, network source, or external data connection is used.
    """
    adapter = resolve_excel_application_adapter(adapter=excel_adapter, factory=excel_factory)
    workbook_ops = resolve_excel_workbook_operations(workbook_operations)
    default = adapter.requires_com_initialization
    excel_factory = adapter.create_application

    initialized = False
    excel = workbook = None
    excel_pid = None
    owned_process = None
    owned_process_exited = False
    security = None
    cleanup_error = None
    version = build = 'unknown'
    supported = []
    reasons = {}
    names = {
        'definition_only': 'PQM_Probe_Definition',
        'connection_only': 'PQM_Probe_Connection',
        'new_sheet': 'PQM_Probe_Sheet',
        'existing_sheet': 'PQM_Probe_Existing_Sheet',
        'data_model': 'PQM_Probe_Model',
    }
    formula = 'let Source = #table({"Value"}, {{1}, {2}}) in Source'

    def bounded_probe_call(callback, operation):
        if not default:
            return callback()
        if not excel_pid:
            raise RuntimeError('Excel process ID unavailable; capability probe cannot safely bound its COM calls.')
        return call_with_process_timeout(excel_pid, callback, PROBE_CALL_TIMEOUT_SECONDS,
            operation=operation, owned_process=owned_process)

    with tempfile.TemporaryDirectory(prefix='pqm-capability-', dir=staging_parent) as temp:
        path = Path(temp) / 'probe.xlsx'
        try:
            if default:
                import pythoncom
                pythoncom.CoInitialize()
                initialized = True
            excel = excel_factory()
            if default:
                hwnd = int(excel.Hwnd)
                excel_pid = owned_excel_pid(excel)
                owned_process = open_owned_excel_process(excel_pid, hwnd=hwnd)
            excel.Visible = False
            excel.DisplayAlerts = False
            version, build = bounded_probe_call(lambda: (str(excel.Version), str(excel.Build)),
                                                 'Excel capability version detection')
            try:
                security = excel.AutomationSecurity
                excel.AutomationSecurity = 3
            except Exception as exc:
                raise RuntimeError('Excel macro security could not be set to force-disable; capability probe stopped.') from exc
            workbook = bounded_probe_call(lambda: workbook_ops.add_workbook(excel),
                'Excel capability workbook creation')
            bounded_probe_call(lambda: workbook_ops.save_workbook_as(workbook, path, 51),
                'Excel capability workbook save')

            workbook.Queries.Add(names['definition_only'], formula)
            workbook.Queries.Add(names['connection_only'], formula)
            try:
                bounded_probe_call(lambda: _query_connection(workbook, names['connection_only']),
                                   'connection-only capability creation')
            except Exception as exc:
                if isinstance(exc, ExcelOperationTimeout):
                    raise
                reasons['connection_only'] = f'Excel connection probe failed: {type(exc).__name__}: {safe_error(exc)}'

            workbook.Queries.Add(names['new_sheet'], formula)
            try:
                sheet = workbook.Worksheets.Add(After=workbook.Worksheets.Item(workbook.Worksheets.Count))
                sheet.Name = 'PQM Capability Output'
                table = sheet.ListObjects.Add(3,
                    'OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;'
                    f'Location={names["new_sheet"]};Extended Properties=""',
                    True, 1, sheet.Range('A1'))
                query_table = table.QueryTable
                query_table.CommandType = 2
                query_table.CommandText = f'SELECT * FROM [{names["new_sheet"]}]'
                query_table.BackgroundQuery = False
                bounded_probe_call(lambda: query_table.Refresh(False), 'new worksheet load capability refresh')
                query_table.WorkbookConnection.Name = f'Query - {names["new_sheet"]}'
            except Exception as exc:
                if isinstance(exc, ExcelOperationTimeout):
                    raise
                reasons['new_sheet'] = f'Excel worksheet load probe failed: {type(exc).__name__}: {safe_error(exc)}'

            workbook.Queries.Add(names['existing_sheet'], formula)
            try:
                sheet = workbook.Worksheets.Item(1)
                table = sheet.ListObjects.Add(3,
                    'OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;'
                    f'Location={names["existing_sheet"]};Extended Properties=""',
                    True, 1, sheet.Range('B3'))
                query_table = table.QueryTable
                query_table.CommandType = 2
                query_table.CommandText = f'SELECT * FROM [{names["existing_sheet"]}]'
                query_table.BackgroundQuery = False
                bounded_probe_call(lambda: query_table.Refresh(False), 'existing worksheet load capability refresh')
                query_table.WorkbookConnection.Name = f'Query - {names["existing_sheet"]}'
            except Exception as exc:
                if isinstance(exc, ExcelOperationTimeout):
                    raise
                reasons['existing_sheet'] = f'Excel existing-sheet load probe failed: {type(exc).__name__}: {safe_error(exc)}'

            workbook.Queries.Add(names['data_model'], formula)
            try:
                model_connection = bounded_probe_call(
                    lambda: _query_connection(workbook, names['data_model'], model=True),
                    'Data Model capability creation')
                bounded_probe_call(model_connection.Refresh, 'Data Model capability refresh')
                if not bool(model_connection.InModel):
                    raise RuntimeError('Connection.InModel was false after refresh.')
            except Exception as exc:
                if isinstance(exc, ExcelOperationTimeout):
                    raise
                reasons['data_model'] = f'Excel Data Model load probe failed: {type(exc).__name__}: {safe_error(exc)}'

            bounded_probe_call(lambda: workbook_ops.save_workbook(workbook), 'Excel capability workbook save')
            bounded_probe_call(lambda: workbook_ops.close_workbook(workbook, False),
                'Excel capability workbook close')
            workbook = None
            workbook = bounded_probe_call(lambda: workbook_ops.open_workbook(excel, path, UpdateLinks=0, ReadOnly=True,
                Password='', WriteResPassword='', IgnoreReadOnlyRecommended=True, AddToMru=False),
                'Excel capability verification open')
            if not bool(getattr(workbook, 'ReadOnly', True)):
                raise RuntimeError('Capability probe could not reopen its workbook read-only.')

            try:
                if str(workbook.Queries.Item(names['connection_only']).Formula) != formula:
                    raise RuntimeError('Static query formula did not persist.')
                connection = workbook.Connections.Item(f'Query - {names["connection_only"]}')
                if bool(connection.InModel):
                    raise RuntimeError('Connection-only query was not distinct from model/table loads.')
                if any(str(table.QueryTable.WorkbookConnection.Name).casefold() ==
                       f'query - {names["connection_only"]}'.casefold()
                       for _sheet, table in _sheet_listobjects(workbook)):
                    raise RuntimeError('Connection-only query unexpectedly has a worksheet table.')
                supported.append('connection_only')
            except Exception as exc:
                if isinstance(exc, ExcelOperationTimeout):
                    raise
                reasons.setdefault('connection_only', f'Connection did not persist after reopen: {type(exc).__name__}: {safe_error(exc)}')

            try:
                query = workbook.Queries.Item(names['new_sheet'])
                expected = f'Query - {names["new_sheet"]}'.casefold()
                table_location = None
                for sheet, table in _sheet_listobjects(workbook):
                    if str(table.QueryTable.WorkbookConnection.Name).casefold() == expected:
                        table_location = (str(sheet.Name), str(table.Range.Cells(1, 1).Address))
                if str(query.Formula) != formula or table_location != ('PQM Capability Output', '$A$1'):
                    raise RuntimeError(f'Worksheet query did not persist at the expected destination: {table_location!r}.')
                supported.append('new_sheet')
            except Exception as exc:
                if isinstance(exc, ExcelOperationTimeout):
                    raise
                reasons.setdefault('new_sheet', f'Worksheet table did not persist after reopen: {type(exc).__name__}: {safe_error(exc)}')

            try:
                query = workbook.Queries.Item(names['existing_sheet'])
                expected = f'Query - {names["existing_sheet"]}'.casefold()
                location = None
                for sheet, table in _sheet_listobjects(workbook):
                    if str(table.QueryTable.WorkbookConnection.Name).casefold() == expected:
                        location = (str(sheet.Name), str(table.Range.Cells(1, 1).Address))
                if str(query.Formula) != formula or location != (str(workbook.Worksheets.Item(1).Name), '$B$3'):
                    raise RuntimeError(f'Existing-sheet query did not persist at the requested cell: {location!r}.')
                supported.append('existing_sheet')
            except Exception as exc:
                if isinstance(exc, ExcelOperationTimeout):
                    raise
                reasons.setdefault('existing_sheet', f'Existing-sheet table did not persist after reopen: {type(exc).__name__}: {safe_error(exc)}')

            try:
                query = workbook.Queries.Item(names['data_model'])
                connection = workbook.Connections.Item(f'Query - {names["data_model"]}')
                if str(query.Formula) != formula or not bool(connection.InModel):
                    raise RuntimeError('Data Model connection or query formula did not persist.')
                tables = tuple(workbook.Model.ModelTables.Item(i)
                               for i in range(1, int(workbook.Model.ModelTables.Count) + 1))
                if not any(str(table.SourceWorkbookConnection.Name).casefold() ==
                           f'query - {names["data_model"]}'.casefold() for table in tables):
                    raise RuntimeError('No persisted Data Model table points to the probe query connection.')
                supported.append('data_model')
            except Exception as exc:
                if isinstance(exc, ExcelOperationTimeout):
                    raise
                reasons.setdefault('data_model', f'Data Model load did not persist after reopen: {type(exc).__name__}: {safe_error(exc)}')

            try:
                query = workbook.Queries.Item(names['connection_only'])
                connection = workbook.Connections.Item(f'Query - {names["connection_only"]}')
                if str(query.Formula) != formula or bool(connection.InModel):
                    raise RuntimeError('Connection-only formula or InModel state was not preserved.')
                if any(str(table.QueryTable.WorkbookConnection.Name).casefold() ==
                       f'query - {names["connection_only"]}'.casefold() for _sheet, table in _sheet_listobjects(workbook)):
                    raise RuntimeError('Connection-only query unexpectedly has a worksheet table.')
                if 'connection_only' not in supported:
                    supported.append('connection_only')
            except Exception as exc:
                if isinstance(exc, ExcelOperationTimeout):
                    raise
                reasons.setdefault('connection_only', f'Connection-only state did not persist after reopen: {type(exc).__name__}: {safe_error(exc)}')

            if str(workbook.Queries.Item(names['definition_only']).Formula) == formula:
                supported.append('definition_only')
            else:
                reasons['definition_only'] = 'Excel did not preserve the definition-only query formula.'
            # This mode never creates or converts a load; it keeps whatever Excel already has.
            supported.append('preserve_existing')
        except Exception as exc:
            if isinstance(exc, ExcelOperationTimeout):
                owned_process_exited = bool(exc.process_exited)
                supported.clear()
                reason = f'Excel capability probe timed out: {safe_error(exc)}'
                for mode in ('definition_only', 'connection_only', 'new_sheet', 'existing_sheet', 'data_model', 'preserve_existing'):
                    reasons[mode] = reason
            else:
                reasons.setdefault('definition_only', f'Excel definition probe failed: {type(exc).__name__}: {safe_error(exc)}')
                for mode in ('connection_only', 'new_sheet', 'existing_sheet', 'data_model'):
                    reasons.setdefault(mode, 'Definition probe failed; load mode is unavailable.')
        finally:
            if workbook is not None:
                if not owned_process_exited:
                    try:
                        bounded_probe_call(lambda: workbook_ops.close_workbook(workbook, False),
                            'Excel capability cleanup close')
                    except Exception:
                        pass
                workbook = None
            if excel is not None:
                connection = model_connection = table = query_table = sheet = tables = model_tables = query = model_table = None
                gc.collect()
                if security is not None and not owned_process_exited:
                    try:
                        bounded_probe_call(lambda: setattr(excel, 'AutomationSecurity', security),
                                           'Excel capability macro-security restore')
                    except Exception:
                        pass
                if not owned_process_exited:
                    quit_error=None
                    try:
                        bounded_probe_call(excel.Quit, 'Excel capability process shutdown')
                    except Exception as exc:
                        quit_error=exc
                    if excel_pid:
                        try:
                            exited=wait_for_owned_excel_exit(owned_process,5000)
                            if not exited:
                                # Probe Excel owns only the disposable workbook. If
                                # graceful Quit stalls or its COM proxy disconnected,
                                # terminate only the PID captured from this DispatchEx.
                                terminate_owned_excel(owned_process)
                                exited=wait_for_owned_excel_exit(owned_process,1000)
                            if not exited:
                                raise TimeoutError('Owned Excel process did not exit after shutdown.')
                        except Exception as exc:
                            cleanup_error=type(exc).__name__
                    elif quit_error is not None:
                        cleanup_error=type(quit_error).__name__
                if owned_process is not None:
                    try:
                        owned_process.close()
                    except Exception as exc:
                        cleanup_error = type(exc).__name__
            if initialized:
                try:
                    import pythoncom
                    pythoncom.CoUninitialize()
                except Exception:
                    pass
    if cleanup_error:
        supported.clear()
        reason = f'Excel capability probe could not confirm owned-process cleanup ({cleanup_error}); all modes are disabled.'
        for mode in ('definition_only', 'connection_only', 'new_sheet', 'existing_sheet', 'data_model', 'preserve_existing'):
            reasons[mode] = reason
    return ExcelCapabilities(version, build, tuple(dict.fromkeys(supported)), reasons)
