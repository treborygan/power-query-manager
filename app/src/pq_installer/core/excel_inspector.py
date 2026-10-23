"""Read-only Excel query inventory. All calls must run on a COM initialized worker thread."""
from pathlib import Path
from .manager import QueryInfo
from .validation import validate_workbook
from .excel_lifetime import (call_with_timeout, OPEN_TIMEOUT_SECONDS,
                             ExcelOperationTimeout, quit_owned_excel)
from .redaction import safe_error
from .excel_adapter import resolve_excel_application_adapter, resolve_excel_workbook_operations

_XL_SRC_RANGE = 1


class ExcelQueryInspector:
    def __init__(self, excel_factory=None, *, excel_adapter=None, workbook_operations=None):
        self._adapter = resolve_excel_application_adapter(adapter=excel_adapter, factory=excel_factory)
        self._workbook_operations = resolve_excel_workbook_operations(workbook_operations)
        self._default = self._adapter.requires_com_initialization
        self._factory = self._adapter.create_application

    @staticmethod
    def _items(collection, call=None, label='Excel collection'):
        """Iterate Excel COM collections without relying on Python COM iteration."""
        call = call or (lambda callback, _operation: callback())
        count = int(call(lambda: collection.Count, f'{label} count'))
        return tuple(call(lambda index=index: collection.Item(index), f'{label} item {index}')
                     for index in range(1, count + 1))

    @classmethod
    def _load_states(cls, workbook, query_names, *, call=None):
        """Infer only load states supported by directly observed workbook objects."""
        read = call or (lambda callback, _operation: callback())
        try:
            connections_collection = read(lambda: workbook.Connections, 'workbook connections')
            connections = cls._items(connections_collection, read, 'workbook connections')
            connections_by_name = {}
            for connection in connections:
                name = read(lambda: connection.Name, 'connection name')
                connections_by_name.setdefault(str(name).casefold(), []).append(connection)
            sheet_loads = {}
            incomplete_connections=set()
            incomplete_table_link=False
            # Access failure means the mapping is unknown; never infer connection-only
            # merely because worksheet enumeration was unavailable.
            sheets_collection = read(lambda: workbook.Worksheets, 'workbook worksheets')
            sheets = cls._items(sheets_collection, read, 'workbook worksheets')
            for sheet in sheets:
                list_objects = read(lambda: sheet.ListObjects, 'worksheet tables')
                for table in cls._items(list_objects, read, 'worksheet tables'):
                    try:
                        # A normal range-backed ListObject has no QueryTable. Excel's
                        # xlSrcRange value is an explicit non-query classification.
                        source_type = read(lambda: table.SourceType, 'table source type')
                        if int(source_type) == _XL_SRC_RANGE:
                            continue
                        connection_name = read(
                            lambda: table.QueryTable.WorkbookConnection.Name,
                            'query table connection name').casefold()
                    except ExcelOperationTimeout:
                        raise
                    except Exception:
                        # Without a connection link this table could belong to any query.
                        incomplete_table_link=True
                        continue
                    try:
                        address = read(lambda: table.Range.Address, 'query table address')
                        sheet_name = read(lambda: sheet.Name, 'worksheet name')
                        sheet_loads.setdefault(connection_name, []).append(f'{sheet_name}!{address}')
                    except ExcelOperationTimeout:
                        raise
                    except Exception:
                        incomplete_connections.add(connection_name)
        except ExcelOperationTimeout:
            raise
        except Exception:
            return {name: 'unknown' for name in query_names}

        states = {}
        for name in query_names:
            if incomplete_table_link:
                states[name]='unknown'
                continue
            # Excel's standard Power Query connection name is Query - <query>.
            key=f'query - {name}'.casefold()
            matches=connections_by_name.get(key,[])
            if len(matches)>1:
                states[name]='unknown'
                continue
            if not matches:
                states[name] = 'definition_only'
                continue
            if key in incomplete_connections:
                states[name]='unknown'
                continue
            connection=matches[0]
            locations = sheet_loads.get(key, [])
            try:
                in_model = bool(read(lambda: connection.InModel, 'connection Data Model state'))
            except ExcelOperationTimeout:
                raise
            except Exception:
                states[name] = 'unknown'
                continue
            modes = []
            if locations:
                modes.append('worksheet: ' + ', '.join(locations))
            if in_model:
                modes.append('data_model')
            states[name] = '; '.join(modes) if modes else 'connection_only'
        return states

    def inspect(self, workbook_path):
        path=validate_workbook(workbook_path)
        excel=workbook=None; initialized=False; previous_security=None
        try:
            if self._default:
                import pythoncom
                pythoncom.CoInitialize(); initialized=True
            excel=self._factory()
            def call(callback, operation):
                return call_with_timeout(excel, callback, OPEN_TIMEOUT_SECONDS,
                    enabled=self._default, operation=f'workbook inspection {operation}')
            call(lambda: setattr(excel, 'Visible', False), 'hide Excel window')
            call(lambda: setattr(excel, 'DisplayAlerts', False), 'suppress Excel alerts')
            try:
                previous_security=call(lambda: excel.AutomationSecurity, 'read macro security setting')
                call(lambda: setattr(excel, 'AutomationSecurity', 3), 'force-disable workbook macros')
            except ExcelOperationTimeout:
                raise
            except Exception as exc:
                raise RuntimeError('Excel macro security could not be set to force-disable; workbook inspection was not started.') from exc
            workbook=call(
                lambda: self._workbook_operations.open_workbook(excel, path, UpdateLinks=0,ReadOnly=True,
                    Password='',WriteResPassword='',IgnoreReadOnlyRecommended=True,AddToMru=False),
                'read-only workbook inspection open')
            if not bool(call(lambda: getattr(workbook,'ReadOnly',True), 'verify workbook read-only state')):
                raise RuntimeError('Excel did not open the workbook read-only; inspection stopped.')
            queries=call(lambda: workbook.Queries, 'workbook query collection')
            count=int(call(lambda: queries.Count, 'workbook query collection count'))
            query_objects = tuple(call(lambda index=index: queries.Item(index),
                f'workbook query item {index}') for index in range(1, count + 1))
            query_names = tuple(str(call(lambda query=query: query.Name,
                'workbook query name')) for query in query_objects)
            load_states = self._load_states(workbook, query_names, call=call)
            inventory=[]
            for query in query_objects:
                name=str(call(lambda query=query: query.Name, 'workbook query name'))
                try: formula=str(call(lambda query=query: query.Formula, 'workbook query formula'))
                except ExcelOperationTimeout: raise
                except Exception: formula=None
                try: description=str(call(lambda query=query: query.Description, 'workbook query description'))
                except ExcelOperationTimeout: raise
                except Exception: description=None
                inventory.append(QueryInfo(name,formula,description,load_states.get(name, 'unknown')))
            return tuple(inventory)
        except ExcelOperationTimeout:
            raise
        except Exception as exc:
            raise RuntimeError(f'Workbook inspection failed: {type(exc).__name__}: {safe_error(exc)}') from exc
        finally:
            if workbook is not None:
                try:
                    if excel is not None:
                        call(lambda: self._workbook_operations.close_workbook(workbook, False), 'close inspected workbook')
                    else:
                        self._workbook_operations.close_workbook(workbook, False)
                except Exception: pass
                workbook = None
            # Release query/connection proxies before quitting the owned Excel instance.
            queries = query = connections = query_objects = load_states = None
            import gc
            gc.collect()
            if excel is not None:
                if previous_security is not None:
                    try: call(lambda: setattr(excel, 'AutomationSecurity', previous_security), 'restore macro security setting')
                    except Exception: pass
                try:
                    if self._default:
                        def release_com_apartment():
                            nonlocal excel, initialized
                            excel = None
                            if initialized:
                                import pythoncom
                                pythoncom.CoUninitialize()
                                initialized = False
                        quit_owned_excel(excel, quit_call=lambda:
                            call(lambda: excel.Quit(), 'quit owned Excel instance'),
                            after_quit=release_com_apartment)
                    else:
                        excel.Quit()
                except Exception as exc:
                    if self._default:
                        if initialized:
                            try:
                                import pythoncom
                                pythoncom.CoUninitialize()
                                initialized = False
                            except Exception:
                                pass
                        raise RuntimeError(f'Excel cleanup failed ({type(exc).__name__}); close the app-owned Excel window if it remains open.') from exc
            if initialized:
                try:
                    import pythoncom
                    pythoncom.CoUninitialize()
                except Exception: pass
