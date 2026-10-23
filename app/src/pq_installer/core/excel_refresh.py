"""Explicit, staged Power Query refresh with bounded waits and per-query results."""
from dataclasses import dataclass, replace
from pathlib import Path
import hashlib
import time

from .backup import create_timestamped_backup
from .validation import validate_workbook_for_write
from .windows_file_users import ensure_no_active_file_users
from .excel_lifetime import (call_with_timeout, call_with_process_timeout,
                             terminate_owned_excel, open_owned_excel_process,
                             OPEN_TIMEOUT_SECONDS, ExcelOperationTimeout,
                             quit_owned_excel)
from .redaction import safe_error
from .excel_adapter import resolve_excel_application_adapter, resolve_excel_workbook_operations
from .excel_transaction import LocalWorkbookTransactionAdapter, cleanup_stage_and_report

def _sha256(path):
    digest = hashlib.sha256()
    with Path(path).open('rb') as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b''):
            digest.update(chunk)
    return digest.hexdigest()


@dataclass(frozen=True)
class RefreshResult:
    query_name: str
    status: str
    message: str
    duration_seconds: float = 0.0


class RefreshOperationError(RuntimeError):
    def __init__(self, message, results=()):
        super().__init__(message)
        self.results = tuple(results)


class ExcelQueryRefresher:
    """Refresh selected/all queries in a staged copy; only this DispatchEx instance is quit."""

    def __init__(self, excel_factory=None, *, excel_adapter=None, workbook_transaction_adapter=None,
                 workbook_operations=None, sleep=time.sleep, monotonic=time.monotonic):
        self._adapter = resolve_excel_application_adapter(adapter=excel_adapter, factory=excel_factory)
        self._workbook_operations = resolve_excel_workbook_operations(workbook_operations)
        self._default = self._adapter.requires_com_initialization
        self._factory = self._adapter.create_application
        self._workbook_transactions = (workbook_transaction_adapter if workbook_transaction_adapter is not None else
            LocalWorkbookTransactionAdapter(check_active=ensure_no_active_file_users))
        self._sleep = sleep
        self._monotonic = monotonic
        self._excel_com_unavailable_after_timeout = False
        self._excel_hwnd = None

    @staticmethod
    def _items(collection):
        count = int(collection.Count)
        return tuple(collection.Item(index) for index in range(1, count + 1))

    @staticmethod
    def _owned_excel_pid(excel):
        """Resolve the process behind this DispatchEx Excel window for timeout cleanup."""
        import win32process
        hwnd = int(excel.Hwnd)
        return int(win32process.GetWindowThreadProcessId(hwnd)[1])

    @staticmethod
    def _terminate_owned_excel(pid):
        """Terminate only the Excel PID obtained from this operation's DispatchEx instance."""
        return terminate_owned_excel(pid)

    def _bounded_process_call(self, excel_pid, callback, timeout_seconds, operation):
        if self._excel_com_unavailable_after_timeout:
            raise RuntimeError('Excel operation timed out; no further COM calls are safe.')
        if not self._default:
            return callback()
        if not excel_pid:
            raise RuntimeError('Excel process ID unavailable; bounded refresh cannot safely target its owned process.')
        owned_process = open_owned_excel_process(excel_pid, hwnd=self._excel_hwnd)
        try:
            return call_with_process_timeout(excel_pid, callback, timeout_seconds,
                operation=operation, terminate=self._terminate_owned_excel,
                owned_process=owned_process)
        except ExcelOperationTimeout:
            self._excel_com_unavailable_after_timeout = True
            raise
        finally:
            owned_process.close()

    def _bounded_excel_call(self, excel, callback, operation, timeout_seconds=None):
        if self._excel_com_unavailable_after_timeout:
            raise RuntimeError('Excel operation timed out; no further COM calls are safe.')
        if timeout_seconds is None:
            timeout_seconds = OPEN_TIMEOUT_SECONDS
        try:
            return call_with_timeout(excel, callback, timeout_seconds, enabled=self._default,
                operation=operation)
        except ExcelOperationTimeout:
            self._excel_com_unavailable_after_timeout = True
            raise

    @staticmethod
    def _collection_item(collection, name):
        try:
            return collection.Item(name)
        except Exception:
            return None

    @staticmethod
    def _plan_data_model_refresh(selected_connection_names, model_tables):
        """Avoid refreshing unselected model sources when Excel's table API is not isolating them."""
        selected = tuple(str(name) for name in selected_connection_names)
        selected_keys = {name.casefold() for name in selected}
        if len(selected_keys) != len(selected):
            raise RuntimeError('Selected Data Model connection mapping is ambiguous; refresh was stopped.')
        mapped = []
        seen = set()
        try:
            for table in model_tables:
                connection_name = str(table.SourceWorkbookConnection.Name)
                key = connection_name.casefold()
                if key in seen:
                    raise RuntimeError('Data Model contains duplicate source-connection mappings; refresh was stopped.')
                seen.add(key)
                mapped.append((connection_name, table))
        except Exception as exc:
            if isinstance(exc, RuntimeError):
                raise
            raise RuntimeError('Could not verify Data Model table sources; refresh was stopped.') from exc
        if not selected_keys:
            return 'none', tuple(mapped)
        if selected_keys != seen:
            raise RuntimeError(
                'A subset of this workbook’s Data Model queries cannot be refreshed safely because Excel may '
                'request other model sources; select all Data Model queries to refresh them together.'
            )
        return ('single' if len(mapped) == 1 else 'all'), tuple(mapped)

    @staticmethod
    def _is_excel_busy_error(exc):
        try:
            hresult = int(exc.args[0]) & 0xFFFFFFFF
        except (AttributeError, IndexError, TypeError, ValueError):
            return False
        return hresult in (0x80010001, 0x8001010A)

    def _prepare_refresh_target(self, workbook, query_name, connection, *, excel_pid):
        """Use a real worksheet QueryTable so Mashup execution has a bounded completion signal."""
        def prepare():
            if bool(connection.InModel):
                expected = f'query - {query_name}'.casefold()
                try:
                    model_tables = self._items(workbook.Model.ModelTables)
                    matches = tuple(table for table in model_tables
                        if str(table.SourceWorkbookConnection.Name).casefold() == expected)
                except Exception as exc:
                    raise RuntimeError(
                        'Could not verify the selected query’s Data Model table; refresh was stopped.'
                    ) from exc
                if len(matches) != 1:
                    raise RuntimeError(
                        'Could not identify exactly one Data Model table for the selected query; '
                        'refresh was stopped.'
                    )
                return ((matches[0], 'data_model', None),)
            expected = f'query - {query_name}'.casefold()
            targets = []
            worksheets = workbook.Worksheets
            for sheet_index in range(1, int(worksheets.Count) + 1):
                sheet = worksheets.Item(sheet_index)
                tables = sheet.ListObjects
                for table_index in range(1, int(tables.Count) + 1):
                    table = tables.Item(table_index)
                    try:
                        if int(table.SourceType) == 1:  # xlSrcRange: ordinary, non-query table
                            continue
                        workbook_connection = table.QueryTable.WorkbookConnection
                        if str(workbook_connection.Name).casefold() == expected:
                            targets.append((table.QueryTable, False, None))
                    except Exception as exc:
                        raise RuntimeError(
                            'Could not verify a worksheet query-load association; refresh was stopped.'
                        ) from exc
            if targets:
                return tuple(targets)

            # A connection-only query has no QueryTable whose state can be observed.
            # Evaluate it in a scratch sheet in the staged workbook and remove only
            # the scratch table, sheet, and connection before saving the stage.
            connection_names = {
                str(workbook.Connections.Item(index).Name).casefold()
                for index in range(1, int(workbook.Connections.Count) + 1)
            }
            ole = connection.OLEDBConnection
            connection_string = str(ole.Connection)
            command_text = str(ole.CommandText)
            if 'Microsoft.Mashup.OleDb.1' not in connection_string:
                raise RuntimeError('The selected query connection is not a supported Power Query Mashup connection.')
            if not command_text.strip():
                raise RuntimeError('The selected Power Query connection has no verifiable command text.')
            original_active_sheet = workbook.ActiveSheet
            scratch_sheet = worksheets.Add(After=worksheets.Item(int(worksheets.Count)))
            try:
                scratch_sheet.Visible = 2  # xlSheetVeryHidden
                query_table = scratch_sheet.QueryTables.Add(
                    connection_string, scratch_sheet.Range('$A$1'), command_text
                )
                generated_connection = query_table.WorkbookConnection
                generated_name = str(generated_connection.Name)
                created_connection = (generated_connection
                                      if generated_name.casefold() not in connection_names else None)
                return ((query_table, True,
                         (scratch_sheet, created_connection, original_active_sheet)),)
            except Exception:
                try:
                    scratch_sheet.Delete()
                except Exception:
                    pass
                raise

        if self._default:
            targets = self._bounded_process_call(excel_pid, prepare, OPEN_TIMEOUT_SECONDS,
                f'prepare {query_name} Mashup QueryTable')
        else:
            try:
                targets = prepare()
            except AttributeError:
                # Dependency-free fake-COM tests may model the workbook connection
                # as the refresh target; real COM workbooks must expose Worksheets.
                targets = ((connection, False, None),)
        return tuple(targets)

    def _refresh_data_model_table(self, query_name, connection, model_table, *, timeout_seconds,
                                  excel_pid, asynchronous):
        """Refresh one loaded model table and verify its selected connection remains readable."""
        started = self._monotonic()
        expected_connection = f'query - {query_name}'.casefold()

        def verify_association():
            actual = str(model_table.SourceWorkbookConnection.Name).casefold()
            if actual != expected_connection:
                raise RuntimeError(
                    f'Data Model table association changed during refresh of {query_name!r}; '
                    'staged workbook was discarded.'
                )
            count = int(model_table.RecordCount)
            if count < 0:
                raise RuntimeError('Excel returned an invalid Data Model row count; staged workbook was discarded.')
            return count

        before = self._bounded_process_call(excel_pid, verify_association, OPEN_TIMEOUT_SECONDS,
            f'verify {query_name} Data Model table before refresh')
        refresh_result = self._bounded_process_call(excel_pid, lambda: model_table.Refresh(),
            timeout_seconds, f'refresh {query_name} Data Model table')
        if refresh_result is False:
            raise RuntimeError(f'Excel cancelled the Data Model refresh for {query_name!r}.')
        after = self._bounded_process_call(excel_pid, verify_association, OPEN_TIMEOUT_SECONDS,
            f'verify {query_name} Data Model table after refresh')
        if after < 0:
            raise RuntimeError(f'Excel returned an invalid Data Model row count for {query_name!r}.')
        return RefreshResult(
            query_name, 'refreshed',
            f'Excel completed the selected Data Model table refresh synchronously; its connection and '
            f'{after} model row(s) verified. Excel does not expose a background option for this API.'
            + ('' if before == after else f' Row count changed from {before}.'),
            max(0.0, self._monotonic() - started),
        )

    def _refresh_all_data_model_tables(self, workbook, mapped_tables, *, timeout_seconds, excel_pid):
        """Use Excel's model-wide API only after preflight proved every model source is selected."""
        started = self._monotonic()
        before = []
        for connection_name, table in mapped_tables:
            def read_before(table=table, expected=connection_name):
                actual = str(table.SourceWorkbookConnection.Name)
                if actual.casefold() != expected.casefold():
                    raise RuntimeError('Data Model source mapping changed before refresh; staged workbook was discarded.')
                count = int(table.RecordCount)
                if count < 0:
                    raise RuntimeError('Excel returned an invalid Data Model row count before refresh.')
                return count
            before.append(self._bounded_process_call(excel_pid, read_before, OPEN_TIMEOUT_SECONDS,
                f'verify {connection_name} Data Model table before refresh'))

        refresh_result = self._bounded_process_call(excel_pid, lambda: workbook.Model.Refresh(),
            timeout_seconds, 'refresh all selected Data Model tables')
        if refresh_result is False:
            raise RuntimeError('Excel cancelled the Data Model refresh; staged workbook was discarded.')

        results = []
        for (connection_name, table), old_count in zip(mapped_tables, before):
            def read_after(table=table, expected=connection_name):
                actual = str(table.SourceWorkbookConnection.Name)
                if actual.casefold() != expected.casefold():
                    raise RuntimeError('Data Model source mapping changed during refresh; staged workbook was discarded.')
                count = int(table.RecordCount)
                if count < 0:
                    raise RuntimeError('Excel returned an invalid Data Model row count after refresh.')
                return count
            count = self._bounded_process_call(excel_pid, read_after, OPEN_TIMEOUT_SECONDS,
                f'verify {connection_name} Data Model table after refresh')
            query_name = connection_name[len('Query - '):] if connection_name.casefold().startswith('query - ') else connection_name
            message = (f'Excel synchronously refreshed all selected Data Model tables; {count} model row(s) '
                f'verified for {query_name!r}. The workbook-wide refresh was used only because every model source '
                'mapped uniquely to a selected Power Query. Excel does not expose a background option for this API.')
            if count != old_count:
                message += f' Row count changed from {old_count}.'
            results.append(RefreshResult(query_name, 'refreshed', message,
                max(0.0, self._monotonic() - started)))
        return tuple(results)

    def _cleanup_temporary_refresh_target(self, query_table, cleanup, *, excel_pid, query_name):
        if cleanup is None or self._excel_com_unavailable_after_timeout:
            return
        sheet, generated_connection, original_active_sheet = cleanup
        def cleanup_owned_objects():
            query_table.Delete()
            sheet.Visible = -1  # xlSheetVisible; Excel rejects deletion of a hidden active sheet.
            original_active_sheet.Activate()
            sheet.Delete()
            if generated_connection is not None:
                generated_connection.Delete()
        self._bounded_process_call(excel_pid, cleanup_owned_objects, OPEN_TIMEOUT_SECONDS,
            f'remove temporary {query_name} refresh table')

    def _refresh_query(self, workbook, query_name, connection, *, timeout_seconds,
                       asynchronous, excel_pid, logger=None):
        targets = self._prepare_refresh_target(workbook, query_name, connection, excel_pid=excel_pid)
        results = []
        for query_table, is_fresh, cleanup in targets:
            failure = None
            try:
                if is_fresh == 'data_model':
                    results.append(self._refresh_data_model_table(
                        query_name, connection, query_table, timeout_seconds=timeout_seconds,
                        excel_pid=excel_pid, asynchronous=asynchronous,
                    ))
                else:
                    results.append(self._refresh_connection(
                        query_name, connection, query_table=query_table,
                        fresh_query_table=is_fresh, timeout_seconds=timeout_seconds,
                        asynchronous=asynchronous, excel_pid=excel_pid, logger=logger,
                    ))
            except Exception as exc:
                failure = exc
            try:
                self._cleanup_temporary_refresh_target(
                    query_table, cleanup, excel_pid=excel_pid, query_name=query_name
                )
            except Exception as cleanup_exc:
                message = f'Temporary refresh table cleanup failed: {safe_error(cleanup_exc)}'
                if failure is not None:
                    message = f'{safe_error(failure)}; {message}'
                raise RuntimeError(message) from (failure or cleanup_exc)
            if failure is not None:
                raise failure
        return replace(results[-1], duration_seconds=sum(item.duration_seconds for item in results))

    def _refresh_connection(self, query_name, connection, *, query_table=None,
                            fresh_query_table=False, timeout_seconds, asynchronous,
                            excel_pid, logger=None):
        started = self._monotonic()
        if self._default and not excel_pid:
            raise RuntimeError('Excel process ID unavailable; bounded refresh cannot safely target its owned process.')
        if query_table is None:
            query_table = connection
        try:
            previous_background = bool(self._bounded_process_call(excel_pid,
                lambda: query_table.BackgroundQuery, OPEN_TIMEOUT_SECONDS,
                f'read {query_name} background refresh setting'))
            self._bounded_process_call(excel_pid,
                lambda: setattr(query_table, 'BackgroundQuery', bool(asynchronous)),
                OPEN_TIMEOUT_SECONDS, f'set {query_name} background refresh setting')
        except ExcelOperationTimeout:
            raise
        except Exception as exc:
            raise RuntimeError(f'Excel cannot set refresh mode for {query_name!r}: {safe_error(exc)}') from exc

        timed_out = False
        try:
            if logger:
                logger(RefreshResult(query_name, 'in_progress', 'Starting the Excel Power Query refresh.'))
            deadline = self._monotonic() + timeout_seconds
            refresh_result = self._bounded_process_call(excel_pid,
                lambda: query_table.Refresh(bool(asynchronous)), timeout_seconds,
                f'refresh for {query_name}')
            if self._default and refresh_result is not True:
                raise RuntimeError('Excel did not confirm that the Power Query refresh completed or started.')
            if refresh_result is False:
                raise RuntimeError('Excel cancelled the Power Query refresh.')
            if asynchronous:
                observed_in_progress = False
                while True:
                    remaining = deadline - self._monotonic()
                    if remaining <= 0:
                        try:
                            self._bounded_process_call(excel_pid, lambda: query_table.CancelRefresh(),
                                OPEN_TIMEOUT_SECONDS, f'cancel {query_name} refresh')
                        except ExcelOperationTimeout:
                            timed_out = True
                            raise
                        except Exception as exc:
                            timed_out = True
                            self._excel_com_unavailable_after_timeout = True
                            raise TimeoutError(f'Refresh timed out after {timeout_seconds:g} seconds; '
                                               f'cancellation failed ({type(exc).__name__}); staged workbook will be discarded.') from exc
                        timed_out = True
                        self._excel_com_unavailable_after_timeout = True
                        raise TimeoutError(f'Refresh timed out after {timeout_seconds:g} seconds; cancellation was requested and staged workbook will be discarded.')
                    try:
                        refreshing = self._bounded_process_call(excel_pid,
                            lambda: bool(query_table.Refreshing), min(OPEN_TIMEOUT_SECONDS, remaining),
                            f'poll {query_name} refresh completion')
                    except Exception as exc:
                        if not self._is_excel_busy_error(exc):
                            raise
                        self._sleep(min(0.1, remaining))
                        continue
                    if not refreshing:
                        if self._default and not (observed_in_progress or fresh_query_table):
                            raise RuntimeError(
                                'Excel did not expose a verifiable active refresh state; staged workbook was discarded.'
                            )
                        break
                    observed_in_progress = True
                    self._sleep(min(0.2, remaining))
            if self._default:
                result_range = self._bounded_process_call(excel_pid,
                    lambda: query_table.ResultRange, OPEN_TIMEOUT_SECONDS,
                    f'verify {query_name} refreshed result range')
                if result_range is None:
                    raise RuntimeError('Excel returned no query result range; staged workbook was discarded.')
                row_count = int(self._bounded_process_call(excel_pid,
                    lambda: result_range.Rows.Count, OPEN_TIMEOUT_SECONDS,
                    f'verify {query_name} refreshed result rows'))
                if row_count < 1:
                    raise RuntimeError('Excel returned an invalid empty query result range; staged workbook was discarded.')
        except ExcelOperationTimeout:
            timed_out = True
            raise
        finally:
            if not timed_out:
                try:
                    self._bounded_process_call(excel_pid,
                        lambda: setattr(query_table, 'BackgroundQuery', previous_background),
                        OPEN_TIMEOUT_SECONDS, f'restore {query_name} background refresh setting')
                except ExcelOperationTimeout:
                    timed_out = True
                    raise
                except Exception as exc:
                    raise RuntimeError(f'Excel could not restore the refresh mode for {query_name!r}: {safe_error(exc)}') from exc

        return RefreshResult(query_name, 'refreshed',
                             'Excel query refresh completed and its query result was verified.',
                             max(0.0, self._monotonic() - started))

    def refresh(self, workbook_path, query_names=None, *, timeout_seconds=300,
                asynchronous=True, make_backup=True, logger=None):
        """Refresh query connections on a staged workbook and commit only if all refreshes pass.

        ``asynchronous`` selects Excel's background-query mode. In either mode this
        method waits on its worker thread for completion and enforces the timeout.
        """
        self._excel_com_unavailable_after_timeout = False
        path = validate_workbook_for_write(workbook_path)
        if isinstance(timeout_seconds, bool) or not isinstance(timeout_seconds, (int, float)) or not 1 <= timeout_seconds <= 7200:
            raise ValueError('Refresh timeout must be between 1 and 7200 seconds.')
        if query_names is not None:
            query_names = tuple(str(name) for name in query_names)
            if not query_names:
                raise ValueError('Select at least one workbook query to refresh.')
            if len({name.casefold() for name in query_names}) != len(query_names):
                raise ValueError('The selected query list contains duplicates.')

        stage = None
        original_hash = _sha256(path)
        excel = workbook = None
        initialized = False
        previous_security = None
        excel_pid = None
        results = []
        try:
            stage = self._workbook_transactions.prepare_stage(path, original_hash, operation='refresh')
            if self._default:
                import pythoncom
                pythoncom.CoInitialize()
                initialized = True
            excel = self._factory()
            if self._default:
                try:
                    self._excel_hwnd = int(excel.Hwnd)
                    excel_pid = self._owned_excel_pid(excel)
                except Exception:
                    self._excel_hwnd = None
                    excel_pid = None
            self._bounded_excel_call(excel, lambda: setattr(excel, 'Visible', False),
                'hide Excel window for refresh')
            self._bounded_excel_call(excel, lambda: setattr(excel, 'DisplayAlerts', False),
                'suppress Excel alerts for refresh')
            try:
                previous_security = self._bounded_excel_call(excel,
                    lambda: excel.AutomationSecurity, 'read refresh macro security setting')
                self._bounded_excel_call(excel,
                    lambda: setattr(excel, 'AutomationSecurity', 3),
                    'force-disable macros for refresh')
            except ExcelOperationTimeout:
                raise
            except Exception as exc:
                raise RuntimeError('Excel macro security could not be set to force-disable; refresh was not started.') from exc
            workbook = self._bounded_excel_call(excel,
                lambda: self._workbook_operations.open_workbook(excel, stage, UpdateLinks=0, ReadOnly=False,
                    Password='', WriteResPassword='', IgnoreReadOnlyRecommended=True, AddToMru=False),
                'staged workbook refresh open')
            if bool(self._bounded_excel_call(excel,
                    lambda: getattr(workbook, 'ReadOnly', False), 'verify refresh workbook write access')):
                raise RuntimeError('Workbook opened read-only; close it or choose a writable copy and retry.')
            if bool(self._bounded_excel_call(excel,
                    lambda: getattr(workbook, 'ProtectStructure', False), 'check refresh workbook structure protection')):
                raise RuntimeError('Workbook structure is protected; refresh was stopped.')

            def read_inventory():
                query_objects = self._items(workbook.Queries)
                query_by_name = {str(query.Name).casefold(): str(query.Name) for query in query_objects}
                if query_names is None:
                    selected = [str(query.Name) for query in query_objects]
                else:
                    missing = [name for name in query_names if name.casefold() not in query_by_name]
                    if missing:
                        raise ValueError(f'Query not found in workbook inventory: {", ".join(missing)}')
                    selected = [query_by_name[name.casefold()] for name in query_names]

                query_formulas = {}
                for query in query_objects:
                    name = str(query.Name)
                    if name in selected:
                        try:
                            query_formulas[name.casefold()] = str(query.Formula)
                        except Exception as exc:
                            raise RuntimeError(f'Cannot safely refresh {name!r}: its current M formula cannot be verified.') from exc

                connections = self._items(workbook.Connections)
                connection_by_name = {str(connection.Name).casefold(): connection for connection in connections}
                refreshable = []
                selected_model_connections = []
                skipped = []
                for name in selected:
                    connection = connection_by_name.get(f'query - {name}'.casefold())
                    if connection is None:
                        result = RefreshResult(name, 'skipped', 'No refreshable Power Query connection is present; definition-only queries are unchanged.')
                        skipped.append(result)
                        if logger:
                            logger(result)
                    else:
                        refreshable.append((name, connection))
                        try:
                            if bool(connection.InModel):
                                selected_model_connections.append(str(connection.Name))
                        except Exception as exc:
                            raise RuntimeError(
                                f'Could not verify whether query {name!r} is loaded to the Data Model; refresh was stopped.'
                            ) from exc
                if selected_model_connections:
                    model_tables = self._items(workbook.Model.ModelTables)
                    model_strategy, mapped_model_tables = self._plan_data_model_refresh(
                        selected_model_connections, model_tables)
                else:
                    model_strategy, mapped_model_tables = 'none', ()
                return (query_formulas, selected, refreshable, skipped,
                    model_strategy, mapped_model_tables)

            (query_formulas, selected, refreshable, skipped,
                model_strategy, mapped_model_tables) = self._bounded_excel_call(
                excel, read_inventory, 'staged workbook refresh query inventory')
            results.extend(skipped)

            if not refreshable:
                return tuple(results)
            if make_backup:
                backup = create_timestamped_backup(path)
                if logger:
                    logger(f'Backup created: {backup}')

            model_query_keys = {name.casefold() for name, connection in refreshable
                if bool(self._bounded_excel_call(excel, lambda connection=connection: bool(connection.InModel),
                    f'verify {connection.Name} Data Model state'))}
            data_model_group_refreshed = False
            for name, connection in refreshable:
                if logger:
                    logger(RefreshResult(name, 'queued', 'Waiting for the Excel refresh worker.'))
                refresh_started = self._monotonic()
                try:
                    if (model_strategy == 'all' and name.casefold() in model_query_keys):
                        if not data_model_group_refreshed:
                            group_results = self._refresh_all_data_model_tables(
                                workbook, mapped_model_tables, timeout_seconds=timeout_seconds,
                                excel_pid=excel_pid)
                            results.extend(group_results)
                            data_model_group_refreshed = True
                        continue
                    result = self._refresh_query(workbook, name, connection,
                        timeout_seconds=timeout_seconds, asynchronous=asynchronous,
                        excel_pid=excel_pid, logger=logger)
                except Exception as exc:
                    elapsed = max(0.0, self._monotonic() - refresh_started)
                    result = RefreshResult(name, 'failed', f'{type(exc).__name__}: {safe_error(exc)}', elapsed)
                    results = [replace(item, status='rolled_back',
                                       message='Refresh succeeded in the staged workbook but was discarded because the operation failed.')
                               if item.status == 'refreshed' else item for item in results]
                    results.append(result)
                    result_by_name = {item.query_name.casefold(): item for item in results}
                    results = [result_by_name[name.casefold()] for name in selected if name.casefold() in result_by_name]
                    raise RefreshOperationError(
                        f'Refresh failed for {name!r}; staged changes were discarded. {result.message}', results
                    ) from exc
                results.append(result)

            result_by_name = {item.query_name.casefold(): item for item in results}
            results = [result_by_name[name.casefold()] for name in selected if name.casefold() in result_by_name]

            self._bounded_excel_call(excel, lambda: self._workbook_operations.save_workbook(workbook),
                'staged workbook refresh save')
            self._bounded_excel_call(excel, lambda: self._workbook_operations.close_workbook(workbook, False),
                'close staged workbook after refresh')
            workbook = None
            verification = self._bounded_excel_call(excel,
                lambda: self._workbook_operations.open_workbook(excel, stage, UpdateLinks=0, ReadOnly=True,
                    Password='', WriteResPassword='', IgnoreReadOnlyRecommended=True, AddToMru=False),
                'staged workbook refresh verification open')
            try:
                def verify_refresh():
                    if not bool(getattr(verification, 'ReadOnly', False)):
                        raise RuntimeError('Excel did not reopen the staged refresh read-only; original workbook was retained.')
                    for name, _connection in refreshable:
                        verified_query = verification.Queries.Item(name)
                        if str(verified_query.Formula) != query_formulas[name.casefold()]:
                            raise RuntimeError(f'Query formula changed unexpectedly during refresh: {name!r}.')
                        verification.Connections.Item(f'Query - {name}')
                self._bounded_excel_call(excel, verify_refresh,
                    'verify staged workbook refresh results')
            finally:
                try:
                    self._bounded_excel_call(excel,
                        lambda: self._workbook_operations.close_workbook(verification, False),
                        'close staged refresh verification workbook')
                except Exception:
                    pass
            self._workbook_transactions.commit_stage(stage, path, original_hash, operation='refresh')
            stage = None
            results = [replace(item, message=(
                f'{item.message.rstrip(". ")} Staged workbook verification passed and the workbook was committed.'
            ))
                       if item.status == 'refreshed' else item for item in results]
            if logger:
                for item in results:
                    if item.status == 'refreshed':
                        logger(item)
            return tuple(results)
        except RefreshOperationError:
            raise
        except Exception as exc:
            results = [replace(item, status='rolled_back',
                               message='Refresh succeeded in the staged workbook but was discarded because the workbook could not be committed.')
                       if item.status == 'refreshed' else item for item in results]
            raise RefreshOperationError(f'{type(exc).__name__}: {safe_error(exc)}; original workbook was retained.', results) from exc
        finally:
            if workbook is not None:
                try:
                    if excel is not None:
                        self._bounded_excel_call(excel,
                            lambda: self._workbook_operations.close_workbook(workbook, False),
                            'close staged workbook after refresh failure')
                    else:
                        self._workbook_operations.close_workbook(workbook, False)
                except Exception:
                    pass
            if stage is not None:
                cleanup_stage_and_report(self._workbook_transactions, stage, logger=logger)
            if excel is not None:
                if previous_security is not None:
                    try:
                        self._bounded_excel_call(excel,
                            lambda: setattr(excel, 'AutomationSecurity', previous_security),
                            'restore refresh macro security setting')
                    except Exception:
                        pass
                try:
                    if self._default:
                        def release_com_apartment():
                            nonlocal excel, initialized
                            excel = None
                            if initialized:
                                import pythoncom
                                pythoncom.CoUninitialize()
                                initialized = False
                        terminated = quit_owned_excel(excel, quit_call=lambda:
                            self._bounded_excel_call(excel, lambda: excel.Quit(),
                                'quit owned Excel instance after refresh'),
                            after_quit=release_com_apartment)
                        if logger:
                            logger(f'Closed captured app-owned Excel process; forced termination={terminated}.')
                    else:
                        excel.Quit()
                except Exception as cleanup_exc:
                    if logger:
                        logger(f'Excel cleanup could not be confirmed ({type(cleanup_exc).__name__}); close the app-owned Excel window if it remains open.')
            if initialized:
                try:
                    import pythoncom
                    pythoncom.CoUninitialize()
                except Exception:
                    pass
            self._excel_hwnd = None
