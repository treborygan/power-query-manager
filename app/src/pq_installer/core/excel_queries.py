from pathlib import Path
import hashlib
from .models import ExistingQueryMode, InstallResult, QueryStatus
from .excel_lifetime import (call_with_timeout, OPEN_TIMEOUT_SECONDS,
                             ExcelOperationTimeout, owned_excel_pid, quit_owned_excel)
from .query_loader import read_m_code
from .validation import validate_query_sources, validate_workbook_for_write, ordered_sources
from .windows_file_users import ensure_no_active_file_users
from .redaction import safe_error
from .excel_adapter import resolve_excel_application_adapter, resolve_excel_workbook_operations
from .excel_transaction import LocalWorkbookTransactionAdapter, cleanup_stage_and_report
from .mashup_privacy import read_workbook_privacy

_XL_SRC_RANGE = 1

def _sha256(path):
    digest = hashlib.sha256()
    with Path(path).open('rb') as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b''):
            digest.update(chunk)
    return digest.hexdigest()


def _formula_matches_excel_roundtrip(expected, actual):
    """Match exact M text, allowing Excel to remove terminal line endings on save."""
    expected = str(expected)
    actual = str(actual)
    return expected == actual or expected.rstrip('\r\n') == actual.rstrip('\r\n')


def _com_exception_scode(exc):
    """Return the provider EXCEPINFO.scode only for a DISP_E_EXCEPTION."""
    try:
        outer_hresult = int(exc.args[0]) & 0xFFFFFFFF
        exception_info = exc.args[2]
        if outer_hresult != 0x80020009:
            return None
        return int(exception_info[5]) & 0xFFFFFFFF
    except (AttributeError, IndexError, TypeError, ValueError):
        return None


class InstallOperationError(RuntimeError):
    """Failed staged import with explicit per-query outcomes after rollback."""
    def __init__(self,message,results=()):
        super().__init__(message)
        self.results=tuple(results)

class ExcelQueryInstaller:
    def __init__(self, excel_factory=None, *, excel_adapter=None, workbook_transaction_adapter=None,
                 workbook_operations=None):
        self._adapter = resolve_excel_application_adapter(adapter=excel_adapter, factory=excel_factory)
        self._workbook_operations = resolve_excel_workbook_operations(workbook_operations)
        self._uses_default_factory = self._adapter.requires_com_initialization
        self._excel_factory = self._adapter.create_application
        self._workbook_transactions = (workbook_transaction_adapter if workbook_transaction_adapter is not None else
            LocalWorkbookTransactionAdapter(check_active=ensure_no_active_file_users))
        self._excel_com_unavailable_after_timeout = False

    def _bounded_excel_call(self, excel, callback, operation, timeout_seconds=None):
        """Bound importer COM work and stop touching an instance once its process exited."""
        if self._excel_com_unavailable_after_timeout:
            raise RuntimeError('Excel operation timed out; no further COM calls are safe.')
        if timeout_seconds is None:
            timeout_seconds = OPEN_TIMEOUT_SECONDS
        try:
            return call_with_timeout(excel, callback, timeout_seconds,
                enabled=self._uses_default_factory, operation=operation)
        except ExcelOperationTimeout:
            self._excel_com_unavailable_after_timeout = True
            raise

    @staticmethod
    def _get_existing_query(workbook, name):
        try:
            return workbook.Queries.Item(name)
        except Exception as exc:
            # On this supported Excel build, Mashup reports a missing query as
            # DISP_E_EXCEPTION with TYPE_E_ELEMENTNOTFOUND in EXCEPINFO.scode.
            # Treat only that observed provider signature as absence; any other
            # automation failure must stop before a create/replacement decision.
            if _com_exception_scode(exc) == 0x8002802B:
                return None
            raise RuntimeError(
                f'Excel query {name!r} lookup could not be checked; refusing to assume it is absent: {safe_error(exc)}'
            ) from exc

    @staticmethod
    def _get_existing_connection(workbook, name):
        try:
            return workbook.Connections.Item(name)
        except Exception as exc:
            # Workbook.Connections reports a missing item as DISP_E_EXCEPTION
            # with DISP_E_BADINDEX in EXCEPINFO.scode on the observed build.
            if _com_exception_scode(exc) == 0x8002000B:
                return None
            raise RuntimeError(
                f'Excel connection {name!r} lookup could not be checked; refusing to assume it is absent: {safe_error(exc)}'
            ) from exc

    @staticmethod
    def _query_load_mapping(workbook, query_name):
        try:
            signature = ExcelQueryInstaller._query_load_signature(workbook, query_name)
            if signature == ('definition_only',):
                return 'definition_only'
            _name, in_model, destinations = signature
            states = []
            if in_model:
                states.append('data_model')
            if destinations:
                states.append('worksheet')
            return '+'.join(states) if states else 'connection_only'
        except Exception:
            return 'unknown'

    @staticmethod
    def _query_worksheet_tables(workbook, query_name):
        expected = f'query - {query_name}'.casefold()
        matches = []
        for sheet_index in range(1, int(workbook.Worksheets.Count) + 1):
            sheet = workbook.Worksheets.Item(sheet_index)
            for table_index in range(1, int(sheet.ListObjects.Count) + 1):
                table = sheet.ListObjects.Item(table_index)
                try:
                    # An ordinary xlSrcRange ListObject has no QueryTable; its
                    # source type explicitly proves it is unrelated to PQ.
                    if int(table.SourceType) == _XL_SRC_RANGE:
                        continue
                    connection_name = str(table.QueryTable.WorkbookConnection.Name).casefold()
                except Exception as exc:
                    raise RuntimeError('Could not verify an existing worksheet load association; refusing to change it.') from exc
                if connection_name == expected:
                    matches.append((sheet, table))
        return matches

    @staticmethod
    def _query_load_signature(workbook, query_name):
        """Capture load associations or fail closed when any relevant mapping is ambiguous."""
        try:
            connections = workbook.Connections
            expected_name = f'query - {query_name}'.casefold()
            matching_connections = []
            for index in range(1, int(connections.Count) + 1):
                candidate = connections.Item(index)
                if str(candidate.Name).casefold() == expected_name:
                    matching_connections.append(candidate)
            if len(matching_connections) > 1:
                raise RuntimeError('Multiple connections use the same Power Query identity.')
            if not matching_connections:
                return ('definition_only',)
            connection = matching_connections[0]
            in_model = bool(connection.InModel)
            worksheets = workbook.Worksheets
            destinations = []
            for sheet_index in range(1, int(worksheets.Count) + 1):
                sheet = worksheets.Item(sheet_index)
                list_objects = sheet.ListObjects
                for table_index in range(1, int(list_objects.Count) + 1):
                    table = list_objects.Item(table_index)
                    try:
                        # Plain Excel tables report xlSrcRange and have no QueryTable.
                        # Their non-query status is explicit, so they need not make a
                        # Power Query load signature unknowable.
                        if int(table.SourceType) == _XL_SRC_RANGE:
                            continue
                        table_connection = str(table.QueryTable.WorkbookConnection.Name)
                    except Exception as exc:
                        raise RuntimeError('A worksheet table connection could not be identified.') from exc
                    if table_connection.casefold() == str(connection.Name).casefold():
                        destinations.append((str(sheet.Name).casefold(),
                                             str(table.Range.Cells(1, 1).Address).replace('$', '').upper(),
                                             str(table.Name).casefold()))
            return (str(connection.Name).casefold(), in_model, tuple(sorted(destinations)))
        except Exception as exc:
            raise RuntimeError(f'Cannot verify existing load associations for {query_name!r}; import stopped: {safe_error(exc)}') from exc

    def _bounded_load(self, callback, excel, timeout_seconds):
        return self._bounded_excel_call(excel, callback, 'Excel query evaluation', timeout_seconds)

    @staticmethod
    def _find_worksheet_by_name(workbook, name):
        """Enumerate worksheet names; never infer absence from a failed named lookup."""
        try:
            worksheets = workbook.Worksheets
            count = int(worksheets.Count)
            for index in range(1, count + 1):
                sheet = worksheets.Item(index)
                if str(sheet.Name).casefold() == str(name).casefold():
                    return sheet
            return None
        except Exception as exc:
            raise RuntimeError(
                'Excel worksheet inventory could not be checked; refusing to assume a destination is absent: '
                f'{safe_error(exc)}') from exc

    @staticmethod
    def _add_worksheet_load(workbook, source):
        """Create a Power Query-backed Excel table at a reviewed, empty destination."""
        sheet_name = source.sheet_name
        if source.load_mode.value == 'new_sheet':
            if not sheet_name:
                sheet_name = ('PQ_' + ''.join(c for c in source.name if c.isalnum() or c == '_'))[:31] or 'PowerQuery'
                base, suffix = sheet_name, 2
                while True:
                    if ExcelQueryInstaller._find_worksheet_by_name(workbook, sheet_name) is None:
                        break
                    sheet_name = f'{base[:27]}_{suffix}'
                    suffix += 1
            else:
                if ExcelQueryInstaller._find_worksheet_by_name(workbook, sheet_name) is not None:
                    raise RuntimeError(f'Worksheet {sheet_name!r} already exists; choose a new sheet name.')
            sheet = workbook.Worksheets.Add(After=workbook.Worksheets.Item(workbook.Worksheets.Count))
            sheet.Name = sheet_name
            cell = 'A1'
        else:
            sheet = ExcelQueryInstaller._find_worksheet_by_name(workbook, source.sheet_name)
            if sheet is None:
                raise RuntimeError(f'Worksheet {source.sheet_name!r} was not found; choose an existing destination sheet.')
            cell = source.cell
            # Unknown result size makes overwrite risk unacceptable. Accept only an empty sheet.
            if (int(sheet.ListObjects.Count) or int(sheet.QueryTables.Count)
                    or int(sheet.PivotTables().Count) or int(sheet.Shapes.Count)):
                raise RuntimeError(f'Worksheet {source.sheet_name!r} contains a table, query, or pivot; choose an empty destination sheet.')
            try:
                used = sheet.UsedRange
                if int(used.Count) > 1 or used.Value not in (None, ''):
                    raise RuntimeError(f'Worksheet {source.sheet_name!r} is not empty; choose an empty destination sheet.')
            except RuntimeError:
                raise
            except Exception:
                raise RuntimeError(f'Could not prove worksheet {source.sheet_name!r} is empty; choose another destination.')
        connection_string = (f'OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;'
                             f'Location={source.name};Extended Properties=""')
        table = sheet.ListObjects.Add(3, connection_string, True, 1, sheet.Range(cell))
        table.Name = source.table_name or ('PQ_' + ''.join(c for c in source.name if c.isalnum() or c == '_'))[:250]
        qt = table.QueryTable
        qt.CommandType = 2
        qt.CommandText = f'SELECT * FROM [{source.name}]'
        qt.BackgroundQuery = False
        qt.Refresh(False)
        try: qt.WorkbookConnection.Name = f'Query - {source.name}'
        except Exception as exc: raise RuntimeError(f'Excel did not create the expected Power Query connection for {source.name!r}.') from exc
        return {'sheet': str(sheet.Name), 'cell': str(table.Range.Cells(1, 1).Address),
                'table_name': str(table.Name), 'address': str(table.Range.Address)}

    @staticmethod
    def _add_model_load(workbook, source):
        connection_name = f'Query - {source.name}'
        connection_string = (f'OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;'
                             f'Location={source.name};Extended Properties=""')
        connection = workbook.Connections.Add2(connection_name, f'Power Query - {source.name}',
            connection_string, f'SELECT * FROM [{source.name}]', 2, True, False)
        connection.Refresh()
        if not bool(connection.InModel):
            raise RuntimeError(f'Excel did not mark query {source.name!r} as loaded to the Data Model.')

    def install(self, workbook_path, sources, mode: ExistingQueryMode, logger=None, capabilities=None,
                load_timeout_seconds=300, expected_workbook_sha256=None):
        self._excel_com_unavailable_after_timeout = False
        workbook_path = Path(workbook_path)
        validate_workbook_for_write(workbook_path)
        validate_query_sources(sources)
        sequence = ordered_sources(sources)
        requested_modes = {getattr(source.load_mode, 'value', source.load_mode) for source in sequence}
        for requested_mode in requested_modes - {'definition_only', 'preserve_existing'}:
            if capabilities is None or requested_mode not in capabilities.supported_load_modes:
                raise RuntimeError(f'{requested_mode.replace("_", " ").title()} mode requires a successful Excel capability probe for this Office build.')
        if isinstance(load_timeout_seconds, bool) or not isinstance(load_timeout_seconds, (int, float)) or not 1 <= load_timeout_seconds <= 7200:
            raise ValueError('Load timeout must be between 1 and 7200 seconds.')
        # Parse every file before Excel is opened so bad inputs cannot partially change a workbook.
        formulas = []
        for source in sequence:
            raw = Path(source.path).read_bytes()
            digest = hashlib.sha256(raw).hexdigest()
            if source.source_sha256 and digest != source.source_sha256:
                raise RuntimeError(f'Source changed after review; reload the query before applying: {source.name}')
            try:
                formula = raw.decode('utf-8-sig')
            except UnicodeDecodeError as exc:
                raise ValueError(f'Cannot decode {source.path} as UTF-8. Save the Power Query source as UTF-8 and retry.') from exc
            if not formula.strip():
                raise ValueError(f'Power Query source file is empty: {source.name}')
            formulas.append((source, formula, digest))
        reviewed_workbook_hash = _sha256(workbook_path) if workbook_path.is_file() else None
        if expected_workbook_sha256 and reviewed_workbook_hash != expected_workbook_sha256:
            raise RuntimeError('Workbook changed after import review; reload the workbook and review again.')
        original_privacy = read_workbook_privacy(workbook_path) if workbook_path.is_file() else None
        excel = workbook = None
        previous_security = None
        staged_path = None
        original_hash = None
        com_initialized = False
        journal = []
        results = []
        plan = []
        current_source = None
        current_phase = 'opening workbook'
        load_expectations = {}
        original_load_signatures = {}
        try:
            if self._uses_default_factory:
                import pythoncom
                pythoncom.CoInitialize()
                com_initialized = True
            excel = self._excel_factory()
            if capabilities is not None:
                version, build = self._bounded_excel_call(excel, lambda: (
                    str(getattr(excel, 'Version', '')),
                    str(getattr(getattr(excel, 'Build', ''), 'value', getattr(excel, 'Build', '')))),
                    'read Excel version for import')
                if version and version != capabilities.office_version:
                    raise RuntimeError('Excel version changed after the capability probe; probe again before applying.')
                if build and capabilities.build and build != capabilities.build:
                    raise RuntimeError('Excel build changed after the capability probe; probe again before applying.')
            try:
                previous_security = self._bounded_excel_call(excel, lambda: (
                    setattr(excel, 'Visible', False),
                    setattr(excel, 'DisplayAlerts', False),
                    getattr(excel, 'AutomationSecurity'))[2],
                    'configure Excel and read import macro security')
                self._bounded_excel_call(excel,
                    lambda: setattr(excel, 'AutomationSecurity', 3),
                    'force-disable macros for import')
            except ExcelOperationTimeout:
                raise
            except Exception as exc:
                raise RuntimeError('Excel macro security could not be set to force-disable; workbook was not opened.') from exc
            open_path = workbook_path
            if workbook_path.is_file():
                original_hash = reviewed_workbook_hash
                staged_path = self._workbook_transactions.prepare_stage(
                    workbook_path, original_hash, operation='import')
                open_path = staged_path
            workbook = self._bounded_excel_call(excel,
                lambda: self._workbook_operations.open_workbook(excel, open_path, UpdateLinks=0, ReadOnly=False,
                    Password='', WriteResPassword='', IgnoreReadOnlyRecommended=True, AddToMru=False),
                'staged workbook open')
            if logger and self._uses_default_factory:
                try:
                    from win32process import GetWindowThreadProcessId
                    workbook_hwnd = int(workbook.Windows.Item(1).Hwnd)
                    logger(f'Excel ownership check: application PID {owned_excel_pid(excel)}, workbook-window PID {int(GetWindowThreadProcessId(workbook_hwnd)[1])}.')
                except Exception as ownership_error:
                    logger(f'Excel ownership diagnostic unavailable ({type(ownership_error).__name__}).')
            readonly, protected = self._bounded_excel_call(excel, lambda: (
                bool(getattr(workbook, 'ReadOnly', False)),
                bool(getattr(workbook, 'ProtectStructure', False))),
                'check staged workbook access and protection')
            if readonly:
                raise RuntimeError('Workbook opened read-only; close it or choose a writable copy and retry.')
            if protected:
                raise RuntimeError('Workbook structure is protected; remove protection before applying query changes.')

            def capture_existing_loads():
                for source, _formula, _digest in formulas:
                    if self._get_existing_query(workbook, source.name) is not None:
                        original_load_signatures[source.name.casefold()] = self._query_load_signature(workbook, source.name)
            self._bounded_excel_call(excel, capture_existing_loads, 'inspect existing query load associations')

            # First create/update every definition. No connection or load action may evaluate M
            # until every reviewed definition is present.
            current_phase='installing query definitions'
            for source, formula, digest in formulas:
                current_source=source
                existing = self._bounded_excel_call(excel,
                    lambda: self._get_existing_query(workbook, source.name),
                    f'find existing query {source.name}')
                replacing_existing = existing is not None
                if existing is not None and mode == ExistingQueryMode.SKIP:
                    result = InstallResult(source.name, QueryStatus.SKIPPED, 'Existing query kept unchanged.')
                    plan.append((source, formula, False, None, result))
                    continue
                if source.load_mode.value == 'preserve_existing' and not replacing_existing:
                    raise RuntimeError(f'Query {source.name!r} uses preserve-existing mode but does not exist in the workbook.')
                if existing is not None:
                    try:
                        old_formula = self._bounded_excel_call(excel,
                            lambda: existing.Formula, f'read existing formula for {source.name}')
                    except Exception as exc:
                        raise RuntimeError(f'Query {source.name!r} cannot safely replace: existing formula could not be captured.') from exc
                    if old_formula is None:
                        raise RuntimeError(f'Query {source.name!r} cannot safely replace: existing formula is unavailable.')
                    journal.append(('updated', source.name, old_formula))
                    # Record before calling COM: a setter can mutate Excel and then raise,
                    # leaving no journal entry if it is appended after the call.
                    self._bounded_excel_call(excel,
                        lambda: setattr(existing, 'Formula', formula),
                        f'update formula for {source.name}')
                    status = QueryStatus.REPLACED
                else:
                    self._bounded_excel_call(excel,
                        lambda: workbook.Queries.Add(source.name, formula),
                        f'add query {source.name}')
                    journal.append(('created', source.name, None))
                    status = QueryStatus.CREATED
                result = InstallResult(source.name, status, f'{status.value} query definition; load mode: definition only (no connection or worksheet load).')
                plan.append((source, formula, replacing_existing, status, result))

            if plan and all(status is None for _source, _formula, _replacing, status, _result in plan):
                results = [result for _source, _formula, _replacing, _status, result in plan]
                current_source = None
                if logger:
                    for result in results:
                        logger(result)
                return results

            # Then apply explicit load choices after all dependency definitions exist.
            current_phase='applying requested load mode'
            for source, formula, replacing_existing, status, result in plan:
                current_source=source
                if status is None:
                    results.append(result)
                    continue
                if source.load_mode.value == 'connection_only':
                    connection_name=f'Query - {source.name}'
                    connection=self._bounded_excel_call(excel,
                        lambda: self._get_existing_connection(workbook,connection_name),
                        f'find connection for {source.name}')
                    if connection is not None:
                        mapping = self._bounded_excel_call(excel,
                            lambda: self._query_load_mapping(workbook, source.name),
                            f'inspect load mapping for {source.name}')
                        if mapping != 'connection_only':
                            raise RuntimeError(f'Query {source.name!r} already has load mapping {mapping!r}; connection-only cannot replace it implicitly.')
                    if connection is None:
                        connection_string=(f'OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;'
                                           f'Location={source.name};Extended Properties=""')
                        connection=self._bounded_excel_call(excel,
                            lambda: workbook.Connections.Add2(connection_name,f'Power Query - {source.name}',
                                connection_string,f'SELECT * FROM [{source.name}]',2,False,False),
                            f'create connection for {source.name}')
                        journal.append(('connection_created',connection_name,None))
                        message = ('Added an explicit workbook connection; any previous load associations were left untouched. No refresh was requested.'
                                   if replacing_existing else
                                   'Created an explicit workbook connection; no worksheet or Data Model load and no refresh were requested.')
                        result=InstallResult(source.name,status,message)
                    else:
                        result=InstallResult(source.name,status,'Existing Power Query connection retained unchanged; worksheet/model load associations were not modified and no refresh was requested.')
                elif source.load_mode.value in ('new_sheet', 'existing_sheet', 'data_model'):
                    if source.load_mode.value in ('new_sheet', 'existing_sheet'):
                        mapping = (self._bounded_excel_call(excel,
                            lambda: self._query_load_mapping(workbook, source.name),
                            f'inspect load mapping for {source.name}')
                            if replacing_existing else 'definition_only')
                        if mapping == 'worksheet':
                            matches = self._bounded_excel_call(excel,
                                lambda: self._query_worksheet_tables(workbook, source.name),
                                f'inspect worksheet load for {source.name}')
                            if len(matches) != 1:
                                raise RuntimeError(f'Query {source.name!r} has an ambiguous worksheet mapping; inspect it before reimporting.')
                            sheet, table = matches[0]
                            def refresh_existing_table():
                                cell = str(table.Range.Cells(1, 1).Address)
                                sheet_name = str(sheet.Name)
                                table_name = str(table.Name)
                                if source.load_mode.value == 'existing_sheet':
                                    desired_cell = source.cell.replace('$', '').upper()
                                    observed_cell = cell.replace('$', '').upper()
                                    if sheet_name.casefold() != source.sheet_name.casefold() or observed_cell != desired_cell:
                                        raise RuntimeError(f'Query {source.name!r} is already loaded to {sheet_name}!{cell}; refusing to move its existing output implicitly.')
                                elif source.sheet_name and sheet_name.casefold() != source.sheet_name.casefold():
                                    raise RuntimeError(f'Query {source.name!r} is already loaded to {sheet_name}; requested new worksheet {source.sheet_name!r} would duplicate the load.')
                                if source.table_name and table_name.casefold() != source.table_name.casefold():
                                    raise RuntimeError(f'Query {source.name!r} already uses table {table_name!r}; refusing to rename it implicitly.')
                                location = {'sheet': sheet_name, 'cell': cell, 'table_name': table_name,
                                            'address': str(table.Range.Address)}
                                table.QueryTable.Refresh(False)
                                return location
                            load_expectations[source.name.casefold()] = self._bounded_load(
                                refresh_existing_table, excel, load_timeout_seconds)
                        elif mapping == 'definition_only':
                            load_expectations[source.name.casefold()] = self._bounded_load(
                                lambda: self._add_worksheet_load(workbook, source), excel, load_timeout_seconds)
                            journal.append(('connection_created', f'Query - {source.name}', None))
                        else:
                            raise RuntimeError(f'Query {source.name!r} already has load mapping {mapping!r}; refusing to alter it implicitly.')
                        location = load_expectations[source.name.casefold()]
                        result = InstallResult(source.name, status, f'Worksheet table preserved/refreshed at {location["sheet"]}!{location["address"]}; query evaluation was requested by this load.')
                    else:
                        mapping = (self._bounded_excel_call(excel,
                            lambda: self._query_load_mapping(workbook, source.name),
                            f'inspect load mapping for {source.name}')
                            if replacing_existing else 'definition_only')
                        if mapping == 'data_model':
                            connection = self._bounded_excel_call(excel,
                                lambda: self._get_existing_connection(workbook, f'Query - {source.name}'),
                                f'find Data Model connection for {source.name}')
                            self._bounded_load(connection.Refresh, excel, load_timeout_seconds)
                        elif mapping == 'definition_only':
                            self._bounded_load(lambda: self._add_model_load(workbook, source), excel, load_timeout_seconds)
                            journal.append(('connection_created', f'Query - {source.name}', None))
                        else:
                            raise RuntimeError(f'Query {source.name!r} already has load mapping {mapping!r}; refusing to alter it implicitly.')
                        load_expectations[source.name.casefold()] = {'mode': 'data_model'}
                        result = InstallResult(source.name, status, 'Data Model load preserved/refreshed or created and verified.')
                elif source.load_mode.value == 'preserve_existing':
                    result = InstallResult(source.name, status, 'Updated the existing query definition; its current load association was preserved unchanged.')
                elif replacing_existing:
                    result=InstallResult(source.name,status,'Updated the existing query definition in place; its current connection, worksheet, and Data Model associations were left unchanged. No refresh was requested.')
                results.append(result)
            current_source=None
            current_phase='saving and verifying staged workbook'
            self._bounded_excel_call(excel, lambda: self._workbook_operations.save_workbook(workbook),
                'staged workbook save')
            if staged_path is not None:
                self._bounded_excel_call(excel, lambda: self._workbook_operations.close_workbook(workbook, False),
                    'close staged workbook before verification')
                workbook = None
                staged_privacy = read_workbook_privacy(staged_path)
                if original_privacy is not None and staged_privacy != original_privacy:
                    raise RuntimeError('Staged workbook verification failed: Power Query privacy settings changed.')
                verification = self._bounded_excel_call(excel,
                    lambda: self._workbook_operations.open_workbook(excel, staged_path, UpdateLinks=0, ReadOnly=True,
                        Password='', WriteResPassword='', IgnoreReadOnlyRecommended=True, AddToMru=False),
                    'staged workbook verification open')
                try:
                    def verify_staged_workbook():
                        if not bool(getattr(verification, 'ReadOnly', True)):
                            raise RuntimeError('Staged workbook verification failed: Excel did not reopen it read-only.')
                        for source, expected_formula, replacing_existing, expected_status, _result in plan:
                            if expected_status is None:
                                continue
                            query = self._get_existing_query(verification, source.name)
                            if query is None or not _formula_matches_excel_roundtrip(
                                    expected_formula, query.Formula):
                                raise RuntimeError(f'Staged workbook verification failed for query {source.name!r}: saved M formula did not match the reviewed source.')
                            if replacing_existing:
                                old_signature = original_load_signatures[source.name.casefold()]
                                requested = source.load_mode.value
                                if not (old_signature == ('definition_only',) and requested in
                                        ('connection_only', 'new_sheet', 'existing_sheet', 'data_model')):
                                    observed_signature = self._query_load_signature(verification, source.name)
                                    if observed_signature != old_signature:
                                        raise RuntimeError(f'Staged workbook verification failed for {source.name!r}: existing load associations changed from {old_signature!r} to {observed_signature!r}.')
                            if source.load_mode.value == 'connection_only':
                                connection = self._get_existing_connection(verification, f'Query - {source.name}')
                                if connection is None or bool(connection.InModel):
                                    raise RuntimeError(f'Staged workbook verification failed for connection-only query {source.name!r}.')
                            elif source.load_mode.value == 'data_model':
                                connection = self._get_existing_connection(verification, f'Query - {source.name}')
                                if connection is None or not bool(connection.InModel):
                                    raise RuntimeError(f'Staged workbook verification failed for Data Model query {source.name!r}.')
                                model_tables = tuple(verification.Model.ModelTables.Item(index)
                                    for index in range(1, int(verification.Model.ModelTables.Count) + 1))
                                if not any(str(table.SourceWorkbookConnection.Name).casefold() ==
                                           f'query - {source.name}'.casefold() for table in model_tables):
                                    raise RuntimeError(f'Staged workbook verification failed for Data Model table {source.name!r}.')
                            elif source.load_mode.value in ('new_sheet', 'existing_sheet'):
                                expected = load_expectations[source.name.casefold()]
                                found = None
                                for index in range(1, int(verification.Worksheets.Count) + 1):
                                    sheet = verification.Worksheets.Item(index)
                                    for table_index in range(1, int(sheet.ListObjects.Count) + 1):
                                        table = sheet.ListObjects.Item(table_index)
                                        try:
                                            connection_name = str(table.QueryTable.WorkbookConnection.Name).casefold()
                                        except Exception:
                                            continue
                                        if connection_name == f'query - {source.name}'.casefold():
                                            found = (str(sheet.Name), str(table.Range.Cells(1, 1).Address), str(table.Name))
                                if found != (expected['sheet'], expected['cell'], expected['table_name']):
                                    raise RuntimeError(f'Staged workbook verification failed for worksheet load {source.name!r}: expected {expected!r}, observed {found!r}.')
                    self._bounded_excel_call(excel, verify_staged_workbook,
                        'verify staged workbook query formulas and loads')
                except Exception as exc:
                    if isinstance(exc, RuntimeError) and 'verification' in str(exc).lower():
                        raise RuntimeError(safe_error(exc)) from exc
                    raise RuntimeError(f'Staged workbook verification failed: {type(exc).__name__}: {safe_error(exc)}') from exc
                finally:
                    try:
                        self._bounded_excel_call(excel,
                            lambda: self._workbook_operations.close_workbook(verification, False),
                            'close staged workbook verification')
                    except Exception as close_exc:
                        verification = None
                        raise RuntimeError(
                            f'Could not close the read-only verification workbook: '
                            f'{type(close_exc).__name__}: {safe_error(close_exc)}') from close_exc
                verification = None
                query = connection = model_tables = table = sheet = matches = None
                self._workbook_transactions.commit_stage(
                    staged_path, workbook_path, original_hash, operation='import')
                staged_path = None
            if logger:
                for result in results:
                    logger(result)
            return results
        except Exception as exc:
            rollback_errors = []
            if workbook is not None:
                for action, name, previous in reversed(journal):
                    try:
                        def rollback_change():
                            query = self._get_existing_query(workbook, name)
                            if action == 'created':
                                if query is not None: query.Delete()
                                if self._get_existing_query(workbook,name) is not None:
                                    raise RuntimeError('created query still exists after rollback')
                            elif action == 'updated':
                                if query is None:
                                    raise RuntimeError('updated query is missing during rollback')
                                query.Formula = previous
                                if str(query.Formula) != str(previous):
                                    raise RuntimeError('restored query formula did not verify')
                            elif action == 'connection_created':
                                connection=self._get_existing_connection(workbook,name)
                                if connection is not None:connection.Delete()
                                if self._get_existing_connection(workbook,name) is not None:
                                    raise RuntimeError('created connection still exists after rollback')
                            elif action == 'replaced':
                                if query is not None: query.Delete()
                                if previous is not None: workbook.Queries.Add(name, previous)
                                restored=self._get_existing_query(workbook,name)
                                if previous is not None and (restored is None or str(restored.Formula)!=str(previous)):
                                    raise RuntimeError('re-added query formula did not verify')
                        self._bounded_excel_call(excel, rollback_change, f'rollback {name}')
                    except Exception as rollback_exc:
                        rollback_errors.append(f'{name}: {type(rollback_exc).__name__}: {safe_error(rollback_exc)}')
            detail = f'{type(exc).__name__}: {safe_error(exc)}'
            if journal:
                detail += f' (rolled back {len(journal)} change(s).)'
            if rollback_errors:
                detail += ' Rollback errors: ' + '; '.join(rollback_errors)
            error_query=current_source.name.casefold() if current_source is not None else None
            rollback_by_query={}
            rolled_back=set()
            for item in rollback_errors:
                name,_,message=item.partition(':')
                rollback_by_query[name.casefold()]=message.strip() or item
            rolled_back={name.casefold() for action,name,_previous in journal
                         if name.casefold() not in rollback_by_query}
            completed={row.query_name.casefold():row for row in results}
            plan_by_query={source.name.casefold():(source,status,result)
                           for source,_formula,_replacing,status,result in plan}
            failure_results=[]
            for source in sequence:
                key=source.name.casefold()
                if key==error_query:
                    status=QueryStatus.FAILED
                    message=f'Failed during {current_phase}: {safe_error(exc)}'
                elif key in rollback_by_query:
                    status=QueryStatus.FAILED
                    message='Rollback could not be verified: '+rollback_by_query[key]
                elif key in rolled_back:
                    status=QueryStatus.ROLLED_BACK
                    message='Staged changes were reverted; the original workbook was retained.'
                elif key in completed:
                    prior=completed[key]
                    status=prior.status;message=prior.message
                elif key in plan_by_query and plan_by_query[key][1] is None:
                    status=plan_by_query[key][2].status;message=plan_by_query[key][2].message
                else:
                    status=QueryStatus.NOT_ATTEMPTED
                    message='No change was committed because the operation stopped earlier.'
                failure_results.append(InstallResult(source.name,status,message))
            raise InstallOperationError(detail,failure_results) from exc
        finally:
            if workbook is not None:
                try: self._bounded_excel_call(excel,
                    lambda: self._workbook_operations.close_workbook(workbook, False),
                    'close staged workbook after import')
                except Exception: pass
                workbook = None
            if excel is not None:
                # Release child COM proxies before Quit. ModelTable proxies in particular can
                # keep an otherwise closed Excel instance alive past Application.Quit.
                existing = query = connection = model_tables = table = sheet = query_table = matches = None
                verification = None
                import gc
                gc.collect()
                if previous_security is not None:
                    try: self._bounded_excel_call(excel,
                        lambda: setattr(excel, 'AutomationSecurity', previous_security),
                        'restore import macro security')
                    except Exception: pass
                try:
                    if self._uses_default_factory:
                        owned_pid = owned_excel_pid(excel)
                        def release_com_apartment():
                            nonlocal excel, com_initialized
                            excel = None
                            if com_initialized:
                                import pythoncom
                                pythoncom.CoUninitialize()
                                com_initialized = False
                        terminated = quit_owned_excel(excel, quit_call=lambda:
                            self._bounded_excel_call(excel, lambda: excel.Quit(),
                                'quit owned Excel instance after import'),
                            after_quit=release_com_apartment)
                        if logger:
                            logger(f'Closed captured app-owned Excel PID {owned_pid}; forced termination={terminated}.')
                    else:
                        self._bounded_excel_call(excel, lambda: excel.Quit(),
                            'quit owned Excel instance after import')
                except Exception as cleanup_exc:
                    if logger:
                        logger(f'Excel cleanup could not be confirmed ({type(cleanup_exc).__name__}); close the app-owned Excel window if it remains open.')
                excel = None
            if com_initialized:
                try:
                    import pythoncom
                    pythoncom.CoUninitialize()
                except Exception: pass
            if staged_path is not None:
                cleanup_stage_and_report(self._workbook_transactions, staged_path, logger=logger)
