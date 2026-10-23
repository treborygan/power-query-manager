from pathlib import Path
import os
import re
import tempfile
from .models import LoadMode
from .query_loader import SUPPORTED_EXTENSIONS
from .windows_file_users import ensure_no_active_file_users
from .path_safety import contains_reparse_point

_INVALID_QUERY_CHARS = set('[]:*?/\\')
_INVALID_SHEET_CHARS = set('[]:*?/\\')
_TABLE_NAME = re.compile(r'^[A-Za-z_][A-Za-z0-9_.]{0,254}$')


def validate_sheet_name(name):
    if not name or not name.strip() or name != name.strip() or len(name) > 31:
        raise ValueError('Worksheet name must be 1–31 characters with no leading or trailing spaces.')
    if any(char in _INVALID_SHEET_CHARS for char in name) or name.startswith("'") or name.endswith("'"):
        raise ValueError(f'Worksheet name {name!r} contains characters Excel does not allow.')
    return name


def validate_cell_address(value):
    if not isinstance(value, str) or not re.fullmatch(r'\$?[A-Za-z]{1,3}\$?[1-9][0-9]{0,6}', value):
        raise ValueError('Starting cell must be an A1-style address such as A1 or $D$7.')
    match = re.fullmatch(r'\$?([A-Za-z]{1,3})\$?([1-9][0-9]{0,6})', value)
    column = 0
    for char in match.group(1).upper():
        column = column * 26 + ord(char) - ord('A') + 1
    row = int(match.group(2))
    if column > 16384 or row > 1048576:
        raise ValueError('Starting cell is outside Excel worksheet limits.')
    return value.replace('$', '').upper()


def validate_table_name(value):
    if not isinstance(value, str) or not _TABLE_NAME.fullmatch(value):
        raise ValueError('Excel table name must start with a letter or underscore and contain only letters, numbers, underscores, or periods.')
    if re.fullmatch(r'[A-Za-z]{1,3}[1-9][0-9]{0,6}', value) or value.upper().startswith('R') and 'C' in value.upper():
        raise ValueError('Excel table name cannot look like a cell reference.')
    return value

def validate_workbook(path):
    path = Path(path)
    if not path.exists() or not path.is_file():
        raise ValueError('The selected Excel workbook does not exist.')
    if path.suffix.lower() not in {'.xlsx', '.xlsm'}:
        raise ValueError('Workbook must be an .xlsx or .xlsm file.')
    if os.name == 'nt':
        zone_path = Path(str(path) + ':Zone.Identifier')
        try:
            with zone_path.open('rb') as stream:
                raw_zone = stream.read(4097)
        except FileNotFoundError:
            raw_zone = None
        except OSError as exc:
            raise ValueError('Windows security-zone metadata could not be checked; workbook automation was stopped.') from exc
        if raw_zone is not None:
            if len(raw_zone) > 4096:
                raise ValueError('Windows security-zone metadata is unexpectedly large; workbook automation was stopped.')
            try:
                zone_text = raw_zone.decode('utf-8-sig')
                zone_matches = re.findall(r'(?im)^\s*ZoneId\s*=\s*(\d+)\s*$', zone_text)
                if not zone_matches or len({int(value) for value in zone_matches}) != 1:
                    raise ValueError('Windows security-zone metadata is present but cannot be interpreted; workbook automation was stopped.')
                zone_id = int(zone_matches[0])
            except (UnicodeError, ValueError) as exc:
                if isinstance(exc, ValueError) and str(exc).startswith('Windows security-zone'):
                    raise
                raise ValueError('Windows security-zone metadata is invalid; workbook automation was stopped.') from exc
            if zone_id >= 3:
                raise ValueError(
                    f'Workbook carries a Windows Internet or Restricted zone mark (ZoneId={zone_id}); '
                    'review it in Excel/Windows and resolve the mark before using workbook automation.')
    return path

def validate_workbook_for_write(path,*,active_file_users=None):
    """Fail closed when the workbook or its same-folder staging area is not writable."""
    path=validate_workbook(path)
    if contains_reparse_point(path):
        raise ValueError('The selected workbook path contains a symbolic link or Windows reparse point. Select a path without reparse points before writing.')
    lock_path=path.with_name('~$'+path.name)
    if lock_path.exists():
        raise ValueError(f'The workbook appears to be open in Excel (owner lock file exists: {lock_path.name}). Close it and retry.')
    try:ensure_no_active_file_users(path,active_file_users)
    except RuntimeError as exc:raise ValueError(str(exc)) from exc
    try:
        descriptor=os.open(path,os.O_RDWR)
    except OSError as exc:
        raise ValueError('The workbook is read-only or cannot be opened for writing. Close it or choose a writable copy.') from exc
    else:
        os.close(descriptor)
    try:
        with tempfile.NamedTemporaryFile(prefix='.pqm-write-check-',dir=path.parent,delete=True):
            pass
    except OSError as exc:
        raise ValueError('The workbook folder is not writable for a verified stage and backup. Choose a writable folder.') from exc
    return path

def validate_query_name(name):
    if not name or not name.strip():
        raise ValueError('Query name cannot be empty.')
    if name != name.strip():
        raise ValueError(f'Query name {name!r} cannot start or end with whitespace.')
    if len(name) > 255:
        raise ValueError(f'Query name {name!r} is longer than Excel’s 255-character limit.')
    bad = sorted(set(name) & _INVALID_QUERY_CHARS)
    if bad:
        raise ValueError(f'Query name {name!r} contains unsupported character(s): {", ".join(bad)}.')
    return name

def validate_query_sources(sources):
    enabled = [s for s in sources if s.enabled]
    if not enabled:
        raise ValueError('Select or enable at least one Power Query source file.')
    seen = set()
    for source in enabled:
        validate_query_name(source.name)
        key = source.name.casefold()
        if key in seen:
            raise ValueError(f'Duplicate query name detected (names are case-insensitive): {source.name}')
        seen.add(key)
        if not source.path.exists() or not source.path.is_file():
            raise ValueError(f'Query source file does not exist or is not a file: {source.path}')
        if source.path.suffix.lower() not in SUPPORTED_EXTENSIONS:
            raise ValueError(f'Unsupported query source extension for {source.path}. Use .m or .pq.')
        if not isinstance(source.order, int) or source.order < 0:
            raise ValueError(f'Order for {source.name} must be a whole number zero or greater.')
        if source.load_mode not in set(LoadMode):
            raise ValueError(f'Unsupported load mode for {source.name}.')
        if source.load_mode in {LoadMode.CONNECTION_ONLY, LoadMode.NEW_SHEET, LoadMode.EXISTING_SHEET, LoadMode.DATA_MODEL} and (len(source.name)>246 or not re.fullmatch(r'[\w .-]+',source.name,re.UNICODE)):
            raise ValueError(f'Query name {source.name!r} contains characters or length unsupported by the Power Query connection string.')
        if source.load_mode == LoadMode.EXISTING_SHEET:
            validate_sheet_name(source.sheet_name)
            validate_cell_address(source.cell)
        elif source.load_mode == LoadMode.NEW_SHEET:
            if source.cell is not None:
                raise ValueError(f'New-sheet load for query {source.name!r} does not accept a cell address.')
            if source.sheet_name is not None:
                validate_sheet_name(source.sheet_name)
        elif source.load_mode in {LoadMode.DEFINITION_ONLY, LoadMode.CONNECTION_ONLY, LoadMode.DATA_MODEL, LoadMode.PRESERVE_EXISTING}:
            if any(value is not None for value in (source.sheet_name, source.cell, source.table_name)):
                raise ValueError(f'Load mode {source.load_mode.value!r} does not accept worksheet destination fields.')
        if source.table_name is not None and source.load_mode in {LoadMode.NEW_SHEET, LoadMode.EXISTING_SHEET}:
            validate_table_name(source.table_name)
    return sources

def order_warnings(sources):
    buckets = {}
    for s in sources:
        if s.enabled:
            buckets.setdefault(s.order, []).append(s)
    return [f'Order {n} is shared by: {", ".join(sorted(x.name for x in items))}; filename order breaks the tie.' for n, items in sorted(buckets.items()) if len(items) > 1]

def ordered_sources(sources):
    return sorted((s for s in sources if s.enabled), key=lambda s: (s.order, s.path.name.casefold(), str(s.path).casefold()))
