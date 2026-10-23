"""Local, bounded failure diagnostics that never persist exception messages or M code."""
from datetime import datetime, timezone
from dataclasses import dataclass
import json
import os
from pathlib import Path
import re
import threading
from .redaction import safe_error
from .path_safety import contains_reparse_point


class ApprovalRequired(ValueError):
    """The operation is fail-closed until the user supplies explicit approval."""

_LOCK=threading.RLock()
_MAX_BYTES=1024*1024
_BACKUP_COUNT=3
_OPERATION=re.compile(r'^[a-z][a-z0-9_-]{0,39}$')


@dataclass(frozen=True)
class FailureCategory:
    code:str
    title:str
    guidance:str


_CATEGORIES={
    'timeout':FailureCategory('timeout','Excel operation timed out',
        'Close any Excel dialog and retry. The staged workbook should not be treated as committed.'),
    'unsupported_capability':FailureCategory('unsupported_capability','Load mode is not supported by this Excel build',
        'Run the Excel capability probe, or choose a mode that the probe confirms.'),
    'workbook_changed':FailureCategory('workbook_changed','Workbook changed during review',
        'Reopen the current workbook, inspect it again, and review the operation before retrying.'),
    'backup_failed':FailureCategory('backup_failed','Workbook backup could not be verified',
        'Check available disk space and folder permissions. The import was stopped.'),
    'verification_failed':FailureCategory('verification_failed','Saved workbook verification failed',
        'Do not treat the operation as successful. Check the target and keep the verified backup.'),
    'workbook_locked_or_read_only':FailureCategory('workbook_locked_or_read_only','Workbook is locked or read-only',
        'Close it in Excel, remove any stale owner lock only after confirming it is closed, and check write permissions.'),
    'permission_or_access':FailureCategory('permission_or_access','File or folder access was denied',
        'Choose a writable local folder and confirm the workbook is not open or controlled by another process.'),
    'query_conflict':FailureCategory('query_conflict','Existing query or load mapping is ambiguous',
        'Inspect the workbook query and its worksheet/Data Model loads before replacing it.'),
    'source_changed':FailureCategory('source_changed','A query source changed after review',
        'Reload the source and review its current content before applying.'),
    'excel_automation':FailureCategory('excel_automation','Excel automation could not complete',
        'Close modal Excel dialogs, confirm desktop Excel is available, and retry with a disposable workbook.'),
    'input_validation':FailureCategory('input_validation','Input needs correction',
        'Check the workbook path, source file, query name, encoding, and destination settings.'),
    'approval_required':FailureCategory('approval_required','Explicit approval is required',
        'Review the displayed operation or mapping preview and provide its exact confirmation phrase. No write was approved.'),
    'operation_failed':FailureCategory('operation_failed','Operation could not be completed',
        'Review the operation log and local diagnostic record before retrying.'),
}


def default_log_directory():
    """Return a per-user local-data directory without requiring elevation."""
    root=os.environ.get('LOCALAPPDATA') or os.environ.get('APPDATA')
    if root:
        base=Path(root)
    else:
        base=Path.home()/'.local'/'share'
    return base/'PowerQueryManager'/'logs'


def classify_failure(error):
    """Map an error to a stable, user-facing category and next step."""
    message=str(error).casefold()
    if isinstance(error,ApprovalRequired):
        key='approval_required'
    elif isinstance(error,TimeoutError) or any(word in message for word in ('timeout','timed out','exceeded')):
        key='timeout'
    elif any(word in message for word in ('requires a successful capability probe','not supported','unsupported','not available on this office build')):
        key='unsupported_capability'
    elif 'source changed after review' in message:
        key='source_changed'
    elif any(word in message for word in ('workbook changed after review','changed during import','concurrent change')):
        key='workbook_changed'
    elif 'backup' in message:
        key='backup_failed'
    elif any(word in message for word in ('verification','verify','hash mismatch')):
        key='verification_failed'
    elif any(word in message for word in ('read-only','readonly','lock file','locked')):
        key='workbook_locked_or_read_only'
    elif any(word in message for word in ('permission','access denied','not writable')):
        key='permission_or_access'
    elif any(word in message for word in ('query conflict','duplicate query','duplicate connection','load association','query mapping','cannot safely replace')):
        key='query_conflict'
    elif any(word in message for word in ('com error','excel automation','excel could not','excel did not','excel is not')):
        key='excel_automation'
    elif isinstance(error,(ValueError,TypeError)):
        key='input_validation'
    else:
        key='operation_failed'
    return _CATEGORIES[key]


def format_failure(error):
    """Render an actionable error while applying the standard secret redaction."""
    category=classify_failure(error)
    message=f'{category.title} ({category.code}). {category.guidance}\nDetails: {safe_error(error)}'
    results=tuple(getattr(error,'results',()))
    if results:
        lines=[]
        for row in results:
            status=getattr(row.status,'value',row.status)
            lines.append(f'  {row.query_name}: {status} — {safe_error(row.message)}')
        message+='\nPer-query outcomes:\n'+'\n'.join(lines)
    return message


def _rotate(path):
    if any(contains_reparse_point(path.with_name(path.name + suffix))
           for suffix in ('', '.1', '.2', '.3')):
        raise OSError('Refusing to rotate a diagnostic log through a symbolic link or reparse point.')
    oldest=path.with_name(path.name+f'.{_BACKUP_COUNT}')
    try:oldest.unlink()
    except FileNotFoundError:pass
    for index in range(_BACKUP_COUNT-1,0,-1):
        current=path.with_name(path.name+f'.{index}')
        if current.exists():os.replace(current,path.with_name(path.name+f'.{index+1}'))
    if path.exists():os.replace(path,path.with_name(path.name+'.1'))


def record_failure(operation,error,*,directory=None):
    """Append a compact error record; diagnostics never contain exception text or formulas.

    Logging is best-effort so a full or restricted user profile cannot mask the
    original operation failure. Pass ``directory`` only for tests or explicit
    local support tooling; normal calls use the current user's local data path.
    """
    operation=str(operation).casefold()
    if not _OPERATION.fullmatch(operation):operation='operation'
    root=Path(directory) if directory is not None else default_log_directory()
    path=root/'PowerQueryManager'/'logs'/'errors.jsonl' if directory is not None else root/'errors.jsonl'
    record={
        'timestamp_utc':datetime.now(timezone.utc).isoformat(timespec='seconds'),
        'operation':operation,
        'error_type':type(error).__name__[:80],
        'category':classify_failure(error).code,
    }
    line=(json.dumps(record,ensure_ascii=True,separators=(',',':'))+'\n').encode('utf-8')
    try:
        with _LOCK:
            if contains_reparse_point(path.parent) or contains_reparse_point(path):
                return None
            path.parent.mkdir(parents=True,exist_ok=True)
            if contains_reparse_point(path.parent) or contains_reparse_point(path):
                return None
            if path.exists() and path.stat().st_size+len(line)>_MAX_BYTES:_rotate(path)
            if contains_reparse_point(path.parent) or contains_reparse_point(path):
                return None
            with path.open('ab') as stream:stream.write(line)
        return path
    except OSError:
        return None
