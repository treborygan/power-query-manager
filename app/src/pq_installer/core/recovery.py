"""Conservative recovery helpers for abandoned import/refresh stage files."""
from pathlib import Path
import hashlib
import os
import shutil
import tempfile
import time
import zipfile
from datetime import datetime

from .validation import ensure_no_active_file_users
from .path_safety import contains_reparse_point


def _sha256(path):
    digest=hashlib.sha256()
    with Path(path).open('rb') as stream:
        for chunk in iter(lambda:stream.read(1024*1024),b''):
            digest.update(chunk)
    return digest.hexdigest()


def find_staged_workbooks(workbook_path):
    """Return matching hidden stage files beside a selected workbook."""
    workbook=Path(os.path.abspath(os.fspath(workbook_path)))
    if contains_reparse_point(workbook):
        return ()
    if workbook.suffix.casefold() not in {'.xlsx','.xlsm'}:
        return ()
    if not workbook.parent.is_dir():
        return ()
    candidates=(path for path in workbook.parent.glob(f'.{workbook.stem}.pqstage-*{workbook.suffix}')
                if not contains_reparse_point(path) and path.is_file()
                and Path(os.path.abspath(path)).parent==workbook.parent
                and Path(os.path.abspath(path))!=workbook)
    return tuple(sorted(candidates,key=lambda path:(path.stat().st_mtime_ns,path.name),reverse=True))


def diagnose_staged_workbook(workbook_path, staged_path, *, now=None):
    """Report stage age and basic Open XML container structure without modifying it.

    This is only a triage check. Excel must still open the candidate read-only
    before any recovery copy can be published.
    """
    workbook=Path(os.path.abspath(os.fspath(workbook_path)))
    stage=Path(os.path.abspath(os.fspath(staged_path)))
    if contains_reparse_point(workbook) or contains_reparse_point(stage):
        raise ValueError('The workbook or staged path contains a symbolic link or Windows reparse point.')
    if stage not in find_staged_workbooks(workbook):
        raise ValueError('The selected file is not a matching staged workbook for this target.')
    info=stage.stat()
    modified_at=datetime.fromtimestamp(info.st_mtime).astimezone().isoformat(timespec='seconds')
    age_seconds=max(0.0,(time.time() if now is None else float(now))-info.st_mtime)
    result={'path':str(stage),'size_bytes':info.st_size,
            'modified_at':modified_at,'age_seconds':age_seconds}
    if info.st_size<=0:
        result.update(status='empty',message='The staged file is empty; it cannot be recovered.')
        return result
    try:
        with zipfile.ZipFile(stage,'r') as archive:
            names=set(archive.namelist())
    except (OSError,EOFError,zipfile.BadZipFile,zipfile.LargeZipFile):
        result.update(status='unreadable_container',
                      message='The staged file does not have a readable ZIP container.')
        return result
    missing={'[Content_Types].xml','xl/workbook.xml'}-names
    if missing:
        result.update(status='missing_workbook_parts',
                      message='The ZIP is readable but missing required workbook part(s): '
                              +', '.join(sorted(missing))+'.')
        return result
    result.update(status='structure_present',
                  message='Required Open XML workbook parts are present; Excel read-only verification is still required.')
    return result


def save_verified_recovery_copy(workbook_path, staged_path, output_path, *, expected_sha256):
    """Atomically save a validated stage to a new sibling/user-selected path.

    This helper deliberately never replaces the original workbook or an existing
    output. Callers must first verify the candidate by opening it read-only in Excel.
    """
    if any(contains_reparse_point(candidate) for candidate in (workbook_path, staged_path, output_path)):
        raise ValueError('The selected stage is not a matching staged workbook, or a recovery path contains a symbolic link or Windows reparse point.')
    workbook=Path(os.path.abspath(os.fspath(workbook_path)))
    stage=Path(os.path.abspath(os.fspath(staged_path)))
    output=Path(os.path.abspath(os.fspath(output_path)))
    if workbook.suffix.casefold() not in {'.xlsx','.xlsm'}:
        raise ValueError('Recovery requires an .xlsx or .xlsm workbook.')
    if stage not in find_staged_workbooks(workbook):
        raise ValueError('The selected file is not a matching staged workbook for this target.')
    if output==workbook or output==stage:
        raise ValueError('Choose a new recovery-copy path; the target and stage are preserved.')
    if output.suffix.casefold()!=workbook.suffix.casefold():
        raise ValueError('The recovery-copy extension must match the selected workbook.')
    if not output.parent.is_dir():
        raise ValueError('The recovery-copy folder does not exist.')
    if output.exists():
        raise FileExistsError('The recovery-copy destination already exists; choose a new name.')
    ensure_no_active_file_users(stage)
    if _sha256(stage)!=expected_sha256:
        raise RuntimeError('The staged workbook changed after verification; verify it again before recovery.')
    descriptor,temporary_name=tempfile.mkstemp(prefix=f'.{output.name}.recover-',suffix=output.suffix,dir=output.parent)
    temporary=Path(temporary_name)
    try:
        with os.fdopen(descriptor,'wb') as destination,stage.open('rb') as source:
            shutil.copyfileobj(source,destination,1024*1024)
            destination.flush();os.fsync(destination.fileno())
        if _sha256(stage)!=expected_sha256 or _sha256(temporary)!=expected_sha256:
            raise RuntimeError('The staged workbook changed during recovery copy; no output was published.')
        if any(contains_reparse_point(candidate) for candidate in
               (workbook, stage, output, temporary)):
            raise ValueError('A recovery path changed to contain a symbolic link or Windows reparse point; no output was published.')
        try:
            os.link(temporary,output)
        except FileExistsError:
            raise FileExistsError('The recovery-copy destination was created concurrently; choose a new name.')
        return output
    finally:
        try:temporary.unlink(missing_ok=True)
        except OSError:pass
