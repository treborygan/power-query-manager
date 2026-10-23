"""Injectable file boundary for staged workbook transactions."""
from __future__ import annotations

import hashlib
import os
import shutil
import time
import uuid
from pathlib import Path
from typing import Protocol

from .windows_file_users import ensure_no_active_file_users
from .path_safety import contains_reparse_point


class WorkbookTransactionAdapter(Protocol):
    """File transaction port used by import and refresh COM services."""

    def prepare_stage(self, workbook, expected_sha256: str, *, operation: str) -> Path: ...
    def commit_stage(self, stage, workbook, expected_sha256: str, *, operation: str) -> None: ...
    def cleanup_stage(self, stage) -> bool: ...


def _sha256(path) -> str:
    digest = hashlib.sha256()
    with Path(path).open('rb') as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b''):
            digest.update(chunk)
    return digest.hexdigest()


def _extended_windows_path(path) -> str:
    value = os.path.abspath(str(path))
    if value.startswith('\\\\?\\'):
        return value
    if value.startswith('\\\\'):
        return '\\\\?\\UNC\\' + value[2:]
    return '\\\\?\\' + value


def _call_replace_file(replaced, replacement, backup):
    import ctypes
    kernel32 = ctypes.WinDLL('kernel32', use_last_error=True)
    replace_file = kernel32.ReplaceFileW
    replace_file.argtypes = [ctypes.c_wchar_p, ctypes.c_wchar_p, ctypes.c_wchar_p,
                             ctypes.c_uint32, ctypes.c_void_p, ctypes.c_void_p]
    replace_file.restype = ctypes.c_int
    if not replace_file(_extended_windows_path(replaced), _extended_windows_path(replacement),
                        _extended_windows_path(backup) if backup else None, 0, None, None):
        raise ctypes.WinError(ctypes.get_last_error())


def _call_move_file(source, destination):
    import ctypes
    kernel32 = ctypes.WinDLL('kernel32', use_last_error=True)
    move_file = kernel32.MoveFileW
    move_file.argtypes = [ctypes.c_wchar_p, ctypes.c_wchar_p]
    move_file.restype = ctypes.c_int
    if not move_file(_extended_windows_path(source), _extended_windows_path(destination)):
        raise ctypes.WinError(ctypes.get_last_error())


def _replace_existing_file(stage: Path, target: Path, expected_sha256: str) -> None:
    """Atomically replace an existing workbook and retain recoverable state on failure."""
    stage = Path(stage)
    target = Path(target)
    replacement_sha256 = _sha256(stage)
    if os.name != 'nt':
        # The application targets Windows; keep a checked fallback for portable
        # tests and development environments.
        if not target.is_file():
            raise FileNotFoundError(f'Workbook disappeared before commit: {target.name}')
        os.replace(stage, target)
        return

    rollback = target.with_name(f'.{target.name}.pqcommit-{uuid.uuid4().hex}.bak')

    try:
        # ReplaceFile requires the destination to exist. Its backup argument also
        # lets us restore the old name if Windows reports a partial replacement.
        _call_replace_file(target, stage, rollback)
    except Exception as exc:
        if rollback.is_file() and not target.exists():
            try:
                _call_move_file(rollback, target)
            except Exception as restore_exc:
                raise RuntimeError(
                    f'Workbook replacement failed and the original could not be restored; '
                    f'a recovery copy remains at {rollback.name}.') from restore_exc
        error_code = getattr(exc, 'winerror', None)
        if error_code is None and getattr(exc, 'args', ()) and isinstance(exc.args[0], int):
            error_code = exc.args[0]
        if error_code in (2, 3) and not target.exists():
            raise FileNotFoundError(f'Workbook disappeared before commit: {target.name}') from exc
        raise

    if _sha256(target) != replacement_sha256 or _sha256(rollback) != expected_sha256:
        if rollback.is_file() and _sha256(rollback) == expected_sha256:
            try:
                _call_replace_file(target, rollback, stage)
            except Exception as restore_exc:
                raise RuntimeError(
                    f'Committed workbook verification failed and rollback was not confirmed; '
                    f'a recovery copy remains at {rollback.name}.') from restore_exc
        raise RuntimeError('Atomic workbook replacement did not preserve the verified stage and original workbook hashes.')
    try:
        rollback.unlink()
    except OSError:
        # This is only the adapter's transient rollback copy; the independently
        # verified timestamped user backup remains in place.
        pass


class LocalWorkbookTransactionAdapter:
    """Same-directory stage/copy/commit operations with verified atomic publish."""

    def __init__(self, *, check_active=None):
        self._check_active = check_active or ensure_no_active_file_users

    @staticmethod
    def _operation(value):
        if value not in {'import', 'refresh'}:
            raise ValueError('Workbook transaction operation must be import or refresh.')
        return value

    def prepare_stage(self, workbook, expected_sha256: str, *, operation: str) -> Path:
        operation = self._operation(operation)
        target = Path(workbook)
        if contains_reparse_point(target):
            raise ValueError('The selected workbook path contains a symbolic link or Windows reparse point. Select a path without reparse points before writing.')
        stage = target.with_name(f'.{target.stem}.pqstage-{uuid.uuid4().hex}{target.suffix}')
        try:
            shutil.copy2(target, stage)
            if _sha256(target) != expected_sha256 or _sha256(stage) != expected_sha256:
                raise RuntimeError(
                    f'Workbook changed while the {operation} stage was being created; {operation} was stopped.')
            return stage
        except Exception:
            self.cleanup_stage(stage)
            raise

    def commit_stage(self, stage, workbook, expected_sha256: str, *, operation: str) -> None:
        operation = self._operation(operation)
        stage = Path(stage)
        target = Path(workbook)
        if contains_reparse_point(target):
            raise RuntimeError('Workbook path changed to contain a symbolic link or Windows reparse point; staged changes were not committed.')
        if contains_reparse_point(stage):
            raise RuntimeError('Staged workbook path contains a symbolic link or Windows reparse point; staged changes were not committed.')
        try:
            if stage.parent.resolve(strict=True) != target.parent.resolve(strict=True):
                raise RuntimeError('Staged workbook is not a sibling of the target; staged changes were not committed.')
        except OSError as exc:
            raise RuntimeError('Staged workbook directory could not be verified; staged changes were not committed.') from exc
        if not stage.is_file() or stage.stat().st_size == 0:
            if operation == 'import':
                message = 'Staged workbook verification failed: Excel produced an empty file.'
            else:
                message = 'Excel did not produce a verifiable staged workbook; original file was retained.'
            raise RuntimeError(message)
        if _sha256(target) != expected_sha256:
            raise RuntimeError(
                f'Original workbook changed during {operation}; staged changes were discarded to preserve the newer file.')
        self._check_active(target)
        if _sha256(target) != expected_sha256:
            raise RuntimeError(
                f'Original workbook changed during final {operation} commit checks; staged changes were discarded.')
        # The active-use and hash checks may take time. Revalidate paths at the
        # last practical boundary before replacing the user-selected workbook.
        if contains_reparse_point(target) or contains_reparse_point(stage):
            raise RuntimeError('Workbook or staged path changed to contain a symbolic link or Windows reparse point; staged changes were not committed.')
        try:
            if stage.parent.resolve(strict=True) != target.parent.resolve(strict=True):
                raise RuntimeError('Staged workbook is no longer a sibling of the target; staged changes were not committed.')
        except OSError as exc:
            raise RuntimeError('Workbook directories could not be reverified; staged changes were not committed.') from exc
        _replace_existing_file(stage, target, expected_sha256)

    def cleanup_stage(self, stage) -> bool:
        if stage is None:
            return True
        path = Path(stage)
        for attempt in range(10):
            try:
                path.unlink()
                return True
            except FileNotFoundError:
                return True
            except OSError:
                if attempt < 9:
                    time.sleep(0.1)
        try:
            path.stat()
        except FileNotFoundError:
            return True
        except OSError:
            # An unreadable path is not proof that the staged workbook is gone.
            return False
        return False


def cleanup_stage_and_report(adapter: WorkbookTransactionAdapter, stage, *, logger=None) -> bool:
    """Remove a failed stage and surface a retained recovery candidate to the caller."""
    if stage is None:
        return True
    try:
        cleaned = bool(adapter.cleanup_stage(stage))
    except Exception:
        cleaned = False
    if not cleaned and logger is not None:
        try:
            logger(f'Temporary staged workbook remains and may need recovery: {Path(stage).name}')
        except Exception:
            pass
    return cleaned
