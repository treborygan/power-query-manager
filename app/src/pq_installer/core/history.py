"""Local, redacted audit history for Power Query Manager actions."""

from __future__ import annotations

import json
import hashlib
import os
import re
import shutil
import sqlite3
import tempfile
import time
import uuid
from contextlib import closing, contextmanager
from datetime import datetime, timedelta, timezone
from pathlib import Path

from pq_installer.core.redaction import safe_error
from pq_installer.core.path_safety import contains_reparse_point


_DEFAULT_RETENTION_DAYS = 365
_MAX_RETENTION_DAYS = 3650
_FORMULA_KEYS = re.compile(
    r"(?i)(?:^|[_ -])(?:formula|m[ _-]?code|expression|query[ _-]?text|script|code)(?:$|[_ -])"
)
_SECRET_KEY = re.compile(r'(?i)(?:password|passwd|pwd|token|secret|credential|api[_ -]?key)')
_SAFE_METADATA_KEYS = frozenset({
    'query_name', 'profile_name', 'workbook', 'comparison', 'query_count',
    'refreshed', 'status', 'action', 'snapshot_id', 'snapshot_sha256',
    'snapshot_size', 'snapshot_name', 'previous_snapshot_id',
    'before_snapshot_id', 'after_snapshot_id',
    'added', 'modified', 'unchanged', 'removed', 'conflict',
    'path_mapping_count', 'environment_mapping_count',
    'created_count', 'replaced_count', 'skipped_count',
    'workbook_count',
})
_MAX_SNAPSHOT_BYTES = 2 * 1024 * 1024 * 1024


class HistoryCorruptionError(sqlite3.DatabaseError):
    """The local history database is unreadable or incompatible with its schema."""


def _sha256(path):
    digest=hashlib.sha256()
    with Path(path).open('rb') as stream:
        for chunk in iter(lambda:stream.read(1024*1024),b''):
            digest.update(chunk)
    return digest.hexdigest()


def _utc(value: datetime | None = None) -> datetime:
    value = value or datetime.now(timezone.utc)
    if value.tzinfo is None:
        raise ValueError('History timestamps must include a timezone.')
    return value.astimezone(timezone.utc)


def _retention(value: int) -> int:
    if isinstance(value, bool) or not isinstance(value, int) or not 1 <= value <= _MAX_RETENTION_DAYS:
        raise ValueError(f'Retention must be a whole number from 1 to {_MAX_RETENTION_DAYS} days.')
    return value


def _safe_details(value, depth=0):
    if depth > 12:
        raise ValueError('History details are nested too deeply.')
    if isinstance(value, dict):
        clean = {}
        for key, item in value.items():
            if not isinstance(key, str):
                raise ValueError('History detail keys must be text.')
            if _FORMULA_KEYS.search(key):
                continue
            if _SECRET_KEY.search(key):
                clean[key[:128]] = '[redacted]'
                continue
            if key.casefold() not in _SAFE_METADATA_KEYS:
                continue
            clean[key[:128]] = _safe_details(item, depth + 1)
        return clean
    if isinstance(value, (list, tuple)):
        if len(value) > 500:
            raise ValueError('History detail lists cannot contain more than 500 items.')
        return [_safe_details(item, depth + 1) for item in value]
    if isinstance(value, str):
        return safe_error(value)[:4096]
    if value is None or isinstance(value, (bool, int, float)):
        return value
    raise ValueError(f'Unsupported history detail value: {type(value).__name__}.')


class HistoryStore:
    """SQLite-backed history with redaction, retention, and explicit deletion."""

    def __init__(self, path, retention_days=None):
        self.path = Path(path).expanduser()
        if contains_reparse_point(self.path):
            raise ValueError('Local history database path contains a symbolic link or Windows reparse point.')
        self.path.parent.mkdir(parents=True, exist_ok=True)
        if contains_reparse_point(self.path):
            raise ValueError('Local history database path changed to a symbolic link or Windows reparse point.')
        self.recovery_archive = None
        try:
            self._initialize()
        except HistoryCorruptionError:
            self.recovery_archive = self._recover_corrupt_database()
        if retention_days is not None:
            self.set_retention_days(retention_days)
        else:
            with self._connect() as db:
                db.execute(
                    'INSERT OR IGNORE INTO settings(key, value) VALUES (?, ?)',
                    ('retention_days', str(_DEFAULT_RETENTION_DAYS)),
                )
        self.recovered_snapshot_files = self._recover_unregistered_snapshot_files()

    @contextmanager
    def _connect(self):
        db = sqlite3.connect(self.path, timeout=10)
        db.row_factory = sqlite3.Row
        db.execute('PRAGMA busy_timeout=10000')
        try:
            yield db
            db.commit()
        except Exception:
            db.rollback()
            raise
        finally:
            db.close()

    def _initialize(self):
        self._validate_existing_database()
        with self._connect() as db:
            db.execute('''CREATE TABLE IF NOT EXISTS history (
                version_id TEXT PRIMARY KEY,
                occurred_at TEXT NOT NULL,
                event TEXT NOT NULL,
                details_json TEXT NOT NULL
            )''')
            db.execute('''CREATE TABLE IF NOT EXISTS settings (
                key TEXT PRIMARY KEY,
                value TEXT NOT NULL
            )''')
            db.execute('''CREATE TABLE IF NOT EXISTS workbook_snapshots (
                snapshot_id TEXT PRIMARY KEY,
                occurred_at TEXT NOT NULL,
                workbook_path TEXT NOT NULL,
                snapshot_path TEXT NOT NULL,
                sha256 TEXT NOT NULL,
                size INTEGER NOT NULL
            )''')
            db.execute("INSERT OR IGNORE INTO settings(key,value) VALUES ('snapshots_enabled','0')")

    def _validate_existing_database(self):
        if not self.path.exists() or self.path.stat().st_size == 0:
            return
        expected = {
            'history': {'version_id', 'occurred_at', 'event', 'details_json'},
            'settings': {'key', 'value'},
            'workbook_snapshots': {'snapshot_id', 'occurred_at', 'workbook_path', 'snapshot_path', 'sha256', 'size'},
        }
        try:
            database_uri = self.path.resolve().as_uri() + '?mode=ro'
            with closing(sqlite3.connect(database_uri, timeout=10, uri=True)) as db:
                result = db.execute('PRAGMA quick_check').fetchone()
                if not result or result[0] != 'ok':
                    raise HistoryCorruptionError('Local history database failed SQLite quick_check.')
                tables = {row[0] for row in db.execute("SELECT name FROM sqlite_master WHERE type='table'")}
                for table, required in expected.items():
                    if table not in tables:
                        continue
                    columns = {row[1] for row in db.execute(f'PRAGMA table_info("{table}")')}
                    if not required.issubset(columns):
                        raise HistoryCorruptionError(f'Local history table {table!r} has an incompatible schema.')
        except sqlite3.DatabaseError as exc:
            message = str(exc).casefold()
            if any(marker in message for marker in ('file is not a database', 'malformed', 'unsupported file format')):
                raise HistoryCorruptionError('Local history database is corrupt or uses an unsupported format.') from exc
            raise

    def _recover_corrupt_database(self):
        """Quarantine unreadable history bytes, then initialize a fresh local store."""
        lock = Path(str(self.path) + '.recovery.lock')
        token = uuid.uuid4().hex
        deadline = time.monotonic() + 5
        descriptor = None
        while descriptor is None:
            try:
                descriptor = os.open(lock, os.O_CREAT | os.O_EXCL | os.O_WRONLY, 0o600)
            except FileExistsError:
                if time.monotonic() >= deadline:
                    raise RuntimeError(
                        'History recovery is already running or its lock is stale; no history files were changed. '
                        f'Recovery lock: {lock}. Close every Power Query Manager instance; after confirming no '
                        f'instance is recovering, remove only that recovery lock file: {lock}. Retry afterward. '
                        'Do not remove or move the history database or its sidecars.'
                    )
                time.sleep(0.05)
            else:
                try:
                    os.write(descriptor, token.encode('ascii'))
                except Exception:
                    os.close(descriptor)
                    descriptor = None
                    try:
                        if lock.exists():
                            lock.unlink()
                    except OSError:
                        pass
                    raise
        archive = self.path.with_name(
            f'{self.path.stem}.corrupt-{datetime.now(timezone.utc):%Y%m%dT%H%M%S}-{uuid.uuid4().hex[:8]}{self.path.suffix}')
        suffixes = ('-wal', '-shm', '-journal')
        moved = []
        main_moved = False
        try:
            # Another process may have recovered the store while this process waited.
            try:
                self._validate_existing_database()
            except HistoryCorruptionError:
                pass
            else:
                self._initialize()
                return None
            for suffix in suffixes:
                sidecar = Path(str(self.path) + suffix)
                if sidecar.exists():
                    archived_sidecar = Path(str(archive) + suffix)
                    os.replace(sidecar, archived_sidecar)
                    moved.append((sidecar, archived_sidecar))
            if not self.path.exists():
                raise RuntimeError('Corrupt history database disappeared during recovery; no replacement was created.')
            os.replace(self.path, archive)
            moved.append((self.path, archive))
            main_moved = True
            try:
                self._initialize()
            except Exception:
                # The original is still recoverable in the quarantine path.
                if not self.path.exists():
                    for original, archived in reversed(moved):
                        if archived.exists() and not original.exists():
                            os.replace(archived, original)
                    moved.clear()
                raise
            return archive
        except Exception:
            if not main_moved or not self.path.exists():
                for original, archived in reversed(moved):
                    if archived.exists() and not original.exists():
                        os.replace(archived, original)
            raise
        finally:
            os.close(descriptor)
            try:
                if lock.read_text(encoding='ascii') == token:
                    lock.unlink()
            except (FileNotFoundError, OSError, UnicodeError):
                pass

    @property
    def retention_days(self):
        with self._connect() as db:
            row = db.execute("SELECT value FROM settings WHERE key='retention_days'").fetchone()
        return _retention(int(row['value'])) if row else _DEFAULT_RETENTION_DAYS

    @property
    def snapshots_enabled(self):
        with self._connect() as db:
            row=db.execute("SELECT value FROM settings WHERE key='snapshots_enabled'").fetchone()
        return bool(int(row['value'])) if row else False

    def set_snapshots_enabled(self,enabled):
        if not isinstance(enabled,bool):raise ValueError('Snapshot preference must be boolean.')
        with self._connect() as db:
            db.execute('INSERT INTO settings(key,value) VALUES (?,?) ON CONFLICT(key) DO UPDATE SET value=excluded.value',
                       ('snapshots_enabled','1' if enabled else '0'))
        return enabled

    @property
    def snapshots_directory(self):
        return self.path.parent/(self.path.stem+'_snapshots')

    def _recover_unregistered_snapshot_files(self):
        """Move manager-owned snapshot files absent from the index to a visible recovery folder.

        A process or machine can stop after the atomic snapshot-file publish but before the
        SQLite row commits. These files are deliberately preserved rather than deleted.
        """
        directory = self.snapshots_directory
        if not directory.exists():
            return []
        if contains_reparse_point(directory):
            raise RuntimeError('History snapshot directory contains a symbolic link or Windows reparse point.')
        with self._connect() as db:
            # Serialize orphan discovery with snapshot publication so another
            # process cannot move a just-published file before its index commits.
            db.execute('BEGIN IMMEDIATE')
            rows = db.execute('SELECT snapshot_path FROM workbook_snapshots').fetchall()
            indexed = {
                os.path.normcase(os.path.abspath(os.fspath(row['snapshot_path'])))
                for row in rows
            }
            recovery_directory = self.path.with_name(self.path.name + '.recovered')
            recovered = []
            for candidate in directory.iterdir():
                is_manager_temp = candidate.name.startswith('.pqm-snapshot-')
                is_uuid_workbook = False
                if candidate.suffix.casefold() in {'.xlsx', '.xlsm'}:
                    try:
                        is_uuid_workbook = str(uuid.UUID(candidate.stem)).casefold() == candidate.stem.casefold()
                    except (ValueError, AttributeError):
                        pass
                if not (is_manager_temp or is_uuid_workbook):
                    continue
                if is_manager_temp and time.time() - candidate.stat().st_mtime < 24 * 60 * 60:
                    # Do not mistake another running manager's in-progress copy for crash residue.
                    continue
                lexical = os.path.normcase(os.path.abspath(os.fspath(candidate)))
                if lexical in indexed or contains_reparse_point(candidate) or not candidate.is_file():
                    continue
                if contains_reparse_point(recovery_directory):
                    raise RuntimeError('History recovery directory contains a symbolic link or Windows reparse point.')
                recovery_directory.mkdir(parents=True, exist_ok=True)
                if contains_reparse_point(recovery_directory):
                    raise RuntimeError('History recovery directory changed to a symbolic link or Windows reparse point.')
                recovered_path = recovery_directory / (
                    f'{candidate.stem}.recovered-{uuid.uuid4().hex}{candidate.suffix}'
                )
                try:
                    os.rename(candidate, recovered_path)
                except FileNotFoundError:
                    # Another manager instance may have recovered this same path first.
                    continue
                recovered.append(recovered_path)
        return recovered

    def _owned_snapshot_path(self, value):
        """Resolve a database path only when it remains inside our snapshot directory."""
        lexical_path=Path(os.path.abspath(os.fspath(value)))
        lexical_root=Path(os.path.abspath(os.fspath(self.snapshots_directory)))
        if contains_reparse_point(lexical_root) or contains_reparse_point(lexical_path):
            raise ValueError('Snapshot directory or file contains a symbolic link or Windows reparse point.')
        path=lexical_path.resolve()
        root=lexical_root.resolve()
        if root not in path.parents:
            raise ValueError('Snapshot path is outside the local history directory.')
        return path

    def create_snapshot(self,workbook_path,*,event='workbook_snapshot_created'):
        """Keep a bounded, hash-verified local workbook version and history reference."""
        workbook=Path(workbook_path).expanduser().resolve()
        if workbook.suffix.casefold() not in {'.xlsx','.xlsm'} or not workbook.is_file():
            raise ValueError('Snapshot source must be an existing .xlsx or .xlsm workbook.')
        size=workbook.stat().st_size
        if size<=0 or size>_MAX_SNAPSHOT_BYTES:
            raise ValueError('Workbook snapshot size must be between 1 byte and 2 GiB.')
        if not isinstance(event,str) or not event.strip():raise ValueError('Snapshot event is required.')
        before=_sha256(workbook)
        self.prune()
        directory=self.snapshots_directory
        directory.mkdir(parents=True,exist_ok=True)
        snapshot_id=str(uuid.uuid4())
        destination=directory/(snapshot_id+workbook.suffix.lower())
        temporary=None
        published=False
        try:
            with tempfile.NamedTemporaryFile(prefix='.pqm-snapshot-',suffix=workbook.suffix,
                    dir=directory,delete=False) as stream:
                temporary=Path(stream.name)
            shutil.copyfile(workbook,temporary)
            # Flush snapshot bytes before making the final UUID path visible or indexing it.
            with temporary.open('r+b') as snapshot_stream:
                snapshot_stream.flush()
                os.fsync(snapshot_stream.fileno())
            copied_size=temporary.stat().st_size
            copied_hash=_sha256(temporary)
            if copied_size!=size or copied_hash!=before or _sha256(workbook)!=before:
                raise RuntimeError('Workbook changed while its history snapshot was being captured.')
            at=_utc().isoformat(timespec='microseconds')
            with self._connect() as db:
                # Coordinate file publication with startup recovery in other
                # processes. The lock spans the exclusive link and index commit.
                db.execute('BEGIN IMMEDIATE')
                # Hard-link publication is atomic and exclusive: it cannot replace
                # an existing path if a UUID collision or local race occurs.
                os.link(temporary,destination,follow_symlinks=False)
                published=True
                temporary.unlink()
                temporary=None
                db.execute('INSERT INTO workbook_snapshots(snapshot_id,occurred_at,workbook_path,snapshot_path,sha256,size) VALUES (?,?,?,?,?,?)',
                    (snapshot_id,at,str(workbook),str(destination),copied_hash,copied_size))
                history_version_id=self._record_in_transaction(db,event,{'workbook':workbook.name,'action':'snapshot_created',
                    'snapshot_id':snapshot_id,'snapshot_sha256':copied_hash,'snapshot_size':copied_size,
                    'snapshot_name':workbook.name},occurred_at=_utc())
            return {'snapshot_id':snapshot_id,'history_version_id':history_version_id,
                'workbook_path':str(workbook),'snapshot_path':str(destination),
                'sha256':copied_hash,'size':copied_size,'occurred_at':at}
        except Exception:
            with self._connect() as db:db.execute('DELETE FROM workbook_snapshots WHERE snapshot_id=?',(snapshot_id,))
            if published:
                try:destination.unlink()
                except FileNotFoundError:pass
            raise
        finally:
            if temporary is not None:
                try:temporary.unlink()
                except FileNotFoundError:pass

    def list_snapshots(self,*,limit=500):
        if isinstance(limit,bool) or not isinstance(limit,int) or not 1<=limit<=10000:
            raise ValueError('Snapshot limit must be a whole number from 1 to 10000.')
        self.prune()
        with self._connect() as db:
            rows=db.execute('SELECT snapshot_id,occurred_at,workbook_path,sha256,size FROM workbook_snapshots ORDER BY occurred_at DESC,snapshot_id DESC LIMIT ?',
                            (limit,)).fetchall()
        return [{'snapshot_id':row['snapshot_id'],'occurred_at':row['occurred_at'],
            'workbook_path':row['workbook_path'],'workbook_name':Path(row['workbook_path']).name,
            'sha256':row['sha256'],'size':row['size']} for row in rows]

    def get_snapshot(self,snapshot_id):
        self.prune()
        try:normalized=str(uuid.UUID(str(snapshot_id)))
        except (ValueError,TypeError,AttributeError) as exc:raise ValueError('Invalid workbook snapshot ID.') from exc
        with self._connect() as db:
            row=db.execute('SELECT * FROM workbook_snapshots WHERE snapshot_id=?',(normalized,)).fetchone()
        if row is None:raise ValueError('Workbook snapshot was not found or has expired.')
        path=self._owned_snapshot_path(row['snapshot_path'])
        if not path.is_file() or path.stat().st_size!=row['size'] or _sha256(path)!=row['sha256']:
            raise ValueError('Workbook snapshot is missing or failed SHA-256 verification.')
        return {'snapshot_id':row['snapshot_id'],'occurred_at':row['occurred_at'],
            'workbook_path':row['workbook_path'],'snapshot_path':str(path),
            'sha256':row['sha256'],'size':row['size']}

    def restore_snapshot(self,snapshot_id,destination,*,overwrite=False,expected_current_sha256=None):
        """Restore as a new file by default; overwrite requires reviewed hash and leaves a verified backup."""
        snapshot=self.get_snapshot(snapshot_id)
        target=Path(os.path.abspath(os.fspath(Path(destination).expanduser())))
        if contains_reparse_point(target):
            raise ValueError('Restore destination contains a symbolic link or Windows reparse point.')
        target=target.resolve()
        if target.suffix.casefold() not in {'.xlsx','.xlsm'} or target.suffix.casefold()!=Path(snapshot['workbook_path']).suffix.casefold():
            raise ValueError('Restore destination must use the same supported workbook extension as the snapshot.')
        existed=target.exists()
        if existed and not overwrite:raise FileExistsError(f'Refusing to overwrite restore destination: {target}')
        if overwrite:
            if not existed:raise ValueError('Overwrite restore requires an existing target workbook.')
            if target!=Path(snapshot['workbook_path']).resolve():
                raise ValueError('Overwrite restore is limited to the workbook path recorded with this snapshot.')
            if not isinstance(expected_current_sha256,str) or not re.fullmatch(r'[0-9a-fA-F]{64}',expected_current_sha256):
                raise ValueError('Overwrite restore requires the expected current SHA-256 hash.')
            expected_current_sha256=expected_current_sha256.casefold()
            if _sha256(target)!=expected_current_sha256:
                raise RuntimeError('Workbook changed after review; restore was stopped.')
            lock_file=target.with_name('~$'+target.name)
            if lock_file.exists():raise RuntimeError('Excel lock file is present; close the workbook before restore.')
            from .windows_file_users import ensure_no_active_file_users
            ensure_no_active_file_users(target)
        previous_snapshot=None
        if overwrite and self.snapshots_enabled:
            previous_snapshot=self.create_snapshot(target,event='workbook_snapshot_before_restore')
            if _sha256(target)!=expected_current_sha256:
                raise RuntimeError('Workbook changed while the pre-restore version was being captured.')
        target.parent.mkdir(parents=True,exist_ok=True)
        stage=None;backup_path=None
        try:
            with tempfile.NamedTemporaryFile(prefix='.pqm-restore-',suffix=target.suffix,dir=target.parent,delete=False) as stream:
                stage=Path(stream.name)
            if contains_reparse_point(target) or contains_reparse_point(stage):
                raise RuntimeError('Restore destination or staging path contains a symbolic link or Windows reparse point.')
            shutil.copyfile(snapshot['snapshot_path'],stage)
            if stage.stat().st_size!=snapshot['size'] or _sha256(stage)!=snapshot['sha256']:
                raise RuntimeError('Restore staging file failed SHA-256 verification.')
            if overwrite:
                from .backup import create_timestamped_backup
                if contains_reparse_point(target) or contains_reparse_point(stage):
                    raise RuntimeError('Restore destination or staging path changed to contain a symbolic link or Windows reparse point.')
                backup_path=create_timestamped_backup(target)
                if _sha256(backup_path)!=expected_current_sha256:
                    raise RuntimeError('Current workbook backup failed verification; restore was stopped.')
                if _sha256(target)!=expected_current_sha256:
                    raise RuntimeError('Workbook changed after review; restore was stopped.')
                from .windows_file_users import ensure_no_active_file_users
                ensure_no_active_file_users(target)
                if _sha256(target)!=expected_current_sha256:
                    raise RuntimeError('Workbook changed during final restore checks; restore was stopped.')
                if contains_reparse_point(target) or contains_reparse_point(stage):
                    raise RuntimeError('Restore destination or staging path changed during final checks; restore was stopped.')
                # Match import/refresh commit semantics: never create a new
                # selected-workbook path if it disappears at this boundary.
                # ReplaceFileW also preserves target metadata and retains a
                # rollback name until the published snapshot hash is verified.
                from .excel_transaction import _replace_existing_file
                _replace_existing_file(stage, target, expected_current_sha256)
                stage=None
            else:
                if contains_reparse_point(target) or contains_reparse_point(stage):
                    raise RuntimeError('Restore destination or staging path changed before publication.')
                os.link(stage,target)
                stage.unlink();stage=None
            if _sha256(target)!=snapshot['sha256']:
                if backup_path:
                    with tempfile.NamedTemporaryFile(prefix='.pqm-rollback-',suffix=target.suffix,
                            dir=target.parent,delete=False) as stream:rollback=Path(stream.name)
                    try:
                        shutil.copyfile(backup_path,rollback)
                        if _sha256(rollback)!=expected_current_sha256:
                            raise RuntimeError('Verified rollback copy failed hash validation.')
                        from .excel_transaction import _replace_existing_file
                        current_target_hash = _sha256(target)
                        _replace_existing_file(rollback,target,current_target_hash)
                        if _sha256(target)!=expected_current_sha256:
                            raise RuntimeError('Restore failed and the original workbook could not be verified after rollback.')
                    except Exception as rollback_exc:
                        raise RuntimeError(
                            f'Restore verification failed and automatic rollback was not confirmed; '
                            f'the verified pre-restore backup remains at {backup_path}.') from rollback_exc
                    finally:
                        try:rollback.unlink()
                        except FileNotFoundError:pass
                raise RuntimeError('Restored workbook failed verification; the verified pre-restore backup was retained.')
            try:
                event_id=self.record('workbook_version_restored',{'workbook':target.name,'action':'snapshot_restored',
                    'snapshot_id':snapshot['snapshot_id'],'snapshot_sha256':snapshot['sha256'],
                    'snapshot_size':snapshot['size'],'snapshot_name':target.name,
                    **({'previous_snapshot_id':previous_snapshot['snapshot_id']} if previous_snapshot else {})})
                history_warning=None
            except Exception as exc:
                event_id=None;history_warning=safe_error(exc)
            return {'snapshot_id':snapshot['snapshot_id'],'restored_path':str(target),
                'sha256':snapshot['sha256'],'backup_path':str(backup_path) if backup_path else None,
                'history_version_id':event_id,'history_warning':history_warning}
        finally:
            if stage is not None:
                try:stage.unlink()
                except FileNotFoundError:pass

    def set_retention_days(self, days):
        days = _retention(days)
        with self._connect() as db:
            db.execute(
                'INSERT INTO settings(key, value) VALUES (?, ?) '
                'ON CONFLICT(key) DO UPDATE SET value=excluded.value',
                ('retention_days', str(days)),
            )
        self.prune()
        return days

    def record(self, event, details, *, occurred_at=None):
        with self._connect() as db:
            return self._record_in_transaction(db,event,details,occurred_at=occurred_at)

    def _record_in_transaction(self, db, event, details, *, occurred_at=None):
        """Insert one redacted audit row using the caller's existing SQLite transaction."""
        if not isinstance(event, str) or not event.strip() or len(event) > 128:
            raise ValueError('History event must be non-empty text of at most 128 characters.')
        if not isinstance(details, dict):
            raise ValueError('History details must be an object.')
        cleaned = _safe_details(details)
        encoded = json.dumps(cleaned, ensure_ascii=False, sort_keys=True, separators=(',', ':'))
        if len(encoded.encode('utf-8')) > 64 * 1024:
            raise ValueError('History details exceed the 64 KiB limit.')
        at = _utc(occurred_at)
        version_id = str(uuid.uuid4())
        row = db.execute("SELECT value FROM settings WHERE key='retention_days'").fetchone()
        days = _retention(int(row['value'])) if row else _DEFAULT_RETENTION_DAYS
        cutoff = (datetime.now(timezone.utc) - timedelta(days=days)).isoformat(timespec='microseconds')
        db.execute('DELETE FROM history WHERE occurred_at < ?', (cutoff,))
        db.execute(
            'INSERT INTO history(version_id, occurred_at, event, details_json) VALUES (?, ?, ?, ?)',
            (version_id, at.isoformat(timespec='microseconds'), safe_error(event.strip()), encoded),
        )
        return version_id

    def list(self, *, limit=500):
        if isinstance(limit, bool) or not isinstance(limit, int) or not 1 <= limit <= 10000:
            raise ValueError('History limit must be a whole number from 1 to 10000.')
        self.prune()
        with self._connect() as db:
            rows = db.execute(
                'SELECT version_id, occurred_at, event, details_json FROM history '
                'ORDER BY occurred_at DESC, version_id DESC LIMIT ?', (limit,)
            ).fetchall()
        return [
            {'version_id': row['version_id'], 'occurred_at': row['occurred_at'],
             'event': row['event'], 'details': json.loads(row['details_json'])}
            for row in rows
        ]

    def prune(self, *, now=None):
        at = _utc(now)
        cutoff = (at - timedelta(days=self.retention_days)).isoformat(timespec='microseconds')
        with self._connect() as db:
            cursor = db.execute('DELETE FROM history WHERE occurred_at < ?', (cutoff,))
            history_removed=cursor.rowcount
            expired=db.execute('SELECT snapshot_id,snapshot_path FROM workbook_snapshots WHERE occurred_at < ?',
                               (cutoff,)).fetchall()
            db.execute('DELETE FROM workbook_snapshots WHERE occurred_at < ?',(cutoff,))
        for row in expired:
            try:self._owned_snapshot_path(row['snapshot_path']).unlink()
            except FileNotFoundError:pass
            except (OSError,ValueError):pass
        return history_removed+len(expired)

    def delete(self, version_id):
        try:
            normalized = str(uuid.UUID(str(version_id)))
        except (ValueError, TypeError, AttributeError):
            return False
        with self._connect() as db:
            row=db.execute('SELECT event,details_json FROM history WHERE version_id=?',(normalized,)).fetchone()
            cursor = db.execute('DELETE FROM history WHERE version_id=?', (normalized,))
        if cursor.rowcount==1 and row:
            try:snapshot_id=json.loads(row['details_json']).get('snapshot_id')
            except Exception:snapshot_id=None
            try:action=json.loads(row['details_json']).get('action')
            except Exception:action=None
            if snapshot_id and action=='snapshot_created':self._delete_snapshot(snapshot_id)
        return cursor.rowcount == 1

    def _delete_snapshot(self,snapshot_id):
        with self._connect() as db:
            row=db.execute('SELECT snapshot_path FROM workbook_snapshots WHERE snapshot_id=?',(snapshot_id,)).fetchone()
            db.execute('DELETE FROM workbook_snapshots WHERE snapshot_id=?',(snapshot_id,))
        if row:
            try:self._owned_snapshot_path(row['snapshot_path']).unlink()
            except (FileNotFoundError,OSError,ValueError):pass

    def clear(self):
        with self._connect() as db:
            cursor = db.execute('DELETE FROM history')
            snapshots=db.execute('SELECT snapshot_path FROM workbook_snapshots').fetchall()
            db.execute('DELETE FROM workbook_snapshots')
        for row in snapshots:
            try:self._owned_snapshot_path(row['snapshot_path']).unlink()
            except (FileNotFoundError,OSError,ValueError):pass
        return cursor.rowcount
