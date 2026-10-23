import json
import sqlite3
from datetime import datetime, timedelta, timezone
from pathlib import Path

import pytest

from pq_installer.core.history import HistoryStore


def test_history_records_redacted_versioned_metadata_and_omits_formula_bodies(tmp_path):
    store = HistoryStore(tmp_path / 'history.sqlite3')
    version = store.record(
        'query_imported',
        {
            'query_name': 'Sales',
            'workbook': 'report.xlsx; Password=hunter2',
            'formula': 'let Secret = "private" in Secret',
            'notes': 'let source = "private" in source',
            'nested': {'api_token': 'abc123', 'status': 'completed'},
        },
    )

    rows = store.list()
    assert len(rows) == 1
    assert rows[0]['version_id'] == version
    assert rows[0]['event'] == 'query_imported'
    serialized = json.dumps(rows[0]['details'])
    assert 'hunter2' not in serialized and '[redacted]' in serialized
    assert 'abc123' not in serialized and 'private' not in serialized
    assert rows[0]['details']['query_name'] == 'Sales'
    assert rows[0]['details']['workbook'] == 'report.xlsx; Password=[redacted]'
    assert 'notes' not in rows[0]['details'] and 'nested' not in rows[0]['details']


def test_history_retains_allowlisted_package_comparison_counts(tmp_path):
    store=HistoryStore(tmp_path/'history.sqlite3')
    store.record('query_package_update_previewed',{
        'ADDED':2,'MODIFIED':3,'UNCHANGED':4,'REMOVED':1,'CONFLICT':0,
        'formula':'let secret = 1 in secret','source_path':'C:/private/source.pq'})

    details=store.list()[0]['details']
    assert details=={'ADDED':2,'MODIFIED':3,'UNCHANGED':4,'REMOVED':1,'CONFLICT':0}


def test_history_retention_and_user_deletion(tmp_path):
    store = HistoryStore(tmp_path / 'history.sqlite3', retention_days=30)
    now = datetime(2026, 10, 21, tzinfo=timezone.utc)
    old = store.record('old_event', {}, occurred_at=now)
    recent = store.record('recent_event', {}, occurred_at=now)

    with sqlite3.connect(store.path) as db:
        db.execute('UPDATE history SET occurred_at=? WHERE version_id=?',
                   ((now - timedelta(days=31)).isoformat(), old))
    assert store.prune(now=now) == 1
    assert [row['version_id'] for row in store.list()] == [recent]
    assert store.delete(recent) is True
    assert store.list() == []
    assert store.delete(old) is False


@pytest.mark.parametrize('days', [0, -1, 3651, True, 1.5])
def test_history_rejects_invalid_retention(days, tmp_path):
    with pytest.raises(ValueError):
        HistoryStore(tmp_path / 'history.sqlite3', retention_days=days)


def test_history_clear_deletes_all_entries(tmp_path):
    store = HistoryStore(tmp_path / 'history.sqlite3')
    store.record('one', {})
    store.record('two', {})
    assert store.clear() == 2
    assert store.list() == []


def test_workbook_snapshots_are_opt_in_hash_verified_and_restorable(tmp_path):
    import hashlib
    workbook=tmp_path/'report.xlsx';original=b'workbook before import';workbook.write_bytes(original)
    store=HistoryStore(tmp_path/'history.sqlite3')
    assert store.snapshots_enabled is False
    assert store.set_snapshots_enabled(True) is True
    snapshot=store.create_snapshot(workbook,event='before_import')
    assert snapshot['sha256']==hashlib.sha256(original).hexdigest()
    assert Path(snapshot['snapshot_path']).is_file()
    assert store.get_snapshot(snapshot['snapshot_id'])['sha256']==snapshot['sha256']
    audit=next(row for row in store.list() if row['version_id']==snapshot['history_version_id'])
    assert audit['event']=='before_import'
    assert audit['details']['snapshot_id']==snapshot['snapshot_id']
    workbook.write_bytes(b'workbook after import')
    current_hash=hashlib.sha256(workbook.read_bytes()).hexdigest()
    restored=store.restore_snapshot(snapshot['snapshot_id'],workbook,overwrite=True,
                                    expected_current_sha256=current_hash)
    assert workbook.read_bytes()==original
    assert hashlib.sha256(Path(restored['backup_path']).read_bytes()).hexdigest()==current_hash
    assert restored['sha256']==snapshot['sha256']
    assert any(row['event']=='workbook_version_restored' for row in store.list())


@pytest.mark.skipif(__import__('os').name != 'nt', reason='exercises native ReplaceFileW commit boundary')
def test_snapshot_overwrite_does_not_recreate_target_removed_at_commit_boundary(tmp_path, monkeypatch):
    import hashlib
    import pq_installer.core.excel_transaction as transaction
    import pq_installer.core.windows_file_users as users

    workbook = tmp_path / 'report.xlsx'
    workbook.write_bytes(b'current workbook')
    store = HistoryStore(tmp_path / 'history.sqlite3')
    store.set_snapshots_enabled(True)
    snapshot = store.create_snapshot(workbook)
    workbook.write_bytes(b'changed workbook')
    current_hash = hashlib.sha256(workbook.read_bytes()).hexdigest()
    real_replace = transaction._call_replace_file
    removed = False

    def remove_target_then_replace(target, stage, rollback):
        nonlocal removed
        if Path(target) == workbook and not removed:
            workbook.unlink()
            removed = True
        return real_replace(target, stage, rollback)

    monkeypatch.setattr(transaction, '_call_replace_file', remove_target_then_replace)
    monkeypatch.setattr(users, 'ensure_no_active_file_users', lambda _path: None)

    with pytest.raises(FileNotFoundError, match='disappeared before commit'):
        store.restore_snapshot(snapshot['snapshot_id'], workbook, overwrite=True,
                               expected_current_sha256=current_hash)

    assert removed
    assert not workbook.exists(), 'restore must not publish into a destination removed after review'
    backups = list(tmp_path.glob('report_backup_*.xlsx'))
    assert len(backups) == 1
    assert hashlib.sha256(backups[0].read_bytes()).hexdigest() == current_hash
    assert not list(tmp_path.glob('.pqm-restore-*'))


@pytest.mark.skipif(__import__('os').name != 'nt', reason='exercises native ReplaceFileW rollback boundary')
def test_snapshot_restore_verification_rollback_does_not_recreate_disappeared_target(tmp_path, monkeypatch):
    import hashlib
    import pq_installer.core.excel_transaction as transaction
    import pq_installer.core.windows_file_users as users

    workbook = tmp_path / 'report.xlsx'
    workbook.write_bytes(b'original snapshot')
    store = HistoryStore(tmp_path / 'history.sqlite3')
    store.set_snapshots_enabled(True)
    snapshot = store.create_snapshot(workbook)
    workbook.write_bytes(b'current reviewed version')
    current_hash = hashlib.sha256(workbook.read_bytes()).hexdigest()
    real_commit = transaction._replace_existing_file
    real_replace = transaction._call_replace_file
    commit_count = 0
    rollback_boundary = False

    def commit_then_damage_first_result(stage, target, expected):
        nonlocal commit_count
        result = real_commit(stage, target, expected)
        commit_count += 1
        if commit_count == 1:
            workbook.write_bytes(b'corrupted after snapshot publication')
        return result

    def remove_target_before_rollback(target, replacement, backup):
        nonlocal rollback_boundary
        if Path(replacement).name.startswith('.pqm-rollback-'):
            workbook.unlink()
            rollback_boundary = True
        return real_replace(target, replacement, backup)

    monkeypatch.setattr(transaction, '_replace_existing_file', commit_then_damage_first_result)
    monkeypatch.setattr(transaction, '_call_replace_file', remove_target_before_rollback)
    monkeypatch.setattr(users, 'ensure_no_active_file_users', lambda _path: None)

    with pytest.raises(RuntimeError, match='verified pre-restore backup remains'):
        store.restore_snapshot(snapshot['snapshot_id'], workbook, overwrite=True,
                               expected_current_sha256=current_hash)

    assert rollback_boundary
    assert not workbook.exists(), 'rollback must not publish into a destination removed after verification failure'
    backups = list(tmp_path.glob('report_backup_*.xlsx'))
    assert len(backups) == 1
    assert hashlib.sha256(backups[0].read_bytes()).hexdigest() == current_hash


def test_snapshot_restore_requires_concurrency_hash_and_refuses_overwrite_by_default(tmp_path):
    import hashlib
    workbook=tmp_path/'report.xlsx';workbook.write_bytes(b'old')
    store=HistoryStore(tmp_path/'history.sqlite3');snapshot=store.create_snapshot(workbook)
    with pytest.raises(ValueError,match='expected current'):
        store.restore_snapshot(snapshot['snapshot_id'],workbook,overwrite=True)
    with pytest.raises(FileExistsError):
        store.restore_snapshot(snapshot['snapshot_id'],workbook)
    destination=tmp_path/'restored.xlsx'
    restored=store.restore_snapshot(snapshot['snapshot_id'],destination)
    assert destination.read_bytes()==b'old' and restored['backup_path'] is None
    workbook.write_bytes(b'changed')
    with pytest.raises(RuntimeError,match='changed after review'):
        store.restore_snapshot(snapshot['snapshot_id'],workbook,overwrite=True,
            expected_current_sha256=hashlib.sha256(b'different').hexdigest())
    assert workbook.read_bytes()==b'changed'


def test_snapshot_restore_refuses_active_workbook_without_owner_lock(tmp_path,monkeypatch):
    import pq_installer.core.windows_file_users as users
    import hashlib
    workbook=tmp_path/'open.xlsx';workbook.write_bytes(b'current')
    store=HistoryStore(tmp_path/'history.sqlite3');snapshot=store.create_snapshot(workbook)
    assert not workbook.with_name('~$'+workbook.name).exists()
    monkeypatch.setattr(users,'active_file_users',lambda path:('EXCEL.EXE',))
    with pytest.raises(RuntimeError,match='currently open or in use'):
        store.restore_snapshot(snapshot['snapshot_id'],workbook,overwrite=True,
            expected_current_sha256=hashlib.sha256(b'current').hexdigest())
    assert workbook.read_bytes()==b'current'


def test_history_retention_and_deletion_remove_associated_snapshots(tmp_path):
    from datetime import datetime,timedelta,timezone
    import sqlite3
    workbook=tmp_path/'report.xlsx';workbook.write_bytes(b'contents')
    store=HistoryStore(tmp_path/'history.sqlite3',retention_days=10)
    snapshot=store.create_snapshot(workbook)
    assert Path(snapshot['snapshot_path']).exists()
    assert store.delete(snapshot['history_version_id']) is True
    assert not Path(snapshot['snapshot_path']).exists()
    snapshot=store.create_snapshot(workbook)
    assert store.clear()>0
    assert not Path(snapshot['snapshot_path']).exists()
    assert store.list_snapshots()==[]

    snapshot=store.create_snapshot(workbook)
    old=(datetime.now(timezone.utc)-timedelta(days=11)).isoformat()
    with sqlite3.connect(store.path) as db:
        db.execute('UPDATE workbook_snapshots SET occurred_at=? WHERE snapshot_id=?',(old,snapshot['snapshot_id']))
    store.prune(now=datetime.now(timezone.utc))
    assert not Path(snapshot['snapshot_path']).exists()
    assert store.list_snapshots()==[]


def test_corrupted_snapshot_is_never_restored(tmp_path):
    workbook=tmp_path/'report.xlsx';workbook.write_bytes(b'original')
    store=HistoryStore(tmp_path/'history.sqlite3');snapshot=store.create_snapshot(workbook)
    workbook.write_bytes(b'current')
    Path(snapshot['snapshot_path']).write_bytes(b'tampered')
    with pytest.raises(ValueError,match='failed SHA-256'):
        store.restore_snapshot(snapshot['snapshot_id'],workbook,overwrite=True,
            expected_current_sha256=__import__('hashlib').sha256(b'current').hexdigest())
    assert workbook.read_bytes()==b'current'


def test_history_store_closes_sqlite_handles_after_each_operation(tmp_path):
    store=HistoryStore(tmp_path/'history.sqlite3')
    store.record('comparison',{})
    moved=tmp_path/'history-moved.sqlite3'
    store.path.replace(moved)
    assert moved.exists()


def test_install_worker_captures_opted_in_pre_change_snapshot_before_excel_mutation(tmp_path,monkeypatch):
    from pq_installer.core.models import InstallResult,QuerySource,QueryStatus
    from pq_installer.ui import install_worker
    workbook=tmp_path/'report.xlsx';workbook.write_bytes(b'pre-import')
    source_path=tmp_path/'query.pq';source_path.write_text('1')
    store=HistoryStore(tmp_path/'history.sqlite3');store.set_snapshots_enabled(True)
    class Installer:
        def install(self,*args,**kwargs):
            workbook.write_bytes(b'post-import')
            return (InstallResult('Q',QueryStatus.CREATED,'created'),)
    monkeypatch.setattr(install_worker,'ExcelQueryInstaller',Installer)
    outcomes=[];failures=[]
    worker=install_worker.InstallWorker(workbook,[QuerySource(source_path,'Q')],None,False,history_store=store)
    worker.completed.connect(outcomes.append);worker.failed.connect(failures.append);worker.run()
    assert failures==[] and len(outcomes)==1
    outcome=outcomes[0]
    snapshot=store.get_snapshot(outcome['snapshot_id'])
    assert Path(snapshot['snapshot_path']).read_bytes()==b'pre-import'
    assert workbook.read_bytes()==b'post-import'


def test_install_worker_retains_verified_backup_when_excel_operation_fails(tmp_path,monkeypatch):
    from pq_installer.core.models import QuerySource
    from pq_installer.ui import install_worker
    workbook=tmp_path/'report.xlsx';workbook.write_bytes(b'pre-import bytes')
    source=tmp_path/'query.pq';source.write_text('1')
    class FailingInstaller:
        def install(self,*args,**kwargs):raise RuntimeError('simulated staged import failure')
    monkeypatch.setattr(install_worker,'ExcelQueryInstaller',FailingInstaller)
    monkeypatch.setattr(install_worker,'record_failure',lambda *args,**kwargs:None)
    failures=[]
    worker=install_worker.InstallWorker(workbook,[QuerySource(source,'Q')],None,True)
    worker.failed.connect(failures.append);worker.run()
    backups=list(tmp_path.glob('report_backup_*.xlsx'))
    assert failures
    assert len(backups)==1 and backups[0].read_bytes()==b'pre-import bytes'
    assert workbook.read_bytes()==b'pre-import bytes'


def test_install_worker_does_not_claim_workbook_saved_when_all_queries_are_skipped(tmp_path,monkeypatch):
    from pq_installer.core.manager import QueryInfo
    from pq_installer.core.models import ExistingQueryMode,InstallResult,QuerySource,QueryStatus
    from pq_installer.ui import install_worker
    workbook=tmp_path/'report.xlsx';workbook.write_bytes(b'unchanged workbook')
    source=tmp_path/'query.pq';source.write_text('replacement formula')
    store=HistoryStore(tmp_path/'history.sqlite3');store.set_snapshots_enabled(True)
    class Inspector:
        def inspect(self,_path):return (QueryInfo('Q','existing formula'),)
    class SkippingInstaller:
        def install(self,*args,**kwargs):
            raise AssertionError('No-op skip should not start a workbook writer.')
    monkeypatch.setattr(install_worker,'ExcelQueryInstaller',SkippingInstaller)
    monkeypatch.setattr(install_worker,'ExcelQueryInspector',Inspector,raising=False)
    logs=[];outcomes=[];failures=[]
    worker=install_worker.InstallWorker(workbook,[QuerySource(source,'Q')],ExistingQueryMode.SKIP,True,history_store=store)
    worker.log.connect(logs.append);worker.completed.connect(outcomes.append);worker.failed.connect(failures.append)

    worker.run()

    assert failures==[] and len(outcomes)==1
    assert workbook.read_bytes()==b'unchanged workbook'
    assert not list(tmp_path.glob('report_backup_*.xlsx'))
    assert store.list_snapshots()==[]
    assert any('No workbook changes were committed' in line for line in logs)
    assert not any('Workbook saved' in line for line in logs)


def test_install_worker_aborts_before_excel_when_opted_in_snapshot_fails(tmp_path,monkeypatch):
    from pq_installer.core.models import QuerySource
    from pq_installer.ui import install_worker
    workbook=tmp_path/'report.xlsx';workbook.write_bytes(b'pre-import')
    source_path=tmp_path/'query.pq';source_path.write_text('1')
    class BrokenHistory:
        snapshots_enabled=True
        def create_snapshot(self,*args,**kwargs):raise OSError('Password=private; query formula: let x=1 in x')
    called=[]
    monkeypatch.setattr(install_worker,'ExcelQueryInstaller',lambda:called.append(True))
    from pq_installer.core.error_log import record_failure
    diagnostic_root=tmp_path/'appdata'
    monkeypatch.setattr(install_worker,'record_failure',lambda operation,error:record_failure(operation,error,directory=diagnostic_root))
    failures=[]
    logs=[]
    worker=install_worker.InstallWorker(workbook,[QuerySource(source_path,'Q')],None,False,history_store=BrokenHistory())
    worker.failed.connect(failures.append);worker.log.connect(logs.append);worker.run()
    assert not called and failures and workbook.read_bytes()==b'pre-import'
    assert 'Operation could not be completed (operation_failed)' in failures[0]
    assert 'private' not in failures[0]
    diagnostic=diagnostic_root/'PowerQueryManager'/'logs'/'errors.jsonl'
    content=diagnostic.read_text(encoding='utf-8')
    assert 'history_store' not in content and 'private' not in content and 'let x' not in content
    assert any(str(diagnostic) in row for row in logs)


def test_install_worker_reports_per_query_failure_and_rollback_outcomes(tmp_path,monkeypatch):
    from pq_installer.core.models import InstallResult,QuerySource,QueryStatus
    from pq_installer.core.excel_queries import InstallOperationError
    from pq_installer.ui import install_worker
    workbook=tmp_path/'book.xlsx';workbook.write_bytes(b'disposable')
    source=tmp_path/'query.pq';source.write_text('1')
    error=InstallOperationError('Excel Add failed',(
        InstallResult('Earlier',QueryStatus.ROLLED_BACK,'staged change reverted'),
        InstallResult('Current',QueryStatus.FAILED,'failed during definition install'),
        InstallResult('Later',QueryStatus.NOT_ATTEMPTED,'operation stopped earlier')))
    class Installer:
        def install(self,*args,**kwargs):raise error
    monkeypatch.setattr(install_worker,'ExcelQueryInstaller',Installer)
    monkeypatch.setattr(install_worker,'record_failure',lambda *args,**kwargs:None)
    logs=[];failures=[]
    worker=install_worker.InstallWorker(workbook,[QuerySource(source,'Current')],None,False)
    worker.log.connect(logs.append);worker.failed.connect(failures.append);worker.run()
    assert failures and 'Per-query outcomes' in failures[0]
    assert any('Earlier: Rolled back' in line for line in logs)
    assert any('Current: Failed' in line for line in logs)
    assert any('Later: Not attempted' in line for line in logs)


def test_install_worker_accepts_cleanup_status_strings_from_installer_logger(tmp_path,monkeypatch):
    from pq_installer.core.models import InstallResult,QuerySource,QueryStatus
    from pq_installer.ui import install_worker
    workbook=tmp_path/'book.xlsx';workbook.write_bytes(b'disposable')
    source=tmp_path/'query.m';source.write_text('let Value = 1 in Value')
    outcome=InstallResult('Q',QueryStatus.CREATED,'Created query definition.')
    class Installer:
        def install(self,*args,logger,**kwargs):
            logger(outcome)
            logger('Excel Quit left the app-owned process running; the captured process was closed.')
            return [outcome]
    monkeypatch.setattr(install_worker,'ExcelQueryInstaller',Installer)
    monkeypatch.setattr(install_worker,'record_failure',lambda *args,**kwargs:None)
    logs=[];completed=[];failures=[]
    worker=install_worker.InstallWorker(workbook,[QuerySource(source,'Q')],None,False)
    worker.log.connect(logs.append);worker.completed.connect(completed.append);worker.failed.connect(failures.append)
    worker.run()
    assert failures==[]
    assert len(completed)==1
    assert any('Q: Created' in line for line in logs)
    assert any('captured process was closed' in line for line in logs)


def test_refresh_worker_captures_snapshot_before_refresh_and_links_result(tmp_path,monkeypatch):
    from pq_installer.core.excel_refresh import RefreshResult
    from pq_installer.ui import refresh_worker
    workbook=tmp_path/'report.xlsx';workbook.write_bytes(b'pre-refresh')
    store=HistoryStore(tmp_path/'history.sqlite3');store.set_snapshots_enabled(True)
    class Refresher:
        def refresh(self,path,*args,**kwargs):
            workbook.write_bytes(b'post-refresh')
            return (RefreshResult('Q','refreshed','done'),)
    monkeypatch.setattr(refresh_worker,'ExcelQueryRefresher',Refresher)
    outcomes=[];failures=[]
    worker=refresh_worker.RefreshWorker(workbook,('Q',),timeout_seconds=60,
        asynchronous=True,make_backup=False,history_store=store)
    worker.completed.connect(outcomes.append);worker.failed.connect(lambda message,results:failures.append(message))
    worker.run()
    assert failures==[] and len(outcomes)==1
    snapshot=store.get_snapshot(outcomes[0]['snapshot_id'])
    assert Path(snapshot['snapshot_path']).read_bytes()==b'pre-refresh'
    assert workbook.read_bytes()==b'post-refresh'

def test_snapshot_pruning_never_unlinks_path_outside_history_directory(tmp_path):
    import sqlite3
    from datetime import datetime, timedelta, timezone

    workbook = tmp_path / 'report.xlsx'
    workbook.write_bytes(b'workbook')
    unrelated = tmp_path / 'important.txt'
    unrelated.write_text('keep this file', encoding='utf-8')
    store = HistoryStore(tmp_path / 'history.sqlite3', retention_days=1)
    snapshot = store.create_snapshot(workbook)
    old = (datetime.now(timezone.utc) - timedelta(days=2)).isoformat()
    with sqlite3.connect(store.path) as db:
        db.execute(
            'UPDATE workbook_snapshots SET occurred_at=?,snapshot_path=? WHERE snapshot_id=?',
            (old, str(unrelated), snapshot['snapshot_id']),
        )

    store.prune(now=datetime.now(timezone.utc))

    assert unrelated.read_text(encoding='utf-8') == 'keep this file'
    assert not store.list_snapshots()

def test_expired_snapshot_cannot_be_retrieved_or_restored_directly(tmp_path):
    import hashlib
    import sqlite3
    from datetime import datetime, timedelta, timezone

    workbook = tmp_path / 'report.xlsx'
    workbook.write_bytes(b'current')
    store = HistoryStore(tmp_path / 'history.sqlite3', retention_days=1)
    snapshot = store.create_snapshot(workbook)
    old = (datetime.now(timezone.utc) - timedelta(days=2)).isoformat()
    with sqlite3.connect(store.path) as db:
        db.execute(
            'UPDATE workbook_snapshots SET occurred_at=? WHERE snapshot_id=?',
            (old, snapshot['snapshot_id']),
        )

    with pytest.raises(ValueError, match='not found or has expired'):
        store.get_snapshot(snapshot['snapshot_id'])
    with pytest.raises(ValueError, match='not found or has expired'):
        store.restore_snapshot(
            snapshot['snapshot_id'], workbook, overwrite=True,
            expected_current_sha256=hashlib.sha256(b'current').hexdigest(),
        )
    assert workbook.read_bytes() == b'current'

def test_snapshot_comparison_history_keeps_snapshot_ids_and_redacts_formula_and_secrets(tmp_path):
    import uuid

    store = HistoryStore(tmp_path / 'history.sqlite3')
    before_id, after_id = str(uuid.uuid4()), str(uuid.uuid4())
    store.record('workbook_snapshots_compared', {
        'action': 'snapshot_comparison_completed',
        'before_snapshot_id': before_id,
        'after_snapshot_id': after_id,
        'comparison': 'completed',
        'formula': 'let Secret = "private" in Secret',
        'password': 'never-store-this',
    })

    entry = store.list()[0]
    assert entry['event'] == 'workbook_snapshots_compared'
    assert entry['details'] == {
        'action': 'snapshot_comparison_completed',
        'before_snapshot_id': before_id,
        'after_snapshot_id': after_id,
        'comparison': 'completed',
        'password': '[redacted]',
    }
from pathlib import Path
import pytest
from pq_installer.core.history import HistoryStore
import pq_installer.core.history as history_module


def test_snapshot_restore_rejects_destination_reparse_point_before_resolving(tmp_path, monkeypatch):
    workbook = tmp_path / 'source.xlsx'
    workbook.write_bytes(b'saved version')
    store = HistoryStore(tmp_path / 'history.sqlite3')
    snapshot = store.create_snapshot(workbook)
    destination = tmp_path / 'linked' / 'report.xlsx'
    calls = []

    def contains_reparse_point(path):
        path = Path(path)
        calls.append(path)
        return path == destination

    monkeypatch.setattr(history_module, 'contains_reparse_point', contains_reparse_point, raising=False)
    with pytest.raises(ValueError, match='symbolic link or Windows reparse point'):
        store.restore_snapshot(snapshot['snapshot_id'], destination)
    assert destination in calls
    assert not (tmp_path / 'linked' / 'report.xlsx').exists()

def test_snapshot_restore_rechecks_target_path_immediately_before_commit(tmp_path, monkeypatch):
    import hashlib
    import pq_installer.core.history as history_module

    workbook = tmp_path / 'report.xlsx'
    workbook.write_bytes(b'current workbook')
    store = HistoryStore(tmp_path / 'history.sqlite3')
    snapshot = store.create_snapshot(workbook)
    workbook.write_bytes(b'new snapshot contents')
    observed = 0

    def contains_reparse_point(path):
        nonlocal observed
        if Path(path) == workbook:
            observed += 1
            return observed >= 4
        return False

    monkeypatch.setattr(history_module, 'contains_reparse_point', contains_reparse_point)
    monkeypatch.setattr('pq_installer.core.windows_file_users.ensure_no_active_file_users', lambda path: None)
    with pytest.raises(RuntimeError, match='changed during final checks'):
        store.restore_snapshot(snapshot['snapshot_id'], workbook, overwrite=True,
                               expected_current_sha256=hashlib.sha256(b'new snapshot contents').hexdigest())
    assert workbook.read_bytes() == b'new snapshot contents'
    assert not list(tmp_path.glob('.pqm-restore-*'))

def test_snapshot_lookup_rejects_reparse_point_snapshot_root(tmp_path, monkeypatch):
    import pq_installer.core.history as history_module

    workbook = tmp_path / 'report.xlsx'
    workbook.write_bytes(b'workbook')
    store = HistoryStore(tmp_path / 'history.sqlite3')
    snapshot = store.create_snapshot(workbook)
    root = store.snapshots_directory
    monkeypatch.setattr(history_module, 'contains_reparse_point',
                        lambda path: Path(path) == root, raising=False)
    with pytest.raises(ValueError, match='symbolic link or Windows reparse point'):
        store.get_snapshot(snapshot['snapshot_id'])

def test_snapshot_restore_rechecks_new_destination_before_publish(tmp_path, monkeypatch):
    import pq_installer.core.history as history_module

    workbook = tmp_path / 'source.xlsx'
    workbook.write_bytes(b'saved version')
    store = HistoryStore(tmp_path / 'history.sqlite3')
    snapshot = store.create_snapshot(workbook)
    destination = tmp_path / 'copy.xlsx'
    observed = 0

    def contains_reparse_point(path):
        nonlocal observed
        if Path(path) == destination:
            observed += 1
            return observed >= 3
        return False

    monkeypatch.setattr(history_module, 'contains_reparse_point', contains_reparse_point)
    with pytest.raises(RuntimeError, match='changed before publication'):
        store.restore_snapshot(snapshot['snapshot_id'], destination)
    assert not destination.exists()
    assert not list(tmp_path.glob('.pqm-restore-*'))

def test_corrupt_history_database_is_preserved_and_reinitialized(tmp_path):
    path = tmp_path / 'history.sqlite3'
    corrupt = b'not a sqlite database\x00\x01'
    path.write_bytes(corrupt)

    store = HistoryStore(path)

    archive = store.recovery_archive
    assert archive is not None and archive.is_file() and archive != path
    assert archive.read_bytes() == corrupt
    assert store.list() == []
    version_id = store.record('history_recovered', {'action': 'recovery_complete'})
    assert store.list()[0]['version_id'] == version_id
    assert path.is_file()

def test_corrupt_history_recovery_preserves_sqlite_sidecars(tmp_path):
    path = tmp_path / 'history.sqlite3'
    corrupt = b'bad db bytes'
    wal = b'preserved wal bytes'
    path.write_bytes(corrupt)
    Path(str(path) + '-wal').write_bytes(wal)

    store = HistoryStore(path)

    assert store.recovery_archive.read_bytes() == corrupt
    assert Path(str(store.recovery_archive) + '-wal').read_bytes() == wal
    assert store.list() == []


def test_history_operational_error_does_not_archive_database(tmp_path, monkeypatch):
    import sqlite3

    path = tmp_path / 'history.sqlite3'
    original = b'preserve while storage is unavailable'
    path.write_bytes(original)

    def io_failure(self):
        raise sqlite3.OperationalError('disk I/O error')

    monkeypatch.setattr(HistoryStore, '_validate_existing_database', io_failure)
    with pytest.raises(sqlite3.OperationalError, match='disk I/O error'):
        HistoryStore(path)
    assert path.read_bytes() == original
    assert not list(tmp_path.glob('history.corrupt-*'))

def test_incompatible_history_schema_is_preserved_before_reset(tmp_path):
    import sqlite3

    path = tmp_path / 'history.sqlite3'
    db = sqlite3.connect(path)
    try:
        db.execute('CREATE TABLE history (unexpected TEXT)')
        db.commit()
    finally:
        db.close()
    original = path.read_bytes()

    store = HistoryStore(path)

    assert store.recovery_archive is not None
    assert store.recovery_archive.read_bytes() == original
    assert store.list() == []

def test_stale_history_recovery_lock_refuses_to_move_database(tmp_path, monkeypatch):
    import pq_installer.core.history as history_module

    path = tmp_path / 'history.sqlite3'
    original = b'preserve corrupt bytes while recovery is locked'
    path.write_bytes(original)
    lock = Path(str(path) + '.recovery.lock')
    lock.write_text('another process', encoding='ascii')
    ticks = iter((0.0, 6.0))
    monkeypatch.setattr(history_module.time, 'monotonic', lambda: next(ticks))
    monkeypatch.setattr(history_module.time, 'sleep', lambda _seconds: None)

    with pytest.raises(RuntimeError, match='lock is stale') as caught:
        HistoryStore(path)
    assert str(lock) in str(caught.value)
    assert 'Close every Power Query Manager instance' in str(caught.value)
    assert f'remove only that recovery lock file: {lock}' in str(caught.value)
    assert 'Do not remove or move the history database or its sidecars' in str(caught.value)
    assert path.read_bytes() == original
    assert lock.read_text(encoding='ascii') == 'another process'
    assert not list(tmp_path.glob('history.corrupt-*'))

def test_failed_fresh_history_initialization_keeps_archive_sidecars_paired(tmp_path, monkeypatch):
    path = tmp_path / 'history.sqlite3'
    original = b'bad original history database'
    wal = b'original wal bytes'
    path.write_bytes(original)
    Path(str(path) + '-wal').write_bytes(wal)
    real_initialize = HistoryStore._initialize
    calls = 0

    def fail_after_initializing_fresh_store(self):
        nonlocal calls
        calls += 1
        real_initialize(self)
        if calls == 2:
            raise OSError('simulated storage failure after creating fresh store')

    monkeypatch.setattr(HistoryStore, '_initialize', fail_after_initializing_fresh_store)
    with pytest.raises(OSError, match='simulated storage failure'):
        HistoryStore(path)
    archives = list(tmp_path.glob('history.corrupt-*'))
    database_archives = [candidate for candidate in archives if candidate.suffix == '.sqlite3']
    assert len(database_archives) == 1 and database_archives[0].read_bytes() == original
    assert Path(str(database_archives[0]) + '-wal').read_bytes() == wal
    assert path.is_file()
    assert not Path(str(path) + '-wal').exists()

def test_power_loss_after_snapshot_publish_preserves_and_surfaces_unindexed_workbook(tmp_path):
    import os
    import subprocess
    import sys

    history_path = tmp_path / 'history.sqlite3'
    workbook = tmp_path / 'source.xlsx'
    workbook.write_bytes(b'disposable workbook snapshot bytes')
    script = r'''
import os
import sys
from pathlib import Path
sys.path.insert(0, sys.argv[1])
from pq_installer.core import history as module
from pq_installer.core.history import HistoryStore
history_path = Path(sys.argv[2])
workbook = Path(sys.argv[3])
store = HistoryStore(history_path)
link = module.os.link
def publish_then_power_loss(source, destination, **kwargs):
    result = link(source, destination, **kwargs)
    if Path(destination).parent == store.snapshots_directory:
        os._exit(73)
    return result
module.os.link = publish_then_power_loss
store.create_snapshot(workbook)
'''
    env = os.environ.copy()
    env['PYTHONPATH'] = str(Path(__file__).resolve().parents[1] / 'src')
    child = subprocess.run(
        [sys.executable, '-c', script, env['PYTHONPATH'], str(history_path), str(workbook)],
        env=env, capture_output=True, text=True, timeout=20,
    )
    assert child.returncode == 73, child.stderr
    orphaned = list((tmp_path / 'history_snapshots').glob('*.xlsx'))
    published = [path for path in orphaned if not path.name.startswith('.pqm-snapshot-')]
    in_progress = [path for path in orphaned if path.name.startswith('.pqm-snapshot-')]
    assert len(published) == len(in_progress) == 1
    assert published[0].read_bytes() == workbook.read_bytes()
    assert in_progress[0].read_bytes() == workbook.read_bytes()

    recovered = HistoryStore(history_path)

    assert recovered.recovered_snapshot_files
    recovered_file = Path(recovered.recovered_snapshot_files[0])
    assert recovered_file.is_file()
    assert recovered_file.read_bytes() == workbook.read_bytes()
    assert not published[0].exists()
    assert in_progress[0].is_file()
    assert in_progress[0].read_bytes() == workbook.read_bytes()
    assert recovered.list_snapshots() == []

def test_snapshot_recovery_ignores_live_temp_and_unrelated_files(tmp_path):
    history_path = tmp_path / 'history.sqlite3'
    workbook = tmp_path / 'source.xlsx'
    workbook.write_bytes(b'disposable workbook bytes')
    store = HistoryStore(history_path)
    stored = store.create_snapshot(workbook)
    live_temp = store.snapshots_directory / '.pqm-snapshot-active-copy.xlsx'
    live_temp.write_bytes(b'copy still in progress')
    unrelated = store.snapshots_directory / 'notes.xlsx'
    unrelated.write_bytes(b'user file')

    reopened = HistoryStore(history_path)

    assert reopened.recovered_snapshot_files == []
    assert Path(stored['snapshot_path']).is_file()
    assert live_temp.read_bytes() == b'copy still in progress'
    assert unrelated.read_bytes() == b'user file'

def test_snapshot_file_is_flushed_before_atomic_publication(tmp_path, monkeypatch):
    import pq_installer.core.history as history_module

    history_path = tmp_path / 'history.sqlite3'
    workbook = tmp_path / 'source.xlsx'
    workbook.write_bytes(b'disposable workbook bytes')
    store = HistoryStore(history_path)
    events = []
    real_fsync = history_module.os.fsync
    real_link = history_module.os.link

    def recording_fsync(descriptor):
        events.append('flush')
        return real_fsync(descriptor)

    def recording_link(source, destination, **kwargs):
        if Path(destination).parent == store.snapshots_directory:
            events.append('publish')
        return real_link(source, destination, **kwargs)

    monkeypatch.setattr(history_module.os, 'fsync', recording_fsync)
    monkeypatch.setattr(history_module.os, 'link', recording_link)

    created = store.create_snapshot(workbook)

    assert events.index('flush') < events.index('publish')
    assert Path(created['snapshot_path']).read_bytes() == workbook.read_bytes()

def test_snapshot_publication_refuses_existing_uuid_destination(tmp_path, monkeypatch):
    import uuid
    import pq_installer.core.history as history_module

    history_path = tmp_path / 'history.sqlite3'
    workbook = tmp_path / 'source.xlsx'
    workbook.write_bytes(b'disposable workbook bytes')
    store = HistoryStore(history_path)
    fixed_id = uuid.UUID('aaaaaaaa-aaaa-4aaa-8aaa-aaaaaaaaaaaa')
    destination = store.snapshots_directory / f'{fixed_id}.xlsx'
    destination.parent.mkdir(parents=True, exist_ok=True)
    destination.write_bytes(b'pre-existing unrelated bytes')
    monkeypatch.setattr(history_module.uuid, 'uuid4', lambda: fixed_id)

    with pytest.raises(FileExistsError):
        store.create_snapshot(workbook)

    assert destination.read_bytes() == b'pre-existing unrelated bytes'
    assert store.list_snapshots() == []
    assert list(store.snapshots_directory.glob('.pqm-snapshot-*')) == []

def test_concurrent_history_startup_does_not_recover_inflight_snapshot(tmp_path, monkeypatch):
    import threading
    import pq_installer.core.history as history_module

    history_path = tmp_path / 'history.sqlite3'
    workbook = tmp_path / 'source.xlsx'
    workbook.write_bytes(b'disposable workbook snapshot bytes')
    store = HistoryStore(history_path)
    published = threading.Event()
    allow_create_to_continue = threading.Event()
    recovery_finished = threading.Event()
    created = []
    recovered_stores = []
    errors = []
    real_link = history_module.os.link

    def pause_after_publication(source, destination, **kwargs):
        result = real_link(source, destination, **kwargs)
        if Path(destination).parent == store.snapshots_directory:
            published.set()
            if not allow_create_to_continue.wait(5):
                raise TimeoutError('test did not release snapshot publication barrier')
        return result

    monkeypatch.setattr(history_module.os, 'link', pause_after_publication)

    def create():
        try:
            created.append(store.create_snapshot(workbook))
        except Exception as exc:
            errors.append(exc)

    def recover():
        try:
            recovered_stores.append(HistoryStore(history_path))
        except Exception as exc:
            errors.append(exc)
        finally:
            recovery_finished.set()

    creator = threading.Thread(target=create)
    creator.start()
    assert published.wait(5)
    recovery = threading.Thread(target=recover)
    recovery.start()
    try:
        assert not recovery_finished.wait(0.25), 'startup moved the published file while its writer was active'
    finally:
        allow_create_to_continue.set()
    creator.join(10)
    recovery.join(10)

    assert not creator.is_alive() and not recovery.is_alive()
    assert errors == []
    assert len(created) == 1 and len(recovered_stores) == 1
    assert recovered_stores[0].recovered_snapshot_files == []
    assert store.get_snapshot(created[0]['snapshot_id'])['sha256'] == created[0]['sha256']

def test_history_store_refuses_database_symlink_without_touching_target(tmp_path):
    import hashlib
    import os

    target = tmp_path / 'target.sqlite3'
    HistoryStore(target)
    before = hashlib.sha256(target.read_bytes()).hexdigest()
    link = tmp_path / 'history.sqlite3'
    try:
        os.symlink(target, link)
    except OSError as exc:
        pytest.skip(f'file symlink creation is unavailable: {exc}')

    with pytest.raises(ValueError, match='symbolic link or Windows reparse point'):
        HistoryStore(link)

    assert link.is_symlink()
    assert hashlib.sha256(target.read_bytes()).hexdigest() == before

def test_history_store_refuses_symlinked_parent_before_creating_database(tmp_path):
    import os

    real_directory = tmp_path / 'real-history'
    real_directory.mkdir()
    linked_directory = tmp_path / 'linked-history'
    try:
        os.symlink(real_directory, linked_directory, target_is_directory=True)
    except OSError as exc:
        pytest.skip(f'directory symlink creation is unavailable: {exc}')

    with pytest.raises(ValueError, match='symbolic link or Windows reparse point'):
        HistoryStore(linked_directory / 'history.sqlite3')

    assert not (real_directory / 'history.sqlite3').exists()

def test_stale_partial_snapshot_temp_is_preserved_for_manual_recovery(tmp_path):
    import os
    import time

    history_path = tmp_path / 'history.sqlite3'
    store = HistoryStore(history_path)
    store.snapshots_directory.mkdir(parents=True)
    abandoned = store.snapshots_directory / '.pqm-snapshot-interrupted.xlsx'
    abandoned.write_bytes(b'partial disposable workbook copy')
    old_time = time.time() - 48 * 60 * 60
    os.utime(abandoned, (old_time, old_time))

    recovered = HistoryStore(history_path)

    assert len(recovered.recovered_snapshot_files) == 1
    recovered_file = recovered.recovered_snapshot_files[0]
    assert recovered_file.read_bytes() == b'partial disposable workbook copy'
    assert not abandoned.exists()
    assert recovered.list_snapshots() == []

def test_power_loss_between_snapshot_row_and_history_event_rolls_back_pair(tmp_path):
    import os
    import subprocess
    import sys

    history_path = tmp_path / 'history.sqlite3'
    workbook = tmp_path / 'source.xlsx'
    workbook.write_bytes(b'disposable workbook bytes for atomic audit test')
    script = r'''
import os
import sys
from pathlib import Path
sys.path.insert(0, sys.argv[1])
from pq_installer.core.history import HistoryStore
history_path = Path(sys.argv[2])
workbook = Path(sys.argv[3])
store = HistoryStore(history_path)
def stop_after_snapshot_insert(self, db, event, details, *, occurred_at=None):
    os._exit(74)
HistoryStore._record_in_transaction = stop_after_snapshot_insert
store.create_snapshot(workbook)
'''
    env = os.environ.copy()
    env['PYTHONPATH'] = str(Path(__file__).resolve().parents[1] / 'src')
    child = subprocess.run(
        [sys.executable, '-c', script, env['PYTHONPATH'], str(history_path), str(workbook)],
        env=env, capture_output=True, text=True, timeout=20,
    )
    assert child.returncode == 74, child.stderr

    recovered = HistoryStore(history_path)

    assert recovered.list_snapshots() == []
    assert recovered.recovered_snapshot_files
    assert recovered.recovered_snapshot_files[0].read_bytes() == workbook.read_bytes()
