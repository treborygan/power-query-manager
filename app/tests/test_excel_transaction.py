import hashlib
import os
from pathlib import Path
import shutil
import subprocess

import pytest

from pq_installer.core.excel_transaction import LocalWorkbookTransactionAdapter
from pq_installer.core.excel_queries import ExcelQueryInstaller
from pq_installer.core.excel_refresh import ExcelQueryRefresher
from pq_installer.core.validation import validate_workbook_for_write


def sha256(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def test_local_transaction_adapter_prepares_and_commits_verified_sibling_stage(tmp_path):
    workbook = tmp_path / 'report.xlsx'
    workbook.write_bytes(b'original bytes')
    digest = sha256(workbook)
    adapter = LocalWorkbookTransactionAdapter(check_active=lambda _path: None)

    stage = adapter.prepare_stage(workbook, digest, operation='import')

    assert stage.parent == workbook.parent
    assert stage.name.startswith('.report.pqstage-')
    assert stage.read_bytes() == workbook.read_bytes()
    stage.write_bytes(b'committed stage bytes')
    adapter.commit_stage(stage, workbook, digest, operation='import')

    assert workbook.read_bytes() == b'committed stage bytes'
    assert not stage.exists()
    assert not list(tmp_path.glob('*.pqcommit-*.bak'))


def test_cleanup_stage_does_not_claim_success_when_final_stat_is_denied(tmp_path, monkeypatch):
    from pq_installer.core.excel_transaction import cleanup_stage_and_report

    stage = tmp_path / '.book.pqstage-retained.xlsx'
    stage.write_bytes(b'recovery candidate')

    def refuse_unlink(_path):
        raise PermissionError('sharing violation')

    def refuse_stat(_path, *args, **kwargs):
        raise PermissionError('access denied')

    monkeypatch.setattr(Path, 'unlink', refuse_unlink)
    monkeypatch.setattr(Path, 'stat', refuse_stat)
    # pathlib versions that suppress permission/stat errors can make exists()
    # report False even though the file cannot be verified absent.
    monkeypatch.setattr(Path, 'exists', lambda _path: False)
    adapter = LocalWorkbookTransactionAdapter(check_active=lambda _path: None)
    reports = []

    assert cleanup_stage_and_report(adapter, stage, logger=reports.append) is False
    assert reports == [f'Temporary staged workbook remains and may need recovery: {stage.name}']
    monkeypatch.undo()
    assert stage.read_bytes() == b'recovery candidate'


@pytest.mark.skipif(os.name != 'nt', reason='Real Windows sharing semantics require Windows.')
def test_cleanup_stage_reports_real_windows_handle_that_denies_delete_share(tmp_path):
    import win32con
    import win32file
    from pq_installer.core.excel_transaction import cleanup_stage_and_report

    stage = tmp_path / '.book.pqstage-sharing.xlsx'
    stage.write_bytes(b'disposable staged workbook')
    # Deliberately omit FILE_SHARE_DELETE. This process-owned handle makes unlink
    # fail with a real Windows sharing violation while still allowing stat/read.
    handle = win32file.CreateFile(
        str(stage), win32con.GENERIC_READ, win32con.FILE_SHARE_READ,
        None, win32con.OPEN_EXISTING, win32con.FILE_ATTRIBUTE_NORMAL, None)
    adapter = LocalWorkbookTransactionAdapter(check_active=lambda _path: None)
    reports = []
    try:
        assert cleanup_stage_and_report(adapter, stage, logger=reports.append) is False
        assert reports == [f'Temporary staged workbook remains and may need recovery: {stage.name}']
        assert stage.read_bytes() == b'disposable staged workbook'
    finally:
        win32file.CloseHandle(handle)

    assert adapter.cleanup_stage(stage) is True
    assert not stage.exists()


def test_local_transaction_adapter_discards_failed_stage_creation(tmp_path, monkeypatch):
    import pq_installer.core.excel_transaction as transaction

    workbook = tmp_path / 'report.xlsx'
    workbook.write_bytes(b'original bytes')
    digest = sha256(workbook)
    real_copy = transaction.shutil.copy2

    def change_source_after_copy(source, destination):
        result = real_copy(source, destination)
        Path(source).write_bytes(b'changed during staging')
        return result

    monkeypatch.setattr(transaction.shutil, 'copy2', change_source_after_copy)
    adapter = LocalWorkbookTransactionAdapter(check_active=lambda _path: None)

    with pytest.raises(RuntimeError, match='stage was being created'):
        adapter.prepare_stage(workbook, digest, operation='refresh')

    assert workbook.read_bytes() == b'changed during staging'
    assert not list(tmp_path.glob('*.pqstage-*'))


def test_local_transaction_adapter_commit_rechecks_hash_and_retains_stage_on_failure(tmp_path):
    workbook = tmp_path / 'report.xlsx'
    workbook.write_bytes(b'original bytes')
    adapter = LocalWorkbookTransactionAdapter(check_active=lambda _path: None)
    stage = adapter.prepare_stage(workbook, sha256(workbook), operation='import')
    stage.write_bytes(b'new stage bytes')
    workbook.write_bytes(b'changed concurrently')

    expected = hashlib.sha256(b'original bytes').hexdigest()
    with pytest.raises(RuntimeError, match='changed during import'):
        adapter.commit_stage(stage, workbook, expected, operation='import')

    assert workbook.read_bytes() == b'changed concurrently'
    assert stage.read_bytes() == b'new stage bytes'
    assert adapter.cleanup_stage(stage) is True
    assert not stage.exists()


def test_local_transaction_adapter_preserves_target_when_replace_fails(tmp_path, monkeypatch):
    import pq_installer.core.excel_transaction as transaction

    workbook = tmp_path / 'report.xlsx'
    workbook.write_bytes(b'original bytes')
    digest = sha256(workbook)
    adapter = LocalWorkbookTransactionAdapter(check_active=lambda _path: None)
    stage = adapter.prepare_stage(workbook, digest, operation='import')
    stage.write_bytes(b'new stage bytes')
    monkeypatch.setattr(transaction, '_replace_existing_file',
                        lambda *_args: (_ for _ in ()).throw(PermissionError('denied')))

    with pytest.raises(PermissionError, match='denied'):
        adapter.commit_stage(stage, workbook, digest, operation='import')

    assert workbook.read_bytes() == b'original bytes'
    assert stage.read_bytes() == b'new stage bytes'

@pytest.mark.skipif(os.name != 'nt', reason='The production Windows replacement API is under test.')
def test_commit_does_not_create_workbook_if_target_disappears_at_replace(tmp_path, monkeypatch):
    import pq_installer.core.excel_transaction as transaction

    workbook=tmp_path/'report.xlsx'
    workbook.write_bytes(b'original bytes')
    digest=sha256(workbook)
    adapter=LocalWorkbookTransactionAdapter(check_active=lambda _path:None)
    stage=adapter.prepare_stage(workbook,digest,operation='import')
    stage.write_bytes(b'new stage bytes')
    real_replace=transaction._replace_existing_file

    def remove_target_then_replace(source,destination,expected_sha256):
        Path(destination).unlink()
        return real_replace(source,destination,expected_sha256)

    monkeypatch.setattr(transaction,'_replace_existing_file',remove_target_then_replace)
    with pytest.raises(FileNotFoundError):
        adapter.commit_stage(stage,workbook,digest,operation='import')
    assert not workbook.exists()
    assert stage.read_bytes()==b'new stage bytes'

@pytest.mark.skipif(os.name != 'nt', reason='Native Windows replacement recovery is under test.')
def test_partial_replace_failure_restores_original_name_and_retains_stage(tmp_path,monkeypatch):
    import pq_installer.core.excel_transaction as transaction
    workbook=tmp_path/'report.xlsx';workbook.write_bytes(b'original bytes')
    digest=sha256(workbook)
    adapter=LocalWorkbookTransactionAdapter(check_active=lambda _path:None)
    stage=adapter.prepare_stage(workbook,digest,operation='import')
    stage.write_bytes(b'new stage bytes')

    def partial_replace(replaced,replacement,backup):
        Path(replaced).rename(backup)
        raise OSError(1177,'simulated ERROR_UNABLE_TO_MOVE_REPLACEMENT_2')

    monkeypatch.setattr(transaction,'_call_replace_file',partial_replace)
    with pytest.raises(OSError,match='1177'):
        adapter.commit_stage(stage,workbook,digest,operation='import')

    assert workbook.read_bytes()==b'original bytes'
    assert stage.read_bytes()==b'new stage bytes'
    assert not list(tmp_path.glob('.*.pqcommit-*.bak'))

@pytest.mark.skipif(os.name != 'nt', reason='Native Windows replacement recovery is under test.')
def test_failed_partial_replace_restore_keeps_named_recovery_copy(tmp_path,monkeypatch):
    import pq_installer.core.excel_transaction as transaction
    workbook=tmp_path/'report.xlsx';workbook.write_bytes(b'original bytes')
    digest=sha256(workbook)
    adapter=LocalWorkbookTransactionAdapter(check_active=lambda _path:None)
    stage=adapter.prepare_stage(workbook,digest,operation='import')
    stage.write_bytes(b'new stage bytes')

    def partial_replace(replaced,replacement,backup):
        Path(replaced).rename(backup)
        raise OSError(1177,'simulated partial replacement')

    def fail_restore(_source,_destination):
        raise PermissionError('simulated restore denial')

    monkeypatch.setattr(transaction,'_call_replace_file',partial_replace)
    monkeypatch.setattr(transaction,'_call_move_file',fail_restore)
    with pytest.raises(RuntimeError,match='recovery copy remains') as failure:
        adapter.commit_stage(stage,workbook,digest,operation='import')

    recovery=list(tmp_path.glob('.*.pqcommit-*.bak'))
    assert len(recovery)==1
    assert recovery[0].read_bytes()==b'original bytes'
    assert failure.value.args[0].endswith(recovery[0].name+'.')
    assert stage.read_bytes()==b'new stage bytes'

@pytest.mark.skipif(os.name != 'nt', reason='Windows named-stream preservation requires NTFS semantics.')
def test_commit_preserves_target_zone_stream_when_excel_stage_omits_it(tmp_path):
    workbook=tmp_path/'report.xlsx'
    workbook.write_bytes(b'original workbook')
    zone=Path(str(workbook)+':Zone.Identifier')
    try:
        zone.write_bytes(b'[ZoneTransfer]\r\nZoneId=2\r\n')
    except OSError as exc:
        pytest.skip(f'Alternate data streams unavailable on this filesystem: {exc}')
    adapter=LocalWorkbookTransactionAdapter(check_active=lambda _path:None)
    stage=adapter.prepare_stage(workbook,sha256(workbook),operation='import')
    Path(str(stage)+':Zone.Identifier').unlink(missing_ok=True)
    stage.write_bytes(b'updated workbook')

    adapter.commit_stage(stage,workbook,sha256(workbook),operation='import')

    assert workbook.read_bytes()==b'updated workbook'
    assert zone.read_bytes()==b'[ZoneTransfer]\r\nZoneId=2\r\n'
    assert not list(tmp_path.glob('*.pqcommit-*.bak'))


def test_process_crash_after_stage_write_before_commit_keeps_original_and_recovery_stage(tmp_path):
    workbook = tmp_path / 'report.xlsx'
    workbook.write_bytes(b'original workbook bytes')
    env = os.environ.copy()
    env['PYTHONPATH'] = str(Path(__file__).resolve().parents[1] / 'src')
    env['PQM_CRASH_TARGET'] = str(workbook)
    script = """
import hashlib, os
from pathlib import Path
from pq_installer.core.excel_transaction import LocalWorkbookTransactionAdapter
target = Path(os.environ['PQM_CRASH_TARGET'])
digest = hashlib.sha256(target.read_bytes()).hexdigest()
adapter = LocalWorkbookTransactionAdapter(check_active=lambda _path: None)
stage = adapter.prepare_stage(target, digest, operation='import')
stage.write_bytes(b'complete staged workbook bytes')
os._exit(71)
"""

    process = subprocess.run([os.sys.executable, '-c', script], env=env, capture_output=True)

    assert process.returncode == 71
    assert workbook.read_bytes() == b'original workbook bytes'
    stages = list(tmp_path.glob('.report.pqstage-*'))
    assert len(stages) == 1
    assert stages[0].read_bytes() == b'complete staged workbook bytes'
    stages[0].unlink()


def test_process_crash_immediately_after_atomic_replace_leaves_complete_commit(tmp_path):
    workbook = tmp_path / 'report.xlsx'
    workbook.write_bytes(b'original workbook bytes')
    env = os.environ.copy()
    env['PYTHONPATH'] = str(Path(__file__).resolve().parents[1] / 'src')
    env['PQM_CRASH_TARGET'] = str(workbook)
    script = """
import hashlib, os
from pathlib import Path
import pq_installer.core.excel_transaction as transaction
target = Path(os.environ['PQM_CRASH_TARGET'])
digest = hashlib.sha256(target.read_bytes()).hexdigest()
adapter = transaction.LocalWorkbookTransactionAdapter(check_active=lambda _path: None)
stage = adapter.prepare_stage(target, digest, operation='import')
stage.write_bytes(b'complete staged workbook bytes')
real_replace = transaction._replace_existing_file
def replace_then_crash(source, destination, expected_sha256):
    real_replace(source, destination, expected_sha256)
    os._exit(72)
transaction._replace_existing_file = replace_then_crash
adapter.commit_stage(stage, target, digest, operation='import')
"""

    process = subprocess.run([os.sys.executable, '-c', script], env=env, capture_output=True)

    assert process.returncode == 72
    assert workbook.read_bytes() == b'complete staged workbook bytes'
    assert list(tmp_path.glob('.report.pqstage-*')) == []


def test_import_and_refresh_consumers_accept_falsey_injected_transaction_adapter():
    class FalseyAdapter:
        def __bool__(self):
            return False

    adapter = FalseyAdapter()
    importer = ExcelQueryInstaller(lambda: None, workbook_transaction_adapter=adapter)
    refresher = ExcelQueryRefresher(lambda: None, workbook_transaction_adapter=adapter)

    assert importer._workbook_transactions is adapter
    assert refresher._workbook_transactions is adapter


def test_stage_preparation_rejects_symbolic_link_workbook(tmp_path):
    target=tmp_path/'real.xlsx';target.write_bytes(b'real workbook bytes')
    link=tmp_path/'selected.xlsx'
    try:
        link.symlink_to(target)
    except OSError as exc:
        pytest.skip(f'Windows symbolic-link creation unavailable: {exc}')
    adapter=LocalWorkbookTransactionAdapter(check_active=lambda _path:None)
    with pytest.raises(ValueError,match='symbolic link'):
        adapter.prepare_stage(link,sha256(target),operation='import')
    assert target.read_bytes()==b'real workbook bytes'
    assert link.is_symlink()
    assert not list(tmp_path.glob('*.pqstage-*'))


def test_commit_rejects_target_changed_to_symbolic_link(tmp_path):
    target=tmp_path/'book.xlsx';target.write_bytes(b'original bytes')
    external=tmp_path/'other.xlsx';external.write_bytes(b'original bytes')
    adapter=LocalWorkbookTransactionAdapter(check_active=lambda _path:None)
    stage=adapter.prepare_stage(target,sha256(target),operation='import')
    stage.write_bytes(b'new staged bytes')
    target.unlink()
    try:
        target.symlink_to(external)
    except OSError as exc:
        pytest.skip(f'Windows symbolic-link creation unavailable: {exc}')
    with pytest.raises(RuntimeError,match='symbolic link'):
        adapter.commit_stage(stage,target,hashlib.sha256(b'original bytes').hexdigest(),operation='import')
    assert external.read_bytes()==b'original bytes'
    assert stage.read_bytes()==b'new staged bytes'


def test_commit_rejects_windows_reparse_point_before_replace(tmp_path, monkeypatch):
    import pq_installer.core.excel_transaction as transaction

    workbook=tmp_path/'book.xlsx';workbook.write_bytes(b'original bytes')
    adapter=LocalWorkbookTransactionAdapter(check_active=lambda _path:None)
    stage=adapter.prepare_stage(workbook,sha256(workbook),operation='import')
    stage.write_bytes(b'new staged bytes')
    monkeypatch.setattr(transaction, 'contains_reparse_point', lambda path: Path(path) == workbook)
    with pytest.raises(RuntimeError, match='reparse point'):
        adapter.commit_stage(stage,workbook,hashlib.sha256(b'original bytes').hexdigest(),operation='import')
    assert workbook.read_bytes()==b'original bytes'
    assert stage.read_bytes()==b'new staged bytes'


def test_commit_rechecks_paths_after_active_user_check(tmp_path, monkeypatch):
    import pq_installer.core.excel_transaction as transaction

    workbook=tmp_path/'book.xlsx';workbook.write_bytes(b'original bytes')
    adapter=LocalWorkbookTransactionAdapter(check_active=lambda _path:None)
    stage=adapter.prepare_stage(workbook,sha256(workbook),operation='import')
    stage.write_bytes(b'new staged bytes')
    target_checks=0
    def changes_to_reparse_point(path):
        nonlocal target_checks
        if Path(path)==workbook:
            target_checks += 1
            return target_checks >= 2
        return False
    monkeypatch.setattr(transaction,'contains_reparse_point',changes_to_reparse_point)
    with pytest.raises(RuntimeError,match='reparse point'):
        adapter.commit_stage(stage,workbook,hashlib.sha256(b'original bytes').hexdigest(),operation='import')
    assert workbook.read_bytes()==b'original bytes'
    assert stage.read_bytes()==b'new staged bytes'


def test_real_windows_junction_is_rejected_at_preflight_and_stage_creation(tmp_path):
    if os.name != 'nt':
        pytest.skip('Windows junction acceptance requires Windows.')
    powershell = shutil.which('powershell.exe')
    if not powershell:
        pytest.skip('Windows PowerShell is required to create a disposable junction.')
    target_dir = tmp_path / 'junction-target'
    target_dir.mkdir()
    workbook = target_dir / 'book.xlsx'
    workbook.write_bytes(b'disposable workbook bytes')
    junction = tmp_path / 'junction-path'
    env = os.environ.copy()
    env['PQM_TEST_JUNCTION_LINK'] = str(junction)
    env['PQM_TEST_JUNCTION_TARGET'] = str(target_dir)
    create = subprocess.run(
        [powershell, '-NoProfile', '-NonInteractive', '-Command',
         'New-Item -ItemType Junction -Path $env:PQM_TEST_JUNCTION_LINK -Target $env:PQM_TEST_JUNCTION_TARGET | Out-Null'],
        env=env, capture_output=True, text=True, check=False)
    if create.returncode:
        pytest.skip(f'Unable to create disposable junction: {create.stderr.strip()}')
    try:
        selected = junction / workbook.name
        assert not junction.is_symlink()  # Windows junction, not a symbolic-link test.
        with pytest.raises(ValueError, match='reparse point'):
            validate_workbook_for_write(selected, active_file_users=lambda _path: None)
        with pytest.raises(ValueError, match='reparse point'):
            LocalWorkbookTransactionAdapter(check_active=lambda _path: None).prepare_stage(
                selected, sha256(workbook), operation='import')
        adapter = LocalWorkbookTransactionAdapter(check_active=lambda _path: None)
        stage = adapter.prepare_stage(workbook, sha256(workbook), operation='import')
        stage.write_bytes(b'staged data')
        with pytest.raises(RuntimeError, match='reparse point'):
            adapter.commit_stage(stage, selected, sha256(workbook), operation='import')
        assert workbook.read_bytes() == b'disposable workbook bytes'
        assert stage.read_bytes() == b'staged data'
        stage.unlink()
    finally:
        os.rmdir(junction)
