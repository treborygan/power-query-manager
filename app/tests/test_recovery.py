import hashlib
import os
from pathlib import Path
import shutil
import subprocess

import pytest

from pq_installer.core.recovery import find_staged_workbooks, save_verified_recovery_copy


def test_find_staged_workbooks_only_returns_matching_sibling_candidates(tmp_path):
    workbook=tmp_path/'report.xlsx';workbook.write_bytes(b'original')
    stage=tmp_path/'.report.pqstage-0123456789abcdef.xlsx';stage.write_bytes(b'staged')
    unrelated=tmp_path/'.other.pqstage-0123456789abcdef.xlsx';unrelated.write_bytes(b'other')
    assert find_staged_workbooks(workbook)==(stage,)


def test_recovery_copy_requires_excel_verified_hash_and_new_destination(tmp_path):
    workbook=tmp_path/'report.xlsx';workbook.write_bytes(b'original')
    stage=tmp_path/'.report.pqstage-0123456789abcdef.xlsx';stage.write_bytes(b'verified staged workbook')
    output=tmp_path/'report_recovered.xlsx'
    expected=hashlib.sha256(stage.read_bytes()).hexdigest()
    result=save_verified_recovery_copy(workbook,stage,output,expected_sha256=expected)
    assert result==output and output.read_bytes()==stage.read_bytes()
    assert stage.read_bytes()==b'verified staged workbook'


def test_recovery_copy_refuses_existing_destination_and_changed_candidate(tmp_path):
    workbook=tmp_path/'report.xlsx';workbook.write_bytes(b'original')
    stage=tmp_path/'.report.pqstage-0123456789abcdef.xlsx';stage.write_bytes(b'verified stage')
    expected=hashlib.sha256(stage.read_bytes()).hexdigest()
    output=tmp_path/'recovered.xlsx';output.write_bytes(b'keep me')
    with pytest.raises(FileExistsError):
        save_verified_recovery_copy(workbook,stage,output,expected_sha256=expected)
    output.unlink()
    stage.write_bytes(b'changed after Excel verification')
    with pytest.raises(RuntimeError,match='changed after verification'):
        save_verified_recovery_copy(workbook,stage,output,expected_sha256=expected)
    assert not output.exists()


def test_recovery_copy_never_overwrites_concurrent_output(tmp_path,monkeypatch):
    import pq_installer.core.recovery as recovery
    workbook=tmp_path/'report.xlsx';workbook.write_bytes(b'original')
    stage=tmp_path/'.report.pqstage-0123456789abcdef.xlsx';stage.write_bytes(b'verified stage')
    output=tmp_path/'recovered.xlsx'
    expected=hashlib.sha256(stage.read_bytes()).hexdigest()
    def collide(source,destination):
        Path(destination).write_bytes(b'concurrent output')
        raise FileExistsError(destination)
    monkeypatch.setattr(recovery.os,'link',collide)
    with pytest.raises(FileExistsError,match='created concurrently'):
        save_verified_recovery_copy(workbook,stage,output,expected_sha256=expected)
    assert output.read_bytes()==b'concurrent output'
    assert stage.read_bytes()==b'verified stage' and workbook.read_bytes()==b'original'


def test_recovery_worker_excel_verifies_then_saves_separate_copy(tmp_path,monkeypatch):
    from pq_installer.core.manager import QueryInfo
    from pq_installer.ui import recovery_worker
    workbook=tmp_path/'report.xlsx';workbook.write_bytes(b'original')
    stage=tmp_path/'.report.pqstage-0123456789abcdef.xlsx';stage.write_bytes(b'staged bytes')
    output=tmp_path/'report_recovered.xlsx'
    inspected=[]
    class Inspector:
        def inspect(self,path):
            inspected.append(Path(path))
            return (QueryInfo('Q','let x=1 in x'),)
    monkeypatch.setattr(recovery_worker,'ExcelQueryInspector',Inspector)
    monkeypatch.setattr(recovery_worker,'record_failure',lambda *args,**kwargs:None)
    completed=[];failed=[]
    worker=recovery_worker.StagedRecoveryWorker(workbook,stage,output)
    worker.completed.connect(completed.append);worker.failed.connect(failed.append);worker.run()
    assert failed==[] and len(completed)==1
    assert inspected==[stage]
    assert completed[0]['query_count']==1
    assert output.read_bytes()==b'staged bytes'
    assert workbook.read_bytes()==b'original' and stage.read_bytes()==b'staged bytes'


def test_recovery_worker_does_not_publish_when_excel_verification_fails(tmp_path,monkeypatch):
    from pq_installer.ui import recovery_worker
    workbook=tmp_path/'report.xlsx';workbook.write_bytes(b'original')
    stage=tmp_path/'.report.pqstage-0123456789abcdef.xlsx';stage.write_bytes(b'unreadable staged bytes')
    output=tmp_path/'report_recovered.xlsx'
    class Inspector:
        def inspect(self,path):raise RuntimeError('simulated invalid workbook')
    monkeypatch.setattr(recovery_worker,'ExcelQueryInspector',Inspector)
    monkeypatch.setattr(recovery_worker,'record_failure',lambda *args,**kwargs:None)
    completed=[];failed=[]
    worker=recovery_worker.StagedRecoveryWorker(workbook,stage,output)
    worker.completed.connect(completed.append);worker.failed.connect(failed.append);worker.run()
    assert completed==[] and failed
    assert not output.exists()
    assert workbook.read_bytes()==b'original' and stage.read_bytes()==b'unreadable staged bytes'

def test_recovery_ignores_symlinked_stage_outside_workbook_directory(tmp_path):
    workbook = tmp_path / 'report.xlsx'
    workbook.write_bytes(b'original')
    unrelated = tmp_path.parent / f'{tmp_path.name}-unrelated.xlsx'
    unrelated.write_bytes(b'unrelated external workbook')
    stage = tmp_path / '.report.pqstage-attacker.xlsx'
    try:
        stage.symlink_to(unrelated)
    except (OSError, NotImplementedError) as exc:
        pytest.skip(f'file symlinks are unavailable in this environment: {exc}')

    assert find_staged_workbooks(workbook) == ()
    with pytest.raises(ValueError, match='not a matching staged workbook'):
        save_verified_recovery_copy(
            workbook, stage, tmp_path / 'recovered.xlsx',
            expected_sha256=hashlib.sha256(unrelated.read_bytes()).hexdigest(),
        )
    assert unrelated.read_bytes() == b'unrelated external workbook'
    assert not (tmp_path / 'recovered.xlsx').exists()


def test_recovery_copy_rejects_junctioned_output_parent(tmp_path):
    if os.name != 'nt':
        pytest.skip('Windows junction acceptance requires Windows.')
    powershell = shutil.which('powershell.exe')
    if not powershell:
        pytest.skip('Windows PowerShell is required to create a disposable junction.')
    workbook=tmp_path/'report.xlsx';workbook.write_bytes(b'original')
    stage=tmp_path/'.report.pqstage-verified.xlsx';stage.write_bytes(b'verified stage')
    external=tmp_path/'redirected-output';external.mkdir()
    junction=tmp_path/'recovery-output'
    env=os.environ.copy()
    env['PQM_TEST_JUNCTION_LINK']=str(junction)
    env['PQM_TEST_JUNCTION_TARGET']=str(external)
    create=subprocess.run(
        [powershell,'-NoProfile','-NonInteractive','-Command',
         'New-Item -ItemType Junction -Path $env:PQM_TEST_JUNCTION_LINK -Target $env:PQM_TEST_JUNCTION_TARGET | Out-Null'],
        env=env,capture_output=True,text=True,check=False)
    if create.returncode:
        pytest.skip(f'Unable to create disposable junction: {create.stderr.strip()}')
    try:
        output=junction/'recovered.xlsx'
        with pytest.raises(ValueError,match='reparse point'):
            save_verified_recovery_copy(
                workbook,stage,output,expected_sha256=hashlib.sha256(stage.read_bytes()).hexdigest())
        assert not (external/'recovered.xlsx').exists()
        assert workbook.read_bytes()==b'original' and stage.read_bytes()==b'verified stage'
    finally:
        os.rmdir(junction)


def test_recovery_copy_rechecks_paths_before_exclusive_publish(tmp_path, monkeypatch):
    import pq_installer.core.recovery as recovery

    workbook=tmp_path/'report.xlsx';workbook.write_bytes(b'original')
    stage=tmp_path/'.report.pqstage-verified.xlsx';stage.write_bytes(b'verified stage')
    output=tmp_path/'recovered.xlsx'
    real_check=recovery.contains_reparse_point
    output_checks=0
    def changes_to_reparse_point(path):
        nonlocal output_checks
        if Path(path)==output:
            output_checks+=1
            return output_checks>=2
        return real_check(path)
    monkeypatch.setattr(recovery,'contains_reparse_point',changes_to_reparse_point)
    with pytest.raises(ValueError,match='reparse point'):
        recovery.save_verified_recovery_copy(
            workbook,stage,output,expected_sha256=hashlib.sha256(stage.read_bytes()).hexdigest())
    assert workbook.read_bytes()==b'original' and stage.read_bytes()==b'verified stage'
    assert not output.exists()
    assert not list(tmp_path.glob('.recovered.xlsx.recover-*'))

def test_stage_diagnosis_reports_openxml_structure_and_age(tmp_path):
    from zipfile import ZipFile
    from pq_installer.core.recovery import diagnose_staged_workbook

    workbook = tmp_path / 'report.xlsx'
    workbook.write_bytes(b'original')
    stage = tmp_path / '.report.pqstage-diagnostic.xlsx'
    with ZipFile(stage, 'w') as archive:
        archive.writestr('[Content_Types].xml', '<Types/>')
        archive.writestr('xl/workbook.xml', '<workbook/>')
    old_time = stage.stat().st_mtime - 7200
    __import__('os').utime(stage, (old_time, old_time))

    diagnosis = diagnose_staged_workbook(workbook, stage)

    assert diagnosis['status'] == 'structure_present'
    assert diagnosis['size_bytes'] > 0
    assert diagnosis['age_seconds'] >= 7199
    assert diagnosis['modified_at']
    assert 'Excel read-only verification is still required' in diagnosis['message']


def test_stage_diagnosis_distinguishes_invalid_container_and_missing_parts(tmp_path):
    from zipfile import ZipFile
    from pq_installer.core.recovery import diagnose_staged_workbook

    workbook = tmp_path / 'report.xlsx'
    workbook.write_bytes(b'original')
    invalid = tmp_path / '.report.pqstage-invalid.xlsx'
    invalid.write_bytes(b'partial file')
    missing = tmp_path / '.report.pqstage-missing.xlsx'
    with ZipFile(missing, 'w') as archive:
        archive.writestr('notes.txt', 'not an Excel workbook')

    assert diagnose_staged_workbook(workbook, invalid)['status'] == 'unreadable_container'
    diagnosis = diagnose_staged_workbook(workbook, missing)
    assert diagnosis['status'] == 'missing_workbook_parts'
    assert '[Content_Types].xml' in diagnosis['message']
