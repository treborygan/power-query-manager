from pq_installer.core.redaction import safe_error
from pathlib import Path
import os
import pytest


def test_safe_error_redacts_sensitive_assignment_forms():
    message='PWD=alpha; Password="beta gamma"; token:secret-value; api_key=key-123'
    result=safe_error(message)
    for secret in ('alpha','beta gamma','secret-value','key-123'):
        assert secret not in result
    assert result.count('[redacted]') == 4


def test_safe_error_redacts_uri_passwords():
    result=safe_error('Could not connect to https://svc:top-secret@example.test/api')
    assert 'top-secret' not in result
    assert 'svc:[redacted]@example.test' in result


def test_safe_error_redacts_authorization_headers_and_signed_url_parameters():
    message=('Request failed; Authorization: Bearer bearer-sentinel-123456789; '
             '{"Authorization": "Basic basic-sentinel"}; '
             'https://example.test/data?sig=signature-sentinel&X-Amz-Signature=aws-sig-sentinel'
             '&client_secret=client-sentinel')
    result=safe_error(message)
    for secret in ('bearer-sentinel-123456789','basic-sentinel','signature-sentinel',
                   'aws-sig-sentinel','client-sentinel'):
        assert secret not in result


def test_safe_error_redacts_quoted_json_token_fields():
    message=('Response: {"refresh_token":"refresh-sentinel",'
             '"id_token":"id-sentinel","client_secret":"json-client-sentinel"}')
    result=safe_error(message)
    for secret in ('refresh-sentinel','id-sentinel','json-client-sentinel'):
        assert secret not in result


def test_safe_error_redacts_compound_secret_key_names():
    message=('Response: {"AWS_SECRET_ACCESS_KEY":"aws-sentinel",'
             '"x_api_key":"x-api-sentinel","private_key":"private-sentinel"}')
    result=safe_error(message)
    for secret in ('aws-sentinel','x-api-sentinel','private-sentinel'):
        assert secret not in result


def test_safe_error_redacts_other_authorization_schemes():
    message=('Proxy-Authorization: Negotiate negotiate-sentinel\n'
             'Authorization: Digest username="user", response="digest-sentinel"')
    result=safe_error(message)
    assert 'negotiate-sentinel' not in result
    assert 'digest-sentinel' not in result


def test_safe_error_redacts_digest_parameters_after_semicolons():
    message='Authorization: Digest username="user", response="safe-prefix; digest-tail-sentinel"'
    result=safe_error(message)
    assert 'digest-tail-sentinel' not in result


def test_local_failure_log_uses_user_directory_and_omits_raw_formula_and_secrets(tmp_path):
    from pq_installer.core.error_log import record_failure
    class SensitiveFailure(RuntimeError):pass
    failure=SensitiveFailure('Password=private; query formula: let Source = "private formula" in Source')
    path=record_failure('import',failure,directory=tmp_path)
    assert path==tmp_path/'PowerQueryManager'/'logs'/'errors.jsonl'
    content=path.read_text(encoding='utf-8')
    assert 'SensitiveFailure' in content and 'import' in content
    assert 'private' not in content and 'Source =' not in content
    assert '"category"' in content and '"timestamp_utc"' in content


def test_local_failure_log_rotates_without_failing_the_operation(tmp_path):
    from pq_installer.core.error_log import record_failure
    directory=tmp_path/'PowerQueryManager'/'logs'
    log=directory/'errors.jsonl';directory.mkdir(parents=True)
    log.write_text('x'*1_048_500,encoding='utf-8')
    assert record_failure('refresh',TimeoutError('timeout'),directory=tmp_path)==log
    assert (directory/'errors.jsonl.1').is_file()
    assert log.stat().st_size<500


def test_default_failure_log_directory_is_per_user_local_data(monkeypatch,tmp_path):
    from pq_installer.core.error_log import default_log_directory
    monkeypatch.setenv('LOCALAPPDATA',str(tmp_path/'Local AppData'))
    monkeypatch.setenv('APPDATA',str(tmp_path/'Roaming AppData'))
    assert default_log_directory()==tmp_path/'Local AppData'/'PowerQueryManager'/'logs'

def test_error_log_refuses_reparse_point_directory(tmp_path,monkeypatch):
    import pq_installer.core.error_log as error_log

    log_directory=tmp_path/'PowerQueryManager'/'logs'
    monkeypatch.setattr(error_log,'contains_reparse_point',lambda path:Path(path)==log_directory,raising=False)

    assert error_log.record_failure('import',RuntimeError('secret=hidden'),directory=tmp_path) is None
    assert not log_directory.exists()

def test_error_log_refuses_reparse_point_file_before_rotation_or_append(tmp_path,monkeypatch):
    import pq_installer.core.error_log as error_log

    log_directory=tmp_path/'PowerQueryManager'/'logs'
    log_directory.mkdir(parents=True)
    log_file=log_directory/'errors.jsonl'
    log_file.write_text('must remain untouched',encoding='utf-8')
    monkeypatch.setattr(error_log,'contains_reparse_point',lambda path:Path(path)==log_file,raising=False)

    assert error_log.record_failure('import',RuntimeError('secret=hidden'),directory=tmp_path) is None
    assert log_file.read_text(encoding='utf-8')=='must remain untouched'
    assert not log_file.with_name(log_file.name+'.1').exists()

def test_error_log_refuses_actual_directory_symlink(tmp_path):
    from pq_installer.core.error_log import record_failure

    redirected = tmp_path / 'PowerQueryManager'
    target = tmp_path / 'outside'
    target.mkdir()
    try:
        os.symlink(target, redirected, target_is_directory=True)
    except OSError as exc:
        pytest.skip(f'Host cannot create a disposable directory symlink: {type(exc).__name__}')

    assert record_failure('import',RuntimeError('Password=discard-me'),directory=tmp_path) is None
    assert not (target/'logs'/'errors.jsonl').exists()

def test_error_log_refuses_actual_file_symlink_before_append_or_rotation(tmp_path):
    from pq_installer.core.error_log import record_failure

    log_directory=tmp_path/'PowerQueryManager'/'logs'
    log_directory.mkdir(parents=True)
    redirected=tmp_path/'outside-errors.jsonl'
    original='external target must remain unchanged'
    redirected.write_text(original,encoding='utf-8')
    log_file=log_directory/'errors.jsonl'
    try:
        os.symlink(redirected,log_file)
    except OSError as exc:
        pytest.skip(f'Host cannot create a disposable file symlink: {type(exc).__name__}')

    assert record_failure('import',RuntimeError('Password=discard-me'),directory=tmp_path) is None
    assert redirected.read_text(encoding='utf-8')==original
    assert not log_file.with_name(log_file.name+'.1').exists()


def test_failure_categories_include_actionable_guidance_and_redact_details():
    from pq_installer.core.error_log import classify_failure,format_failure
    failure=RuntimeError('Workbook opened read-only; Password=private')
    category=classify_failure(failure)
    formatted=format_failure(failure)
    assert category.code=='workbook_locked_or_read_only'
    assert category.title and category.guidance
    assert category.title in formatted and category.guidance in formatted
    assert 'private' not in formatted and '[redacted]' in formatted


def test_failure_categories_distinguish_timeouts_verification_and_unsupported_modes():
    from pq_installer.core.error_log import classify_failure
    assert classify_failure(TimeoutError('open timed out')).code=='timeout'
    assert classify_failure(RuntimeError('Staged workbook verification failed')).code=='verification_failed'
    assert classify_failure(RuntimeError('Connection-only mode requires a successful capability probe')).code=='unsupported_capability'
    assert classify_failure(RuntimeError('Source changed after review')).code=='source_changed'


def test_missing_user_approval_has_its_own_actionable_category(tmp_path):
    from pq_installer.core.error_log import ApprovalRequired,classify_failure,format_failure,record_failure
    failure=ApprovalRequired('Pass --confirm DEPLOY after reviewing the plan.')
    category=classify_failure(failure)
    formatted=format_failure(failure)
    assert category.code=='approval_required'
    assert 'approval' in category.title.casefold()
    assert 'confirm' in category.guidance.casefold()
    assert category.title in formatted and category.guidance in formatted
    path=record_failure('deploy',failure,directory=tmp_path)
    assert '"category":"approval_required"' in path.read_text(encoding='utf-8')
