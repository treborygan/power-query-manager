import hashlib
import json
import threading
from pathlib import Path
from types import SimpleNamespace
import pytest

from pq_installer.core.manager import QueryInfo
from pq_installer.core.history import HistoryStore
from pq_installer import cli


@pytest.fixture(autouse=True)
def isolate_cli_local_history(monkeypatch, tmp_path):
    monkeypatch.setenv('LOCALAPPDATA', str(tmp_path / 'LocalAppData'))


def test_cli_multitarget_plan_uses_inspector_worker_and_writes_no_overwrite_json(tmp_path,monkeypatch):
    source=tmp_path/'q.pq';source.write_text('let x=1 in x',encoding='utf-8')
    profile=tmp_path/'profile.json';profile.write_text(json.dumps({
      'schema_version':1,'profile_name':'test','workbook':{'path':'unused.xlsx'},
      'options':{'existing_query_policy':'replace','backup':True,'dry_run':True,'refresh':'none'},
      'queries':[{'id':'stable-q','source_path':'q.pq','query_name':'Q','enabled':True,'order':0,'load':{'mode':'definition_only'}}]}))
    targets=[tmp_path/'one.xlsx',tmp_path/'two.xlsm']
    for target in targets:target.write_bytes(b'disposable')
    before=[hashlib.sha256(target.read_bytes()).hexdigest() for target in targets]
    thread_ids=[]
    class Inspector:
        def inspect(self,path):
            thread_ids.append(threading.get_ident())
            return [QueryInfo('Q','let x=1 in x')] if Path(path)==targets[0].resolve() else []
    monkeypatch.setattr(cli,'ExcelQueryInspector',lambda:Inspector())
    output=tmp_path/'plans.json'
    result=cli.main(['plan','--dry-run','--profile',str(profile),
                     '--workbook',str(targets[0]),'--workbook',str(targets[1]),'--output',str(output)])
    assert result==0 and len(thread_ids)==2 and all(tid!=threading.get_ident() for tid in thread_ids)
    bundle=json.loads(output.read_text(encoding='utf-8'))
    assert [plan['items'][0]['action'] for plan in bundle['plans']]==['skip','create']
    assert [hashlib.sha256(target.read_bytes()).hexdigest() for target in targets]==before
    events=HistoryStore(tmp_path/'LocalAppData'/'PowerQueryManager'/'history.sqlite3').list()
    assert len(events)==1
    assert events[0]['event']=='cli_plan_created'
    assert events[0]['details']=={'action':'dry_run','added':1,'comparison':'completed',
                                  'conflict':0,'modified':0,'query_count':1,'status':'completed',
                                  'unchanged':1,'workbook_count':2}
    assert str(source) not in json.dumps(events)
    previous=output.read_bytes()
    assert cli.main(['plan','--dry-run','--profile',str(profile),'--workbook',str(targets[0]),'--output',str(output)])==2
    assert output.read_bytes()==previous


def test_cli_requires_explicit_dry_run_and_target(tmp_path):
    assert cli.main(['plan','--profile',str(tmp_path/'missing.json')])==2


def test_cli_explicit_confirmation_argument_allows_noninteractive_deployment(tmp_path,capsys):
    source=tmp_path/'q.pq';source.write_text('let x=1 in x',encoding='utf-8')
    profile=tmp_path/'profile.json';profile.write_text(json.dumps({
      'schema_version':1,'profile_name':'test','workbook':{'path':'unused.xlsx'},
      'options':{'existing_query_policy':'replace','backup':False,'dry_run':False,'refresh':'none'},
      'queries':[{'id':'q','source_path':'q.pq','query_name':'Q','enabled':True,'order':0,'load':{'mode':'definition_only'}}]}))
    target=tmp_path/'target.xlsx';target.write_bytes(b'disposable')
    calls=[]
    class Inspector:
        def inspect(self,path):return ()
    class Installer:
        def install(self,*args,**kwargs):
            calls.append('install')
            return [type('Result',(),{'query_name':'Q','status':'Created','message':'created'})()]
    result=cli.main(['deploy','--apply','--profile',str(profile),'--workbook',str(target),
        '--confirm','DEPLOY'],inspector_factory=lambda:Inspector(),installer_factory=lambda _path:Installer())
    assert result==0
    assert calls==['install']
    assert '"status": "applied"' in capsys.readouterr().out
    events=HistoryStore(tmp_path/'LocalAppData'/'PowerQueryManager'/'history.sqlite3').list()
    assert len(events)==1 and events[0]['event']=='cli_deployment_completed'
    assert events[0]['details']=={'action':'apply','created_count':1,'query_count':1,
                                  'replaced_count':0,'skipped_count':0,'status':'completed',
                                  'workbook_count':1}
    assert str(target) not in json.dumps(events)


def test_cli_confirmation_argument_does_not_approve_wrong_operation_phrase(tmp_path,capsys):
    source=tmp_path/'q.pq';source.write_text('let x=1 in x',encoding='utf-8')
    profile=tmp_path/'profile.json';profile.write_text(json.dumps({
      'schema_version':1,'profile_name':'test','workbook':{'path':'unused.xlsx'},
      'options':{'existing_query_policy':'replace','backup':False,'dry_run':False,'refresh':'none'},
      'queries':[{'id':'q','source_path':'q.pq','query_name':'Q','enabled':True,'order':0,'load':{'mode':'definition_only'}}]}))
    target=tmp_path/'target.xlsx';target.write_bytes(b'disposable')
    calls=[]
    class Inspector:
        def inspect(self,path):return ()
    class Installer:
        def install(self,*args,**kwargs):calls.append('install');return []
    result=cli.main(['deploy','--apply','--profile',str(profile),'--workbook',str(target),
        '--confirm','IMPORT'],inspector_factory=lambda:Inspector(),installer_factory=lambda _path:Installer())
    assert result==3
    assert calls==[]
    assert 'cancelled' in capsys.readouterr().out.casefold()


def test_cli_returns_validation_exit_code_for_invalid_profile(tmp_path):
    profile=tmp_path/'bad.json';profile.write_text('{',encoding='utf-8')
    target=tmp_path/'target.xlsx';target.write_bytes(b'disposable')
    assert cli.main(['plan','--dry-run','--profile',str(profile),'--workbook',str(target)])==2


def test_cli_returns_excel_exit_code_for_inspection_failure(tmp_path,monkeypatch):
    source=tmp_path/'q.pq';source.write_text('1',encoding='utf-8')
    profile=tmp_path/'profile.json';profile.write_text(json.dumps({
      'schema_version':1,'profile_name':'test','workbook':{'path':'unused.xlsx'},
      'options':{'existing_query_policy':'replace','backup':True,'dry_run':True,'refresh':'none'},
      'queries':[{'id':'stable-q','source_path':'q.pq','query_name':'Q','enabled':True,'order':0,'load':{'mode':'definition_only'}}]}))
    target=tmp_path/'target.xlsx';target.write_bytes(b'disposable')
    class Inspector:
        def inspect(self,path):raise RuntimeError('Workbook inspection failed: COM unavailable')
    monkeypatch.setattr(cli,'ExcelQueryInspector',lambda:Inspector())
    assert cli.main(['plan','--dry-run','--profile',str(profile),'--workbook',str(target)])==4


def test_cli_returns_unsupported_exit_code_for_unprobed_load_mode(tmp_path,monkeypatch):
    source=tmp_path/'q.pq';source.write_text('1',encoding='utf-8')
    profile=tmp_path/'profile.json';profile.write_text(json.dumps({
      'schema_version':1,'profile_name':'test','workbook':{'path':'unused.xlsx'},
      'options':{'existing_query_policy':'replace','backup':True,'dry_run':True,'refresh':'none'},
      'queries':[{'id':'stable-q','source_path':'q.pq','query_name':'Q','enabled':True,'order':0,'load':{'mode':'connection_only'}}]}))
    target=tmp_path/'target.xlsx';target.write_bytes(b'disposable')
    class Capabilities:
        supported_load_modes=('definition_only',)
    monkeypatch.setattr(cli,'probe_excel_capabilities',lambda:Capabilities())
    assert cli.main(['plan','--dry-run','--profile',str(profile),'--workbook',str(target)])==3


def test_cli_environment_mapping_requires_explicit_review_and_uses_temporary_copy(tmp_path,capsys):
    source=tmp_path/'q.pq';source.write_text('File.Contents("C:\\old\\q.csv")')
    other=tmp_path/'other.pq';other.write_text('let x = 2 in x')
    profile=tmp_path/'profile.json';profile.write_text(json.dumps({
      'schema_version':1,'profile_name':'test','workbook':{'path':'unused.xlsx'},
      'options':{'existing_query_policy':'replace','backup':True,'dry_run':True,'refresh':'none'},
      'environment_mappings':[{'key':'disk','from':'C:\\old','to':'D:\\new','enabled':True}],
      'queries':[{'id':'q','source_path':'q.pq','query_name':'Q','enabled':True,'order':0,'load':{'mode':'definition_only'}},
        {'id':'other','source_path':'other.pq','query_name':'Other','enabled':True,'order':1,'load':{'mode':'definition_only'}}]}))
    target=tmp_path/'target.xlsx';target.write_bytes(b'disposable')
    original=source.read_bytes()
    assert cli.main(['plan','--dry-run','--profile',str(profile),'--workbook',str(target)])==2
    seen=[]
    class Inspector:
        def inspect(self,path):
            seen.append(source.read_bytes())
            return ()
    output=tmp_path/'plan.json'
    code=cli.main(['plan','--dry-run','--profile',str(profile),'--workbook',str(target),
        '--apply-environment-mappings','--output',str(output)],
        inspector_factory=lambda:Inspector(),input_func=lambda prompt:'APPLY M-CODE MAPPINGS')
    assert code==0
    assert source.read_bytes()==original and seen==[original]
    bundle=json.loads(output.read_text(encoding='utf-8'))
    assert bundle['environment_mapping_preview'][0]['changed'] is True
    assert [row['action'] for row in bundle['plans'][0]['items']]==['create','create']
    mapped_hash=hashlib.sha256(b'File.Contents("D:\\new\\q.csv")').hexdigest()
    assert bundle['plans'][0]['items'][0]['formula_sha256']==mapped_hash
    assert bundle['plans'][0]['items'][1]['formula_sha256']==hashlib.sha256(other.read_bytes()).hexdigest()
    assert 'File.Contents("D:\\new\\q.csv")' in capsys.readouterr().out


def test_cli_environment_mapping_decline_has_cancel_exit_and_never_starts_excel(tmp_path):
    source=tmp_path/'q.pq';source.write_text('File.Contents("C:\\old\\q.csv")')
    original=source.read_bytes()
    profile=tmp_path/'profile.json';profile.write_text(json.dumps({
      'schema_version':1,'profile_name':'test','workbook':{'path':'unused.xlsx'},
      'options':{'existing_query_policy':'replace','backup':True,'dry_run':False,'refresh':'none'},
      'environment_mappings':[{'key':'disk','from':'C:\\old','to':'D:\\new','enabled':True}],
      'queries':[{'id':'q','source_path':'q.pq','query_name':'Q','enabled':True,'order':0,'load':{'mode':'definition_only'}}]}))
    target=tmp_path/'target.xlsx';target.write_bytes(b'disposable')
    calls=[]
    assert cli.main(['plan','--dry-run','--profile',str(profile),'--workbook',str(target),
        '--apply-environment-mappings'],input_func=lambda prompt:'NO')==3
    assert cli.main(['deploy','--apply','--profile',str(profile),'--workbook',str(target),
        '--apply-environment-mappings'],input_func=lambda prompt:'NO',
        inspector_factory=lambda:(calls.append('inspector') or None),
        installer_factory=lambda *_:calls.append('installer'))==3
    assert source.read_bytes()==original and target.read_bytes()==b'disposable' and calls==[]


def test_cli_eof_during_mapping_approval_is_a_safe_cancel(tmp_path, capsys):
    source=tmp_path/'q.pq';source.write_text('File.Contents("C:\\old\\q.csv")')
    original=source.read_bytes()
    profile=tmp_path/'profile.json';profile.write_text(json.dumps({
      'schema_version':1,'profile_name':'EOF approval','workbook':{'path':'unused.xlsx'},
      'options':{'existing_query_policy':'replace','backup':True,'dry_run':True,'refresh':'none'},
      'environment_mappings':[{'key':'disk','from':'C:\\old','to':'D:\\new','enabled':True}],
      'queries':[{'id':'q','source_path':'q.pq','query_name':'Q','enabled':True,'order':0,'load':{'mode':'definition_only'}}]}))
    target=tmp_path/'target.xlsx';target.write_bytes(b'disposable')
    calls=[]

    def closed_input(_prompt):
        raise EOFError

    result=cli.main(['plan','--dry-run','--profile',str(profile),'--workbook',str(target),
        '--apply-environment-mappings'],input_func=closed_input,
        inspector_factory=lambda:(calls.append('inspector') or None))

    assert result==3
    assert 'declined' in capsys.readouterr().err.casefold()
    assert source.read_bytes()==original and target.read_bytes()==b'disposable'
    assert calls==[]


def test_cli_combines_path_and_m_code_mappings_in_approved_source_order(tmp_path):
    original_dir=tmp_path/'sources';original_dir.mkdir()
    mapped_dir=tmp_path/'mapped';mapped_dir.mkdir()
    original=original_dir/'q.pq';original.write_text('File.Contents("C:\\old\\original.csv")')
    mapped_source=mapped_dir/'q.pq';mapped_source.write_text('File.Contents("C:\\old\\mapped.csv")')
    original_bytes=original.read_bytes();mapped_bytes=mapped_source.read_bytes()
    profile=tmp_path/'profile.json';profile.write_text(json.dumps({
      'schema_version':1,'profile_name':'combined mappings','workbook':{'path':'unused.xlsx'},
      'options':{'existing_query_policy':'replace','backup':True,'dry_run':True,'refresh':'none'},
      'path_mappings':[{'key':'folder','from':'sources','to':'mapped','enabled':True}],
      'environment_mappings':[{'key':'disk','from':'C:\\old','to':'D:\\new','enabled':True}],
      'queries':[{'id':'q','source_path':'sources/q.pq','query_name':'Q','enabled':True,'order':0,'load':{'mode':'definition_only'}}]}))
    target=tmp_path/'target.xlsx';target.write_bytes(b'disposable')
    output=tmp_path/'plan.json'
    code=cli.main(['plan','--dry-run','--profile',str(profile),'--workbook',str(target),
      '--apply-path-mappings','--apply-environment-mappings','--output',str(output)],
      inspector_factory=lambda:SimpleNamespace(inspect=lambda _path:()),
      input_func=lambda prompt:('APPLY PATH MAPPINGS' if 'PATH MAPPINGS' in prompt else 'APPLY M-CODE MAPPINGS'))
    assert code==0
    bundle=json.loads(output.read_text(encoding='utf-8'))
    assert bundle['path_mapping_preview'][0]['mapped_path']==str(mapped_source.resolve())
    assert bundle['environment_mapping_preview'][0]['source_path']==str(mapped_source.resolve())
    expected=b'File.Contents("D:\\new\\mapped.csv")'
    assert bundle['plans'][0]['items'][0]['formula_sha256']==hashlib.sha256(expected).hexdigest()
    assert original.read_bytes()==original_bytes and mapped_source.read_bytes()==mapped_bytes


def test_cli_unmatched_environment_mapping_still_plans_and_clears_rules(tmp_path):
    source=tmp_path/'q.pq';source.write_text('let Value = 1 in Value')
    profile=tmp_path/'profile.json';profile.write_text(json.dumps({
      'schema_version':1,'profile_name':'unmatched','workbook':{'path':'unused.xlsx'},
      'options':{'existing_query_policy':'replace','backup':True,'dry_run':True,'refresh':'none'},
      'environment_mappings':[{'key':'disk','from':'X:\\missing','to':'Y:\\mapped','enabled':True}],
      'queries':[{'id':'q','source_path':'q.pq','query_name':'Q','enabled':True,'order':0,'load':{'mode':'definition_only'}}]}))
    target=tmp_path/'target.xlsx';target.write_bytes(b'disposable')
    result=cli.main(['plan','--dry-run','--profile',str(profile),'--workbook',str(target),
        '--apply-environment-mappings'],inspector_factory=lambda:SimpleNamespace(inspect=lambda _path:()),
        input_func=lambda _prompt:'APPLY M-CODE MAPPINGS')
    assert result==0


def test_cli_path_mapping_decline_is_cancelled_before_excel_starts(tmp_path):
    source_dir=tmp_path/'sources';source_dir.mkdir()
    source=source_dir/'q.pq';source.write_text('1')
    mapped=tmp_path/'mapped';mapped.mkdir();(mapped/'q.pq').write_text('2')
    profile=tmp_path/'profile.json';profile.write_text(json.dumps({
      'schema_version':1,'profile_name':'path decline','workbook':{'path':'unused.xlsx'},
      'options':{'existing_query_policy':'replace','backup':True,'dry_run':False,'refresh':'none'},
      'path_mappings':[{'key':'folder','from':'sources','to':'mapped','enabled':True}],
      'queries':[{'id':'q','source_path':'sources/q.pq','query_name':'Q','enabled':True,'order':0,'load':{'mode':'definition_only'}}]}))
    target=tmp_path/'target.xlsx';target.write_bytes(b'disposable')
    calls=[]
    result=cli.main(['deploy','--apply','--profile',str(profile),'--workbook',str(target),
        '--apply-path-mappings'],input_func=lambda _prompt:'NO',
        inspector_factory=lambda:(calls.append('inspector') or None),
        installer_factory=lambda *_:calls.append('installer'))
    assert result==3 and calls==[]
    assert source.read_text()=='1' and (mapped/'q.pq').read_text()=='2'
    assert target.read_bytes()==b'disposable'


def test_cli_multi_workbook_deploy_requires_apply_acknowledgement(tmp_path):
    profile=tmp_path/'profile.json';profile.write_text('{}')
    target=tmp_path/'target.xlsx';target.write_bytes(b'disposable')
    assert cli.main(['deploy','--profile',str(profile),'--workbook',str(target)])==2


def test_cli_multi_workbook_deploy_reports_per_target_results_and_partial_exit(tmp_path,capsys):
    from pq_installer.core.models import QueryStatus
    source=tmp_path/'q.pq';source.write_text('1')
    profile=tmp_path/'profile.json';profile.write_text(json.dumps({
      'schema_version':1,'profile_name':'deploy','workbook':{'path':'unused.xlsx'},
      'options':{'existing_query_policy':'replace','backup':False,'dry_run':False,'refresh':'none'},
      'queries':[{'id':'q','source_path':'q.pq','query_name':'Q','enabled':True,'order':0,'load':{'mode':'definition_only'}}]}))
    targets=[tmp_path/'one.xlsx',tmp_path/'two.xlsx']
    for target in targets:target.write_bytes(b'disposable')
    class Inspector:
        def inspect(self,path):return ()
    class Installer:
        def __init__(self,path):self.path=path
        def install(self,*args,**kwargs):
            if self.path==targets[1]:raise RuntimeError('second target failed safely')
            return [SimpleNamespace(query_name='Q',status=QueryStatus.CREATED,message='created')]
    report=tmp_path/'report.json'
    code=cli.main(['deploy','--apply','--profile',str(profile),
      '--workbook',str(targets[0]),'--workbook',str(targets[1]),'--output',str(report)],
      inspector_factory=lambda:Inspector(),installer_factory=lambda path:Installer(path),
      input_func=lambda prompt:'DEPLOY')
    bundle=json.loads(report.read_text(encoding='utf-8'))
    assert code==5
    assert [row['status'] for row in bundle['results']]==['applied','failed']
    assert 'Preflight complete' in capsys.readouterr().out
    events=HistoryStore(tmp_path/'LocalAppData'/'PowerQueryManager'/'history.sqlite3').list()
    assert len(events)==1 and events[0]['event']=='cli_deployment_completed'
    assert events[0]['details']=={'action':'apply','created_count':1,'query_count':1,
        'replaced_count':0,'skipped_count':0,'status':'partial','workbook_count':2}
    audit_text=json.dumps(events)
    assert str(source) not in audit_text and str(targets[0]) not in audit_text


def test_cli_deploy_validates_report_destination_before_any_commit(tmp_path,capsys):
    source=tmp_path/'q.pq';source.write_text('1')
    profile=tmp_path/'profile.json';profile.write_text(json.dumps({
      'schema_version':1,'profile_name':'deploy','workbook':{'path':'unused.xlsx'},
      'options':{'existing_query_policy':'replace','backup':False,'dry_run':False,'refresh':'none'},
      'queries':[{'id':'q','source_path':'q.pq','query_name':'Q','enabled':True,'order':0,'load':{'mode':'definition_only'}}]}))
    target=tmp_path/'one.xlsx';target.write_bytes(b'disposable')
    report=tmp_path/'exists.json';report.write_text('preserve')
    applied=[]
    code=cli.main(['deploy','--apply','--profile',str(profile),'--workbook',str(target),'--output',str(report)],
      installer_factory=lambda path:SimpleNamespace(install=lambda *args,**kwargs:applied.append(path)),
      inspector_factory=lambda:SimpleNamespace(inspect=lambda path:()),input_func=lambda prompt:'DEPLOY')
    assert code==2 and not applied and report.read_text()=='preserve'


def test_cli_inspect_runs_read_only_inspector_on_worker_and_writes_json(tmp_path):
    target=tmp_path/'workbook.xlsx';target.write_bytes(b'disposable')
    calling_thread=threading.get_ident(); observed=[]
    class Inspector:
        def inspect(self,path):
            observed.append((Path(path),threading.get_ident()))
            return (QueryInfo('Revenue','let x = 1 in x',load_state='connection_only'),)
    output=tmp_path/'inventory.json'
    assert cli.main(['inspect','--workbook',str(target),'--output',str(output)],
                    inspector_factory=lambda:Inspector())==0
    result=json.loads(output.read_text(encoding='utf-8'))
    assert observed==[(target.resolve(),observed[0][1])]
    assert observed[0][1]!=calling_thread
    assert result['queries'][0]['load_state']=='connection_only'
    assert result['queries'][0]['formula']=='let x = 1 in x'


def test_cli_validate_checks_enabled_sources_without_excel(tmp_path,capsys):
    source=tmp_path/'q.pq';source.write_text('let x = 1 in x',encoding='utf-8')
    profile=tmp_path/'profile.json';profile.write_text(json.dumps({
      'schema_version':1,'profile_name':'validation','workbook':{'path':'unused.xlsx'},
      'options':{'existing_query_policy':'replace','backup':True,'dry_run':True,'refresh':'none'},
      'queries':[{'id':'q','source_path':'q.pq','query_name':'Q','enabled':True,'order':0,'load':{'mode':'definition_only'}}]}))
    assert cli.main(['validate','--profile',str(profile)])==0
    result=json.loads(capsys.readouterr().out)
    assert result['valid'] is True and result['queries'][0]['bytes']==source.stat().st_size


def test_cli_validate_rejects_missing_enabled_source(tmp_path,capsys):
    profile=tmp_path/'profile.json';profile.write_text(json.dumps({
      'schema_version':1,'profile_name':'validation','workbook':{'path':'unused.xlsx'},
      'options':{'existing_query_policy':'replace','backup':True,'dry_run':True,'refresh':'none'},
      'queries':[{'id':'q','source_path':'missing.pq','query_name':'Q','enabled':True,'order':0,'load':{'mode':'definition_only'}}]}))
    assert cli.main(['validate','--profile',str(profile)])==2


def test_cli_failure_records_local_diagnostic_without_persisting_error_text(monkeypatch,tmp_path,capsys):
    from pq_installer import cli
    recorded=[]
    monkeypatch.setattr(cli,'record_failure',lambda operation,error:recorded.append((operation,type(error).__name__)))
    assert cli.main(['validate','--profile',str(tmp_path/'missing.json')])==2
    assert recorded==[('validate_profile','ValueError')]
    assert 'Input needs correction (input_validation)' in capsys.readouterr().err


def test_cli_export_creates_versioned_package_from_read_only_worker_inventory(tmp_path,capsys):
    target=tmp_path/'source.xlsx';target.write_bytes(b'disposable')
    package=tmp_path/'queries.pqm.zip';threads=[]
    class Inspector:
        def inspect(self,path):
            threads.append(threading.get_ident())
            return (QueryInfo('A','let x = 1 in x'),QueryInfo('B','let x = 2 in x'))
    assert cli.main(['export','--workbook',str(target),'--output',str(package),'--package'],
                    inspector_factory=lambda:Inspector())==0
    assert threads and threads[0]!=threading.get_ident() and package.is_file()
    import zipfile
    with zipfile.ZipFile(package) as archive:
        manifest=json.loads(archive.read('pqm-package.json'))
        assert len(manifest['queries'])==2 and archive.testzip() is None
    assert json.loads(capsys.readouterr().out)['query_count']==2


def test_cli_export_refuses_nonempty_folder_before_inspection(tmp_path):
    target=tmp_path/'source.xlsx';target.write_bytes(b'disposable')
    folder=tmp_path/'out';folder.mkdir();(folder/'existing.txt').write_text('keep')
    called=[]
    assert cli.main(['export','--workbook',str(target),'--output',str(folder)],
                    inspector_factory=lambda:SimpleNamespace(inspect=lambda path:called.append(path)))==2
    assert not called and (folder/'existing.txt').read_text()=='keep'


def test_cli_refresh_requires_explicit_confirmation_and_runs_on_worker_thread(tmp_path,capsys):
    from pq_installer.core.excel_refresh import RefreshResult
    target=tmp_path/'refresh.xlsx';target.write_bytes(b'disposable')
    caller=threading.get_ident();threads=[];calls=[]
    class Inspector:
        def inspect(self,path):
            threads.append(threading.get_ident())
            return (QueryInfo('Query A','1',load_state='connection_only'),)
    class Refresher:
        def refresh(self,path,names,**kwargs):
            threads.append(threading.get_ident());calls.append((path,names,kwargs))
            return (RefreshResult(query_name='Query A',status='refreshed',message='done',duration_seconds=0.1),)
    code=cli.main(['refresh','--apply','--query','query a','--workbook',str(target),'--synchronous','--confirm','REFRESH'],
      inspector_factory=lambda:Inspector(),refresher_factory=lambda:Refresher())
    emitted=capsys.readouterr().out
    report=json.loads(emitted[emitted.find('{'):])
    assert code==0 and len(calls)==1
    assert all(thread_id!=caller for thread_id in threads)
    assert calls[0][1]==('Query A',) and calls[0][2]['asynchronous'] is False and calls[0][2]['make_backup'] is True
    assert report['results'][0]['status']=='refreshed' and report['mode']=='synchronous'


def test_cli_refresh_decline_never_calls_refresher(tmp_path):
    target=tmp_path/'refresh.xlsx';target.write_bytes(b'disposable');calls=[]
    inspector=lambda:SimpleNamespace(inspect=lambda path:(QueryInfo('Q','1'),))
    refresher=lambda:SimpleNamespace(refresh=lambda *args,**kwargs:calls.append(args))
    assert cli.main(['refresh','--apply','--all','--workbook',str(target),'--confirm','IMPORT'],inspector_factory=inspector,
      refresher_factory=refresher)==3
    assert calls==[]


def test_cli_refresh_requires_apply_and_selection(tmp_path):
    target=tmp_path/'refresh.xlsx';target.write_bytes(b'disposable')
    assert cli.main(['refresh','--workbook',str(target),'--all'])==2


def test_cli_import_package_preserves_names_order_and_requires_typed_confirmation(tmp_path,capsys):
    from pq_installer.core.excel_queries import ExistingQueryMode
    from pq_installer.core.models import InstallResult,QueryStatus
    from pq_installer.core.manager import export_query_package
    package=tmp_path/'queries.pqm.zip'
    export_query_package((QueryInfo('Second','2'),QueryInfo('First','1')),package)
    target=tmp_path/'target.xlsx';target.write_bytes(b'disposable')
    calls=[];prompts=[]
    class Installer:
        def install(self,path,sources,mode,**kwargs):
            calls.append((tuple((source.name,source.order,source.load_mode.value) for source in sources),mode))
            return tuple(InstallResult(source.name,QueryStatus.CREATED,'created') for source in sources)
    code=cli.main(['import','--apply','--package',str(package),'--workbook',str(target),'--policy','replace'],
      inspector_factory=lambda:SimpleNamespace(inspect=lambda path:()),
      installer_factory=lambda path:Installer(),input_func=lambda prompt:(prompts.append(prompt) or 'IMPORT'))
    captured=capsys.readouterr().out
    report=json.loads(captured[captured.find('{'):])
    assert code==0 and report['results'][0]['status']=='applied'
    assert calls[0][0]==(('Second',0,'definition_only'),('First',1,'definition_only'))
    assert calls[0][1] is ExistingQueryMode.REPLACE
    assert prompts==['Type IMPORT to continue: ']


def test_cli_import_package_decline_does_not_apply(tmp_path):
    from pq_installer.core.manager import export_query_package
    package=tmp_path/'queries.pqm.zip';export_query_package((QueryInfo('Q','1'),),package)
    target=tmp_path/'target.xlsx';target.write_bytes(b'disposable');calls=[]
    assert cli.main(['import','--apply','--package',str(package),'--workbook',str(target)],
      inspector_factory=lambda:SimpleNamespace(inspect=lambda path:()),
      installer_factory=lambda path:SimpleNamespace(install=lambda *a,**k:calls.append(a)),
      input_func=lambda prompt:'NO')==3
    assert calls==[]


def test_cli_import_package_rejects_unsafe_zip_without_workbook_inspection(tmp_path):
    import zipfile
    package=tmp_path/'unsafe.zip'
    with zipfile.ZipFile(package,'w') as archive:archive.writestr('../escape.pq','1')
    target=tmp_path/'target.xlsx';target.write_bytes(b'disposable');inspected=[]
    assert cli.main(['import','--apply','--package',str(package),'--workbook',str(target)],
      inspector_factory=lambda:SimpleNamespace(inspect=lambda path:inspected.append(path)),
      input_func=lambda prompt:'IMPORT')==2
    assert inspected==[] and not (tmp_path.parent/'escape.pq').exists()


def test_cli_import_package_dry_run_is_read_only_and_needs_no_confirmation(tmp_path,capsys):
    import hashlib
    from pq_installer.core.manager import export_query_package
    package=tmp_path/'queries.pqm.zip';export_query_package((QueryInfo('Q','1'),),package)
    target=tmp_path/'target.xlsx';target.write_bytes(b'disposable')
    before=hashlib.sha256(target.read_bytes()).hexdigest();applied=[]
    code=cli.main(['import','--dry-run','--package',str(package),'--workbook',str(target)],
      inspector_factory=lambda:SimpleNamespace(inspect=lambda path:()),
      installer_factory=lambda path:SimpleNamespace(install=lambda *args,**kwargs:applied.append(path)))
    report=json.loads(capsys.readouterr().out)
    assert code==0 and report['bundle_type']=='query-package-import-plan'
    assert report['plans'][0]['items'][0]['action']=='create'
    assert hashlib.sha256(target.read_bytes()).hexdigest()==before and applied==[]


def test_cli_bundle_without_noninteractive_mapping_approval_reports_required_approval(tmp_path,capsys):
    source_dir=tmp_path/'sources';source_dir.mkdir()
    (source_dir/'q.pq').write_text('1')
    mapped_dir=tmp_path/'mapped';mapped_dir.mkdir()
    (mapped_dir/'q.pq').write_text('2')
    profile=tmp_path/'profile.json';profile.write_text(json.dumps({
      'schema_version':1,'profile_name':'approval required','workbook':{'path':'unused.xlsx'},
      'options':{'existing_query_policy':'replace','backup':True,'dry_run':True,'refresh':'none'},
      'path_mappings':[{'key':'folder','from':'sources','to':'mapped','enabled':True}],
      'queries':[{'id':'q','source_path':'sources/q.pq','query_name':'Q','enabled':True,'order':0,'load':{'mode':'definition_only'}}]}))
    output=tmp_path/'bundle.pqmdeploy'
    result=cli.main(['bundle','--profile',str(profile),'--output',str(output),'--apply-path-mappings'])
    assert result==3
    assert 'approval required' in capsys.readouterr().err.casefold()
    assert not output.exists()


def test_cli_deploy_without_noninteractive_confirmation_reports_required_approval(tmp_path,capsys):
    source=tmp_path/'q.pq';source.write_text('1')
    profile=tmp_path/'profile.json';profile.write_text(json.dumps({
      'schema_version':1,'profile_name':'approval required','workbook':{'path':'unused.xlsx'},
      'options':{'existing_query_policy':'replace','backup':False,'dry_run':False,'refresh':'none'},
      'queries':[{'id':'q','source_path':'q.pq','query_name':'Q','enabled':True,'order':0,'load':{'mode':'definition_only'}}]}))
    target=tmp_path/'target.xlsx';target.write_bytes(b'disposable')
    calls=[]
    result=cli.main(['deploy','--apply','--profile',str(profile),'--workbook',str(target)],
      inspector_factory=lambda:SimpleNamespace(inspect=lambda _path:()),
      installer_factory=lambda _path:(calls.append('installer') or SimpleNamespace(install=lambda *_a,**_k:[])))
    assert result==3 and not calls
    assert 'approval required' in capsys.readouterr().err.casefold()


def test_cli_refresh_without_noninteractive_confirmation_reports_required_approval(tmp_path,capsys):
    target=tmp_path/'refresh.xlsx';target.write_bytes(b'disposable');calls=[]
    inspector=lambda:SimpleNamespace(inspect=lambda _path:(QueryInfo('Q','1',load_state='connection_only'),))
    refresher=lambda:(calls.append('factory') or SimpleNamespace(refresh=lambda *_a,**_k:()))
    result=cli.main(['refresh','--apply','--all','--workbook',str(target)],
      inspector_factory=inspector,refresher_factory=refresher)
    assert result==3 and not calls
    assert 'approval required' in capsys.readouterr().err.casefold()


def test_cli_package_import_without_noninteractive_confirmation_reports_required_approval(tmp_path,capsys):
    from pq_installer.core.manager import export_query_package
    package=tmp_path/'queries.pqm.zip';export_query_package((QueryInfo('Q','1'),),package)
    target=tmp_path/'target.xlsx';target.write_bytes(b'disposable')
    calls=[]
    result=cli.main(['import','--apply','--package',str(package),'--workbook',str(target)],
      inspector_factory=lambda:SimpleNamespace(inspect=lambda _path:()),
      installer_factory=lambda _path:(calls.append('installer') or SimpleNamespace(install=lambda *_a,**_k:())))
    assert result==3 and not calls
    assert 'approval required' in capsys.readouterr().err.casefold()


def test_cli_plan_path_only_mapping_requires_noninteractive_approval(tmp_path,capsys):
    sources=tmp_path/'sources';sources.mkdir();(sources/'q.pq').write_text('1')
    mapped=tmp_path/'mapped';mapped.mkdir();(mapped/'q.pq').write_text('2')
    profile=tmp_path/'profile.json';profile.write_text(json.dumps({
      'schema_version':1,'profile_name':'path-only plan','workbook':{'path':'unused.xlsx'},
      'options':{'existing_query_policy':'replace','backup':True,'dry_run':True,'refresh':'none'},
      'path_mappings':[{'key':'folder','from':'sources','to':'mapped','enabled':True}],
      'queries':[{'id':'q','source_path':'sources/q.pq','query_name':'Q','enabled':True,'order':0,'load':{'mode':'definition_only'}}]}))
    target=tmp_path/'target.xlsx';target.write_bytes(b'disposable');output=tmp_path/'plan.json'
    result=cli.main(['plan','--dry-run','--profile',str(profile),'--workbook',str(target),
      '--apply-path-mappings','--output',str(output)],
      inspector_factory=lambda:SimpleNamespace(inspect=lambda _path:()))
    assert result==3 and not output.exists()
    error=capsys.readouterr().err.casefold()
    assert 'approval required' in error
    assert 'explicit approval is required' in error and 'apply path mappings' in error
    approved_output=tmp_path/'approved-plan.json'
    approved=cli.main(['plan','--dry-run','--profile',str(profile),'--workbook',str(target),
      '--apply-path-mappings','--confirm','APPLY PATH MAPPINGS','--output',str(approved_output)],
      inspector_factory=lambda:SimpleNamespace(inspect=lambda _path:()))
    assert approved==0
    preview=json.loads(approved_output.read_text(encoding='utf-8'))['path_mapping_preview'][0]
    assert preview['mapped_path']==str((mapped/'q.pq').resolve())
