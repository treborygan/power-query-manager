import hashlib
from pathlib import Path
from types import SimpleNamespace

import pytest

from pq_installer.core.deployment import build_multi_workbook_dry_run
from pq_installer.core.manager import QueryInfo


def profile_with_sources(root):
    queries=[]
    for index,name in enumerate(('New','Changed','Same','Unknown','Disabled')):
        path=root/f'{name}.pq';path.write_text(f'let Value = {index} in Value',encoding='utf-8')
        queries.append({'id':name.casefold(),'source_path':path.name,'query_name':name,
                        'enabled':name!='Disabled','order':index,'load':{'mode':'definition_only'}})
    return {'schema_version':1,'profile_name':'batch','workbook':{'path':'unused.xlsx'},
      'options':{'existing_query_policy':'replace','backup':True,'dry_run':True,'refresh':'none'},'queries':queries}


def test_multi_workbook_dry_run_plans_are_inventory_aware_and_read_only(tmp_path):
    source_dir=tmp_path/'sources';source_dir.mkdir();profile=profile_with_sources(source_dir)
    targets=[]
    for name in ('one.xlsx','two.xlsm'):
        path=tmp_path/name;path.write_bytes(b'disposable workbook bytes');targets.append(path)
    original_hashes=[hashlib.sha256(path.read_bytes()).hexdigest() for path in targets]
    def inventory_provider(path):
        if path==targets[0]:
            return (QueryInfo('Changed','old formula'),QueryInfo('Same','let Value = 2 in Value'),QueryInfo('Unknown',None))
        return ()
    plans=build_multi_workbook_dry_run(profile,source_dir/'profile.json',targets,
                                      inventory_provider=inventory_provider)
    assert len(plans)==2 and plans[0]['workbook_path']==str(targets[0].resolve())
    assert all(plan['dry_run'] is True and plan.get('refresh','none')=='none' for plan in plans)
    actions={item['query_name']:item['action'] for item in plans[0]['items']}
    assert actions=={'New':'create','Changed':'replace','Same':'skip','Unknown':'inspect'}
    assert len(plans[1]['items'])==4 and all(item['action']=='create' for item in plans[1]['items'])
    assert all('Disabled' not in [item['query_name'] for item in plan['items']] for plan in plans)
    assert [hashlib.sha256(path.read_bytes()).hexdigest() for path in targets]==original_hashes


def test_multi_workbook_deploy_ignores_missing_disabled_source(tmp_path):
    from pq_installer.core.deployment import deploy_multi_workbook

    source_dir=tmp_path/'sources';source_dir.mkdir();profile=profile_with_sources(source_dir)
    (source_dir/'Disabled.pq').unlink()
    target=tmp_path/'target.xlsx';target.write_bytes(b'disposable workbook bytes')

    reports=deploy_multi_workbook(profile,source_dir/'profile.json',[target],
        inventory_provider=lambda path:())

    assert [report.status for report in reports]==['planned']
    assert [item['query_name'] for item in reports[0].plan['items']]==[
        'New','Changed','Same','Unknown']


def test_multi_workbook_dry_run_rejects_duplicate_targets_and_detects_inspector_mutation(tmp_path):
    sources=tmp_path/'sources';sources.mkdir();profile=profile_with_sources(sources)
    target=tmp_path/'book.xlsx';target.write_bytes(b'original')
    with pytest.raises(ValueError,match='duplicate workbook'):
        build_multi_workbook_dry_run(profile,sources/'p.json',[target,target],inventory_provider=lambda p:())
    def mutating_inspector(path):
        path.write_bytes(b'changed')
        return ()
    with pytest.raises(RuntimeError,match='changed during dry-run'):
        build_multi_workbook_dry_run(profile,sources/'p.json',[target],inventory_provider=mutating_inspector)


def test_generated_multi_target_plans_validate_against_frozen_contract_when_available(tmp_path):
    jsonschema=pytest.importorskip('jsonschema')
    source_dir=tmp_path/'sources';source_dir.mkdir();profile=profile_with_sources(source_dir)
    target=tmp_path/'book.xlsx';target.write_bytes(b'disposable')
    plans=build_multi_workbook_dry_run(profile,source_dir/'profile.json',[target],inventory_provider=lambda path:())
    contract=Path(__file__).resolve().parents[2]/'contracts'/'plan.schema.json'
    schema=__import__('json').loads(contract.read_text(encoding='utf-8'))
    jsonschema.Draft202012Validator(schema).validate(plans[0])


def test_multi_workbook_deploy_isolates_failure_and_keeps_successful_target(tmp_path):
    from pq_installer.core.deployment import deploy_multi_workbook
    from pq_installer.core.models import QueryStatus
    source=tmp_path/'q.pq';source.write_text('let x=1 in x')
    profile={'schema_version':1,'profile_name':'Batch','workbook':{'path':'unused.xlsx'},
      'options':{'existing_query_policy':'replace','backup':True,'dry_run':False,'refresh':'none'},
      'queries':[{'id':'q-id','source_path':'q.pq','query_name':'Q','enabled':True,'order':0,'load':{'mode':'definition_only'}}]}
    profile_path=tmp_path/'profile.json'
    targets=[tmp_path/'good.xlsx',tmp_path/'bad.xlsx']
    for target in targets:target.write_bytes(b'disposable')
    calls=[];events=[]
    class Installer:
        def __init__(self,path):self.path=path
        def install(self,workbook,sources,mode,**kwargs):
            calls.append(workbook)
            events.append(('apply',workbook))
            assert kwargs['expected_workbook_sha256']
            assert sources[0].source_sha256
            if workbook==targets[1]:raise RuntimeError('target rejected safely')
            return [SimpleNamespace(query_name='Q',status=QueryStatus.CREATED,message='created')]
    def inventory(path):events.append(('preflight',path));return ()
    reports=deploy_multi_workbook(profile,profile_path,targets,
      inventory_provider=inventory,installer_factory=lambda path:Installer(path))
    assert [row.status for row in reports]==['applied','failed']
    assert calls==targets
    assert [phase for phase,_ in events]==['preflight','preflight','apply','apply']
    assert all(Path(row.backup_path).is_file() for row in reports)
    assert [row.workbook_path for row in reports]==[str(path.resolve()) for path in targets]
    assert 'still matches the pre-deployment' in reports[1].message


def test_multi_workbook_deploy_reports_target_hash_after_installer_raises(tmp_path):
    from pq_installer.core.deployment import deploy_multi_workbook
    from pq_installer.core.models import QueryStatus

    source=tmp_path/'q.pq';source.write_text('let x=1 in x')
    profile={'schema_version':1,'profile_name':'Batch','workbook':{'path':'unused.xlsx'},
      'options':{'existing_query_policy':'replace','backup':True,'dry_run':False,'refresh':'none'},
      'queries':[{'id':'q-id','source_path':'q.pq','query_name':'Q','enabled':True,'order':0,'load':{'mode':'definition_only'}}]}
    targets=[tmp_path/'changed.xlsx',tmp_path/'next.xlsx']
    originals={target:b'original '+target.name.encode() for target in targets}
    for target,content in originals.items():target.write_bytes(content)
    backups=[]
    def make_backup(target):
        backup=target.with_suffix(target.suffix+'.verified-backup')
        backup.write_bytes(target.read_bytes())
        backups.append(backup)
        return backup
    class Installer:
        def __init__(self,path):self.path=path
        def install(self,workbook,*args,**kwargs):
            if workbook==targets[0]:
                workbook.write_bytes(b'changed before exception')
                raise RuntimeError('verification failed after publication')
            return [SimpleNamespace(query_name='Q',status=QueryStatus.CREATED,message='created')]

    reports=deploy_multi_workbook(profile,tmp_path/'profile.json',targets,
      inventory_provider=lambda _path:(),installer_factory=lambda path:Installer(path),
      backup_factory=make_backup)

    assert [report.status for report in reports]==['failed','applied']
    failed=reports[0]
    assert Path(failed.backup_path).read_bytes()==originals[targets[0]]
    assert 'workbook state changed' in failed.message.casefold()
    assert hashlib.sha256(targets[0].read_bytes()).hexdigest() in failed.message
    assert 'verified pre-deployment backup' in failed.message.casefold()
    assert targets[1].read_bytes()==originals[targets[1]]


def test_multi_workbook_deploy_reports_missing_target_after_installer_failure(tmp_path):
    from pq_installer.core.deployment import deploy_multi_workbook

    source=tmp_path/'q.pq';source.write_text('1')
    target=tmp_path/'vanished.xlsx';target.write_bytes(b'disposable original workbook')
    profile={'schema_version':1,'profile_name':'Batch','workbook':{'path':'unused.xlsx'},
      'options':{'existing_query_policy':'replace','backup':True,'dry_run':False,'refresh':'none'},
      'queries':[{'id':'q','source_path':'q.pq','query_name':'Q','enabled':True,'order':0,'load':{'mode':'definition_only'}}]}
    backup=target.with_suffix('.verified-backup')
    def make_backup(path):
        backup.write_bytes(path.read_bytes())
        return backup
    class Installer:
        def install(self,workbook,*args,**kwargs):
            workbook.unlink()
            raise RuntimeError('workbook disappeared during final verification')

    reports=deploy_multi_workbook(profile,tmp_path/'profile.json',[target],
      inventory_provider=lambda _path:(),installer_factory=lambda _path:Installer(),
      backup_factory=make_backup)

    assert reports[0].status=='failed'
    assert not target.exists()
    assert backup.read_bytes()==b'disposable original workbook'
    assert 'final workbook state could not be verified' in reports[0].message.casefold()
    assert 'verified pre-deployment backup' in reports[0].message.casefold()


@pytest.mark.parametrize('statuses', [
    ('Rolled back', 'Not attempted'),
    ('Created', 'Rolled back'),
    ('Created', 'Not attempted'),
])
def test_multi_workbook_deploy_never_reports_noncommitted_results_as_applied(tmp_path,statuses):
    from pq_installer.core.deployment import deploy_multi_workbook
    source=tmp_path/'q.pq';source.write_text('let x=1 in x')
    source2=tmp_path/'q2.pq';source2.write_text('let x=2 in x')
    profile={'schema_version':1,'profile_name':'Batch','workbook':{'path':'unused.xlsx'},
      'options':{'existing_query_policy':'replace','backup':True,'dry_run':False,'refresh':'none'},
      'queries':[{'id':'q','source_path':'q.pq','query_name':'Q','enabled':True,'order':0,'load':{'mode':'definition_only'}},
                 {'id':'q2','source_path':'q2.pq','query_name':'Q2','enabled':True,'order':1,'load':{'mode':'definition_only'}}]}
    target=tmp_path/'target.xlsx';target.write_bytes(b'disposable')
    class Installer:
        def install(self,*args,**kwargs):
            return [SimpleNamespace(query_name=('Q','Q2')[i],status=status,
                                    message='simulated transaction outcome')
                    for i,status in enumerate(statuses)]
    reports=deploy_multi_workbook(profile,tmp_path/'profile.json',[target],
        inventory_provider=lambda _path:(),installer_factory=lambda _path:Installer())
    assert reports[0].status=='failed'
    assert reports[0].results[-1]['status'] in {'Rolled back','Not attempted'}


@pytest.mark.parametrize('outcomes', [
    (),
    (('Q','Created'),),
    (('Q','Created'),('Q','Replaced')),
    (('Q','Created'),('Q2','Unexpected')),
])
def test_multi_workbook_deploy_requires_complete_success_results(tmp_path,outcomes):
    from pq_installer.core.deployment import deploy_multi_workbook
    source=tmp_path/'q.pq';source.write_text('1')
    source2=tmp_path/'q2.pq';source2.write_text('2')
    profile={'schema_version':1,'profile_name':'Batch','workbook':{'path':'unused.xlsx'},
      'options':{'existing_query_policy':'replace','backup':True,'dry_run':False,'refresh':'none'},
      'queries':[{'id':'q','source_path':'q.pq','query_name':'Q','enabled':True,'order':0,'load':{'mode':'definition_only'}},
                 {'id':'q2','source_path':'q2.pq','query_name':'Q2','enabled':True,'order':1,'load':{'mode':'definition_only'}}]}
    target=tmp_path/'target.xlsx';target.write_bytes(b'disposable')
    class Installer:
        def install(self,*args,**kwargs):
            return [SimpleNamespace(query_name=name,status=status,message='fake result')
                    for name,status in outcomes]
    reports=deploy_multi_workbook(profile,tmp_path/'profile.json',[target],
        inventory_provider=lambda _path:(),installer_factory=lambda _path:Installer())
    assert reports[0].status=='failed'
    assert 'did not complete successfully' in reports[0].message


def test_multi_workbook_deploy_write_preflight_blocks_readonly_target_before_backup(tmp_path):
    from pq_installer.core.deployment import deploy_multi_workbook
    from pq_installer.core.models import QueryStatus
    source=tmp_path/'q.pq';source.write_text('1')
    profile={'schema_version':1,'profile_name':'Batch','workbook':{'path':'unused.xlsx'},
      'options':{'existing_query_policy':'replace','backup':True,'dry_run':False,'refresh':'none'},
      'queries':[{'id':'q','source_path':'q.pq','query_name':'Q','enabled':True,'order':0,'load':{'mode':'definition_only'}}]}
    targets=[tmp_path/'good.xlsx',tmp_path/'readonly.xlsx']
    for target in targets:target.write_bytes(b'disposable')
    checked=[];inventoried=[];installed=[]
    def write_preflight(path):
        checked.append(path)
        if path==targets[1]:raise ValueError('The workbook is read-only or cannot be opened for writing.')
    class Installer:
        def install(self,workbook,*args,**kwargs):
            installed.append(workbook)
            return [SimpleNamespace(query_name='Q',status=QueryStatus.CREATED,message='created')]
    reports=deploy_multi_workbook(profile,tmp_path/'profile.json',targets,
      inventory_provider=lambda path:(inventoried.append(path) or ()),
      installer_factory=lambda path:Installer(),write_preflight=write_preflight)
    assert checked==targets
    assert inventoried==[targets[0].resolve()]
    assert installed==[targets[0].resolve()]
    assert [row.status for row in reports]==['applied','blocked_preflight']
    assert Path(reports[0].backup_path).is_file()
    assert reports[1].backup_path is None


def test_multi_workbook_deploy_blocks_ambiguous_preflight_and_skips_noop_targets(tmp_path):
    from pq_installer.core.deployment import deploy_multi_workbook
    source=tmp_path/'q.pq';source.write_text('1')
    profile={'schema_version':1,'profile_name':'Batch','workbook':{'path':'unused.xlsx'},
      'options':{'existing_query_policy':'replace','backup':True,'dry_run':True,'refresh':'none'},
      'queries':[{'id':'q','source_path':'q.pq','query_name':'Q','enabled':True,'order':0,'load':{'mode':'definition_only'}}]}
    targets=[tmp_path/'ambiguous.xlsx',tmp_path/'same.xlsx']
    for target in targets:target.write_bytes(b'disposable')
    installer_calls=[]
    class Installer:
        def install(self,*args,**kwargs):installer_calls.append(args[0]);return []
    def inventory(path):
        if path==targets[0].resolve():return [QueryInfo('Q',None),QueryInfo('q',None)]
        return [QueryInfo('Q','1')]
    reports=deploy_multi_workbook(profile,tmp_path/'profile.json',targets,
      inventory_provider=inventory,installer_factory=lambda path:Installer())
    assert [row.status for row in reports]==['blocked_preflight','no_changes']
    assert installer_calls==[]
    assert not list(tmp_path.glob('*.backup*'))


def test_multi_workbook_deploy_confirmation_decline_stops_all_commits(tmp_path):
    from pq_installer.core.deployment import deploy_multi_workbook
    source=tmp_path/'q.pq';source.write_text('1')
    profile={'schema_version':1,'profile_name':'Batch','workbook':{'path':'unused.xlsx'},
      'options':{'existing_query_policy':'replace','backup':True,'dry_run':False,'refresh':'none'},
      'queries':[{'id':'q','source_path':'q.pq','query_name':'Q','enabled':True,'order':0,'load':{'mode':'definition_only'}}]}
    targets=[tmp_path/'one.xlsx',tmp_path/'two.xlsx']
    for target in targets:target.write_bytes(b'disposable')
    installed=[]
    class Installer:
        def install(self,*args,**kwargs):installed.append(args[0]);return []
    confirmed=[]
    reports=deploy_multi_workbook(profile,tmp_path/'profile.json',targets,
      inventory_provider=lambda path:(),installer_factory=lambda path:Installer(),
      confirm_apply=lambda plans:(confirmed.extend(plans) or False))
    assert len(confirmed)==2
    assert [item.status for item in reports]==['cancelled','cancelled']
    assert installed==[] and not list(tmp_path.glob('*.backup*'))
