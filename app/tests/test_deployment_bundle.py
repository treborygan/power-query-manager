import hashlib
import json
from pathlib import Path
import zipfile

import pytest


def _profile(root):
    source=root/'query.pq';source.write_bytes(b'let Value = 7 in Value\r\n')
    return source,{
        'schema_version':1,
        'profile_name':'Release profile',
        'workbook':{'path':str(root/'default.xlsx')},
        'options':{'existing_query_policy':'replace','backup':True,'dry_run':False,'refresh':'none'},
        'queries':[
            {'id':'sales','source_path':source.name,'query_name':'Sales','enabled':True,
             'order':0,'load':{'mode':'definition_only'}},
            {'id':'disabled','source_path':'not-installed.pq','query_name':'Disabled','enabled':False,
             'order':1,'load':{'mode':'definition_only'}},
        ],
    }


def test_deployment_bundle_is_portable_and_hash_verified(tmp_path):
    from pq_installer.core.manager import export_deployment_bundle
    from pq_installer.core.packages import extract_query_zip,validate_deployment_bundle_manifest

    source,profile=_profile(tmp_path)
    profile_path=tmp_path/'profile.json'
    profile_path.write_text(json.dumps(profile),encoding='utf-8')
    package=export_deployment_bundle(profile_path,tmp_path/'release.pqmdeploy.zip')

    extracted=extract_query_zip(package,staging_parent=tmp_path)
    try:
        manifest=validate_deployment_bundle_manifest(extracted)
        bundled_profile=manifest['profile']
        bundled_source=extracted/bundled_profile['queries'][0]['source_path']
        assert bundled_source.read_bytes()==source.read_bytes()
        assert hashlib.sha256(bundled_source.read_bytes()).hexdigest()==manifest['sources'][0]['sha256']
        assert bundled_profile['queries'][1]['enabled'] is False
        assert not (extracted/bundled_profile['queries'][1]['source_path']).exists()
        assert not Path(bundled_profile['queries'][0]['source_path']).is_absolute()
    finally:
        import shutil
        shutil.rmtree(extracted)


def test_deployment_bundle_rejects_source_hash_tampering(tmp_path):
    from pq_installer.core.manager import export_deployment_bundle
    from pq_installer.core.packages import extract_query_zip

    source,profile=_profile(tmp_path)
    profile_path=tmp_path/'profile.json'
    profile_path.write_text(json.dumps(profile),encoding='utf-8')
    package=export_deployment_bundle(profile_path,tmp_path/'release.pqmdeploy.zip')
    tampered=tmp_path/'tampered.zip'
    with zipfile.ZipFile(package) as original,zipfile.ZipFile(tampered,'w') as modified:
        for item in original.infolist():
            data=original.read(item.filename)
            if item.filename.endswith('.pq'):data+=b'changed'
            modified.writestr(item.filename,data)
    with pytest.raises(ValueError,match='hash mismatch'):
        extract_query_zip(tampered,staging_parent=tmp_path)


def test_deployment_bundle_never_overwrites_existing_output(tmp_path):
    from pq_installer.core.manager import export_deployment_bundle
    source,profile=_profile(tmp_path)
    profile_path=tmp_path/'profile.json'
    profile_path.write_text(json.dumps(profile),encoding='utf-8')
    output=tmp_path/'release.pqmdeploy.zip';output.write_bytes(b'keep')
    with pytest.raises(FileExistsError):export_deployment_bundle(profile_path,output)
    assert output.read_bytes()==b'keep'


@pytest.mark.parametrize('content,match',[(b'\xff','valid UTF-8'),(b'  \r\n','empty')])
def test_deployment_bundle_rejects_unusable_enabled_source(tmp_path,content,match):
    from pq_installer.core.manager import export_deployment_bundle
    source,profile=_profile(tmp_path);source.write_bytes(content)
    profile_path=tmp_path/'profile.json';profile_path.write_text(json.dumps(profile),encoding='utf-8')
    with pytest.raises(ValueError,match=match):
        export_deployment_bundle(profile_path,tmp_path/'bad.pqmdeploy')
    assert not (tmp_path/'bad.pqmdeploy').exists()


def test_cli_creates_bundle_and_plans_from_it_without_original_source_path(tmp_path,capsys):
    from pq_installer import cli
    source,profile=_profile(tmp_path)
    profile_path=tmp_path/'profile.json'
    profile_path.write_text(json.dumps(profile),encoding='utf-8')
    package=tmp_path/'portable.pqmdeploy'
    assert cli.main(['bundle','--profile',str(profile_path),'--output',str(package)])==0
    created=json.loads(capsys.readouterr().out)
    assert created['bundle_type']=='multi-workbook-deployment'
    source.unlink()

    target=tmp_path/'target.xlsx';target.write_bytes(b'disposable workbook')
    class Inspector:
        def inspect(self,path):return ()
    assert cli.main(['plan','--dry-run','--profile',str(package),'--workbook',str(target)],
                    inspector_factory=lambda:Inspector())==0
    plan=json.loads(capsys.readouterr().out)
    assert [row['action'] for row in plan['plans'][0]['items']]==['create']
    assert plan['plans'][0]['items'][0]['query_name']=='Sales'


def test_cli_deploys_from_portable_bundle_and_cleans_extraction(tmp_path,capsys):
    from pq_installer import cli
    from pq_installer.core.models import QueryStatus
    source,profile=_profile(tmp_path)
    profile['options']['backup']=False
    profile_path=tmp_path/'profile.json'
    profile_path.write_text(json.dumps(profile),encoding='utf-8')
    package=tmp_path/'portable.pqmdeploy'
    assert cli.main(['bundle','--profile',str(profile_path),'--output',str(package)])==0
    capsys.readouterr()
    target=tmp_path/'target.xlsx';target.write_bytes(b'disposable workbook')
    calls=[]
    class Inspector:
        def inspect(self,path):return ()
    class Installer:
        def install(self,workbook,sources,*args,**kwargs):
            calls.append((Path(workbook),[(row.name,Path(row.path).read_bytes()) for row in sources]))
            return [type('Result',(),{'query_name':'Sales','status':QueryStatus.CREATED,'message':'created'})()]
    code=cli.main(['deploy','--apply','--profile',str(package),'--workbook',str(target)],
        inspector_factory=lambda:Inspector(),installer_factory=lambda path:Installer(),
        input_func=lambda _prompt:'DEPLOY')
    assert code==0
    assert calls==[(target.resolve(),[('Sales',b'let Value = 7 in Value\r\n')])]
    output=capsys.readouterr().out
    assert '"status": "applied"' in output
    extracted=list(tmp_path.glob('pq-import-*'))
    assert extracted==[]


def test_cli_bundles_only_after_approved_source_path_mapping(tmp_path,capsys):
    from pq_installer import cli
    from pq_installer.core.packages import extract_query_zip,validate_deployment_bundle_manifest
    source,profile=_profile(tmp_path)
    mapped_dir=tmp_path/'mapped';mapped_dir.mkdir()
    mapped_source=mapped_dir/'query.pq';mapped_source.write_bytes(b'let Value = 99 in Value')
    profile['path_mappings']=[{'key':'portable-root','from':'.','to':'mapped','enabled':True}]
    profile_path=tmp_path/'profile.json';profile_path.write_text(json.dumps(profile),encoding='utf-8')
    declined=tmp_path/'declined.zip'
    assert cli.main(['bundle','--profile',str(profile_path),'--output',str(declined)])==2
    assert not declined.exists()
    capsys.readouterr()

    package=tmp_path/'approved.zip'
    assert cli.main(['bundle','--profile',str(profile_path),'--output',str(package),
                     '--apply-path-mappings'],input_func=lambda _prompt:'APPLY PATH MAPPINGS')==0
    capsys.readouterr()
    root=extract_query_zip(package,staging_parent=tmp_path)
    try:
        manifest=validate_deployment_bundle_manifest(root)
        assert manifest['profile']['path_mappings']==[]
        bundled=root/manifest['profile']['queries'][0]['source_path']
        assert bundled.read_bytes()==mapped_source.read_bytes()
        assert bundled.read_bytes()!=source.read_bytes()
    finally:
        import shutil
        shutil.rmtree(root)
