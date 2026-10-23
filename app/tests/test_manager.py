import json
import zipfile
from pathlib import Path
import pytest
from pq_installer.core.manager import Comparison, QueryInfo, compare_source, export_queries, load_profile, save_profile


def test_comparison_exact_unicode_and_newlines():
    assert compare_source('x', 'é\n// c', QueryInfo('x', 'é\n// c')) == Comparison.UNCHANGED
    assert compare_source('x', 'é', QueryInfo('x', 'e')) == Comparison.MODIFIED
    assert compare_source('x', 'm', None) == Comparison.NEW
    assert compare_source('x', 'm', QueryInfo('x', None)) == Comparison.UNKNOWN
    assert compare_source('x', 'm', QueryInfo('x', 'old'), skip=True) == Comparison.SKIP
    assert compare_source('x', 'm', QueryInfo('X', 'm')) == Comparison.CONFLICT


def test_exact_diff_retains_unicode_comments_and_newline_changes():
    from pq_installer.core.manager import diff_m_code
    diff=diff_m_code('let x = "café" in x\n', 'let x = "tea" in x\n// tail\n')
    assert '-let x = "tea" in x' in diff and '+let x = "café" in x' in diff
    assert '-// tail' in diff


def test_comparison_normalizes_line_endings_without_changing_diff_text():
    from pq_installer.core.manager import diff_m_code
    source='let\n    x = 1\nin\n    x\n'
    workbook='let\r\n    x = 1\r\nin\r\n    x\r\n'
    assert compare_source('Q',source,QueryInfo('Q',workbook))==Comparison.UNCHANGED
    assert 'let\n' in diff_m_code(source,workbook)


def test_comparison_ignores_terminal_line_ending_removed_by_excel():
    source='let Source = #table({"Value"}, {{5}}) in Source\r\n'
    workbook='let Source = #table({"Value"}, {{5}}) in Source'
    assert compare_source('Q',source,QueryInfo('Q',workbook))==Comparison.UNCHANGED


def test_export_names_are_safe_and_collisions_explicit(tmp_path):
    paths=export_queries([QueryInfo('../same', 'é'), QueryInfo('..\\same', 'x')],tmp_path)
    assert [p.name for p in paths] == ['_same.pq', '_same_2.pq']
    assert paths[0].read_text(encoding='utf-8') == 'é'


def test_export_truncates_max_length_query_names_for_filesystems(tmp_path):
    paths=export_queries([QueryInfo('Q' * 255, 'let x = 1 in x'),
                          QueryInfo('🧪' * 100, 'let y = 2 in y')],tmp_path)
    assert [len(path.stem) for path in paths] == [60, 60]
    assert all(len(path.stem.encode('utf-8')) <= 240 for path in paths)
    assert paths[0].read_text(encoding='utf-8') == 'let x = 1 in x'


def test_query_package_long_filename_keeps_full_query_name_in_manifest(tmp_path):
    from pq_installer.core.manager import export_query_package
    name='🧪' * 255
    package=export_query_package([QueryInfo(name, 'let x = 1 in x')],tmp_path/'queries.zip')
    with zipfile.ZipFile(package) as archive:
        entries=[entry for entry in archive.namelist() if entry.endswith('.pq')]
        manifest=json.loads(archive.read('pqm-package.json'))
    assert len(entries)==1
    assert len(entries[0].split('/')[-1].encode('utf-8'))<=255
    assert manifest['queries'][0]['query_name']==name


def test_profile_roundtrip_and_reject_invalid(tmp_path):
    data={'schema_version':1,'profile_name':'Pöwer','workbook':{'path':'../b.xlsx'},
          'options':{'existing_query_policy':'replace','backup':True,'dry_run':False,'refresh':'none'},
          'queries':[{'id':'stable','source_path':'q.pq','query_name':'Q','enabled':True,'order':0,'load':{'mode':'definition_only'}}]}
    path=tmp_path/'p.json'; save_profile(data,path); assert load_profile(path)==data
    data['queries'][0]['order']=-1
    with pytest.raises(ValueError): save_profile(data,path)


@pytest.mark.parametrize('version_field', [None, 0])
def test_profile_loader_migrates_unversioned_v0_to_v1_without_rewriting_file(tmp_path, version_field):
    document={'profile_name':'Legacy','workbook':{'path':'book.xlsx'},
        'options':{'existing_query_policy':'replace','backup':True,'dry_run':False,'refresh':'none'},
        'queries':[{'id':'stable','source_path':'query.pq','query_name':'Q','enabled':True,
                    'order':0,'load':{'mode':'definition_only'}}]}
    if version_field is not None:
        document['schema_version']=version_field
    path=tmp_path/'legacy-profile.json'
    original=json.dumps(document)
    path.write_text(original,encoding='utf-8')

    migrated=load_profile(path)

    assert migrated['schema_version']==1
    assert migrated['profile_name']=='Legacy'
    assert path.read_text(encoding='utf-8')==original


@pytest.mark.parametrize(('version_field','extra'), [(2, {}), (True, {}), (0, {'legacy_guess':True})])
def test_profile_loader_refuses_unknown_or_ambiguous_migrations(tmp_path, version_field, extra):
    document={'schema_version':version_field,'profile_name':'Legacy','workbook':{'path':'book.xlsx'},
        'options':{'existing_query_policy':'replace','backup':True,'dry_run':False,'refresh':'none'},
        'queries':[{'id':'stable','source_path':'query.pq','query_name':'Q','enabled':True,
                    'order':0,'load':{'mode':'definition_only'}}],**extra}
    path=tmp_path/'unsupported-profile.json'
    path.write_text(json.dumps(document),encoding='utf-8')
    with pytest.raises(ValueError,match='Cannot load profile'):
        load_profile(path)


def test_profile_schema_version_must_be_an_integer(tmp_path):
    from pq_installer.core.manager import validate_profile
    document={'schema_version':1.0,'profile_name':'P','workbook':{'path':'book.xlsx'},
        'options':{'existing_query_policy':'replace','backup':True,'dry_run':False,'refresh':'none'},
        'queries':[{'id':'stable','source_path':'query.pq','query_name':'Q','enabled':True,
                    'order':0,'load':{'mode':'definition_only'}}]}
    with pytest.raises(ValueError,match='schema_version'):
        validate_profile(document)


def test_profile_builder_keeps_row_ids_and_uses_relative_paths(tmp_path):
    from pq_installer.core.manager import build_import_profile
    from pq_installer.core.models import QuerySource
    source_path=tmp_path/'sources'/'café.pq';source_path.parent.mkdir();source_path.write_text('let x=1 in x')
    source=QuerySource(source_path,'Café',order=4,enabled=False)
    profile=build_import_profile('Morning',tmp_path/'book.xlsx',[source],['stable-row'],base_dir=tmp_path)
    assert profile['queries'][0]['id']=='stable-row'
    assert profile['queries'][0]['source_path']=='sources/café.pq'
    assert profile['workbook']['path']=='book.xlsx'
    from pq_installer.core.manager import validate_profile
    validate_profile(profile)


def test_profile_load_resolves_relative_paths_and_preserves_stable_ids(tmp_path):
    from pq_installer.core.manager import build_import_profile,profile_to_sources
    from pq_installer.core.models import QuerySource
    query=tmp_path/'queries'/'a.pq';query.parent.mkdir();query.write_text('let x=1 in x')
    profile_path=tmp_path/'profiles'/'nightly.json';profile_path.parent.mkdir()
    document=build_import_profile('Nightly',tmp_path/'book.xlsx',[QuerySource(query,'A',3)],['row-123'],base_dir=profile_path.parent)
    workbook,sources,ids=profile_to_sources(document,profile_path)
    assert workbook==tmp_path/'book.xlsx' and sources[0].path==query and ids==['row-123']


def test_connection_profile_requires_capability_then_roundtrips_mode(tmp_path):
    from pq_installer.core.manager import build_import_profile,profile_to_sources
    from pq_installer.core.models import QuerySource,LoadMode
    query=tmp_path/'q.pq';query.write_text('let x=1 in x')
    document=build_import_profile('Conn',tmp_path/'b.xlsx',[QuerySource(query,'Q',load_mode=LoadMode.CONNECTION_ONLY)],['r'],base_dir=tmp_path)
    with pytest.raises(ValueError,match='capability probe'):profile_to_sources(document,tmp_path/'p.json')
    _,sources,ids=profile_to_sources(document,tmp_path/'p.json',supported_load_modes=('definition_only','connection_only'))
    assert sources[0].load_mode==LoadMode.CONNECTION_ONLY and ids==['r']


def test_profile_round_trip_preserves_worksheet_destination(tmp_path):
    from pq_installer.core.manager import build_import_profile,profile_to_sources
    from pq_installer.core.models import LoadMode,QuerySource
    query=tmp_path/'load.pq';query.write_text('let x=1 in x')
    source=QuerySource(query,'Output',load_mode=LoadMode.EXISTING_SHEET,
                       sheet_name='Results',cell='$D$5',table_name='PQ_Output')
    profile_path=tmp_path/'load-profile.json'
    document=build_import_profile('Workbook load',tmp_path/'book.xlsx',[source],['load-row'],base_dir=tmp_path)
    _,sources,ids=profile_to_sources(document,profile_path,
                                     supported_load_modes=('definition_only','existing_sheet'))
    assert ids==['load-row']
    assert (sources[0].sheet_name,sources[0].cell,sources[0].table_name)==('Results','$D$5','PQ_Output')


def test_profile_load_refuses_unavailable_load_and_refresh_modes(tmp_path):
    from pq_installer.core.manager import profile_to_sources
    document={'schema_version':1,'profile_name':'P','workbook':{'path':'book.xlsx'},
              'options':{'existing_query_policy':'replace','backup':True,'dry_run':False,'refresh':'selected'},
              'queries':[{'id':'r','source_path':'q.pq','query_name':'Q','enabled':True,'order':0,'load':{'mode':'data_model'}}]}
    with pytest.raises(ValueError,match='refresh'):profile_to_sources(document,tmp_path/'p.json')
    document['options']['refresh']='none'
    with pytest.raises(ValueError,match='Load mode'):profile_to_sources(document,tmp_path/'p.json')


def test_profile_loader_rejects_duplicate_json_properties(tmp_path):
    path=tmp_path/'duplicate.json';path.write_text('{"schema_version":1,"schema_version":1}')
    with pytest.raises(ValueError,match='duplicate'):load_profile(path)


def test_profile_path_mapping_preview_is_anchored_and_longest_prefix(tmp_path):
    from pq_installer.core.manager import preview_profile_path_mappings
    root=tmp_path/'profile';root.mkdir()
    source=root/'queries'/'shared'/'q.pq';source.parent.mkdir(parents=True);source.write_text('let x=1 in x')
    document={'schema_version':1,'profile_name':'P','workbook':{'path':'book.xlsx'},
      'options':{'existing_query_policy':'replace','backup':True,'dry_run':False,'refresh':'none'},
      'path_mappings':[{'key':'general','from':'queries','to':'target','enabled':True},
                       {'key':'shared','from':'queries/shared','to':'target/special','enabled':True}],
      'queries':[{'id':'q1','source_path':'queries/shared/q.pq','query_name':'Q','enabled':True,'order':0,'load':{'mode':'definition_only'}}]}
    rows=preview_profile_path_mappings(document,root/'profile.json')
    assert len(rows)==1 and rows[0].mapping_key=='shared'
    assert rows[0].mapped_path==str((root/'target'/'special'/'q.pq').resolve())
    document['queries'][0]['source_path']='queries_other/q.pq'
    rows=preview_profile_path_mappings(document,root/'profile.json')
    assert rows[0].mapping_key is None and rows[0].mapped_path.endswith('queries_other\\q.pq')


def test_profile_mapping_apply_requires_preview_and_rejects_ambiguous_mapping(tmp_path):
    from pq_installer.core.manager import profile_to_sources
    root=tmp_path/'p';root.mkdir();(root/'a').mkdir();(root/'b').mkdir()
    query=root/'a'/'q.pq';query.write_text('let x=1 in x')
    document={'schema_version':1,'profile_name':'P','workbook':{'path':'book.xlsx'},
      'options':{'existing_query_policy':'replace','backup':True,'dry_run':False,'refresh':'none'},
      'path_mappings':[{'key':'one','from':'a','to':'b','enabled':True}],
      'queries':[{'id':'q','source_path':'a/q.pq','query_name':'Q','enabled':True,'order':0,'load':{'mode':'definition_only'}}]}
    with pytest.raises(ValueError,match='preview and approve'):
      profile_to_sources(document,root/'profile.json')
    _,sources,_=profile_to_sources(document,root/'profile.json',apply_path_mappings=True)
    assert sources[0].path==(root/'b'/'q.pq').resolve()
    document['path_mappings'].append({'key':'also','from':'a','to':'b/other','enabled':True})
    with pytest.raises(ValueError,match='ambiguous'):
      profile_to_sources(document,root/'profile.json',apply_path_mappings=True)


def test_profile_environment_mapping_preview_and_materialized_copy_preserve_original(tmp_path):
    from pq_installer.core.manager import (preview_profile_environment_mappings,
        materialize_profile_environment_mapping_copies)
    source=tmp_path/'sources'/'one.pq';source.parent.mkdir();source.write_bytes(
      b'\xef\xbb\xbf// keep\r\nFile.Contents("C:\\old\\folder\\data.csv")\r\n')
    document={'schema_version':1,'profile_name':'P','workbook':{'path':'book.xlsx'},
      'options':{'existing_query_policy':'replace','backup':True,'dry_run':False,'refresh':'none'},
      'queries':[{'id':'row-1','source_path':'sources/one.pq','query_name':'One','enabled':True,'order':0,'load':{'mode':'definition_only'}}],
      'environment_mappings':[{'key':'disk','from':'C:\\old','to':'D:\\new','enabled':True}]}
    profile_path=tmp_path/'profile.json'
    previews=preview_profile_environment_mappings(document,profile_path)
    assert len(previews)==1 and previews[0].changed
    assert 'File.Contents("D:\\new\\folder\\data.csv")' in previews[0].mapped_formula
    mapped_paths,staging=materialize_profile_environment_mapping_copies(previews)
    assert mapped_paths['row-1'].read_bytes()==b'\xef\xbb\xbf// keep\r\nFile.Contents("D:\\new\\folder\\data.csv")\r\n'
    assert source.read_bytes().startswith(b'\xef\xbb\xbf// keep\r\nFile.Contents("C:\\old')
    from shutil import rmtree
    rmtree(staging)


def test_profile_environment_mapping_preview_skips_disabled_missing_source(tmp_path):
    from pq_installer.core.manager import preview_profile_environment_mappings
    source=tmp_path/'enabled.pq';source.write_text('File.Contents("C:\\old\\x.csv")',encoding='utf-8')
    document={'schema_version':1,'profile_name':'P','workbook':{'path':'book.xlsx'},
      'options':{'existing_query_policy':'replace','backup':True,'dry_run':False,'refresh':'none'},
      'queries':[
        {'id':'enabled','source_path':'enabled.pq','query_name':'Enabled','enabled':True,'order':0,'load':{'mode':'definition_only'}},
        {'id':'disabled','source_path':'not-present.pq','query_name':'Disabled','enabled':False,'order':1,'load':{'mode':'definition_only'}}],
      'environment_mappings':[{'key':'disk','from':'C:\\old','to':'D:\\new','enabled':True}]}
    previews=preview_profile_environment_mappings(document,tmp_path/'profile.json')
    assert [row.query_id for row in previews]==['enabled']
    assert previews[0].changed


def test_profile_environment_mapping_materialization_rejects_source_drift(tmp_path):
    from pq_installer.core.manager import (preview_profile_environment_mappings,
        materialize_profile_environment_mapping_copies)
    source=tmp_path/'q.pq';source.write_text('File.Contents("C:\\old\\x.csv")',encoding='utf-8')
    document={'schema_version':1,'profile_name':'P','workbook':{'path':'book.xlsx'},
      'options':{'existing_query_policy':'replace','backup':True,'dry_run':False,'refresh':'none'},
      'queries':[{'id':'row-1','source_path':'q.pq','query_name':'Q','enabled':True,'order':0,'load':{'mode':'definition_only'}}],
      'environment_mappings':[{'key':'disk','from':'C:\\old','to':'D:\\new','enabled':True}]}
    previews=preview_profile_environment_mappings(document,tmp_path/'profile.json')
    source.write_text('File.Contents("C:\\new\\x.csv")',encoding='utf-8')
    with pytest.raises(RuntimeError,match='changed after preview'):
        materialize_profile_environment_mapping_copies(previews)


def test_profile_save_keeps_original_source_path_after_temporary_mapped_copy(tmp_path):
    from pq_installer.core.manager import build_import_profile
    from pq_installer.core.models import QuerySource
    original=tmp_path/'source'/'q.pq';original.parent.mkdir();original.write_text('File.Contents("C:\\old")')
    mapped=tmp_path/'staging'/'001_q.pq';mapped.parent.mkdir();mapped.write_text('File.Contents("D:\\new")')
    source=QuerySource(mapped,'Q',profile_source_path=original)
    profile=build_import_profile('Mapped',tmp_path/'book.xlsx',[source],['row'],base_dir=tmp_path)
    assert profile['queries'][0]['source_path']=='source/q.pq'


def test_profile_m_code_environment_mapping_requires_explicit_approval(tmp_path):
    from pq_installer.core.manager import profile_to_sources
    source=tmp_path/'q.pq';source.write_text('File.Contents("C:\\old\\q.csv")')
    profile={'schema_version':1,'profile_name':'P','workbook':{'path':'book.xlsx'},
      'options':{'existing_query_policy':'replace','backup':True,'dry_run':False,'refresh':'none'},
      'environment_mappings':[{'key':'disk','from':'C:\\old','to':'D:\\new','enabled':True}],
      'queries':[{'id':'q1','source_path':'q.pq','query_name':'Q','enabled':True,'order':0,'load':{'mode':'definition_only'}}]}
    with pytest.raises(ValueError,match='preview and approve'):
        profile_to_sources(profile,tmp_path/'p.json')
    _,sources,_=profile_to_sources(profile,tmp_path/'p.json',apply_environment_mappings=True)
    assert sources[0].path==source


def test_materializing_unmatched_environment_mapping_does_not_create_temp_tree():
    from pq_installer.core.manager import materialize_profile_environment_mapping_copies
    assert materialize_profile_environment_mapping_copies(())==({},None)


def test_profile_requires_all_schema_options_and_rejects_unknown_properties():
    data={'schema_version':1,'profile_name':'P','workbook':{'path':'book.xlsx'},
          'options':{'existing_query_policy':'replace','backup':True,'dry_run':True,'refresh':'none'},
          'queries':[{'id':'id','source_path':'x.pq','query_name':'X','enabled':True,'order':0,'load':{'mode':'existing_sheet','sheet_name':'Data','cell':'B2'}}]}
    from pq_installer.core.manager import validate_profile
    validate_profile(data)
    del data['options']['backup']
    with pytest.raises(ValueError,match='backup'): validate_profile(data)
    data['options']['backup']=True;data['surprise']=1
    with pytest.raises(ValueError,match='unexpected'): validate_profile(data)


def test_export_writes_manifest_and_never_overwrites_manifest(tmp_path):
    queries=[QueryInfo('Original Name','let a=1 in a')]
    outputs=export_queries(queries,tmp_path)
    manifest=tmp_path/'query-manifest.json'
    assert manifest.exists() and json.loads(manifest.read_text(encoding='utf-8'))['queries'][0]['query_name']=='Original Name'
    export_queries(queries,tmp_path)
    assert (tmp_path/'query-manifest-2.json').exists()


def test_query_zip_rejects_traversal_duplicate_and_symlink(tmp_path):
    from pq_installer.core.packages import extract_query_zip
    archive=tmp_path/'bad.zip'
    with zipfile.ZipFile(archive,'w') as z:
        z.writestr('../escape.m','x')
    with pytest.raises(ValueError,match='unsafe path'): extract_query_zip(archive)
    with zipfile.ZipFile(archive,'w') as z:
        z.writestr('a/../x.m','x')
    with pytest.raises(ValueError,match='unsafe path'): extract_query_zip(archive)
    with zipfile.ZipFile(archive,'w') as z:
        z.writestr('a.m','one');z.writestr('./a.m','two')
    with pytest.raises(ValueError,match='duplicate'): extract_query_zip(archive)
    with zipfile.ZipFile(archive,'w') as z:
        link=zipfile.ZipInfo('link.m');link.create_system=3;link.external_attr=(0o120777<<16);z.writestr(link,'target')
    with pytest.raises(ValueError,match='symbolic link'): extract_query_zip(archive)


def test_query_zip_extracts_only_safe_m_sources_in_isolated_folder(tmp_path):
    from pq_installer.core.packages import extract_query_zip
    archive=tmp_path/'good.zip'
    with zipfile.ZipFile(archive,'w') as z:z.writestr('queries/1_base.m','let x=1 in x')
    extracted=extract_query_zip(archive,staging_parent=tmp_path)
    assert extracted.parent==tmp_path and (extracted/'queries/1_base.m').read_text()=='let x=1 in x'


def test_query_zip_enforces_uncompressed_size_and_windows_reserved_names(tmp_path):
    from pq_installer.core.packages import extract_query_zip
    archive=tmp_path/'oversized.zip'
    with zipfile.ZipFile(archive,'w',compression=zipfile.ZIP_STORED) as z:z.writestr('large.m','x'*100)
    with pytest.raises(ValueError,match='size limit'): extract_query_zip(archive,max_uncompressed=50)
    with zipfile.ZipFile(archive,'w') as z:z.writestr('CON.m','x')
    with pytest.raises(ValueError,match='reserved'): extract_query_zip(archive)


def test_query_zip_rejects_oversized_archive_before_temp_extraction(tmp_path):
    from pq_installer.core.packages import extract_query_zip

    archive=tmp_path/'oversized-archive.zip'
    with zipfile.ZipFile(archive,'w') as z:
        z.writestr('query.m','let x=1 in x')
    archive.write_bytes(archive.read_bytes()+b'padding'*300)

    with pytest.raises(ValueError,match='archive size limit'):
        extract_query_zip(archive,staging_parent=tmp_path,max_archive_bytes=1024)

    assert list(tmp_path.glob('pq-import-*')) == []


def test_query_zip_size_check_and_parser_use_same_open_file(tmp_path, monkeypatch):
    import pq_installer.core.packages as package_module

    archive=tmp_path/'selected.zip'
    replacement=tmp_path/'replacement.zip'
    with zipfile.ZipFile(archive,'w') as z:
        z.writestr('selected.m','let x=1 in x')
    with zipfile.ZipFile(replacement,'w',compression=zipfile.ZIP_STORED) as z:
        z.writestr('replacement.m','let x=2 in x')
        z.writestr('padding.m','x'*2048)
    real_zipfile=package_module.zipfile.ZipFile
    parser_inputs=[]

    def replace_path_after_open(file,*args,**kwargs):
        is_open_file=hasattr(file,'fileno')
        parser_inputs.append(is_open_file)
        # Model a path replacement at the parser boundary without relying on
        # Windows allowing replacement of a file that is open without delete
        # sharing. A path-based parser would see replacement.zip; a parser
        # given the checked open handle must continue reading selected.zip.
        return real_zipfile(file if is_open_file else replacement,*args,**kwargs)

    monkeypatch.setattr(package_module.zipfile,'ZipFile',replace_path_after_open)
    extracted=package_module.extract_query_zip(
        archive,staging_parent=tmp_path,max_archive_bytes=1024)

    assert parser_inputs==[True]
    assert {path.name for path in extracted.rglob('*.m')}=={'selected.m'}


def test_versioned_query_package_roundtrips_names_order_and_exact_formula(tmp_path):
    from pq_installer.core.manager import export_query_package
    from pq_installer.core.packages import extract_query_zip, apply_query_package_manifest
    from pq_installer.core.models import QuerySource
    from pq_installer.core.query_loader import discover_query_files
    queries=[QueryInfo('Sales Raw','let Source = "café" in Source'),QueryInfo('Helper','// keep CRLF\r\nlet x=1 in x')]
    package=tmp_path/'bundle.pqm.zip'
    export_query_package(queries,package)
    extracted=extract_query_zip(package,staging_parent=tmp_path)
    sources=discover_query_files([extracted],recursive=True)
    restored=apply_query_package_manifest(sources,extracted)
    assert [(s.name,s.order) for s in restored]==[('Sales Raw',0),('Helper',1)]
    assert [Path(s.path).read_bytes().decode('utf-8') for s in restored]==[q.formula for q in queries]
    with zipfile.ZipFile(package) as archive:manifest=json.loads(archive.read('pqm-package.json'))
    schema=json.loads((Path(__file__).resolve().parents[2]/'contracts'/'query-package.schema.json').read_text(encoding='utf-8'))
    from jsonschema import Draft202012Validator
    Draft202012Validator(schema).validate(manifest)
    assert manifest['schema_version']==2
    assert manifest['queries'][0]['dependencies']==[]
    assert manifest['queries'][0]['dependency_status']=='known'


def test_query_package_update_preview_classifies_all_safe_change_states(tmp_path):
    from pq_installer.core.manager import export_query_package, compare_query_packages
    base=export_query_package([
        QueryInfo('Stable','let Value = 1 in Value'),
        QueryInfo('Changed','let Value = 1 in Value'),
        QueryInfo('Removed','let Value = 9 in Value')],tmp_path/'base.pqm.zip')
    update=export_query_package([
        QueryInfo('Stable','let Value = 1 in Value'),
        QueryInfo('Changed','let Value = 2 in Value'),
        QueryInfo('Added','let Value = 3 in Value')],tmp_path/'update.pqm.zip')
    preview=compare_query_packages(base,update)
    assert [(row.name,row.status) for row in preview]==[
        ('Stable','UNCHANGED'),('Changed','MODIFIED'),('Added','ADDED'),('Removed','REMOVED')]
    modified=next(row for row in preview if row.name=='Changed')
    assert '-let Value = 1 in Value' in modified.diff
    assert '+let Value = 2 in Value' in modified.diff
    removed=next(row for row in preview if row.name=='Removed')
    assert 'Removed' in removed.diff and 'let Value = 9' in removed.diff


def test_query_package_update_preview_flags_case_only_rename_as_conflict(tmp_path):
    from pq_installer.core.manager import export_query_package, compare_query_packages
    base=export_query_package([QueryInfo('Sales','1')],tmp_path/'base.pqm.zip')
    update=export_query_package([QueryInfo('sales','1')],tmp_path/'update.pqm.zip')
    preview=compare_query_packages(base,update)
    assert [(row.name,row.status) for row in preview]==[('sales','CONFLICT')]


def test_query_package_update_preview_rejects_tampered_package_and_cleans_staging(tmp_path):
    import zipfile
    from pq_installer.core.manager import export_query_package, compare_query_packages
    base=export_query_package([QueryInfo('Q','1')],tmp_path/'base.pqm.zip')
    original=export_query_package([QueryInfo('Q','2')],tmp_path/'valid-update.pqm.zip')
    tampered=tmp_path/'tampered.pqm.zip'
    with zipfile.ZipFile(original) as src,zipfile.ZipFile(tampered,'w') as dst:
        for item in src.infolist():
            data=src.read(item.filename)
            if item.filename.endswith('.pq'):data=b'let Changed = 3 in Changed'
            dst.writestr(item.filename,data)
    with pytest.raises(ValueError,match='hash mismatch'):
        compare_query_packages(base,tampered)
    assert not list(tmp_path.glob('pqm-*'))


def test_query_package_manifest_records_only_unambiguous_dependencies(tmp_path):
    from pq_installer.core.manager import export_query_package
    from pq_installer.core.packages import extract_query_zip,validate_query_package_manifest
    queries=[QueryInfo('Helper','1'),QueryInfo('Report','let Source = Helper in Source'),
             QueryInfo('Dynamic','Expression.Evaluate("Helper")')]
    package=export_query_package(queries,tmp_path/'deps.pqm.zip')
    extracted=extract_query_zip(package,staging_parent=tmp_path)
    manifest=validate_query_package_manifest(extracted)
    assert manifest['queries'][1]['dependencies']==['Helper']
    assert manifest['queries'][1]['dependency_status']=='known'
    assert manifest['queries'][2]['dependencies']==[]
    assert manifest['queries'][2]['dependency_status']=='ambiguous'
    assert manifest['dependency_cycles']==[]


def test_query_package_manifest_records_cycles_without_reordering_rows(tmp_path):
    from pq_installer.core.manager import export_query_package
    from pq_installer.core.packages import extract_query_zip,validate_query_package_manifest
    package=export_query_package([QueryInfo('A','B'),QueryInfo('B','A')],tmp_path/'cycle.pqm.zip')
    root=extract_query_zip(package,staging_parent=tmp_path)
    manifest=validate_query_package_manifest(root)
    assert manifest['dependency_cycles']==[['A','B','A']]
    assert [row['query_name'] for row in manifest['queries']]==['A','B']


def test_query_package_v2_rejects_tampered_dependency_metadata(tmp_path):
    from pq_installer.core.manager import export_query_package
    from pq_installer.core.packages import extract_query_zip,validate_query_package_manifest
    package=export_query_package([QueryInfo('Helper','1'),QueryInfo('Report','Helper')],tmp_path/'deps.pqm.zip')
    tampered=tmp_path/'tampered-deps.pqm.zip'
    with zipfile.ZipFile(package) as source,zipfile.ZipFile(tampered,'w') as destination:
        for info in source.infolist():
            data=source.read(info.filename)
            if info.filename=='pqm-package.json':
                manifest=json.loads(data);manifest['queries'][1]['dependencies']=[]
                data=json.dumps(manifest).encode('utf-8')
            destination.writestr(info,data)
    with pytest.raises(ValueError,match='dependency metadata'):
        extract_query_zip(tampered,staging_parent=tmp_path)
    assert not list(tmp_path.glob('pq-import-*'))


def test_query_package_writer_rejects_duplicate_names_and_oversize_payloads(tmp_path):
    from pq_installer.core.manager import export_query_package
    with pytest.raises(ValueError,match='duplicate case-insensitive'):
        export_query_package([QueryInfo('Q','1'),QueryInfo('q','2')],tmp_path/'duplicate.pqm.zip')
    with pytest.raises(ValueError,match='configured size limit'):
        export_query_package([QueryInfo('Large','123456789')],tmp_path/'large.pqm.zip',max_uncompressed=8)


def test_version_one_query_package_manifest_remains_importable(tmp_path):
    from pq_installer.core.manager import export_query_package
    from pq_installer.core.packages import extract_query_zip,validate_query_package_manifest,migrate_query_package_manifest,apply_query_package_manifest
    from pq_installer.core.query_loader import discover_query_files
    package=export_query_package([QueryInfo('Legacy','1'),QueryInfo('Report','let Source = Legacy in Source')],tmp_path/'legacy.pqm.zip')
    legacy=tmp_path/'legacy-v1.pqm.zip'
    with zipfile.ZipFile(package) as source,zipfile.ZipFile(legacy,'w') as destination:
        for info in source.infolist():
            if info.filename=='pqm-package.json':
                manifest=json.loads(source.read(info.filename));manifest['schema_version']=1
                manifest.pop('dependency_cycles')
                for row in manifest['queries']:
                    row.pop('dependencies');row.pop('dependency_status')
                data=json.dumps(manifest).encode('utf-8')
            else:data=source.read(info.filename)
            destination.writestr(info,data)
    extracted=extract_query_zip(legacy,staging_parent=tmp_path)
    assert validate_query_package_manifest(extracted)['schema_version']==1
    migrated=migrate_query_package_manifest(extracted)
    assert migrated['schema_version']==2
    assert migrated['queries'][1]['dependencies']==['Legacy']
    assert migrated['queries'][1]['dependency_status']=='known'
    restored=apply_query_package_manifest(discover_query_files([extracted],recursive=True),extracted)
    assert [(source.name,source.order) for source in restored]==[('Legacy',0),('Report',1)]


def test_versioned_query_package_verifies_hash_and_never_overwrites(tmp_path):
    import zipfile
    from pq_installer.core.manager import export_query_package
    from pq_installer.core.packages import extract_query_zip
    output=tmp_path/'bundle.zip'
    export_query_package([QueryInfo('Q','let x=1 in x')],output)
    original=output.read_bytes()
    with pytest.raises(FileExistsError):export_query_package([QueryInfo('Q','changed')],output)
    assert output.read_bytes()==original
    tampered=tmp_path/'tampered.zip'
    with zipfile.ZipFile(output) as src,zipfile.ZipFile(tampered,'w') as dst:
        for item in src.infolist():
            data=src.read(item.filename)
            if item.filename.endswith('.pq'):data=b'changed formula'
            dst.writestr(item.filename,data)
    with pytest.raises(ValueError,match='hash mismatch'):extract_query_zip(tampered,staging_parent=tmp_path)
    assert not list(tmp_path.glob('pq-import-*'))
