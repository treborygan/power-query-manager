"""Read-only query management services; adapters are injected and never refresh."""
from dataclasses import dataclass
from enum import Enum
from difflib import unified_diff
from pathlib import Path
import json
import re
import hashlib
import os
import tempfile
import zipfile
from .models import QuerySource, LoadMode
from .m_source_mapping import preview_m_source_mapping
from .dependencies import analyze_query_references, find_dependency_cycles


class Comparison(str, Enum):
    NEW = 'NEW'
    MODIFIED = 'MODIFIED'
    UNCHANGED = 'UNCHANGED'
    SKIP = 'SKIP'
    CONFLICT = 'CONFLICT'
    UNKNOWN = 'UNKNOWN'


@dataclass(frozen=True)
class QueryInfo:
    name: str
    formula: str | None
    description: str | None = None
    load_state: str = 'unknown'


@dataclass(frozen=True)
class PathMappingPreview:
    query_id: str
    source_path: str
    mapped_path: str
    mapping_key: str | None
    changed: bool
    target_exists: bool
    ambiguous: bool = False


@dataclass(frozen=True)
class ProfileEnvironmentMappingPreview:
    query_id: str
    query_name: str
    source_path: str
    source_sha256: str
    original_formula: str
    mapped_formula: str
    changes: tuple
    warnings: tuple[str, ...]

    @property
    def changed(self):
        return self.original_formula != self.mapped_formula


@dataclass(frozen=True)
class QueryPackageChange:
    name: str
    status: str
    previous_name: str | None
    old_formula: str | None
    new_formula: str | None
    diff: str


def compare_source(name: str, formula: str, existing: QueryInfo | None, *, skip=False) -> Comparison:
    if existing is None:
        return Comparison.NEW
    if skip:
        return Comparison.SKIP
    if existing.name != name:
        return Comparison.CONFLICT
    if existing.formula is None:
        return Comparison.UNKNOWN
    def digest(value):
        normalized=value.replace('\r\n','\n').replace('\r','\n')
        # Excel removes terminal CR/LF when it saves a query formula. Treat
        # that observed round-trip change as equivalent for status only;
        # diff_m_code still shows the original source and workbook text.
        normalized=normalized.rstrip('\n')
        return hashlib.sha256(normalized.encode('utf-8')).digest()
    return Comparison.UNCHANGED if digest(formula) == digest(existing.formula) else Comparison.MODIFIED


def diff_m_code(source_formula: str, workbook_formula: str | None, *, source_name='Source', workbook_name='Workbook') -> str:
    """Return an exact line diff; no normalization is used for status decisions."""
    if workbook_formula is None:
        return f'Workbook formula for {workbook_name} is unavailable; comparison is unknown.\n'
    return ''.join(unified_diff(workbook_formula.splitlines(keepends=True),
                                source_formula.splitlines(keepends=True),
                                fromfile=workbook_name, tofile=source_name))


def compare_query_packages(base_package, updated_package) -> tuple[QueryPackageChange, ...]:
    """Compare two untrusted query packages without changing either package or workbook.

    Removed rows are represented for review only; the result never authorizes workbook deletion.
    """
    from .packages import extract_query_zip, apply_query_package_manifest
    from .query_loader import discover_query_files, read_m_code

    def read_package(path, parent):
        root = extract_query_zip(path, staging_parent=parent)
        sources = apply_query_package_manifest(
            discover_query_files([root], recursive=True), root)
        formulas = {}
        display_names = {}
        for source in sources:
            key = source.name.casefold()
            if key in formulas:
                raise ValueError('Query package contains duplicate case-insensitive query names.')
            formulas[key] = read_m_code(source)
            display_names[key] = source.name
        return formulas, display_names

    with tempfile.TemporaryDirectory(prefix='pqm-update-preview-') as temporary:
        base_dir = Path(temporary) / 'base'
        update_dir = Path(temporary) / 'update'
        base_dir.mkdir(); update_dir.mkdir()
        old, old_names = read_package(base_package, base_dir)
        new, new_names = read_package(updated_package, update_dir)

    results=[]
    for key, new_formula in new.items():
        name=new_names[key]
        old_formula=old.get(key)
        previous_name=old_names.get(key)
        if previous_name is not None and previous_name != name:
            status='CONFLICT'
        elif old_formula is None:
            status='ADDED'
        elif old_formula == new_formula:
            status='UNCHANGED'
        else:
            status='MODIFIED'
        diff=''.join(unified_diff(
            (old_formula or '').splitlines(keepends=True),
            new_formula.splitlines(keepends=True),
            fromfile=previous_name or '/dev/null',tofile=name))
        results.append(QueryPackageChange(name,status,previous_name,old_formula,new_formula,diff))
    for key, old_formula in old.items():
        if key in new:continue
        name=old_names[key]
        diff=''.join(unified_diff(old_formula.splitlines(keepends=True),[],
                                  fromfile=name,tofile='/dev/null'))
        results.append(QueryPackageChange(name,'REMOVED',name,old_formula,None,diff))
    return tuple(results)


def inspect_workbook(adapter) -> tuple[QueryInfo, ...]:
    """Call adapter inventory only. It must not save, refresh, or modify the workbook."""
    rows = adapter.inspect_queries()
    return tuple(QueryInfo(str(q.name), getattr(q, 'formula', None), getattr(q, 'description', None),
                           getattr(q, 'load_state', 'unknown')) for q in rows)


_RESERVED_QUERY_FILE_STEMS = {
    'con', 'prn', 'aux', 'nul', *(f'com{i}' for i in range(1, 10)),
    *(f'lpt{i}' for i in range(1, 10)),
}


def _safe_query_file_stem(name: str) -> str:
    """Map a query name to a portable, bounded filename component."""
    stem = re.sub(r'[<>:"/\\|?*\x00-\x1f]', '_', str(name)).strip(' .') or 'query'
    # 60 Unicode code points stay under common 255-byte UTF-8 and 255-unit
    # UTF-16 component limits, even for four-byte characters. Leave room for
    # package order prefixes and the .pq extension.
    stem = stem[:60].rstrip(' .') or 'query'
    if stem.split('.')[0].casefold() in _RESERVED_QUERY_FILE_STEMS:
        stem = '_' + stem
    return stem


def export_queries(queries, folder) -> tuple[Path, ...]:
    folder = Path(folder); folder.mkdir(parents=True, exist_ok=True)
    used = set(); written = []; manifest_rows = []
    for query in queries:
        if query.formula is None:
            continue
        stem = _safe_query_file_stem(query.name)
        candidate = stem; index = 2
        while candidate.casefold() in used or (folder / f'{candidate}.pq').exists():
            candidate = f'{stem}_{index}'; index += 1
        used.add(candidate.casefold())
        target = folder / f'{candidate}.pq'
        with target.open('x', encoding='utf-8', newline='') as stream:
            stream.write(query.formula)
        written.append(target)
        manifest_rows.append({'query_name': query.name, 'file': target.name,
                              'sha256': hashlib.sha256(query.formula.encode('utf-8')).hexdigest()})
    manifest = {'schema_version': 1, 'queries': manifest_rows}
    manifest_path = folder / 'query-manifest.json'
    suffix = 2
    while manifest_path.exists():
        manifest_path = folder / f'query-manifest-{suffix}.json'; suffix += 1
    with manifest_path.open('x', encoding='utf-8', newline='') as stream:
        stream.write(json.dumps(manifest, ensure_ascii=False, indent=2) + '\n')
    return tuple(written)


def export_query_package(queries, path, *, max_uncompressed=25 * 1024 * 1024, max_queries=255) -> Path:
    """Write a versioned, hash-verified query package without overwriting output."""
    destination=Path(path).expanduser().resolve()
    destination.parent.mkdir(parents=True,exist_ok=True)
    if destination.exists():raise FileExistsError(f'Package already exists: {destination}')
    rows=[];used_names=set()
    exportable=[query for query in queries if query.formula is not None]
    if len(exportable)>max_queries:raise ValueError('Query package exceeds the configured file-count limit.')
    names=[str(query.name) for query in exportable]
    if len({name.casefold() for name in names})!=len(names):
        raise ValueError('Cannot package duplicate case-insensitive query names.')
    analyses={str(query.name):analyze_query_references(query.formula,names) for query in exportable}
    graph={}
    for query in exportable:
        analysis=analyses[str(query.name)]
        graph[str(query.name)]=list(analysis.dependencies) if not analysis.ambiguous else []
    cycles=find_dependency_cycles(graph)
    for order,query in enumerate(exportable):
        if not isinstance(query.formula,str) or not query.formula.strip():
            raise ValueError(f'Query {query.name!r} has no usable M formula and cannot be packaged.')
        from .validation import validate_query_name
        validate_query_name(str(query.name))
        stem=_safe_query_file_stem(query.name)
        filename=f'queries/{order:04d}_{stem}.pq'
        if filename.casefold() in used_names:raise ValueError(f'Duplicate package filename generated for query {query.name!r}.')
        used_names.add(filename.casefold())
        formula_bytes=query.formula.encode('utf-8')
        analysis=analyses[str(query.name)]
        rows.append({'query_name':str(query.name),'file':filename,'order':len(rows),
                     'sha256':hashlib.sha256(formula_bytes).hexdigest(),
                     'dependencies':[] if analysis.ambiguous else list(analysis.dependencies),
                     'dependency_status':'ambiguous' if analysis.ambiguous else 'known',
                     'content':formula_bytes})
    if not rows:raise ValueError('No query formulas are available to package.')
    manifest={'schema_version':2,'package_type':'power-query-manager-query-package',
              'queries':[{k:v for k,v in row.items() if k!='content'} for row in rows],
              'dependency_cycles':[list(cycle) for cycle in cycles]}
    manifest_bytes=(json.dumps(manifest,ensure_ascii=False,indent=2)+'\n').encode('utf-8')
    total_size=sum(len(row['content']) for row in rows)+len(manifest_bytes)
    if total_size>max_uncompressed:
        raise ValueError('Query package exceeds the configured size limit.')
    temp_path=None
    try:
        with tempfile.NamedTemporaryFile(prefix='.pqm-package-',suffix='.tmp',dir=destination.parent,delete=False) as temp:
            temp_path=Path(temp.name)
        # Store rather than deflate M sources so repetitive user code cannot make
        # a package exceed the importer's compression-ratio guard.
        with zipfile.ZipFile(temp_path,'w',compression=zipfile.ZIP_STORED) as archive:
            for row in rows:archive.writestr(row['file'],row['content'])
            archive.writestr('pqm-package.json',manifest_bytes)
        # Hard-link creation is atomic and fails if another process created the destination.
        os.link(temp_path,destination)
        return destination
    finally:
        if temp_path is not None:
            try:temp_path.unlink()
            except FileNotFoundError:pass


def export_deployment_bundle(profile_path,path,*,apply_path_mappings=False,
                             max_uncompressed=25*1024*1024,max_queries=255) -> Path:
    """Create a portable deployment profile/source ZIP without overwriting output."""
    profile_path=Path(profile_path).expanduser().resolve()
    profile=load_profile(profile_path)
    if len(profile['queries'])>max_queries:
        raise ValueError('Deployment bundle exceeds the configured query-count limit.')
    # Environment mappings stay as reviewed rules in the portable profile. Only
    # source-path mappings are explicitly applied because the files are embedded.
    resolution_profile=dict(profile)
    resolution_profile.pop('environment_mappings',None)
    _,sources,stable_ids=profile_to_sources(resolution_profile,profile_path,
        supported_load_modes=('definition_only','connection_only','new_sheet','existing_sheet','data_model','preserve_existing'),
        apply_path_mappings=apply_path_mappings)
    by_id=dict(zip(stable_ids,sources))
    portable=json.loads(json.dumps(profile))
    portable['workbook']=dict(portable['workbook'])
    portable['workbook']['path']='deployment-target.xlsx'
    portable['workbook'].pop('output_path',None)
    # The source paths are now bundle-relative. Old root mappings must not be
    # applied a second time after extraction.
    portable['path_mappings']=[]

    source_rows=[];contents=[];total=0
    for index,query in enumerate(portable['queries']):
        source=by_id[query['id']]
        extension=source.path.suffix.casefold()
        if extension not in {'.m','.pq'}:
            raise ValueError(f'Query source {source.name!r} must use an .m or .pq extension.')
        relative=f'sources/{index:04d}{extension}'
        query['source_path']=relative
        if not source.path.is_file():
            if query['enabled']:
                raise ValueError(f'Enabled query source is missing: {source.name}.')
            continue
        content=source.path.read_bytes()
        try:formula=content.decode('utf-8-sig')
        except UnicodeDecodeError as exc:
            raise ValueError(f'Query source {source.name!r} is not valid UTF-8.') from exc
        if not formula.strip():raise ValueError(f'Query source {source.name!r} is empty.')
        total+=len(content)
        if total>max_uncompressed:
            raise ValueError('Deployment bundle exceeds the configured uncompressed-size limit.')
        digest=hashlib.sha256(content).hexdigest()
        source_rows.append({'query_id':query['id'],'file':relative,'sha256':digest})
        contents.append((relative,content))
    if not any(query['enabled'] for query in portable['queries']):
        raise ValueError('Deployment profile has no enabled query sources.')
    manifest={'schema_version':1,'bundle_type':'power-query-manager-deployment-bundle',
              'profile':portable,'sources':source_rows}
    manifest_bytes=(json.dumps(manifest,ensure_ascii=False,indent=2)+'\n').encode('utf-8')
    if total+len(manifest_bytes)>max_uncompressed:
        raise ValueError('Deployment bundle exceeds the configured uncompressed-size limit.')
    destination=Path(path).expanduser().resolve()
    if destination.suffix.casefold() not in {'.zip','.pqmdeploy'}:
        raise ValueError('Deployment bundle output must use .zip or .pqmdeploy.')
    destination.parent.mkdir(parents=True,exist_ok=True)
    if destination.exists():raise FileExistsError(f'Deployment bundle already exists: {destination}')
    temp_path=None
    try:
        with tempfile.NamedTemporaryFile(prefix='.pqm-deploy-',suffix='.tmp',dir=destination.parent,delete=False) as stream:
            temp_path=Path(stream.name)
        with zipfile.ZipFile(temp_path,'w',compression=zipfile.ZIP_STORED) as archive:
            for relative,content in contents:archive.writestr(relative,content)
            archive.writestr('pqm-deployment.json',manifest_bytes)
        os.link(temp_path,destination)
        return destination
    finally:
        if temp_path is not None:
            try:temp_path.unlink()
            except FileNotFoundError:pass


def validate_profile(document: dict) -> None:
    """Small dependency-free validation for profile v1's required invariants."""
    if (not isinstance(document, dict)
            or isinstance(document.get('schema_version'), bool)
            or not isinstance(document.get('schema_version'), int)
            or document.get('schema_version') != 1):
        raise ValueError('Unsupported or invalid profile schema_version (expected 1).')
    if set(document) - {'schema_version','profile_name','description','workbook','options','queries','path_mappings','environment_mappings'}:
        raise ValueError('Profile contains unexpected properties.')
    if not isinstance(document.get('profile_name'), str) or not document['profile_name'].strip():
        raise ValueError('Profile profile_name must be a non-empty string.')
    if len(document['profile_name']) > 128 or (isinstance(document.get('description'), str) and len(document['description']) > 2000):
        raise ValueError('Profile name or description exceeds its schema length limit.')
    if 'description' in document and (not isinstance(document['description'], str) or len(document['description']) > 2000):
        raise ValueError('Profile description must be a string of at most 2000 characters.')
    workbook = document.get('workbook')
    if not isinstance(workbook, dict) or set(workbook) - {'path','output_path'} or not isinstance(workbook.get('path'), str) or not workbook['path']:
        raise ValueError('Profile workbook.path is required.')
    if 'output_path' in workbook and (not isinstance(workbook['output_path'],str) or not workbook['output_path']):
        raise ValueError('Profile workbook.output_path must be a non-empty string.')
    options = document.get('options')
    if not isinstance(options, dict) or set(options) - {'existing_query_policy','backup','dry_run','refresh','import_order','staged_output'}:
        raise ValueError('Profile options contains unexpected properties.')
    required_options = {'existing_query_policy','backup','dry_run','refresh'}
    if required_options - set(options):
        raise ValueError('Profile options require existing_query_policy, backup, dry_run, and refresh.')
    if options.get('existing_query_policy') not in {'replace','skip','review_each'} or options.get('refresh') not in {'none','selected','all'}:
        raise ValueError('Profile policy or refresh option is invalid.')
    if any(not isinstance(options[k], bool) for k in ('backup','dry_run')):
        raise ValueError('Profile backup and dry_run options must be booleans.')
    if 'staged_output' in options and not isinstance(options['staged_output'], bool):
        raise ValueError('Profile staged_output must be boolean.')
    if 'import_order' in options and options['import_order'] not in {'manual','filename_prefix','dependency_suggestion'}:
        raise ValueError('Profile import_order is invalid.')
    if not isinstance(document.get('queries'), list) or not document['queries']:
        raise ValueError('Profile queries must be a non-empty array.')
    ids = set()
    for q in document['queries']:
        if not isinstance(q, dict) or any(k not in q for k in ('id','source_path','query_name','enabled','order','load')):
            raise ValueError('Each profile query requires id, source_path, query_name, enabled, order, and load.')
        if set(q) - {'id','source_path','query_name','enabled','order','existing_policy','load','depends_on'}:
            raise ValueError(f'Query {q.get("id")} contains unexpected properties.')
        if not all(isinstance(q[k], str) and q[k] for k in ('id','source_path','query_name')):
            raise ValueError('Profile query id, source_path, and query_name must be non-empty strings.')
        if q['id'] in ids: raise ValueError('Profile query IDs must be unique.')
        ids.add(q['id'])
        if len(q['id']) > 128 or len(q['query_name']) > 255:
            raise ValueError(f'Query {q["id"]} exceeds a schema length limit.')
        if not isinstance(q['enabled'], bool) or isinstance(q['order'],bool) or not isinstance(q['order'], int) or q['order'] < 0:
            raise ValueError(f'Invalid enabled/order value for query {q["id"]}.')
        if 'existing_policy' in q and q['existing_policy'] not in {'replace','skip','review'}:
            raise ValueError(f'Invalid existing_policy for query {q["id"]}.')
        load=q['load']
        if not isinstance(load, dict) or set(load)-{'mode','sheet_name','cell','table_name'} or load.get('mode') not in {'definition_only','connection_only','new_sheet','existing_sheet','data_model','preserve_existing'}:
            raise ValueError(f'Invalid load mode for query {q["id"]}.')
        mode=load['mode']
        if mode == 'existing_sheet':
            if not isinstance(load.get('sheet_name'), str) or not (1 <= len(load['sheet_name']) <= 31):
                raise ValueError(f'Existing-sheet load for query {q["id"]} requires a valid sheet_name.')
            if not isinstance(load.get('cell'), str) or not re.fullmatch(r'\$?[A-Za-z]{1,3}\$?[1-9][0-9]{0,6}', load['cell']):
                raise ValueError(f'Existing-sheet load for query {q["id"]} requires a valid A1 cell.')
        elif mode == 'new_sheet':
            if 'cell' in load or ('sheet_name' in load and (not isinstance(load['sheet_name'],str) or not (1 <= len(load['sheet_name']) <= 31))):
                raise ValueError(f'New-sheet load for query {q["id"]} has invalid destination fields.')
        elif set(load) & {'sheet_name','cell','table_name'}:
            raise ValueError(f'Load mode for query {q["id"]} does not accept sheet/cell/table fields.')
        if 'table_name' in load and (not isinstance(load['table_name'],str) or not (1 <= len(load['table_name']) <= 255)):
            raise ValueError(f'Invalid table_name for query {q["id"]}.')
        if 'depends_on' in q and (not isinstance(q['depends_on'],list) or len(q['depends_on']) != len(set(q['depends_on'])) or not all(isinstance(x,str) and x for x in q['depends_on'])):
            raise ValueError(f'Invalid depends_on list for query {q["id"]}.')
    if 'path_mappings' in document:
        if not isinstance(document['path_mappings'], list): raise ValueError('Profile path_mappings must be an array.')
        for mapping in document['path_mappings']:
            if not isinstance(mapping,dict) or set(mapping) != {'key','from','to','enabled'} or not isinstance(mapping['key'],str) or not mapping['key'] or not isinstance(mapping['from'],str) or not isinstance(mapping['to'],str) or not isinstance(mapping['enabled'],bool):
                raise ValueError('Invalid profile path mapping.')
    if 'environment_mappings' in document:
        if not isinstance(document['environment_mappings'],list):
            raise ValueError('Profile environment_mappings must be an array.')
        for mapping in document['environment_mappings']:
            if (not isinstance(mapping,dict) or set(mapping)!={'key','from','to','enabled'}
                    or not isinstance(mapping['key'],str) or not mapping['key']
                    or not isinstance(mapping['from'],str) or not mapping['from']
                    or not isinstance(mapping['to'],str) or not mapping['to']
                    or not isinstance(mapping['enabled'],bool)):
                raise ValueError('Invalid profile environment mapping.')


def save_profile(document, path):
    validate_profile(document)
    Path(path).write_text(json.dumps(document, ensure_ascii=False, indent=2) + '\n', encoding='utf-8')


def migrate_profile(document):
    """Migrate the unversioned/v0 profile shape to v1 without guessing fields.

    The legacy layout is the v1 profile document without ``schema_version``
    (or with it explicitly set to 0). All other structure is validated under
    the v1 rules. The input mapping is never mutated.
    """
    if not isinstance(document, dict):
        raise ValueError('Profile must be a JSON object.')
    version = document.get('schema_version', 0)
    if isinstance(version, bool) or not isinstance(version, int):
        raise ValueError('Unsupported or invalid profile schema_version (expected 0 or 1).')
    if version not in (0, 1):
        raise ValueError(f'Unsupported profile schema_version {version}; supported versions are 0 and 1.')
    migrated = dict(document)
    migrated['schema_version'] = 1
    validate_profile(migrated)
    return migrated


def load_profile(path):
    def unique_object(pairs):
        result={}
        for key,value in pairs:
            if key in result:raise ValueError(f'duplicate profile property: {key}')
            result[key]=value
        return result
    try: document = json.loads(Path(path).read_text(encoding='utf-8'),object_pairs_hook=unique_object)
    except (OSError, json.JSONDecodeError, ValueError) as exc: raise ValueError(f'Cannot load profile: {exc}') from exc
    try:
        return migrate_profile(document)
    except ValueError as exc:
        raise ValueError(f'Cannot load profile: {exc}') from exc


def build_import_profile(profile_name, workbook_path, sources, stable_ids, *, base_dir,
                         existing_query_policy='replace', backup=True, dry_run=False):
    base=Path(base_dir).resolve()
    def portable(value):
        path=Path(value).resolve()
        try:return Path(os.path.relpath(path,base)).as_posix()
        except ValueError:return str(path)
    if len(sources)!=len(stable_ids):raise ValueError('Every query source must have a stable row ID.')
    queries=[]
    for source, stable_id in zip(sources, stable_ids):
        load={'mode':source.load_mode.value}
        for key in ('sheet_name','cell','table_name'):
            value=getattr(source,key,None)
            if value is not None:
                load[key]=value
        persisted_path=source.profile_source_path or source.path
        queries.append({'id':str(stable_id),'source_path':portable(persisted_path),'query_name':source.name,
                        'enabled':bool(source.enabled),'order':int(source.order),'load':load})
    document={'schema_version':1,'profile_name':str(profile_name),'workbook':{'path':portable(workbook_path)},
              'options':{'existing_query_policy':existing_query_policy,'backup':bool(backup),'dry_run':bool(dry_run),
                         'refresh':'none','import_order':'manual','staged_output':True},
              'queries':queries}
    validate_profile(document)
    return document


def _resolve_profile_path(value, profile_dir):
    candidate=Path(value).expanduser()
    return candidate.resolve() if candidate.is_absolute() else (profile_dir/candidate).resolve()


def preview_profile_path_mappings(document, profile_path):
    """Preview anchored source-root mappings without changing profile or files."""
    validate_profile(document)
    base=Path(profile_path).resolve().parent
    mappings=[]
    for mapping in document.get('path_mappings',[]):
        if not mapping['enabled']:
            continue
        mappings.append((mapping['key'],_resolve_profile_path(mapping['from'],base),
                         _resolve_profile_path(mapping['to'],base)))
    rows=[]
    for query in document['queries']:
        source=_resolve_profile_path(query['source_path'],base)
        matches=[]
        for key,from_root,to_root in mappings:
            try:
                relative=source.relative_to(from_root)
            except ValueError:
                continue
            matches.append((len(from_root.parts),key,to_root/relative))
        if not matches:
            target=source;key=None;ambiguous=False
        else:
            specificity=max(entry[0] for entry in matches)
            best=[entry for entry in matches if entry[0]==specificity]
            unique_targets={str(entry[2]).casefold() for entry in best}
            ambiguous=len(unique_targets)>1
            if ambiguous:
                target=source;key=None
            else:
                _,key,target=sorted(best,key=lambda entry:entry[1].casefold())[0]
        rows.append(PathMappingPreview(query['id'],str(source),str(target),key,
                                      target!=source, target.is_file(),ambiguous))
    return tuple(rows)


def preview_profile_environment_mappings(document, profile_path):
    """Preview exact M source-literal edits; this function never writes files."""
    validate_profile(document)
    mappings=[mapping for mapping in document.get('environment_mappings',[]) if mapping['enabled']]
    if not mappings:
        return ()
    base=Path(profile_path).resolve().parent
    previews=[]
    for query in document['queries']:
        if not query['enabled']:
            continue
        source=_resolve_profile_path(query['source_path'],base)
        try:raw=source.read_bytes()
        except OSError as exc:raise ValueError(f'Cannot read source for environment mapping: {source.name}') from exc
        try:formula=raw.decode('utf-8-sig')
        except UnicodeDecodeError as exc:raise ValueError(f'Source {source.name!r} is not valid UTF-8.') from exc
        result=preview_m_source_mapping(formula,mappings)
        previews.append(ProfileEnvironmentMappingPreview(query['id'],query['query_name'],str(source),
            hashlib.sha256(raw).hexdigest(),formula,result.mapped,result.changes,result.warnings))
    return tuple(previews)


def materialize_profile_environment_mapping_copies(previews, *, staging_parent=None):
    """Write changed formulas into an isolated temporary directory after UI approval."""
    changed=[preview for preview in previews if preview.changed]
    if not changed:
        return {},None
    root=Path(tempfile.mkdtemp(prefix='pq-map-',dir=staging_parent))
    mapped={}
    try:
        for index,preview in enumerate(changed,1):
            source=Path(preview.source_path)
            raw=source.read_bytes()
            if hashlib.sha256(raw).hexdigest()!=preview.source_sha256:
                raise RuntimeError(f'Source {source.name!r} changed after preview; reload and review the mapping again.')
            bom=raw.startswith(b'\xef\xbb\xbf')
            destination=root/f'{index:03d}_{source.name}'
            data=(b'\xef\xbb\xbf' if bom else b'')+preview.mapped_formula.encode('utf-8')
            with destination.open('xb') as stream:stream.write(data)
            mapped[preview.query_id]=destination
        return mapped,root
    except Exception:
        import shutil
        shutil.rmtree(root,ignore_errors=True)
        raise


def profile_to_sources(document, profile_path, *, supported_load_modes=('definition_only',),
                       apply_path_mappings=False, apply_environment_mappings=False):
    validate_profile(document)
    options=document['options']
    if options.get('refresh','none')!='none':raise ValueError('This release does not support refresh settings saved in profiles.')
    if options.get('existing_query_policy')=='review_each':raise ValueError('This release cannot apply a review-each profile policy.')
    if options.get('import_order','manual')!='manual':raise ValueError('This release supports manual profile order only.')
    if options.get('staged_output',True) is False:raise ValueError('This release always stages workbook changes; the profile disables staging.')
    has_enabled_mappings=any(mapping.get('enabled') for mapping in document.get('path_mappings',[]))
    if has_enabled_mappings and not apply_path_mappings:
        raise ValueError('Enabled path mappings require preview and approve action before profile loading.')
    has_enabled_environment_mappings=any(mapping.get('enabled') for mapping in document.get('environment_mappings',[]))
    if has_enabled_environment_mappings and not apply_environment_mappings:
        raise ValueError('Enabled M-code environment mappings require preview and approve action before profile loading.')
    base=Path(profile_path).resolve().parent
    mapped_paths={}
    if has_enabled_mappings:
        previews=preview_profile_path_mappings(document,profile_path)
        if any(row.ambiguous for row in previews):
            raise ValueError('Path mappings are ambiguous for one or more query sources; resolve the conflict in the profile.')
        mapped_paths={row.query_id:Path(row.mapped_path) for row in previews}
    sources=[];stable_ids=[]
    for row in document['queries']:
        if row['load']['mode'] not in set(supported_load_modes):raise ValueError(f'Load mode {row["load"]["mode"]!r} has not passed this Office build capability probe.')
        if row.get('depends_on'):raise ValueError('Dependency overrides in profiles are not supported by this release.')
        policy=row.get('existing_policy')
        if policy and policy != options['existing_query_policy']:
            raise ValueError('Per-query policies must match the profile-wide existing query policy.')
        source_path=mapped_paths.get(row['id'],_resolve_profile_path(row['source_path'],base))
        load=row['load']
        sources.append(QuerySource(source_path,row['query_name'],row['order'],row['enabled'],LoadMode(load['mode']),
                                   order_automatic=False,sheet_name=load.get('sheet_name'),cell=load.get('cell'),
                                   table_name=load.get('table_name')))
        stable_ids.append(row['id'])
    return _resolve_profile_path(document['workbook']['path'],base),sources,stable_ids
