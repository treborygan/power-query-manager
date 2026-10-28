"""Safe extraction for untrusted Power Query ZIP inputs."""
from pathlib import Path, PurePosixPath
import os
import re
import shutil
import stat
import tempfile
import zipfile
import json
import hashlib
import re
from dataclasses import replace
from pq_installer.core.validation import validate_query_name
from pq_installer.core.dependencies import analyze_query_references,find_dependency_cycles


def _unique_json_object(pairs):
    result={}
    for key,value in pairs:
        if key in result:raise ValueError(f'duplicate package manifest property: {key}')
        result[key]=value
    return result


def validate_query_package_manifest(root):
    """Validate a versioned package manifest and hashes after safe extraction."""
    root=Path(root)
    manifest_path=root/'pqm-package.json'
    candidates=list(root.rglob('pqm-package.json'))
    if not candidates:return None
    if candidates != [manifest_path]:raise ValueError('Versioned query package manifest must appear once at the package root.')
    try:manifest=json.loads(manifest_path.read_text(encoding='utf-8'),object_pairs_hook=_unique_json_object)
    except (OSError,UnicodeError,json.JSONDecodeError,ValueError) as exc:
        raise ValueError(f'Invalid Power Query package manifest: {exc}') from exc
    if not isinstance(manifest,dict) or isinstance(manifest.get('schema_version'),bool):
        raise ValueError('Unsupported or invalid Power Query package manifest.')
    version=manifest.get('schema_version')
    expected_top=({'schema_version','package_type','queries'} if version==1 else
                  {'schema_version','package_type','queries','dependency_cycles'} if version==2 else set())
    if (set(manifest)!=expected_top
            or manifest.get('package_type')!='power-query-manager-query-package'):
        raise ValueError('Unsupported or invalid Power Query package manifest.')
    rows=manifest.get('queries')
    if not isinstance(rows,list) or not rows:raise ValueError('Power Query package manifest must contain query entries.')
    names=set();files=set();orders=set()
    entry_names=[];formulas={}
    for row in rows:
        expected_row=({'query_name','file','order','sha256'} if version==1 else
                      {'query_name','file','order','sha256','dependencies','dependency_status'})
        if not isinstance(row,dict) or set(row)!=expected_row:
            raise ValueError('Invalid query package manifest entry.')
        name=row['query_name'];relative=row['file'];order=row['order'];digest=row['sha256']
        try:
            if not isinstance(name,str):raise ValueError('query name must be text')
            validate_query_name(name)
        except (TypeError,ValueError) as exc:raise ValueError(f'Invalid query name in package manifest: {exc}') from exc
        if not isinstance(relative,str) or not relative or '\\' in relative:
            raise ValueError('Query package file path must be a safe relative POSIX path.')
        parts=PurePosixPath(relative).parts
        if PurePosixPath(relative).is_absolute() or not parts or any(part in {'','.','..'} for part in parts) or ':' in relative:
            raise ValueError('Query package file path is unsafe.')
        if PurePosixPath(relative).suffix.casefold() not in {'.m','.pq'}:
            raise ValueError('Query package entries must reference .m or .pq files.')
        if not isinstance(order,int) or isinstance(order,bool) or order<0 or order in orders:
            raise ValueError('Query package order values must be unique non-negative integers.')
        if not isinstance(digest,str) or not re.fullmatch(r'[0-9a-fA-F]{64}',digest):
            raise ValueError('Query package entry has an invalid SHA-256 value.')
        if name.casefold() in names:raise ValueError('Query package contains duplicate query names.')
        names.add(name.casefold());files.add(relative.casefold());orders.add(order);entry_names.append(name)
        target=root.joinpath(*parts)
        if not target.is_file():raise ValueError(f'Query package file is missing: {relative}.')
        actual=hashlib.sha256(target.read_bytes()).hexdigest()
        if actual.casefold()!=digest.casefold():raise ValueError(f'Query package hash mismatch for {relative}.')
        if version==2:
            dependencies=row['dependencies'];dependency_status=row['dependency_status']
            if (not isinstance(dependencies,list) or not all(isinstance(dep,str) and dep for dep in dependencies)
                    or len({dep.casefold() for dep in dependencies})!=len(dependencies)
                    or not isinstance(dependency_status,str) or dependency_status not in {'known','ambiguous'}):
                raise ValueError('Invalid query package dependency metadata.')
            try:formula=target.read_bytes().decode('utf-8-sig')
            except UnicodeDecodeError as exc:raise ValueError(f'Invalid UTF-8 query source in package: {relative}.') from exc
            formulas[name]=formula
    actual_files={p.relative_to(root).as_posix().casefold() for p in root.rglob('*') if p.is_file() and p.suffix.casefold() in {'.m','.pq'}}
    if actual_files!=files:raise ValueError('Query package manifest file list does not match its source files.')
    if version==2:
        for row in rows:
            if any(dep.casefold() not in names for dep in row['dependencies']):
                raise ValueError('Query package dependency metadata references a missing query.')
            analysis=analyze_query_references(formulas[row['query_name']],entry_names)
            if row['dependency_status']=='ambiguous':
                if not analysis.ambiguous or row['dependencies']:
                    raise ValueError('Query package dependency metadata does not match its M source.')
            elif analysis.ambiguous or row['dependencies']!=list(analysis.dependencies):
                raise ValueError('Query package dependency metadata does not match its M source.')
        graph={row['query_name']:(row['dependencies'] if row['dependency_status']=='known' else []) for row in rows}
        expected_cycles=[list(cycle) for cycle in find_dependency_cycles(graph)]
        if manifest.get('dependency_cycles')!=expected_cycles:
            raise ValueError('Query package dependency cycle metadata does not match its M sources.')
    return manifest


def validate_deployment_bundle_manifest(root):
    """Validate a portable profile/source deployment bundle after safe extraction."""
    root=Path(root)
    manifest_path=root/'pqm-deployment.json'
    candidates=list(root.rglob('pqm-deployment.json'))
    if not candidates:return None
    if candidates != [manifest_path] or (root/'pqm-package.json').exists():
        raise ValueError('Deployment bundle manifest must appear once at the root and cannot share a query-package manifest.')
    try:
        manifest=json.loads(manifest_path.read_text(encoding='utf-8'),object_pairs_hook=_unique_json_object)
    except (OSError,UnicodeError,json.JSONDecodeError,ValueError) as exc:
        raise ValueError(f'Invalid deployment bundle manifest: {exc}') from exc
    if (not isinstance(manifest,dict) or isinstance(manifest.get('schema_version'),bool)
            or set(manifest)!={'schema_version','bundle_type','profile','sources'}
            or manifest.get('schema_version')!=1
            or manifest.get('bundle_type')!='power-query-manager-deployment-bundle'):
        raise ValueError('Unsupported or invalid deployment bundle manifest.')
    from .manager import validate_profile
    profile=manifest['profile']
    validate_profile(profile)
    rows=manifest['sources']
    if not isinstance(rows,list):raise ValueError('Deployment bundle sources must be an array.')
    queries={row['id']:row for row in profile['queries']}
    by_id={};files=set()
    for row in rows:
        if not isinstance(row,dict) or set(row)!={'query_id','file','sha256'}:
            raise ValueError('Invalid deployment bundle source entry.')
        query_id=row['query_id'];relative=row['file'];digest=row['sha256']
        if not isinstance(query_id,str) or query_id not in queries or query_id in by_id:
            raise ValueError('Deployment bundle source references a duplicate or missing profile query.')
        if not isinstance(relative,str) or not relative or '\\' in relative:
            raise ValueError('Deployment bundle source path must be a safe relative POSIX path.')
        parts=PurePosixPath(relative).parts
        if (PurePosixPath(relative).is_absolute() or not parts
                or any(part in {'','.','..'} for part in parts) or ':' in relative
                or PurePosixPath(relative).suffix.casefold() not in {'.m','.pq'}):
            raise ValueError('Deployment bundle source path is unsafe.')
        if not isinstance(digest,str) or not re.fullmatch(r'[0-9a-fA-F]{64}',digest):
            raise ValueError('Deployment bundle source has an invalid SHA-256 value.')
        if relative.casefold() in files:
            raise ValueError('Deployment bundle contains duplicate source paths.')
        if queries[query_id]['source_path']!=relative:
            raise ValueError('Deployment bundle profile source path does not match its manifest.')
        target=root.joinpath(*parts).resolve()
        if root.resolve() not in target.parents or not target.is_file():
            raise ValueError(f'Deployment bundle source is missing or outside the bundle: {relative}.')
        actual=hashlib.sha256(target.read_bytes()).hexdigest()
        if actual.casefold()!=digest.casefold():raise ValueError(f'Deployment bundle hash mismatch for {relative}.')
        by_id[query_id]=row;files.add(relative.casefold())
    for query in profile['queries']:
        relative=query['source_path']
        parts=PurePosixPath(relative).parts if isinstance(relative,str) and '\\' not in relative else ()
        if (not parts or PurePosixPath(relative).is_absolute()
                or any(part in {'','.','..'} for part in parts) or ':' in relative
                or PurePosixPath(relative).suffix.casefold() not in {'.m','.pq'}):
            raise ValueError('Deployment bundle profile contains an unsafe source path.')
        source=root.joinpath(*parts)
        if query['id'] in by_id:
            continue
        if query['enabled'] or source.exists():
            raise ValueError(f'Deployment bundle source is not hash-verified: {relative}.')
    actual_files={path.relative_to(root).as_posix().casefold() for path in root.rglob('*')
                  if path.is_file() and path.suffix.casefold() in {'.m','.pq'}}
    if actual_files!=files:raise ValueError('Deployment bundle file list does not match its source entries.')
    return manifest


def migrate_query_package_manifest(root):
    """Upgrade a validated v1 manifest to the current v2 dependency form in memory."""
    root=Path(root)
    manifest=validate_query_package_manifest(root)
    if manifest is None or manifest['schema_version']==2:return manifest
    names=[row['query_name'] for row in manifest['queries']]
    rows=[];graph={}
    for row in manifest['queries']:
        path=root.joinpath(*PurePosixPath(row['file']).parts)
        try:
            formula=path.read_bytes().decode('utf-8-sig')
            analysis=analyze_query_references(formula,names)
            ambiguous=analysis.ambiguous
            dependencies=[] if ambiguous else list(analysis.dependencies)
        except (OSError,UnicodeDecodeError):
            ambiguous=True;dependencies=[]
        migrated={**row,'dependencies':dependencies,
                  'dependency_status':'ambiguous' if ambiguous else 'known'}
        rows.append(migrated);graph[row['query_name']]=dependencies if not ambiguous else []
    return {'schema_version':2,'package_type':manifest['package_type'],'queries':rows,
            'dependency_cycles':[list(cycle) for cycle in find_dependency_cycles(graph)]}


def apply_query_package_manifest(sources,root):
    manifest=migrate_query_package_manifest(root)
    if manifest is None:return list(sources)
    root=Path(root).resolve();by_path={}
    for row in manifest['queries']:
        path=(root.joinpath(*PurePosixPath(row['file']).parts)).resolve()
        by_path[path]=row
    if len(sources)!=len(by_path):raise ValueError('Query package source count does not match its manifest.')
    result=[]
    for source in sources:
        try:row=by_path[Path(source.path).resolve()]
        except KeyError as exc:raise ValueError(f'Source is missing from the package manifest: {source.path}.') from exc
        result.append(replace(source,name=row['query_name'],order=row['order'],order_from_prefix=False,order_automatic=False))
    return sorted(result,key=lambda source:(source.order,source.path.name.casefold(),str(source.path).casefold()))


def extract_query_zip(archive, *, staging_parent=None, max_files=256,
                      max_uncompressed=25 * 1024 * 1024, max_ratio=200,
                      max_archive_bytes=32 * 1024 * 1024):
    archive = Path(archive)
    if (isinstance(max_archive_bytes, bool) or not isinstance(max_archive_bytes, int)
            or max_archive_bytes < 1):
        raise ValueError('Query ZIP archive-size limit must be a positive whole number of bytes.')
    try:
        archive_stream = archive.open('rb')
    except OSError as exc:
        raise ValueError(f'Invalid query ZIP: {exc}') from exc
    try:
        opened_stat = os.fstat(archive_stream.fileno())
        if not stat.S_ISREG(opened_stat.st_mode) or opened_stat.st_size == 0:
            raise ValueError('Query ZIP is missing, empty, or not a regular file.')
        if opened_stat.st_size > max_archive_bytes:
            raise ValueError('Query ZIP exceeds the configured archive size limit.')
        zf = zipfile.ZipFile(archive_stream)
    except (OSError, zipfile.BadZipFile) as exc:
        archive_stream.close()
        raise ValueError(f'Invalid query ZIP: {exc}') from exc
    except Exception:
        archive_stream.close()
        raise

    try:
        # Windows may report a temporary directory using a short/aliased path,
        # while Path.resolve() canonicalizes child paths to the long spelling.
        # Compare canonical roots on both sides; do not relax traversal guards.
        root = Path(tempfile.mkdtemp(prefix='pq-import-', dir=staging_parent)).resolve()
    except Exception:
        zf.close()
        archive_stream.close()
        raise
    names = set(); total = 0; files = 0
    try:
        with zf:
            infos = zf.infolist()
            if len(infos) > max_files:
                raise ValueError('Query ZIP contains too many entries.')
            for info in infos:
                raw_name = info.filename.replace('\\', '/')
                if not raw_name or raw_name.startswith('/') or raw_name.startswith('//') or re.match(r'^[A-Za-z]:', raw_name):
                    raise ValueError(f'Query ZIP contains unsafe path: {info.filename!r}.')
                parts = raw_name.split('/')
                if any(part == '..' for part in parts):
                    raise ValueError(f'Query ZIP contains unsafe path: {info.filename!r}.')
                normalized = '/'.join(part for part in parts if part not in ('', '.'))
                if not normalized:
                    continue
                if any(':' in part or part.rstrip(' .') != part or '\x00' in part for part in PurePosixPath(normalized).parts):
                    raise ValueError(f'Query ZIP contains unsafe path: {info.filename!r}.')
                reserved={'con','prn','aux','nul',*(f'com{i}' for i in range(1,10)),*(f'lpt{i}' for i in range(1,10))}
                if any(part.split('.')[0].casefold() in reserved for part in PurePosixPath(normalized).parts):
                    raise ValueError(f'Query ZIP contains a reserved Windows path: {normalized}.')
                key = normalized.rstrip('/').casefold()
                if key in names:
                    raise ValueError(f'Query ZIP contains duplicate normalized path: {normalized}.')
                names.add(key)
                unix_mode = (info.external_attr >> 16) & 0xFFFF
                if stat.S_ISLNK(unix_mode):
                    raise ValueError(f'Query ZIP contains a symbolic link: {normalized}.')
                target = root.joinpath(*PurePosixPath(normalized).parts)
                if root not in target.resolve().parents and target.resolve() != root:
                    raise ValueError(f'Query ZIP path escapes extraction directory: {normalized}.')
                if info.is_dir():
                    target.mkdir(parents=True, exist_ok=True)
                    continue
                files += 1
                if files > max_files:
                    raise ValueError('Query ZIP contains too many files.')
                total += info.file_size
                if total > max_uncompressed:
                    raise ValueError('Query ZIP expands beyond the configured size limit.')
                if info.file_size > 1024 * 1024 and info.compress_size == 0:
                    raise ValueError('Query ZIP contains an invalid compressed-size entry.')
                if info.compress_size and info.file_size / info.compress_size > max_ratio:
                    raise ValueError(f'Query ZIP entry has a suspicious compression ratio: {normalized}.')
                ext = target.suffix.casefold()
                if ext not in {'.m', '.pq'} and target.name.casefold() not in {'manifest.json', 'query-manifest.json','pqm-package.json','pqm-deployment.json'}:
                    raise ValueError(f'Unsupported file in query ZIP: {normalized}. Only .m/.pq and manifests are accepted.')
                target.parent.mkdir(parents=True, exist_ok=True)
                actual=0
                with zf.open(info, 'r') as source, target.open('xb') as dest:
                    while True:
                        chunk=source.read(64 * 1024)
                        if not chunk: break
                        actual += len(chunk)
                        if actual > max_uncompressed or total - info.file_size + actual > max_uncompressed:
                            raise ValueError('Query ZIP expands beyond the configured size limit.')
                        dest.write(chunk)
                if actual != info.file_size:
                    raise ValueError(f'Query ZIP size metadata does not match extracted data: {normalized}.')
        if files == 0:
            raise ValueError('Query ZIP contains no files.')
        if not any(p.suffix.casefold() in {'.m', '.pq'} for p in root.rglob('*') if p.is_file()):
            raise ValueError('Query ZIP contains no .m or .pq source files.')
        validate_query_package_manifest(root)
        validate_deployment_bundle_manifest(root)
        return root
    except Exception:
        shutil.rmtree(root, ignore_errors=True)
        raise
    finally:
        archive_stream.close()
