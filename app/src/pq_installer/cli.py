"""Optional CLI for inspection, validation, planning, and staged deployment."""
from __future__ import annotations

import argparse
from contextlib import contextmanager
from concurrent.futures import ThreadPoolExecutor
from dataclasses import asdict
import json
import os
from pathlib import Path
import sys
import tempfile
import shutil
from difflib import unified_diff

from pq_installer.core.deployment import build_multi_workbook_dry_run, deploy_multi_workbook
from pq_installer.core.excel_capabilities import probe_excel_capabilities
from pq_installer.core.excel_inspector import ExcelQueryInspector
from pq_installer.core.excel_queries import ExcelQueryInstaller
from pq_installer.core.excel_refresh import ExcelQueryRefresher, RefreshOperationError
from pq_installer.core.manager import (build_import_profile, export_queries, export_query_package,
    export_deployment_bundle, load_profile, preview_profile_path_mappings, preview_profile_environment_mappings,
    materialize_profile_environment_mapping_copies, save_profile)
from pq_installer.core.models import LoadMode
from pq_installer.core.packages import (apply_query_package_manifest, extract_query_zip,
    validate_deployment_bundle_manifest)
from pq_installer.core.query_loader import discover_query_files
from pq_installer.core.redaction import safe_error
from pq_installer.core.reverse_dependencies import discover_scan_files, scan_reverse_dependencies
from pq_installer.core.error_log import ApprovalRequired,format_failure,record_failure


class MappingApprovalDeclined(RuntimeError):
    """A user explicitly declined reviewed source mapping changes."""


class MappingApprovalRequired(ApprovalRequired):
    """A noninteractive invocation lacks approval for reviewed source mappings."""


def _record_cli_failure(operation,error):
    # Persist category and exception type only; the normal CLI output retains
    # its detailed message without copying workbook data or M code into logs.
    record_failure(operation,error)


def _record_cli_history(event,details):
    """Record safe aggregate CLI activity in the user's local history store."""
    local_app_data=os.environ.get('LOCALAPPDATA')
    if not local_app_data:return
    try:
        from pq_installer.core.history import HistoryStore
        store=HistoryStore(Path(local_app_data)/'PowerQueryManager'/'history.sqlite3')
        store.record(event,details)
    except Exception as exc:
        print(f'Local history could not be saved: {safe_error(exc)}',file=sys.stderr)


def _plan_audit_details(bundle,profile):
    items=[item for plan in bundle.get('plans',()) for item in plan.get('items',())]
    actions=[item.get('action') for item in items]
    return {'action':'dry_run','status':'completed','comparison':'completed',
        'query_count':sum(bool(query.get('enabled',True)) for query in profile.get('queries',())),
        'workbook_count':len(bundle.get('plans',())),
        'added':actions.count('create'),'modified':actions.count('replace'),
        'unchanged':actions.count('skip'),'conflict':actions.count('inspect')}


def _deployment_audit_details(bundle,profile):
    results=bundle.get('results',())
    query_results=[query for workbook in results for query in workbook.get('results',())]
    states=[row.get('status') for row in results]
    status='declined' if 'cancelled' in states else ('completed' if not any(
        value in {'failed','blocked_preflight'} for value in states) else 'partial')
    query_states=[str(row.get('status','')).casefold() for row in query_results]
    return {'action':'apply','status':status,
        'query_count':sum(bool(query.get('enabled',True)) for query in profile.get('queries',())),
        'workbook_count':len(results),
        'created_count':query_states.count('created'),
        'replaced_count':query_states.count('replaced'),
        'skipped_count':query_states.count('skipped')}


def _make_parser():
    parser=argparse.ArgumentParser(prog='power-query-manager-cli',description='Inspect, validate, plan, and deploy Power Query definitions.')
    commands=parser.add_subparsers(dest='command',required=True)
    plan=commands.add_parser('plan',help='Create immutable plans for one or more workbooks; never apply them.')
    plan.add_argument('--dry-run',action='store_true',required=True,help='Required acknowledgement: this command only inspects and plans.')
    plan.add_argument('--profile',required=True,help='Versioned JSON import profile.')
    plan.add_argument('--workbook',action='append',required=True,dest='workbooks',help='Target .xlsx/.xlsm workbook; repeat for multiple targets.')
    plan.add_argument('--output',help='Write a JSON plan bundle to a new path; existing files are never overwritten.')
    plan.add_argument('--apply-path-mappings',action='store_true',help='Preview and approve enabled profile path mappings before planning.')
    plan.add_argument('--apply-environment-mappings',action='store_true',help='Review exact M-code diffs and use temporary mapped source copies in the dry-run plan.')
    plan.add_argument('--confirm',action='append',default=[],help='Supply one exact approval phrase for each requested review, in prompt order; may be repeated.')
    deploy=commands.add_parser('deploy',help='Deploy to multiple workbooks as separate staged transactions.')
    deploy.add_argument('--apply',action='store_true',required=True,help='Required acknowledgement that workbook files may be changed.')
    deploy.add_argument('--profile',required=True,help='Versioned JSON import profile.')
    deploy.add_argument('--workbook',action='append',required=True,dest='workbooks',help='Target .xlsx/.xlsm workbook; repeat for multiple targets.')
    deploy.add_argument('--output',help='Write the per-workbook report to a new path; existing files are never overwritten.')
    deploy.add_argument('--apply-path-mappings',action='store_true',help='Preview and approve enabled profile path mappings before deployment.')
    deploy.add_argument('--apply-environment-mappings',action='store_true',help='Review exact M-code diffs and deploy from temporary mapped source copies.')
    deploy.add_argument('--confirm',action='append',default=[],help='Supply exact approval phrases in review order (for example, --confirm DEPLOY); may be repeated.')
    bundle=commands.add_parser('bundle',help='Create a portable multi-workbook deployment bundle from a profile and its source files.')
    bundle.add_argument('--profile',required=True,help='Versioned JSON import profile.')
    bundle.add_argument('--output',required=True,help='New .pqmdeploy or .zip output path; existing files are never overwritten.')
    bundle.add_argument('--apply-path-mappings',action='store_true',help='Preview and approve enabled source-path mappings before embedding those sources.')
    bundle.add_argument('--confirm',action='append',default=[],help='Supply the exact source-mapping approval phrase after reviewing its preview.')
    inspect=commands.add_parser('inspect',help='Read query names, formulas, and observed load states from a workbook.')
    inspect.add_argument('--workbook',required=True,help='Target .xlsx/.xlsm workbook; opened read-only.')
    inspect.add_argument('--output',help='Write JSON to a new path; existing files are never overwritten.')
    find=commands.add_parser('dependency-scan',help='Find files that refer to selected workbooks/Access databases without opening or refreshing them.')
    find.add_argument('--target',action='append',required=True,dest='targets',help='Dependency target file to find incoming references to; repeat for multiple targets.')
    find.add_argument('--file',action='append',default=[],dest='files',help='File to inspect for outgoing references; repeat to select multiple files.')
    find.add_argument('--folder',action='append',default=[],dest='folders',help='Folder containing candidate Excel/Access files; repeat for multiple folders.')
    find.add_argument('--recursive',action='store_true',help='Include nested folders when discovering candidate files.')
    find.add_argument('--output',help='Save report as JSON to a NEW path; no existing report is overwritten.')
    validate=commands.add_parser('validate',help='Validate a versioned import profile and enabled source files offline.')
    validate.add_argument('--profile',required=True,help='Versioned JSON import profile.')
    export=commands.add_parser('export',help='Export all workbook query definitions to .pq files or a verified package.')
    export.add_argument('--workbook',required=True,help='Source .xlsx/.xlsm workbook; opened read-only.')
    export.add_argument('--output',required=True,help='New output folder for .pq files, or new .pqm.zip path with --package.')
    export.add_argument('--package',action='store_true',help='Write a versioned .pqm.zip package instead of individual .pq files.')
    refresh=commands.add_parser('refresh',help='Explicitly refresh selected/all workbook queries in a staged copy.')
    refresh.add_argument('--apply',action='store_true',required=True,help='Required acknowledgement that Excel may contact external sources and the workbook may be updated.')
    selection=refresh.add_mutually_exclusive_group(required=True)
    selection.add_argument('--all',action='store_true',help='Refresh all queries that have refreshable connections.')
    selection.add_argument('--query',action='append',dest='queries',help='Query name to refresh; repeat for multiple names.')
    refresh.add_argument('--workbook',required=True,help='Target .xlsx/.xlsm workbook.')
    refresh.add_argument('--timeout',type=float,default=300,help='Per-query timeout in seconds (1 to 7200; default: 300).')
    refresh.add_argument('--synchronous',action='store_true',help='Use Excel synchronous refresh mode; default is background mode.')
    refresh.add_argument('--confirm',action='append',default=[],help='Supply the exact REFRESH approval phrase for explicit non-interactive refresh.')
    package_import=commands.add_parser('import',help='Import a verified query ZIP to one workbook using definition-only loads.')
    import_mode=package_import.add_mutually_exclusive_group(required=True)
    import_mode.add_argument('--apply',action='store_true',help='Allow staged definition changes after typed confirmation.')
    import_mode.add_argument('--dry-run',action='store_true',help='Read-only package/workbook comparison; never writes the workbook.')
    package_import.add_argument('--package',required=True,help='Untrusted .zip/.pqm.zip input; safely extracted to a temporary folder.')
    package_import.add_argument('--workbook',required=True,help='Target .xlsx/.xlsm workbook.')
    package_import.add_argument('--policy',choices=('skip','replace'),default='skip',help='Existing-query policy (default: skip).')
    package_import.add_argument('--confirm',action='append',default=[],help='Supply the exact IMPORT approval phrase for an applying import.')
    return parser


def _write_json_output(path,document):
    rendered=json.dumps(document,ensure_ascii=False,indent=2,sort_keys=True)+'\n'
    if path:
        written=_write_new_file(path,rendered)
        print(f'Wrote JSON output: {written}')
    else:
        sys.stdout.write(rendered)


def _validate_profile_sources(profile_path):
    profile=load_profile(profile_path)
    base=Path(profile_path).expanduser().resolve().parent
    rows=[]
    for query in profile['queries']:
        if not query['enabled']:
            rows.append({'id':query['id'],'status':'disabled'})
            continue
        source=Path(query['source_path']).expanduser()
        if not source.is_absolute():source=base/source
        if not source.is_file():raise ValueError(f"Enabled source file is missing for query {query['query_name']!r}.")
        if source.suffix.casefold() not in {'.m','.pq'}:raise ValueError(f"Enabled source file has unsupported extension for query {query['query_name']!r}.")
        rows.append({'id':query['id'],'status':'valid','source_path':str(source.resolve()),'bytes':source.stat().st_size})
    return {'schema_version':1,'valid':True,'profile_name':profile['profile_name'],'queries':rows}


@contextmanager
def _review_cli_environment_mappings(profile,profile_path,*,enabled,input_func,interactive,
                                    apply_path_mappings=False):
    path_mappings=[row for row in profile.get('path_mappings',[]) if row.get('enabled')]
    path_previews=()
    working_profile=profile
    if path_mappings:
        if not apply_path_mappings:
            raise ValueError('Profile enables source path mappings. Re-run with --apply-path-mappings to review them.')
        path_previews=preview_profile_path_mappings(profile,profile_path)
        if any(row.ambiguous for row in path_previews):
            raise ValueError('Path mappings are ambiguous for one or more query sources; resolve the conflict first.')
        mapped_paths={row.query_id:row.mapped_path for row in path_previews}
        working_profile=dict(profile)
        working_profile['queries']=[dict(row,source_path=mapped_paths[row['id']])
                                    for row in profile['queries']]
        working_profile['path_mappings']=[]
    mappings=[row for row in working_profile.get('environment_mappings',[]) if row.get('enabled')]
    if not mappings:
        yield profile_path,(),path_previews
        return
    if not enabled:
        raise ValueError('Profile enables M-code environment mappings. Re-run with --apply-environment-mappings to review exact diffs.')
    previews=preview_profile_environment_mappings(working_profile,profile_path)
    if any(preview.warnings for preview in previews):
        details='; '.join(warning for preview in previews for warning in preview.warnings)
        raise ValueError(f'M-code mapping preview is ambiguous; resolve it in the profile before continuing. {details}')
    print('M-code environment mapping diff preview:')
    for preview in previews:
        if not preview.changed:
            print(f'  {preview.query_id}: unchanged ({Path(preview.source_path).name})')
            continue
        print(f'  {preview.query_id}: {Path(preview.source_path).name}')
        diff=unified_diff(preview.original_formula.splitlines(True),preview.mapped_formula.splitlines(True),
                          fromfile='original/'+Path(preview.source_path).name,
                          tofile='mapped/'+Path(preview.source_path).name)
        sys.stdout.writelines(diff)
    if not interactive:
        raise ApprovalRequired(
            'M-code mapping approval is required; pass --confirm "APPLY M-CODE MAPPINGS" after reviewing the diff.')
    if input_func('Type APPLY M-CODE MAPPINGS to continue: ').strip()!='APPLY M-CODE MAPPINGS':
        raise MappingApprovalDeclined('M-code mapping review was declined; no files or workbooks were changed.')
    mapped,root=materialize_profile_environment_mapping_copies(previews)
    if root is None:
        root=tempfile.mkdtemp(prefix='pq-map-profile-')
    try:
        prepared=dict(working_profile)
        profile_base=Path(profile_path).expanduser().resolve().parent
        prepared_queries=[]
        for row in working_profile['queries']:
            source=Path(row['source_path']).expanduser()
            if not source.is_absolute():source=profile_base/source
            prepared_queries.append(dict(row,source_path=str(mapped.get(row['id'],source.resolve()))))
        prepared['queries']=prepared_queries
        prepared['path_mappings']=[]
        prepared['environment_mappings']=[]
        effective=Path(root)/'approved-profile.json'
        effective.write_text(json.dumps(prepared,ensure_ascii=False,indent=2),encoding='utf-8')
        yield str(effective),previews,path_previews
    finally:
        shutil.rmtree(root,ignore_errors=True)


def _mapping_preview_json(previews):
    return [{'query_id':row.query_id,'query_name':row.query_name,'source_path':row.source_path,
             'source_sha256':row.source_sha256,'changed':row.changed,'changes':[asdict(change) for change in row.changes],
             'warnings':list(row.warnings)} for row in previews]


def _path_mapping_preview_json(previews):
    return [{'query_id':row.query_id,'source_path':row.source_path,
             'mapped_path':row.mapped_path,'mapping_key':row.mapping_key,'changed':row.changed,
             'target_exists':row.target_exists,'ambiguous':row.ambiguous} for row in previews]


@contextmanager
def _deployment_profile_input(profile_path):
    """Yield a JSON profile path, safely unpacking a portable deployment bundle when needed."""
    path=Path(profile_path).expanduser()
    if path.suffix.casefold() not in {'.zip','.pqmdeploy'}:
        yield str(path)
        return
    root=extract_query_zip(path)
    try:
        manifest=validate_deployment_bundle_manifest(root)
        if manifest is None:
            raise ValueError('Profile ZIP is not a verified Power Query Manager deployment bundle.')
        materialized=root/'deployment-profile.json'
        save_profile(manifest['profile'],materialized)
        yield str(materialized)
    finally:
        shutil.rmtree(root,ignore_errors=True)


def _build_bundle(profile_path,workbooks,*,inspector_factory=None,capability_probe=None,apply_path_mappings=False,
                  environment_previews=(),path_previews=None):
    profile=load_profile(profile_path)
    enabled_modes={row['load']['mode'] for row in profile['queries'] if row['enabled']}
    supported=('definition_only',)
    if 'connection_only' in enabled_modes:
        probe=(capability_probe or probe_excel_capabilities)()
        supported=probe.supported_load_modes
    preview=() if path_previews is None else path_previews
    if path_previews is None and any(mapping['enabled'] for mapping in profile.get('path_mappings',[])):
        preview=preview_profile_path_mappings(profile,profile_path)
        if not apply_path_mappings:
            raise ValueError('Profile enables source path mappings. Re-run with --apply-path-mappings to include the shown preview in the dry-run bundle.')
        if any(row.ambiguous for row in preview):
            raise ValueError('Profile path mappings contain an ambiguous source target; resolve the mapping conflict first.')
    factory=inspector_factory or ExcelQueryInspector
    def inventory_provider(path):
        return factory().inspect(path)
    plans=build_multi_workbook_dry_run(profile,profile_path,workbooks,
        inventory_provider=inventory_provider,supported_load_modes=supported,
        apply_path_mappings=apply_path_mappings)
    return {'schema_version':1,'bundle_type':'multi-workbook-dry-run',
            'environment_mapping_preview':_mapping_preview_json(environment_previews),
            'path_mapping_preview':_path_mapping_preview_json(preview),
            'plans':list(plans)}


def _write_new_file(path,text):
    destination=Path(path).expanduser().resolve()
    destination.parent.mkdir(parents=True,exist_ok=True)
    temp_path=None
    try:
        with tempfile.NamedTemporaryFile(prefix='.pqm-plan-',suffix='.tmp',dir=destination.parent,delete=False,mode='w',encoding='utf-8',newline='\n') as stream:
            temp_path=Path(stream.name);stream.write(text)
        os.link(temp_path,destination)
        return destination
    finally:
        if temp_path is not None:
            try:temp_path.unlink()
            except FileNotFoundError:pass


def _prepare_output_path(path):
    destination=Path(path).expanduser().resolve()
    destination.parent.mkdir(parents=True,exist_ok=True)
    if destination.exists():raise FileExistsError(f'Refusing to overwrite existing output: {destination}')
    with tempfile.NamedTemporaryFile(prefix='.pqm-report-check-',dir=destination.parent,delete=True):
        pass
    return destination


def _review_cli_path_mappings(profile,profile_path,*,enabled,input_func,interactive):
    mappings=profile.get('path_mappings',[])
    if not any(row.get('enabled') for row in mappings):return ()
    if not enabled:
        raise ValueError('Profile enables source path mappings. Re-run with --apply-path-mappings to review them.')
    preview=preview_profile_path_mappings(profile,profile_path)
    if any(row.ambiguous for row in preview):
        raise ValueError('Profile path mappings contain an ambiguous source target; resolve it before deployment.')
    print('Source path mapping preview:')
    for row in preview:
        print(f'  {row.query_id}: {row.source_path} -> {row.mapped_path}' if row.changed
              else f'  {row.query_id}: unchanged ({row.source_path})')
    if not interactive:
        raise MappingApprovalRequired(
            'Path mapping approval is required; pass --confirm "APPLY PATH MAPPINGS" after reviewing the preview.')
    if input_func('Type APPLY PATH MAPPINGS to continue: ').strip()!='APPLY PATH MAPPINGS':
        raise MappingApprovalDeclined('Path mapping review was declined; no files or workbooks were changed.')
    return preview


def _deploy_bundle(profile_path,workbooks,*,inspector_factory=None,installer_factory=None,
                   capability_probe=None,apply_path_mappings=False,input_func=None,interactive=False):
    profile=load_profile(profile_path)
    enabled_modes={row['load']['mode'] for row in profile['queries'] if row['enabled']}
    needs_probe=enabled_modes-{'definition_only','preserve_existing'}
    capabilities=(capability_probe or probe_excel_capabilities)() if needs_probe else None
    supported=(capabilities.supported_load_modes if capabilities is not None
               else ('definition_only','preserve_existing'))
    factory=inspector_factory or ExcelQueryInspector
    def inventory_provider(path):return factory().inspect(path)
    apply_factory=installer_factory or (lambda _path:ExcelQueryInstaller())
    def confirm(plans):
        print(f'Preflight complete for {len(plans)} workbook(s). Each workbook is committed independently after staged save/reopen verification.')
        has_evaluating_load=any(item['load']['mode'] in {'new_sheet','existing_sheet','data_model'}
            for plan in plans for item in plan['items'] if item['action']!='skip')
        if has_evaluating_load:
            print('Worksheet/Data Model loads evaluate M and may contact external data sources.')
        for plan in plans:
            print(f"  {plan['workbook_path']}")
            for item in plan['items']:print(f"    {item['action']}: {item['query_name']}")
        if not interactive:
            raise ApprovalRequired('Deployment approval is required; pass --confirm DEPLOY after reviewing the plan.')
        return input_func('Type DEPLOY to continue: ').strip()=='DEPLOY'
    results=deploy_multi_workbook(profile,profile_path,workbooks,
        inventory_provider=inventory_provider,supported_load_modes=supported,
        capabilities=capabilities,apply_path_mappings=apply_path_mappings,
        installer_factory=apply_factory,confirm_apply=confirm)
    return {'schema_version':1,'bundle_type':'multi-workbook-deployment-report',
            'results':[asdict(row) for row in results]}


def _import_query_package(package_path,workbook_path,*,policy='skip',dry_run=False,inspector_factory=None,
                          installer_factory=None,input_func=None,interactive=False):
    import shutil
    extracted=extract_query_zip(package_path)
    try:
        sources=discover_query_files([extracted],recursive=True)
        sources=apply_query_package_manifest(sources,extracted)
        if not sources:raise ValueError('Query package contains no importable .m/.pq sources.')
        for order,source in enumerate(sources):
            source.order=order
            source.load_mode=LoadMode.DEFINITION_ONLY
            source.enabled=True
        profile_path=extracted/'cli-import-profile.json'
        profile=build_import_profile('CLI package import',workbook_path,sources,
            [f'package-query-{index}' for index in range(len(sources))],base_dir=extracted,
            existing_query_policy=policy,backup=True,dry_run=dry_run)
        save_profile(profile,profile_path)
        factory=inspector_factory or ExcelQueryInspector
        def inventory_provider(path):return factory().inspect(path)
        apply_factory=installer_factory or (lambda _path:ExcelQueryInstaller())
        def confirm(plans):
            print('Package import preflight (definition-only; no query refresh will run):')
            for plan in plans:
                print(f"  {plan['workbook_path']}")
                for item in plan['items']:print(f"    {item['action']}: {item['query_name']}")
            if not interactive:
                raise ApprovalRequired('Package import approval is required; pass --confirm IMPORT after reviewing the plan.')
            return input_func('Type IMPORT to continue: ').strip()=='IMPORT'
        with ThreadPoolExecutor(max_workers=1,thread_name_prefix='pqm-cli-excel') as executor:
            if dry_run:
                plans=executor.submit(build_multi_workbook_dry_run,profile,str(profile_path),[workbook_path],
                    inventory_provider=inventory_provider,supported_load_modes=('definition_only','preserve_existing')).result()
                return {'schema_version':1,'bundle_type':'query-package-import-plan','package_name':Path(package_path).name,
                        'policy':policy,'plans':list(plans)}
            results=executor.submit(deploy_multi_workbook,profile,str(profile_path),[workbook_path],
                inventory_provider=inventory_provider,supported_load_modes=('definition_only','preserve_existing'),
                installer_factory=apply_factory,confirm_apply=confirm).result()
        return {'schema_version':1,'bundle_type':'query-package-import-report',
                'package_name':Path(package_path).name,'policy':policy,
                'results':[asdict(row) for row in results]}
    finally:
        shutil.rmtree(extracted,ignore_errors=True)


def _dispatch_parsed(args,parser,*,inspector_factory=None,installer_factory=None,
                     refresher_factory=None,capability_probe=None,input_func=None,interactive=False):
    if args.command=='dependency-scan':
        try:
            if not (args.files or args.folders):
                raise ValueError('Supply at least one --file or --folder to inspect.')
            for source in (*args.files,*args.folders):
                if not Path(source).exists():
                    raise ValueError('A selected scan path does not exist.')
            candidates=discover_scan_files((*args.files,*args.folders),recursive=args.recursive)
            report=scan_reverse_dependencies(args.targets,candidates)
            _write_json_output(args.output,{'schema_version':1,**asdict(report)})
            return 5 if report.issues or report.cancelled else 0
        except (ValueError,OSError) as exc:
            print(f'Dependency scan failed: {format_failure(exc)}',file=sys.stderr)
            return 2
    if args.command=='validate':
        try:
            _write_json_output(None,_validate_profile_sources(args.profile))
            return 0
        except (ValueError,OSError) as exc:
            _record_cli_failure('validate_profile',exc)
            print(f'Profile validation failed: {format_failure(exc)}',file=sys.stderr);return 2
    if args.command=='inspect':
        try:
            with ThreadPoolExecutor(max_workers=1,thread_name_prefix='pqm-cli-excel') as executor:
                inventory=executor.submit((inspector_factory or ExcelQueryInspector)().inspect,args.workbook).result()
            _write_json_output(args.output,{'schema_version':1,'workbook_path':str(Path(args.workbook).resolve()),
                'queries':[{'name':row.name,'formula':row.formula,'description':row.description,'load_state':row.load_state}
                           for row in inventory]})
            return 0
        except RuntimeError as exc:
            _record_cli_failure('inspect',exc)
            message=format_failure(exc);print(f'Excel/COM inspection failed: {message}',file=sys.stderr);return 4
        except (ValueError,OSError) as exc:
            _record_cli_failure('inspect',exc)
            print(f'Workbook inspection validation failed: {format_failure(exc)}',file=sys.stderr);return 2
    if args.command=='bundle':
        try:
            profile=load_profile(args.profile)
            path_previews=_review_cli_path_mappings(profile,args.profile,enabled=args.apply_path_mappings,
                                                    input_func=input_func,interactive=interactive)
            output=export_deployment_bundle(args.profile,args.output,
                apply_path_mappings=args.apply_path_mappings)
            _write_json_output(None,{'schema_version':1,'bundle_type':'multi-workbook-deployment',
                'output_path':str(output),'size_bytes':output.stat().st_size,
                'query_count':len(profile['queries']),'path_mapping_preview':_path_mapping_preview_json(path_previews)})
            return 0
        except MappingApprovalDeclined as exc:
            _record_cli_failure('bundle',exc);print(format_failure(exc),file=sys.stderr);return 3
        except MappingApprovalRequired as exc:
            _record_cli_failure('bundle',exc);print(f'Approval required: {format_failure(exc)}',file=sys.stderr);return 3
        except NotImplementedError as exc:
            _record_cli_failure('bundle',exc);print(f'Deployment bundle unsupported: {format_failure(exc)}',file=sys.stderr);return 3
        except (ValueError,OSError) as exc:
            _record_cli_failure('bundle',exc);print(f'Deployment bundle failed: {format_failure(exc)}',file=sys.stderr);return 2
    if args.command=='export':
        try:
            output=Path(args.output).expanduser().resolve()
            if args.package:
                if output.suffix.casefold()!='.zip':raise ValueError('Package output must use the .zip extension.')
                if output.exists():raise FileExistsError(f'Refusing to overwrite existing output: {output}')
            elif output.exists() and (not output.is_dir() or any(output.iterdir())):
                raise FileExistsError(f'Refusing to export into a non-empty output folder: {output}')
            with ThreadPoolExecutor(max_workers=1,thread_name_prefix='pqm-cli-excel') as executor:
                inventory=executor.submit((inspector_factory or ExcelQueryInspector)().inspect,args.workbook).result()
            if args.package:
                exported=export_query_package(inventory,output)
                document={'schema_version':1,'format':'pqm.zip','output_path':str(exported),'query_count':sum(row.formula is not None for row in inventory)}
            else:
                exported=export_queries(inventory,output)
                document={'schema_version':1,'format':'pq-files','output_folder':str(output),'files':[str(path) for path in exported],
                          'query_count':len(exported)}
            _write_json_output(None,document)
            return 0
        except RuntimeError as exc:
            _record_cli_failure('export',exc)
            print(f'Excel/COM export failed: {format_failure(exc)}',file=sys.stderr);return 4
        except (ValueError,OSError) as exc:
            _record_cli_failure('export',exc)
            print(f'Query export failed: {format_failure(exc)}',file=sys.stderr);return 2
    if args.command=='refresh':
        try:
            if not 1 <= args.timeout <= 7200:
                raise ValueError('Refresh timeout must be between 1 and 7200 seconds.')
            with ThreadPoolExecutor(max_workers=1,thread_name_prefix='pqm-cli-excel') as executor:
                inventory=executor.submit((inspector_factory or ExcelQueryInspector)().inspect,args.workbook).result()
                by_name={row.name.casefold():row.name for row in inventory}
                inventory_by_name={row.name.casefold():row for row in inventory}
                if args.all:
                    selected=None; shown=[row.name for row in inventory]
                else:
                    requested=args.queries or []
                    if len({name.casefold() for name in requested})!=len(requested):
                        raise ValueError('The selected query list contains duplicates.')
                    missing=[name for name in requested if name.casefold() not in by_name]
                    if missing:raise ValueError('Query not found in workbook inventory: '+', '.join(missing))
                    selected=tuple(by_name[name.casefold()] for name in requested);shown=list(selected)
                if not shown:raise ValueError('Workbook contains no queries to refresh.')
                print(f'Refresh selection for {args.workbook}: '+', '.join(
                    f'{name} [{inventory_by_name[name.casefold()].load_state}]' for name in shown))
                print('Excel may contact external data sources. A verified backup is created by default; refresh runs in a staged workbook.')
                if not interactive:
                    raise ApprovalRequired('Refresh approval is required; pass --confirm REFRESH after reviewing the selection.')
                if input_func('Type REFRESH to continue: ').strip()!='REFRESH':
                    print('Refresh was declined.',file=sys.stderr);return 3
                refresher=(refresher_factory or ExcelQueryRefresher)()
                results=executor.submit(refresher.refresh,args.workbook,selected,
                    timeout_seconds=args.timeout,asynchronous=not args.synchronous,make_backup=True).result()
            document={'schema_version':1,'workbook_path':str(Path(args.workbook).resolve()),
                'mode':'synchronous' if args.synchronous else 'background',
                'results':[asdict(row) for row in results]}
            _write_json_output(None,document)
            return 5 if any(row.status in {'failed','rolled_back'} for row in results) else 0
        except ApprovalRequired as exc:
            _record_cli_failure('refresh',exc)
            print(f'Approval required: {format_failure(exc)}',file=sys.stderr);return 3
        except RefreshOperationError as exc:
            _record_cli_failure('refresh',exc)
            document={'schema_version':1,'workbook_path':str(Path(args.workbook).resolve()),
                'error':format_failure(exc),'results':[asdict(row) for row in exc.results]}
            _write_json_output(None,document);return 5
        except RuntimeError as exc:
            _record_cli_failure('refresh',exc)
            print(f'Excel/COM refresh failed: {format_failure(exc)}',file=sys.stderr);return 4
        except (ValueError,OSError) as exc:
            _record_cli_failure('refresh',exc)
            print(f'Refresh validation failed: {format_failure(exc)}',file=sys.stderr);return 2
    if args.command=='import':
        try:
            report=_import_query_package(args.package,args.workbook,policy=args.policy,
                dry_run=args.dry_run,
                inspector_factory=inspector_factory,installer_factory=installer_factory,
                input_func=input_func,interactive=interactive)
            _write_json_output(None,report)
            if args.dry_run:return 0
            states=[row['status'] for row in report['results']]
            return 5 if any(state in {'failed','blocked_preflight'} for state in states) else (3 if 'cancelled' in states else 0)
        except MappingApprovalRequired as exc:
            _record_cli_failure('import',exc);print(f'Approval required: {format_failure(exc)}',file=sys.stderr);return 3
        except ApprovalRequired as exc:
            _record_cli_failure('import',exc);print(f'Approval required: {format_failure(exc)}',file=sys.stderr);return 3
        except NotImplementedError as exc:
            _record_cli_failure('import',exc)
            print(f'Package import unsupported: {format_failure(exc)}',file=sys.stderr);return 3
        except RuntimeError as exc:
            _record_cli_failure('import',exc)
            message=format_failure(exc)
            if any(marker in message.casefold() for marker in ('excel','com','inspection failed')):
                print(f'Excel/COM package import failed: {message}',file=sys.stderr);return 4
            print(f'Package import failed: {message}',file=sys.stderr);return 5
        except (ValueError,OSError) as exc:
            _record_cli_failure('import',exc)
            print(f'Package import validation failed: {format_failure(exc)}',file=sys.stderr);return 2
    if args.command=='deploy':
        try:
            output_path=_prepare_output_path(args.output) if args.output else None
            profile=load_profile(args.profile)
            path_previews=_review_cli_path_mappings(profile,args.profile,enabled=args.apply_path_mappings,
                                                    input_func=input_func,interactive=interactive)
            with _review_cli_environment_mappings(profile,args.profile,enabled=args.apply_environment_mappings,
                    input_func=input_func,interactive=interactive,
                    apply_path_mappings=args.apply_path_mappings) as (effective_profile,environment_previews,path_previews):
                with ThreadPoolExecutor(max_workers=1,thread_name_prefix='pqm-cli-excel') as executor:
                    bundle=executor.submit(_deploy_bundle,effective_profile,args.workbooks,
                        inspector_factory=inspector_factory,installer_factory=installer_factory,
                        capability_probe=capability_probe,apply_path_mappings=args.apply_path_mappings,
                        input_func=input_func,interactive=interactive).result()
                bundle['environment_mapping_preview']=_mapping_preview_json(environment_previews)
                bundle['path_mapping_preview']=_path_mapping_preview_json(path_previews or ())
            rendered=json.dumps(bundle,ensure_ascii=False,indent=2,sort_keys=True)+'\n'
            if output_path:
                try:
                    written=_write_new_file(output_path,rendered)
                    print(f'Wrote multi-workbook deployment report: {written}')
                except OSError:
                    sys.stdout.write(rendered)
                    raise
            else:sys.stdout.write(rendered)
            states=[row['status'] for row in bundle['results']]
            failures=[row for row in bundle['results'] if row['status'] in {'failed','blocked_preflight'}]
            if failures:
                if all(row['status']=='blocked_preflight' and any(token in row['message'].casefold() for token in ('excel','com','inspection failed')) for row in failures) and not any(state=='applied' for state in states):exit_code=4
                else:exit_code=5
            elif 'cancelled' in states:exit_code=3
            else:exit_code=0
            _record_cli_history('cli_deployment_completed',_deployment_audit_details(bundle,profile))
            return exit_code
        except MappingApprovalRequired as exc:
            _record_cli_failure('deploy',exc)
            print(f'Approval required: {format_failure(exc)}',file=sys.stderr);return 3
        except ApprovalRequired as exc:
            _record_cli_failure('deploy',exc)
            print(f'Approval required: {format_failure(exc)}',file=sys.stderr);return 3
        except NotImplementedError as exc:
            _record_cli_failure('deploy',exc)
            print(f'Deployment unsupported: {format_failure(exc)}',file=sys.stderr);return 3
        except MappingApprovalDeclined as exc:
            _record_cli_failure('deploy',exc)
            print(format_failure(exc),file=sys.stderr);return 3
        except RuntimeError as exc:
            _record_cli_failure('deploy',exc)
            message=format_failure(exc)
            if any(marker in message.casefold() for marker in ('excel','com','inspection failed')):
                print(f'Excel/COM operation failed: {message}',file=sys.stderr);return 4
            print(f'Deployment failed: {message}',file=sys.stderr);return 5
        except (ValueError,OSError) as exc:
            _record_cli_failure('deploy',exc)
            message=format_failure(exc);unsupported=any(marker in message.casefold() for marker in ('not supported','unsupported','capability probe','not available on this office build'))
            print(f"{'Deployment unsupported' if unsupported else 'Deployment validation failed'}: {message}",file=sys.stderr)
            return 3 if unsupported else 2
    if args.command!='plan':parser.error('Unsupported command.')
    try:
        profile=load_profile(args.profile)
        _review_cli_path_mappings(profile,args.profile,enabled=args.apply_path_mappings,
                                  input_func=input_func,interactive=interactive)
        with _review_cli_environment_mappings(profile,args.profile,enabled=args.apply_environment_mappings,
                input_func=input_func,interactive=interactive,
                apply_path_mappings=args.apply_path_mappings) as (effective_profile,environment_previews,path_previews):
            # All Excel COM work stays on a dedicated, COM-initialized worker thread.
            with ThreadPoolExecutor(max_workers=1,thread_name_prefix='pqm-cli-excel') as executor:
                bundle=executor.submit(_build_bundle,effective_profile,args.workbooks,
                    inspector_factory=inspector_factory,capability_probe=capability_probe,
                    apply_path_mappings=args.apply_path_mappings,environment_previews=environment_previews,
                    path_previews=path_previews if path_previews else None).result()
        rendered=json.dumps(bundle,ensure_ascii=False,indent=2,sort_keys=True)+'\n'
        if args.output:
            written=_write_new_file(args.output,rendered)
            print(f'Wrote dry-run plan bundle for {len(bundle["plans"])} workbook(s): {written}')
        else:
            sys.stdout.write(rendered)
        _record_cli_history('cli_plan_created',_plan_audit_details(bundle,profile))
        return 0
    except MappingApprovalRequired as exc:
        _record_cli_failure('plan',exc)
        print(f'Approval required: {format_failure(exc)}',file=sys.stderr)
        return 3
    except ApprovalRequired as exc:
        _record_cli_failure('plan',exc)
        print(f'Approval required: {format_failure(exc)}',file=sys.stderr)
        return 3
    except NotImplementedError as exc:
        _record_cli_failure('plan',exc)
        print(f'Plan unsupported: {format_failure(exc)}',file=sys.stderr)
        return 3
    except MappingApprovalDeclined as exc:
        _record_cli_failure('plan',exc)
        print(format_failure(exc),file=sys.stderr)
        return 3
    except RuntimeError as exc:
        _record_cli_failure('plan',exc)
        message=format_failure(exc)
        if any(marker in message.casefold() for marker in ('excel','com','inspection failed')):
            print(f'Excel/COM operation failed: {message}',file=sys.stderr)
            return 4
        print(f'Plan operation failed: {message}',file=sys.stderr)
        return 5
    except (ValueError, OSError) as exc:
        _record_cli_failure('plan',exc)
        message=format_failure(exc)
        unsupported=any(marker in message.casefold() for marker in
            ('not supported','unsupported','capability probe','not available on this office build'))
        code=3 if unsupported else 2
        label='Plan unsupported' if unsupported else 'Plan validation failed'
        print(f'{label}: {message}',file=sys.stderr)
        return code
    except Exception as exc:
        _record_cli_failure('plan',exc)
        print(f'Plan operation failed: {format_failure(exc)}',file=sys.stderr)
        return 5


def main(argv=None,*,inspector_factory=None,installer_factory=None,refresher_factory=None,
         capability_probe=None,input_func=None):
    parser=_make_parser()
    try:args=parser.parse_args(argv)
    except SystemExit as exc:return int(exc.code)
    confirmations=list(getattr(args,'confirm',()) or ())
    interactive=(input_func is not None) or bool(confirmations) or sys.stdin.isatty()
    if input_func is None and confirmations:
        def input_func(prompt):
            if confirmations:
                return confirmations.pop(0)
            if sys.stdin.isatty():
                return input(prompt)
            return ''
    input_func=input if input_func is None else input_func
    approval_reader=input_func
    def input_func(prompt):
        try:
            return approval_reader(prompt)
        except EOFError:
            # A closed console is an explicit absence of approval, never an
            # uncaught exception or implicit authorization.
            return ''
    if args.command in {'plan','deploy'}:
        try:
            with _deployment_profile_input(args.profile) as profile_path:
                args.profile=profile_path
                return _dispatch_parsed(args,parser,inspector_factory=inspector_factory,
                    installer_factory=installer_factory,refresher_factory=refresher_factory,
                    capability_probe=capability_probe,input_func=input_func,interactive=interactive)
        except (ValueError,OSError) as exc:
            _record_cli_failure('deployment_bundle',exc)
            print(f'Deployment profile validation failed: {format_failure(exc)}',file=sys.stderr)
            return 2
    return _dispatch_parsed(args,parser,inspector_factory=inspector_factory,
        installer_factory=installer_factory,refresher_factory=refresher_factory,
        capability_probe=capability_probe,input_func=input_func,interactive=interactive)


if __name__=='__main__':
    raise SystemExit(main())
