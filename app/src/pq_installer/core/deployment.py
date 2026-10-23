"""Read-only multi-workbook planning and isolated per-workbook deployment."""
from __future__ import annotations

import hashlib
from dataclasses import dataclass, replace
from pathlib import Path
import uuid

from .manager import profile_to_sources
from .validation import validate_workbook, validate_workbook_for_write, validate_query_sources
from .models import ExistingQueryMode, QueryStatus
from .backup import create_timestamped_backup
from .excel_queries import ExcelQueryInstaller
from .redaction import safe_error


@dataclass(frozen=True)
class WorkbookDeploymentResult:
    workbook_path: str
    status: str
    plan: dict | None = None
    results: tuple[dict, ...] = ()
    backup_path: str | None = None
    message: str = ''


def deploy_multi_workbook(profile, profile_path, workbook_paths, *, inventory_provider,
                          supported_load_modes=('definition_only',), capabilities=None,
                          apply_path_mappings=False, installer_factory=None,
                          backup_factory=None, confirm_apply=None,
                          write_preflight=validate_workbook_for_write):
    """Plan all targets first, then apply isolated per-workbook staged transactions.

    Call this API from a single COM-capable worker thread. The default installer
    initializes/uninitializes COM per workbook and uses only its own Excel process.
    Errors are returned per target; no cross-workbook atomicity is claimed.
    """
    if not callable(inventory_provider):raise TypeError('A read-only workbook inventory provider is required.')
    installer_factory=installer_factory or (lambda _path:ExcelQueryInstaller())
    backup_factory=backup_factory or create_timestamped_backup
    targets=[];seen=set()
    for value in workbook_paths:
        target=validate_workbook(value).resolve();key=str(target).casefold()
        if key in seen:raise ValueError(f'duplicate workbook target in deployment: {target}')
        seen.add(key);targets.append(target)
    if not targets:raise ValueError('Select at least one workbook target for deployment.')
    options=profile.get('options',{})
    _,sources,stable_ids=profile_to_sources(profile,profile_path,
        supported_load_modes=supported_load_modes,apply_path_mappings=apply_path_mappings)
    enabled=[source for source in sources if source.enabled]
    if not enabled:raise ValueError('The profile has no enabled query sources to deploy.')
    validate_query_sources(enabled)
    source_hashes={}
    frozen_sources=[]
    for source,stable_id in zip(sources,stable_ids):
        if not source.enabled:
            continue
        if not source.path.is_file():raise ValueError(f'Query source is missing: {source.name}')
        digest=hashlib.sha256(source.path.read_bytes()).hexdigest()
        source_hashes[stable_id]=digest
        frozen_sources.append(replace(source,source_sha256=digest))

    prepared={}
    for target in targets:
        try:
            if not options.get('dry_run',False):
                write_preflight(target)
            plan=build_multi_workbook_dry_run(profile,profile_path,[target],
                inventory_provider=inventory_provider,supported_load_modes=supported_load_modes,
                apply_path_mappings=apply_path_mappings)[0]
            if any(item['formula_sha256']!=source_hashes.get(item['id']) for item in plan['items']):
                raise RuntimeError('A query source changed during deployment preflight.')
            prepared[target]=plan
        except Exception as exc:
            prepared[target]=WorkbookDeploymentResult(str(target),'blocked_preflight',message=safe_error(exc))

    eligible_plans=[]
    if not options.get('dry_run',False):
        for target,entry in prepared.items():
            if isinstance(entry,WorkbookDeploymentResult):continue
            if all(item['action']=='skip' for item in entry['items']):continue
            if any(item['action']=='inspect' or (item['action']!='skip' and item['warnings']) for item in entry['items']):continue
            eligible_plans.append(entry)
    approved=True
    if eligible_plans and confirm_apply is not None:
        approved=bool(confirm_apply(tuple(eligible_plans)))

    # Do not write any target until read-only preflight has been attempted for every target.
    reports=[]
    for target in targets:
        entry=prepared[target]
        if isinstance(entry,WorkbookDeploymentResult):
            reports.append(entry);continue
        plan=entry
        unsafe=[item for item in plan['items'] if item['action']=='inspect'
                or (item['action']!='skip' and item['warnings'])]
        if unsafe:
            message='Preflight requires manual review for: '+', '.join(item['query_name'] for item in unsafe)
            reports.append(WorkbookDeploymentResult(str(target),'blocked_preflight',plan,message=message));continue
        if all(item['action']=='skip' for item in plan['items']):
            reports.append(WorkbookDeploymentResult(str(target),'no_changes',plan,message='Preflight found no workbook changes.'));continue
        if options.get('dry_run',False):
            reports.append(WorkbookDeploymentResult(str(target),'planned',plan,message='Profile is marked dry-run; no workbook was changed.'));continue
        if not approved:
            reports.append(WorkbookDeploymentResult(str(target),'cancelled',plan,message='Deployment was not confirmed; no workbook was changed.'));continue
        expected_hash=plan['original_sha256'];backup_path=None
        try:
            if hashlib.sha256(target.read_bytes()).hexdigest()!=expected_hash:
                raise RuntimeError('Workbook changed after preflight; no changes were applied.')
            if options.get('backup',True):
                backup_path=str(backup_factory(target))
                if hashlib.sha256(Path(backup_path).read_bytes()).hexdigest()!=expected_hash:
                    raise RuntimeError('Backup verification failed; workbook deployment was not started.')
            installer=installer_factory(target)
            results=tuple(installer.install(target,frozen_sources,
                ExistingQueryMode(options['existing_query_policy']),
                capabilities=capabilities,expected_workbook_sha256=expected_hash))
            result_rows=tuple({'query_name':result.query_name,
                'status':getattr(result.status,'value',str(result.status)),
                'message':safe_error(getattr(result,'message',''))} for result in results)
            expected_names={source.name.casefold() for source in enabled}
            reported_names=[str(row['query_name']).casefold() for row in result_rows]
            successful_statuses={QueryStatus.CREATED.value,QueryStatus.REPLACED.value,
                                 QueryStatus.SKIPPED.value}
            complete_results=(len(reported_names)==len(expected_names)
                and set(reported_names)==expected_names)
            failed=(not complete_results or any(row['status'] not in successful_statuses
                                                 for row in result_rows))
            reports.append(WorkbookDeploymentResult(str(target),'failed' if failed else 'applied',
                plan,result_rows,backup_path,
                'Deployment did not complete successfully for every planned query.' if failed
                else 'Workbook transaction completed and all planned query results were successful.'))
        except Exception as exc:
            message=safe_error(exc)
            try:
                current_hash=hashlib.sha256(target.read_bytes()).hexdigest()
            except OSError:
                message += ' Final workbook state could not be verified after the failure.'
                if backup_path:
                    message += ' A verified pre-deployment backup is available in the backup_path field.'
            else:
                if current_hash==expected_hash:
                    message += f' Target still matches the pre-deployment SHA-256 ({expected_hash}).'
                else:
                    message += (f' Workbook state changed after the failure; current SHA-256 is '
                                f'{current_hash}.')
                    if backup_path:
                        message += ' A verified pre-deployment backup is available in the backup_path field.'
                    else:
                        message += ' No verified pre-deployment backup is available.'
            reports.append(WorkbookDeploymentResult(str(target),'failed',plan,(),backup_path,message))
    return tuple(reports)


def build_multi_workbook_dry_run(profile, profile_path, workbook_paths, *,
                                 inventory_provider, supported_load_modes=('definition_only',),
                                 apply_path_mappings=False):
    """Build one contract-v1 dry-run plan per workbook from read-only inventories.

    `inventory_provider(path)` must return query inventory records and must not save,
    refresh, or alter the workbook. The caller owns any worker-thread requirement.
    This helper remains read-only; `deploy_multi_workbook` owns separate apply transactions.
    """
    if not callable(inventory_provider):
        raise TypeError('A read-only workbook inventory provider is required.')
    targets=[];seen=set()
    for value in workbook_paths:
        target=validate_workbook(value).resolve()
        key=str(target).casefold()
        if key in seen:raise ValueError(f'duplicate workbook target in dry-run: {target}')
        seen.add(key);targets.append(target)
    if not targets:raise ValueError('Select at least one workbook target for the dry-run.')
    _,sources,stable_ids=profile_to_sources(profile,profile_path,
        supported_load_modes=supported_load_modes,apply_path_mappings=apply_path_mappings)
    enabled=[source for source in sources if source.enabled]
    if not enabled:raise ValueError('The profile has no enabled query sources to plan.')
    validate_query_sources(enabled)

    source_snapshot={}
    for source,stable_id in zip(sources,stable_ids):
        if not source.enabled:continue
        try:raw=Path(source.path).read_bytes()
        except OSError as exc:raise ValueError(f'Cannot read source {source.name!r}: {exc}') from exc
        try:formula=raw.decode('utf-8-sig')
        except UnicodeDecodeError as exc:raise ValueError(f'Source {source.name!r} is not valid UTF-8.') from exc
        if not formula.strip():raise ValueError(f'Source {source.name!r} is empty.')
        source_snapshot[stable_id]=(raw,formula,hashlib.sha256(raw).hexdigest())

    policy=profile['options']['existing_query_policy']
    plans=[]
    for target in targets:
        original_hash=hashlib.sha256(target.read_bytes()).hexdigest()
        inventory=tuple(inventory_provider(target))
        after_inventory_hash=hashlib.sha256(target.read_bytes()).hexdigest()
        if after_inventory_hash!=original_hash:
            raise RuntimeError(f'Workbook changed during dry-run inspection: {target}')
        by_name={};duplicates=set()
        for query in inventory:
            key=str(query.name).casefold()
            if key in by_name:duplicates.add(key)
            else:by_name[key]=query
        items=[]
        sequence=sorted(zip(sources,stable_ids),key=lambda pair:(pair[0].order,pair[0].path.name.casefold(),str(pair[0].path).casefold()))
        for source,stable_id in sequence:
            if not source.enabled:continue
            _,formula,source_hash=source_snapshot[stable_id]
            existing=by_name.get(source.name.casefold())
            warnings=[];expected_existing_hash=None
            if source.name.casefold() in duplicates:
                action='inspect';warnings.append('Multiple workbook queries have this case-insensitive name; resolve the conflict manually.')
            elif existing is None:
                action='create'
            elif getattr(existing,'formula',None) is None:
                action='inspect';warnings.append('Existing query formula is unavailable; action cannot be determined.')
            else:
                existing_formula=str(existing.formula)
                expected_existing_hash=hashlib.sha256(existing_formula.encode('utf-8')).hexdigest()
                if existing_formula==formula:action='skip'
                elif policy=='skip':action='skip'
                else:action='replace'
                if getattr(existing,'load_state','unknown')=='unknown':
                    warnings.append('Existing load mapping is unknown; this plan changes no worksheet or Data Model destination.')
            items.append({'id':stable_id,'query_name':source.name,'formula_sha256':source_hash,
                          'action':action,'order':source.order,'load':{'mode':source.load_mode.value},
                          **({'expected_existing_sha256':expected_existing_hash} if expected_existing_hash else {}),
                          'warnings':warnings})
        plans.append({'schema_version':1,'run_id':uuid.uuid4().hex,'workbook_path':str(target),
                      'original_sha256':original_hash,'dry_run':True,'refresh':'none',
                      'backup_required':bool(profile['options']['backup']),'items':items})

    for stable_id,(raw,_,_) in source_snapshot.items():
        source=next(source for source,row_id in zip(sources,stable_ids) if row_id==stable_id)
        if Path(source.path).read_bytes()!=raw:
            raise RuntimeError(f'Query source changed during dry-run: {source.name}')
    return tuple(plans)
