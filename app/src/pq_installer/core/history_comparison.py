"""Conservative, read-only comparison of query inventories from workbook versions."""
from dataclasses import dataclass
from difflib import unified_diff

from pq_installer.core.manager import QueryInfo


@dataclass(frozen=True)
class QueryVersionChange:
    name: str
    status: str
    before: QueryInfo | None
    after: QueryInfo | None
    reason: str
    diff: str


def compare_query_inventories(before, after) -> tuple[QueryVersionChange, ...]:
    """Compare two read-only workbook inventories without guessing duplicate identities."""
    before=tuple(before);after=tuple(after)
    old_by_key={};new_by_key={}
    for rows,target in ((before,old_by_key),(after,new_by_key)):
        for row in rows:
            if not isinstance(row,QueryInfo):
                raise TypeError('Snapshot inventory rows must be QueryInfo values.')
            target.setdefault(row.name.casefold(),[]).append(row)
    changes=[]
    visited=set()
    for row in before:
        key=row.name.casefold()
        if key in visited:continue
        visited.add(key)
        old_rows=old_by_key[key];new_rows=new_by_key.get(key,[])
        if len(old_rows)>1 or len(new_rows)>1:
            for conflict in old_rows:
                changes.append(QueryVersionChange(conflict.name,'CONFLICT',conflict,None,
                    'Duplicate case-insensitive query names prevent reliable version matching.',''))
            for conflict in new_rows:
                changes.append(QueryVersionChange(conflict.name,'CONFLICT',None,conflict,
                    'Duplicate case-insensitive query names prevent reliable version matching.',''))
            continue
        old=old_rows[0]
        new=new_rows[0] if new_rows else None
        if new is None:
            changes.append(QueryVersionChange(old.name,'REMOVED',old,None,'Query is absent from the later version.',
                ''.join(unified_diff((old.formula or '').splitlines(keepends=True),[],fromfile=old.name,tofile='/dev/null'))))
        elif old.formula is None or new.formula is None:
            changes.append(QueryVersionChange(old.name,'UNKNOWN',old,new,
                'At least one workbook version does not expose this query formula.',''))
        else:
            old_formula=old.formula.replace('\r\n','\n').replace('\r','\n')
            new_formula=new.formula.replace('\r\n','\n').replace('\r','\n')
            formula_changed=old_formula!=new_formula
            loads_known=old.load_state!='unknown' and new.load_state!='unknown'
            load_changed=loads_known and old.load_state!=new.load_state
            name_changed=old.name!=new.name
            status='MODIFIED' if formula_changed or load_changed or name_changed else ('UNCHANGED' if loads_known else 'UNKNOWN')
            reasons=[]
            if name_changed:reasons.append(f'Query display name changed from {old.name} to {new.name}.')
            if formula_changed:reasons.append('M formula changed.')
            if load_changed:reasons.append(f'Load destination changed from {old.load_state} to {new.load_state}.')
            if not loads_known:reasons.append('Load destination is unknown in at least one version.')
            diff=''.join(unified_diff(old.formula.splitlines(keepends=True),new.formula.splitlines(keepends=True),
                                      fromfile=f'{old.name} (earlier)',tofile=f'{new.name} (later)')) if formula_changed else ''
            changes.append(QueryVersionChange(old.name,status,old,new,' '.join(reasons) or 'Formula and known load destination are unchanged.',diff))
    for row in after:
        key=row.name.casefold()
        if key in visited:continue
        visited.add(key)
        rows=new_by_key[key]
        if len(rows)>1:
            for conflict in rows:
                changes.append(QueryVersionChange(conflict.name,'CONFLICT',None,conflict,
                    'Duplicate case-insensitive query names prevent reliable version matching.',''))
        elif row.formula is None:
            changes.append(QueryVersionChange(row.name,'UNKNOWN',None,row,
                'Formula is unavailable in the later workbook version.',''))
        else:
            changes.append(QueryVersionChange(row.name,'ADDED',None,row,'Query is present only in the later version.',
                ''.join(unified_diff([],row.formula.splitlines(keepends=True),fromfile='/dev/null',tofile=row.name))))
    return tuple(changes)
