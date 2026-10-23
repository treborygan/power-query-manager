#!/usr/bin/env python3
"""Verify handoff structure/serialized contracts and pristine source hashes.

Use --allow-modified-app after Codex begins modifying app/. Pass --run-tests to
also execute the existing project's pytest suite if installed.
"""
import argparse
import hashlib
import json
import os
from pathlib import Path
import subprocess
import sys

ROOT = Path(__file__).resolve().parent.parent
REQUIRED = [
    'START_HERE.md','CODEX_KICKOFF_PROMPT.md','AGENTS.md','docs/BASELINE_AUDIT.md',
    'docs/ARCHITECTURE.md','docs/ROADMAP.md','docs/QA_WINDOWS_MATRIX.md',
    'docs/ACCEPTANCE_CRITERIA.md','docs/EXCEL_COM_SPIKES.md',
    'contracts/profile.schema.json','contracts/plan.schema.json',
    'contracts/release_gates.json','examples/profiles/demonstration.profile.json',
    'app/pyproject.toml','app/src/pq_installer/core/excel_queries.py',
    'app/tests/test_excel_queries.py','BASELINE_SHA256SUMS.txt'
]

def digest(path):
    h=hashlib.sha256()
    with path.open('rb') as f:
        for data in iter(lambda: f.read(1024 * 1024), b''):
            h.update(data)
    return h.hexdigest()

def error(msg):
    print('FAIL:',msg)
    return 1

def main():
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--allow-modified-app',action='store_true',help='Skip pristine baseline hash checks after implementation edits')
    parser.add_argument('--run-tests',action='store_true',help='Run app pytest with source PYTHONPATH')
    args=parser.parse_args()
    failures=0
    for rel in REQUIRED:
        if not (ROOT/rel).is_file(): failures+=error(f'missing {rel}')
    for path in sorted(ROOT.rglob('*')):
        if not path.is_file(): continue
        if '__pycache__' in path.parts or '.pytest_cache' in path.parts or '.venv' in path.parts:
            failures+=error(f'cache/env path should not be shipped: {path.relative_to(ROOT)}')
    docs=[p for p in (ROOT/'docs').glob('*.md')]
    if len(docs)<8: failures+=error('not enough specification documents')
    for rel in ('contracts/profile.schema.json','contracts/plan.schema.json','contracts/release_gates.json','examples/profiles/demonstration.profile.json'):
        try: json.loads((ROOT/rel).read_text(encoding='utf-8'))
        except Exception as exc: failures+=error(f'invalid JSON {rel}: {exc}')
    try:
        gate=json.loads((ROOT/'contracts/release_gates.json').read_text())
        assert [x['version'] for x in gate['releases']] == ['1.3.1','1.4.0','1.5.0','2.0.0']
        assert len({f for v in gate['releases'] for f in v['features']}) == 33
        road=(ROOT/'docs/ROADMAP.md').read_text(encoding='utf-8')
        crit=(ROOT/'docs/ACCEPTANCE_CRITERIA.md').read_text(encoding='utf-8')
        matrix=(ROOT/'docs/QA_WINDOWS_MATRIX.md').read_text(encoding='utf-8')
        for release in gate['releases']:
            for feature in release['features']:
                if feature not in road: failures+=error(f'feature ID absent from roadmap: {feature}')
            for ac in release['mandatory_acceptance']:
                if ac not in crit: failures+=error(f'acceptance ID absent: {ac}')
            for w in release['windows_cases']:
                if w not in matrix: failures+=error(f'Windows case absent: {w}')
        print('OK: release gate references and milestones')
    except Exception as exc: failures+=error(f'bad release gate file: {exc}')
    hashes=ROOT/'BASELINE_SHA256SUMS.txt'
    if hashes.exists():
        count=0
        for line in hashes.read_text().splitlines():
            if not line.strip() or line.startswith('#'): continue
            sha, rel = line.split('  ',1)
            file=ROOT/rel
            if not file.is_file(): failures+=error(f'missing baseline file: {rel}')
            elif not args.allow_modified_app and digest(file)!=sha:
                failures+=error(f'baseline hash mismatch: {rel}')
            count+=1
        if count!=22: failures+=error(f'expected 22 baseline files, found {count}')
        else: print(f'OK: {count} baseline files accounted for' + (' (hash bypass enabled)' if args.allow_modified_app else ', hashes match'))
    if args.run_tests:
        env=os.environ.copy()
        env['PYTHONPATH']=str(ROOT/'app/src')+os.pathsep+env.get('PYTHONPATH','')
        p=subprocess.run([sys.executable,'-m','pytest','-q',str(ROOT/'app/tests')],cwd=ROOT/'app',env=env)
        if p.returncode: failures+=error(f'app pytest failed with exit code {p.returncode}')
        else: print('OK: pytest suite completed')
    print('RESULT:', 'PASS' if failures==0 else f'FAIL ({failures} error(s))')
    return 1 if failures else 0
if __name__ == '__main__': raise SystemExit(main())
