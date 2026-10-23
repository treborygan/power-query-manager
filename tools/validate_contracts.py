#!/usr/bin/env python3
"""Validate external JSON schemas and example profile (optional: pip install jsonschema)."""
from pathlib import Path
import json
import sys
root=Path(__file__).resolve().parent.parent
try:
    from jsonschema import Draft202012Validator
except ImportError:
    print('UNVERIFIED: jsonschema not installed; structural syntax checks are in verify_handoff.py')
    raise SystemExit(2)
profile=json.loads((root/'contracts/profile.schema.json').read_text(encoding='utf-8'))
plan=json.loads((root/'contracts/plan.schema.json').read_text(encoding='utf-8'))
package=json.loads((root/'contracts/query-package.schema.json').read_text(encoding='utf-8'))
example=json.loads((root/'examples/profiles/demonstration.profile.json').read_text(encoding='utf-8'))
plan_example=json.loads((root/'examples/plans/demonstration.plan.json').read_text(encoding='utf-8'))
Draft202012Validator.check_schema(profile)
Draft202012Validator.check_schema(plan)
Draft202012Validator.check_schema(package)
Draft202012Validator(profile).validate(example)
Draft202012Validator(plan).validate(plan_example)
package_v1={'schema_version':1,'package_type':'power-query-manager-query-package','queries':[
    {'query_name':'Q','file':'queries/0001_Q.pq','order':0,'sha256':'0'*64}]}
package_v2={'schema_version':2,'package_type':'power-query-manager-query-package','queries':[
    {'query_name':'Q','file':'queries/0001_Q.pq','order':0,'sha256':'0'*64,'dependencies':[],'dependency_status':'known'}],
    'dependency_cycles':[]}
Draft202012Validator(package).validate(package_v1)
Draft202012Validator(package).validate(package_v2)
print('PASS: profile, plan, and query-package Draft 2020-12 schemas are valid; all v1/v2 contract samples conform')
