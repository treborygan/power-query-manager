"""Check that test function names cited by feature traceability exist in app/tests."""
from __future__ import annotations

import ast
from pathlib import Path
import re
import sys

ROOT = Path(__file__).resolve().parents[1]
TRACEABILITY = ROOT / 'docs' / 'FEATURE_TRACEABILITY.md'
TEST_ROOT = ROOT / 'app' / 'tests'


def main() -> int:
    referenced = set(re.findall(r'\btest_[A-Za-z0-9_]+\b', TRACEABILITY.read_text(encoding='utf-8')))
    defined = set()
    for path in TEST_ROOT.rglob('*.py'):
        try:
            tree = ast.parse(path.read_text(encoding='utf-8'), filename=str(path))
        except (OSError, UnicodeError, SyntaxError) as exc:
            print(f'Could not inspect {path.relative_to(ROOT)}: {type(exc).__name__}')
            return 2
        defined.update(
            node.name for node in ast.walk(tree)
            if isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef))
            and node.name.startswith('test_')
        )
    missing = sorted(referenced - defined)
    if missing:
        print(f'FAIL: {len(missing)} traceability test references do not exist:')
        for name in missing:
            print(f'  {name}')
        return 1
    print(f'PASS: all {len(referenced)} cited test functions resolve across app/tests.')
    return 0


if __name__ == '__main__':
    sys.exit(main())
