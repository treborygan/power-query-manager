# Build, support and release operations

## Normal development

`app/` is an editable Python project; on Windows use the existing `app/build.ps1` as starting point and improve it. On Linux without PySide6 use the pure-core suite via `PYTHONPATH=app/src python -m pytest -q app/tests`. Existing 18 tests are a baseline only; new tests needed for all features.

## Windows onedir distribution

1. Create isolated venv from supported Python 3.11–3.13 or deliberately upgrade dependencies for a newer supported Python. Pin versions to tested compatible toolchain, record hashes/version inventory.
2. Run unit+GUI tests; on Windows run Excel integration tests against disposable local workbook(s) and Windows acceptance matrix.
3. Run PyInstaller onedir from clean source and verify all dependencies/Qt platform plugins and `pythoncom`/`pywintypes` resolved. Prefer `--noupx`/no UPX if security tools flag binaries; test both routes only as permitted.
4. Launch EXE from `dist/PowerQueryInstaller` as a non-admin user, confirm window, local scan, COM import, exit, no orphan Excel processes. Do not equate PyInstaller exit 0 with successful launcher.
5. Create `PowerQueryManager-vX.Y.Z-Windows-onedir.zip` containing only `PowerQueryInstaller.exe`, `_internal/`, `README.txt`, `LICENSES/` if needed, and `SHA256SUMS.txt`. If utility launches from arbitrary writable directory, test spaces and Unicode in path.
6. Optional onefile EXE is a separate additional distribution, available only after successfully tested launch and COM on Windows. Do not replace working onedir build just for single-file packaging.

## Source release vs user-facing release

Keep all development source/tests/fixtures/evidence in repo; **never** include these inside user-facing runtime distribution. Include a functional summary in README:
- imports `.m` and `.pq`, editable names/paths/order, skip/replace, backups
- inspector/diff/profile/export when actually implemented
- exact supported Excel load destinations, Office versions/limitations
- Windows/Excel prerequisite; no administrator requirement for normal use if host permits
- restore procedure/backups, known limitations, manual intervention

## User operations

- Instruct to close target Excel workbook and choose disposable workbook for first run.
- Always allow review/dry-run before changes; backup location provided before apply.
- Staged file will be applied only if target hash is unchanged and Excel hasn't taken lock. Cloud-synced conflicts cause abort with preserved copies.
- Recovery: present backup and staged workbook location. Automatic restore only if proven safe; never overwrite an independently modified target.

## Real test gaps

Handoff built under Linux with no Windows Excel or PySide6 COM runtime: package does not include a verified exe. Codex must verify on a qualifying Windows host; exact results must be recorded in `evidence/TEST_REPORT_TEMPLATE.md`. Missing host cannot become PASS.
