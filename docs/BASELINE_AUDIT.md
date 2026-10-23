# Evidence-based baseline source audit — 2026-10-21

The app under `app/` is copied from the user-supplied archive `Power_Query_Installer_1.3.0_Source_Build.zip` without modification. Original hash manifest is `BASELINE_SHA256SUMS.txt`. No changes are present in the baseline at handoff creation.

## Current architecture / key locations

| File | Observed responsibility |
|---|---|
| `app/src/pq_installer/main.py` | Qt application entrypoint |
| `app/src/pq_installer/ui/main_window.py` | Single-window UI, editable five-column QTableWidget, QThread orchestration; hardcoded styling |
| `app/src/pq_installer/ui/install_worker.py` | Backup then COM install work, signal/error reporting |
| `app/src/pq_installer/core/models.py` | QuerySource, LoadMode(QUERY_ONLY), ExistingQueryMode(REPLACE/SKIP), install status |
| `app/src/pq_installer/core/query_loader.py` | Multi-file discovery, .m/.pq, prefixes/deterministic order, UTF-8 BOM reading |
| `app/src/pq_installer/core/validation.py` | Path/extension/query name/order/source validation |
| `app/src/pq_installer/core/excel_queries.py` | `DispatchEx`, `pythoncom.CoInitialize`, `Queries.Add`, `Formula` update, rollback journal, `Save` |
| `app/src/pq_installer/core/backup.py` | Timestamped shutil.copy2 next to workbook |
| `app/PowerQueryInstaller.spec`, `app/build.ps1` | PyInstaller onedir, Windows scripted tests/build |
| `app/tests/` | 18 tests: source discovery/validation/backup, mock COM definitions/rollback and instructions |

## Baseline result

Command used: `cd app && PYTHONPATH=src python -m pytest -q` on Linux host. Result: **18 passed** (not a Windows Excel integration result). The package has NO built Windows executable. No user's production Excel file was opened.

## Observed correctness/safety issues

- `excel_queries.py`: when an existing query has no readable `Formula`, code deletes existing object, calls `Queries.Add`, **then** journals replacement. If Add fails, original is lost in COM session and no rollback journal entry exists. Note that an exception while reading `Formula` already aborts; the reproducible unsafe branch is specifically when `Formula` returns `None`. Although workbook is closed without saving, success of cleanup is not a substitute for safe in-session behavior. Remove this fallback or implement demonstrably reversible mechanism.
- `excel_queries.py`: direct caller does not validate workbook path and directly opens it; `Workbooks.Open()` is not explicit about `UpdateLinks`, read-only/write permission/macro execution. Detection of read-only after open is needed.
- `backup.py`: suffix based on second resolution could collide/overwrite in quick repeated operations. Backup policy is optional, and there is no manifest validation, retention, secure overwrite control, or staged-file commit design.
- `ui/main_window.py`: `_set_busy()` only disables import button, not source grid, workbook picker/browse, settings, open-workbook button. Source objects can still be modified while a worker retains shared references.
- `ui/main_window.py`: row view uses `QTableWidget` and `id(source)` as UID; app has no persistent profile identity, search, sort, bulk editing or release-safe immutable plan representation.
- `validation.py`: only accepts definition-only load; no load destination or Excel state validation, no `.xlsb` input (do not add until verified), no protected/synced conflict handling.
- `query_loader.py`: folder scan only immediate child; `.m`/`.pq` files assumed standalone complete formula; no section documents/multiple definitions support.
- `install_worker.py`: errors displayed, but lacks structured per-stage report, cancellation boundaries and explicit save/commit validation.
- Original README notes an unverified standalone executable and Windows COM acceptance. Preserve the distinction.

## Current verified functions

- Mixed `.m` / `.pq` files can be discovered from selections or immediate folder.
- Names and ordering derived from leading integer + `_`, `-`, or whitespace separator. Unprefixed sorted after highest prefix.
- Query names, source file paths, numeric orders, enabled state edited in UI; tie warnings displayed.
- `ExistingQueryMode` supports replace formula or skip; creation via `Queries.Add`, no worksheet load.
- Workbook backup and failure rollback for previously journaled query-definition writes.

## Deferred/blocked features

No code currently creates/inspects `Workbook.Connections`, `ListObjects`/`QueryTable`, `Workbook.Model`, or performs refresh. Do not state that it already has data-model, worksheet load, or connection-only support. No validation of Excel versions is available in this environment.

## Regression-first approach

Reproduce the destructive fallback and test no existing object is deleted if its formula cannot be captured. Add read-only and macro tests, duplicate backup test, COM ownership cleanup tests, UI mutation lock test, and source/profile migration tests **before** adding advanced functionality.
