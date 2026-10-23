# Power Query Manager implementation report

## Current deliverable

- Source: `app/`; one-folder Windows package: `dist/PowerQueryManager/`.
- Package verifier: **PASS**, 116 files and 76,241,206 payload bytes.
- `dist/SHA256SUMS` SHA-256: `A543D52F5855F2DD8BBB83DF5B94598825561792849FFDC5F9630B3CF907E2FB`.
- GUI SHA-256: `EB53833FC96CBE8FBE3AD40280F6569799D1D4D5FAB223C8713CF67DD509A93C`; CLI SHA-256: `A631F4256483B10EA921CBBD463A28C88318C1C7E7361AC7AD17E6BFA9255C4E`.
- Package contains the GUI, CLI, user manual, and release summary. No `.git` metadata was present, so commit IDs are unavailable.
- Host used for current Windows acceptance: Windows 10 build 26300, Python 3.11.9, PyInstaller 6.13.0, desktop Excel 16.0 build 20430.0. Tests and Excel smokes ran non-elevated where stated.

## Implemented

- **v1.3.1 safety and repeatability:** staged workbook transactions, verified backups and read-only reopen, guarded rollback, skip no-op, macro/security checks, owned Excel process cleanup, redacted diagnostics, and one-folder packaging.
- **v1.4.0 manager:** PySide6 `QAbstractTableModel`/`QTableView` workflow, source ordering and editing, read-only query/load inventory, source/workbook comparison and M diff/editor, query export, versioned profiles, and conservative dependency display.
- **v1.5.0 Excel behavior:** capability-probed definition-only, connection-only, new worksheet, existing worksheet/cell, Data Model and preserve-existing loads; selected/all refresh with timeout, diagnostics, and conservative dependency handling.
- **v2.0.0 advanced tooling:** multi-workbook planning/deployment, source-path and M-code mapping previews, declared-parameter editing, query/deployment packages, local history, package comparison, and optional CLI.
- The shared `ExcelApplicationAdapter`, `ExcelWorkbookOperations` lifecycle port, and staged `WorkbookTransactionAdapter` provide injectable seams for fake-COM tests. Workbook lifecycle calls in inspection, import, refresh, and capability probing use the shared port.
- Current focused verification for the lifecycle port: **91 passed** across adapter, inspector, importer, refresher and capability tests. The rebuilt package passed a disposable definition-only GUI import on Excel 16.0 build 20430.0, packaged CLI `--help`, and package verification. W293–W301 record packaged query-package import, hostile ZIP rejection, VBA preservation, read-only deployment, active-workbook refusal without an owner-lock file, undoable row removal, and redacted CLI history.
- W299 acceptance: while Excel held a disposable workbook open read-only without a `~$` lock file, packaged CLI deployment returned `blocked_preflight` (exit 5); workbook hash stayed unchanged, no backup/stage was created, and no new Excel process started. Only the harness-owned Excel process was closed; existing PID 36464 remained.
- W301: three focused history regressions passed; packaged CLI mixed writable/read-only deployment and packaged GUI definition-only import passed on Excel 16.0 build 20430.0. Broader Windows/Office gates remain **UNVERIFIED**.
- W302: within-batch source-path duplicates are filtered before they enter the table; focused model suite passed 41 tests. The rebuilt current package passed a disposable definition-only GUI import on Excel 16.0 build 20430.0 with exact formula, verified original-hash backup, read-only reopen, manager exit 0 and owned-process cleanup. Packaged keyboard traversal did not complete reliably and remains open; the full application suite was not run.
- W300 implements undo/redo for selected-row removal. The focused UI model file passed **40 tests**; the new regression first failed on the old behavior. The rebuilt package passed its 35-control keyboard cycle and CLI startup check. Screen-reader/physical display review and the full test suite remain unverified for this change.

## Incomplete or unverified

- All four release gates remain **UNVERIFIED**; this package is an implementation preview, not a production-certified release.
- Only one installed Excel build was available. Other Office/Excel builds, bitness combinations, enterprise policies, external providers/connectors, representative production workbooks, complex Data Model relationships, and some workbook object variants remain unverified.
- Screen-reader user testing and physical display/OS scaling review require other people/devices. Automated Qt scaling, viewport and keyboard checks do not replace those checks.
- Arbitrary blocked COM methods, all transaction interruption points, OneDrive/sync races, broader ACL configurations, damaged SQLite page salvage, and directory-entry durability remain open or intentionally preserve data for manual recovery.
- The larger domain workflow protocol remains service-specific; `ExcelWorkbookOperations` isolates workbook lifecycle calls but does not move query discovery, mutation, load configuration, refresh, and verification into one backend.
- The full application suite was not rerun after the W298 adapter change. The targeted 91-test selection and rebuilt package checks passed. Use the release files for each milestone's historical commands and exact scope.

## Release status

| Release | Implementation state | Gate |
|---|---|---|
| v1.3.1 | Safety and packaging features implemented; lifecycle and broad failure matrix incomplete | **UNVERIFIED** |
| v1.4.0 | Manager, inspection, export, diff and profiles implemented; physical accessibility and broader workbook matrix incomplete | **UNVERIFIED** |
| v1.5.0 | Load modes, preservation, dependencies and refresh implemented for probed Excel behavior; other builds/providers incomplete | **UNVERIFIED** |
| v2.0.0 | Deployment, mapping, parameter, package, history and CLI features implemented; broader security, recovery and platform acceptance incomplete | **UNVERIFIED** |

See [release evidence](releases/) and [Windows QA matrix](../docs/QA_WINDOWS_MATRIX.md) for detailed observed results and limitations. No release gate is promoted by this summary.
