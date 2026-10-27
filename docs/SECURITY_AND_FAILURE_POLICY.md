# Security, privacy and failure policy

## Trust boundaries

Untrusted inputs: source M code, zip archives, profile JSON, workbook paths, user-edited query/destination names, Excel COM property values, external data connections and query results. `M` expressions can access files, databases and networks **when refreshed**. An import alone must not evaluate query results unless user explicitly enables refresh or load modes that inherently evaluate query data. Warn when a requested load requires evaluation.

## Never do automatically

- Execute arbitrary code from source archives, shell commands, macros or startup scripts.
- Enable Excel macros or remove Windows security protections; `AutomationSecurity` should force-disable during programmatic workbook open, with documented Excel 4-macro exceptions. Do not pretend this guarantees absence of every prompt.
- Refresh queries without explicit option/confirmation; remote data access may transmit information.
- Overwrite input files while exporting; delete original query if safe formula/load restoration is not established.
- Force-close an Excel instance not launched by the app; use DispatchEx and retain process ownership.
- Store passwords/connection tokens/formula bodies in crash logs, profiles, history or telemetry. The user's explicit export of M code is exception to logging restrictions, not to secure file storage.
- Disable endpoint protection, modify registry, install services or use elevated privileges to evade corporate policy.

## Redaction and diagnostics

Logs may contain query name, status, error category, stage, elapsed time, anonymized workbook basename and hashes. Default log path under user profile with configurable retention; no logs with full M formula, access tokens or complete internal URLs by default. Provide explicit diagnostic bundle review before export. Profiles cannot carry credentials; references to external sources remain user-managed.

Missing required approval has its own `approval_required` diagnostic category. It must name the exact confirmation phrase and distinguish absent approval from an explicit wrong-phrase decline; neither path may start the protected write/refresh operation.

## Safe imports

Use deterministic immutable job plan with input/workbook checksums; verify before applying. Staged workbook approach protects original until successful verification. Backups collision-proof and user-visible with integrity hashes. If staging path is not semantically equivalent (Power Query relative paths, OneDrive sync, protected file), block and offer explicit output-copy workflow rather than unsafe in-place editing. Do not assume `os.replace` atomic across network shares. Test rollback after `Save`/COM partial write/commit conflict.

## Reverse dependency scans

Dependency Finder is a metadata-only operation: no Excel COM launch,
query/SQL execution, VBA, refresh, link update or writeback. Access DAO is
initialized inside the scan worker and opens a **temporary local copy** with
its read-only flag; this avoids creating lock files beside mapped or synced
original Access databases. OS reads, network traffic to fetch file bytes,
cloud placeholder hydration and temporary-disk use can still occur. Copies
are disposed after inspection. If an Access provider is unavailable,
unsupported, password-protected, locked or corrupt, report coverage gaps,
not an empty-success result. XML/ZIP parts have per-file entry/size limits
and a 256 MiB cumulative per-workbook metadata budget.

Reports omit raw M, SQL and connection strings, as these can contain
passwords or private URLs. JSON and CSV outputs are explicitly user-chosen,
new files only. Formula-like CSV cell values are text-escaped to prevent
Excel/CSV formula injection from malicious query/connection names. Matches
from basename or inert literals must remain tentative;
a zero-hit report cannot be used as proof that an old source can be removed.

## Support matrix

Unsupported workbook types (`.xlsb`, `.xls`, encrypted/password protected, Protected View) **must be rejected or explicitly researched**, not silently processed. If user requests expanded formats, add separate feasibility/acceptance matrix before enabling. Queries containing `section` declarations or multiple definitions are not implicitly supported by current standalone-expression importer. Never split them with naive regex.

## Package supply chain

Pinned dependency versions should be updated only after compatibility tests and vulnerability review; produce local dependency inventory and license summary where applicable. Do not fetch/execute untrusted remote build scripts. Build artifact paths must be free of private data and verify SHA-256. Package PyInstaller `onedir` launcher with `_internal` and required licenses; never promise a lone EXE functions if onedir required.

## Exit error categories

`INPUT_VALIDATION`, `PROFILE_SCHEMA`, `WORKBOOK_LOCKED`, `READ_ONLY`, `UNSUPPORTED_EXCEL`, `UNSUPPORTED_LOAD`, `COM_INIT`, `COM_OPERATION`, `LOAD_CONFLICT`, `REFRESH_AUTH`, `REFRESH_TIMEOUT`, `SAVE_VERIFY`, `COMMIT_CONFLICT`, `RESTORE_FAILED`, `CANCELLED_SAFE`. Every message should recommend next practical action, never report success after partial failure.
