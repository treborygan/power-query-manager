# Feature acceptance and test cases

These are independent of internal implementation. Orchestrator links actual automated test names and real Excel test evidence to each ID in release notes. PASS means exercised, not 'implemented in source'. Distinguish automated/fake COM vs Windows Excel/packaged EXE verification.

## Legacy preservation (all phases)

| ID | Acceptance condition |
|---|---|
| LEG-01 | Mixed `.m` and `.pq` imports preserve BOM-independent M formula contents exactly (UTF-8/Unicode tested); unsupported extensions and empty files rejected before Excel opens. |
| LEG-02 | Files named `1_a.pq`, `02-b.m`, `003 c.pq` import in numeric order; `2foo.m` has no prefix; unprefixed alphabetically after; ties stable; user manual order wins. |
| LEG-03 | Enabled, order, query name and source path remain editable; validation surfaces duplicate names, bad paths/order, reserved names; disabling a query prevents import. |
| LEG-04 | Existing-query replace retains object/load where feasible; skip never modifies formula or load; errors clearly reported. |
| LEG-05 | `.xlsx` and `.xlsm` supported, macro contents preserved (not executed by importer); base tests remain green. |

## v1.3.1

| ID | Acceptance condition |
|---|---|
| S-01 | Existing query's Formula cannot be captured; simulated Add raises => existing query not deleted/modified and reported blocked. |
| S-02 | Read-only/locked/protected/invalid workbook rejected safely, target file hash stays unchanged. |
| S-03 | Excel opened with external link update suppressed and programmatic macro security force-disabled/restored, with documented Excel 4 macro caveat; only owned Excel instance quits. |
| S-04 | Two rapid backups create different filenames; each hash matches original; successful rollback/restore verification on injected COM, save and commit failures. |
| S-05 | GUI cannot mutate query list/paths/options/target while worker runs; immutable plan has original source hashes and detects disk changes before apply. |
| S-06 | On failure original workbook untouched by default staged strategy; logs identify backup/stage/recoverability without secrets. |
| S-07 | Onedir `PowerQueryInstaller.exe` launches from extracted directory on supported Windows, Qt plugins present, COM import works using disposable workbook, no orphan process. |

## v1.4.0

| ID | Acceptance condition |
|---|---|
| M-01 | Inspect 0/1/many existing queries; expose formulas/names/descriptions and unknown load mappings without editing workbook. |
| M-02 | Compare imported query to Excel existing query: NEW / MODIFIED / UNCHANGED / SKIP / CONFLICT / UNKNOWN, accurate with Unicode/comments/newlines; diff displays exact content. |
| M-03 | Export queries preserves original M expression, stable safe filenames and original query names manifest; name collisions handled explicitly. |
| M-04 | Profile roundtrip and schema validation work for relative paths, Unicode, all options and migration from older schema; invalid profile rejected safely. |
| M-05 | Grid filtering/sorting/bulk editing do not corrupt stable IDs, selected rows or import order; code editor requires explicit save/apply. |
| M-06 | Inspect-only/dry-run causes no file edits, queries refresh, connections created or dangerous prompts; verified by target hash/inventory comparison. |
| M-07 | ZIP rejects path traversal, symlinks, zip bombs and duplicate normalized names; folder recursion opt-in; all extraction isolated. |

## v1.5.0

| ID | Acceptance condition |
|---|---|
| L-01 | Definition-only creates/updates M formula without claiming a worksheet/data-model load; real Excel reopened formula equality. |
| L-02 | Connection-only offered only after real Excel proves explicit connection object/state and re-open persistence; otherwise UI disabled with why. |
| L-03 | New sheet load creates exactly one functional Excel table at intended sheet/cell, data refresh and reopen succeed on throwaway data with known row count. |
| L-04 | Existing sheet load chooses exact target cell, protects unrelated content and tables, disallows collision; supports safe idempotent rerun. |
| L-05 | Data Model load creates actual in-model table/connection visible in `Workbook.Model`; if unsupported, disabled and documented. |
| L-06 | Preserve-existing update keeps original table, sheet, table identity, connection, model membership where applicable; no destructive replace on ambiguous link. |
| L-07 | Dependency analyzer finds real direct references and cycles while ignoring comments/string literals, local bindings; ambiguous cases must be flagged, not auto-reordered. |
| L-08 | Refresh selected/all has explicit consent, distinguishes queued/in-progress/finished/failed and times out safely; no claim of completion for background refresh without confirmation. |
| L-09 | User-chosen order only changed via explicit approval; prior order reproducible and dependencies advisory. |

## v2.0.0

| ID | Acceptance condition |
|---|---|
| A-01 | Multi-workbook dry-run shows per-workbook changes without modifications; batch proceeds with per-workbook backup/commit/report and isolates failures. |
| A-02 | Environment mapping modifies only explicitly approved M-code spans; preview diff; warns dynamic/ambiguous source expressions; never stores secrets. |
| A-03 | Supported simple parameter edits roundtrip type/escaping without corrupting M; complex cases read-only. |
| A-04 | ZIP query package manifest/hash verified, extraction safe, compatible schema migration or actionable rejection. |
| A-05 | CLI `inspect/plan/import/export/validate/refresh` (if delivered) returns predictable codes: 0 success, 2 validation, 3 unsupported, 4 Excel/COM, 5 partial failure; `--dry-run` produces no side effects. |
| A-06 | Local history redacts credentials/formula bodies by default, has version identifiers and user-adjustable retention/deletion. |
| A-07 | Portable Windows distribution startup/import/refresh on supported desktop Excel verified; installer does not require admin for normal operation. |

## Reverse Dependency Finder (post-v2 enhancement)

| ID | Acceptance condition | Current gate |
| --- | --- | --- |
| DF-01 | Inventory Excel/Access consumers from selected files or folders, with optional recursion and multi-target selection; exclude each target from self-references. | PASS (portable unit/UI) |
| DF-02 | Detect references through real embedded Power Query M, workbook links and connection definitions; classify exact paths vs basename and tentative literals. | PASS (synthetic ZIP fixture); real Windows workbook UNVERIFIED |
| DF-03 | Inspect real `.accdb`/`.mdb` linked tables and saved queries on a local temporary copy; no original lock-file/writeback. | PASS (fake DAO); real Windows ACE/DAO UNVERIFIED |
| DF-04 | No external data refresh, Excel process, VBA execution, Access query execution or modification to original file hashes. | PASS (static code/portable synthetic fixture); Windows integration UNVERIFIED |
| DF-05 | Unsupported/unreadable/dynamic sources, absent DAO and canceled scans are visible; zero hits are never equated with proven independence. | PASS (portable unit); Windows matrix UNVERIFIED |
| DF-06 | GUI worker remains responsive/cancellable; JSON and CSV preserve dependency matches and coverage issues; CLI exit 5 represents partial coverage. | PASS (portable UI/CLI); packaged Windows executable UNVERIFIED |
| DF-07 | Verify Windows mapped-drive/UNC, OneDrive hydration, genuine `.xlsm`/`.xlsx` DataMashup and Access `.mdb`/`.accdb` smoke cases with original SHA-256 before/after. | UNVERIFIED — Windows smoke required |

## Scenarios to add (unit + Excel where possible)

- name clashes case-insensitive, trailing whitespace, Unicode, `#"quoted"` M identifiers, sheet 31-char / punctuation, existing table display name; duplicate imports from two ZIPs
- two files with prefix ties, manual order editing then another import, sort without reorder, missing/disabled dependency, cycle, reference inside text or comment
- workbook already open in Excel, read-only, OneDrive/SharePoint synced, zero-byte/invalid ZIP, password protected `.xlsm`, broken external connection, untrusted M code
- no Excel installed, pywin32 fail, Office 32-/64-bit mismatch, empty workbook, no network, revoked data source permissions, corporate policy restriction
- failure injection: before COM init, after workbook open, after N query adds, after connection create, during refresh, save, close, file replace and checksum verification
- queue/cancel during each safe phase; double-click install; GUI quit while worker active; no orphan COM processes

## Test result taxonomy

- **PASS (unit fake COM)**: deterministic programmatic behavior only.
- **PASS (Windows Excel)**: executed on an actual disposable workbook; record Office/build and facts.
- **PASS (packaged executable)**: built launcher was run as normal non-admin user in real Windows environment.
- **BLOCKED**: reason + steps to resolve and expected command.
- **UNVERIFIED**: code present, no live evidence, must not be advertised as supported production functionality.
- **FAIL**: reproducible defect + issue owner.
