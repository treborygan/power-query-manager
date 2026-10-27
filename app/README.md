# Power Query Manager — Windows application guide

This guide covers the current development-preview source, Excel behavior, local development, and packaging. The implementation includes work across the v1.3.1–v2.0.0 roadmap; some release requirements remain partial or unverified.

Milestone history: [changelog](../docs/CHANGELOG.md). Release evidence and gates are maintained in `../evidence/releases/`.

Windows desktop utility for adding and updating Power Query M query definitions in an Excel workbook using desktop Excel COM automation.

## Portable Windows download

Download the [latest repository ZIP](https://github.com/treborygan/power-query-manager/archive/refs/heads/main.zip), extract it with File Explorer, then open `power-query-manager-main\dist\PowerQueryManager\` and run `PowerQueryManager.exe`. Keep the entire `PowerQueryManager` folder together so its `_internal` runtime files remain beside the application. This is a portable folder, not an MSI or setup wizard; administrator installation is not required. Desktop Microsoft Excel is required for workbook operations. The package manual is at `dist/README.txt`.

If startup reports `No module named 'shiboken6.Shiboken'`, confirm that `dist\PowerQueryManager\_internal\shiboken6\Shiboken.pyd` exists. Re-extract the complete ZIP if it is missing; the executable cannot run by itself.

## Current release status

This source package includes v1.3.1 stability work, the v1.4 manager workflow, and partial v1.5/v2 support. Definition-only, connection-only, worksheet, and Data Model loads use build-specific disposable probes and staged save/reopen verification. Static-M app-level load/import/reimport smokes passed on Excel 16.0 builds 20326.0 and 20430.0; W49 revalidated the process-aware, 60-second-per-call capability probe and all load modes on build 20430.0. The packaged GUI passed definition-only import/reopen on build 20430.0 in W55 and existing-query replacement in W270; the W271 rebuild passed packaged keyboard traversal across 35 controls. The shared Excel application adapter, staged-file transaction protocol, and injectable workbook-lifecycle port are tested. Excel workbook add/open/save/save-as/close calls in inspection, import, refresh, and capability probing now pass through the lifecycle port; query/load-specific COM logic remains in its domain services. The Windows package excludes unused Qt translation catalogs, optional Qt QML/Quick, PDF, virtual keyboard, OpenGL and network/TLS components; the application interface is English-only. The manager itself does not use Qt networking; an explicitly requested query refresh runs through desktop Excel and may contact the sources defined by that query. The full roadmap is not complete or released; VBA-bearing `.xlsm`, other Office builds, external connectors, arbitrary blocked COM calls, and representative production workbooks are not certified. See release evidence for current gates.

## Workbook privacy preservation

Before a staged import opens Excel, the manager reads the workbook DataMashup permission metadata when present. It validates the workbook's DPAPI-bound permission binding and compares FirewallEnabled and WorkbookGroupType before commit; if the saved staged copy changes those settings or the binding cannot be validated for the current Windows user, the import fails closed and leaves the selected workbook unchanged. The check does not inspect or change source credentials or privacy-level assignments. Default FirewallEnabled=true / WorkbookGroupType=nil was preserved in one disposable Excel 16.0 build 20430.0 smoke. Non-default permission values and cross-user DPAPI-bound workbooks remain unverified.

## Functions

### Dependency Finder (new source feature; Windows verification outstanding)

Choose **Dependency Finder…** to find *incoming* dependencies: for example,
"Which `.xlsx`, `.xlsm`, `.accdb` or `.mdb` files reference `Warehouse.accdb`?"
Inventory a mapped/local folder recursively or add individual files, select
which candidate files to inspect, then select one or more targets. Alternatively
choose **Map all selected files** to index their internal links once, then
explore incoming/outgoing edges by file without rescanning. **Quick** checks
queries/connections/links; **Deep** additionally streams worksheet formula
and defined-name metadata. A visual advisory file graph is available alongside
the searchable results table. It scans
embedded Excel Power Query and workbook link metadata without starting Excel,
and Access table/query link definitions from an isolated temporary database
copy using a compatible Windows DAO engine. View filename-only/possible
matches, incomplete coverage and export to a new JSON/CSV report. No source
data is refreshed or modified. The existing packaged EXE has not been rebuilt
or smoke-tested for this feature.

CLI: `power-query-manager-cli dependency-scan --target "C:\Data\Warehouse.accdb" --folder "C:\Reports" --recursive`.
General map: `power-query-manager-cli dependency-map --folder "C:\Reports" --recursive --depth quick`.
See [usage, detection scope and acceptance gates](../docs/DEPENDENCY_FINDER.md).


- Import `.m` and `.pq` source files together, by multi-select or from the selected folder.
- Duplicate source paths are ignored when files are added, including duplicates within the same multi-file batch; the first selected row is kept.
- Edit each row's enabled state, numeric order, query name, source path, and load mode. Use **Browse selected path** to change a source file.
- Review the final import sequence before Excel is changed. Duplicate order numbers are called out and sorted by case-insensitive filename, then path.
- Update an existing query formula in place or skip it. Create a collision-safe, hash-verified timestamped workbook backup by default; the final backup name is claimed with an atomic hard link so a concurrent backup cannot be overwritten. Filesystems that do not support same-volume hard links fail backup creation safely before Excel changes the workbook. Existing workbook changes are made against a sibling staged copy and committed by same-volume replacement only after Excel saves, closes, reopens read-only, verifies reviewed formulas/load mappings, and detects no concurrent edit.
- Excel opens suppress external link updates and force-disable programmatic VBA macros, then restore the owned Excel instance's original macro security setting during cleanup. Excel 4.0/XLM macros are an Office exception: Excel can still show an approval prompt even in force-disable mode. Power Query Manager does not approve or bypass that prompt; XLM-bearing workbook behavior is unverified. See [Microsoft's AutomationSecurity documentation](https://learn.microsoft.com/en-us/office/vba/api/excel.application.automationsecurity).
- Before import, refresh, and snapshot overwrite, Windows Restart Manager checks whether another process is using the workbook. Import and refresh repeat the check just before commit. If Windows cannot verify the file is unused, the operation stops. The check never closes another application; a small race remains if another app opens the workbook immediately after the final check. Excel 16.0 build 20430.0 was detected in a disposable COM smoke; a separate Windows subprocess test covers a held file without an Excel owner-lock file.
- Staged import and refresh reject paths where the workbook, stage, or any parent is a symbolic link or Windows reparse point, including a path changed before final commit. Select a path without reparse points; the app does not replace a link itself.
- Crash recovery: when a matching hidden `.pqstage-*` file is found beside the selected workbook, choose **Recover staged copy…**. Recovery refuses symbolic links and Windows reparse points in the workbook, stage, destination, or any parent path. The candidate list shows its age, size, modified time, and a preliminary ZIP/Open XML structure check. This check does not prove the workbook is valid or that an import committed. Excel still opens the candidate read-only; if that succeeds and the bytes remain unchanged, the manager saves a separately named copy. It never replaces the selected workbook or deletes the stage file. This preserves a candidate for manual inspection after a crash; it does not certify that a partial save contains the intended latest changes.
- If Windows prevents removal of a failed operation's temporary stage, the operation log reports the remaining stage filename. Use **Recover staged copy…** to inspect or save a verified copy; the manager does not claim that the stage was committed.
- Explicitly refresh selected or all existing Power Query connections after read-only inspection. The separate confirmation warns about external source access; background and synchronous modes, per-query timeout, per-query status, staged save/reopen verification, concurrent-edit detection, and default backup are available.
- Validate workbook and enabled source paths, file extensions, query names, duplicate names, and order before opening Excel.
- Record changes and attempt to roll back only query definitions created or updated by the current run if a later step fails.
- Save/load versioned profiles. Edit source-path and M-source environment mapping rules in the manager before saving; cancelled edits do not change active rules. Enabled source-root mappings use anchored path prefixes, show a before/after preview, and require approval; the most specific prefix wins and ambiguous ties are rejected.
- Legacy profiles without `schema_version` or with version `0` migrate only when they already match the known v1 profile structure. Loading does not rewrite the file; unknown versions or extra/ambiguous legacy fields are rejected.
- Profiles may also define opt-in M-code environment mappings. Loading one displays an exact diff before approval and writes changed formulas only to isolated temporary copies; originals remain untouched. The conservative tokenizer rewrites only literal first arguments of `File.Contents`, `Folder.Files`, `Folder.Contents`, `SharePoint.Files`, `SharePoint.Contents`, and `Web.Contents`. Dynamic expressions, ambiguous mappings, and unrecognized source functions are left unchanged with warnings. This is not a full M-language parser. CLI `plan` and `deploy` support these mappings with `--apply-environment-mappings`; the exact unified diff is printed, the user must type `APPLY M-CODE MAPPINGS`, and only temporary source copies are used. When source-root mappings are also enabled, those paths resolve first and the M diff is previewed against the mapped source file. Source-root mappings also require review and the exact `APPLY PATH MAPPINGS` phrase, including for path-only plans. Declining, including when the approval input reaches end-of-file, returns exit code 3 before Excel is started.
- Export selected/all workbook queries as `.pq` or as a versioned `.pqm.zip`. Package manifest v2 records exact formula hashes, stable order, conservative dependency/ambiguity information, and cycle paths; v1 archives remain importable and receive v2 dependency metadata in memory. Imported package queries enter the source grid for review before workbook compare and replace/skip.
- Export filenames shorten sanitized stems to 60 Unicode code points to fit common filesystem limits; manifests retain the complete query names.
- The query inspector provides line-numbered source/workbook panes, basic M token colors, added/removed diff colors, and search in all three panes. Syntax coloring is lexical only. Editing changes an in-memory buffer; **Save edited source copy…** writes a new file, and only the later import operation can change the workbook.
- Export selected/all workbook queries as a versioned `.pqm.zip` package with original names/order, exact formula SHA-256, conservative dependency/ambiguity metadata, and detected cycle paths. Use **Compare package versions…** to select a current and updated package for a read-only added/modified/unchanged/removed/conflict preview with exact M diffs; removed rows never delete workbook queries, and package hashes are validated before comparison. Manifest v2 has a JSON Schema contract; import verifies formula hashes and recomputes dependency metadata before adding extracted sources to the editable manager list. Version-1 packages migrate to v2 dependency metadata in memory. Imported rows can be renamed/reordered/disabled and reviewed against the target workbook before the normal replace/skip import decision.
- Select **Dependency graph…** to view detected query-to-query references, disabled dependencies, and cycles. The bounded scanner observes M line terminators, Unicode regular identifiers, dotted regular identifiers, quoted-identifier character escapes, and hash-prefixed keywords. It avoids partial-name matches, flags dynamic/ambiguous syntax without guessed arrows, and never changes import order; it is not a complete M parser. A read-only text summary accompanies the visual graph for keyboard and assistive-technology access.
- Table cell edits, drag reordering, and selected-row removal support Undo/Redo (`Ctrl+Z` / `Ctrl+Y`). Bulk enable, order, destination, and removal changes undo as a single action; a canceled destination setup applies no partial row edits. Use **Dark theme** in the title row to switch contrast palettes. The initial window size is bounded by the available screen; accessibility and scaling still need the documented manual laptop matrix.
- Show **Query inspector** in the title row (or use **View → Query inspector**) to open a docked, read-only source/workbook M preview for the selected row. It is hidden by default, follows the table selection, reports when a query or formula is unavailable, limits each preview to 100,000 characters, and links to the full compare/editor.
- Failure diagnostics are written locally to `%LOCALAPPDATA%\PowerQueryManager\logs\errors.jsonl` and rotated at 1 MiB (three backups). Records include timestamp, operation, exception type, and a fixed category only; raw error text, workbook paths, query formulas, and credentials are not stored. User-facing failure details redact URI passwords, Authorization and Proxy-Authorization header values through the end of the header line (including semicolon-containing Digest parameters), common signed URL secrets such as `sig` and `X-Amz-Signature`, quoted JSON token fields such as `refresh_token` and `id_token`, and compound key names such as `AWS_SECRET_ACCESS_KEY` and `x_api_key`. The writer refuses diagnostic log files, rotation targets, or parent directories containing symbolic links or Windows reparse points. If the user profile is unavailable, redirected, or unwritable, the operation error is still returned and logging is skipped.
- Error dialogs and CLI failures classify common timeouts, unsupported Office features, changed/locked workbooks, backup/verification failures, query conflicts, source drift, invalid inputs, and missing explicit approval, with a short recovery step. Full details remain in the operation output. Import failures identify each query as failed, rolled back, or not attempted; the local diagnostic file stores no raw error text.
- Optional CLI: `power-query-manager-cli inspect --workbook target.xlsx` inventories formulas and observed load states read-only; `--output inventory.json` writes to a new path. `power-query-manager-cli validate --profile deployment.json` validates the profile and enabled `.m`/`.pq` source paths offline. `power-query-manager-cli export --workbook target.xlsx --output queries.pqm.zip --package` writes a versioned package; omit `--package` to export `.pq` files to a new or empty folder. `power-query-manager-cli import --dry-run --package queries.pqm.zip --workbook target.xlsx` validates and compares without writing; for apply, use `--apply`, review the plan, and type `IMPORT`; noninteractive use supplies `--confirm IMPORT`. Import safely extracts packages with a 32 MiB compressed-archive limit, 25 MiB expanded-size limit, and per-entry compression-ratio checks; it defaults to Skip for existing names and uses definition-only loads. `--policy replace` opts into replacing formulas. Package imports never refresh queries. `power-query-manager-cli plan --dry-run --profile deployment.json --workbook one.xlsx --workbook two.xlsm` creates read-only schema-v1 plans. `power-query-manager-cli refresh --apply --all --workbook target.xlsx` first shows the selected query names/load states, warns about external source access, and requires typing `REFRESH`; it stages the workbook, creates a verified backup, and reports per-query results. Use repeated `--query NAME` to select queries, `--synchronous` for synchronous Excel mode, or `--timeout SECONDS` (1–7200; default 300). For noninteractive use, add `--confirm REFRESH`; without it, the CLI reports that approval is required and exits 3. A wrong supplied phrase is treated as a decline. `power-query-manager-cli deploy --profile deployment.json --workbook one.xlsx --workbook two.xlsm --apply` preflights every target before writes, presents the target/action summary, then requires typing `DEPLOY`; noninteractive use supplies `--confirm DEPLOY`. Missing approval is reported separately from an explicit decline. `--output report.json` writes a per-target report only to a new path. Each workbook has its own staged transaction and backup; the batch is not atomic, and individual failures are reported without claiming full success. `--apply-path-mappings` previews exact mapped source paths before confirmation; the required phrase is `APPLY PATH MAPPINGS` (or `--confirm "APPLY PATH MAPPINGS"` for noninteractive runs). For M-code environment mappings, add `--apply-environment-mappings` to `plan` or `deploy` to print the exact unified diff and require typing `APPLY M-CODE MAPPINGS`; noninteractive use supplies `--confirm "APPLY M-CODE MAPPINGS"`; the approved plan/deployment uses temporary source copies. A profile with `dry_run=true` only plans and never applies.
- CLI exit codes are 0 for complete, 2 for validation/output error, 3 for unsupported, approval-required, or declined action, 4 for Excel/COM preflight error, and 5 for failed/partial operation.
- Open **History…** in the operation log to review locally recorded comparisons, imports, refreshes, exports, profile, CLI plan/deploy, and restore actions. CLI events contain action/status and aggregate counts only; formulas, query names, source paths, and workbook paths are excluded. History uses a local SQLite file under `%LOCALAPPDATA%\PowerQueryManager\history.sqlite3`, refuses redirected database/parent paths containing symbolic links or Windows reparse points, and redacts recognized secret fields, omits formula/code fields, and supports a retention period from 1 to 3650 days, deleting one entry, or clearing all entries. Workbook snapshots are OFF by default. When enabled, full `.xlsx`/`.xlsm` files are copied into the adjacent local `history_snapshots` folder before import and refresh; the snapshot contents are flushed before exclusive hard-link publication (an existing UUID destination is never overwritten), and the snapshot index row and audit event commit in one SQLite transaction. Snapshot publication holds an immediate SQLite writer lock across the exclusive link and index commit; concurrent manager startup waits before scanning for unindexed files, preventing it from recovering a live snapshot. These files can contain sensitive workbook content. Select two snapshots and choose **Compare selected snapshots…** to compare query names, M formulas, and detected load destinations. Excel opens both verified snapshots read-only on a COM worker thread; the operation never refreshes external data, and inaccessible formulas or ambiguous duplicate names are marked unknown/conflicting. Snapshot files are hash-checked, governed by retention and explicit history deletion, and can be restored from a selected snapshot. Restore requires closing the workbook in Excel; overwriting is limited to its recorded original path, checks the current workbook hash, and creates a verified timestamped backup. On Windows, restore publication and verification rollback use the same existing-target ReplaceFileW safety helper as import and refresh; if the selected path disappears after review, the app refuses to recreate it and retains/reports the verified backup. Restore to a new file is the default when the original path is unavailable. If the local history database is corrupt or has an incompatible schema, the app preserves the original SQLite file and any WAL/journal sidecars under timestamped `history.corrupt-*` names, starts a fresh history store, and reports the archive path in the operation log. If Windows or the process stops after a snapshot file is published but before its history row is committed, startup moves the unindexed manager-named file to `%LOCALAPPDATA%\PowerQueryManager\history.sqlite3.recovered` without deleting it, and reports that folder in the operation log for manual review; recovered files are not automatically treated as valid history snapshots. A crash in the brief hard-link publication window can also leave a recent `.pqm-snapshot-*` alias; startup leaves recent in-progress names untouched while preserving the published UUID file for manual review. It does not silently discard the old database or automatically merge damaged entries. If recovery reports a stale lock, close every Power Query Manager instance first; after confirming no instance is recovering, remove only the specific `.recovery.lock` file named in the error and retry. Never remove or move the history database or its `-wal`, `-shm`, or `-journal` sidecars as lock cleanup.
- In the query inspector, **Edit declared parameter…** is enabled only for a standalone Text, Number, or Logical literal followed by `IsParameterQuery=true` and a matching `Type` metadata value. It changes only that literal in the editor; use **Save edited source copy…** to write a new file. Comments, UTF-8 BOM, and original newline style are preserved. Other types, expressions, or ambiguous metadata stay read-only in the parameter control; the regular M editor remains available for explicit manual edits.

## Supported input and ordering

Both extensions contain standalone Power Query M expressions; the complete file text is supplied as the query formula. M-code rewriting is limited to the explicit environment-mapping behavior described above; there is no complete M-language parser or evaluator. Files intended to be a shared `section` document or containing multiple query definitions have not been verified as an input format; use one standalone query expression per file.

A filename prefix is recognized only when one or more digits are followed by `_`, `-`, or whitespace. For example `1_qQueryName.pq`, `02-Helper.m`, and `003 Main.pq` start with orders 1, 2, and 3, and the recognized prefix is removed from the initial query name. A name such as `2foo.m` has no order prefix and keeps `2foo` as its initial query name. Unprefixed files are assigned after prefixed files in case-insensitive filename order. Equal order values remain visible as ties; filename and path provide the deterministic tie-break. A manually edited order is kept when more files are added.

## Excel load behavior and safety

**Definition only** is the default and creates or updates a `Workbook.Queries` definition. **Connection only**, **New worksheet**, **Existing worksheet**, and **Data Model** modes are available only when that mode passes a disposable-workbook probe on the current Office version/build. The probe uses only inline static `#table` M and no external data. Worksheet/model loads evaluate the selected query after the UI's explicit warning and confirmation; each operation is bounded and runs in a stage. Existing worksheet loads are accepted only on a provably empty worksheet; repeated imports preserve and refresh a matching destination, while unknown or mismatched mappings fail closed. Create/reimport/reopen checks passed on Excel 16.0 builds 20326.0 and 20430.0 for the cases listed in release evidence. Read-only inventory links a `Query - <name>` connection to ListObject QueryTables and Data Model `InModel` state; incomplete enumeration remains unknown. Broader workbook mapping and other Office builds remain unverified.

After inspecting the workbook, use **Refresh selected…** or **Refresh all…** to refresh existing Power Query connections. Refresh is never run during import, inspection, or M validation. The dialog requires a second explicit confirmation because Excel may contact external data sources; it offers a per-query timeout, default-on timestamped backup, and background/synchronous Excel modes where the load API supports them. A workbook with one Data Model table uses Excel's synchronous `ModelTable.Refresh()` API. If it contains multiple model tables, a partial selection is refused because Excel may contact other model sources; **Refresh all** is allowed only when every model table maps uniquely to a selected Power Query. That path uses Excel's synchronous model-wide API. The operation runs in the manager worker and the report explains the synchronous-only limitation. Refresh runs against a same-directory staged workbook, is reopened read-only to verify query formulas/connections, and is committed only when no concurrent change to the original workbook is detected. Definition-only queries without connections are reported skipped. Packaged loopback refresh for connection-only, worksheet-loaded, and Data Model queries passed on Excel 16.0 build 20430.0; connection-only timeout rollback also passed. Other Office builds, external providers, provider prompts and the remaining acceptance matrix remain unverified.

Before any Excel COM instance is created, the workbook path check reads the Windows `Zone.Identifier` alternate data stream when present. Internet/Restricted zone marks (ZoneId 3 or higher) and uninterpretable metadata are refused; the app never removes the mark or changes trust settings. This is conservative fail-closed handling, not certification of Excel Protected View UI, organizational zone policy, or other protected-session behaviors.

The app uses a separate, hidden desktop Excel process. Bounded calls target only that process and wait up to five seconds for its exit; if termination fails, the operation reports that cleanup is unconfirmed and asks you to verify the app-owned Excel window. Close the selected workbook in Excel before import or refresh. A verified backup is enabled by default and saved next to the workbook. The installer does not directly edit `.xlsx` package files; definition/connection imports do not evaluate M, worksheet and Data Model load selections evaluate M after confirmation, and the separately confirmed refresh action refreshes existing connections.

## Requirements

- Windows 10 or Windows 11.
- Desktop Microsoft Excel installed and activated; Excel COM must be available to the current Windows user.
- Python 3.11, 3.12, or 3.13 for building/running from source, plus PowerShell.
- The Python process and Excel may use different bitness because Excel automation is out-of-process COM; if Excel activation fails, retry with Python matching the installed Office architecture and inspect the logged COM error.

## Build on Windows

1. Extract this ZIP to a writable local folder.
2. Open PowerShell in the extracted `pq_installer` folder.
3. Run:

   ```powershell
   Set-ExecutionPolicy -Scope Process Bypass
   .\build.ps1
   ```

   The script creates `.venv`, installs pinned packages, runs tests, and builds a one-folder distribution with PyInstaller. Packaging fixes `PYTHONHASHSEED` and `SOURCE_DATE_EPOCH`; two clean builds with the same pinned toolchain on the same Windows build were byte-for-byte identical across all 227 onedir files. Reproducibility across different Windows hosts, Python patch builds, or toolchains is not claimed. The launcher will be at `dist\PowerQueryManager\PowerQueryManager.exe`; keep the adjacent `_internal` directory with it. Do not distribute the EXE by itself.

To run from source after the build:

```powershell
.\.venv\Scripts\python.exe -m pq_installer.main
```

## Windows acceptance checks before release

1. Run `.uild.ps1`; confirm tests pass and the executable and `_internal` folder exist.
2. Launch the packaged EXE on Windows with desktop Excel installed. Confirm the main window opens and can be resized to a narrower laptop window.
3. On a disposable `.xlsx` workbook, import a mixed `.m`/`.pq` set, including a query that references a later query. Confirm the definitions appear under **Data > Queries & Connections**.
4. Test replace and skip behavior, and confirm replacing an existing query retains its previous worksheet load destination.
5. Force a controlled failure partway through import; confirm only this run's query changes are rolled back, pre-existing workbook objects stay intact, and the backup remains usable.
6. Confirm no hidden Excel process remains after success or failure.

## Source package contents

- `src/`: app entry point, UI, validation, ordering, backup, and COM installer.
- `tests/`: portable unit and fake-COM tests.
- `pyproject.toml`, `requirements*.txt`: package metadata and pinned build/runtime tools.
- `PowerQueryInstaller.spec`: one-folder PyInstaller configuration.
- `build.ps1`: fail-fast Windows build and test command.
- `README.md`: requirements, operation, limitations, and acceptance steps.

### 2026-10-22 snapshot retention/path regression and package refresh

- Focused history tests: **22 passed**. Full suite: **260 passed**. Contract validation, compilation, onedir build, CLI help and `dist-next` package verification passed.
- Updated source and packaged manual text to explain retention-before-restore and path-contained snapshot cleanup. No Excel operation was run for this history-only behavior; v2 remains partial and production release remains gated.


### 2026-10-22 local history recovery

- Corrupt or incompatible local history is now preserved under a timestamped `history.corrupt-*` archive before a fresh store is created; SQLite sidecars are preserved with it and the UI reports the archive path. Operational I/O errors do not trigger quarantine.
- The test suite covers corrupt bytes, incompatible schema, sidecar preservation, operational-error fail-closed behavior, and visible recovery reporting. History entry salvage and automatic merge are not implemented; release remains partial.


### 2026-10-22 Windows Internet-zone workbook preflight

- Shared workbook validation refuses Internet/Restricted `Zone.Identifier` marks (ZoneId 3+) and malformed/conflicting zone metadata before inspector/import creates Excel. Local Intranet ZoneId 2 remains allowed. No zone stream is modified.
- Test-first portable/Windows regressions and an Excel-created disposable workbook smoke passed. This blocks a marked file before COM can bypass normal Protected View flow; actual Protected View UI remains unverified.


## Portable multi-workbook deployment bundle

Use `PowerQueryManagerCLI.exe bundle --profile deployment.json --output release.pqmdeploy` to create a portable release containing the versioned profile and exact `.m`/`.pq` source bytes. Source hashes and bundle paths are checked when the archive is extracted. The target workbook path is replaced with a placeholder so the creator's local workbook path is not embedded. Existing output files are never overwritten.

If the profile enables source-path mappings, add `--apply-path-mappings`; the CLI prints the resolved source paths and requires typing `APPLY PATH MAPPINGS` before embedding those mapped sources. M-code environment mapping rules stay in the profile and still require explicit diff review at plan/deploy time.

Plan or deploy from the portable archive by passing it to the existing `--profile` option, for example `PowerQueryManagerCLI.exe plan --dry-run --profile release.pqmdeploy --workbook one.xlsx --workbook two.xlsm` or `PowerQueryManagerCLI.exe deploy --apply --profile release.pqmdeploy --workbook one.xlsx --workbook two.xlsm`. Apply requires the exact `DEPLOY` confirmation. For a terminal session, type it at the prompt; for a deliberate non-interactive invocation, pass `--confirm DEPLOY`. Source-path and M-code mapping reviews retain their distinct phrases (`APPLY PATH MAPPINGS` and `APPLY M-CODE MAPPINGS`) and require a separate `--confirm` for each review, in prompt order. Each workbook is committed independently with its own staged transaction and backup.


### 2026-10-23 — packaged GUI scale acceptance

- The current published executable passed disposable import and UI Automation at Qt scale factors 1.25 and 1.5 on Windows 10 build 26300 / Excel 16.0 build 20430.0. Exact M readback, verified backup, read-only reopen, exit 0 and owned-process cleanup passed. Screen-reader and physical Windows display-scaling acceptance remain open.

### 2026-10-23 — repeatable W203 package build

- Two clean Windows `app/build.ps1` runs passed 426 tests each and produced identical SHA-256 values for all 227 onedir runtime files on Windows 10 build 26300 / Python 3.11.9 / PyInstaller 6.13.0. Cross-host reproducibility is not established.

### 2026-10-23 — current-package replace/skip and mixed-source acceptance

- The rebuilt GUI passed existing-query replacement, preserve-existing connection-only replacement, skip no-op, and mixed numbered `.m`/`.pq` profile import on Excel 16.0 build 20430.0. Skip retained the exact workbook hash/timestamp and created no backup/stage. The mixed profile was serialized in reverse order but imported in numeric-prefix order with exact formulas.


Current source rebuild check (W247): after the query-inspector keyboard change, a clean isolated PyInstaller build matched all 227 onedir runtime files byte-for-byte on Windows 10 build 26300 / Python 3.11.9 / PyInstaller 6.13.0. Cross-host reproducibility is not claimed.

