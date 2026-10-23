# Excel COM spike log

## 2026-10-22 — W73 packaged GUI all-skip no-op

- Host: Windows 10 build 26300; Excel 16.0 build 20430.0; current one-folder PyInstaller GUI; Python UIA driver.
- `tools/run_packaged_gui_import_smoke.py --skip-existing` seeded a disposable `.xlsx` with a static definition-only `GuiSmoke` query, then loaded a versioned profile configured to skip existing names while its source file contained a different formula. The packaged worker's read-only inventory detected the existing query and bypassed backup, history snapshot, and workbook writer.
- PASS: the UI showed “No changes”; original formula, worksheet marker, target SHA-256, modification timestamp, and read-only attribute were unchanged; no backup or stage existed; manager exited 0 and Excel process set returned to baseline. Existing Excel PID 13156 was untouched. No refresh or external data source was used.
- This verifies the all-skip path on a read-only workbook only when every enabled name exists. Mixed skip/create, workbook-open races, and other Office versions remain open.

## 2026-10-22 — W71 packaged VBA-bearing GUI import rerun

- Host: Windows 10 build 26300; Excel 16.0 build 20430.0; one-folder PyInstaller package; Python UIA driver.
- `tools/run_packaged_gui_import_smoke.py --vba` created an isolated disposable `.xlsm` with a harmless `ThisWorkbook.Workbook_Open` event that would change A1. The packaged GUI imported one static definition-only query. Read-only reopen confirmed exact M text, original-hash backup matched, `xl/vbaProject.bin` and event source were retained, and the marker did not change. Manager exit code was 0 and smoke-owned Excel process set returned to baseline.
- No external refresh or external data source was used. This repeats W64 on the same build; it does not generalize to other VBA project forms, XLM, external connectors, Protected View, or additional Office builds. Existing Excel PID 13156 was not touched.

## Macro-security boundary (documented, not XLM-tested)

Every workbook open path uses `UpdateLinks=0` and sets Excel `Application.AutomationSecurity` to `msoAutomationSecurityForceDisable` before `Workbooks.Open`; the prior setting is restored during cleanup of the owned Excel process. Microsoft documents that this disables VBA macros opened programmatically, but **does not disable Excel 4.0 (XLM) macros**: Excel prompts the user to decide whether to open such a file. This app does not approve or bypass that prompt. Excel 4.0/XLM-bearing workbooks have not been tested in the acceptance matrix; treat the open as unverified and cancel any unexpected prompt. [Microsoft: Application.AutomationSecurity](https://learn.microsoft.com/en-us/office/vba/api/excel.application.automationsecurity)

## Windows active-workbook preflight

- API: Windows Restart Manager `RmStartSession` → `RmRegisterResources` for the exact workbook path → `RmGetList` → `RmEndSession`. This is query-only; Power Query Manager never calls Restart Manager shutdown/restart APIs and never closes an external process. Microsoft documents the registered-file resource and process-list behavior in [RmRegisterResources](https://learn.microsoft.com/en-us/windows/win32/api/restartmanager/nf-restartmanager-rmregisterresources) and [RmGetList](https://learn.microsoft.com/en-us/windows/win32/api/restartmanager/nf-restartmanager-rmgetlist).
- Import checks before work begins and again immediately before staged commit; explicit refresh checks before refresh and before commit; snapshot overwrite checks before staging and before replacement. A Restart Manager error blocks the write. The check supplements Excel's `~$` owner file, read/write probe, workbook hash and atomic replace; it narrows but cannot eliminate a race where another process opens the workbook after the final check.
- Portable tests cover injected active-user and API-error outcomes. A Windows subprocess test opens a disposable workbook without an Excel owner file; Restart Manager detects the process and write preflight refuses the workbook. This proves Windows API integration for a held file, not Excel-specific lock behavior or OneDrive/Protected View behavior.
- Real Excel smoke: Windows 10 build 26300, Excel 16.0 build 20430.0, Python 3.11.9 x64. On a COM-initialized worker thread, an owned `DispatchEx` instance created and reopened a disposable `.xlsx`; Restart Manager returned `Microsoft Excel` for that path. The owner lock was present in this real-Excel run. Excel was closed and its exact PID exit was confirmed; the disposable folder was removed. The owner-lock-free case is covered separately by the Python subprocess regression, not claimed as an Excel-specific scenario.

## 2026-10-21 — latest packaged CLI package import

- Host/runtime: Windows 10 build 26300; Excel 16.0 build 20430.0; Python 3.11.9 x64; PyInstaller 6.13.0; PySide6 6.8.3.
- Method: launched `dist/PowerQueryManager/PowerQueryManagerCLI.exe import --apply` in an interactive pseudo-terminal against a newly created disposable `.xlsx` and two static definition-only package entries. Typed the required `IMPORT` confirmation. The CLI Restart Manager preflight ran before staged import; a separate COM worker reopened the committed workbook read-only for formula verification.
- Observation: exit code 0; exact `Sales` and `Inventory` M formulas read back; timestamped backup SHA-256 matched the pre-import workbook (`e63dffa447db5f6c4059063b321b919213318505a34ad9eb81735121091ddc5a2`); CLI and inspection owned Excel processes exited; no external source evaluation occurred. The disposable directory was deleted by the smoke harness.
- Limitations: definition-only package import into a blank `.xlsx` only; no replacement/skip against existing loads, `.xlsm`, external sources, or packaged GUI import.

## 2026-10-21 — packaged multi-workbook CLI dry-run

- Host/runtime: Windows 10 build 26300; Microsoft Excel 16.0 build 20326.0; Python 3.11.9 x64, pywin32 310; PyInstaller 6.13.0.
- Method: created two disposable blank workbooks (`.xlsx` and `.xlsm`) with owned Excel automation; ran the built `PowerQueryManagerCLI.exe plan --dry-run` with one UTF-8 `.pq` source and both workbooks as separate `--workbook` arguments. The CLI invoked `ExcelQueryInspector` on its dedicated COM worker thread. It produced two plan-v1 entries with action `create` and wrote the bundle to a new temporary output path.
- Observation: exit code 0, no stderr; both original workbook SHA-256 values were identical before and after. No workbook was saved or refreshed by the planner, and the M formula was static with no external source. Test fixture folder and output were removed after readback.
- Limitations: only blank disposable workbooks and definition-only plans were exercised; does not certify existing-query replace/skip plans through real Excel, loaded destination mappings, workbook lock scenarios, or actual deployment/apply. CLI intentionally has no apply or refresh command.

## 2026-10-21 — read-only query inventory

- Host/runtime: Windows 10 build 26300; Microsoft Excel 16.0 build 20326.0; Python 3.11.9 x64, pywin32 310.
- Method: opened a disposable `.xlsx` through `ExcelQueryInspector` in a separate owned Excel process with `ReadOnly=True` and external link updates disabled; read its query formula and hash before/after inspection.
- Observation: query formula was returned; load mapping was reported `unknown`; workbook SHA-256 was unchanged. The owned Excel instance was closed. No save or refresh was requested.
- Limitations: the blank/simple disposable workbook does not establish load mapping for arbitrary existing worksheet/Data Model connections, nor `.xlsm`, locks, or synced workbook behavior.

## 2026-10-21 — connection-only capability and staged import

- Host/runtime: Windows 10 build 26300; Microsoft Excel 16.0 build 20326.0; Python 3.11.9 x64, pywin32 310.
- Method: `probe_excel_capabilities()` created a temporary `.xlsx`, added the static scalar M formula `let Source = 1 in Source`, created a connection with `Workbook.Connections.Add2` using `OLEDB;Provider=Microsoft.Mashup.OleDb.1`, `Location=PQM_Capability_Probe`, command text `SELECT * FROM [PQM_Capability_Probe]`, command type 2, `CreateModelConnection=False`, and `ImportRelationships=False`; saved, closed, reopened read-only, and verified the definition and connection persisted. It found no worksheet ListObjects and `InModel=False`.
- Import smoke: a separate disposable `.xlsx` was staged through `ExcelQueryInstaller` with `let Source = 41 in Source` and `LoadMode.CONNECTION_ONLY`; after commit, a separate worker thread initialized COM and opened the committed workbook read-only. It observed the exact formula, connection `Query - Probe`, `InModel=False`, and zero worksheet ListObjects. No refresh was called and no external data source was contacted. The importer and inspection process both closed their owned Excel instances.
- Outcome: the UI may offer **Connection only** only after this per-build probe succeeds; the installer also requires the probe attestation and checks the Excel version/build before applying. Current observed support: definition-only and connection-only on Office 16.0 build 20326.0.
- Limitations: this is one build and a blank disposable `.xlsx`. It does not certify arbitrary existing connection preservation, `.xlsm`, worksheet or Data Model loading, refresh, cancellation/timeouts, locked/shared workbooks, or failure recovery against all Office states. Connection-only import does not refresh the query.

## 2026-10-21 — definition-only staged import

- Host: Windows 10 build 26300; Microsoft Excel version 16.0, build 20326.0; Python 3.11.9 x64, pywin32 310.
- Method: create a disposable `.xlsx` using an owned `DispatchEx` Excel instance; call `ExcelQueryInstaller` with a disposable UTF-8 `.m` file containing `let Source = 41 in Source`; importer copies to sibling stage, opens with `UpdateLinks=0`, adds `Workbook.Queries` definition, saves, closes, and atomically replaces the disposable file; reopen read-only in a second owned Excel process and inspect `Workbook.Queries.Item('Probe').Formula`.
- Observation: result `Created`; reopened formula exactly equaled `let Source = 41 in Source`. Both Excel PIDs created by this smoke were gone afterward.
- Limitations: this verifies definition-only on this one Excel build and a disposable blank workbook. It does not certify connection-only, worksheet/Data Model loads, preservation of existing load destinations, refresh, macros, protected workbooks, OneDrive/SharePoint locks, or failure recovery on Excel.

## 2026-10-21 — explicit staged query refresh

- Host/runtime: Windows 10 build 26300; Microsoft Excel 16.0 build 20326.0; Python 3.11.9 x64, pywin32 310.
- Method: created disposable `static-refresh.xlsx` with `Workbook.Queries.Add('PQM_Refresh_Smoke', 'let Source = #table({"Value"}, {{1}}) in Source')` and `Workbook.Connections.Add2` using `OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=PQM_Refresh_Smoke;Extended Properties=""`, command `SELECT * FROM [PQM_Refresh_Smoke]`, command type 2, `CreateModelConnection=False`, `ImportRelationships=False`. The test M expression contains only an inline constant table and has no external source. Ran `ExcelQueryRefresher.refresh` twice with a 60-second timeout and backup enabled: background mode set `OLEDBConnection.BackgroundQuery=True` and polled `OLEDBConnection.Refreshing`; synchronous mode set it false and waited for `WorkbookConnection.Refresh` to return. Both modes used same-directory staged copies, reopened the saved stage read-only to verify the exact M formula and connection, then replaced the disposable target. Checked the backup hash and owned Excel process exit after each run.
- Observation: both refresh modes returned `refreshed`; each backup SHA-256 matched its pre-refresh target; reopened formula matched exactly and load status was `connection_only`; both owned Excel processes exited.
- Limitations: only an inline static M table was exercised. No external connectors/authentication, loaded worksheet/Data Model result, timeout-triggered cancellation or real synchronous timeout termination, locked workbook, or repeat-refresh behavior on a user workbook was tested. The real smoke does not certify refresh for arbitrary workbook connections. Unit tests cover timeout cancellation and owned-PID targeting; those timeout paths were not induced against real Excel.

## 2026-10-21 — worksheet and Data Model application load/reimport smoke

- Host/runtime: Windows 10 build 26300; Microsoft Excel 16.0 build 20326.0; Python 3.11.9 x64, pywin32 310. Disposable workbooks only.
- Worksheet method: added static query `PQM_Sheet_Spike` with `Workbook.Queries.Add`; called `Worksheet.ListObjects.Add(SourceType=3, Source=<Mashup OLEDB connection string>, LinkSource=True, XlListObjectHasHeaders=1, Destination=Range('B3'))`; set `QueryTable.CommandType=2`, `CommandText='SELECT * FROM [PQM_Sheet_Spike]'`, `BackgroundQuery=False`; refreshed. This produced one ListObject, range `$B$3:$B$5`, values `Value, 1, 2`. Renaming the generated `QueryTable.WorkbookConnection.Name` to `Query - PQM_Sheet_Spike` persisted after save/close/read-only reopen with the formula, range and values intact. `QueryTables.Add` alone did not create a ListObject in the earlier probe.
- Data Model method: added static query `PQM_Model_Spike`; called `Workbook.Connections.Add2('Query - PQM_Model_Spike', ..., <Mashup OLEDB connection string>, 'SELECT * FROM [PQM_Model_Spike]', 2, True, False)` and `connection.Refresh()`. Observed `InModel=True`, one `ModelTable`, `RecordCount=2`, and `SourceWorkbookConnection='Query - PQM_Model_Spike'`; `InModel` and the model table persisted after save/close/read-only reopen.
- App probe: `tools/run_excel_load_smoke.py` ran the per-build `probe_excel_capabilities()` using separate static inline `#table` queries for definition-only, connection-only, new sheet, existing sheet at `B3`, and Data Model. It saved, closed, reopened read-only, and checked formulas, table connection mappings, exact worksheet location, `InModel`, and model-table source connection. Observed support on this build: all five modes.
- App import method: for each of new worksheet, existing worksheet, and Data Model, the smoke created a separate disposable `.xlsx`, performed create, same-formula reimport, then changed-formula reimport, and used read-only inspection after commit. The installer staged sibling copies, applied definitions before loads, bounded synchronous M evaluation to its owned Excel process, reopened the stage read-only to check exact M and load mapping, compared the original hash, and atomically committed. Repeat and changed-formula imports retained the same table/model association. The smoke reopened each output read-only and confirmed values `3, 4` plus an unrelated `PQM_Marker!A1` cell remained unchanged. It tracked and verified exit of every seed/import/inspect/value-check Excel PID.
- Observation: `new_sheet` reopened as `worksheet: PQ_PQM_New_Load!$A$1:$A$3`; `existing_sheet` reopened as `worksheet: Output!$B$3:$B$5`; `data_model` reopened as `data_model`. All three passed create/repeat/changed-formula import, result-value and unrelated-cell checks, read-only reopen verification, and owned-process cleanup. Earlier smoke runs exposed a lingering Data Model process; COM child proxy cleanup and per-operation PID tracking were added, and the final full smoke passed.
- Password modal check: `tools/run_excel_password_safety_smoke.py` created a disposable open-password `.xlsx`. Inspection and staged import each blocked for the smoke's 10-second timeout, terminated only their own Excel process, returned timeout errors, preserved the original workbook SHA-256, and removed the stage. Production open timeout is 60 seconds. This verifies fail-closed modal handling but does not certify password support.
- Limitations: static inline M only, one Office build, disposable `.xlsx` files, and existing-sheet output limited to a fully empty worksheet. External connectors, `.xlsm`, multi-query dependency refresh, diverse existing workbook styles/objects, timeouts induced during actual M evaluation, packaged GUI use, and other Office builds remain unverified. These results do not claim broad Excel compatibility.

## 2026-10-21 — two-workbook deployment apply smoke

- Host/runtime: Windows 10 build 26300; Microsoft Excel 16.0 build 20326.0; Python 3.11.9 x64; pywin32 310. Non-admin launch; local disposable workbooks.
- Method: `app\.venv\Scripts\python.exe tools\run_excel_deployment_smoke.py` created two blank `.xlsx` targets and used the deployment service on a COM worker thread. Each target received the same static inline M definition-only query. The smoke reopened outputs read-only, checked exact formulas and definition-only load state, validated each pre-change backup hash, repeated the batch, and tracked all Excel PIDs created by the smoke.
- Observation: first batch results were `applied`, `applied`; repeat results were `no_changes`, `no_changes`; two backups matched their original SHA-256 values; all nine owned Excel PIDs exited.
- Limitations: static inline M and definition-only loads only. This does not test failure isolation during real Excel deployment, worksheet/Data Model deployment, `.xlsm`, external connectors, shared/locked workbooks, or other Office builds. It is evidence for the simple two-target path, not a batch atomicity guarantee.

## 2026-10-21 — CLI selected refresh smoke

- Host/runtime: Windows 10 build 26300; Microsoft Excel 16.0 build 20326.0; Python 3.11.9 x64; pywin32 310. A disposable local `.xlsx` was used.
- Method: `app\.venv\Scripts\python.exe tools\run_excel_cli_refresh_smoke.py` created a query with inline static `#table` M and its Mashup OLE DB connection. The script invoked the actual CLI `refresh --apply --query ... --workbook ... --timeout 60` command twice, once in background mode and once with `--synchronous`. The test harness supplied the required typed `REFRESH` confirmation; no external source was present. Each operation used its staged workbook and default backup. The smoke checked the exact pre-refresh backup SHA-256, read-only reopened formula and `connection_only` state, and verified that new Excel processes exited.
- Observation: both CLI operations returned `refreshed`; the source formula and connection-only state survived read-only reopen; each backup matched the pre-operation file hash; no CLI-owned Excel process remained.
- Limitations: one Excel build and one static inline table only. Does not certify external connectors/authentication, timeout behavior, worksheet/Data Model refresh, mixed query outcomes, `.xlsm`, or packaged GUI refresh. W15 disconnected/external-source behavior remains open.

## 2026-10-21 — CLI verified query-package import smoke

- Host/runtime: Windows 10 build 26300; Microsoft Excel 16.0 build 20326.0; Python 3.11.9 x64; pywin32 310. A disposable local `.xlsx` and locally generated package-v2 `.pqm.zip` were used.
- Method: `python tools\run_excel_cli_package_import_smoke.py` created two standalone static M definitions in a hash-verified package and a blank workbook, then invoked CLI `import --apply --package ... --workbook ...` with the harness supplying the required typed `IMPORT` confirmation. The CLI used the safe bounded ZIP extractor, compared the target read-only, applied definitions only in a staged workbook, made a verified backup, and reopened the result read-only to compare exact names/formulas.
- Observation: both queries returned `Created`; read-only inventory matched names and formulas; backup SHA-256 matched the original workbook; no CLI-owned Excel process remained.
- Limitations: one Office build, one blank `.xlsx`, package-v2 source, definition-only load, and no existing-query replace/skip scenario. Does not certify `.xlsm`, arbitrary ZIP input beyond portable traversal/size tests, loaded query destinations, or packaged GUI import.
# Local history snapshot restore smoke — 2026-10-21

- Host: Windows desktop Excel 16.0, build 20430.0.
- Method: created a disposable `.xlsx` through an owned `DispatchEx` Excel instance with `A1="before"`; closed it; captured an opt-in local `HistoryStore` snapshot; reopened through a second owned Excel instance, changed A1 to `"after"`, saved and closed; restored the recorded snapshot with the current SHA-256 supplied; reopened read-only and inspected A1.
- Observed: A1 returned to `before`; restored workbook SHA-256 matched the stored snapshot; timestamped pre-restore backup matched the hash of the `after` workbook; all smoke-created Excel process IDs exited. Run command: `python tools/run_excel_history_restore_smoke.py`.
- Limits: this verifies raw workbook snapshot/restore on one `.xlsx` and one Excel build. It does not establish broad `.xlsm` coverage, GUI restore interaction, recovery under forced disk/permission failures, or release readiness. Snapshot files contain the complete workbook and remain local; capture defaults OFF.

# Existing query load and related workbook object preservation — 2026-10-22

- Host: Windows desktop Excel 16.0, build 20430.0.
- Method: created a disposable `.xlsx` containing ordinary range-backed `tblKeep` with style/filter, a workbook defined name, a cell formula and `ForceFullCalculation=True`; imported two static inline-M queries, one connection-only and one worksheet-loaded to `tblPqmReport`; added report formulas and a workbook defined name consuming that query table; changed both M formulas and replaced the definitions; then reopened read-only and checked objects, formulas, load states, and calculated report values.
- Observed: on Excel 16.0 build 20430.0, changed query outputs (30, 40) populated the same report table and both report formulas/name calculated to 70; `tblKeep` retained its name, range, row data, style and active filter; its defined name, cell formula and ForceFullCalculation were preserved; the connection-only query stayed connection-only. Excel reports ordinary tables as `SourceType=1` (`xlSrcRange`); their `QueryTable` property raises COM error `-2146827284`. Inventory and load-signature code skips only this explicit plain-range type and fails closed when other table links cannot be resolved. The first smoke exposed and fixed that mapping defect. A later run passed every assertion; the harness waited 20 seconds and then terminated only the exact smoke-tracked Excel PID after checking it was Excel, all workbooks were closed, and `Quit` had been called. Pre-existing Excel PID 13156 was untouched. Run command: `python tools/run_excel_preserve_smoke.py`.
- Limits: one `.xlsx`, one Excel build, one ordinary table, one connection-only query and one worksheet-loaded static query were exercised. This does not establish preservation for all external-table types, Data Model-loaded report links, `.xlsm`, privacy settings, or other Excel builds. The smoke-owned COM server did not exit gracefully within 20 seconds on the final run and was explicitly cleaned up by exact-PID identity; broader process-cleanup behavior remains open.


## 2026-10-21 — owned Excel process timeout cleanup

- Host/runtime: Windows 10 build 26300; Microsoft Excel 16.0 build 20430.0; Python 3.11.9 x64 with pywin32 310, non-admin local run.
- Method: `app\.venv\Scripts\python.exe tools\run_excel_timeout_cleanup_smoke.py` created one hidden `DispatchEx` Excel instance, captured its window PID and a synchronization handle, then exercised the bounded-call helper with a two-second wait and one-second timeout. It opened no workbook and did not enumerate or terminate unrelated Excel instances.
- Observation: the timeout helper targeted the captured owned PID, waited for process exit, closed its process handle, and the smoke independently observed that exact handle signaled. The run reported `PASS: owned Excel timeout cleanup; Excel 16.0 build 20430.0; Python 3.11.9`.
- Limits: the callback was a controlled sleep, not a blocked Excel COM call, modal dialog, or query evaluation. Real M-evaluation timeout behavior, failure to terminate under access denial, and full crash recovery remain unverified. Fake Win32 tests cover those cleanup-failure branches without killing an unrelated process.


## 2026-10-21 — packaged CLI import on Excel 16.0 build 20430.0

- Host/runtime: Windows 10 build 26300; Microsoft Excel 16.0 build 20430.0; Python 3.11.9 x64 with pywin32 310.
- Method: `python tools\run_excel_cli_package_import_smoke.py` created a disposable blank `.xlsx` and a `.pqm.zip` containing two static inline-M definitions. It launched the built `dist\PowerQueryManager\PowerQueryManagerCLI.exe`, reviewed the printed definition-only plan (`Sales`, `Inventory`), and supplied the required `IMPORT` token in the interactive PTY. No external source or refresh was involved.
- Observation: the packaged binary returned `applied`; read-only Excel inventory matched both names and exact formulas; the timestamped backup matched the pre-import SHA-256; the smoke observed all seed/import/inspection Excel processes it owned exit.
- Limits: definition-only package import and two static M formulas only. This does not verify packaged GUI import, other load modes, `.xlsm`, external data, or mixed deployment failure.


## 2026-10-21 — abandoned staged workbook recovery

- Host/runtime: Windows 10 build 26300; Microsoft Excel 16.0 build 20430.0; Python 3.11.9 x64; PySide6 6.8.3 and pywin32 310.
- Method: `python tools\run_excel_recovery_smoke.py` created a disposable `.xlsx` with one static inline-M query through owned `DispatchEx`, closed it, copied it to the hidden `.pqstage-<uuid>.xlsx` pattern used by import/refresh, and ran `StagedRecoveryWorker`. The worker used real read-only Excel inspection, then atomically published to a new output path; the recovered copy was reopened read-only and the exact query name/formula checked.
- Observation: PASS; Excel reported one query and the exact formula; the selected original SHA-256 remained unchanged, the candidate stage remained present, the output SHA-256 matched the candidate, and the seed plus both recovery-owned Excel PIDs exited.
- Limits: this validates query-bearing `.xlsx` copy-out, not a forced process crash, a partially written Excel save, `.xlsm`, a GUI-driven recovery click sequence, or every interruption phase. Recovery saves a separate copy and does not decide whether its contents should replace the selected workbook.
## CLI M-code environment mapping integration smoke — 2026-10-21

- Host: Windows 10 build 26300; Microsoft Excel version 16.0, build 20430.0; Python 3.11.9; PySide6 6.8.3.
- Procedure: `python tools/run_excel_cli_environment_mapping_smoke.py` created a disposable `.xlsx`, reviewed an exact mapping from `C:\old` to `D:\new`, required the mapping approval and deployment confirmation callbacks, deployed two definition-only queries, then reopened read-only and compared exact formulas. One formula included `File.Contents` but the query stayed definition-only; Excel did not evaluate the expression and no external path was contacted.
- Observed: mapped formula matched the approved diff; the second query formula was unchanged; both original source files retained their pre-run SHA-256; deployment generated a verified backup and reported `applied`; two owned Excel PIDs exited before the smoke returned PASS.
- Limits: this verifies the CLI review/materialization/deployment integration only on this Excel build and for definition-only queries. It does not certify mapping for worksheet/Data Model loads, external data access, other Office builds, packaged CLI interactive input, or Excel evaluation of the mapped external source.

## Mixed read-only multi-workbook deployment — 2026-10-21

- Host: Windows 10 build 26300; Microsoft Excel 16.0, build 20430.0; Python 3.11.9 x64; pywin32 310.
- Procedure: `python tools/run_excel_deployment_mixed_smoke.py` created two disposable `.xlsx` targets with Excel, recorded their SHA-256 values, then set the second target read-only using the Windows file attribute. It ran `deploy_multi_workbook` with a static definition-only query and a real read-only Excel inventory provider.
- Observed: first target status `applied`, second `blocked_preflight`; first target query formula read back exactly and its backup matched the original hash. The read-only target received no backup and remained byte-identical. Both owned Excel processes exited before PASS.
- Code change: non-dry-run deployment now validates per-target write access before read-only inventory, backup, or installer execution. This lets a writable target proceed while an unwritable target is isolated during preflight. Dry-run planning remains read-only and does not require write access.
- Limits: verifies Windows read-only attribute handling on this build. ACL-restricted folders, Excel-open targets, mid-commit COM failures, and broader filesystem types remain unverified.

## Packaged query import with a Unicode name — 2026-10-21

- Host: Windows 10 build 26300; Excel 16.0 build 20430.0; Python 3.11.9 x64; packaged CLI from the current onedir build.
- Procedure: the interactive packaged CLI import smoke created a versioned package with `Sales`, `Inventory`, and `Équipe – Résultats 2026 (Q1)`, imported to a new disposable `.xlsx`, typed `IMPORT`, then reopened read-only.
- Observed: CLI returned `applied`; all three names and exact formulas read back, including Unicode, en dash, spaces, parentheses; pre-import backup hash matched and smoke-owned Excel processes exited.
- Limits: this verifies the named W17 live-workbook export/package/CLI-import path on one Office build. It does not test every Excel-special punctuation character, `.xlsm`, existing-name replacement/skip, other Office builds, or the GUI path.


## Combined source-path and M-code mapping — 2026-10-22

- Host: Windows 10 build 26300; Excel 16.0 build 20430.0; Python 3.11.9 x64; PySide6 6.8.3.
- Procedure: `python tools/run_excel_cli_environment_mapping_smoke.py --packaged` created separate `sources/` and `mapped/` folders with disposable `.pq` files, then a blank `.xlsx`. CLI deployment reviewed and approved path-root mapping (`sources` to `mapped`) first, then previewed and approved M literal mapping (`C:\old` to `D:\new`), and deployed both queries definition-only.
- Observed: deployment report path preview resolves to `mapped/first.pq`; M preview source path is that same mapped file. Excel read back the rewritten first formula and untouched second formula exactly. Hashes for both original and mapped source files stayed unchanged; the backup was verified and both owned Excel processes exited.
- Limits: one build and definition-only loads. Worksheet/Data Model evaluation, external source evaluation and alternate Office builds remain outside this smoke.

## History snapshot inventory comparison — 2026-10-22

- Host: Windows 10 build 26300; Excel 16.0, build 20430.0; Python 3.11.9 x64; pywin32 310.
- Procedure: `python tools/run_excel_history_compare_smoke.py` created a disposable `.xlsx` with static inline-M definitions, captured a local pre-change snapshot, changed one formula and added one query through desktop Excel, then used the production `ExcelQueryInstaller` to change a second query from definition-only to connection-only after a successful Office capability probe. It captured the second snapshot and ran `SnapshotCompareWorker` against both.
- Observed: comparison classified `Changed` as `MODIFIED` with the expected formula diff, `Stable` as `UNCHANGED`, `Added` as `ADDED`, and `LoadSwitch` as `MODIFIED` with `definition_only` → `connection_only` in the reason and read-back inventories. Both snapshot files retained their original SHA-256; no query refreshed; all six smoke-owned Excel instances (fixture creation, update, capability probe, installer, and two read-only inspections) exited.
- Limits: verifies query-presence, formula, and definition/connection-only comparison on one Office build. It does not verify worksheet/Data Model load-destination transitions, duplicate/unknown identities, `.xlsm`, snapshot corruption/concurrency races, external connections, or other Office builds.

## Excel read-only open without an owner-lock file — 2026-10-22

- Host/runtime: Windows 10 build 26300; Microsoft Excel 16.0 build 20430.0; Python 3.11.9 x64; pywin32 310.
- Observation: a disposable workbook opened read-only in desktop Excel had no `~$` owner-lock file and Restart Manager returned no active file user. Excel registered the exact workbook path in the COM Running Object Table. The combined active-use check found `Microsoft Excel (open workbook)`, and `validate_workbook_for_write` rejected before its write-open probe; the workbook SHA-256 remained unchanged and the owned Excel process exited.
- Implementation: `active_file_users` now combines Restart Manager results with exact-path ROT monikers. It inspects moniker display names only and does not bind to or invoke Excel objects. It checks the caller thread's COM state before initialization so nested use does not uninitialize a COM apartment owned by the caller. Errors fail closed.
- Limits: verifies this read-only Excel case on one local Office build. Protected View, remote sessions, OneDrive, other Excel instances/builds and Excel edit-mode open without owner file remain outside this smoke.

## Installer and protected-open recheck after ROT preflight — 2026-10-22

- Host/runtime: Windows 10 build 26300; Excel 16.0 build 20430.0; Python 3.11.9 x64; PySide6 6.8.3; pywin32 310.
- `tools/run_excel_import_smoke.py` passed current-source capability probe, staged connection-only create and repeat-replace, exact read-only formula/load verification, and owned-process cleanup. This confirms the new path-moniker scan does not reject an ordinary closed target.
- `tools/run_excel_password_safety_smoke.py` created a disposable open-password workbook. Both read-only inspection and staged import opened real Excel password prompts, timed out at 10 seconds, confirmed exit of their own Excel processes, and left the workbook hash unchanged.
- Limits: does not certify packaged GUI import, non-modal COM hangs, M evaluation timeouts, or other Office builds.

## All Excel load modes after ROT preflight — 2026-10-22

- Host/runtime: Windows 10 build 26300; Microsoft Excel 16.0 build 20430.0; Python 3.11.9 x64; PySide6 6.8.3; pywin32 310.
- Procedure: `tools/run_excel_load_smoke.py` capability-probed current Excel, then used the production staged installer for `new_sheet`, `existing_sheet`, and `data_model` on separate disposable `.xlsx` files. For each it created, repeated the same import, changed the M source and reimported, inspected the result read-only, then reopened output data to verify values and unrelated-cell preservation.
- Observed: all capability modes reported supported. `new_sheet` reopened as `worksheet: PQ_PQM_New_Load!$A$1:$A$3`; `existing_sheet` as `worksheet: Output!$B$3:$B$5`; Data Model as `data_model`. Each changed formula and output values `3, 4` verified; `PQM_Marker!A1` remained unchanged; all owned Excel processes exited.
- Limits: static inline M and one Office build. No external sources, `.xlsm`, timeout induced during M evaluation, varied workbook-object preservation, or other Office builds.


## Macro-enabled workbook container import — 2026-10-22

- Host/runtime: Windows 10 build 26300; Microsoft Excel 16.0 build 20430.0; Python 3.11.9 x64; PySide6 6.8.3; pywin32 310.
- Procedure: `python tools/run_excel_import_smoke.py --xlsm` created a disposable `.xlsm` with Excel `FileFormat=52` and no VBA project, then used the production installer to create and repeat-replace a static inline-M connection-only query. Excel read the formula and load state back in read-only mode.
- Observed: create/repeat-replace, exact formula and connection-only verification, and cleanup of every smoke-owned Excel process passed.
- Limits: verifies a macro-enabled file container without embedded VBA. It does not prove preservation or safe handling of a VBA project, XLM macros, Protected View, external sources, GUI import, or other Office builds.

## `.xlsx` default-path regression after `.xlsm` option — 2026-10-22

- Host/runtime: Windows 10 build 26300; Microsoft Excel 16.0 build 20430.0; Python 3.11.9 x64; PySide6 6.8.3; pywin32 310.
- Procedure and result: ran `python tools/run_excel_import_smoke.py` without the new flag. The default `.xlsx` staged connection-only create/repeat-replace, exact formula/load readback, and smoke-owned Excel cleanup all passed.
- Limits: static inline M, one Excel build, no external source or query refresh.


## Packaged GUI definition-only import — 2026-10-22

- Host/runtime: Windows 10 build 26300; Microsoft Excel 16.0 build 20430.0; Python 3.11.9 x64; PySide6 6.8.3; pywin32 310; PyInstaller onedir package.
- Procedure: `python tools/run_packaged_gui_import_smoke.py` created a disposable `.xlsx` with desktop Excel, selected a static inline-M `.m` source in the packaged UI, explicitly confirmed definition-only import, then reopened the workbook read-only through Excel. The smoke uses `tools/requirements_ui_smoke.txt` for optional UI automation dependencies and must run in an interactive Windows desktop.
- Observed: UI reported import complete; exact query name and formula were read back; the workbook hash changed; a timestamped backup matched its pre-import hash; unrelated cell `A1` retained its marker; the packaged GUI exited 0; and smoke-owned Excel processes exited. The pre-existing Excel instance remained untouched. DPI-scaled `click_input` coordinates initially missed the review control; the repeatable helper now routes clicks by the accessible screen rectangle and client coordinates.
- Limits: one Office build, `.xlsx`, and definition-only mode. No external query was refreshed. This does not verify standard-user ACLs, `.xlsm` with a VBA project, evaluated loads, external connectors, or other Office builds.

## Rebuilt packaged GUI import and timeout cleanup — 2026-10-22 (W35)

- Host/runtime: Windows 10 build 26300; Excel 16.0 build 20430.0; Python 3.11.9 x64; PyInstaller 6.13.0; PySide6 6.8.3; pywin32 310.
- `app/build.ps1` rebuilt the current source after the lifecycle timeout wait regression was fixed; its embedded portable suite reported 255 passed. The W34 visible-GUI disposable `.xlsx` smoke was repeated against `app/dist/PowerQueryManager/PowerQueryManager.exe` and passed exact formula verification, timestamped backup hash verification, unrelated `A1` preservation, GUI exit 0 and cleanup of owned Excel PIDs. The existing user-open packaged manager under root `dist/` was left untouched.
- `tools/run_excel_timeout_cleanup_smoke.py` also passed on the same Excel build; the helper waited for and confirmed termination of its exact `DispatchEx` PID after a controlled delay. This is not a blocked COM or M evaluation test.
- Limits: the current rebuilt app/dist GUI import remains definition-only, `.xlsx`, one Office build. Standard-user permissions, VBA-bearing `.xlsm`, external data, evaluated load modes, blocked COM hangs, and other Office builds are unverified. The new executable has not replaced the user-open root `dist/` package.

## Shared Excel application adapter build — 2026-10-22 (W36)

- Host/runtime: Windows 10 build 26300; Excel 16.0 build 20430.0; Python 3.11.9 x64; PyInstaller 6.13.0; PySide6 6.8.3; pywin32 310.
- Added shared `ExcelApplicationAdapter` injection to inspector, installer, refresher and capability probe. `DispatchExExcelAdapter` is the production implementation; fake adapters explicitly declare whether the service owns COM initialization; existing `excel_factory` injection remains compatible. Full suite during build: **258 passed**; focused adapter/service suite: **50 passed**; contract validation and compileall passed.
- New onedir GUI passed the W35 disposable definition-only import/reopen smoke against Excel 16.0 build 20430.0; new packaged CLI `--help` exited 0. The root `dist/` process stayed open and was not overwritten.
- Limits: this is an application-creation adapter, not yet the full transaction-level Excel protocol; service methods still perform their own workbook operations. This smoke does not exercise load evaluation or a blocked COM request. Higher-level extraction, non-modal COM timeout behavior, external sources, VBA-bearing `.xlsm`, standard-user rights and other Office builds remain unverified.

## Staged worksheet and Data Model load regression after bounded transaction calls — 2026-10-22 (W45)

- Host/runtime: Windows 10 build 26300; Microsoft Excel 16.0 build 20430.0; Python 3.11.9 x64; PySide6 6.8.3; pywin32 310.
- Procedure: `tools/run_excel_load_smoke.py` invoked the current-source capability probe, then the production `ExcelQueryInstaller` for `new_sheet`, `existing_sheet`, and `data_model` against separate generated `.xlsx` files. Each source was the same static inline M expression `let Source = #table({"Value"}, {{1}, {2}}) in Source`; no external connector or user source was evaluated. Each mode was created, repeat-imported, and reimported after changing expected static results to 3 and 4. Read-only Excel reopen checked formulas, load mapping, output values, and preservation of a marker cell.
- Observed: the probe enabled `connection_only`, `new_sheet`, `existing_sheet`, `data_model`, `definition_only`, and `preserve_existing`. Worksheet locations reopened as `PQ_PQM_New_Load!$A$1:$A$3` and `Output!$B$3:$B$5`; the model connection was present in `Workbook.Model.ModelTables`; all three modes returned values 3 and 4 after the changed-formula import; marker cell remained unchanged; tracked Excel processes exited.
- Limits: one Office build, static inline M, and the listed workbook shapes. This does not validate external data, VBA-bearing `.xlsm`, broader workbook object preservation, interrupted M evaluation, Protected View, OneDrive, or other Excel builds. Capability claims remain scoped to the probe observations.

## Capability-probe timeout boundary and current load-mode rerun — 2026-10-22 (W49)

- Host/runtime: Windows 10 build 26300; Microsoft Excel 16.0 build 20430.0; Python 3.11.9 x64; PySide6 6.8.3; pywin32 310.
- Method update: the production default `DispatchEx` capability probe now bounds potentially blocking static-load evaluation, workbook save/open/close, connection creation, and owned-process shutdown calls to 60 seconds each, targeting only the PID obtained from that probe's Excel window. An injected `ExcelOperationTimeout` regression proves that timeout aborts the remaining probe and disables all unverified modes. A known-exited owned process is not called again through stale COM proxies.
- Observed: `tools/run_excel_load_smoke.py` passed the updated capability probe and current staged importer for new worksheet, existing worksheet at `Output!B3`, and Data Model, with repeated and changed-formula import, read-only formula/load reopen, static output values, unrelated-cell preservation, and owned-process cleanup. The probe enabled definition-only, connection-only, new-sheet, existing-sheet, Data Model, and preserve-existing on this build. Packaged GUI definition-only import also passed after rebuild on the same Excel build.
- Limits: the live run did not induce a 60-second timeout; timeout behavior is verified by injected unit coverage plus the separate controlled owned-process timeout helper smoke. Not every COM property enumeration in capability readback has an individual watchdog. Arbitrary blocked COM, external connectors, macro-bearing `.xlsm`, other Office builds, and real evaluation timeout remain unverified.


## Read-only snapshot comparison of worksheet and Data Model loads — 2026-10-22 (W57)

- Host/runtime: Windows 10 build 26300; Microsoft Excel 16.0 build 20430.0; Python 3.11.9 x64; pywin32 310.
- Method: `python tools/run_excel_history_compare_smoke.py` created a definition-only and connection-only snapshot pair from one disposable `.xlsx`, then created equivalent disposable workbooks with a query loaded to a worksheet table and to the Data Model. `SnapshotCompareWorker` inspected hash-verified snapshots read-only through owned Excel instances; comparison performed no refresh. A portable regression covers the concrete inspector states `worksheet: Output!$A$1:$A$3` and `data_model`.
- Observed: same-workbook definition-only → connection-only showed `MODIFIED`; comparison across equivalent loaded workbook snapshots showed `worksheet: HistoryOutput!$A$1:$A$2` → `data_model` as `MODIFIED`. Formula diff, Added/Unchanged outcomes, unchanged snapshot hashes, no refresh, and exit of all 12 smoke-owned Excel processes passed.
- Limits: worksheet and Data Model snapshots came from separate equivalent disposable workbooks because the importer explicitly refuses implicit load-mode conversion on an existing query. This verifies comparison of known observed states, not same-workbook conversion, pivot/report relationships, `.xlsm`, ambiguous inventories, corruption, external queries, or other Office builds.


## W60 workbook privacy permission preservation — 2026-10-22

- Windows 10 build 26300; Microsoft Excel 16.0 build 20430.0; Python 3.11.9 x64; PySide6 6.8.3; pywin32 310.
- Implemented DataMashup v0 permission metadata inspection and validation of the SHA-256 permission binding using Windows DPAPI CurrentUser with the documented entropy. Staged import compares FirewallEnabled and WorkbookGroupType after save/close and aborts before commit on mismatch or unverifiable binding. A fake-COM regression confirms failure leaves original workbook bytes unchanged.
- Disposable Excel preservation smoke passed formula replacement and reopened verification while preserving FirewallEnabled=True and WorkbookGroupType=nil. Related worksheet/report preservation checks and owned Excel cleanup passed; pre-existing Excel PID 13156 was untouched. Portable tests: 299 passed; build-embedded suite: 299 passed.
- Limits: one Office build and default permission values only. A handcrafted non-default DataMashup fixture with a locally valid DPAPI binding was rejected by Excel during M mutation (COM HRESULT 0x800A03EC), so non-default settings and cross-user DPAPI bindings remain unverified. No external sources were evaluated. v1.5.0 remains partial; T150-03 and release gate are not green.


## W61 disposable macro-enabled container import — 2026-10-22

- Host: Windows 10 build 26300; Excel 16.0 build 20430.0; Python 3.11.9 x64; pywin32 310.
- Method: `PYTHONPATH=app\src python -B tools/run_excel_import_smoke.py --xlsm` created a fresh disposable `.xlsm` with Excel, applied and repeat-replaced a static inline-M connection-only query through the staged importer, and reopened it read-only for exact formula/load-state verification. No external source was contacted.
- Observed: staged create and replace, read-only reopen verification, connection-only state, and cleanup of smoke-owned Excel processes passed. Existing Excel PID 13156 was not targeted.
- Limits: this `.xlsm` contained no VBA project. The test verifies the macro-enabled workbook container path only; preservation of an actual harmless VBA project, macro non-execution, and automation-security restoration with a macro-bearing workbook remain unverified. This does not close v1.3.1 or v1.5.0 gates.


## W62 packaged GUI `.xlsm` import — 2026-10-22

- Host: Windows 10 build 26300; Excel 16.0 build 20430.0; packaged one-folder GUI; `PCRN\Robert` non-admin token.
- Extended `tools/run_packaged_gui_import_smoke.py` with `--xlsm`, which creates a fresh macro-enabled container without a VBA project. Packaged GUI import verified exact definition-only formula after read-only reopen, pre-import backup hash, manager exit 0, and Excel process-set cleanup. The `.xlsx` path also passed after the harness change. Existing Excel PID 13156 remained untouched.
- This tests packaged GUI handling of `.xlsm` container format. It does not verify actual VBA project preservation, macro non-execution, macro-bearing workbook automation-security restoration, or restricted ACL/Group Policy behavior. v1.3.1 remains UNVERIFIED.


## W63 file-use and password-protected workbook refusal — 2026-10-22

- Windows 10 build 26300; Microsoft Excel 16.0 build 20430.0; Python 3.11.9 x64. Disposable workbooks only; existing Excel PID 13156 was not touched.
- `tools/run_excel_open_without_lock_smoke.py` opened a disposable workbook read-only through owned Excel with no `~$` file. Restart Manager/ROT file-use detection found the active Excel workbook, preflight refused mutation, workbook hash remained unchanged, and the owned Excel PID exited.
- `tools/run_excel_password_safety_smoke.py` applied a test-only 10-second open bound to a disposable password-protected workbook. Read-only inspection and staged import both timed out, each owned Excel process exit was confirmed, workbook hash remained byte-identical, and no stage remained. The production timeout remains 60 seconds.
- These tests do not cover Protected View, OneDrive conflict, network-share ACLs, blocked calls outside bounded open operations, or arbitrary Excel builds. v1.3.1 remains UNVERIFIED.


## W64 packaged GUI VBA-bearing workbook preservation and non-execution — 2026-10-22

- Windows 10 build 26300; Microsoft Excel 16.0 build 20430.0; packaged one-folder GUI; current process token was non-admin. The pre-existing Excel PID 13156 was untouched.
- `tools/run_packaged_gui_import_smoke.py --vba` used existing `AccessVBOM=1` (read only; no Trust Center or registry setting was changed) to create a disposable `.xlsm` containing a harmless `ThisWorkbook.Workbook_Open` event that changes A1 to `MACRO_RAN`. The packaged GUI imported one static definition-only query. Post-import checks verified the VBA project remains embedded, event source remains intact, A1 remains `Power Query Manager GUI smoke marker`, the exact M formula reopens, backup matches original SHA-256, manager exits 0, and Excel process set returns to baseline.
- This directly verifies preservation/non-execution for this event macro and tested Office build. It does not certify all VBA projects, add-ins, XLM macro prompts, Protected View, alternate Office builds, or every entry point. v1.3.1 remains UNVERIFIED pending the remaining matrix.


## W65 packaged GUI `.xlsx` control rerun — 2026-10-22

- After adding the opt-in VBA fixture path to `tools/run_packaged_gui_import_smoke.py`, the existing standard `.xlsx` packaged GUI smoke passed again on Windows 10 build 26300 / Excel 16.0 build 20430.0. Exact formula, verified backup, read-only reopen, manager exit 0 and Excel process cleanup passed.
- This confirms the test-harness extension did not break the standard workbook route. Existing Excel PID 13156 remained untouched.


## W67 VBA-bearing workbook refresh safety — 2026-10-22

- Windows 10 build 26300; Excel 16.0 build 20430.0; Python 3.11.9 x64; pywin32 310. Existing Excel PID 13156 was not touched.
- `tools/run_excel_vba_refresh_smoke.py` created two disposable `.xlsm` workbooks with a connection-only static inline-M query and a harmless `ThisWorkbook.Workbook_Open` event that would change a marker cell if run. It exercised both synchronous and background staged refresh through the production `ExcelQueryRefresher`.
- Both modes refreshed successfully; original-hash backups matched; formula and connection-only load state reopened; `xl/vbaProject.bin` remained present; event source remained identical; marker remained `VBA_NOT_RUN`; all smoke-owned Excel processes exited. No external source was used.
- Limits: one Office build, one event macro form, static M, and the currently enabled VBA-project access setting; no policy settings were changed. Other VBA projects, XLM, external connectors, other Office builds, and interruption during refresh remain unverified.


## W68 packaged history-dialog UI Automation names — 2026-10-22

- Windows 10 build 26300; packaged one-folder GUI built from the current source; pywinauto UIA client.
- Added accessible names for the history privacy note, history table, retention period, optional snapshot setting, and history actions. `tools/run_packaged_history_a11y_smoke.py` opened the packaged UI, inspected the Local operation history dialog over Windows UI Automation, and verified the expected control names and roles. Manager exit code was 0. `LOCALAPPDATA` was redirected to a temporary QA directory and confirmed isolated.
- This confirms UI Automation exposes the named controls for this dialog; it is not a screen-reader user study or OS scaling acceptance. The full app suite passed **300 tests** before build; the onedir build passed its embedded **300-test** run. Packaged `.xlsx` and VBA-bearing `.xlsm` GUI import smokes and release package verification passed after rebuilding.


## W70 chronological snapshot comparison — 2026-10-22

- A new UI integration regression exposed that snapshot rows are displayed newest-first but were passed to the comparison worker in visual order, reversing “before” and “after.” The history dialog now orders the two selected snapshots by their recorded UTC timestamp before comparison.
- Regression passed with a fake compare worker and temporary snapshot files; full portable suite: **302 passed in 12.66s**. Windows onedir build passed its embedded 302-test suite. The rebuilt package verified at **229 files / 135,694,494 bytes**; CLI help, packaged history-dialog UI Automation, and disposable `.xlsx` GUI import/readback passed on Excel 16.0 build 20430.0. Existing Excel PID 13156 was untouched.
- This verifies timestamp ordering in the UI path and packaged regression set. Real Excel history comparison across snapshots remains scoped to the W57 data/inventory smoke; broader ambiguity/corruption and multi-build validation remain open.


## W77 — controlled owned-process timeout cleanup — 2026-10-22

- Host: Windows 10 build 26300; Excel 16.0 build 20430.0; Python 3.11.9; pywin32 310.
- Method: `tools/run_excel_timeout_cleanup_smoke.py` creates an isolated Excel instance with `DispatchEx`, calls the shared `call_with_timeout` helper around a controlled two-second callback with a one-second deadline, and verifies the exact process handle signals exit after timeout cleanup. It opens no workbook.
- Result: PASS. Process inventory returned to only the pre-existing Excel PID 13156. The helper terminated only its own Excel process.
- Limits: this exercises timeout lifecycle with a controlled callback, not an actually blocked COM call, blocked property getter, Power Query M evaluation, or arbitrary Office build. Those remain unverified.


## W78 — packaged GUI worksheet load confirmation — 2026-10-22

- Host observed Excel 16.0 build 20430.0. The packaged GUI ran its existing disposable capability probe, then imported a profile targeting a new worksheet using a static inline `#table` M expression.
- PASS: read-only reopen reported the expected query formula and worksheet load state; the created `GuiOutput!tblGuiSmoke` contained 17 and 29. The backup matched the pre-import workbook and the manager exited cleanly.
- This confirms the tested new-worksheet path for this build and static M input. It does not certify the load API across Office versions or query behavior involving external sources.

## Connection-only load and packaged manager acceptance — 2026-10-22 (W104)

- Host/runtime: Windows 10 build 26300; Microsoft Excel 16.0 build 20430.0; Python 3.11.9 x64; PySide6 6.8.3; pywin32 310.
- Method: `tools/run_excel_load_smoke.py` used the production `ExcelQueryInstaller` on a disposable `.xlsx`, created a connection-only query from static inline M, repeated the import, and reimported after changing only the static inline table values. Read-only Excel reopen queried `Workbook.Connections.Item("Query - PQM_Connection_Load")`, inspected worksheet ListObjects and Data Model tables, and verified the unrelated marker. `tools/run_packaged_gui_import_smoke.py --connection-only` then probed capabilities in the packaged GUI, loaded a versioned profile, applied the query, and verified the result via read-only inspection and Excel COM.
- Observed: capability probe enabled `connection_only`; production importer and packaged GUI both produced the `connection_only` inspection state and a `Query - GuiSmoke` workbook connection, with no worksheet output table and no Data Model table. The packaged GUI verified exact formula, original-hash backup, manager exit code 0, and clean exit of all Excel processes created during the smoke. Existing Excel PID 13156 remained untouched.
- Limits: one Office build, disposable `.xlsx`, and static inline M only. No external refresh/evaluation, macro-bearing workbook, other Office build, or production data source was exercised. This evidence confirms a distinct connection-only state for this tested Excel build only.

## Safe provider query names and Unicode load reimport — 2026-10-22 (W105)

- Query names for connection/worksheet/Data Model loads are constrained by `validate_query_sources` to a conservative provider-safe alphabet and length before Excel automation. A new regression tries semicolon connection-property syntax, a closing SQL identifier bracket, and quoted provider syntax; each raises before calling the Excel factory. This verifies current guards, not arbitrary OLE DB escaping.
- The production load smoke then used the permitted Unicode/spaces/dash/underscore name `PQM_Connéxion - Load_1` for connection-only create, repeat import and changed-formula reimport. Excel 16.0 build 20430.0 reopened it with the expected connection and no worksheet or Data Model load. The static M result, unrelated marker and owned-process cleanup were also verified.
- This is one Excel/Mashup provider build. Other builds and external-source M remain outside this observation.

## W107 non-query charts and shapes through staged query replacement — 2026-10-22

- Method: extended `tools/run_excel_preserve_smoke.py` to seed a disposable `.xlsx` with a chart sourced from an unrelated ordinary Excel table plus a separate named shape with alternative text. The production importer replaced the M definition through its staged transaction. A separate read-only Excel instance reopened the workbook and asserted chart name/title and shape marker alongside the existing table/style/filter, names, formulas, report table, connection-only load, and workbook privacy checks.
- Observed: PASS on Windows 10 build 26300 / Excel 16.0 build 20430.0; chart and shape survived, query/report formulas and values remained correct, and all exact smoke-owned Excel processes exited. Existing Excel PID 13156 was untouched.
- Scope: only these chart and shape forms on one build. Other drawing object variants, PivotTables/slicers, embedded objects, external links, macro projects, sync folders and Office builds remain unverified.

## W108 local PivotTable preservation — 2026-10-22

- Method: `tools/run_excel_preserve_smoke.py` seeded a disposable workbook with ordinary data table `tblKeep`, an Excel PivotCache/PivotTable on a separate worksheet, plus the previously tested chart, shape, query-backed report table, names and formulas. The production staged importer replaced static-M query definitions. A separate read-only Excel instance verified PivotTable structure and values after reopen.
- Observed: Excel 16.0 build 20430.0 preserved `PQM_KeepPivot`, the `Category` row field, `PQM Sum of Value` data field, category aggregates 7 and 54, and grand total 61. The rest of the preservation checks passed, privacy settings remained FirewallEnabled=True / WorkbookGroupType=nil, and smoke-owned Excel processes exited. Existing Excel PID 13156 was untouched.
- Limits: one local range-backed PivotTable on one Office build; external and Data Model PivotCaches, slicers/timelines, other Excel builds and unsampled object types remain unverified.


## W123 local PivotTable slicer preservation — Excel 16.0 build 20430.0

- Method: seed a disposable range-backed PivotTable; create a named slicer cache via `Workbook.SlicerCaches.Add` and a named slicer via `SlicerCache.Slicers.Add`; save a non-default item selection; run the production staged definition-only importer; reopen separately read-only. Microsoft documents [`SlicerCaches.Add2`](https://learn.microsoft.com/en-us/dotnet/api/microsoft.office.interop.excel.slicercaches.add2?view=excel-pia) and [`Slicers.Add`](https://learn.microsoft.com/en-us/office/vba/api/excel.slicers.add). The smoke uses `gencache.EnsureDispatch` on the already-owned slicer collection because late-bound pywin32 silently ignored optional middle arguments in this environment; it does not create or attach to another Excel application.
- Observed: local PivotTable and slicer survived; `Keep` remained selected, `Drop` remained excluded, and the filtered aggregate stayed 54. Query formula and owned-process cleanup passed.
- Limits: this certifies one local PivotTable/slicer configuration on build 20430.0. Timelines, external/Data Model PivotTables, alternate slicer styles/selections, other builds, and capability-probe behavior are not certified.

## 2026-10-22 — W125 injected transaction commit failure

- Portable fake-COM/import regression uses the production `WorkbookTransactionAdapter` boundary to raise at final commit after staged verification. The importer reports every staged query as rolled back, retains exact original workbook bytes, and cleans the stage.
- This validates truthful failure reporting and file-level isolation. It does not simulate a live Excel COM failure, filesystem power loss, or a failure after `os.replace` has actually completed. The live replacement smoke passes independently on Excel 16.0 build 20430.0; the injected commit-failure Excel matrix remains unverified.


## W129 — Excel 16.0 build 20430.0 — capability probe owned-process cleanup

- Method: capability probe uses a temporary workbook and `DispatchEx`; after releasing workbook proxies it requests `Quit`, waits on the exact process handle for the DispatchEx PID, terminates only that PID if it remains alive, and verifies exit again. Uncertain cleanup disables load capabilities. This avoids treating a `Quit` COM exception as proof that Excel is still alive or as permission to target another Excel instance.
- Observation: on Windows 10 build 26300 / Excel 16.0 build 20430.0, source probe enabled connection-only, new worksheet, existing worksheet, and Data Model. All four source load smokes passed. Packaged GUI Data Model import verified formula, InModel connection, matching ModelTables membership and row count 2, backup hash, manager exit 0, and owned process cleanup. Preexisting PID 36464 remained untouched.
- Limits: exact process wait/termination is unit-tested with fakes; live smoke exercised successful cleanup, not a live `Quit` exception or forced termination. One Office build and static inline M only; other builds, blocked COM calls and broader model relationships remain unverified.


## W130 — Excel 16.0 build 20430.0 — process-handle timeout cleanup

- Implementation: open one process handle with `PROCESS_TERMINATE | SYNCHRONIZE` for a bounded operation; use that handle for `TerminateProcess` and `WaitForSingleObject`, then close it. Capability probing retains the handle through its DispatchEx lifetime, including graceful Quit and fail-closed cleanup. Refresh passes the held handle through its timeout callback. This closes the PID-reuse window after handle acquisition.
- Verification: 32 focused lifetime/capability/refresh tests and the full 365-test suite passed; onedir build passed with 365 embedded tests. Packaged Data Model load/reopen/cleanup smoke passed on Excel 16.0 build 20430.0.
- Limits: no forced termination of a live blocked COM call was attempted; PID resolution to handle acquisition and other Office builds remain unverified.


## W131 — Excel 16.0 build 20430.0 — packaged GUI load modes after process-handle change

- Method: execute the rebuilt one-folder GUI import harness separately for connection-only, new worksheet, existing worksheet at `Output!C4`, and Data Model using fresh disposable `.xlsx` files and static inline M. Read-only Excel inspection after save checks formula, destination/load state, output values/model membership, original-hash backup, manager exit and smoke-owned process cleanup.
- Observation: all four modes passed; connection-only created a Workbook connection without worksheet/model output, worksheet loads returned values 17/29 at requested anchors, and Data Model membership reported two rows. Existing process PID 36464 was not targeted.
- Limits: one Excel build and static inline M. The method does not certify other builds, external source evaluation, or broad Data Model/report relationship preservation.


## W134 — Excel 16.0 build 20430.0 — HWND owner recheck

- Method: after opening a process handle for the PID resolved from an owned DispatchEx window, query `GetWindowThreadProcessId` for that same HWND again. Reject and close the handle if owner PID changed or cannot be read. A held kernel handle remains the target for timeout termination and process exit wait. Capability probe, regular bounded calls, and refresh supply the HWND.
- Observation: packaged Data Model capability probe/import passed on Windows 10 build 26300 / Excel 16.0 build 20430.0; formula, model table membership/row count 2, backup, manager exit and cleanup verified. A fake Win32 test verifies mismatched HWND owner refusal and handle closure.
- Limit: HWND-to-process-handle acquisition is not atomic in Win32; the test does not synthesize a simultaneous PID and HWND reuse race. Genuinely blocked COM termination remains unverified.


## W135 — Excel 16.0 build 20430.0 — packaged refresh with HWND ownership validation

- Method: packaged CLI refreshes a disposable static-inline-M connection-only query in background and synchronous modes. Check result status, pre-refresh backup SHA-256, post-refresh read-only formula/load state, and Excel process set. Production refresher obtains the owned Excel HWND and PID, opens a process handle, rechecks HWND ownership before bounded calls, and retains that handle through timeout cleanup.
- Observation: both modes returned `refreshed`; backup/readback checks passed; no smoke-owned Excel process remained. No external source was contacted.
- Limits: one Office build and static M only; no deliberately blocked COM call or synthesized HWND/PID reuse.


## 2026-10-23 — W146 — release package COM shutdown and refresh

- Host: Windows 10 build 26300; Excel 16.0 build 20430.0; Python 3.11.9; non-elevated user session.
- On the packaged GUI `.xlsm` import, the Excel application HWND and staged workbook window resolved to the same captured process PID. The prior cleanup waited while the COM apartment was still initialized, repeatedly timing out and force-terminating the app PID. The repair calls `Application.Quit`, drops the automation proxy, uninitializes COM on the same worker, and only then waits up to 15 seconds on the retained process handle. Fallback termination remains scoped to that exact PID.
- Observed PASS: packaged disposable VBA-bearing `.xlsm` import preserved exact M text, a SHA-256-matched backup, `xl/vbaProject.bin`, `Workbook_Open` source and the unchanged marker; read-only reopen and process-set check passed. The installer-owned Excel PID exited gracefully (`forced termination=False`); only pre-existing PID 36464 remained.
- Observed PASS: packaged CLI refresh returned `refreshed` for both background and synchronous modes with static inline M, verified backup hashes, no external source, and owned Excel cleanup.
- Limits: one Excel build and local disposable workbooks. This does not certify other Office builds, external data connectors, XLM, genuinely blocked COM termination, OneDrive interference, or complex workbook relationships. The VBA verification helper itself uses exact-PID cleanup for its read-only verifier process.


## 2026-10-23 — W147 — packaged snapshot comparison dialog

- `python -B tools/run_packaged_gui_import_smoke.py --exe dist/PowerQueryManager/PowerQueryManager.exe --history-dialog` passed on Excel 16.0 build 20430.0. The packaged history flow compared two seeded, local snapshots read-only; UI Automation observed the changes table, formula diff and close action. It performed no refresh and contacted no external source.
- The smoke's disposable GUI import and read-only verification also passed. Manager and owned Excel processes exited; only pre-existing Excel PID 36464 remained. This verifies this build and packaged dialog path, not alternate Office versions or screen-reader behavior.


## 2026-10-23 — W150 — Excel `Application.Wait` did not block through this COM call

- On Excel 16.0 build 20430.0, a hidden disposable `DispatchEx` application was asked to wait until a Python `datetime` 20 seconds in the future. The COM call returned `True` in about 0.031 seconds. This call shape therefore cannot serve as a deterministic blocked-COM timeout fixture on this host. No workbook was opened; the owned process exited afterward and only pre-existing PID 36464 remained.
- The attempted smoke-harness change was reverted. `run_excel_timeout_cleanup_smoke.py` still passes its controlled Python-side sleep timeout, which does not prove cleanup while a COM call is blocked. A separate deterministic COM blocking method and true blocked-call cleanup remain unverified.


## 2026-10-23 — W151 — terminate an owned Excel process during blocked COM refresh

- Command: `python -B tools/run_excel_blocked_com_timeout_smoke.py`. On Windows 10 build 26300 / Excel 16.0 build 20430.0, a hidden disposable `DispatchEx` instance made a synchronous `QueryTable.Refresh(False)` request to a local `127.0.0.1` CSV fixture. The server received the request and withheld the response. The 3-second timeout terminated only the captured Excel PID, and its retained process handle confirmed exit.
- Excel process set was `[36464]` both before and after; only the user's pre-existing PID 36464 remained. No workbook was opened or saved and no external network was used.
- The harness was repeated after tightening setup/cleanup and again passed with exact owned PID termination and baseline process-set preservation. This verifies the lifetime timeout mechanism during one genuinely blocked Excel COM call. It uses a legacy QueryTable fixture rather than a Power Query connection; Power Query refresh timeout semantics and other Office builds remain unverified.

## 2026-10-23 — W152 — Power Query refresh completion signals are not reliable on this build

- Host: Windows 10 build 26300; Excel 16.0 build 20430.0; Python 3.11.9.
- `python -B tools/run_excel_pq_blocked_refresh_timeout_smoke.py` failed acceptance: the packaged CLI exited 0 in about 0.094 seconds and reported `refreshed` / staged workbook committed, but the disposable `Web.Contents` source bound to `127.0.0.1` received no request. This is a false-success observation, not a successful refresh.
- A direct Excel probe configured `OLEDBConnection.BackgroundQuery=False`; `WorkbookConnection.Refresh()` returned `None` in 0.032 seconds. The delayed local request was observed in that run, but `OLEDBConnection.Refreshing` read false, `AfterCalculate` did not fire within the bounded observation window, and `OLEDBConnection.RefreshDate` raised a COM error before and after. Another probe did not reach the local endpoint. Each probe used a separately created Excel instance and disposable local fixture; process inventory returned to only pre-existing PID 36464.
- Conclusion: on this tested Office build, these signals cannot establish Power Query completion or support a trustworthy Power Query timeout. Do not treat the W146 static-M refresh smoke as proof of query execution; it proved only staged file/connection preservation for that fixture. Power Query refresh/timeout is **UNVERIFIED** and the packaged false-success must be fixed before v1.5.0 can pass its release gate. W151's legacy `QueryTable.Refresh(False)` blocked-COM result remains valid for that distinct API path.
- Microsoft documents [`OLEDBConnection.RefreshDate`](https://learn.microsoft.com/en-us/office/vba/api/excel.oledbconnection.refreshdate) and [`OLEDBConnection`](https://learn.microsoft.com/en-us/office/vba/api/excel.oledbconnection); the observed Power Query behavior above did not provide usable values.

## 2026-10-23 — W154/W155 — verified QueryTable refresh and timeout on Excel 16.0 build 20430.0

- Method spike: a temporary worksheet `QueryTable` bound through `OLEDB;Provider=Microsoft.Mashup.OleDb.1` to the existing workbook query executes `SELECT * FROM [<query>]`. With `BackgroundQuery=False`, `QueryTable.Refresh(False)` returned `True` only after the local CSV response, and the query result range contained the header and row. With `BackgroundQuery=True`, `Refresh(True)` returned with `Refreshing=True`; it later transitioned to false after the endpoint response, and the result range contained two rows. This is the state surface used by the refresher instead of `WorkbookConnection.Refresh`.
- Cleanup experiment: the temporary query table creates a separate workbook connection named `Connection`. Deleting the query table alone leaves that connection. Excel also refused deletion while the very-hidden scratch worksheet remained active/hidden. The verified cleanup order is delete table, unhide scratch worksheet, reactivate original worksheet, delete scratch worksheet, then delete only the generated connection; the workbook's original query connection remains.
- W154 packaged success: `python -B tools/run_excel_pq_refresh_smoke.py --cli-exe app/dist/PowerQueryManager/PowerQueryManagerCLI.exe` passed background and synchronous refreshes against a loopback-only CSV endpoint. Three local GETs were observed across the two operations; exact M formula, one original query connection, original worksheet count, pre-refresh backups, read-only reopen and process cleanup passed. Only pre-existing PID 36464 remained.
- W155 packaged timeout: `python -B tools/run_excel_pq_blocked_refresh_timeout_smoke.py --cli-exe app/dist/PowerQueryManager/PowerQueryManagerCLI.exe` passed a deliberately delayed loopback source with the 3-second timeout. CLI exit was 5, timeout was reported, original workbook SHA-256 and backup matched, stage was removed, and Excel process inventory returned to `[36464]`; only the captured app-owned Excel process was terminated.
- Scope at W154/W155: one Windows 10 build 26300 / Excel 16.0 build 20430.0 and connection-only query shape. W157 later adds worksheet-loaded refresh and W158 adds selected Data Model table refresh on the same build. Other Office builds, provider failures and external connectors remain unverified. This evidence does not certify a general M parser or external-source compatibility.

## 2026-10-23 — W157 — existing worksheet QueryTable refresh verified

- Command: `python -B tools/run_excel_pq_refresh_smoke.py` against the published `dist` CLI on Windows 10 build 26300 / Excel 16.0 build 20430.0.
- Extended packaged smoke seeded a worksheet-loaded query with an existing table at `D4`, changed its M formula to a local `Web.Contents` CSV source, and refreshed in background and synchronous modes. Both operations contacted the loopback source and read-only reopen found value `42` in the existing table. The table remained associated with `Query - PQM_Loopback_Success`; worksheet count, formula, connection inventory, and per-operation pre-refresh backup hashes were verified.
- The same run retested connection-only refresh in both modes. Nine local GETs were observed across four operations. All packaged CLI calls reported `refreshed`; no temporary tables, sheets, stages, or connections remained; Excel process inventory returned to `[36464]`. No external network was used.
- At W157, this verified connection-only and worksheet-loaded refresh on one build; Data Model refresh was still refused at that point. W158 below adds the scoped Data Model acceptance. Other Office builds, external connectors, provider failures, and the rest of the matrix remain open.

## 2026-10-23 — W158 — selected Data Model table refresh on Excel 16.0 build 20430.0

- Microsoft documents [`ModelTable.Refresh`](https://learn.microsoft.com/en-us/office/vba/api/excel.modeltable.refresh) as refreshing the model table's source connections; [`Model.Refresh`](https://learn.microsoft.com/en-us/office/vba/api/excel.model.refresh) refreshes every data source associated with the workbook model. The implementation uses the table-level method for a workbook with one model table. With multiple model tables, partial selections are refused; the workbook-wide method is used only when every table maps uniquely to a selected Power Query.
- `python -B tools/run_excel_data_model_refresh_spike.py` passed against a disposable model-loaded workbook. After changing its M definition to a `127.0.0.1` CSV source, `ModelTable.Refresh()` returned `None`, the local endpoint received requests, and the model table row count changed from 1 to 2. This confirms execution and readback for that query/build; it does not establish general external-provider semantics.
- `python -B tools/run_excel_data_model_refresh_smoke.py` passed against the published `dist` package. On a two-model-table disposable workbook, `--all` refreshed both sources in background-requested and synchronous-requested modes; each table reopened read-only with two rows, formulas and `InModel` associations remained exact, worksheet count was unchanged, each backup matched its pre-refresh hash, stages were removed, and process inventory returned to `[36464]`. Selecting only one of the two model queries was refused before any new source request or workbook-byte change.
- Excel exposes both model refresh APIs as synchronous methods. The manager report states this limitation; its worker runs separately from the UI. Other Office builds, external connectors, credentials/privacy prompts, complex model relationships and the broader Windows matrix remain unverified.

## 2026-10-23 — W187 — read-only Excel open without an owner-lock file

- Host: Windows 10 build 26300; Excel 16.0 build 20430.0; Python 3.11.9 x64.
- Command: `python -B tools/run_excel_open_without_lock_smoke.py`.
- Excel opened a valid disposable `.xlsx` with `ReadOnly=True` and no `~$active.xlsx`. `active_file_users()` returned `('Microsoft Excel (open workbook)',)` through the exact-path ROT check. `validate_workbook_for_write()` refused the target; before/after workbook hashes matched. The harness closed its own Excel process; only pre-existing PID 36464 remained.
- The paired `test_windows_restart_manager_detects_an_open_file_without_excel_owner_lock` focused regression passed against a separate Python subprocess holding a disposable `.xlsx` open without an owner-lock file; Restart Manager reported an active user and write preflight refused. These checks verify detection only; they do not attempt import/refresh/restore while another owner is active, and do not cover OneDrive.

## 2026-10-23 — W188 — write-service refusal with Excel ROT owner and no lock file

- Host: Windows 10 build 26300; Excel 16.0 build 20430.0; Python 3.11.9 x64.
- Command: `python -B tools/run_excel_open_without_lock_smoke.py`.
- A disposable `.xlsx` was opened read-only in Excel without a `~$` file. Exact-path ROT detection found `Microsoft Excel (open workbook)`. Import, refresh and history overwrite-restore each refused before a stage, backup, or Excel automation factory was reached. The workbook hash was unchanged. Harness-owned Excel exited; pre-existing PID 36464 remained untouched.
- One Office build only; OneDrive and broader matrix cases remain unverified.

## 2026-10-23 — W197 — query formula terminal-newline canonicalization

- Host: Windows 10 build 26300; Excel 16.0 build 20430.0; Python 3.11.9 x64.
- Method: created a disposable `.xlsx` with a hidden `DispatchEx` instance, added a static query using `Queries.Add` with an internal CRLF and terminal CRLF, saved, closed, reopened read-only, then read `Query.Formula`. Excel preserved the internal CRLF exactly but omitted the terminal CRLF. The owned Excel instance exited; pre-existing PID 36464 was untouched.
- Result: strict formula equality incorrectly rejects a valid source with a terminal line ending. Staged and independent readback verification now allow only trailing CR/LF removal; changes to internal line endings or any other M content still fail. A rebuilt packaged GUI import of UTF-8-BOM/CRLF `.m` into a disposable `.xlsx` passed formula readback under this rule, unchanged source SHA-256, verified pre-import backup, read-only reopen, manager exit 0 and owned-process cleanup.
- Scope: one Excel build and static definition-only query. Other Office builds and non-static query/load cases remain unverified.

## 2026-10-23 — W235 — rebuilt package load and refresh reacceptance

- Host: Windows 10 build 26300, Excel 16.0 build 20430.0, Python 3.11.9, PyInstaller 6.13.0. Each GUI load test used a generated disposable workbook and a non-elevated packaged GUI. New worksheet and `Output!C4` sheet loads read back exact formulas, values 17/29, backups and unrelated marker (existing sheet); connection-only reopened with a distinct connection and no sheet/model table; Data Model reopened with membership and row count 2. Owned Excel processes exited; PID 36464 remained.
- The rebuilt packaged CLI passed connection-only and worksheet-loaded Power Query refresh in background and synchronous modes against a temporary `127.0.0.1` CSV endpoint. The worksheet value was 42; formulas, connections, worksheet count, four pre-refresh backup hashes, temporary-object removal and owned process cleanup passed. Seven loopback GETs were observed; no external network was used.
- The same package passed a 3-second delayed-loopback refresh timeout: CLI exit 5, timeout report, original and backup hashes preserved, stage removed, owned Excel cleanup passed. This is one Office build; external providers and alternate builds remain unverified.

## W245 — rebuilt packaged definition-only GUI import (2026-10-23)

- Host: Windows 10 build 26300; non-elevated interactive session; desktop Excel 16.0 build 20430.0.
- Method: `python -B tools/run_packaged_gui_import_smoke.py --exe dist/PowerQueryManager/PowerQueryManager.exe` created a disposable `.xlsx` plus a static inline-M `.m` source, launched the one-folder GUI, traversed named controls with Windows UI Automation, and confirmed definition-only import. The import used the staged workbook path and the application-owned Excel COM worker. The harness then opened the committed workbook read-only with Excel and checked formula/backup state.
- Observed: Excel formula matched the expected static M; a timestamped backup matched the pre-import SHA-256; read-only reopen passed; GUI exited 0. Seed Excel PID 34508 exited, app-owned PID 43900 closed without forced termination, and verification PID 35764 was terminated during its cleanup. Process set returned to baseline `[36464]`; the pre-existing Excel PID 36464 was not touched. The temporary workbook and source were deleted.
- Limits: definition-only mode does not evaluate M or contact a source. This is one Excel build and one `.xlsx` fixture; it does not certify other Office builds, `.xlsm`/VBA forms, worksheet/connection/Data Model loads, providers, refresh, or the remaining workbook/security/accessibility matrix.

## W246 — packaged query replacement (2026-10-23)

- Host: Windows 10 build 26300; non-elevated session; Excel 16.0 build 20430.0.
- Method: `python -B tools/run_packaged_gui_import_smoke.py --exe dist/PowerQueryManager/PowerQueryManager.exe --replace-existing` created a disposable `.xlsx` with an existing definition-only query, then used the visible packaged GUI and its worker-thread Excel automation to replace the formula. The harness opened the committed workbook read-only to verify results. No refresh was requested.
- Observed: `GuiSmoke` returned Replaced; read-only reopen returned the exact new M formula; the timestamped backup SHA-256 equaled the original workbook SHA-256; an unrelated worksheet marker remained. Manager exit 0. Seed PID 26924 and app-owned PID 43216 exited; verification PID 38424 was cleaned up. Process set returned to `[36464]`, preserving the pre-existing Excel process. Temporary fixture files were removed.
- Limits: one Office build, `.xlsx`, and a definition-only source query. This does not verify replacement of worksheet/Data Model loaded queries, other Office builds, external providers, refresh, or the remaining security/accessibility matrix.

## W248 — packaged load-mode acceptance after UI rebuild (2026-10-23)

- Host: Windows 10 build 26300, Excel 16.0 build 20430.0, non-elevated user. Package: current rebuilt one-folder GUI. Four separate invocations of `tools/run_packaged_gui_import_smoke.py` used `--new-sheet`, `--connection-only`, `--existing-sheet`, and `--data-model` respectively.
- Each smoke created a disposable `.xlsx` with static inline M. The visible GUI selected an explicit destination; worksheet/Data Model modes requested evaluation of the static table; connection-only did not refresh or create a worksheet/model table. Staged import made a timestamped backup, and a separate read-only Excel COM reopen verified formula/load state and expected output.
- Observed: new-worksheet output at `GuiOutput!A1:A3` and existing-worksheet output at `Output!C4` each returned 17/29; the unrelated marker remained. Connection-only created a workbook connection without worksheet or Data Model output. Data Model connection membership and row count 2 were present after reopen. Every formula and pre-import backup hash matched. Four manager exits were 0. For the runs respectively, seed/app/verification PIDs were 41224/39312/25196, 41520/3116/24160, 23684/47328/32736, and 14780/1768/47376; smoke-owned processes exited or were cleaned during verification. Each process set returned to `[36464]`; that pre-existing Excel process was not touched.
- Limits: one Excel build, macro-free `.xlsx`, static in-memory M only. This provides no external-provider or cross-version certification, and does not close screen-reader, physical-display, or remaining workbook/security acceptance.

## W249 — packaged CLI background and synchronous refresh (2026-10-23)

- Host: Windows 10 build 26300; desktop Excel 16.0 build 20430.0; rebuilt packaged CLI; non-elevated execution context.
- Method: `python -B tools/run_excel_cli_refresh_smoke.py --cli-exe dist/PowerQueryManager/PowerQueryManagerCLI.exe` created a disposable `.xlsx` with query `PQM_CLI_Refresh` defined as a static inline M `#table` and a `Microsoft.Mashup.OleDb.1` workbook connection. The CLI ran refresh once in background mode and once synchronously. After each, the harness compared the new timestamped backup SHA-256 with the exact immediately preceding workbook hash and performed read-only query inventory.
- Observed: both JSON results returned status `refreshed`; formula and `connection_only` state persisted; both exact backup hash comparisons passed. The report recorded no external source. Smoke-owned seed/refresh/verification Excel processes exited.
- Limits: the static inline query has no network source. This verifies one connection-only query on one Excel build; it does not certify external providers, loaded worksheet/Data Model refresh for this exact package, other Office builds, or timeouts.

## W250 — packaged CLI Power Query connection and worksheet refresh (2026-10-23)

- Host: Windows 10 build 26300; Excel 16.0 build 20430.0; rebuilt packaged CLI; non-elevated.
- Method: `python -B tools/run_excel_pq_refresh_smoke.py --cli-exe dist/PowerQueryManager/PowerQueryManagerCLI.exe` created disposable workbooks and a temporary HTTP server bound only to `127.0.0.1`. The M query used `Web.Contents` to read generated CSV. The packaged CLI refreshed connection-only and existing-worksheet-loaded queries in background and synchronous modes; query execution is observed through Excel's temporary/existing worksheet QueryTable.
- Observed: four reports returned `refreshed`; the loopback server saw seven requests; all four backup hashes equaled immediate pre-refresh hashes; formulas and connection identity were retained; worksheet value 42 read back; temporary connection/sheet objects were removed. Excel process set returned to `[36464]`; pre-existing PID 36464 was untouched. The harness reports `external_network_used=false`.
- Limits: temporary local CSV only, one Excel build. External providers, authentication/privacy prompts, Data Model refresh in this exact rebuild, other Office builds and remaining workbook/security/accessibility acceptance are not certified.

## W251 — packaged synchronous refresh timeout (2026-10-23)

- Host: Windows 10 build 26300; Excel 16.0 build 20430.0; current packaged CLI; non-elevated.
- Method: `python -B tools/run_excel_pq_blocked_refresh_timeout_smoke.py --cli-exe dist/PowerQueryManager/PowerQueryManagerCLI.exe` created a disposable `.xlsx` with `PQM_Loopback_Timeout`, whose M used `Web.Contents` to request a delayed CSV from a server bound only to `127.0.0.1`. It invoked packaged CLI `refresh --apply --synchronous --timeout 3 --confirm REFRESH` and released the test-owned delayed response during cleanup.
- Observed: the request arrived; CLI exit code 5 and structured timeout details were verified. Original workbook SHA-256 remained unchanged; exactly one timestamped backup matched that hash; no stage remained. Excel process set returned to `[36464]`; `external_network_used=false`.
- Limits: this verifies one blocking synchronous Power Query path and one Office build, not all COM stalls, asynchronous timeout cases, external providers, or other Office versions. Remaining release/security/accessibility matrix stays open.


## W261 — Excel-managed refresh after Qt Network removal

- Host/package: Windows 10 build 26300; desktop Excel 16.0 build 20430.0; current compact packaged CLI; non-elevated context.
- Method: `python -B tools/run_excel_pq_refresh_smoke.py --cli-exe dist/PowerQueryManager/PowerQueryManagerCLI.exe`. The harness created a disposable `.xlsx` and an HTTP CSV server bound to `127.0.0.1`; Power Query `Web.Contents` ran inside Excel. The packaged app has no Qt Network/TLS components.
- Observed: connection-only and worksheet refresh in background and synchronous modes all returned `refreshed`; ten loopback requests arrived; formulas persisted; four backups matched exact pre-refresh SHA-256 values; worksheet value 42 reopened; temporary connections and sheets were removed. Process set returned to `[36464]`, and `external_network_used=false`.
- Limits: loopback-only, one Office build. External providers, authentication/privacy prompts, Data Model refresh on this exact package, alternate Office versions and remaining workbook/security matrix acceptance remain **UNVERIFIED**.

## W262 — Data Model load on current compact GUI

- Host/package: Windows 10 build 26300; Excel 16.0 build 20430.0; current compact PyInstaller GUI; non-elevated token.
- Method: `python -B tools/run_packaged_gui_import_smoke.py --exe dist/PowerQueryManager/PowerQueryManager.exe --data-model`. Capability probe enabled Data Model. A disposable `.xlsx` and static inline M were used; no external source was refreshed.
- Observed: manager reported Data Model load; read-only reopen confirmed exact formula, model membership and row count 2. Timestamped backup matched original workbook SHA-256. Manager exit 0; owned process cleanup returned Excel list to `[36464]`.
- Limits: one Excel build/package. Data Model refresh on this exact package, alternate Office versions and remaining workbook/security matrix items are **UNVERIFIED**.

## W263 — packaged Data Model refresh spike

- Host/package: Windows 10 build 26300; Excel desktop 16.0 build 20430.0; current compact PyInstaller GUI/CLI; non-elevated.
- Method: `python -B tools/run_packaged_gui_import_smoke.py --exe dist/PowerQueryManager/PowerQueryManager.exe --data-model --refresh-data-model`; the harness imported a static inline-M query into a disposable workbook's Data Model, then invoked packaged CLI refresh with explicit `--confirm REFRESH`.
- Observed: CLI returned `refreshed` in 3.031 seconds. `ModelTable.Refresh()` completed synchronously; connection association and `RecordCount=2` were verified before/after. Staged workbook verification passed, exact original-hash backup matched, read-only reopen preserved formula and model membership, and process cleanup returned to `[36464]`. No external network was used.
- Limitation: the Data Model table API is synchronous; it does not expose the background refresh option used for worksheet/connection QueryTables. One Office build only. Other Excel builds/providers and remaining security/workbook matrix items remain **UNVERIFIED**.

## W264 — packaged replacement preserves an existing Data Model load

- Host/package: Windows 10 build 26300; Excel desktop 16.0 build 20430.0; compact packaged GUI; non-elevated.
- Method: `python -B tools/run_packaged_gui_import_smoke.py --exe dist/PowerQueryManager/PowerQueryManager.exe --replace-preserve-model-load`. The harness seeded a one-row static-M query loaded to the Data Model, then used the GUI to replace only its M definition with a two-row definition under preserve-existing behavior.
- Observed: replacement completed and read-only reopen verified the new formula, `Connections.Item('Query - GuiSmoke').InModel == True`, existing model-table association and unchanged `RecordCount == 1`. This confirms the changed definition was not implicitly evaluated/refreshed. Original-hash backup, unrelated marker, manager exit 0 and owned Excel cleanup passed; process set returned to `[36464]`.
- Limitation: one Excel build and a simple static query. Other Office versions and workbook/Data Model relationship variants remain **UNVERIFIED**.

## W265 — packaged replacement preserves an existing worksheet load

- Host/package: Windows 10 build 26300; Excel desktop 16.0 build 20430.0; compact packaged GUI; non-elevated.
- Method: `python -B tools/run_packaged_gui_import_smoke.py --exe dist/PowerQueryManager/PowerQueryManager.exe --replace-preserve-worksheet-load`. The harness seeded a one-row static-M query loaded as table `tblGuiSmoke` at `Output!C4`, then replaced only the M definition with a two-row expression under preserve-existing behavior.
- Observed: GUI replacement completed; Excel read-only reopen verified the new formula, same table identity and C4 anchor, and unchanged one-row value 5. No implicit refresh or relocation occurred. Original-hash backup, unrelated marker, manager exit 0 and cleanup passed; process set returned to `[36464]`.
- Limitation: one Excel build and a simple static query. Other Office versions and workbook/table relationship variants remain **UNVERIFIED**.

## W267 — packaged GUI comparison after Excel terminal-newline normalization

- Host/package: Windows 10 build 26300; Excel desktop 16.0 build 20430.0; W266 compact onedir GUI; non-elevated.
- Method: `python -B tools/run_packaged_gui_import_smoke.py --exe dist/PowerQueryManager/PowerQueryManager.exe --compare-terminal-newline`. The harness seeded a disposable workbook query with static M lacking a final line ending and supplied the same `.m` formula ending CRLF.
- Observed: the packaged UI inventory reported `UNCHANGED`. GUI replacement then completed. Read-only reopen verified formula equality under Excel's terminal-CR/LF behavior; backup matched the pre-import workbook hash; the source SHA-256 was unchanged; manager exited 0 and owned Excel cleanup returned the process set to `[36464]`. No external source was contacted.
- Limitation: one Office build and static inline M. Other Office builds/providers and wider workbook/security/accessibility cases remain **UNVERIFIED**.

## W279 — current packaged GUI replacement after authorization redaction update

- Host/package: Windows 10 build 26300; Excel desktop 16.0 build 20430.0; W278 compact onedir GUI; non-elevated.
- Method: `python -B tools/run_packaged_gui_import_smoke.py --exe dist/PowerQueryManager/PowerQueryManager.exe --replace-existing`. The harness created a disposable workbook, seeded an existing query, and used the visible GUI's replace-existing flow.
- Observed: exact replacement formula readback, verified pre-import backup, unrelated worksheet marker preservation, and manager exit 0. The manager-owned Excel process exited without force, and read-only verification cleanup returned the process set to baseline PID `[36464]`.
- Limitation: this smoke validates the GUI import path but does not exercise authorization redaction. It covers one Office build and one static query; wider release gates remain **UNVERIFIED**.

## W280 — current packaged GUI mixed `.m` / `.pq` import

- Host/package: Windows 10 build 26300; Excel desktop 16.0 build 20430.0; W278 compact onedir GUI; non-elevated.
- Method: `python -B tools/run_packaged_gui_import_smoke.py --exe dist/PowerQueryManager/PowerQueryManager.exe --mixed-import`. The harness loaded a versioned profile containing three numbered mixed-extension sources and drove the packaged GUI import path against a disposable workbook.
- Observed: reviewed order and Excel query collection order matched; formulas matched sources exactly; backup hash, read-only reopen, manager exit 0, and owned Excel process cleanup passed. Process set returned to existing PID `[36464]`; no refresh or external source was used.
- Limitation: this fixture verifies three numbered entries only. Prefix ties, manual reorder precedence, and other Office builds remain **UNVERIFIED**.

## W281 — current packaged GUI all-skip no-op

- Host/package: Windows 10 build 26300; Excel desktop 16.0 build 20430.0; W278 compact onedir GUI; non-elevated.
- Method: `python -B tools/run_packaged_gui_import_smoke.py --exe dist/PowerQueryManager/PowerQueryManager.exe --skip-existing`. The harness seeded a query, marked the disposable `.xlsx` read-only, loaded a skip profile and changed source formula, then used the visible GUI.
- Observed: query was skipped; UI reported no workbook changes. Original formula, SHA-256, timestamp and read-only attribute stayed unchanged; no backup or stage appeared. Manager exited 0 and process set returned to existing PID `[36464]`.
- Limitation: this acceptance covers `.xlsx` on one Office build. Other workbook formats and Office builds remain **UNVERIFIED**.

## W283 — current W282 package GUI replacement

- Host/package: Windows 10 build 26300; Excel desktop 16.0 build 20430.0; W282 compact onedir GUI; non-elevated.
- Method: `python -B tools/run_packaged_gui_import_smoke.py --exe dist/PowerQueryManager/PowerQueryManager.exe --replace-existing` against the current package.
- Observed: exact replacement formula, verified pre-import backup, unrelated worksheet marker, and manager exit 0. The app-owned Excel process exited normally; read-only verification cleanup returned the process set to baseline PID `[36464]`. No refresh or external source was used.
- Limitation: one Office build and definition-only replacement; the broader workbook, security and accessibility gates remain **UNVERIFIED**.

## W284 — current W282 package Data Model import

- Host/package: Windows 10 build 26300; Excel desktop 16.0 build 20430.0; W282 compact onedir GUI; non-elevated.
- Method: `python -B tools/run_packaged_gui_import_smoke.py --exe dist/PowerQueryManager/PowerQueryManager.exe --data-model`. The visible GUI imported a static query after the live capability probe.
- Observed: probe enabled `data_model`; read-only reopen verified exact formula, Data Model membership and row count 2. Backup hash matched, manager exited 0, and owned Excel cleanup returned process set to baseline PID `[36464]`.
- Limitation: simple inline M and one Excel build. Complex Data Model relationships, other versions and external providers remain **UNVERIFIED**.

## W285 — current W282 package Data Model refresh

- Host/package: Windows 10 build 26300; Excel desktop 16.0 build 20430.0; W282 compact GUI and CLI; non-elevated.
- Method: `python -B tools/run_packaged_gui_import_smoke.py --exe dist/PowerQueryManager/PowerQueryManager.exe --data-model --refresh-data-model`. The GUI capability-probed and imported a static inline-M query; the packaged CLI explicitly refreshed it with `--confirm REFRESH`.
- Observed: CLI reported `refreshed` in 3.047 seconds. The connection and two model rows verified; Excel's Data Model API completed synchronously. Staged read-only verification passed, the backup matched exact pre-refresh workbook bytes, commit changed the staged workbook, and no stage remained. Owned process set returned to `[36464]`; no external source was contacted.
- Limitation: simple static M and one Excel build. Other providers, builds, timeout modes and broader workbook/security cases remain **UNVERIFIED**.

## W286 — current W282 package existing-worksheet load

- Host/package: Windows 10 build 26300; Excel desktop 16.0 build 20430.0; W282 compact onedir GUI; non-elevated.
- Method: `python -B tools/run_packaged_gui_import_smoke.py --exe dist/PowerQueryManager/PowerQueryManager.exe --existing-sheet`. The capability probe ran and the GUI loaded the query to `Output!C4` on a disposable workbook.
- Observed: probe enabled `existing_sheet`; read-only reopen verified exact formula, table anchor `Output!C4`, and values 17/29. Unrelated marker and original-hash backup were verified; manager exited 0 and process set returned to `[36464]`.
- Limitation: one build and simple inline M. Other Office versions, provider behavior and the full workbook matrix remain **UNVERIFIED**.

## W287 — current W282 package connection-only load

- Host/package: Windows 10 build 26300; Excel desktop 16.0 build 20430.0; W282 compact onedir GUI; non-elevated.
- Method: `python -B tools/run_packaged_gui_import_smoke.py --exe dist/PowerQueryManager/PowerQueryManager.exe --connection-only`. The capability probe ran and the GUI imported a static query with connection-only destination.
- Observed: capability enabled; read-only reopen found exact formula and an explicit workbook connection. No worksheet table, Data Model membership, or refresh was created/requested. Backup hash matched; manager exited 0 and process set returned to `[36464]`.
- Limitation: one Office build and static inline M. Other Excel versions and providers remain **UNVERIFIED**.

## W288 — current W282 package new-worksheet load

- Host/package: Windows 10 build 26300; Excel desktop 16.0 build 20430.0; W282 compact onedir GUI; non-elevated.
- Method: `python -B tools/run_packaged_gui_import_smoke.py --exe dist/PowerQueryManager/PowerQueryManager.exe --new-sheet`. The capability probe ran and the GUI imported a static query to a new worksheet.
- Observed: probe enabled `new_sheet`; read-only reopen verified the exact formula, worksheet load state, `GuiOutput!A1` table anchor and values 17/29. Backup hash matched, manager exited 0, and owned Excel cleanup returned to `[36464]`.
- Limitation: one build and static inline M. Other Office versions, refresh providers and the full workbook matrix remain **UNVERIFIED**.

## W289 — current W282 package definition-only import and dialog UIA

- Host/package: Windows 10 build 26300; Excel desktop 16.0 build 20430.0; W282 compact onedir GUI; non-elevated.
- Method: `python -B tools/run_packaged_gui_import_smoke.py --exe dist/PowerQueryManager/PowerQueryManager.exe` (default definition-only path).
- Observed: exact formula, verified backup and read-only reopen passed; operation report stated no connection or worksheet load. Manager exited 0 and process set returned to `[36464]`. During the same smoke, Windows UI Automation found named controls in the query inspector, dependency graph and profile mapping dialogs; mapping edits were cancelled.
- Limitation: definition-only static M and one Office build. UIA named-control discovery does not replace keyboard/screen-reader or broader accessibility acceptance.

## W291 — current W282 package deployment bundle apply/repeat

- Host/package: Windows 10 build 26300; Excel desktop 16.0 build 20430.0; W282 packaged CLI; non-elevated.
- Method: `python -B tools/run_excel_deployment_bundle_smoke.py --cli-exe dist/PowerQueryManager/PowerQueryManagerCLI.exe`. The harness generated two disposable `.xlsx` workbooks, a profile and portable `.pqmdeploy` bundle, then ran packaged CLI deploy/apply and repeat.
- Observed: first batch statuses `applied/applied`; repeat statuses `no_changes/no_changes`. Exact formula/load readback and two pre-deployment backup hashes passed; disabled missing source was ignored. Owned Excel PIDs 4348, 25344 and 39896 exited; pre-existing PID 36464 remained.
- Limitation: two local disposable `.xlsx` files and one Office build. Mid-commit batch failures, sync conflicts, `.xlsm`, and other Office builds remain **UNVERIFIED**.

## W292 — current W282 package multi-workbook dry-run plan

- Host/package: Windows 10 build 26300; Excel desktop 16.0 build 20430.0; W282 packaged CLI; non-elevated.
- Method: Created two disposable `.xlsx` files with isolated `DispatchEx`, then ran packaged CLI `plan --dry-run --profile <temporary-profile> --workbook <target-one> --workbook <target-two> --output <new-report>` using one static definition-only M source. No `--apply`, confirmation phrase, or refresh was supplied.
- Observed: both plans reported `create`; original target SHA-256 values matched after the CLI run; no backup or stage existed; Excel process set remained at baseline PID `[36464]` after the CLI run. No M evaluation was requested.
- Limitation: two new, empty `.xlsx` targets and one build. Existing-query/no-op dry-run, `.xlsm`, other Office builds and broader failure cases remain **UNVERIFIED**.

## W293 — packaged CLI query-package import

- Observed environment: Windows 10 build 26300, desktop Excel 16.0 build 20430.0, non-elevated; package `dist/PowerQueryManager/PowerQueryManagerCLI.exe`.
- Method: `python -B tools/run_excel_cli_package_import_smoke.py` created two disposable `.xlsx` workbooks through isolated Excel automation, installed three static inline-M definitions into the source workbook, exported a `.pqm.zip`, and invoked the packaged CLI with explicit `--confirm IMPORT` against the target. No refresh or external-source evaluation was requested.
- Observation: CLI exited successfully; read-only inspection returned exact names and formulas, including `Équipe – Résultats 2026 (Q1)`. Source-workbook SHA-256 was unchanged by export. The target backup matched its exact pre-import hash. Owned Excel instances exited; only pre-existing PID 36464 remained.
- Limitation: one workbook shape and one Office build only. This does not certify package import across Office versions, workbook protection/macro variants, provider types, or remaining security matrix cases.
## W295 — current packaged GUI VBA preservation

- Observed environment: Windows 10 build 26300, Excel 16.0 build 20430.0, non-elevated, current compact GUI package.
- Command: `python -B tools/run_packaged_gui_import_smoke.py --exe dist/PowerQueryManager/PowerQueryManager.exe --xlsm --vba`.
- The harness created a disposable VBA-bearing `.xlsm` with a harmless `Workbook_Open` marker and used the packaged GUI for definition-only import. Read-only reopen verified exact M, preserved VBA source and unchanged marker. The timestamped backup verified; the manager exited 0 and owned Excel processes exited, leaving only pre-existing PID 36464.
- This does not certify other macro projects, Trust Center policies or Office builds.
## W297 — packaged CLI batch with a read-only target

- Observed environment: Windows 10 build 26300, Excel 16.0 build 20430.0, non-elevated, current compact CLI package.
- Command: `python -B tools/run_excel_deployment_mixed_smoke.py --cli-exe dist/PowerQueryManager/PowerQueryManagerCLI.exe`.
- The harness created two disposable `.xlsx` files and set the second target read-only. Packaged `deploy --apply --confirm DEPLOY` returned `applied` for the writable target and `blocked_preflight` for the read-only target, with exit code 5. Packaged `inspect` verified the formula on the successful target; its backup matched the original hash. The blocked target had no backup and retained its original hash. CLI-owned Excel processes exited; process set returned to `[36464]`.
- This confirms isolated partial deployment through the packaged CLI on one build. Other ACLs, mid-commit failures and Office builds remain unverified.
## W298 — workbook lifecycle adapter package smoke

- Host: Windows 10 build 26300; desktop Excel 16.0 build 20430.0; Python 3.11.9; PyInstaller 6.13.0; non-elevated.
- Change: introduced injectable `ExcelWorkbookOperations` for add/open/save/save-as/close, used by workbook inspection, staged import, staged refresh and capability probing. COM calls remain inside their existing bounded-call wrappers.
- Focused fake-COM/service tests: 91 passed. Rebuilt one-folder GUI/CLI; packaged CLI help and package manifest verification passed.
- Real Excel method: `python -B tools/run_packaged_gui_import_smoke.py --exe dist/PowerQueryManager/PowerQueryManager.exe` imported a definition-only static query into a disposable `.xlsx`. Read-only reopen verified exact M and original-hash backup; GUI exited 0 and owned Excel processes exited. Pre-existing PID 36464 was untouched.
- Limit: one definition-only GUI path/build validates lifecycle routing; other import/load/refresh lifecycle paths are covered by focused fake-COM tests and their earlier Excel evidence, not all rerun against this build. Other Office versions and broader matrix remain unverified.
## W299 — packaged CLI active-use refusal without Excel owner-lock file

- Environment: Windows 10 build 26300, non-elevated; desktop Excel 16.0 build 20430.0; W298 onedir CLI package.
- Method: a harness-owned Excel instance opened a disposable `.xlsx` read-only; the test confirmed `ReadOnly=True` and absence of `~$active.xlsx`, then invoked packaged CLI `deploy --apply --confirm DEPLOY` using a temporary versioned profile.
- Observation: CLI returned exit 5 with one `blocked_preflight` result explaining the workbook was open or in use. Target SHA-256 stayed unchanged; no backup/stage appeared; no additional Excel process started. The harness closed only its own workbook/application and waited for its PID to exit; the pre-existing Excel PID 36464 remained.
- Limit: one package and Office build; concurrent external Excel launch, network/sync paths, other lock-file providers and other Office builds remain **UNVERIFIED**.
