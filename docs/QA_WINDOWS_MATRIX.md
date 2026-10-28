# Windows Excel acceptance: manual + script-assisted test matrix

## W300 — final packaged header fix and end-to-end Windows checks, 2026-10-28

- Exact tested application commit: `3e45bc450482b5311a6a30938c27ea19fca50a45`. Windows 10 build 26300, Python 3.11.9 x64, PySide6 6.8.3, pywin32 310, PyInstaller 6.13.0, Excel 16.0 build 20430.0; acceptance harness confirmed non-elevated launch.
- Fixed the clipped `Enabled` source-table heading by reserving 132 px for that user-resizable checkbox column. Added a regression assertion for its resize mode and minimum initial width. Fresh packaged UI screenshot at 1182×732 shows the complete heading and a synthetic `StaticTable` source row.
- `app/build.ps1`: **542 passed in 44.40s**, including the added regression. Full UI suite at simulated Qt scale 1.25: **92 passed in 9.04s**; at 1.5: **92 passed in 8.93s**. These do not substitute for physical screen-resolution or OS-level DPI checks.
- Rebuilt/staged package verifier: **PASS, 116 files / 76,330,196 payload bytes**. Packaged CLI `--help` exited 0. GUI SHA-256 `CFD598065D357484D3C79935E2465587858FBAD4D7930A051F706A48EDD646AB`; CLI SHA-256 `6B30D61A87A2A164949C35EFB8BEEE2226B40344741CD64BA974365F752F0382`.
- Packaged Excel GUI create, all-skip, replacement, and mixed three-source `.m`/`.pq` import passed with exact formulas/order, pre-import backup verification where writes occurred, unchanged workbook hash/timestamp and no backup/stage on all-skip, read-only reopen, manager exit 0, and natural cleanup of manager-owned and verification Excel processes. Packaged history comparison, hash-verified history restore with exact pre-restore backup, and query-package comparison dialogs exposed their named controls and passed.
- Packaged DAO 12.0 Windows acceptance report: `%TEMP%\\pqm-vector-dependency-acceptance-3e45bc4-20261028.json`; exact source commit above, overall **PASS** on local synthetic fixtures (10 files: 4 Access, 4 Excel, plus source/support files). Source hashes/sizes/timestamps stayed unchanged; no lock files, Access temp directories or residual Excel processes remained. Quick and Deep each returned 12 edges across 31,200 formula cells, with no Quick-only or Deep-only edges; JSON/CSV exports verified all 12 edges, filter counts matched, and UI progress/cancellation were observed. Excel seed process exited naturally (2.656 s); no forced termination. DAO 12.0 evidence does not prove DAO 16.0.
- Profile save/load had already been verified in W299; the later mixed-source import additionally exercised profile-driven query order against real Excel. A profile load alone does not prove every profile option/mapping scenario.
- **Still open:** physical Windows layout at 960×600 and 1366×768, physical 125%/150% OS DPI, full keyboard-only navigation and remaining untested controls/workflows. CUA resize attempt left the live window at 1182×732. ACE DAO 16.0 and mapped SMB/UNC remain explicitly **WAIVED / NOT VERIFIED**, never PASS. Keep PR #2 open until all non-waived acceptance checks and review are complete.

## W299 — packaged Vector source/profile workflow, 2026-10-28

- Exact tested application source commit: `536ee6ce5d901eabbe7345c8c2ffc287cc31d372`; tested package: `app/vector-release/PowerQueryManager/PowerQueryManager.exe`, rebuilt through `app/build.ps1`. Package hashes (from W298): GUI SHA-256 `34AC00810E01D6E1E4C23ED667491E2059F74AC9074ECDAF970A08BF494665DC`; CLI SHA-256 `6B30D61A87A2A164949C35EFB8BEEE2226B40344741CD64BA974365F752F0382`.
- On Windows build 26300, packaged GUI added a synthetic `StaticTable.m` file and displayed its `StaticTable` row, source path, definition-only mode and enabled/import-order metrics. The packaged source inspector opened and displayed the exact source M text; workbook comparison remained unavailable/unknown because no query was inspected from the selected workbook.
- Saved and then loaded a JSON import profile through the packaged UI. The app reported `Loaded 1 source row(s). Review the paths and query names before applying.` The saved profile contains the relative synthetic workbook/source paths and the expected definition-only row. The synthetic `Dashboard.xlsx` SHA-256 remained `2E12C3926965C9B96F47A23533AADE6D53743114C44E23EED2662DAA91888149`; the `StaticTable.m` SHA-256 remained `3E8737D902F97C739104567B6FEDA9ECFED418B001A21DB18CA8B1751BFA6325`. The profile round-trip selected the workbook path but did not open or modify the workbook. Packaged GUI exited normally; no PowerQueryManager process remained.
- The live window remained 1182×732. An attempted resize to 960×600 did not change its bounds. 960×600 / 1366×768 Qt layout assertions and 125% / 150% Qt scale-factor tests are not physical Windows display/DPI verification. The Enabled heading looked clipped and was fixed in W300; a fresh packaged screenshot now shows the full label.
- Other interactive workflows (dependency scans/graphs in this session, history, filters, exports, progress/cancellation, and workbook import) were not exercised by this check. The complete build test suite and package verification results remain recorded in W298; this session adds no claim about these untested UI flows. ACE DAO 16.0 and mapped SMB/UNC remain **WAIVED / NOT VERIFIED**, never PASS. PR #2 remains open; do not merge until all non-waived acceptance gates and the review are complete.

## Vector candidate — 2026-10-28, rebuilt and visually checked against Concept 03

- Exact tested application commit: `e934432e2f8f86b421bf51d2b7869fad5326318c`. Windows build 26300, Python 3.11.9, PySide6 6.8.3, PyInstaller 6.13.0, Excel 16.0 build 20430.0; packaged GUI launched as the current user. No workbook was selected or modified during this visual pass. The QA documentation update is a separate follow-on commit and does not change application sources.
- `app/build.ps1`: **541 passed in 37.13s**; GUI and CLI rebuilt. `tools/package_release.py` staged the release and `tools/verify_release_package.py` passed. Packaged CLI `--help` exited 0. Release manifest lists 116 payload files / 76,330,010 bytes (manifest itself excluded). GUI SHA-256 `145F60264FD182B96787ED42F405C513FD3A7274155FF7A1C9FA336756D0F2B5`; CLI SHA-256 `6B30D61A87A2A164949C35EFB8BEEE2226B40344741CD64BA974365F752F0382`.
- Compared the actual main workspace to the supplied Vector Concept 03 board. Implemented responsive 4-to-2 metric cards, 4-to-2 workbook action layout, 3-to-2 source actions, and a left title/right primary action composition. Added layout regression assertions for 960×600 and 1366×768. The packaged main workspace and its synthetic query table were visually observed. Native Windows title bar/menu remain outside the styled canvas.
- Packaged query Dependency graph was inspected with five synthetic source queries and one detected `Report -> QueryHelper` dependency. A label stacking defect was fixed by placing title/status text above node cards, and a regression assertion checks text z-order. The Finder empty state was opened and its accessibility tree confirmed mode, Quick/Deep depth, file/target selection, scan/cancel, results, coverage, and export controls. No packaged Finder scan or populated file graph was executed during this pass.
- Full GUI workflow, inspector and other dialogs were not all manually exercised. DPI values 125%/150% are covered by Qt scale-factor tests in the UI suite; physical Windows DPI and on-screen checks at 960×600 / 1366×768 were not performed. Keyboard focus traversal, all graph interactions, progress/cancel and workbook import remain covered by existing tests and earlier harness evidence, but were not repeated in this visual pass.
- **Open gates:** complete interactive Windows workflow and visual inspection of remaining dialogs and populated file graph; physical DPI/resolution review. ACE DAO 16.0 and mapped SMB/UNC remain explicitly **WAIVED / NOT VERIFIED**, never PASS. Preserve those waivers. Do not merge until every non-waived acceptance check passes and review is complete.

## Vector feature candidate — Windows verification, 2026-10-27

- Tested application source commit: `d72c25440231def2c4bd91234523da72064d1456`; Windows build 26300, Python 3.11.9, PySide6 6.8.3, PyInstaller 6.13.0, Excel 16.0 build 20430.0. The later smoke-harness adjustments only change test driving, not the packaged application.
- `app/build.ps1`: **540 passed**, then both onedir executables built. The isolated staged release verified **116 files / 76,326,415 bytes**. GUI SHA-256 `08DE89817BB721D4B72EE8D5FF70DA2EF9B9D82D746D5157E7E6193F0CC25106`; CLI SHA-256 `6B30D61A87A2A164949C35EFB8BEEE2226B40344741CD64BA974365F752F0382`. Packaged CLI `--help` exited successfully.
- UI tests: **90 passed** with each of `QT_SCALE_FACTOR=1.25` and `1.5` using Qt offscreen. The laptop layout assertion checks 960×600 and 1366×768 without horizontal workspace scrolling. This is not OS-level DPI or physical-display verification. Packaged UI Automation traversed 43 controls by Tab, including workbook, table, backup and import; actual GUI launch, main workspace, command palette, empty inspector and Finder were visually observed. The original concept board was unavailable for side-by-side comparison.
- Local disposable Excel/Access scanner harness with `DAO.DBEngine.120`: **PASS** at the tested application commit. Ten fixtures retained hashes and timestamps; no Access locks or temp directories remained; Excel exited naturally; Quick/Deep, filter counts, 12 graph edges, CSV/JSON exports, progress and cancellation were checked. The in-process Finder checks do not prove every packaged GUI interaction.
- Packaged GUI definition-only create and existing-query replacement succeeded with exact formula readback and matching pre-import backups; the manager-owned Excel processes exited without forced termination. The separate read-only verification process required forced cleanup in the first runs. After the harness was corrected to release its COM reference and apartment before waiting, repeated create/history and replacement runs exited both manager-owned and verification Excel processes naturally. Inspector, query-graph, mapping and snapshot-comparison dialog accessible names were checked.
- **Open gates:** full live workflow and visual coverage (including populated graphs, all dialogs, DPI and resolutions), and review against the actual Vector board. ACE DAO 16.0 and mapped SMB/UNC remain **WAIVED / NOT VERIFIED**, not passed. Do not merge based on these partial results alone.

**Never use production workbooks.** Create disposable `.xlsx` and `.xlsm` with desktop Excel; include M sources in `examples/sources/`. Source packages must contain no company data.

## Test host capture

Record UTC timestamp, OS/build, Python version, PySide6 version, pywin32 version, PyInstaller version, Excel edition + `Application.Version`/`Application.Build`, 32/64-bit Office, workbook path category (local vs synchronized), non-admin launch confirmed, test file SHA-256 before/after, screenshots and process observations. Do **not** log usernames or internal paths in shared artifacts.

## Test fixtures (create using real Excel on test PC)

- F1: clean `.xlsx`, sheet `Existing`, marker text at `A1`, blank target range elsewhere.
- F2: `.xlsx` with existing Power Query named `QueryHelper`, sheet table bound to it, and a separate unrelated table `Unrelated`.
- F3: `.xlsm` with harmless test macro and independent worksheet data; verify code preserved but not run by programmatic open. Do not use malicious macros.
- F4: workbook with query loaded to Data Model, plus workbook connections where Office supports them.
- F5: locked/open/read-only or password-protected workbook; disposable copy only.
- F6: workbook with query names with spaces, non-ASCII, punctuation requiring safe quoting.
- F7: workbook in non-corporate OneDrive test sync folder (optional); record actual supported behavior; don't sync secrets.

## Execution checklist

### DF — reverse Dependency Finder (2026-10-27; NOT VERIFIED)

Use **only disposable** sample Excel and Access files and an isolated test
folder. Do not use work or personal database contents. This is a new source
feature; the checked-in packaged binary has not been rebuilt.

1. Create `Warehouse.accdb`, `Stock.mdb`, `Report.xlsx` and `Queries.xlsm`;
   use actual DAO linked tables / saved Access queries pointing to
   `Warehouse.accdb`, and Excel Power Query `Access.Database(File.Contents(...))`
   and a regular Excel external connection to the disposable target.
2. Record SHA-256, timestamp and any `.laccdb`/`.ldb` files before the scan.
   Use the GUI's **Dependency Finder…** to scan the folder recursively and
   choose multiple targets. Repeat with explicit multiple-file selection.
3. Confirm incoming rows correctly identify Excel query/connection names and
   Access linked-table/query names. Filename-only and unresolvable dynamic
   expressions must not appear as confidently exact dependencies.
4. Run `power-query-manager-cli dependency-scan --target PATH --folder DIR --recursive`
   on the same fixtures and compare its JSON results against the GUI export.
5. Confirm **every original input hash and timestamp is unchanged**, and no
   Access lock file was created beside any original. Check temp cleanup,
   no Excel process launched, no macros/queries executed and no data refreshed.
6. Repeat against a mapped network drive, a UNC alias, a cloud-synced test
   placeholder, an unavailable/incompatible DAO engine, an encrypted Access
   file, malformed `.xlsx`, `.xls`/`.xlsb`, cancel after progress, and a
   filename collision in a different folder. Verify coverage warnings.
7. Build the Windows onedir package, test both GUI and packaged CLI under a
   normal non-admin token, and record release/version/provider bitness.
8. Switch to **Map all selected files**. Check that an `.xlsx` consumer which
   is *also* a dependency target still appears as a source of outgoing links;
   use each file's Incoming/Outgoing/Both filters and confirm they update
   without another scan. Confirm the visual graph and complete JSON/CSV
   include all relationships, unresolved references and coverage statuses.
9. Repeat Quick versus Deep with a large disposable worksheet containing
   external-link formulas. Measure peak RSS and elapsed time, document the
   256 MiB metadata budget behavior, and verify unchanged original hashes.
   The streaming parser and temporary Access disk-copy behavior are not yet
   validated with very large, real corporate-sized files.

Acceptance: **UNVERIFIED** until actual Windows/Access/Excel host results and
unchanged-file hashes are recorded. Synthetic portable tests alone are not a
substitute.

### DF current Windows build check — 2026-10-27

- Tested source commit before this entry: `c5009db52d9c321f86d5e5bd9b11e7b73fa93aeb`.
- Windows 10 build 26300, Python 3.11.9, PySide6 6.8.3, pywin32 310,
  pytest 8.3.5, PyInstaller 6.13.0. `app/build.ps1` installed/confirmed the
  pinned Windows dependencies, ran the complete suite (**536 passed in
  37.79s**), and rebuilt the GUI and CLI.
- `tools/package_release.py` staged the rebuilt release; the package verifier
  passed (**116 files, 76,310,185 bytes**). The packaged CLI `--help` lists
  `dependency-scan` and `dependency-map`; the latter lists `quick` and `deep`.
  Both packaged EXEs have fresh hashes in `dist/SHA256SUMS`.
- The packaged GUI launched and its Windows accessibility tree exposed
  **Find files that depend on selected Excel or Access databases and
  workbooks** and **Dependency graph**. This establishes launch/control
  presence only; no scan, filter, graph, or export workflow was exercised.
- **NOT VERIFIED:** real `.accdb`/`.mdb` DAO metadata reads, linked tables and
  saved queries; genuine Excel Mashup/external workbook references; incoming,
  outgoing and both filters; GUI/CLI result parity; JSON/CSV correctness;
  large-workbook Quick/Deep timing and peak working set; local/mapped SMB/UNC/
  synchronized path execution; scan-level source hashes/timestamps, lock-file
  and temp cleanup. No `.accdb` or `.mdb` fixture was present in the checkout
  or the inspected test fixture locations. Therefore the required acceptance
  criteria remain **FAIL / OPEN** despite the passing suite and package build.
- The committed CLI and synthetic scanner tests exercise Quick/Deep coverage
  rules, but are not a substitute for large real workbook comparison. No
  failure in the full suite required a source-code repair. This is a build and
  portable suite result, not completed integration acceptance.


### W73 — 2026-10-22 — packaged GUI all-skip no-op

- Command: `python -B tools/run_packaged_gui_import_smoke.py --skip-existing` using the current one-folder package.
- Host: Windows 10 build 26300; Excel 16.0 build 20430.0. The smoke seeded a disposable definition-only query, marked the workbook read-only, loaded a versioned skip profile, and submitted a changed source formula with backup enabled in the profile.
- PASS: UI reported “No changes”; existing formula and worksheet marker remained unchanged; workbook SHA-256 and timestamp were byte-for-byte/metadata identical; read-only attribute remained set; no backup or stage was created; manager exit was 0 and all smoke-owned Excel processes exited. Existing Excel PID 13156 remained untouched.
- Scope: one Office build and all selected query names already existed. Mixed skip/create, workbook-open conflicts, and other builds remain unverified.

### W72 — 2026-10-22 — abrupt process exit around atomic workbook commit

- Two portable subprocess tests use a temporary workbook and production transaction adapter. Hard exit after complete stage write/before commit preserves original bytes and leaves a complete stage; hard exit immediately after `os.replace` leaves complete new bytes and no stage.
- Results: focused **2 passed**; full suite **306 passed in 11.68s**.
- Scope excludes Excel COM call interruption, power loss, and transaction phases outside the two injected boundaries.

### W71 — 2026-10-22 — repeat packaged VBA-bearing GUI import

- Host: Windows 10 build 26300; Excel 16.0 build 20430.0; packaged one-folder GUI; non-admin user context.
- Command: `python -B tools/run_packaged_gui_import_smoke.py --vba`.
- Result: PASS. Disposable `.xlsm` with harmless `Workbook_Open` imported a static definition-only query; exact M readback, verified original-hash backup, preserved VBA project/source, unchanged marker, manager exit 0, and smoke-owned Excel cleanup verified.
- Existing Excel PID 13156 was unchanged. This does not cover other Office builds, macro forms, external M sources, Protected View, or interrupted commit phases.

### W71 — 2026-10-22 — repeatable package rename regression

- Focused tests prove bounded retry for Windows winerror 5/32/33 only and no retry for unrelated errors. Full portable suite: 304 passed. Package manifest verified: 229 files / 135,694,592 bytes.
- An injected test cannot establish the cause of the earlier real transient access denial. Repeatable package publishing on broader Windows policy/AV configurations remains unverified.

| Case | Actions | Pass evidence | Minimum release |
|---|---|---|---|
| W01 | Launch onedir EXE as normal user | App opens, no missing DLL/Qt errors, correct version, no admin UAC | v1.3.1 |
| W02 | Import 3 mixed files with numbered prefixes into F1 | Excel queries present with exact M text & order log; file saved/reopens | v1.3.1 |
| W03 | Repeat replace vs skip on F2 | Existing load/table/links preserved, skip truly untouched | v1.3.1/1.5 |
| W04 | Force simulated fail during apply, confirm original+backup hash | Original unchanged, backup equals pre-change and truthful outcome | v1.3.1 |
| W05 | F3 macro security and workbook integrity | Macro not executed, macro project remains, automation settings restored | v1.3.1 |
| W06 | F5 read-only/lock protection | Graceful block and zero original changes | v1.3.1 |
| W07 | Inspect F1/F2; export all, close without apply | Query mapping/status correct; hash unchanged; exported text equal | v1.4 |
| W08 | Load profile, mutate row, save/reopen | Stable order, paths, load mode; schema validation and migration | v1.4 |
| W09 | UI 1366x768 and 125%/150% scaling | No cut-off important controls, keyboard works, contrast and error labels | v1.4 |
| W10 | Source to new sheet | Functional table, correct row values, no unrelated changes | v1.5 |
| W11 | Source to exact existing cell | Correct ListObject top-left and no collisions/overwrites | v1.5 |
| W12 | Data Model load if available | Model membership observed; re-open and repeat import stable | v1.5 |
| W13 | Connection-only distinct state (if available) | Connection exists and no sheet/model; else feature disabled | v1.5 |
| W14 | Dependency cycle and missing reference | Warn, prevent dangerous apply; no silent reordering | v1.5 |
| W15 | Refresh selected/all, one disconnected source | Accurate per-query outcome, timeout, no erroneous success | v1.5 |
| W16 | 2 disposable batch workbooks, 1 read-only | One completes; other safely fails; separate backups/results | v2.0 |
| W17 | Package export/import with non-ASCII and special names | Files/schema/hash roundtrip and safe extraction | v2.0 |
| W18 | CLI `--dry-run` if implemented | JSON deterministic; original file hash unchanged | v2.0 |
| W19 | Repeat launch/import 3x then exit | No Excel.exe COM-process leaks owned by app; user Excel instances survive | all |
| W20 | Non-admin app in restricted corporate-like user context | App launches from local writable folder if policy allows; no admin installs | all |
| W21 | Replace query in workbook with a separate ordinary range-backed Excel table | Query formula changes; ordinary table name, range, data, and load associations remain verified after reopen | v1.5 |
| W22 | Hold a disposable workbook open in another process without an Excel `~$` owner file, then run import preflight; repeat with Excel holding a workbook | Restart Manager detects active use; import/refresh/restore refuse before mutation, and no other application is closed | v1.3.1 |

### W74 — 2026-10-22 — packaged GUI mixed `.m`/`.pq` numeric-prefix order

- Command: `python -B tools/run_packaged_gui_import_smoke.py --mixed-import` using the current one-folder package.
- Host: Windows 10 build 26300; Excel 16.0 build 20430.0. Three disposable static sources (`01_First.m`, `02_Second.pq`, `03_Third.m`) were loaded through a versioned profile whose rows were serialized in reverse order.
- PASS: the packaged UI review showed numeric-prefix order 1/2/3; Excel query collection order matched; read-only reopen returned exact M formulas for all three; original-workbook backup hash matched; manager exited 0 and all smoke-owned Excel processes exited. Existing Excel PID 13156 was untouched.
- Scope: this verifies the tested packaged GUI/profile path and static definitions on one Office build. Dependency evaluation, external sources, other builds and broader W02 workbook fixtures remain outside this acceptance.

### W75 — 2026-10-22 — packaged GUI replacement of an existing query

- Command: `python -B tools/run_packaged_gui_import_smoke.py --replace-existing` using the current one-folder package.
- Host: Windows 10 build 26300; Excel 16.0 build 20430.0. The disposable workbook was seeded with `GuiSmoke` using an older static formula; the versioned profile selected Replace and imported changed source text.
- PASS: packaged UI completed replacement; read-only Excel reopen returned the exact new formula; the timestamped backup matched the original workbook SHA-256; the unrelated worksheet marker stayed unchanged; manager exit was 0 and all smoke-owned Excel processes exited. Existing Excel PID 13156 was untouched.
- Scope: definition-only replacement of one query. Existing worksheet/Data Model load associations, related tables and links, other Office builds, and non-static/external M remain unverified.

### W76 — 2026-10-22 — packaged GUI preserves existing connection-only load

- A first acceptance attempt exposed that `preserve_existing` was rejected as an unsupported profile load mode before import. The UI had treated it like an Excel capability that needed probing, although it is a preservation policy.
- Added regressions that failed before their fixes and passed after: the model permits `preserve_existing` before/after probing, and the post-probe load-mode delegate continues to offer it. Capability-probed destinations remain gated. Rebuilt with `app/build.ps1`: full app suite **310 passed in 16.06s**, then PyInstaller onedir build passed.
- `python -B tools/run_packaged_gui_import_smoke.py --replace-preserve-load --exe app/dist/PowerQueryManager/PowerQueryManager.exe` passed on Windows 10 build 26300 / Excel 16.0 build 20430.0. A disposable connection-only query was replaced through a versioned GUI profile; exact new formula readback, preserved `connection_only` state, verified pre-import backup, manager exit 0 and smoke-owned Excel cleanup passed. Existing Excel PID 13156 remained untouched.
- Scope: one Office build and connection-only preservation. Worksheet/Data Model association preservation through packaged GUI, other Office builds and external sources remain open.

### Observed results (2026-10-21)

- **W16 homogeneous batch PASS (real Excel):** Windows 10 build 26300, Excel 16.0 build 20326.0, Python 3.11.9 x64, pywin32 310. `tools/run_excel_deployment_smoke.py` applied one static definition-only query to two disposable local `.xlsx` files; both first-run results were `applied`, both repeat results were `no_changes`, both backup hashes matched, and all nine smoke-owned Excel PIDs exited. See `EXCEL_COM_SPIKES.md`.
- **W16 mixed read-only PASS (real Excel):** Windows 10 build 26300, Excel 16.0 build 20430.0, Python 3.11.9 x64. `tools/run_excel_deployment_mixed_smoke.py` deployed to two disposable workbooks, with the second target's Windows read-only attribute set. Results were `applied` and `blocked_preflight`; the writable target's backup matched its original hash and its query formula read back correctly. The blocked target had no backup and its SHA-256 remained unchanged. Both owned Excel PIDs exited. This verifies read-only-target isolation on this host/build, not arbitrary mid-commit COM failures or all filesystem ACL configurations.
- **W15 partial PASS (real Excel):** On the same host/build, `tools/run_excel_cli_refresh_smoke.py` exercised the CLI refresh command in background and synchronous modes against one connection-only static `#table` query. Both returned `refreshed`; formula/load state and exact backups were verified and owned Excel processes exited. The disconnected/external-source case and failure/timeouts remain unverified. See `EXCEL_COM_SPIKES.md`.
- **W17 partial PASS (real Excel):** On the same host/build, `tools/run_excel_cli_package_import_smoke.py` safely imported two package-v2 standalone definitions into a blank disposable `.xlsx`; names/formulas, original-hash backup, and owned-process cleanup passed. Special/non-ASCII names, `.xlsm`, and replacement/skip against existing loaded queries remain unverified. See `EXCEL_COM_SPIKES.md`.
- **W17 packaged import baseline (real Excel):** Windows 10 build 26300, Excel 16.0 build 20430.0, Python 3.11.9 x64, pywin32 310. The original two-query packaged CLI import smoke passed with exact formulas, backup hash, and process exit. The Unicode-name extension below supersedes this baseline. See `EXCEL_COM_SPIKES.md`.
- **W17 PASS (real Excel):** Windows 10 build 26300, Excel 16.0 build 20430.0, Python 3.11.9 x64. `tools/run_excel_cli_package_import_smoke.py` created three queries in a disposable source workbook, exported its read-only inventory to a versioned query package, then launched the packaged CLI to import that package into a second disposable `.xlsx`. The package included `Équipe – Résultats 2026 (Q1)`. All names and exact formulas read back, the source workbook hash stayed unchanged during export, the target pre-import backup matched, and owned Excel processes exited. This verifies the named W17 path on one Office build; `.xlsm`, replace/skip against existing queries, GUI import, and other Office builds remain unverified.
- **W19 partial PASS (real Excel):** the controlled timeout cleanup smoke on Excel 16.0 build 20430.0 terminated one hidden DispatchEx instance and independently observed its exact process handle signal exit. It did not simulate a blocked COM call; the repeated three-launch/import cycle in W19 remains unverified.
- **W21 partial PASS (real Excel):** Windows 10 build 26300, Excel 16.0 build 20430.0, Python 3.11.9 x64, pywin32 310. `tools/run_excel_preserve_smoke.py` replaced a connection-only static-M query while preserving a separate plain-range table's name/range/data/style/filter, a workbook defined name, a cell formula, and the ForceFullCalculation property; query connection state remained connection-only and all smoke-owned Excel PIDs exited. The test does not cover worksheet/Data Model loads, `.xlsm`, external tables, privacy settings, report references, or other Office builds. See `EXCEL_COM_SPIKES.md`.
- **W22 partial PASS (initial Restart Manager baseline):** Windows 10 build 26300. A subprocess held a disposable `.xlsx` open with no Excel owner-lock file; Restart Manager detected it and write preflight refused. Excel 16.0 build 20430.0 was independently detected while holding another disposable workbook. The initial test did not cover a read-only Excel open with no lock file; W28 now covers that case through the exact-path ROT supplement. OneDrive and Protected View remain unverified. See `EXCEL_COM_SPIKES.md`.
- **W23 partial PASS (abandoned stage recovery):** Windows 10 build 26300, Excel 16.0 build 20430.0, Python 3.11.9 x64, PySide6 6.8.3, pywin32 310. `tools/run_excel_recovery_smoke.py` created a query-bearing disposable `.xlsx`, copied it to the hidden import/refresh stage pattern, verified it read-only with the recovery worker, wrote a new copy, and reopened that copy to check the exact query name/formula. The original hash stayed unchanged, the stage remained, output hash matched, and seed plus both recovery-owned Excel PIDs exited. Forced crash, partial-save, `.xlsm`, and GUI interaction remain unverified. See `EXCEL_COM_SPIKES.md`.
- **W24 partial PASS (packaged GUI startup/close):** Windows 10 build 26300. Launched the one-folder GUI EXE, found the visible Power Query Manager window, confirmed Qt6Core/Qt6Gui/Qt6Widgets modules loaded, sent WM_CLOSE to that exact process-owned window, and observed exit code 0. This did not perform Excel COM import through the GUI or test standard-user privileges.
- **W25 automated UI PASS (layout/keyboard):** Windows Qt offscreen suite at 960×600 and 1366×768 confirms the workspace has no horizontal scroll, the query table remains at least 140 px high, Tab moves through workbook controls and query cells, and key controls expose accessible names. This is not a physical laptop, screen-reader, or 125%/150% DPI acceptance.
- **W26 combined source mapping PASS (real Excel):** Windows 10 build 26300, Excel 16.0 build 20430.0, Python 3.11.9 x64. `tools/run_excel_cli_environment_mapping_smoke.py --packaged` used the interactive packaged CLI to approve path-root mapping first, then preview/approve M-code rewriting against the mapped source file. A disposable definition-only deployment read back the mapped and untouched formulas exactly; hashes of all original/mapped source files stayed unchanged; backup and owned-process checks passed. Other Office builds and worksheet/Data Model evaluation remain unverified.
- **W27 partial PASS (real Excel):** Windows 10 build 26300, Excel 16.0 build 20430.0, Python 3.11.9 x64, pywin32 310. `tools/run_excel_history_compare_smoke.py` captured two snapshots of a disposable workbook around a static-M formula update, query addition, and capability-probed definition-only → connection-only transition. `SnapshotCompareWorker` returned formula/load `MODIFIED` results with the exact diff and destination reason, `UNCHANGED`, and `ADDED`; both snapshot hashes stayed unchanged and all six smoke-owned Excel processes exited. No query refreshed. Worksheet/Data Model destination transitions, duplicate/unknown identities, `.xlsm`, concurrency/corruption cases, and other Office builds remain unverified.
- **W28 PASS (real Excel open without owner lock):** Windows 10 build 26300, Excel 16.0 build 20430.0, Python 3.11.9 x64, pywin32 310. `tools/run_excel_open_without_lock_smoke.py` opened a disposable workbook read-only in desktop Excel; Excel kept the workbook registered by exact path in the COM Running Object Table but created no `~$` owner-lock file and Restart Manager alone returned no process. The enhanced active-use check found `Microsoft Excel (open workbook)`, and the real `validate_workbook_for_write` rejected before opening a write handle. The workbook hash stayed unchanged and the smoke-owned Excel PID exited. This validates local current-user Excel ROT detection on one build, not other Office apps, remote sessions, or every Excel workbook-open state.
- **W29 PASS (installer after ROT preflight change):** Windows 10 build 26300, Excel 16.0 build 20430.0, Python 3.11.9 x64, PySide6 6.8.3, pywin32 310. Re-ran `tools/run_excel_import_smoke.py` using the current source and venv. Capability probe, staged connection-only create, repeat replace, read-only formula/load verification, and owned Excel cleanup all passed. This checks the new ROT check does not break an ordinary closed-target import; it does not verify packaged GUI import or all load modes.
- **W30 partial PASS (protected-workbook modal timeout after ROT preflight change):** Windows 10 build 26300, Excel 16.0 build 20430.0, Python 3.11.9 x64, pywin32 310. `tools/run_excel_password_safety_smoke.py` caused real Excel password prompts for a disposable protected `.xlsx`; both read-only inspection and staged import timed out at the configured 10 seconds, confirmed termination/exit of only their owned Excel processes, and preserved the workbook SHA-256. This verifies password-modal open timeouts, not a hung non-modal COM call or M evaluation timeout.
- **W31 PASS (all load paths after ROT preflight change):** Windows 10 build 26300, Excel 16.0 build 20430.0, Python 3.11.9 x64, PySide6 6.8.3, pywin32 310. `tools/run_excel_load_smoke.py` capability-probed all five modes; definition-only/preserve-existing were also reported supported. It created and repeat/changed-formula reimported new-sheet, existing-sheet at `Output!B3`, and Data Model queries. Read-only reopen confirmed exact formulas, expected worksheet/Data Model mappings, result values `3, 4`, and unrelated marker preservation. All seed/import/inspection/value-check Excel processes exited. Static inline M only; external connectors, `.xlsm`, timeouts during evaluation, and other Office builds remain unverified.
- **W32 partial PASS (.xlsm file container):** Windows 10 build 26300, Excel 16.0 build 20430.0, Python 3.11.9 x64, PySide6 6.8.3, pywin32 310. `tools/run_excel_import_smoke.py --xlsm` used Excel FileFormat 52 to create a disposable `.xlsm` without a VBA project. Staged connection-only create/repeat-replace, exact formula/load-state readback, and cleanup of smoke-owned Excel processes passed. Real VBA-project preservation, macro security, XLM prompts, `.xlsm` GUI import, and other builds remain unverified.
- **W33 PASS (.xlsx default-path regression):** Windows 10 build 26300, Excel 16.0 build 20430.0, Python 3.11.9 x64, PySide6 6.8.3, pywin32 310. After adding the `--xlsm` option to `tools/run_excel_import_smoke.py`, the default invocation passed staged `.xlsx` connection-only create/repeat-replace, exact formula/load readback, and cleanup of smoke-owned Excel processes.
- **W34 PASS (packaged GUI import):** Windows 10 build 26300, Excel 16.0 build 20430.0, Python 3.11.9 x64, PySide6 6.8.3, pywin32 310. `tools/run_packaged_gui_import_smoke.py` created a disposable `.xlsx` in desktop Excel, selected a static `.m` source in the visible packaged GUI, confirmed definition-only import, and verified the exact M formula through read-only Excel reopen. The target hash changed, the timestamped backup matched the original hash, worksheet marker `A1` was preserved, the packaged manager exited 0, and all smoke-owned Excel PIDs exited while the pre-existing Excel instance was untouched. One build and `.xlsx` only; standard-user permissions, macro-bearing `.xlsm`, and evaluated/external loads remain open.
- **W35 PASS (rebuilt GUI + timeout cleanup regression):** Windows 10 build 26300, Excel 16.0 build 20430.0, Python 3.11.9 x64, PyInstaller 6.13.0, PySide6 6.8.3, pywin32 310. `app/build.ps1` rebuilt the current source and passed 255 tests; `tools/run_packaged_gui_import_smoke.py --exe .\app\dist\PowerQueryManager\PowerQueryManager.exe` passed the same disposable `.xlsx` definition-only flow with exact formula readback, original-hash backup, unchanged A1 marker, GUI exit 0, and owned Excel cleanup. `tools/run_excel_timeout_cleanup_smoke.py` also confirmed exit of its exact smoke-owned Excel process after the bounded timeout. Its callback is a controlled delay, not a blocked COM call or M evaluation. The user-open `dist/` executable was not replaced; standard-user rights, VBA-bearing `.xlsm`, other builds, and evaluated/external loads remain open.
- **W36 PASS (shared application adapter build):** Windows 10 build 26300, Excel 16.0 build 20430.0, Python 3.11.9 x64, PyInstaller 6.13.0, PySide6 6.8.3, pywin32 310. `app/build.ps1` rebuilt the shared `ExcelApplicationAdapter` source with **258 tests passed**. The rebuilt `app/dist/PowerQueryManager/PowerQueryManager.exe` passed the disposable GUI definition-only import/reopen smoke with exact formula readback, original-hash backup, preserved A1 marker, GUI exit 0 and owned Excel cleanup. Its paired `PowerQueryManagerCLI.exe --help` exited 0. Root `dist/` remained untouched because its GUI is open. This is one Excel build and `.xlsx` definition-only coverage; adapter transaction extraction, blocked COM/M-evaluation hangs, standard-user rights, VBA-bearing `.xlsm`, and other builds remain open.
- **W37 PASS (history path containment and retention):** Windows 10 build 26300, Python 3.11.9 x64. Two test-first regressions demonstrated that pruning a tampered SQLite snapshot path deleted an unrelated file, and that an expired snapshot could still be retrieved/restored directly before pruning. Cleanup now resolves and constrains stored paths to the snapshot directory, and retrieval prunes expired versions first. Full suite: **260 passed**; contract validation and compileall passed; `app/build.ps1` passed the same 260 tests and rebuilt the onedir app; packaged CLI `--help` exited 0; release package verification passed for `dist-next` (227 files / 135,082,015 bytes). No new Excel workbook operation was run for this history-only change. Root `dist/` remains unchanged because its GUI is still live; overall Excel release gates remain open.

**Important:** a fake-COM `Workbook.Queries.Add` PASS is never a worksheet/table/Data Model load PASS. For modes depending on data source access, prefer a pure M fixture `Table.FromRecords` with no external credentials. Document if firewall/privacy warnings are from the Excel test host rather than importer.

## Windows host local commands

Open PowerShell in package root:

```powershell
python .\tools\verify_handoff.py
.\tools\check_windows_env.ps1
.\tools\run_tests.ps1
.\tools\build_and_smoke.ps1

# Optional packaged GUI import smoke: run in an interactive, non-elevated desktop session.
python -m pip install -r .\tools\requirements_ui_smoke.txt
python .\tools\run_packaged_gui_import_smoke.py
```

The scripts do NOT automatically open production workbooks. A manual tester must supply disposable fixtures and explicitly confirm COM tests, then record output in `evidence/`.

## Evidence requirements

For each row record `PASS/FAIL/BLOCKED/UNVERIFIED`, test ID, timestamp, Office/Windows builds, expected/actual result, artifact path and corrective issue reference. No completed status without a real observed value. Example blank form in `evidence/TEST_REPORT_TEMPLATE.md`.

- **W38 PASS (current packaged GUI):** Windows 10 build 26300, Excel 16.0 build 20430.0, Python 3.11.9 x64, PyInstaller 6.13.0, PySide6 6.8.3, pywin32 310. `tools/run_packaged_gui_import_smoke.py --exe .\dist-next\PowerQueryManager\PowerQueryManager.exe` completed a visible GUI definition-only import into its disposable `.xlsx`. Exact formula readback, timestamped backup hash, preserved A1 marker, manager exit 0, and Excel process-set equality passed. Pre-existing Excel PID 13156 and user-open root package GUI PID 26288 were untouched. This is a repeat of the scoped W34 case using the latest rebuilt package; `.xlsm`, evaluated/external loads, standard-user rights, blocked COM timeout, more Office builds, and manual accessibility/scaling remain open.

- **W39 PASS (staged recovery symlink containment and rebuilt package):** Windows 10 build 26300, Excel 16.0 build 20430.0, Python 3.11.9 x64, PyInstaller 6.13.0, PySide6 6.8.3, pywin32 310. A test-first regression showed `find_staged_workbooks()` followed a matching symlink to a file outside the target directory. Discovery now resolves the workbook folder and refuses symlink candidates or candidates resolving outside it. Focused recovery tests: **7 passed**; full suite and build: **261 passed**; contracts and compileall passed. Root onedir package verifier passed at 227 files / 135,082,362 bytes; CLI help exited 0. Packaged GUI definition-only import/reopen passed on Excel 16.0 build 20430.0 with a disposable `.xlsx`, exact formula, verified backup, preserved marker, manager exit and owned-process cleanup; pre-existing Excel PID 13156 remained untouched. Forced interruption phases, blocked COM timeout, `.xlsm`, and stale-stage diagnosis remain open.

- **W40 PASS (stale-stage triage and rebuilt package):** Windows 10 build 26300, Excel 16.0 build 20430.0, Python 3.11.9 x64, PyInstaller 6.13.0, PySide6 6.8.3, pywin32 310. `diagnose_staged_workbook()` now reports stage size/age/mtime and preliminary ZIP/Open XML structure; this does not prove transaction completion. The picker warns that Excel read-only verification remains required and handles candidates that disappear/change. Focused recovery tests: **9 passed**; full suite: **263 passed**; contracts, compileall, and build passed. Root package verifier: 227 files / 135,084,909 bytes. Packaged GUI import/reopen passed on Excel 16.0 build 20430.0 with exact formula, backup, marker, clean exit, and owned-process checks. The recovery dialog’s actual Excel path is still only covered by the prior worker smoke; forced interruption, blocked COM/M evaluation, `.xlsm`, standard-user rights, and manual accessibility/scaling remain open.

- **W41 PASS (portable final staged-commit failure):** Fake-COM regression injects a final `os.replace` failure after staged save/reopen. The original workbook bytes remain unchanged, the original-hash backup remains, stage cleanup completes, and the query outcome is `ROLLED_BACK` with an operation error. Focused test: 1 passed; full portable suite: 264 passed; contracts and compileall passed. This does not verify a real Excel lock, forced termination, or blocked COM call.

- **W42 PASS (snapshot comparison duplicate ambiguity):** A test-first regression reproduced a dropped later query when one earlier query matched multiple later case-insensitive duplicate names. The comparison now returns conflict records for every query on both sides without pairing formulas or load states. Focused suite: 6 passed; full suite: 265 passed; contracts and compileall passed. Windows build, 227-file package verification (135,084,785 bytes), packaged CLI help, and the independent GUI import smoke passed. The GUI smoke does not certify snapshot comparison. T200-05 real Excel ambiguity/worksheet/model/multi-build matrix remains open.
- **W43 PASS (shared staged-file transaction adapter):** Extracted import/refresh sibling-stage copy, original fingerprint and active-user checks, atomic replace, and stage cleanup into the injectable `WorkbookTransactionAdapter`; COM workbook operations remain service-owned. Focused transaction/import/refresh tests: 43 passed; full portable suite: 270 passed; Draft 2020-12 contracts and compileall passed. Windows onedir build passed its embedded 270-test run. Root package verified at 227 files / 135,084,785 bytes; packaged CLI help exited 0. Packaged GUI definition-only disposable `.xlsx` import passed on Excel 16.0 build 20430.0 with exact formula, verified backup, read-only reopen, manager exit 0, and owned-process cleanup. This does not verify blocked COM, forced termination during each transaction phase, Protected View, OneDrive, VBA-bearing `.xlsm`, standard-user rights, or other Office builds. T131-02 remains partial and release gates remain UNVERIFIED.
- **W43 refresh extension PASS:** The real-Excel background refresh smoke used a generated disposable workbook with a static `#table` M expression and no external source. Excel 16.0 build 20430.0 reported a successful refresh; the original-hash backup matched, the query reopened with its exact formula/load state, and the refresh-owned Excel process exited. This verifies the extracted refresh staging/commit path only for this narrow scenario.
- **W44 PASS (bounded staged verification and save calls):** Test-first regressions found that final read-only verification opens in import/refresh bypassed the owned-Excel timeout helper. Both staged saves and verification opens now go through the bounded process-aware wrapper. Full suite: 272 passed; contracts and compileall passed; Windows onedir build passed its embedded 272 tests. Static background refresh and packaged GUI definition-only import both passed against disposable workbooks on Excel 16.0 build 20430.0 with backup/readback and owned-process checks. Root package verified at 227 files / 135,090,027 bytes; packaged CLI help exited 0. This does not validate modal dialogs or arbitrary blocked COM methods; release gates remain UNVERIFIED.
- **W45 PASS (current-source Excel load matrix):** `tools/run_excel_load_smoke.py` passed the live capability probe and staged installer for new worksheet, existing worksheet at `Output!B3`, and Data Model on Windows 10 build 26300 / Excel 16.0 build 20430.0. Create, repeat import, changed-formula reimport, formula/load readback, output values 3/4, unrelated marker preservation, and owned-process cleanup passed for all three. Only static inline M and generated disposable `.xlsx` workbooks were used. External connectors, VBA-bearing workbooks, Protected View, OneDrive, other Excel builds, and interruption during M evaluation remain unverified.
- **W46 PASS (retained-stage warning and rebuilt package):** Test-first fake-adapter regressions reproduced silent stage retention after cleanup failure in both import and refresh. Both paths now log the retained stage filename without masking the original operation error. Full suite: 274 passed; contracts and compileall passed; onedir build passed its embedded tests. Root package verified at 227 files / 135,090,793 bytes; packaged CLI help exited 0; packaged GUI import/reopen passed on Excel 16.0 build 20430.0. Stage-retention injection is portable fake-COM coverage; remaining recovery and Office matrix items remain open.
- **W47 PASS (automated Qt scale-factor UI checks):** All 53 UI tests passed with `QT_SCALE_FACTOR=1.25` and again with `QT_SCALE_FACTOR=1.5` on Windows 10 build 26300 / Python 3.11.9 / PySide6 6.8.3. This checks Qt's scaled offscreen/test geometry and keyboard assertions; it is not a substitute for Windows display-scaling, screen-reader, or physical laptop acceptance.
- **W48 PASS (release documentation freshness):** The package stager's regression now requires retained-stage recovery guidance and current W46/W47 acceptance notes in generated docs. It passed; the full suite passed 274 tests, contracts and compileall passed, and `package_release.py --docs-only` refreshed root docs/SHA256SUMS without touching the executables. Package verifier passed 227 files / 135,091,032 bytes. Packaged GUI acceptance remains W46; this docs-only refresh does not claim a new binary build.
- **W49 PASS (bounded capability probe and rebuilt package):** Added 60-second owned-process bounds around capability-probe load evaluation, save/open/close, connection creation and shutdown; injected timeout now aborts the probe and disables unverified modes. Full suite: 276 passed; contracts and compileall passed; Windows onedir build succeeded. Current-source live probe/load matrix passed new worksheet, existing sheet, and Data Model on Excel 16.0 build 20430.0; packaged GUI import/reopen also passed after rebuild. Initial root package verified at 227 files / 135,094,444 bytes; CLI help exited 0. Arbitrary blocked COM/property calls and actual evaluation timeout remain unverified.
- **W50 PASS (post-W49 manual refresh):** After updating the generated release manual and summary with W49, `package_release.py --docs-only` refreshed the root documentation and SHA256SUMS without replacing binaries. Package verifier passed at 227 files / 135,094,766 bytes; package-staging test passed. This did not rebuild or retest the executable.
- **W51 PASS (symbolic-link write protection):** Test-first Windows regressions showed write preflight, stage preparation, and final commit accepted workbook symlink paths. Those boundaries now reject links; refresh uses write preflight and stage commit also requires a real sibling stage. Full suite: 280 passed; contracts/compileall and Windows onedir build passed. Root package verified at 227 files / 135,096,061 bytes; CLI help exited 0. Packaged GUI import/reopen and current-source static background refresh passed on Excel 16.0 build 20430.0. The tests preserve link target bytes and retain stage on target-link substitution. Junction/reparse-point behavior and a real concurrent link-swap race remain untested.
- **W52 PASS (Windows reparse-point rejection):** Added shared `lstat` detection for POSIX symlinks and Windows `FILE_ATTRIBUTE_REPARSE_POINT` across the workbook/stage and each lexical parent path; write preflight, stage preparation, and final commit reject either. Full suite: 285 passed in 11.63s; Draft 2020-12 contracts and compileall passed; Windows onedir build passed with embedded 285-test run. Root package verifier passed at 227 files / 135,099,416 bytes and CLI `--help` exited 0. Packaged GUI definition-only import/reopen and static-M background refresh passed on Excel 16.0 build 20430.0; both used disposable workbooks, verified backups/readback, and cleaned up only their own Excel process. The packaged GUI smoke ran under the current non-elevated `pcrn\robert` token (`IsInRole(Administrator) = False`). A real Windows junction test verified preflight/stage/commit refusal while preserving target and staged bytes; a concurrent path-swap race was not exercised. Broader Excel release gates remain unverified.
- **W53 PASS (last-boundary path recheck):** Test-first regression reproduced a path becoming a reparse point after initial commit checks; commit now repeats workbook/stage reparse and sibling-directory validation immediately before `os.replace`. Full suite: 286 passed in 13.75s; contracts and compileall passed; onedir build passed its embedded suite. Root package verified at 227 files / 135,099,926 bytes; CLI help exited 0; package-staging and focused late-change/real-junction regressions passed. Packaged GUI import and static-M refresh passed again on Excel 16.0 build 20430.0. This narrows but does not eliminate the OS-level race between the final path check and atomic replace; no kernel-level race claim is made.
- **W54 PASS (recovery reparse protection):** Test-first real-junction regression showed recovery-copy output through a junction was published into the junction target. Recovery discovery, diagnosis and copy-out now reject reparse points across workbook/stage/output paths and their parents, preserving source and target bytes; path handling uses lexical absolute paths instead of resolving through links. Recovery tests: 10 passed, including actual junction refusal. Full suite: 287 passed in 15.23s; contracts and compileall passed; Windows onedir build passed its embedded suite. Package verified at 227 files / 135,100,561 bytes; CLI help passed. Rebuilt packaged GUI definition-only import/reopen passed on Excel 16.0 build 20430.0 with exact formula, backup, read-only reopen and owned-process cleanup. Broader recovery interruption and Excel matrix gates remain open.
- **W55 PASS (recovery prepublish recheck):** Test-first regression reproduced a recovery destination changing to a reparse point after initial validation and before publication. Recovery now repeats workbook/stage/destination/temp path checks after the copy is verified and immediately before `os.link`. The late-change regression and all 11 recovery tests passed. Full suite: 288 passed in 12.77s; contracts and compileall passed; Windows onedir build passed its embedded suite. Package verified at 227 files / 135,100,924 bytes; CLI help and recovery/package-staging focused tests passed. Rebuilt packaged GUI import/reopen passed on Excel 16.0 build 20430.0. This reduces but does not remove the OS-level race after the final check.

- **W56 PASS (related report-reference preservation):** Test-first regression showed that a plain `xlSrcRange` ListObject with no QueryTable blocked unrelated Power Query replacement; the query-table mapper now skips only this explicit source type. A disposable Excel 16.0 build 20430.0 workbook then passed changed-formula replacement while retaining a normal table's name/range/data/style/filter, a defined name, a cell formula, `ForceFullCalculation`, a connection-only query and a worksheet-loaded query table used by report formulas and a defined name; both report values recalculated to 70. Cleanup used exact smoke-owned Excel PIDs only; a process that did not exit within 20 seconds after workbook close/`Quit` was identity-checked as Excel and stopped, leaving pre-existing PID 13156 untouched. Full suite: 289 passed in 12.58s; build's embedded suite: 289 passed in 11.97s; contracts, compileall and handoff verifier passed; package verified at 227 files / 135,101,262 bytes; packaged CLI `--help` exited 0; rebuilt packaged GUI definition-only import/reopen passed with exact formula, backup, app exit and Excel cleanup. One Office build and static inline M only; privacy, `.xlsm` VBA preservation, Data Model report references, external connectors, blocked M evaluation and broader desktop accessibility remain unverified.


- **W57 PASS (history comparison of observed load modes):** Added a regression asserting that an observed `worksheet: Output!$A$1:$A$3` state compared with `data_model` is reported as `MODIFIED` with both destinations in the reason. The Excel smoke retained the same-workbook definition-only → connection-only history case and created equivalent disposable `.xlsx` files separately for new-sheet and Data Model modes; snapshot comparison detected `worksheet: HistoryOutput!$A$1:$A$2` → `data_model`. Exact M diffs, Added/Unchanged/Modified outcomes, immutable snapshot hashes, no refresh, and exit of all 12 smoke-owned Excel processes were verified on Windows 10 build 26300 / Excel 16.0 build 20430.0. Portable full suite: 290 passed in 11.83s; compileall and contracts passed; docs-only package refresh and verifier passed at 227 files / 135,101,514 bytes. The importer refuses implicit conversion of an existing query's load mode, so this is cross-snapshot comparison evidence, not same-workbook conversion acceptance. Other Office builds, `.xlsm`, corruption/ambiguous inventory and broader recovery remain unverified.


- **W58 PASS (fail-closed load-mode conversion):** A new fake-COM regression requests a worksheet load for a query already mapped connection-only. The importer reports that implicit destination alteration is refused, preserves the original formula and connection object, adds no query definition, and shuts down its injected Excel adapter. The regression passed; full suite: **291 passed in 13.22s** and compileall passed. This records the supported fail-closed behavior; it does not add load-mode conversion capability. Package-staging regression passed; docs-only package refresh and verifier passed at 227 files / 135,101,633 bytes.


- **W59 PASS (packaged GUI import under non-admin token):** Current root `dist` onedir GUI performed definition-only import into a disposable `.xlsx` on Excel 16.0 build 20430.0. Exact formula, verified backup, read-only reopen, manager exit 0, and Excel process cleanup passed. The current token was `PCRN\Robert`; `WindowsPrincipal.IsInRole(Administrator)` returned `False`. QA-only `pywinauto`/`comtypes` dependencies were installed into a temp directory outside the workspace. This verifies that tested path under a non-admin token, not broader ACL, Group Policy, network share, or multi-user install conditions. Docs-only package refresh and verification passed at 227 files / 135,101,801 bytes.


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


## W69 snapshot comparison event recording — 2026-10-22

- A test-first regression proved `HistoryStore` dropped the before/after snapshot IDs from comparison metadata. The history allowlist now retains those UUID fields, and the UI records `workbook_snapshots_compared` only when the read-only comparison completes. Formula/code fields are filtered and secret keys redacted.
- Focused regression passed; full app suite: **301 passed in 13.22s**. The onedir build passed its embedded **301-test** run. Contracts passed. The package was rebuilt and verified at **229 files / 135,694,204 bytes**; CLI help exited 0; packaged history UI Automation and disposable `.xlsx` GUI import passed on Windows 10 build 26300 / Excel 16.0 build 20430.0.
- History records only completion and the two snapshot IDs; it does not store workbook paths, query formulas, or diff contents. Broader history corruption/recovery and Office matrix validation remain open.


## W70 chronological snapshot comparison — 2026-10-22

- A new UI integration regression exposed that snapshot rows are displayed newest-first but were passed to the comparison worker in visual order, reversing “before” and “after.” The history dialog now orders the two selected snapshots by their recorded UTC timestamp before comparison.
- Regression passed with a fake compare worker and temporary snapshot files; full portable suite: **302 passed in 12.66s**. Windows onedir build passed its embedded 302-test suite. The rebuilt package verified at **229 files / 135,694,494 bytes**; CLI help, packaged history-dialog UI Automation, and disposable `.xlsx` GUI import/readback passed on Excel 16.0 build 20430.0. Existing Excel PID 13156 was untouched.
- This verifies timestamp ordering in the UI path and packaged regression set. Real Excel history comparison across snapshots remains scoped to the W57 data/inventory smoke; broader ambiguity/corruption and multi-build validation remain open.


### W77 — 2026-10-22 — controlled owned-Excel timeout cleanup

- `python -B tools/run_excel_timeout_cleanup_smoke.py` passed on Excel 16.0 build 20430.0 / Python 3.11.9. A 2-second callback with a 1-second timeout triggered exact owned-process termination and signaled exit; the Excel process inventory returned to pre-existing PID 13156 only.
- No workbook was opened. This does not cover arbitrary blocked COM property calls, real M evaluation timeout, or other Office builds.


## W78 packaged GUI new-worksheet load — 2026-10-22

- Windows 10 build 26300; Excel 16.0 build 20430.0; packaged one-folder GUI; static inline M only. Existing Excel PID 13156 was not touched.
- `tools/run_packaged_gui_import_smoke.py --new-sheet` probed the current Excel build in the GUI, loaded a versioned profile targeting a new worksheet, and imported into a disposable `.xlsx`.
- PASS: the capability probe enabled `new_sheet`; Excel readback verified exact formula and worksheet load state; `GuiOutput!tblGuiSmoke` contained values 17 and 29; unrelated marker content remained; the backup matched the original hash; manager exit was 0 and owned Excel processes exited.
- Limits: one Office build, one static inline M query, and one-folder package. This does not validate external connectors, other Office builds, or arbitrary M evaluation/refresh timeouts.


## W80 post-W79 packaged GUI definition-only import — 2026-10-22

- Windows 10 build 26300; Excel 16.0 build 20430.0; current root onedir package. Existing Excel PID 13156 was not touched.
- `python -B tools/run_packaged_gui_import_smoke.py` passed against a disposable `.xlsx`: exact M formula, pre-import backup hash, read-only reopen, manager exit 0, and smoke-owned Excel process cleanup all verified.
- This confirms the rebuilt package's standard GUI/import path after the history restore safety change. It does not expand the workbook format, load-mode, connector, or Office-version matrix.


## W81 snapshot storage-root path regression — 2026-10-22

- `tests/test_history.py::test_snapshot_lookup_rejects_reparse_point_snapshot_root` reproduced the old lookup accepting a redirected snapshots directory, then passed after lexical root/candidate reparse checks were added. Related restore-path regressions verify rejection before resolution and refusal at a simulated late commit-path change.
- Full suite: **313 passed in 15.73s**; PyInstaller onedir GUI/CLI build succeeded. Packaged CLI help passed and the rebuilt GUI passed disposable `.xlsx` definition-only import on Excel 16.0 build 20430.0. These GUI checks serve as package regressions; no Excel workbook was opened for the history-only path tests.
- This does not verify arbitrary kernel races, history database corruption recovery, or other Windows/Office builds.


## W82 post-W81 rebuilt-package smoke — 2026-10-22

- After rebuilding with the final snapshot destination publication guard, the packaged GUI definition-only import passed on Windows 10 build 26300 / Excel 16.0 build 20430.0 using a disposable `.xlsx`. Exact formula, pre-import backup hash, read-only reopen, manager exit 0 and owned-process cleanup passed. Existing Excel PID 13156 was untouched.
- This is a package regression; no history restore or external connector was invoked.


## W83 local history corruption recovery regressions — 2026-10-22

- Portable tests passed for arbitrary corrupt DB bytes and schema mismatch: the original bytes are retained under a unique timestamped archive, a fresh store accepts new events, and WAL sidecars are preserved. An injected `disk I/O error` remains fail-closed and does not move the database. UI test verifies the archive path and recovery notice appear in the operation log.
- Full suite: **319 passed in 14.35s**. No workbook was opened or modified for history database recovery tests. Actual user profile recovery and multi-process Windows race cases remain unverified.


## W84 packaged corrupt-history recovery — 2026-10-22

- `python -B tools/run_packaged_history_a11y_smoke.py --exe dist/PowerQueryManager/PowerQueryManager.exe --corrupt-history` passed. It seeded a corrupt database only under a temporary `LOCALAPPDATA`, opened the packaged UI, verified the exact corrupt bytes were retained in a timestamped archive, confirmed the fresh SQLite store was usable, and confirmed UI Automation exposed the archive name in the operation log. History dialog accessible names/roles passed and manager exit was 0.
- This is a package smoke of the recovery path, not salvage of damaged entries or a real user-profile recovery. The temporary profile was removed; unrelated profile data was not touched.


## W85 final packaged import after history recovery — 2026-10-22

- After the history recovery code was built and packaged, `python -B tools/run_packaged_gui_import_smoke.py` passed on Windows 10 build 26300 / Excel 16.0 build 20430.0 against a disposable `.xlsx`: exact formula, original-hash backup, read-only reopen, manager exit 0 and owned Excel process cleanup. Existing Excel PID 13156 was untouched.


## W86-W87 final history recovery package — 2026-10-22

- Added a failure-path regression after fresh-database initialization fails: the original corrupt database and its WAL remain paired in the archive rather than restoring the WAL beside a partial new store. Full `app/build.ps1` suite: **321 passed in 17.00s**; PyInstaller onedir GUI/CLI build succeeded.
- Package verification passed at **229 files / 135,704,065 bytes**; packaged CLI `--help` exited 0. W84 recovery smoke passed on the final package with isolated temporary `LOCALAPPDATA`, exact corrupt-byte and WAL preservation, functional fresh DB, UI archive-path visibility, history controls and exit 0.
- W85 repeated packaged disposable `.xlsx` definition-only import on Excel 16.0 build 20430.0; exact formula, original-hash backup, read-only reopen, app exit and smoke-owned Excel cleanup passed. Existing Excel PID 13156 was untouched.
- Test results do not cover corrupted user profile recovery with concurrent app instances; stale recovery lock intentionally stops for manual investigation; damaged history rows are not salvaged/merged.


## W88 Windows Internet-zone workbook preflight — 2026-10-22

- Windows 10 build 26300; Excel 16.0 build 20430.0. `tools/run_excel_zone_mark_safety_smoke.py` created a disposable `.xlsx` with an isolated `DispatchEx` Excel instance, closed it, then added a real NTFS `Zone.Identifier` stream with ZoneId=3.
- Production `ExcelQueryInspector` and `ExcelQueryInstaller` both refused before their injected Excel application factories ran. Workbook SHA-256 and Zone.Identifier bytes remained unchanged; no stage was created; the exact smoke-owned seed Excel PID 23424 exited. Existing Excel PID 13156 was not targeted.
- This verifies conservative fail-closed handling of an Internet-marked workbook before COM automation. It does not open Excel Protected View or certify the Protected View UI/security-center policy; OneDrive conflict and other Office session behavior remain unverified.

## W89 rebuilt package GUI regression after zone validation — 2026-10-22

- `app/build.ps1` passed **326 tests in 16.98s** and rebuilt the GUI/CLI. Packaged CLI help exited 0. The current package passed disposable `.xlsx` definition-only GUI import on Excel 16.0 build 20430.0 with exact formula, verified backup, read-only reopen, manager exit 0, and cleanup of its own Excel process. Existing Excel PID 13156 remained untouched.
- Package verified at 229 files / 135,706,475 bytes before the documentation-only refresh.


## W94 packaged GUI existing worksheet destination — 2026-10-22

- Extended `tools/run_packaged_gui_import_smoke.py` with `--existing-sheet`. The packaged GUI probed Excel, loaded a saved versioned profile, and imported static inline M to an empty `Output!C4` destination in a disposable `.xlsx`.
- Excel 16.0 build 20430.0 read-only verification found exact formula and `worksheet: Output!$C$4:$C$6`; the table anchor was `$C$4`, data values were 17/29, unrelated marker was preserved, backup SHA-256 matched the original, manager exited 0, and the smoke-owned Excel process exited. Existing Excel PID 13156 was untouched.
- This verifies the existing-sheet GUI profile route for this Office build and an empty destination. Occupied-sheet GUI refusal and other Office builds remain unverified.


## W95 bounded inspector COM access — 2026-10-22

- Production inspector COM property reads, collection counts/items, query formulas/descriptions, and load mapping calls now each use the owned Excel PID timeout boundary. Timeout exceptions propagate to callers instead of yielding a misleading unknown inventory state.
- Test-first injected timeout regression and inspector tests passed (8 tests); full app build passed 327 tests in 16.91s and rebuilt the onedir GUI/CLI. Rebuilt packaged `.xlsx` definition-only GUI import passed on Excel 16.0 build 20430.0, including read-only inspector verification, matching backup and owned Excel cleanup.
- This does not yet bound every importer/refresh COM property chain or prove cleanup for a real blocked COM call. Other Office builds remain unverified.


## W96-W97 bounded refresh COM calls — 2026-10-22

- W96 bounded connection object lookup, BackgroundQuery read/set/restore, refresh invocation, asynchronous completion polling, and cancellation against the exact owned Excel PID. W97 extends bounds to refresh inventory, staged verification, workbook open/save/close and cleanup. After any operation timeout, further COM cleanup is suppressed because process state may be unknown.
- Test-first regressions cover refresh operation boundaries, timeout propagation, stage discard and no further COM use after confirmed exit. Focused refresh suite: **16 passed**; full `app/build.ps1`: **330 passed in 17.50s** and PyInstaller onedir GUI/CLI build succeeded.
- `tools/run_excel_refresh_smoke.py` passed on Windows 10 build 26300 / Excel 16.0 build 20430.0 using a disposable workbook and static inline M: background refresh succeeded, backup hash matched, query reopened, and the owned Excel process exited. No external source was used. Packaged GUI import smoke passed with exact query formula, verified backup, read-only reopen, manager exit 0, and owned-process cleanup.
- Package verification passed at **229 files / 135,716,084 bytes** and packaged CLI help exited 0. Existing Excel PID 13156 was untouched.
- These checks do not prove behavior when COM termination itself fails or bound every importer property/mutation. Other Office builds and the full v1.3.1 gate remain unverified.


## W98 bounded importer COM calls — 2026-10-22

- Production importer COM calls now use the operation-owned Excel timeout wrapper across Excel configuration and version checks, workbook access/security, query/load inventory, formula capture and mutation, connection creation, worksheet/Data Model load operations, staged save/read-only verification, rollback and shutdown. After any importer timeout, subsequent COM work and cleanup are suppressed because process state may be unknown.
- Test-first regression simulates a query Add that mutates staged state before a timeout with unknown termination result; the stage is discarded, original bytes remain unchanged, and no follow-on COM calls occur; importer suite: **36 passed**. Full `app/build.ps1`: **331 passed in 15.16s** and PyInstaller onedir GUI/CLI build succeeded.
- Packaged definition-only GUI import and existing-sheet import passed against disposable `.xlsx` files on Windows 10 build 26300 / Excel 16.0 build 20430.0. The existing-sheet route verified exact M, `Output!C4` load, table anchor and values 17/29, unrelated marker preservation, matching backup hash, manager exit 0 and owned Excel cleanup. Existing Excel PID 13156 was untouched.
- Bounded operations do not prove timeout behavior if terminating an actually blocked COM process fails, or validate other Office builds and remaining Protected View/OneDrive matrix items.


## W99 current-importer password modal timeout — 2026-10-22

- Updated the disposable password safety smoke to accept the explicit `ExcelOperationTimeout` type now propagated by production COM boundaries. It created a password-protected `.xlsx`, then ran read-only inspection and staged import. Inspection timed out after 10 seconds and importer open after its production 60-second bound; both exact smoke-owned Excel processes exited. Original workbook SHA-256 stayed unchanged and no stage remained.
- The smoke exited 0. Existing Excel PID 13156 remained untouched. This verifies password-modal timeout handling on Excel 16.0 build 20430.0; a truly blocked non-modal COM call, failed process termination, and other Office builds remain unverified.


## W100 runtime-configured COM timeout — 2026-10-22

- Bounded importer and refresh calls now resolve the module timeout default at call time rather than capturing it during import, so runtime policy/test overrides are honored consistently. Two regressions cover the importer and refresh helper settings.
- Importer/refresh focused suite: **54 passed**. Full `app/build.ps1`: **333 passed in 16.39s**; onedir GUI/CLI build succeeded. Root package verified at **229 files / 135,721,094 bytes**; packaged CLI help exited 0.
- Repeated `tools/run_excel_password_safety_smoke.py` against a disposable password-protected `.xlsx`: both inspection and import timed out at the configured **10 seconds**, each owned Excel process exited, source hash stayed unchanged and no stage remained. Rebuilt packaged GUI definition-only import passed on Excel 16.0 build 20430.0 with exact formula, backup, read-only reopen, manager exit 0 and cleanup. Existing Excel PID 13156 remained untouched.
- This confirms configurable modal-open timeout behavior on one Office build; blocked non-modal COM, termination failure and other Office builds remain unverified.


## W101 fail-closed query lookup — 2026-10-22

- Importer lookup previously converted every `Queries.Item` exception into “query absent.” On Excel 16.0 build 20430.0, a disposable workbook showed that genuine absence has outer HRESULT `DISP_E_EXCEPTION` and EXCEPINFO.scode `TYPE_E_ELEMENTNOTFOUND`. The importer now treats only that signature as absence and reports all other lookup errors as unsafe to proceed. A test-first regression failed on the old behavior; focused importer suite: **38 passed**.
- Full `app/build.ps1`: **334 passed in 15.95s** and PyInstaller GUI/CLI onedir build succeeded. The current-source live Excel classifier probe passed. Packaged GUI creation and replacement both passed on disposable `.xlsx` workbooks: exact formula, matching original-hash backup, read-only verification, unrelated marker preservation, manager exit 0 and owned Excel cleanup. Package verification passed at **229 files / 135,721,982 bytes**; CLI `--help` exited 0. Existing Excel PID 13156 was untouched.
- This classifier is based on the observed Mashup signature on one Office build. Other Office versions/providers remain fail-closed until their actual missing-query signatures are verified; the broader release gates remain open.
### W111 — Packaged manager main-window UI Automation — 2026-10-22

- Command: `python -B tools/run_packaged_gui_import_smoke.py --replace-existing` against the current onedir package.
- PASS on Windows 10 build 26300 / Excel 16.0 build 20430.0: UI Automation exposed the accessible names and roles for target workbook, query settings table, source filter, existing-query policy, backup toggle, import action and operation log. The same run replaced a disposable query; exact formula, original-hash backup, unrelated marker, manager exit 0 and owned Excel cleanup passed. Existing Excel PID 13156 was untouched.
- This confirms packaged accessibility metadata and keyboard-independent control discovery; it is not a screen-reader user test, keyboard-only workflow acceptance, OS-level DPI test or physical laptop review.

## W113 packaged main-window keyboard traversal — 2026-10-22

- Command: `python -B tools/run_packaged_keyboard_smoke.py --exe dist/PowerQueryManager/PowerQueryManager.exe`.
- PASS: Windows UI Automation `GetFocusedElement` observed the packaged manager's full 34-control Tab cycle. Workbook path, Add files, query settings table, existing-query policy, backup toggle and import action were reachable in that order and focus stayed within the manager process.
- The manager exited 0. No workbook was opened and no Excel process was created; pre-existing Excel PID 13156 remained untouched. This covers keyboard focus traversal in the packaged main window on Windows 10 build 26300 / Python 3.11.9; screen-reader user testing, physical display scaling, other Office builds and laptop ergonomics remain unverified.

## W114 enabled-source multi-workbook deployment — 2026-10-22

- A test-first regression reproduced that a disabled profile entry with a missing source file aborted deployment before planning. `deploy_multi_workbook` now hashes and freezes enabled sources only and passes only enabled sources to the installer; profile planning still excludes disabled entries. Focused deployment tests: **8 passed**.
- Full `app/build.ps1`: **338 passed in 17.69s**; PyInstaller GUI/CLI onedir build succeeded. The refreshed release package verified.
- `tools/run_excel_deployment_smoke.py` passed on Windows 10 build 26300 / Excel 16.0 build 20430.0: two disposable workbooks applied, repeat deployment returned `no_changes` for both, both original-hash backups verified, and a disabled row referencing a nonexistent `.pq` did not block deployment. Nine smoke-owned Excel PIDs exited. Rebuilt packaged GUI replacement and keyboard traversal also passed; existing Excel PID 13156 was untouched.
- This verifies the tested static definition-only batch path on one Office build. Mid-commit Excel failure, wider Office/workbook matrix, and physical sync/conflict cases remain open.

## W115 Excel explicit Protected View UI — 2026-10-22

- Added `tools/run_excel_protected_view_smoke.py`. It creates a blank disposable `.xlsx` in a `DispatchEx` Excel instance, explicitly opens it through `Application.ProtectedViewWindows.Open`, reads the Protected View UI through Windows UI Automation, closes without enabling editing, and verifies the workbook SHA-256 is unchanged.
- PASS on Windows 10 build 26300 / Excel 16.0 build 20430.0: Protected View collection count was 1; visible UIA title and banner identified Protected View; workbook hash stayed unchanged; owned Excel PID 22908 exited. Pre-existing Excel PID 13156 was untouched.
- Microsoft documents that `ProtectedViewWindows.Open` opens the named workbook in a new Protected View window: https://learn.microsoft.com/en-us/office/vba/api/excel.protectedviewwindows.open . This acceptance explicitly invokes that API. It does not prove automatic ZoneId=3 triggering through Explorer, actual Protected View policy configuration, message-center behavior, or workbook editing prevention across policy states. The app's separate Internet-zone preflight refusal remains independently verified.

## W116 fail-closed multi-workbook transaction reporting — 2026-10-22

- Three test-first cases reproduced false `applied` reports when an installer returned `Rolled back` or `Not attempted` results. Four more cases cover empty, missing, duplicate and unknown-status result sets. `deploy_multi_workbook` now marks a workbook failed unless exactly one `Created`, `Replaced` or `Skipped` result exists for every enabled query. Focused deployment suite: **15 passed**.
- Full `app/build.ps1`: **345 passed in 17.35s**; onedir GUI/CLI rebuild succeeded. The refreshed package's verifier passed. The packaged CLI `--help`, packaged GUI existing-query replacement, and 34-control keyboard traversal passed.
- Real two-workbook smoke on Windows 10 build 26300 / Excel 16.0 build 20430.0 passed both apply and repeat-no-op batches with exact query formula and backup-hash checks; the disabled missing-source profile entry remained ignored. Nine owned Excel processes exited; existing PID 13156 was untouched.
- The failure classifications use deterministic fake installer outcomes; mid-commit COM failure and multi-version Excel behavior remain open.

## W117 current Excel Internet Protected View policy — 2026-10-22

- Read-only query of `HKCU\Software\Microsoft\Office\16.0\Excel\Security\ProtectedView` returned `DisableInternetFilesInPV=1`, `DisableAttachmentsInPV=1`, and `DisableUnsafeLocationsInPV=1`; the corresponding `HKCU\Software\Policies\Microsoft\Office\16.0\Excel\Security\ProtectedView` key was absent. No Trust Center or registry setting was changed.
- Microsoft Support describes Protected View for Internet-originating files as a Trust Center setting and advises consulting an administrator before changing these controls: https://support.microsoft.com/en-us/office/collab-files/what-is-protected-view . As currently configured, this host cannot certify the enabled automatic Internet-file Protected View path. W115 separately certifies explicit ProtectedViewWindows.Open UI; the app's ZoneId=3 fail-closed check is independently covered by W88.

## W118 portable multi-workbook deployment release bundle — 2026-10-22

- Added bounded `.pqmdeploy`/ZIP creation from a versioned profile and exact `.m`/`.pq` bytes. The archive replaces the creator's workbook path with a placeholder, uses bundle-relative source paths, preserves disabled missing-source rows, and records SHA-256 per included source. Existing output is never overwritten. Safe extraction rejects tampered hashes and validates profile/manifest consistency. Enabled source-path mappings require preview plus typed approval before their mapped files are embedded; M-code environment mappings remain in the profile for deployment-time diff review.
- CLI `bundle` creates the release; `plan --profile release.pqmdeploy` and `deploy --profile release.pqmdeploy` consume it through the shared bounded extractor and remove temporary extraction after completion. Focused deployment-bundle, CLI, and package tests: **68 passed**; includes exact-byte round trip, hash-tamper refusal, no overwrite, mapping approval, dry-run, deploy, and extraction cleanup.
- `tools/run_excel_deployment_bundle_smoke.py` passed on Windows 10 build 26300 / Excel 16.0 build 20430.0: two disposable workbooks applied and read back the exact static formula; both original-hash backups verified; repeat deployment was a no-op; disabled missing source was ignored; nine owned Excel processes exited.
- This certifies static definition-only portable deployment on one Office build. Mid-commit COM failure and wider Office/workbook version coverage remain open.


## W119 packaged CLI portable deployment — 2026-10-22

- `tools/run_excel_deployment_bundle_smoke.py --cli-exe app/dist/PowerQueryManager/PowerQueryManagerCLI.exe` passed on Windows 10 build 26300 / Excel 16.0 build 20430.0. The packaged executable consumed the explicit `--confirm DEPLOY` phrase, deployed to two disposable workbooks, read back the exact static M formula and definition-only load state, verified both original-hash backups, then returned `no_changes` on repeat. A disabled missing-source row was ignored. All deployment/inspection Excel processes exited; pre-existing PID 13156 was untouched.
- This is a packaged CLI smoke for static definition-only deployment on one Office build, not a claim of cross-version or mid-commit failure coverage.


## W120 packaged CLI refresh — 2026-10-22

- `tools/run_excel_cli_refresh_smoke.py --cli-exe dist/PowerQueryManager/PowerQueryManagerCLI.exe` passed on Windows 10 build 26300 / Excel 16.0 build 20430.0. The packaged CLI accepted `--confirm REFRESH` and completed both background and synchronous refresh of a disposable connection-only query using static inline M. Both backups matched the pre-refresh file hashes; formula and load state reopened correctly; no external source was contacted; owned Excel processes exited and pre-existing PID 13156 was untouched.
- Static inline M on one Office build does not establish connector-backed refresh, alternate builds, or timeout behavior for genuinely blocked COM.


## W121 packaged CLI query-package export/import — 2026-10-22

- `tools/run_excel_cli_package_import_smoke.py` passed on Windows 10 build 26300 / Excel 16.0 build 20430.0. A disposable source workbook with three static queries was exported by the production package exporter and imported by the packaged CLI using `--confirm IMPORT`. Source bytes remained unchanged; the target reopened with exact names/formulas (including `Équipe – Résultats 2026 (Q1)`); its backup matched the original SHA-256; no external source was evaluated; owned Excel processes exited; PID 13156 was untouched.
- Static definition-only query-package path on one Office build only; alternate builds and broad failure scenarios remain unverified.


## W123 local PivotTable slicer preservation — 2026-10-22

- `python -B tools/run_excel_slicer_preserve_smoke.py` passed on Windows 10 build 26300 / Excel 16.0 build 20430.0. The disposable workbook contained a local range-backed PivotTable and named slicer with `Drop` deselected. Production staged definition-only import added a static query; read-only reopen verified exact query formula, PivotTable field, slicer name/caption, saved selected/excluded items, and filtered value 54. All smoke-owned Excel instances exited; Excel process inventory returned to baseline; no external source was evaluated.
- A preliminary combined test with the broader load-capability probe failed before import during probe cleanup (`com_error`); the focused definition-only preservation path passed. Do not infer slicer compatibility for load capability probing from this evidence. Timelines, external/Data Model PivotTables, other Office builds and screen interaction remain unverified.

## W125 — staged final-commit failure handling — 2026-10-22

- Portable regression injected a failure at the file transaction adapter's final commit, after the staged workbook save and verification boundary. For a two-query import, both changed rows were reported rolled back, the original workbook bytes stayed unchanged, and the stage was removed. Focused test: 1 passed; full build: 359 passed in 15.64s.
- The rebuilt packaged GUI replacement smoke independently passed on Excel 16.0 build 20430.0 with exact formula, backup hash, unrelated marker and owned-process checks. The failure injection itself was not performed inside live Excel/COM, so that matrix cell remains unverified.

## 2026-10-22 — W126 package comparison audit metadata

- The history writer's explicit metadata allowlist previously omitted the comparison count keys emitted by the package update preview UI. Added only the five fixed aggregate counts (added, modified, unchanged, removed, conflict); formula and source-path fields remain filtered. A regression verifies the exact saved record.
- Focused history tests: 34 passed. Full Windows build: 360 passed in 20.46s. PyInstaller rebuilt GUI/CLI. Packaged GUI replacement then passed on Excel 16.0 build 20430.0 with exact formula/backup/marker readback and process cleanup. Final package verified at 229 files / 135,738,985 bytes. This improves local audit completeness without recording query contents or file locations.


## W127 — 2026-10-22 — packaged GUI smoke cleanup race

- First rerun passed UI Automation but cleanup resolved the main-window HWND after the manager had already exited. The smoke harness now checks the owned app process before looking up the HWND and continues only if exit/readback verification can be completed.
- Rerun PASS on Windows 10 build 26300 / Excel 16.0 build 20430.0: exact formula, verified original-hash backup, unrelated marker, manager exit 0 and owned Excel cleanup.


## W128 — 2026-10-22 — audit metadata build acceptance

- Portable UI history regression checks that a mixed create/replace/skip result stores only outcome counts, action/status and workbook basename; query names and full workbook path are omitted. History/UI suite: 35 passed; full build: 361 passed in 19.43s.
- Rebuilt packaged GUI replacement passed on Windows 10 build 26300 / Excel 16.0 build 20430.0 with exact M readback, original-hash backup, unrelated marker and owned Excel cleanup. This import smoke verifies packaging/Excel behavior; history persistence is verified by the portable Qt test.


## W129 — 2026-10-22 — capability cleanup and packaged Data Model

- Environment: Windows 10 build 26300; Excel 16.0 build 20430.0; Python 3.11.9 x64; PySide6 6.8.3; pywin32 310; packaged one-folder GUI.
- Source matrix: connection-only, new worksheet, existing worksheet, and Data Model creation/repeat/changed-formula/reopen checks passed.
- Packaged Data Model path: PASS for capability probe, exact formula, connection `InModel` state, Data Model table membership and row count 2, verified original backup hash, manager exit code 0, and smoke-owned Excel process cleanup. The process set returned to only preexisting Excel PID 36464.
- No external source was refreshed or evaluated. One Office build and static inline M only; alternate builds, complex model relationships and external connectors remain unverified.


## W130 — 2026-10-22 — process-handle timeout cleanup

- Windows 10 build 26300; Python 3.11.9 x64; PySide6 6.8.3; pywin32 310; Excel 16.0 build 20430.0.
- Focused lifetime/capability/refresh suite: 32 passed. Full portable suite: 365 passed. `app/build.ps1`: 365 embedded tests passed and onedir GUI/CLI build succeeded.
- Packaged GUI Data Model smoke passed with read-only formula/model membership/row-count checks, verified backup, manager exit 0 and owned-process cleanup. Packaged CLI `--help` exited 0.
- The process-handle reuse guarantee is covered by mocked Win32 API tests. Live smoke covered normal cleanup only; forced termination of a blocked COM call and PID-to-handle capture race remain unverified.


## W131 — 2026-10-22 — rebuilt packaged GUI load matrix

- Windows 10 build 26300; Excel 16.0 build 20430.0; rebuilt one-folder GUI; disposable `.xlsx` workbooks; static inline M.
- PASS: connection-only (connection persisted; no worksheet/model table), new worksheet (table values 17/29), existing worksheet (`Output!C4`, values 17/29, unrelated marker preserved), and Data Model (connection `InModel`, model membership, row count 2).
- All four runs also verified the exact M formula after read-only reopen, pre-import backup SHA-256, manager exit code 0 and cleanup of every smoke-owned Excel process. Preexisting Excel PID 36464 remained untouched.
- Limit: one Office build, static inline M, and narrow disposable workbook fixtures. No external connector was evaluated.


## W132 — 2026-10-22 — main-window accessibility and layout regression

- Environment: Windows 10 build 26300; Python 3.11.9 x64; PySide6 6.8.3.
- Focused widget-tree regression verifies explicit accessible names on visible main-window interactive controls. It caught the Qt table-corner select-all button; the app now names it “Select all query source table cells”.
- UI suite: 60 passed at `QT_SCALE_FACTOR=1.25`; 60 passed at `QT_SCALE_FACTOR=1.5`. Windows onedir build and embedded suite: 366 passed. Packaged Windows UI Automation and disposable Excel definition-only import passed on Excel 16.0 build 20430.0.
- This is automated Qt/UIA coverage only; it does not replace screen-reader user testing, Windows display-scaling acceptance or physical laptop review.


## W133 — 2026-10-22 — packaged select-all UI Automation

- Windows 10 build 26300; PySide6 6.8.3; packaged one-folder manager.
- UIA initially omitted Qt's internal table-corner button, so the app disables it and provides a standard `Select all query rows` button; packaged UIA smoke now requires the button name and `Button` role and passed. Portable Qt test confirms it selects every row.
- Full build: 367 passed. UI suite: 61 passed at 125% and again at 150% scale. Packaged definition-only import passed on Excel 16.0 build 20430.0 with formula, backup, reopen, manager exit and owned process cleanup checks.
- Physical display scaling and screen-reader behavior remain manual acceptance.


## W134 — 2026-10-22 — Excel HWND/process ownership recheck

- Windows 10 build 26300; Python 3.11.9 x64; pywin32 310; Excel 16.0 build 20430.0.
- Fake Win32 regression: when the HWND owner differs from the captured PID after process-handle acquisition, close the handle and refuse the bounded COM call. Focused lifetime/capability/refresh/importer tests: 34 passed; full suite 368 passed. Build passed with 368 embedded tests.
- Rebuilt packaged GUI Data Model smoke passed with formula, model membership/row count 2, backup, app exit and owned-process cleanup. Existing Excel PID 36464 remained untouched.
- This does not simulate simultaneous HWND/PID reuse or forced termination of a genuinely blocked COM call; those remain unverified.


## W135 — 2026-10-22 — packaged CLI refresh after HWND recheck

- Windows 10 build 26300; Microsoft Excel 16.0 build 20430.0; packaged CLI; disposable `.xlsx` containing static inline M.
- PASS: background and synchronous refresh each reported `refreshed`; hash-verified backup per refresh; read-only formula/load-state verification; no external source; smoke-owned Excel process set returned to baseline.
- This exercises refresh using HWND-validated process handles. Simultaneous PID/HWND reuse and blocked-COM termination remain unverified; other Office builds remain open.


## W136 — 2026-10-22 — modal dialog controls and row selection

- Automated Qt regression covers Query Inspector, package update preview, profile mappings, history comparison and dependency graph: visible enabled interactive controls must expose accessible names. Profile mapping select-all buttons select all rows.
- Focused dialogs pass; full suite 370 passed. UI suite 63 passed at 125% and 150% Qt scale. Windows onedir build passed with 370 embedded tests. Packaged main-window UIA and disposable GUI import passed on Excel 16.0 build 20430.0.
- Modal dialog Windows UIA was not run for each dialog; screen-reader use, physical laptop review and OS-level display scaling remain unverified.


## W137 — 2026-10-22 — packaged modal accessibility verification

- Extended `tools/run_packaged_gui_import_smoke.py` to inspect the packaged query inspector and both profile-mapping categories through Windows UI Automation. It verifies names and roles for the M source/workbook/diff editors, search and actions; mapping tables, tabs, row actions and save/cancel buttons. Profile edits were cancelled.
- Packaged Windows UIA and definition-only import passed on Windows 10 build 26300 / Excel 16.0 build 20430.0. Disposable `.xlsx` formula readback, original backup, read-only reopen, manager exit 0 and owned Excel cleanup all passed. Pre-existing Excel PID 36464 was not touched.
- Scope: package update, history comparison and dependency graph are covered by Qt widget tests but were not inspected through packaged UIA. Screen-reader user testing, physical laptop review and OS-level display scaling remain open; v1.4.0 gate remains **UNVERIFIED**.


## W138 — 2026-10-22 — packaged dependency graph accessibility verification

- Packaged Windows UIA exposed the dependency diagram under the `Group` role and the named text summary as `Edit`; the action group and Close button were also present. The repeatable smoke passed with a disposable workbook on Excel 16.0 build 20430.0, including exact formula, backup and process-cleanup verification.
- Package update and history comparison remain unverified through packaged UIA. Screen-reader, physical laptop and OS-level scaling acceptance remain open.


## W139 — 2026-10-22 — packaged query-package preview accessibility

- Packaged Windows UIA verified the package update preview's summary, changes table, diff, safety notice and Close control; package files were generated locally. Disposable GUI import passed with exact formula, verified backup, read-only reopen and owned Excel cleanup on Excel 16.0 build 20430.0.
- The local history comparison dialog did not appear within 90 seconds in a disposable snapshot smoke. Its packaged UIA gate remains open; only pre-existing Excel PID 36464 remained after cleanup.


## W140 — 2026-10-22 — history worker UI-thread fix

- Failing-first QThread regressions exposed off-thread callbacks in comparison and restore. A QObject slot bridge now marshals history worker success/failure/completion to the GUI thread. Full suite 372 passed; UI suite 65 passed at both 125% and 150% Qt scale; Windows build passed 372 embedded tests.
- Packaged UIA passed for the history comparison table, formula diff and Close action; combined UIA also passed inspector, dependency graph, profile mapping and package update dialogs. Disposable Excel import passed on 16.0 build 20430.0 with backup, read-only formula verification, and owned-process cleanup.
- Other Office builds, screen-reader use, physical laptop review and OS-level scaling remain unverified.




## W141 — packaged local history restore acceptance

- Packaged GUI UI Automation selected a history snapshot for a disposable workbook, invoked Restore, confirmed the guarded overwrite, and verified the target SHA-256 returned to the snapshot hash. The UI confirmation exposed the generated backup path; the backup hash exactly matched the disposable workbook bytes immediately before restore.
- The same packaged run then completed definition-only import on a disposable `.xlsx` with exact formula readback, verified source backup, successful read-only reopen, manager exit 0, and cleanup of its own Excel process. Environment: Windows 10 build 26300, Excel 16.0 build 20430.0. Pre-existing Excel PID 36464 remained untouched.
- This closes the packaged UIA history-restore behavior on this host only. Other Office builds, screen-reader use, physical display acceptance, and the v1.4.0 release gate remain **UNVERIFIED**.


## W142 — packaged keyboard and history-store recovery

- `python -B tools/run_packaged_keyboard_smoke.py --exe dist/PowerQueryManager/PowerQueryManager.exe` passed. UI Automation verified the required main controls in keyboard order, a complete 35-control Tab cycle, no focus escape, no Excel launch, and exit success.
- `python -B tools/run_packaged_history_a11y_smoke.py --exe dist/PowerQueryManager/PowerQueryManager.exe --corrupt-history` passed. A synthetic corrupt SQLite file in isolated temporary `%LOCALAPPDATA%` was retained byte-for-byte in exactly one `history.corrupt-*` archive; a fresh database initialized; the UI exposed the recovery archive; the app exited 0.
- Scope is packaged UIA/keyboard behavior on this Windows host. Docs-only package refresh and verifier passed at 229 files / 135,751,359 bytes; contract and handoff checks passed. No screen-reader user evaluation or physical laptop review was performed; alternate Office builds and all remaining acceptance-matrix items stay open.


## W143 — query-bearing staged recovery for `.xlsx` and `.xlsm` containers

- `python -B tools/run_excel_recovery_smoke.py` and `python -B tools/run_excel_recovery_smoke.py --xlsm` passed on Windows 10 build 26300 / Excel 16.0 build 20430.0. Each disposable stage contained one query. Read-only reopen verified the exact M formula; original SHA-256 stayed unchanged; stage stayed preserved; output hash matched; two recovery-owned Excel instances and the seed instance exited.
- `.xlsm` was a macro-enabled container without a VBA project. VBA-bearing preservation, actual process interruption, and other Office builds remain unverified.


## W144 — worker-finish UI dispatch and rebuilt package

- Five remaining `QThread.finished` callbacks used plain lambdas to update busy UI state. They now connect to `MainWindow._on_background_thread_finished`, a typed Qt slot. QThread regression verifies callback thread identity is the Qt application thread. Focused test passed; full portable suite 373 passed; UI suite 66 passed at Qt scale 1.25 and 1.5; `app/build.ps1` passed 373 embedded tests.
- Rebuilt package verifier passed for 229 files / 135,751,591 bytes and CLI `--help` passed. Packaged keyboard smoke completed its 35-control cycle. Packaged GUI definition-only import passed on Windows 10 build 26300 / Excel 16.0 build 20430.0 with formula equality, verified backup, read-only reopen, manager exit 0 and owned Excel cleanup.
- Other Office builds, blocked COM termination, screen-reader use and physical display acceptance remain open.


## W145 — non-elevated packaged GUI import

- The packaged GUI import acceptance harness now rejects elevated sessions before creating any workbook or Excel process. In a `PCRN\Robert` session where `IsUserAnAdmin()` returned false, it explicitly reported non-elevated execution and passed packaged `.xlsx` definition-only import on Excel 16.0 build 20430.0: exact formula, verified backup, read-only reopen, manager exit 0, and owned Excel cleanup.
- Scope is this user profile and workspace ACL. Other standard-user/group-policy combinations remain unverified.


### W146 — 2026-10-23 — final package VBA import and refresh cleanup

- Host: Windows 10 build 26300; desktop Excel 16.0 build 20430.0; non-elevated launch; existing Excel PID 36464 recorded before each run and left untouched.
- `python -B tools/run_packaged_gui_import_smoke.py --exe dist/PowerQueryManager/PowerQueryManager.exe --vba` PASS on a disposable VBA-bearing `.xlsm`: exact M formula, verified pre-import backup, read-only reopen, VBA project and `Workbook_Open` source preserved, macro marker unchanged, manager exit 0, and no new Excel process. The installer-owned Excel PID shut down without forced termination.
- `python -B tools/run_excel_cli_refresh_smoke.py --cli-exe dist/PowerQueryManager/PowerQueryManagerCLI.exe` PASS for background and synchronous refresh of static inline M. Both verified backup hashes and readback; no external source was contacted; all owned Excel processes exited.
- Build/full test suite: **379 passed**. Package verifier: **229 files / 135,757,954 bytes**. Scope is one Windows/Office build, static inline M, and the tested disposable workbook shape; broader Excel, external-source, XLM, accessibility and display matrix cases remain open. Release gates remain **UNVERIFIED**.


## W147 — 2026-10-23 — packaged history comparison UIA rerun

- Command: `python -B tools/run_packaged_gui_import_smoke.py --exe dist/PowerQueryManager/PowerQueryManager.exe --history-dialog`.
- PASS on Windows 10 build 26300 / Excel 16.0 build 20430.0 under a non-elevated token. Packaged UI Automation opened local history, selected two seeded disposable snapshots, started read-only comparison, and exposed the named changes table, formula diff and close controls. The disposable GUI import portion also passed exact formula readback, verified backup, read-only reopen and manager exit 0.
- Seed and manager-owned Excel processes exited; the read-only verification process was cleaned up by its exact PID. Process set returned to the baseline containing only pre-existing PID 36464, which was untouched.
- This closes the packaged comparison-dialog UIA gap on this host only. Screen-reader, physical display, alternate Office builds and remaining Windows matrix coverage remain open; all release gates remain **UNVERIFIED**.


## W148 — 2026-10-23 — packaged CLI source-path and M-code mapping

- Command: `python -B tools/run_excel_cli_environment_mapping_smoke.py --packaged`. Windows 10 build 26300, desktop Excel 16.0 build 20430.0.
- First run exposed a smoke-harness defect: the child CLI received no typed mapping approvals and exited on `EOFError`. The harness now supplies the exact `APPLY PATH MAPPINGS`, `APPLY M-CODE MAPPINGS`, and `DEPLOY` phrases through repeated `--confirm` options and bounds the process to 240 seconds.
- Rerun PASS through both `python -B tools/run_excel_cli_environment_mapping_smoke.py` and its `--packaged` path. Both displayed the source-path preview and exact M diff, deployed two definition-only queries to a generated disposable workbook, and read-only Excel inspection verified the mapped formula and unchanged second formula. All four source files retained their original hashes. The mapping target is a nonexistent path embedded in static M; Excel did not evaluate it or contact external data.
- Packaged CLI returned exit 0; all smoke-owned Excel processes exited and pre-existing PID 36464 remained untouched. This validates the approval and mapping workflow on one build, not general M parsing or external-source refresh. Release gates remain **UNVERIFIED**.


## W149 — 2026-10-23 — closed stdin safely declines CLI approvals

- Added a failing-first regression for EOF while the CLI presents an M-code mapping approval. Before the fix, the uncaught `EOFError` was converted into a generic exit 5. CLI input now treats EOF as an empty response so existing approval guards return cancellation (exit 3) without starting Excel or changing source/target bytes. Focused CLI suite: **29 passed**; full Windows build/test pipeline: **380 passed** on Python 3.11.9 / Windows 10 build 26300.
- Rebuilt packaged GUI and CLI. Package verifier passed (**229 files**). Packaged CLI `--help` returned 0. The post-build packaged mapping integration smoke passed on Excel 16.0 build 20430.0 with exact approvals, mapped and unchanged formula readback, unchanged source hashes, and no owned process leak. This is the intended-approval path; the EOF path itself is covered by the unit regression.
- Excel process inventory returned only pre-existing PID 36464 after tests; it was untouched. Release gates remain **UNVERIFIED** for other Office builds and remaining matrix cases.


## W150 — 2026-10-23 — blocked-COM timeout probe characterization (inconclusive)

- On Windows 10 build 26300 / Excel 16.0 build 20430.0, a disposable `DispatchEx` instance received `Application.Wait(future_datetime)` with a target 20 seconds ahead. The COM call returned `True` in about 0.031 seconds instead of blocking, so it did not exercise timeout termination of a blocked COM server. No workbook was opened. The experimental change to the timeout smoke was reverted.
- The established `python -B tools/run_excel_timeout_cleanup_smoke.py` was rerun and passed its owned-process timeout cleanup check; it uses Python-side `sleep`, so it is not evidence for termination during a genuinely blocked COM call. The only remaining Excel process was pre-existing PID 36464.
- Genuine blocked-COM termination remains **UNVERIFIED**; do not treat this characterization as a pass for that acceptance item.


## W151 — 2026-10-23 — blocked Excel COM refresh timeout cleanup

- Command: `python -B tools/run_excel_blocked_com_timeout_smoke.py`. Windows 10 build 26300, Excel 16.0 build 20430.0, Python 3.11.9.
- PASS: a private `DispatchEx` Excel instance ran a synchronous `QueryTable.Refresh(False)` against an HTTP CSV fixture bound only to `127.0.0.1`. The local handler confirmed the request and deliberately withheld its response. The 3-second `call_with_timeout` terminated only the captured Excel process; its process handle signaled exit. Excel process IDs before/after were `[36464]`; pre-existing PID 36464 remained untouched. No workbook was opened or saved, and no external network was used.
- The harness was repeated after tightening setup/cleanup; it passed again with owned PID 10700. Process IDs again returned to `[36464]`.
- Scope: genuine blocked Excel COM `QueryTable.Refresh` call on one Office build. This is not a Power Query connection refresh test and does not certify other Office versions, alternative blocked COM methods, or all timeout/cleanup races.

## W152 — 2026-10-23 — Power Query refresh false-success acceptance failure

- Command: `python -B tools/run_excel_pq_blocked_refresh_timeout_smoke.py` against a disposable `.xlsx` and loopback-only delayed `Web.Contents` source. Windows 10 build 26300 / Excel 16.0 build 20430.0 / Python 3.11.9.
- FAIL: packaged CLI exited 0 in 0.094 seconds, reported `refreshed` and staged workbook commit, while the source received no request. No external network was used. The harness removed its temporary folder; Excel inventory returned to pre-existing PID 36464 only.
- Direct follow-up: `WorkbookConnection.Refresh()` returned `None` promptly; background `Refreshing` was false during a delayed source request; `AfterCalculate` did not signal completion in the observation window; Power Query `RefreshDate` raised a COM error. Another run did not contact the fixture. Diagnostic-owned Excel processes exited; PID 36464 was untouched.
- Gate at the time of W152: Power Query refresh completion, timeout, and no-false-success behavior **FAIL / UNVERIFIED** on this build. W154/W155 and W157 later superseded this finding for the tested connection-only and existing worksheet-loaded loopback cases on this same Excel build; Data Model refresh, other Office builds and external providers remain unverified. This is separate from W151, which validates a blocked legacy `QueryTable.Refresh(False)` call.

## W153 — 2026-10-23 — recovered packaged GUI import smoke

- Command: `python -B tools/run_packaged_gui_import_smoke.py --exe dist/PowerQueryManager/PowerQueryManager.exe`. The harness ran under a non-elevated token on Windows 10 build 26300 / Excel 16.0 build 20430.0.
- PASS: packaged UI Automation found named main-window, query-grid, filter, replace, backup, import, and operation-log controls; inspector, dependency graph, and profile mapping dialogs exposed named controls. Mapping changes were cancelled. A disposable `.xlsx` definition-only import preserved exact M formula, verified backup hash, read-only reopen, manager exit 0, and graceful owned Excel exit. Only pre-existing PID 36464 remained.
- This verifies the recovered packaged GUI/import path on this host. It does not resolve the independent W152 Power Query refresh false-success finding or complete the release matrix.

## W154 — 2026-10-23 — successful packaged loopback Power Query refresh

- Command: `python -B tools/run_excel_pq_refresh_smoke.py` against the published `dist` package. Windows 10 build 26300 / Excel 16.0 build 20430.0 / Python 3.11.9. The same test also passed against `app/dist` before publication.
- PASS in background and synchronous modes against a local `127.0.0.1` CSV source. Three expected local GET requests were observed; CLI results were `refreshed`; each backup matched its immediate pre-refresh workbook hash; read-only reopen verified exact M text, one original query connection, and original worksheet count. Temporary tables/connections were removed. Only pre-existing Excel PID 36464 remained; no external network was used.
- This fixes the false-success path for the tested connection-only shape. See W157 for the additional worksheet-loaded case; Data Model and other Office builds remain open.

## W155 — 2026-10-23 — packaged Power Query timeout and rollback

- Command: `python -B tools/run_excel_pq_blocked_refresh_timeout_smoke.py` against the published `dist` package. Same Windows/Excel host as W154; disposable workbook and delayed loopback-only source. The same test also passed against `app/dist` before publication.
- PASS: 3-second synchronous refresh timeout returned CLI exit 5 and an explicit timeout. The source request was received; original workbook SHA-256 stayed unchanged; exactly one backup matched the original hash; no stage remained; Excel PID inventory returned to `[36464]` after cleanup of the captured app-owned process.
- W154/W155 resolve W152 for the tested connection-only loopback path; W157/W158 later add worksheet-loaded and selected Data Model refresh on the same build. Alternate Office builds, provider prompts, and external connectors remain unverified.

## W156 — 2026-10-23 — rebuilt published GUI import acceptance

- Command: `python -B tools/run_packaged_gui_import_smoke.py --exe dist/PowerQueryManager/PowerQueryManager.exe` after publishing the refreshed GUI/CLI build. Non-elevated Windows 10 build 26300 / Excel 16.0 build 20430.0.
- PASS: main manager, inspector, dependency graph, and profile mapping dialogs exposed named UIA controls; mapping edits were cancelled. Disposable GUI definition-only import passed exact M formula readback, verified backup, read-only reopen and manager exit 0. Manager-owned Excel exited gracefully; verifier-owned Excel PID was cleaned; process inventory returned to pre-existing PID 36464.
- This verifies this build's import/UI path only. Screen-reader, physical display and other Office builds remain open.

## W157 — 2026-10-23 — published package refreshes worksheet-loaded queries

- Command: `python -B tools/run_excel_pq_refresh_smoke.py` against the published `dist` package. Windows 10 build 26300 / Excel 16.0 build 20430.0.
- PASS for both connection-only and existing worksheet-loaded query modes, each in background and synchronous mode. The worksheet table at `D4` refreshed from seeded value `1` to local loopback result `42`; its table location, query connection, M formula and worksheet count were retained. Each transaction's backup matched its pre-refresh hash. Connection-only and worksheet-loaded refresh requests reached only the local CSV server (nine GETs total); only pre-existing Excel PID 36464 remained.
- At W157, connection-only and worksheet-loaded refresh checks passed on this one host/build; Data Model refresh was still refused then. W158 below adds the scoped Data Model acceptance. Alternate Office builds and external providers remain unverified.

## W158 — 2026-10-23 — packaged selected Data Model table refresh

- Command: `python -B tools/run_excel_data_model_refresh_smoke.py` against published `dist`. Disposable `.xlsx` with two Data Model queries; both source URLs bound only to loopback; Windows 10 build 26300 / Excel 16.0 build 20430.0.
- PASS: background-requested and synchronous-requested `--all` packaged CLI runs refreshed both model tables. Both loopback endpoints were contacted; both tables reopened read-only with two rows, exact M/connection associations, and unchanged worksheet count. Each run's backup matched its pre-refresh workbook, stage was removed, and CLI reported both selected queries refreshed.
- PASS: a separate one-query selection from that same two-table model was refused during preflight with the original workbook hash unchanged, no new loopback request, and no retained stage. Excel process inventory returned to `[36464]` before and after; the pre-existing user-owned PID was untouched. Excel's Data Model APIs are synchronous and have no background option; the operation report says so. Other Office builds, external providers, authentication/privacy prompts, model relationships and remaining Windows acceptance remain unverified.


### W169 — 2026-10-23 — rebuilt package smoke after history transaction fix

- Package: `dist/PowerQueryManager/PowerQueryManager.exe` rebuilt by `app/build.ps1`; compact package verification passed at 229 files / 135,784,108 bytes. Packaged CLI `--help`, contract validation and 108-reference feature traceability passed.
- Host: Windows 10 build 26300, Python 3.11.9, desktop Excel 16.0 build 20430.0; packaged smoke ran non-elevated with disposable `.xlsx`.
- PASS: main manager and inspector/dependency/profile mapping UI Automation controls exposed; definition-only import preserved exact M formula; verified backup and read-only reopen passed; manager exited 0; smoke-owned Excel process exited. Only pre-existing Excel PID 36464 remained and was not touched.
- Scope: this smoke verifies the rebuilt GUI/import path; W169 hard-exit/SQLite rollback is separately covered by a portable subprocess regression. Alternate Office builds, other workbook classes and remaining QA matrix items stay unverified.


### W170 — rebuilt GUI package smoke after exclusive snapshot publication

- `app/build.ps1`: 401 portable tests passed; one-folder GUI/CLI rebuilt. Compact package verification: 229 files / 135,784,395 bytes. Contract validation and 109 traceability references passed; packaged CLI `--help` exited 0.
- Rebuilt GUI smoke on Windows 10 build 26300 / Excel 16.0 build 20430.0 under non-elevated token: UI Automation controls passed; disposable `.xlsx` definition-only import preserved exact M formula; verified backup, read-only reopen, manager exit 0 and owned-process cleanup passed. Only pre-existing Excel PID 36464 remained.
- Snapshot exclusive-collision and hard-exit link-window checks are portable filesystem/SQLite tests and did not interact with Excel. Alternate Office builds and remaining workbook/accessibility/security matrix items remain unverified.


### W171 — clean onedir build reproducibility

- Two consecutive clean builds using `app/build.ps1`, Windows 10 build 26300, Python 3.11.9 and PyInstaller 6.13.0 each passed 401 tests. After fixing `PYTHONHASHSEED=0` and `SOURCE_DATE_EPOCH=1704067200` for PyInstaller, all **227** payload file paths and SHA-256 hashes matched exactly (0 missing, extra or changed).
- This verifies repeatability on the same host/toolchain. Cross-host, alternate Python patch version, and alternate PyInstaller build reproducibility are not claimed.


### W172 — rebuilt package after concurrent snapshot recovery fix

- The concurrent snapshot/startup barrier regression passes; full suite: **402 passed**. The package verifier passed at 229 files / 135,784,626 bytes; contracts, 110 test-traceability references and packaged CLI help passed.
- Packaged GUI smoke on Windows 10 build 26300 / Excel 16.0 build 20430.0 under a non-elevated token passed UI Automation and disposable definition-only import: exact formula, backup, read-only reopen, manager exit 0 and owned-process cleanup. Only pre-existing Excel PID 36464 remained.
- Concurrency regression is portable SQLite/filesystem coverage; Excel was not involved in its interleaving test. Other Office builds and remaining Excel matrix items remain unverified.


### W173 — history database reparse-point rejection

- On Windows, actual file and directory symlinks were created in disposable temp folders. `HistoryStore` refused each before following/creating through the path; the target SQLite bytes remained unchanged and the target directory had no new database. The configured LocalAppData history path had no reparse-point component. Focused history suite: **44 passed**; full suite: **404 passed**. Rebuilt package verification passed at 229 files / 135,784,777 bytes; contracts, 112 traceability references and packaged CLI help passed. The same package passed GUI UI Automation and disposable definition-only import on Excel 16.0 build 20430.0; exact formula, backup, read-only reopen and owned-process cleanup passed, leaving only pre-existing PID 36464.
- The symlink regressions are local filesystem/database coverage; no workbook or Excel process was opened for those tests. Other reparse-point types and concurrent path replacement remain unverified.


### W174 — oversized compressed query ZIP refusal

- Portable regression adds valid ZIP data plus trailing padding to exceed a supplied 1,024-byte cap. Extraction rejects it before ZIP parsing and creates no `pq-import-*` staging directory. Focused manager suite: **35 passed**; full suite: **405 passed**. Package verification passed at 229 files / 135,785,193 bytes; contracts, 113 traceability references and CLI help passed. Packaged GUI smoke also passed on Excel 16.0 build 20430.0 with UI Automation and disposable definition-only import, exact formula, verified backup, read-only reopen and owned-process cleanup.
- Existing path traversal, symlink, duplicate, Windows reserved name, expanded-size and compression-ratio cases remain covered. The archive-limit regression itself did not interact with Excel.

### W175 — ZIP archive handle consistency

- The compressed-size check and ZIP parsing now use the same open file descriptor; focused regression plus deployment bundle suite: **44 passed** and full suite/build: **406 passed**. This is portable filesystem coverage. The regression simulates parser-boundary path substitution because Windows disallows renaming the normal open file handle used by the test. Rebuilt packaged GUI smoke passed on Windows 10 build 26300 / Excel 16.0 build 20430.0 under a non-elevated token. UI Automation, disposable definition-only import, exact formula, verified backup, read-only reopen, manager exit 0, and cleanup of smoke-owned Excel processes passed; pre-existing PID 36464 remained. Screen-reader, physical display, alternate Office builds and full Excel matrix remain unverified.

### W176 — current-package skip / replace / mixed-source import

- Re-ran current compact-package GUI acceptance on Windows 10 build 26300 / Excel 16.0 build 20430.0, non-elevated. All-skip preserved exact workbook hash/timestamp with no backup or stage. Replace preserved an unrelated worksheet marker and verified exact new M plus original-hash backup. Three numbered `.m`/`.pq` files matched review and Excel query order with exact formula readback. Manager exit was 0; harness-owned Excel processes were cleaned, leaving only pre-existing PID 36464. Other Office builds and remaining matrix items are still unverified.

### W177 — current-package load destination acceptance

- Re-ran connection-only, new worksheet, existing worksheet, and Data Model flows through the compact GUI package on Windows 10 build 26300 / Excel 16.0 build 20430.0, non-elevated. Connection-only had no sheet/model; new-sheet table values were 17/29; existing-sheet table anchored at Output!C4 with unrelated marker preserved; Data Model membership and row count 2 were verified. All modes verified exact M and backups. Owned process cleanup passed; pre-existing PID 36464 was untouched. Tests use static inline M; external providers and other Office builds remain unverified.

### W178 — packaged CLI background and synchronous refresh

- Current packaged CLI refreshed a disposable inline-static-M query in both background and synchronous modes on Excel 16.0 build 20430.0. Both statuses were `refreshed`, backups matched, no external source was used, and owned Excel processes exited. External providers and other Office builds remain unverified.

### W180 — ordinary objects and related report preservation

- `tools/run_excel_preserve_smoke.py` passed on Excel 16.0 build 20430.0. It verified ordinary table name/range/rows/style/filter; defined names, cell formulas, calculation mode, chart, shape and PivotTable values; query-backed report formulas and name; distinct connection-only/worksheet load mappings; and privacy metadata. Owned Excel processes exited. Other object types/builds remain unverified.

### W181 — staged verification close failure

- Portable fake-COM regression injects an error closing the staged workbook's read-only verification view. The importer must abort before commit, preserve original bytes, discard the stage, and report `ROLLED_BACK`. Focused importer suite: 42 passed; full Windows build: 407 passed; compact package verifier: 229 files / 135,788,307 bytes. Rebuilt packaged GUI import passed on Excel 16.0 build 20430.0 with exact formula, backup, read-only reopen and process cleanup. Real Office injection of a close failure remains unverified.

### W182 — rebuilt package replacement after W181

- Current rebuilt package existing-query replacement passed on Excel 16.0 build 20430.0, non-elevated: exact new formula, verified pre-import backup, unrelated marker preservation, manager exit 0 and owned-process cleanup. This validates the normal commit path only; real verification-close failure injection and other Office builds remain unverified.

### W183/W184 — UIA harness recovery and VBA acceptance

- W183 adds two tests proving the smoke harness treats unknown UIA control types as an unavailable optional log, not an import failure; full build: 409 passed and compact package: 229 files / 135,788,815 bytes. W184 reran the current packaged GUI with a disposable harmless VBA `Workbook_Open`: VBA source survived, macro marker did not change, exact M and backup matched, manager exited 0 and owned Excel processes were cleaned. Windows 10 build 26300 / Excel 16.0 build 20430.0, non-elevated. Other macro forms and Office builds remain unverified.

### W185 — worksheet name inventory failure

- Portable fake-COM test verifies a failed worksheet name read aborts before adding a destination sheet. Importer suite: **44 passed**. Full build: 410 passed; package verifier: 229 files / 135,790,075 bytes. Current packaged new-worksheet and existing-worksheet modes passed on Excel 16.0 build 20430.0 with exact formulas, values 17/29, Output!C4 anchor, backup, unrelated marker, read-only reopen and process cleanup. Real Office collection-error injection remains unverified.

### W186 — keyboard navigation on current package

- Current compact package completed a 35-control Tab cycle through the manager; import controls were reached in order and focus remained in the window. No workbook or Excel was opened. Pywinauto printed its STA-thread warning and reset to STA; exit 0. Screen-reader and physical OS scaling acceptance remain open.

### W187 — active workbook detection without owner-lock file

- W22 recheck passed both subcases on Windows 10 build 26300 / Excel 16.0 build 20430.0. Restart Manager detected a separate Python process holding a disposable `.xlsx` with no `~$` file; write preflight refused. Real Excel read-only opened a valid disposable workbook without an owner-lock file; the exact-path ROT check detected the workbook, preflight refused, hash was unchanged and owned Excel process exited. Pre-existing PID 36464 remained untouched. OneDrive and mutation-path tests remain separate/open.

### W188 — application preflight against lock-file-free read-only Excel owner

- `python -B tools/run_excel_open_without_lock_smoke.py` passed on Windows 10 build 26300 / Excel 16.0 build 20430.0. With a disposable workbook read-only open in Excel and no owner-lock file, packaged-service import, refresh and overwrite-restore were refused before staging or backup. Exact source hash was preserved; only the harness-owned Excel process exited.
- This adds mutation-path evidence to W187 detection results. OneDrive behavior, other Office builds and the remaining matrix are still open.

### W190 — repeated launch/import process cleanup

- Three sequential packaged GUI acceptance runs each imported a query into an isolated disposable workbook, verified exact M, matching backup and read-only reopen, and exited 0. After each run, the Excel process set returned to the pre-existing PID 36464; no unrelated Excel process was touched.
- UI Automation did not expose the optional operation-log text on two runs, so repeatable operation-log accessibility remains unverified. Windows 10 build 26300 / Excel 16.0 build 20430.0 only.

### W192 — keyboard navigation on the current package

- Packaged UI Automation completed a 35-control Tab cycle without focus escape; workbook path, Add files, query grid, replace policy, backup and import were reachable in order. No workbook or Excel was opened. Exit 0; pywinauto emitted its STA reset warning.
- This does not replace screen-reader user testing or physical OS display-scaling review.

### W193 — final packaged build

- The current published one-folder package was rebuilt from the current source after the 413-test suite passed. Package: 229 files / 135,793,203 bytes; package verifier, contracts, 118 feature references and CLI help passed.
- Strict UIA operation-log lookup on the published package found the final completion entry. Definition-only import verified exact M, original-hash backup, read-only reopen, app exit 0 and cleanup of its owned Excel process. Pre-existing PID 36464 remained untouched.
- Other Office builds, screen-reader and physical display cases remain open.

### W194 — stage cleanup reporting

- Cleanup now confirms deletion by observing `FileNotFoundError`; permission, sharing or stat errors cannot be interpreted as “gone.” A failing-first regression verifies the retained stage is named in the operation report. Focused tests: **15 passed**; current suite: **414 passed**.
- The current package was rebuilt and verified, and packaged GUI import/strict operation-log UIA passed on Excel 16.0 build 20430.0. The regression simulates a filesystem sharing/access denial; real Excel-induced deletion failure and other Office builds remain unverified.

### W195 — Windows stage cleanup under denied delete sharing

- A disposable `CreateFileW` handle omitted `FILE_SHARE_DELETE`; stage cleanup returned false and named the retained candidate without changing its bytes. Closing the owned test handle allowed a subsequent cleanup to remove the file. Focused test: **1 passed**; full suite: **415 passed**.
- This is a real Windows sharing violation but not one created by Excel. Other Office builds and the rest of the Windows matrix remain open.

### W196 — editor contrast across light and dark themes

- M syntax colors, diff colors, line-number gutter and current-line highlight now track the manager theme. Automated contrast assertions require **4.5:1** or higher for token, diff and gutter text in light and dark mode; nested inspector editors are tested while the parent theme toggles.
- 70 UI tests passed at each Qt scale factor (125%, 150%). Full app suite passed **418**. Packaged GUI import passed on Excel 16.0 build 20430.0.
- Screen-reader evaluation and physical OS display scaling remain manual acceptance items.

### W197 — UTF-8 BOM/CRLF M source create and replacement

- The COM probe on Windows 10 build 26300 / Excel 16.0 build 20430.0 observed that `Queries.Add` preserves internal CRLF but removes terminal CR/LF after add/save/read-only reopen. The importer and readback verifier now accept only this terminal line-ending normalization while requiring exact internal content.
- The published onedir GUI passed both create and existing-query replacement with BOM-prefixed CRLF source on separate disposable `.xlsx` workbooks. Each verified M readback under the narrow rule, source-byte SHA-256, pre-import backup, read-only reopen, manager exit 0, and owned Excel cleanup; create also verified the final operation log through UI Automation. The pre-existing Excel PID 36464 remained untouched.
- These checks cover one static definition-only query on one Office build. Alternate Office builds, external data, other load modes for BOM/CRLF sources, and wider workbook variants remain open.

### W198 — existing-target atomic staged replacement

- A failing-first Windows regression demonstrated that `os.replace` can create the target path if a concurrent actor removes the selected workbook at the final boundary. The adapter now uses `ReplaceFileW`, which requires an existing target, keeps a same-volume rollback copy, verifies the original and staged hashes, and attempts recovery on a partial failure. A disappearing-target race test leaves the path absent and the stage intact rather than publishing into the new path.
- On Windows, the NTFS named-stream test confirms `Zone.Identifier` is retained from the selected workbook even if Excel's staged copy omits it. The rebuilt GUI existing-query replacement passed on Excel 16.0 build 20430.0 with exact M readback, verified backup, unrelated marker, manager exit 0, and owned Excel cleanup.
- The exact published GUI package also passed existing-query replacement from a BOM-prefixed CRLF `.m` source; it verified the unchanged source hash and Excel formula readback under W197's terminal-newline rule.
- This verifies the normal replacement path and a deterministic target-disappearance boundary on one Windows host. W200 simulates API-reported partial replacement and recovery denial at the injected Win32 boundary; other filesystem/policy combinations and OS-generated partial errors remain unverified.

### W199 — repeatable onedir payload after W198

- A second clean `app/build.ps1` run passed **422 tests** on Windows 10 build 26300 / Python 3.11.9 / PyInstaller 6.13.0 / PySide6 6.8.3. Its published runtime payload compared against W198: **227 files, no missing/extra files, and zero hash differences**.
- Same-host/toolchain repeatability only; other Windows environments and Office builds remain unverified.

### W200 — partial ReplaceFile recovery states

- Injected API-boundary tests simulate Win32 error 1177 after the original has moved to the rollback path. Successful recovery uses real `MoveFileW` and verifies original hash, stage preservation and no remaining rollback file. A second test injects restore denial and verifies the original-hash recovery copy remains named and retained with the stage.
- This does not induce a native `ReplaceFileW` partial failure. The freshly rebuilt GUI separately passed BOM/CRLF existing-query replacement on Excel 16.0 build 20430.0 with source hash, formula readback, backup, exit and process cleanup.

### W201 — package launch from a read/execute-only directory

- `tools/run_packaged_readonly_install_smoke.ps1` copies the published onedir package to a disposable temp folder, applies a current-user RX ACL, verifies that writing is denied, and launches the packaged import acceptance from that location. The script cleans its disposable tree in `finally`.
- Pass: Windows 10 build 26300, Excel 16.0 build 20430.0, non-elevated. The packaged manager imported a definition-only query; exact M, verified backup, read-only reopen, UI Automation, exit 0 and owned-process cleanup passed. Existing Excel PID 36464 remained untouched.
- One local ACL/profile/build only. Alternate ACL inheritance, OneDrive redirection, endpoint policy, screen-reader/display and other Office versions remain unverified.

## W203 — history-restore commit safety

- Snapshot overwrite and its verification rollback now use the shared Windows `ReplaceFileW` helper. Failing-first tests confirm a target deleted at either commit boundary is not recreated and the verified pre-restore backup is retained.
- Full build: **426 passed**. Packaged history restore passed on Excel 16.0 build 20430.0 with exact restored snapshot hash, exact timestamped backup, UI Automation and owned-process cleanup. PID 36464 was untouched.
- Native partial-error generation, cloud-sync races, other Office builds and remaining accessibility/workbook matrix cases remain unverified.

## W206 — packaged scale-factor acceptance

- The current published GUI passed UI Automation and disposable import at Qt scale factors 1.25 and 1.5. Exact M, backup, read-only reopen, exit 0 and owned Excel cleanup passed on Excel 16.0 build 20430.0; pre-existing PID 36464 stayed untouched.
- Screen-reader, physical OS display, and other Office builds remain unverified.

### W209 — all v1.5 load modes rechecked on the rebuilt package

- The current package passed real Excel GUI acceptance for definition-only (W203), new worksheet, existing worksheet at `Output!C4`, connection-only, and Data Model. Worksheet values 17/29, connection/sheet/model distinctions, exact formulas, backup hashes, read-only reopen, exit 0 and owned-process cleanup passed on Excel 16.0 build 20430.0.
- This is one host/build with built-in static M. Other Office versions, external providers and workbook variants remain unverified; v1.5.0 remains **UNVERIFIED**.

## W210 — rebuilt package replace, skip and mixed legacy import

- Replace passed with exact formula readback, original-hash backup, unrelated marker preservation and owned Excel cleanup. Preserve-existing replacement retained an existing connection-only load. Skip preserved exact workbook hash, timestamp and formula, and created no backup or stage.
- Mixed numbered `.m`/`.pq` sources passed through a versioned profile serialized in reverse order; numeric-prefix ordering, Excel query order, exact formulas and backup passed. Excel 16.0 build 20430.0 only; PID 36464 remained untouched.

## W211 — current-package refresh verification

- Loopback-only packaged Power Query refresh passed in background and synchronous modes for connection-only and worksheet loads; worksheet value updated to 42 and four exact backups were verified. Six HTTP requests were observed.
- The 3-second Power Query timeout correctly returned exit 5, preserved original hash and backup, removed the stage and cleaned its owned Excel process. Data Model refresh in both requested modes returned two rows for selected and helper tables with verified backups.
- Excel 16.0 build 20430.0 only, localhost only; no external network. Pre-existing PID 36464 survived all runs.

## W213 — rebuilt `.xlsm` macro safety

- Current package imported into a disposable VBA-bearing `.xlsm`. The harmless `Workbook_Open` marker did not execute, `vbaProject.bin` remained intact, exact formula/backup/reopen passed, app exited 0 and owned Excel processes exited on Excel 16.0 build 20430.0. Existing PID 36464 was untouched.
- XLM, other VBA forms and other Office builds remain unverified.

## W214 — current packaged deployment-bundle acceptance

- Packaged CLI built and applied a versioned `.pqmdeploy` bundle to two disposable workbooks. Both first-run statuses were `applied`, both formulas/load states and original-hash backups verified; repeated deploy returned `no_changes` for both, and the disabled missing source was ignored. Three owned Excel PIDs exited.
- Pass used stable baseline `[36464]` on Excel 16.0 build 20430.0. An earlier run stopped because its process guard observed PID 2200 disappear; ownership/cause is unknown and the run is inconclusive. The passing retry preserved PID 36464.
- Mid-batch failure isolation, sync conflicts, other Office builds and larger deployments remain open.

## W215 — deployment read-only target isolation

- Mixed two-workbook deployment passed: writable target `applied` with exact original-hash backup; read-only target `blocked_preflight` before backup and remained byte-identical. Both owned Excel PIDs exited; pre-existing PID 36464 stayed.
- Excel 16.0 build 20430.0 only; mid-commit partial batch failure and sync conflicts remain unverified.

## W216/W217 — current packaged CLI mapping and query-package checks

- Packaged source-root and M-literal mapping showed exact previews, required typed approvals, changed only the approved definition, and left original sources unchanged. Both queries read back correctly; no external path was evaluated.
- Packaged CLI query-package export/import passed with three queries, exact formula/name readback including Unicode punctuation, unchanged source workbook hash, verified target backup, and owned Excel cleanup. Excel 16.0 build 20430.0 only.

## W221 — packaged deployment and failure-state reporting

- Portable regression coverage reproduces both changed-target and unchanged-target outcomes after an installer exception. Reports remain `failed`, include the observed current target hash, and identify when a verified pre-deployment backup is available. Focused deployment suite: **16 passed**; full Windows build: **427 passed**.
- The rebuilt packaged CLI applied a disposable `.pqmdeploy` bundle to two workbooks (`applied`, `applied`), then repeated as (`no_changes`, `no_changes`) with exact formula/load readback and verified backups. The rebuilt packaged GUI import also passed on Windows 10 build 26300 / Excel 16.0 build 20430.0. All smoke-owned Excel processes exited; PID 36464 remained.
- Native Excel partial commit errors, cross-workbook atomicity, cloud-sync conflicts, other Office versions, screen-reader use and physical display acceptance remain **UNVERIFIED**.

## W223 — rebuilt-package parameter edit and import

- Packaged GUI directly edited the declared Text parameter from the test source, saved a separate copy, retained the original source hash, and imported the copy into a disposable workbook. Exact formula, verified backup, read-only reopen, unrelated marker, exit and owned-process cleanup passed on Excel 16.0 build 20430.0. Existing Excel PID 36464 was untouched.
- This is one Office build only; other builds and remaining accessibility/security/workbook cases remain unverified.

## W224 — vanished-target deployment failure

- Fake-installer regression removes the disposable workbook path after the verified backup and raises. The app reports failure and unverified final state, keeps the backup, and does not recreate the path. Deployment suite: **17 passed**; portable suite: **428 passed**.
- Real Excel deletion/partial commit is not exercised. The current binary is still W221, verified on Excel 16.0 build 20430.0; remaining matrix items are unverified.

## W225 — dependency scanning follows M line boundaries

- Microsoft M line-comment syntax ends at CR, CRLF, LF, NEL, U+2028, or U+2029. The dependency scanner now recognizes these boundaries; failing-first unit tests cover each form and a dependency-graph UI regression verifies the resulting edge/summary. It remains advisory and does not claim full M grammar parsing.
- Portable full suite/build and current package verification are recorded in W226. Excel-specific release gates remain unverified.

## W227 — Unicode regular identifiers in the dependency scanner

- Fixed four false negatives from M regular identifiers containing Unicode combining, connector and formatting characters or starting with Unicode category Nl. The tokenizer follows the documented Unicode categories; failing-first tests pass, along with the dependency graph UI tests.
- Scanner remains advisory and no M source was evaluated. The full rebuilt package and remaining gates are recorded in W228.

## W228 — rebuilt Windows package after dependency scanner fixes

- Windows 10 build 26300 / Python 3.11.9 / PyInstaller 6.13.0: **439 tests passed**. Package staging/verifier tests: **8 passed**; contracts and 123 traceability references pass. The compact package verifies at 229 files / 135,809,314 bytes.
- Packaged GUI import/readback passed on Excel 16.0 build 20430.0. Exact M, verified backup, read-only reopen, exit and owned Excel cleanup passed; pre-existing PID 36464 remained.
- Dependency analysis handles the six specified M line endings and the documented Unicode categories, but remains advisory. Other Office versions, full M grammar, providers, accessibility and security matrix remain unverified.

## W229 — quoted identifier escape handling

- Failing-first scanner tests exposed missed references in M quoted identifiers containing control/Unicode escapes, escaped `#`, or doubled quotes, and an invalid Unicode escape that was accepted as a literal. The scanner now decodes valid escapes and marks invalid escapes ambiguous. Focused scanner tests: 23 passed; graph dialog tests: 5 passed.
- Full suite: 445 passed; Windows build succeeded on Windows 10 build 26300 / Python 3.11.9 / PyInstaller 6.13.0. Package staging/verifier/smoke-support tests: 13 passed; contracts and 125 test traceability references pass. The compact package verifies at 229 files / 135,812,357 bytes, with 227/227 runtime hashes matching `app/dist`; packaged CLI help exits 0.
- Packaged GUI import/readback passed on Excel 16.0 build 20430.0 with UI Automation, exact formula, verified backup, read-only reopen, app exit and owned-process cleanup. PID 36464 was present before and after the smoke. The scanner remains advisory and does not claim full M-language parsing. Other Office builds and broader accessibility/security/workbook matrix gates remain **UNVERIFIED**.

## W230 — dotted M identifier matching

- Microsoft’s grammar allows dot-separated components in a regular identifier. A failing-first regression showed that `Base.Field` could be reported as a reference to `Base`. The scanner now compares the full dotted name; regressions also cover exact dotted query names, record field labels, quoted identifiers, and dynamic `Record.Field` ambiguity. Focused dependency and graph tests: 32 passed.
- Full app suite: 449 passed; Windows 10 build 26300 / Python 3.11.9 / PyInstaller 6.13.0 build succeeded. Package staging/verifier/smoke-support tests: 13 passed; contracts and 129 traceability functions pass. Compact package: 229 files / 135,813,787 bytes; manifest SHA-256 `EB101082C6A45AACB41127C06A17C374ECE6A5B5E753285A42BBE86AF4B9FE6F`; all 227 runtime files match `app/dist`; CLI help exits 0.
- Rebuilt packaged GUI import/readback passed on Excel 16.0 build 20430.0 with UI Automation, exact formula, verified backup, read-only reopen, app exit, and owned-process cleanup. PID 36464 was present before and after the smoke. The scanner does not claim full M grammar support; other Office builds and broader accessibility/security/workbook gates remain **UNVERIFIED**.

## W231 — hash-prefixed M keywords

- Failing-first scanner tests exposed false query edges for reserved `#` forms such as `#date` and `#table`; `#shared` and `#sections` are now flagged ambiguous, and unknown hash-prefixed identifiers fail closed. Focused dependency and graph tests: 44 passed.
- Full app suite: 461 passed; Windows build succeeded on Windows 10 build 26300 / Python 3.11.9 / PyInstaller 6.13.0. Package staging/verifier/smoke-support tests: 13 passed; contracts and 132 traceability functions pass. Compact package verifies at 229 files / 135,815,185 bytes; manifest SHA-256 `06F410D8BED362AB7E4487B82DCB6D8991CC200312B0AC635A18185F979BB8F3`; all 227 runtime hashes match `app/dist`; CLI help exits 0.
- Packaged GUI import/readback passed on Excel 16.0 build 20430.0 with UI Automation, exact formula, verified backup, read-only reopen, app exit 0, and owned-process cleanup. PID 36464 was present before and after the smoke. Other Office builds and broader accessibility/security/workbook acceptance remain **UNVERIFIED**.

## W232 — non-query object preservation

- Source-based `run_excel_preserve_smoke.py` passed on Excel 16.0 build 20430.0. In addition to table/filter/style, chart, shape, PivotTable values, names, formulas, query-backed report links and privacy metadata, it now verifies data validation type/source, conditional-format rule/fill, and hyperlink address/display text after staged query replacement and read-only reopen.
- The fixture uses only static inline M tables and no external source. Owned Excel processes exited; pre-existing PID 36464 remained. Other Office builds and additional workbook-object relationships are **UNVERIFIED**.
- Package docs-only refresh verifies at 229 files / 135,815,952 bytes with manifest SHA-256 `F74868674D687D158C0E205B88B57BF12607F01AA01BFF4DBA030BFF7518E7B1`; runtime binaries match `app/dist` exactly. Package tests: 13 passed; contracts, 132 traceability references, verifier and CLI help pass. The docs refresh left the executable payload unchanged.

## W233 — current-package keyboard traversal

- `tools/run_packaged_keyboard_smoke.py --exe dist/PowerQueryManager/PowerQueryManager.exe` passed on the current Windows host. UI Automation visited 35 focusable controls, completed a full Tab cycle within the manager, and reached workbook path, Add files, query table, existing-query policy, backup and import in order. The process exited 0; no workbook or Excel instance was opened.
- The only Excel process before and after was pre-existing PID 36464. Main-window keyboard traversal is verified for this package. Modal keyboard paths, screen-reader use, physical display scaling, other Office builds and remaining Windows acceptance remain **UNVERIFIED**.

## W234 — 2026-10-23 — unresolved quoted M reference warning

- A failing-first test reproduced a quoted M identifier outside the available source list being shown as “No direct query references detected.” The scanner now reports it as ambiguous; graph tests verify no guessed edge is drawn. The warning does not claim the reference is absent from the workbook because the graph only receives the current source list.
- Focused dependency and graph tests: **46 passed**. Full suite and Windows build: **463 passed**. Rebuilt one-folder GUI and CLI published; package replacement smoke passed on Excel 16.0 build 20430.0 with exact formula, backup, read-only reopen and owned-process cleanup, and keyboard smoke passed across 35 controls. Excel acceptance used a separate ordinary formula and did not evaluate this unresolved reference.
- The scanner is advisory, not a complete M parser. The broader W14 cycle/missing-reference effect, other Office builds, screen-reader and physical display checks remain **UNVERIFIED**.
- Final docs-only package refresh after W234 verifies at 229 files / 135,817,933 bytes; SHA256SUMS SHA-256 `1841EAC79B0B159019B3B0219D3B010DFF661514021D5AEC5EB8349DC1215AD7`. Runtime hash parity, package staging tests (3), manifest verification and CLI help all pass. No executable payload changed during the docs refresh.

## W235 — rebuilt package v1.5 load and refresh reacceptance

- Re-ran the current compact GUI on disposable workbooks. New worksheet load, existing worksheet load at `Output!C4`, connection-only, and Data Model load all passed on Windows 10 build 26300 / Excel 16.0 build 20430.0 under a non-elevated token. Exact formulas, values 17/29, placement, connection/sheet/model distinctions, backup hashes, manager exit, and cleanup of owned Excel processes passed.
- The current packaged CLI refreshed connection-only and worksheet-loaded queries in background and synchronous modes against `127.0.0.1`; the worksheet output reopened with value 42. Four backups matched pre-refresh hashes; formulas/connections/sheet count remained correct; temporary tables/sheets/connections/stages were removed. Seven loopback GETs occurred; no external network was used.
- A delayed loopback Power Query refresh timed out at 3 seconds with CLI exit 5, truthful timeout report, unchanged original hash, matching backup, removed stage, and owned Excel process cleanup. Each smoke process set returned to pre-existing PID 36464.
- This reaccepts the rebuilt package on one Office build and inline/local fixtures only. Other Office builds, external providers, privacy/auth scenarios and remaining matrix cases remain **UNVERIFIED**.
- Docs-only refresh after W235 verifies at 229 files / 135,818,797 bytes with manifest SHA-256 `8BBD7A9302F3DB4EEE54F3EAC437E76E031D658993B369E4F4FF4330407C5746`. Runtime hash parity, package staging tests (3), manifest verification and CLI help passed; runtime executables were unchanged.

## W236 — 2026-10-23 — current package legacy import and replace/skip regression acceptance

- Current compact GUI passed mixed numbered `.m`/`.pq` import with the review sequence matching Excel query order and exact formulas. It passed all-skip with identical workbook hash/timestamp and no backup/stage. It also replaced an existing connection-only query while preserving that load association.
- Packaged GUI acceptance ran on Excel 16.0 build 20430.0 under a non-elevated token against generated disposable workbooks. Backup/read-only reopen checks where applicable, exit 0 and owned Excel cleanup passed. PID 36464 remained. Alternate Office builds and broader workbook relationships remain **UNVERIFIED**.
- W236 final docs-only package refresh verified 229 files / 135,819,461 bytes; manifest SHA-256 `3611CA0A803972996DA4FCB2170B509038107A1F47B24EA1171A49135813792C`. Runtime hash parity, package staging tests (3), manifest verification and CLI help passed. No runtime binaries changed.

## W237 — 2026-10-23 — current packaged deployment, package transfer and UI dialogs

- Current packaged deployment bundle applied to two disposable workbooks (`applied`, `applied`) and repeated as (`no_changes`, `no_changes`); exact formula/load readback, two verified original backups, disabled missing source ignored and owned-process cleanup passed. A separate mixed-target fake/service deployment passed with writable `applied`, read-only `blocked_preflight`, unchanged read-only hash/no backup and verified writable backup/formula.
- Packaged CLI query-package roundtrip verified three formulas, Unicode/punctuation name, unchanged source-workbook hash, verified target backup and no M evaluation. Packaged GUI UIA found named inspector, dependency graph, profile-mapping and package-comparison controls; profile changes were cancelled. Disposable GUI definition-only import passed exact formula/backup/read-only-reopen/process checks.
- Excel 16.0 build 20430.0, Windows 10 build 26300; only generated disposable files and non-elevated launches. PID 36464 remained. Other Office builds, provider variants and remaining policy/workbook matrix remain **UNVERIFIED**.
- W237 docs-only package refresh verified 229 files / 135,820,464 bytes; SHA256SUMS SHA-256 `29D7C2652A53E592C6AB7A0546107B3ED82319961039914F266DC86E520F881A`. Runtime parity, package tests (3), verifier and CLI help pass; binaries unchanged.

## W238 — 2026-10-23 — current package read/execute-only launch and import

- The current compact package was copied to a disposable directory with only read/execute ACLs. The write probe failed as expected. The non-elevated packaged GUI launched there and completed definition-only import to a separate disposable `.xlsx`; exact formula, verified backup, read-only reopen, app exit and owned Excel cleanup passed.
- The temporary read-only package was removed; pre-existing Excel PID 36464 remained. Windows 10 build 26300 / Excel 16.0 build 20430.0. Restricted enterprise application-control and profile policies remain **UNVERIFIED**.
- W238 final docs-only package refresh verifies 229 files / 135,821,085 bytes with manifest SHA-256 `226A69978C008FB35D223995EA8130F43A5BABA30F36A9DF14A61FFC5C083725`; all runtime hashes match the current Windows build. Package staging tests (3), verifier and CLI help pass; binaries unchanged.

## W239 — fresh source and package verification

- `app/.venv/Scripts/python.exe -m pytest -q`: **463 passed** on Windows 10 build 26300 / Python 3.11.9. Contract/schema validation passed and all **134** traceability references resolved.
- `tools/verify_release_package.py dist`: **PASS**, 229 files / 135,821,849 payload bytes. Hash comparison found 227 source runtime files, 227 packaged runtime files, zero missing, zero extra, zero different. Packaged CLI `--help` exited 0.
- No Excel process was started or changed during W239; pre-existing PID 36464 remained. W238 is the most recent package-based live Excel check. No new COM behavior is claimed here. Remaining Office-version, connector, policy, screen-reader, physical-display and workbook/security cases remain **UNVERIFIED**.
- The handoff verifier's default pristine-source check is no longer applicable after implementation edits and generated test caches; the compact release-package verifier remains the distribution-scope check.
- Final W239 docs-only package refresh manifest SHA-256: E7EC2D147E97064FB2638C4384905E37911FCB944A63D3BFA7A78A6EE157C0B1; verified payload remains 229 entries / 135,821,849 bytes.


## W240 — clean onedir build reproducibility

- A fresh PyInstaller `--clean` build of `PowerQueryInstaller.spec` used an isolated temporary work/dist path and the pinned deterministic environment (`PYTHONHASHSEED=0`, `SOURCE_DATE_EPOCH=1704067200`). Current and rebuilt payloads each contained 227 files; **0 missing, 0 extra, 0 hash differences**.
- Windows 10 build 26300 / Python 3.11.9 / PyInstaller 6.13.0. The temporary output was removed; published `app/dist` and `dist` stayed unchanged. This verifies same-host/toolchain repeatability only.
- No Excel process was started or changed; pre-existing PID 36464 remained. Alternate host/toolchain reproducibility and all open Excel/security/accessibility gates remain **UNVERIFIED**.
- W240 docs-only refresh verifies the final package at 229 manifest entries / 135,822,550 payload bytes; SHA256SUMS SHA-256 258FFAC12F6C6E745310FD0607BDA81D645CC7C4C192C9370817FD2DDBB0EBF7. Runtime files were unchanged.

## W241 — current-package modal keyboard acceptance

- `tools/run_packaged_keyboard_smoke.py --exe dist/PowerQueryManager/PowerQueryManager.exe --dialogs` passed UI Automation on the current compact package. Main-window traversal visited 35 focusable controls in a complete Tab cycle. The source-mapping dialog Tab cycles reached the path and M-source tables and their actions; Right Arrow switched the tab category; Escape dismissed the dialog. Process exit was 0.
- The smoke opened no workbook or Excel instance. Pre-existing Excel PID 36464 remained. Full portable suite: **465 passed**; focused smoke-helper and packaging tests: **10 passed**. Runtime executables were unchanged.
- Other modal dialogs, screen-reader use, physical display scaling and alternate Office builds remain **UNVERIFIED**.
- W241 docs-only package refresh verifies 229 manifest entries / 135,823,259 payload bytes; SHA256SUMS SHA-256 E728A92F157C0C47A0BFAF08C5102048E40FD450A25E52192C36F6905E1C7223. Runtime executables were unchanged.

## W242 — current-package operation-history keyboard path

- `tools/run_packaged_keyboard_smoke.py --exe dist/PowerQueryManager/PowerQueryManager.exe --dialogs` passed the operation-history modal Tab cycle through query history table, retention, snapshot toggle, refresh, retention apply, delete and close. Escape closed the dialog.
- `LOCALAPPDATA` pointed to a temporary profile, removed after the packaged manager exited. No user history was read or changed; no workbook or Excel process was opened. Pre-existing PID 36464 remained. Screen-reader and physical-display acceptance remain **UNVERIFIED**.
- W242 docs-only package refresh verifies 229 manifest entries / 135,823,957 payload bytes; SHA256SUMS SHA-256 FEDCA7BB01516FCDADCE923AF3F581BC453A0FC8E7ECBA03A3338984A6C841CD. Runtime files were unchanged.

## W243 — current-package dependency-graph keyboard path

- Current packaged keyboard smoke opened the empty dependency graph, completed a Tab cycle through its text summary and Close action, and dismissed it with Escape. No workbook or Excel process was opened. The combined smoke also passed main-window, source-mapping and temporary-profile operation-history navigation.
- This verifies the empty graph dialog's keyboard path only. Populated graph navigation, screen-reader, physical display and alternate Office acceptance remain **UNVERIFIED**.
- W243 docs-only package refresh verifies 229 manifest entries / 135,824,492 payload bytes; SHA256SUMS SHA-256 FD1E91708B3451677F91F04EA43A53911A7F03959F6BD0750D1E99F24A31A686. Runtime executables were unchanged.

## W244 — 2026-10-23 — query-inspector keyboard traversal

- Packaged GUI on Windows 10 build 26300, Python 3.11.9, PyInstaller 6.13.0: main window (35-control full Tab cycle), operation-history dialog, inspector search/Find actions/source-workbook-diff editors/save-copy/close, empty dependency graph, and source/M mapping tabs all passed UI Automation. Exit code 0.
- `app/build.ps1`: full suite **467 passed in 23.34s**; PyInstaller GUI/CLI onedir rebuild passed. Focused UI/package regressions passed **13 tests**.
- Temporary profile and `.pq` fixture only; no workbook or Excel launched. Pre-existing Excel PID 36464 remained. The history dialog was checked before the inspector because inspection records an audit-history entry. Screen-reader, physical display scaling, other Office versions/builds and other matrix items remain **UNVERIFIED**.
- Compact package verification: **229 manifest entries / 135,825,519 payload bytes**; `SHA256SUMS` SHA-256 `372081D56B27568BA6F85EF0562D9416DCBCD25D7F9C59F6DC1702789B96409A`; 227/227 runtime hashes match; verifier and CLI `--help` pass.

## W245 — 2026-10-23 — rebuilt packaged GUI import

- Command: `python -B tools/run_packaged_gui_import_smoke.py --exe dist/PowerQueryManager/PowerQueryManager.exe`; non-elevated Windows 10 build 26300 / Python 3.11.9, Excel 16.0 build 20430.0.
- The smoke created a disposable `.xlsx` and static inline-M source, imported in definition-only mode through the packaged GUI, then reopened read-only in Excel. UI Automation verified named controls and completion log; exact M readback, original-hash timestamped backup, manager exit 0 and cleanup passed.
- Seed/app/verification Excel PIDs 34508/43900/35764 exited. The process set returned to baseline `[36464]`; that pre-existing Excel process was untouched. Temporary fixtures were removed. No external source was contacted or refreshed. Other Office builds/providers and remaining matrix checks remain **UNVERIFIED**.

## W246 — 2026-10-23 — current-package existing-query replacement

- `python -B tools/run_packaged_gui_import_smoke.py --exe dist/PowerQueryManager/PowerQueryManager.exe --replace-existing` passed on Windows 10 build 26300 / Excel 16.0 build 20430.0, non-elevated. A disposable `.xlsx` existing definition-only query was replaced.
- Exact M readback, timestamped pre-import backup SHA-256, unrelated worksheet marker, read-only reopen and manager exit 0 passed. Seed/app/verification Excel processes were cleaned up; the process set returned to the pre-run `[36464]`; that unrelated existing instance was not touched. No refresh or external source was used.
- Other Office builds, replacement with loaded associations on this exact package, and remaining workbook/security/accessibility matrix items remain **UNVERIFIED**.
- Final W246 package contains 229 manifest entries / 135,827,136 payload bytes; `SHA256SUMS` SHA-256 `9BB0A72CD81637088DAC027E98F1C2CA73D24A4373B0CE01A751413074AD7A12`.

## W247 — 2026-10-23 — current-source clean rebuild parity

- A clean isolated PyInstaller rebuild of the current onedir source matched the published `app/dist/PowerQueryManager`: **227 files each; 0 missing, 0 extra, 0 changed**. `PYTHONHASHSEED=0`, `SOURCE_DATE_EPOCH=1704067200`; Windows 10 build 26300 / Python 3.11.9 / PyInstaller 6.13.0. The temporary output directory was removed automatically.
- This is same-host/toolchain evidence; cross-host/toolchain reproducibility remains **UNVERIFIED**.
- W247 compact package verification: 229 manifest entries / 135,827,789 payload bytes; `SHA256SUMS` SHA-256 `47A8848B716DBF4F68E0CA1A03DB60E8005F568232C8FB492A91AB1827269295`.

## W248 — 2026-10-23 — current-package load-mode reacceptance

- Four separate commands against the rebuilt package passed on Windows 10 build 26300 / Excel 16.0 build 20430.0, non-elevated: `--new-sheet`, `--connection-only`, `--existing-sheet`, and `--data-model`.
- All used disposable `.xlsx` files and static inline M. Exact M, timestamped original-hash backup and read-only reopen passed. New/existing worksheet outputs returned values 17/29; `Output!C4` placement and unrelated marker passed; connection-only had no worksheet/model table; Data Model membership and two-row count passed. Each GUI exited 0 and owned Excel process cleanup returned to baseline PID 36464. No external source was contacted.
- Other Office builds, connectors and remaining Windows matrix items remain **UNVERIFIED**.
- W248 release package verifies at 229 manifest entries / 135,828,764 payload bytes; `SHA256SUMS` SHA-256 `459A7E0B8DD70D371A960A54562BC4113A919354A908D1B04FB710A16195B6DF`.

## W249 — 2026-10-23 — current packaged CLI refresh

- Command: `python -B tools/run_excel_cli_refresh_smoke.py --cli-exe dist/PowerQueryManager/PowerQueryManagerCLI.exe`; Windows 10 build 26300 / Excel 16.0 build 20430.0.
- On a disposable workbook seeded with a static inline-M query and Mashup connection, the rebuilt packaged CLI ran background and synchronous refresh. Both reports returned `refreshed`; two backup hashes matched exact pre-refresh bytes; read-only inventory after each refresh preserved the formula and connection-only state. No external source contacted; smoke-owned Excel processes exited.
- Other Office builds, external providers and load-specific refresh paths on this current binary remain **UNVERIFIED**.
- W249 compact package: 229 manifest entries / 135,829,542 payload bytes; SHA-256 `4BBCEA17C88937471BB7CA5F4D08F345973CA34C1B2300D23D03F87A36936F63`.

## W250 — 2026-10-23 — current packaged CLI Power Query refresh

- Rebuilt packaged CLI ran connection-only and worksheet-loaded Power Query refresh, each in background and synchronous mode, against a temporary localhost CSV endpoint on Excel 16.0 build 20430.0.
- Four `refreshed` reports; seven loopback requests; exact M/load shape; four pre-refresh backup SHA-256 matches; worksheet value 42; temporary connections/sheets removed. Process set returned to `[36464]`; no external network was used.
- Other Office builds, Data Model refresh on this executable, external providers and remaining matrix items remain **UNVERIFIED**.
- W250 package: 229 manifest entries / 135,830,358 payload bytes; SHA-256 `0EC3A59B09A30FF2A76178C7B35C8C57BF457948F01B2C7466BC804FF8AB20C6`.

## W251 — 2026-10-23 — current packaged Power Query timeout

- Current packaged CLI synchronous refresh against a deliberately delayed loopback `Web.Contents` source returned exit code 5 after its configured 3-second timeout on Excel 16.0 build 20430.0. The request was observed.
- Target hash stayed unchanged; exactly one original-hash backup remained; stage was removed; process set returned to `[36464]`; no external network was contacted.
- Other Office builds and providers remain **UNVERIFIED**.
- W251 package: 229 manifest entries / 135,831,024 payload bytes; SHA-256 `69288982BAB19D796F3AF51E722E22D8301976AED4AC08E6981630E6FCE390DD`.

## W254 — 2026-10-23 — current packaged manager accessibility

- Rebuilt the packaged GUI after adding explicit accessible names to the workbook inventory and refresh-options controls. The Windows UI Automation keyboard smoke passed its 35-control main-window cycle and traversed operation history, query inspector, dependency graph, and source mapping. No workbook or Excel was opened; manager exited cleanly and the Excel process set stayed `[36464]`.
- The inventory and refresh-options controls were verified with five focused Qt UI tests, including a failing-first regression for the unnamed inventory table. They were not traversed via live UI Automation because that inventory is opened after Excel inspection.
- Package verification passed at 229 files / 135,833,330 bytes; `SHA256SUMS` SHA-256 `D76A562712B7900BFEA330DFECC876B8F82F89151C456C7310FE516F87B9E193`.
- Screen-reader testing, alternate OS display scaling, and full Office matrix remain **UNVERIFIED**.

## W255 — 2026-10-23 — latest packaged GUI replacement

- The rebuilt packaged GUI replaced an existing query definition in a disposable `.xlsx` on Windows 10 build 26300 / Excel 16.0 build 20430.0. Read-only reopen confirmed the exact new formula; the timestamped backup matched the original hash; an unrelated worksheet marker remained unchanged.
- Manager exited 0. App-owned Excel PID 41352 exited without forced termination; read-only verification PID 9168 was cleaned up. The process set returned to `[36464]`. An earlier invocation's command handle was not retained and is excluded; the directly captured rerun passed.
- The current package verified at 229 files / 135,833,330 bytes; SHA256SUMS hash `D76A562712B7900BFEA330DFECC876B8F82F89151C456C7310FE516F87B9E193`.
- This reaccepts one definition replacement path on one Office build. Other Office versions and remaining workbook/security cases remain **UNVERIFIED**.


## W257 — current packaged definition-only GUI import

- Current rebuilt GUI imported an inline M query into a disposable `.xlsx` on Excel 16.0 build 20430.0. Exact formula and original-hash backup passed read-only reopen; manager exited 0 and Excel process list returned to `[36464]`. No external source was contacted.
- Other Office builds and remaining workbook/security matrix gates remain **UNVERIFIED**.

## W258 — compact current Windows package

- Removed all 94 unused Qt `.qm` catalogs from PyInstaller collection; the English-only app does not load Qt translation catalogs. The verifier now rejects their return. Focused package verifier tests pass.
- Current GUI and CLI were rebuilt. CLI help passed; packaged UI automation and definition-only disposable workbook import passed on Excel 16.0 build 20430.0 with exact formula/backup readback and process cleanup. Runtime payload decreased by 6,277,341 bytes.
- The broader Office/version/security matrix remains **UNVERIFIED**.

## W259 — compact Qt runtime

- The package removes hook-collected Qt QML/Quick, PDF, virtual-keyboard and OpenGL runtime components not used by the current interface. The verifier rejects their reappearance; package verifier tests passed (7).
- Rebuilt package: 124 manifest files / 88,553,103 bytes. Packaged CLI help and UI/import smoke passed on Excel 16.0 build 20430.0; exact formula, verified backup, read-only reopen, process cleanup and existing PID 36464 were confirmed. Runtime payload fell a further 41,004,504 bytes.
- Other Office builds and untested optional Qt component surfaces remain **UNVERIFIED**.

- W259 final manifest: 124 files / 88,553,103 payload bytes; SHA256SUMS SHA-256 `36A3776805249AA6666A5639AA12D86F3E204A9332E14B5C3D206C3C1939EF7F`.

## W260 — Qt Network/TLS removal

- Removed unused Qt Network/TLS bindings, plugins, and Qt OpenSSL copies. The app source contains no QtNetwork use; refresh executes through explicit desktop Excel automation. The verifier regression and package-verifier suite passed (8 tests).
- CLI help and current packaged GUI UIA/definition-only import passed on Excel 16.0 build 20430.0. Package contains 116 files / 76,193,358 bytes; formula, original-hash backup, read-only reopen and owned-process cleanup passed. Other builds and external-source refresh remain **UNVERIFIED**.

- W260 final package manifest: 116 files / 76,193,358 payload bytes; SHA256SUMS SHA-256 `5DDC03E6FB531348E2F842AB8E515641F92D697E725E73947851A025D927836A`.

## W261 — network-trimmed CLI refresh

- The current compact CLI passed Excel Power Query refresh in background and synchronous modes for connection-only and worksheet loads using a loopback-only CSV server; all four reports returned `refreshed`, backups and formula were preserved, worksheet value 42 reopened, and temporary objects were removed.
- Excel 16.0 build 20430.0; ten loopback requests; process list returned to `[36464]`; no external network used. Other Office builds/providers and Data Model refresh on this exact package remain **UNVERIFIED**.

- W261 final package: 116 files / 76,193,358 payload bytes; SHA256SUMS SHA-256 `26EABD1707BE77F1B151C553D25343A3B9DD07C04812810A2F5408AE258D4273`.

## W262 — current compact Data Model import

- Current packaged GUI's Excel capability probe enabled Data Model, and a disposable static-M import passed on Excel 16.0 build 20430.0. Exact formula, Data Model membership, 2-row count, matching backup hash, manager exit and Excel process cleanup passed; PID 36464 remained.
- No external source used. Other builds and Data Model refresh on this package remain **UNVERIFIED**.

- W262 final package: 116 files / 76,193,358 payload bytes; SHA256SUMS SHA-256 `8647F7D5E6E4E21DEF5F7974F8263C4C3F45779C8997ED8C2466C38C84B45A94`.

## W263 — current packaged Data Model refresh

- Data Model import followed by packaged CLI refresh returned `refreshed` in 3.031s on Excel 16.0 build 20430.0. Connection association, two rows, exact backup hash, staged read-only verification and formula/model readback passed. Process set returned to `[36464]`; no external source was used.
- Data Model refresh is synchronous only through the observed Excel API. Other builds and remaining workbook/security matrix remain **UNVERIFIED**.
- W263 final package: 116 files / 76,193,889 payload bytes; SHA256SUMS SHA-256 `01F8798055DBB528E027AF43BE7D630B663B4892A6F2BA775152B54FF26995DF`.

- W264 verifies packaged GUI replacement preserves an existing Data Model association and does not implicitly refresh it: changed formula, `InModel`, unchanged row count 1, original-hash backup, read-only reopen, manager exit and owned-process cleanup passed on Excel 16.0 build 20430.0. Other Office builds and complex model relationships remain **UNVERIFIED**.

- W265 verifies packaged GUI replacement preserves an existing worksheet query table at `Output!C4` and its preexisting one-row data without implicit refresh. Exact changed formula, table identity, anchor, backup, read-only reopen and owned-process cleanup passed on Excel 16.0 build 20430.0. Other Office builds and workbook relationships remain **UNVERIFIED**.

- W266: Source comparison ignores terminal CR/LF that Excel removes when saving M formulas, avoiding a false MODIFIED status after a round trip. Diff text remains exact. Failing-first regression reproduced the defect; 3 comparison tests passed, and the full Windows suite passed 479 tests during `app/build.ps1`. PyInstaller rebuilt GUI and CLI; W266 is a comparison-only change and does not have new real-Excel behavior evidence.

- W267 verifies the rebuilt packaged GUI reports `UNCHANGED` for a source differing from the Excel-stored formula only by terminal CRLF, then completes replacement with exact backup, source hash, read-only formula and process checks on Excel 16.0 build 20430.0. Other Office builds remain **UNVERIFIED**.

- W269 profile migration/validation: 7 focused tests passed; `app/build.ps1` passed 484 tests and rebuilt GUI/CLI. Contract validation, compact package verification, and packaged CLI `--help` passed. The package contains the updated loader. Other Office builds and broader legacy profile shapes remain **UNVERIFIED**.
- W269 packaged CLI smoke validated an unversioned-compatible/schema-0 profile, returned a v1 validation report, exited 0, and preserved the input file. No Excel automation was needed.

- W270 closes a profile save/load inconsistency: v1 schema validation now rejects floating-point `1.0` even though it compares equal to integer 1 in Python. The failing-first regression and seven focused profile tests pass; `app/build.ps1` completed its full suite and rebuilt the GUI/CLI. Contract validation, package verification (116 files / 76,198,528 bytes), and packaged CLI help pass. A disposable packaged GUI replacement smoke also passed on Windows 10 build 26300 / Excel 16.0 build 20430.0 under a non-admin token, with exact formula readback, backup, marker preservation, exit 0, and return to baseline Excel PID `[36464]`. This Excel exercise verifies the current package's import path; the profile schema change itself does not alter Excel behavior. Overall release gates remain UNVERIFIED.
- W271 failing-first export regression reproduced Windows filename failure for a valid 255-character Excel query name. After bounding folder and package filename stems, the focused export/package selection passed (3 tests); a fresh onedir rebuild passed packaged GUI keyboard traversal (35 controls), CLI `--help`, contract validation, and package verification. Final compact package: 116 files / 76,199,132 bytes; `SHA256SUMS` SHA-256 `479F59F9AF3EDFB86720FF518A084F1B66AFA69F51DEE8A2D9FD9678EFB15135`. This change does not exercise or alter Excel COM behavior; complete release gates remain UNVERIFIED.
- W272 validated schema-0 profile migration using the current packaged CLI: the JSON result was valid schema v1, exit code 0, and the profile file was byte-identical afterward. The offline smoke did not open a workbook or Excel.
- W273 reran the packaged GUI replacement smoke against the exact current W271 compact package. On Windows 10 build 26300 / Excel 16.0 build 20430.0, non-admin: exact formula, backup hash, marker, manager exit 0, and return to baseline PID `[36464]` passed. The operation-owned Excel process exited normally; the read-only verification process required forced cleanup. Existing PID 36464 was untouched; no refresh ran.
- W274 targeted package-export test passed for a 255-character Unicode query name, checking a bounded ZIP filename component and full-name manifest preservation. No runtime binary changed.
- W275 reproduced display-path leakage of Bearer/Basic authorization credentials and signed URL values. After extending sanitization for authorization headers, client secrets, and common signed parameters, the focused redaction module passed **13 tests**. The rebuilt compact package passed CLI help, contract validation, package verification (116 files / 76,200,376 bytes), and GUI keyboard traversal (35 controls); no workbook or Excel process was opened.
- W276 reproduced leakage of quoted JSON `refresh_token`, `id_token`, and `client_secret` fields. The focused redaction module passed **14 tests** after fix; the rebuilt package passed CLI help, contract and package verification (116 files / 76,200,852 bytes), and GUI keyboard traversal. No workbook or Excel instance was opened.
- W277 reproduced compound key leakage for `AWS_SECRET_ACCESS_KEY`, `x_api_key`, and `private_key`. The focused redaction module passed **15 tests** after fix; the onedir rebuild passed package/contract verification, CLI help, and GUI keyboard traversal (35 controls). No workbook or Excel instance was opened; full suite remains unrerun.
- W278 reproduced leakage from Negotiate `Proxy-Authorization` and Digest `Authorization` header values. The focused redaction module passed **16 tests** after fix; onedir rebuild, package and contract verification, CLI help, and GUI keyboard traversal (35 controls) passed. No workbook/Excel instance was opened, and the full suite was not rerun.
- W279 reran the replacement flow through the W278 compact packaged GUI on a disposable workbook, non-elevated, Windows 10 build 26300 / Excel 16.0 build 20430.0. Exact formula readback, original-hash backup, unrelated marker, manager exit 0, and return to baseline Excel PID `[36464]` passed. The manager-owned process exited normally; read-only verification Excel required forced cleanup. No refresh or external source was used.
- W280 ran the W278 compact package's mixed `.m`/`.pq` profile through the visible GUI on a disposable workbook, non-elevated, Windows 10 build 26300 / Excel 16.0 build 20430.0. Three numbered entries were reviewed and created in the same order; exact formulas, backup, read-only reopen, manager exit 0, and process cleanup passed. Existing Excel PID 36464 remained untouched. Other prefixes, ties, manual reorder, and Office builds remain **UNVERIFIED**.
- W281 reran `--skip-existing` through the W278 packaged GUI on a disposable `.xlsx`, Windows 10 build 26300 / Excel 16.0 build 20430.0, non-elevated. Workbook SHA-256/timestamp and query formula were unchanged; no backup/stage; no-change UI result; manager exit 0; owned Excel cleanup returned to baseline PID 36464. Existing PID 36464 was untouched.
- W282 fixed the reproduced leak from a semicolon-containing quoted Digest parameter. Focused redaction module passed **17 tests**; onedir rebuild, package/contract verification, CLI help, and 35-control GUI keyboard traversal passed. Excel was not exercised because this change affects error-text redaction only.
- W283 reran replace-existing on the W282 package using a disposable `.xlsx`, non-elevated, Windows 10 build 26300 / Excel 16.0 build 20430.0. Exact formula, pre-import backup hash, unrelated marker, manager exit 0, and return to baseline Excel PID 36464 passed. Manager-owned Excel exited normally; read-only verification Excel required forced cleanup.
- W284 ran Data Model import through the W282 packaged GUI, non-elevated, Windows 10 build 26300 / Excel 16.0 build 20430.0. Capability probe enabled the mode; read-only reopen verified formula, model membership and row count 2; backup hash, manager exit 0, and process cleanup passed. PID 36464 remained untouched. Complex models and other Office builds remain **UNVERIFIED**.
- W285 ran packaged GUI Data Model import followed by explicit packaged CLI refresh on the same disposable workbook. Excel reported synchronous completion (3.047 seconds), connection and two model rows verified, staged verification passed, pre-refresh backup hash matched, no stage remained, and owned process cleanup returned to baseline PID 36464. No external source used; other builds remain **UNVERIFIED**.
- W286 ran existing-worksheet load through the W282 packaged GUI on a disposable workbook, non-elevated, Windows 10 build 26300 / Excel 16.0 build 20430.0. Capability probe enabled the mode; read-only reopen verified exact formula, `Output!C4` table anchor, and values 17/29. Unrelated marker and backup hash matched; manager exit 0 and owned-process cleanup passed. Existing PID 36464 remained untouched.
- W287 ran connection-only load through the W282 packaged GUI on a disposable workbook. Capability probe enabled the mode; explicit workbook connection and exact formula persisted after read-only reopen with no worksheet/model table and no refresh. Backup hash, manager exit 0, and owned-process cleanup passed; existing PID 36464 remained untouched.
- W288 ran new-worksheet load through the W282 packaged GUI on a disposable workbook, non-elevated, Windows 10 build 26300 / Excel 16.0 build 20430.0. Probe enabled the mode; read-only reopen verified exact formula, `GuiOutput!A1` table and values 17/29. Backup hash, manager exit 0, and process cleanup passed. PID 36464 remained untouched.
- W289 ran default definition-only import through the W282 packaged GUI on a disposable `.xlsx`: exact formula, backup, read-only reopen and manager exit 0 passed; operation report showed no connection/worksheet load. Same run's UI Automation checked named query inspector, dependency graph and profile mapping controls; mapping changes were cancelled. Baseline Excel PID 36464 remained.
- W290 ran the W282 packaged CLI's offline profile validator against the example profile. It returned `valid: true`, schema v1, and three valid sources; profile SHA-256 was identical before/after, and the Excel process set did not change.
- W291 ran the W282 packaged CLI against a portable `.pqmdeploy` bundle and two disposable `.xlsx` workbooks on Excel 16.0 build 20430.0. First batch `applied/applied`; repeat `no_changes/no_changes`; formula/load readback and two original-hash backups verified; disabled missing source ignored; all owned Excel PIDs exited and PID 36464 remained.
- W292 ran packaged CLI `plan --dry-run` against two Excel-created disposable `.xlsx` targets. Both plans reported `create`; target hashes were unchanged, no backup/stage was created, and the CLI run left the Excel process set unchanged at baseline PID 36464.

## W293 current compact-package acceptance

- Host: Windows 10 build 26300, non-elevated; Excel 16.0 build 20430.0.
- Command: `python -B tools/run_excel_cli_package_import_smoke.py` using `dist/PowerQueryManager/PowerQueryManagerCLI.exe`.
- Result: **PASS**. Three query definitions imported from a disposable export workbook into a disposable blank target. Names/formulas were read back exactly, including `Équipe – Résultats 2026 (Q1)`. Export did not change the source workbook; the single target backup matched its original SHA-256. No external source evaluation occurred.
- Cleanup: the smoke's owned Excel instances exited; only the pre-existing Excel PID 36464 remained.
- Scope: one Office build and workbook shape; other versions and the broader workbook/provider/security/accessibility matrix remain **UNVERIFIED**.
## W294 packaged unsafe query-package rejection

- Host/package: Windows 10 build 26300, non-elevated; current `PowerQueryManagerCLI.exe`.
- A disposable ZIP containing `../escape.pq` was submitted with `import --dry-run` against a sentinel target. The CLI exited 2 and reported the unsafe ZIP path before workbook inspection.
- Target SHA-256 remained unchanged, no traversal file was created outside the temporary directory, and Excel was not launched.
- Scope: one traversal entry only; other archive attack shapes retain their separate test/evidence status.
## W295 current packaged GUI VBA-bearing workbook

- Command: `python -B tools/run_packaged_gui_import_smoke.py --exe dist/PowerQueryManager/PowerQueryManager.exe --xlsm --vba`.
- Windows 10 build 26300 / Excel 16.0 build 20430.0, non-elevated. The packaged GUI imported a definition-only query into a disposable VBA-bearing `.xlsm`; exact formula, VBA source and unchanged `Workbook_Open` marker were verified after read-only reopen. Backup verified, manager exited 0, owned Excel processes were cleaned up, and baseline PID 36464 remained.
- Scope: one harmless macro project and one Office build; other macro-security configurations remain **UNVERIFIED**.
## W296 release-doc refresh

- Updated `tools/package_release.py` so a docs-only release refresh uses the full v1.3.1–v2.0.0 scope and includes W293–W295 packaged acceptance results.
- Command: `python -B tools/package_release.py --docs-only`; followed by `python -B tools/verify_release_package.py`.
- Result: **PASS**, 116 files / 76,203,691 payload bytes. `dist/SHA256SUMS` SHA-256 is `946CE3E511A48F74B7D2F58C793FB3F717EB5A86508CDC37F46850E1FA64DD7D`. GUI and CLI executable timestamps were unchanged; no rebuild occurred.
- Release gates remain **UNVERIFIED** for the wider Windows/Office matrix.
## W297 current packaged CLI mixed read-only deployment

- Command: `python -B tools/run_excel_deployment_mixed_smoke.py --cli-exe dist/PowerQueryManager/PowerQueryManagerCLI.exe`.
- Windows 10 build 26300 / Excel 16.0 build 20430.0, non-elevated. Two disposable workbooks were deployed as one CLI batch: writable target `applied`, read-only target `blocked_preflight`; the CLI returned partial-failure exit code 5.
- Successful target formula and original-hash backup verified. Read-only target had no backup and its hash stayed unchanged. Both packaged CLI deployment and read-only inspection returned to baseline Excel PID 36464.
- This verifies read-only target isolation for the shipped CLI on one Office build; mid-commit failures, other ACL configurations and alternate builds remain **UNVERIFIED**.
- The package docs were refreshed without rebuilding binaries. Verifier: 116 files / 76,204,147 payload bytes; manifest SHA-256 `33D45AEAF496D75C9F02D930B8368614E74D55BB03A3758ECC315969466CD6DC`.
## W298 injectable workbook lifecycle adapter

- Added a shared port for workbook add/open/save/save-as/close and routed inspector, import, refresh and capability-probe lifecycle calls through it. Existing Excel application and transaction adapters remain supported.
- Focused fake-COM/service selection: **91 passed**. Rebuilt on Windows 10 build 26300 / Python 3.11.9 / PyInstaller 6.13.0.
- Packaged GUI definition-only import passed on Excel 16.0 build 20430.0 with exact formula readback, verified backup, read-only reopen, manager exit 0, and owned Excel cleanup. CLI `--help` and package verifier passed; only pre-existing Excel PID 36464 remained.
- Package: 116 files / 76,208,339 payload bytes; `SHA256SUMS` SHA-256 `AE23240EC59E1DBE2127849923AB025FEFA39106D0A10F9C289005E8663B44E2`. The full suite and broader Office matrix remain **UNVERIFIED**.
## W299 — packaged CLI refusal for an open read-only workbook without an owner-lock file

- Command: `app/.venv/Scripts/python.exe -B tools/run_excel_open_without_lock_smoke.py --cli-exe dist/PowerQueryManager/PowerQueryManagerCLI.exe`.
- Windows 10 build 26300 / Excel desktop 16.0 build 20430.0, non-elevated. The harness created a disposable `.xlsx`, opened it read-only in its own Excel instance, and confirmed Excel had not created `~$active.xlsx`. The packaged CLI ran `deploy --apply --confirm DEPLOY` and returned `blocked_preflight` with exit 5 and the active-use message.
- The target hash was unchanged; no backup or stage was created; Excel process IDs remained the initial PID 36464 plus the harness-owned PID during CLI execution. The pre-existing process was untouched. Existing service import/refresh/overwrite-restore refusals also passed; the harness-owned Excel process exited during cleanup.
- Other Excel builds, lock-file providers and wider failure cases remain **UNVERIFIED**.

### W300 — query-grid removal undo

- The focused Qt model suite passed 40 tests after adding one undo command for removing selected rows. The regression verifies one history entry, original stable IDs and row order after undo, and the same removal after redo.
- This is source-level Qt verification; the full suite, packaged keyboard test, screen-reader review and physical display review were not run for this change.
- The updated compact package was rebuilt with PyInstaller 6.13.0 and Python 3.11.9, then its manager passed the 35-control Windows UI Automation keyboard cycle; CLI `--help` and manifest verification passed. Removal itself is covered by source Qt model tests, not this UIA run.

### W301 — CLI local history for planning and deployment

- Focused command: `python -B -m pytest -q tests/test_cli.py::test_cli_multitarget_plan_uses_inspector_worker_and_writes_no_overwrite_json tests/test_cli.py::test_cli_explicit_confirmation_argument_allows_noninteractive_deployment tests/test_cli.py::test_cli_multi_workbook_deploy_reports_per_target_results_and_partial_exit`.
- Result: **3 passed**. Tests use an isolated temporary `LOCALAPPDATA` and verify dry-run comparison, successful deployment, and partial deployment entries in the shared `HistoryStore`. They assert aggregate details and confirm source/workbook paths are excluded.
- No Excel COM operation is needed for these CLI fixtures. W301 packaged CLI runtime acceptance passed on Excel 16.0 build 20430.0 using disposable writable and read-only workbooks; it verified partial-result isolation, target backup/hash, process cleanup, and a temporary local history event with aggregate counts only. W301 packaged GUI definition-only import also passed with exact M readback and backup verification. Other Office builds and remaining workbook, security, and accessibility matrix rows remain **UNVERIFIED**.

### W302 — current-package source-grid duplicate handling

- The source model now filters duplicate paths in the same add batch, retaining the first selected row. Focused Qt model suite: **41 passed**; the rebuilt package manifest verifies.
- Packaged keyboard traversal could not be accepted in this run: one attempt reported ambiguous matching window titles and another reported focus leaving the manager. A separate inventory observed one visible main window. Treat keyboard traversal as **UNVERIFIED** pending a reliable run. The current package separately passed the primary disposable definition-only GUI import on Excel 16.0 build 20430.0; other load modes/builds remain unverified.

## W294 — Dependency Finder Windows acceptance harness (2026-10-27)

- Added `tools/run_dependency_windows_acceptance.py`. It requires Excel and
  `DAO.DBEngine.160` by default, creates disposable linked-table/saved-query
  `.accdb`/`.mdb` fixtures plus genuine Excel Mashup, external-link and large
  formula workbooks, and checks expected edges, ambiguous/unresolved names,
  Quick/Deep parity, CLI multi-target scans, GUI filters/graph/JSON/CSV,
  progress/cancellation, file SHA-256/mtime, Access locks, temp cleanup and
  owned Excel process cleanup. Optional existing mapped, UNC and synchronized
  roots use `--mapped-root`, `--unc-root`, and `--sync-root`; it does not create
  shares, map drives, or change sync settings. See the script help for fixture
  retention and JSON report options.
- Exact source commit at the start of the diagnostic run: `7c9d94b5f3d7502b0204de7873042604fb36ee24`; the run included then-uncommitted harness and UTF-16 Mashup fix changes, so this SHA is a base revision, not a clean tested revision. Final clean commit SHA is recorded in evidence/STATUS.md after commit.
- Windows build command `app/build.ps1`: **537 passed in 39.46 s**; GUI and CLI rebuilt with PyInstaller 6.13.0 / Python 3.11.9. Package staging/verification: **116 files / 76,310,327 bytes**. Packaged CLI help exposed `dependency-scan`, `dependency-map`, Quick and Deep. The package contains the regression fix for UTF-16LE/BE Office customXml Mashup markers.
- Diagnostic command: `python tools/run_dependency_windows_acceptance.py --dao-progid DAO.DBEngine.120 --report %TEMP%\\pqm-dependency-harness-final.json`. This explicit DAO 12.0 compatibility run is diagnostic only. Required `DAO.DBEngine.160` is not registered on this host; only DAO 12.0 was available. The harness's default ACE 16.0 acceptance gate therefore remains **FAIL / OPEN**.
- Excel 16.0 build 20430.0 created real Power Query Mashup and external workbook references. Four top-level Access databases and four Excel workbooks plus two duplicate-basename Access fixtures were scanned. Both modes found the eight expected source-target pairs (12 dependency records), plus one ambiguity issue and two unresolved references (duplicate `Ambiguous.accdb` and out-of-scope `OutsideScope.xlsx`). Multi-target CLI matched six relationships. Quick: 6.43 s, peak process-tree working set 189,259,776 bytes. Deep: 3.63 s, 186,060,800 bytes. The large workbook contained 31,200 external-link formula cells. Both modes found the external workbook edge through Excel's external-link relationship metadata; neither emitted formula-reference edges, and there were no mode-exclusive edges in this fixture.
- GUI acceptance checks passed in this diagnostic: incoming/outgoing/both row counts, 12-edge visual graph, JSON and CSV exports (12 edges each), progress observation and cancellation after one inspected source. Source hashes and timestamps were unchanged; no Access lock files remained; DAO temp directories were cleaned; no Excel PIDs remained; fixture tree was removed. However, Excel's fixture-owned process did not exit after `Quit()` and required forced termination after 21.33 s, so the cleanup acceptance gate correctly reports **FAIL**, despite returning to the original process set.
- Mapped SMB, UNC and synchronized folder repetitions were **NOT RUN** because no roots were supplied. Overall packaged-app acceptance remains **FAIL / OPEN** until DAO 16.0 and natural Excel process cleanup pass and the three nonlocal location types are exercised.
- A clean-revision ACE 16.0 preflight on `f542a72f5c9087b44dae85cff614f5b0ed07b89f` failed closed before fixture creation: `DAO.DBEngine.160` is not registered; only DAO 12.0 is available. This is the exact committed harness revision; the longer DAO 12.0 run was a diagnostic of the same source plus earlier uncommitted build state and is not presented as a clean-commit acceptance pass.

## W295 — Excel cleanup repair, synchronized-path run, and repeated performance (2026-10-27)

- Supersedes W294's Excel-cleanup result. Excel fixture setup and PID inventory are now separate from unnecessary parent COM initialization; DAO fixtures run in a short-lived worker so their COM lifetime cannot leak into Excel automation. The complete real DAO→Excel fixture sequence with 31,200 formulas exited naturally in **3.63 s**, with no cleanup errors, no forced termination, and no Excel process remaining. The earlier isolated 130-cell check also exited normally. This cleanup repair is validated on the current working tree; committed SHA is recorded in the matching evidence/STATUS.md follow-up.
- Final diagnostic command used `DAO.DBEngine.120` and `--sync-root C:\\Users\\Robert\\OneDrive --performance-repeats 3`. Result: harness **PASS** for local and OneDrive synchronized roots. Each run scanned 10 files (4 top-level Access databases and 4 top-level Excel workbooks plus two duplicate-basename Access databases). The 31,200-cell external-formula workbook and genuine Excel 16.0 build 20430.0 Mashup/external links were created with desktop Excel.
- Local, 3 interleaved repetitions per scan mode: Quick median **3.842 s** (2.922–8.667 s), max observed process-tree working set **120,893,440 bytes**; Deep median **3.841 s** (3.646–3.996 s), max **118,755,328 bytes**. OneDrive synchronized folder: Quick median **3.241 s** (2.821–7.259 s), max **120,717,312 bytes**; Deep median **3.546 s** (3.362–3.615 s), max **118,595,584 bytes**. Each mode found 12 dependency records and two unresolved references (one ambiguous duplicate basename plus the out-of-scope reference); Quick and Deep had no mode-exclusive edges in this fixture. The formula references were represented through Excel external-link relationship metadata.
- Both locations passed source SHA-256/size/mtime preservation, no Access lock files, Access temporary-directory cleanup, Excel PID-set restoration and disposable fixture cleanup. Local GUI directional filters, visual graph, JSON/CSV exports, progress and cancellation passed again; the OneDrive repeat exercised packaged CLI scans and read-only integrity checks.
- ACE 16.0 preflight rerun on the working tree still fails closed before fixture creation: `DAO.DBEngine.160` is not registered; `DAO.DBEngine.120` is available. The Windows session is not elevated, so no Office/runtime install was attempted. Microsoft SMB mappings list is empty and `net use` reports no connections; mapped SMB and UNC repetitions remain **NOT RUN**. OneDrive is the synchronized-folder case exercised. Overall required ACE 16.0 plus SMB/UNC acceptance remains **FAIL / OPEN**, although the DAO 12.0 diagnostic and local/OneDrive acceptance are successful.

## W296 — clean-revision confirmation (2026-10-27)

- Exact tested harness commit: `a13ff507e76e4cca80fd204c69781a033ffbed39`. Clean-run JSON: `%TEMP%\\pqm-dependency-a13ff50-clean.json`; ACE 16.0 preflight JSON: `%TEMP%\\pqm-dependency-a13ff50-ace16.json`.
- DAO 12.0 diagnostic acceptance returned **PASS** on both local and OneDrive synchronized roots. Excel 16.0 build 20430.0 exited naturally in **3.23 s**, without forced termination; no Excel process remained. Each scan mode returned 12 dependency records, with two unresolved references and one ambiguity issue. The local GUI checks passed; the synchronized-root run repeated packaged CLI and read-only integrity checks.
- Three interleaved packaged CLI repetitions on the exact commit: local Quick median **3.060 s** (2.910–3.307 s), peak process-tree working set **190,595,072 bytes**; Deep median **3.501 s** (3.388–3.682 s), peak **192,208,896 bytes**. OneDrive Quick median **3.174 s** (3.112–3.318 s), peak **194,142,208 bytes**; Deep median **3.610 s** (3.565–3.733 s), peak **191,352,832 bytes**. The fixture included 31,200 external-link formula cells.
- Source hashes/sizes/timestamps remained unchanged at both locations; no Access lock files or scan temporary directories remained; both disposable fixture trees were removed.
- ACE 16.0 preflight on the same SHA **FAIL**: provider class is unregistered; only `DAO.DBEngine.120` is available. No Access/Office runtime was installed because this Windows session is not elevated. `Get-SmbMapping` and `net use` returned no connections, so mapped SMB and UNC remain **NOT RUN**. Overall Windows acceptance remains **FAIL / OPEN** pending DAO 16.0 and SMB/UNC access.

## W297 — scope-limited feature acceptance waiver (2026-10-27)

**Decision:** the repository owner explicitly waived two Windows acceptance checks for merging PR #1, not because they passed. This permits **Dependency Finder feature-branch integration with documented limitations**; it does **not** certify those environments or mark the whole Power Query Manager release matrix complete.

| Original check | Recorded evidence | Decision for PR #1 | Remaining limitation |
| --- | --- | --- | --- |
| ACE DAO 16.0 (`DAO.DBEngine.160`) | Clean revision `a13ff507e76e4cca80fd204c69781a033ffbed39`: provider not registered, preflight stopped before fixture creation; DAO 12.0 real `.accdb`/`.mdb` diagnostic passed | **WAIVED / NOT TESTED** | No ACE 16.0 compatibility claim; rerun on that engine before relying on ACE 16.0 deployments |
| Mapped SMB and UNC | No mappings or share connections were present on the Windows test host; no scan ran | **WAIVED / NOT TESTED** | No mapped-network/UNC compatibility or path-identity assurance; verify before use on network shares |

**Checks actually passed (W296, not re-executed by this documentation change):** 537 Windows tests in the earlier build report; packaged GUI/CLI; DAO 12.0 disposable Access/Excel graph; local and OneDrive Quick/Deep scans (31,200 external-formula cells); GUI filter/graph/JSON/CSV/progress/cancel; unchanged original SHA-256/size/timestamps; no Access lock/temp remnants; clean Excel exit in 3.23 seconds without forced termination. Both modes yielded 12 dependency records, two unresolved references and one ambiguity warning. The clean tested harness revision is `a13ff507e76e4cca80fd204c69781a033ffbed39`; evidence is also summarized in W296 and `evidence/STATUS.md`.

**Scope:** PR #1 may be merged with the two waivers clearly disclosed. Waiver is *not* equivalent to PASS, does not disable fail-closed DAO/provider diagnostics, and does not alter scanner security or source integrity requirements. Other project-wide release and host-compatibility gates remain independently applicable. Do not advertise ACE 16.0 or mapped SMB/UNC as validated. Future testing should record actual results rather than inheriting this waiver as proof.

## W298 — Vector graph viewport regression and rebuilt package (2026-10-28)

- On Windows build 26300 / Python 3.11.9 / PySide6 6.8.3, the packaged Finder mapped two synthetic OOXML workbooks to one exact internal edge (`Dashboard.xlsx` → `Inventory.xlsx`). Finder reported 0 outside-scope/unresolvable references. SHA-256 values before and after scanning matched: Dashboard `2E12C3926965C9B96F47A23533AADE6D53743114C44E23EED2662DAA91888149`; Inventory `D7B749C529EA41E9C2C0544C97415698A8D20277D4BE930E614CAF6191FCBC1F`.
- Visual inspection found the graph initially opened panned beyond the left edge. Added first-show auto-fit in `file_dependency_graph_dialog.py` plus a 960×600 graph-bound assertion. Exact app-source revision: `536ee6ce5d901eabbe7345c8c2ffc287cc31d372`.
- Rebuilt with `app/build.ps1`: **541 passed in 41.26 s**; PyInstaller completed successfully. Staged package verifier: **PASS, 116 files / 76,330,146 bytes**. Packaged CLI `--help` exited successfully and lists dependency-scan/map commands. GUI process launch succeeded. The Qt regression covers the auto-fit behavior; a fresh packaged graph check is recorded in the follow-up bullet below.
- Packaged EXE SHA-256: GUI `34AC00810E01D6E1E4C23ED667491E2059F74AC9074ECDAF970A08BF494665DC`; CLI `6B30D61A87A2A164949C35EFB8BEEE2226B40344741CD64BA974365F752F0382`.
- Follow-up packaged GUI check on Windows, app-code revision `536ee6ce5d901eabbe7345c8c2ffc287cc31d372`: a fresh Quick map of the same two synthetic files produced one exact edge; the graph opened with both nodes and the lime directed edge fully visible on first show, and its accessible text listed `Dashboard.xlsx depends on Inventory.xlsx [exact_path]`. The launched GUI was closed; a subsequent process check found no `PowerQueryManager` process. EXE hashes were re-read and match the values above. This visual check used the current 1180×730 app window and does not cover the requested display-size/DPI matrix.
- Additional packaged UI checks at 1180×730: Ctrl+K opened the command palette with workbook, dependency, import, history, undo and redo actions; the inspector checkbox opened the right-side Source M / Workbook M panel; the light theme rendered the workspace and Dependency Finder controls with readable text and visible focus outline. Finder opened with Quick/Deep mode selectors and file/target sections; no scan was run in this light-theme check. This does not establish full keyboard-only navigation or cover every dialog/workflow.
- Physical 125%/150% DPI and 960×600 / 1366×768 host matrix and broader remaining workflows/dialogs remain **OPEN / NOT VERIFIED**. ACE DAO 16.0 and mapped SMB/UNC remain **WAIVED / NOT VERIFIED**, never PASS.

## W302 — physical display layout and scale observation (2026-10-28)

- Using the freshly packaged GUI from app-source commit `3e45bc450482b5311a6a30938c27ea19fca50a45`, Windows was set through Display Settings to 1366×768 at 100%. The maximized main workspace was visually inspected; the Vector metric cards and action sections fit without visible clipping.
- At 1366×768 and 125% OS scaling, the packaged app rendered in a 1078×614 window capture. Metric cards reflowed from four columns to two and the workspace remained scrollable; no visible clipping was found in the observed top sections. This is a limited visual observation, not full workflow acceptance.
- Windows exposed only 100% and 125% in the normal scale selector. A physical 150% OS scaling check is **NOT AVAILABLE / NOT VERIFIED**. Qt offscreen UI tests separately passed at scale factors 1.25 and 1.5 (92 tests each); those do not substitute for physical DPI checks.
- The host offered 1366×768 but not 960×600 as a display mode. Existing Qt layout checks cover 960×600, but physical 960×600 window presentation remains **NOT VERIFIED**.
- Display settings were restored and confirmed at 1920×1080 and 100%. The application code/package was not changed by this observation. ACE DAO 16.0 and mapped SMB/UNC remain **WAIVED / NOT VERIFIED**, never PASS. PR #2 remains open and unmerged while these and other manual acceptance gates remain open.

## W303 — packaged Finder modes, file graph, and keyboard accessibility (2026-10-28)

- Revalidated the staged package with `tools/verify_release_package.py app/vector-release`: **PASS, 116 files / 76,330,196 bytes**. GUI and CLI hashes remain unchanged from W300.
- On the packaged GUI, the Finder added a disposable pair (`Dashboard.xlsx`, `Inventory.xlsx`) from `%TEMP%\pqm-vector-file-graph-20261028`. Quick / Map all selected files returned **1 exact internal edge**, 0 outside-scope or unresolvable references; the file graph showed both nodes and the lime directed edge. Clicking a node selected it and updated the selected-file label. Fit graph returned a panned view to its initial full-fit state.
- Quick / Find dependents with only `Inventory.xlsx` as target returned **1 reference** from `Dashboard.xlsx`. A deliberate no-consumer case (both workbooks selected as targets, excluding all candidate files) returned partial status with `Attempted 0 file(s)` and the `no_consumers` scope warning; no incorrect success or missing-reference claim was shown.
- Packaged keyboard smoke `python tools/run_packaged_keyboard_smoke.py --exe app/vector-release/PowerQueryManager/PowerQueryManager.exe --dialogs` exited **0**. Main-window Tab traversal reached **43** controls in order without opening Excel. Operation history, source inspector, query dependency graph, and both source-mapping tabs passed required control-cycle assertions; Right Arrow changed mapping tabs and Escape closed dialogs. Smoke used an isolated temporary user profile and temporary `.pq` source.
- Interactive packaged app was closed with Escape/Alt+F4; the smoke harness exited cleanly. No PowerQueryManager or Excel process remained. This additional pass does not close the physical 960×600 or 150% DPI gates, and does not verify every button/workflow. ACE DAO 16.0 and mapped SMB/UNC remain **WAIVED / NOT VERIFIED**. PR #2 remains open and unmerged.

## W304 — refreshed package query graph and inspector replay (2026-10-28)

- On the rebuilt package from app-code revision `a6b32e733fff3f48daf42cf1f90e1a93d840a5bf`, loaded disposable `Base.pq` and `Final.pq` (`let X = Base in X`). The query table reported `1. Base → 2. Final`; the query graph showed both nodes and `Final depends on Base`. The source inspector displayed the exact M text and correctly reported workbook comparison as unknown because no workbook was selected. Closed without saving; no workbook was opened or modified.
- Current staged package re-verification: **PASS, 116 files / 76,330,151 bytes**. Packaged CLI `--help` exited successfully. GUI SHA-256 `1BBB016C3C1FAD2062A23370BFC99397CB8D395A14BACE5C634D7D3E66A4B804`; CLI SHA-256 `6B30D61A87A2A164949C35EFB8BEEE2226B40344741CD64BA974365F752F0382`.
- This does not close other workflow or integration gates. Physical 960×600 and 150% DPI, broad workflow/keyboard coverage remain **OPEN / NOT VERIFIED**. ACE DAO 16.0 and mapped SMB/UNC remain **WAIVED / NOT VERIFIED**, never PASS. PR #2 remains open and unmerged.

## W305 — packaged history dialog accessibility and corruption recovery (2026-10-28)

- Ran `python tools/run_packaged_history_a11y_smoke.py --exe app/vector-release/PowerQueryManager/PowerQueryManager.exe --corrupt-history` against the verified refreshed package. **PASS:** Windows UI Automation found the history table, retention control, optional snapshot setting and history actions. The seeded corrupt SQLite bytes were preserved in a recoverable archive and the app reported the archive; the fresh history store initialized. The packaged manager exited 0 and all history data remained under the isolated temporary `LOCALAPPDATA` profile.
- This check covers the history dialog and corrupt-store recovery only; it does not prove the other remaining manual workflows or display gates. Physical 960×600 / 150% DPI and broad workflow/integration acceptance remain **OPEN / NOT VERIFIED**. ACE DAO 16.0 and mapped SMB/UNC remain **WAIVED / NOT VERIFIED**, never PASS. PR #2 remains open and unmerged.

## W306 — refreshed-package DAO 12.0 Finder diagnostic (2026-10-28)

- Ran `tools/run_dependency_windows_acceptance.py --dao-progid DAO.DBEngine.120 --cli-exe app/vector-release/PowerQueryManager/PowerQueryManagerCLI.exe --performance-repeats 1` on the refreshed package built from app-code revision `a6b32e733fff3f48daf42cf1f90e1a93d840a5bf`; harness/report revision was `0dfee363228b9bfce5215292e56a84586b7dd969`. This is a DAO 12.0 diagnostic, not ACE DAO 16.0 acceptance.
- **PASS for this scope:** disposable local fixture scan produced the expected 12 dependency edges; Quick and Deep each found 12 with no mode-exclusive edges, including the 31,200-cell formula workbook. Source hashes/sizes/timestamps were unchanged; no Access locks remained; DAO temp folders and fixtures were removed; Excel PID set returned to baseline. Packaged GUI exercised incoming/outgoing/both filters, 12-edge graph and verified 12-edge JSON/CSV exports; progress was observed and cancellation returned a canceled result after one file. Multi-target reverse scan matched six relationships.
- ACE DAO 16.0 remains **WAIVED / NOT VERIFIED**. Mapped SMB/UNC were not exercised here. This local DAO 12.0 run does not close the broader deployment matrix.

## W307 — refreshed-package existing-query import, backup, and cleanup (2026-10-28)

- Ran `tools/run_packaged_gui_import_smoke.py --exe app/vector-release/PowerQueryManager/PowerQueryManager.exe --replace-existing` on the current package, app-code revision `a6b32e733fff3f48daf42cf1f90e1a93d840a5bf`, Windows 10 build 26300 / Excel 16.0 build 20430.0, non-elevated. **PASS:** packaged UI Automation exposed the workbook, query grid, filter, replace policy, backup, import action and operation log controls. On a generated disposable workbook, the existing query was replaced in place, the exact new M formula was read back, the pre-import backup was verified, and an unrelated worksheet marker was preserved. Manager exited 0; owned Excel processes exited without forced termination. Post-run process inventory showed no Excel or PowerQueryManager process.
- This verifies the existing-query definition replacement path for this package and Excel build. It does not certify other Office builds, DAO 16.0, mapped SMB/UNC or all other remaining workbook/load/security cases. Physical 960×600 and 150% DPI remain **OPEN / NOT VERIFIED**; PR #2 remains open and unmerged.

## W308 — refreshed-package profile-backed mixed-source import (2026-10-28)

- Ran `tools/run_packaged_gui_import_smoke.py --exe app/vector-release/PowerQueryManager/PowerQueryManager.exe --mixed-import` on app-code revision `a6b32e733fff3f48daf42cf1f90e1a93d840a5bf`, Windows 10 build 26300 / Excel 16.0 build 20430.0, non-elevated. **PASS:** the versioned profile supplied three numbered `.m`/`.pq` sources; the UI review order matched the resulting Excel query order; exact formulas matched their source files; backup verification and read-only reopen passed. Packaged manager exited 0, owned Excel exited naturally (no forced termination), and the post-run process inventory showed no Excel or manager process.
- This covers the profile-backed mixed-import path on one Office build with disposable data. Other profile variants, Office builds and remaining workbook/security gates remain open. ACE DAO 16.0 and mapped SMB/UNC remain waived; physical 960×600 and 150% DPI remain **NOT VERIFIED**.

## W309 — refreshed-package all-skip import no-op (2026-10-28)

- Ran `tools/run_packaged_gui_import_smoke.py --exe app/vector-release/PowerQueryManager/PowerQueryManager.exe --skip-existing` on app-code revision `a6b32e733fff3f48daf42cf1f90e1a93d840a5bf`, Windows 10 build 26300 / Excel 16.0 build 20430.0, non-elevated. **PASS:** packaged UI reported the existing query was kept unchanged and no workbook changes were committed. The disposable `.xlsx` hash, timestamp, and existing formula matched before/after; no backup or stage was created. Manager exited 0 and owned Excel exited without forced termination; no Excel or manager process remained.
- Scope is the all-skip case only. Other Office builds and remaining workbook/security cases stay open; physical 960×600 and 150% DPI remain **NOT VERIFIED**. ACE DAO 16.0 and mapped SMB/UNC remain **WAIVED / NOT VERIFIED**.

## W310 — refreshed-package keyboard navigation (2026-10-28)

- Ran `tools/run_packaged_keyboard_smoke.py --exe app/vector-release/PowerQueryManager/PowerQueryManager.exe --dialogs` on the refreshed package built from app-code revision `a6b32e733fff3f48daf42cf1f90e1a93d840a5bf`; harness revision was feature branch `54b97d1a91ef8dfd230c768d3b078e7fe3686b28`. **PASS:** main-window Tab traversal completed and stayed in the manager; 43 controls were visited in order, including workbook path, Add files, query table, existing-query policy, backup and import action. Operation-history, inspector, query graph and source-mapping dialogs completed their required Tab cycles; mapping Right Arrow switched categories and Escape closed dialogs. No workbook or Excel was opened.
- This updates keyboard smoke evidence for the refreshed package. It does not certify screen-reader usage or every modal. Physical 960×600 and 150% DPI remain **NOT VERIFIED**; broader non-waived manual acceptance remains open. ACE DAO 16.0 and mapped SMB/UNC remain **WAIVED / NOT VERIFIED**.

## W311 — refreshed-package history restore and definition-only import (2026-10-28)

- Ran `tools/run_packaged_gui_import_smoke.py --exe app/vector-release/PowerQueryManager/PowerQueryManager.exe --history-restore` on app-code revision `a6b32e733fff3f48daf42cf1f90e1a93d840a5bf`, Windows 10 build 26300 / Excel 16.0 build 20430.0, non-elevated. **PASS:** packaged UI Automation found named inspector, dependency-graph and profile-mapping controls; mapping changes were cancelled. History restore returned the disposable workbook to the selected snapshot hash and preserved an exact pre-restore backup.
- The same harness continued through packaged definition-only import: exact formula readback, verified pre-import backup, read-only reopen, manager exit 0 and natural owned-Excel cleanup. No Excel or manager process remained after the run.
- Scope is this history restore and definition-only import scenario on one Excel build. Physical 960×600 and 150% DPI and broader remaining workbook/security cases remain **NOT VERIFIED**. ACE DAO 16.0 and mapped SMB/UNC remain **WAIVED / NOT VERIFIED**.

## W312 — refreshed-package parameter edit and smoke harness correction (2026-10-28)

- Ran `tools/run_packaged_gui_import_smoke.py --exe app/vector-release/PowerQueryManager/PowerQueryManager.exe --parameter-edit` against the package built from app-code revision `a6b32e733fff3f48daf42cf1f90e1a93d840a5bf`, using harness fix commit `319c1d5dbc05d4eb2492df3cb77380fdb8aa5f19`, Windows 10 build 26300 / Excel 16.0 build 20430.0, non-elevated. **PASS:** edited the declared text parameter through the inspector, saved a separate `.m` copy with exact expected bytes, preserved the original source bytes, imported it, verified the exact formula and pre-import backup, passed read-only reopen, and cleaned up Excel without forced termination; manager exited 0.
- The first attempt timed out because the smoke used raw Win32 `SendMessage` clicks to activate the inspector. The analogous packaged-dialog smoke uses UI Automation `invoke()` successfully; changing this harness action to UIA `invoke()` made the replay pass. Assertions were unchanged. This was a smoke-harness activation defect, not an application failure.
- No Excel or PowerQueryManager process remained after the passing run. Scope is parameter editing with a synthetic source and one Excel build. Physical 960×600 and 150% DPI and broader workbook/security cases remain **NOT VERIFIED**; ACE DAO 16.0 and mapped SMB/UNC remain **WAIVED / NOT VERIFIED**.

## W313 — refreshed-package Finder Deep scan and file graph (2026-10-28)

- Launched `app/vector-release/PowerQueryManager/PowerQueryManager.exe` and used Ctrl+K to open Dependency Finder. In the GUI, selected **Map all selected files and their dependencies** and **Deep — also worksheet formulas and names**, then added the two disposable fixtures from `%TEMP%\pqm-vector-file-graph-20261028`: `Dashboard.xlsx` and `Inventory.xlsx`.
- **PASS for this scenario:** packaged GUI reported `Read-only scan finished`, 2 selected files, 1 internal dependency edge, and 0 outside-scope or unresolvable references. The result row identified `Dashboard.xlsx` as depending on `Inventory.xlsx`, type `Excel connection`, query/connection `SyntheticSource`, confidence `exact_path`. The opened visual graph showed the lime arrow from Dashboard to Inventory; its accessible text read `Dashboard.xlsx depends on Inventory.xlsx [exact_path]`.
- This was a read-only scan of synthetic disposable workbooks and did not launch Excel. It covers this Finder Deep/general-map path and graph view only. It does not establish broad Quick/Deep parity, every graph interaction, Excel/Access integration, package shutdown or remaining workflows. ACE DAO 16.0 and mapped SMB/UNC remain **WAIVED / NOT VERIFIED**; physical 960×600 and 150% DPI remain **NOT VERIFIED**. Tested app package remains based on app-code commit `a6b32e733fff3f48daf42cf1f90e1a93d840a5bf`; repository evidence update is on the feature branch and PR #2 remains unmerged.

## W314 — exact-head Windows rebuild and package integrity (2026-10-28)

- On feature branch commit `36238ce97adf03dbf348ebdedd0b83200c0426d5`, ran `app/build.ps1` on Windows 10 build 26300 with the repository's Python 3.11.9 / PyInstaller 6.13.0 environment. **PASS:** the full suite completed with **543 passed in 48.97 s** and both onedir executables rebuilt successfully.
- Staged the fresh `app/dist/PowerQueryManager` output into `app/vector-release` using `tools/package_release.py`; `tools/verify_release_package.py app/vector-release` passed at **116 files / 76,330,151 bytes**. The staged CLI `--help` exited 0. GUI SHA-256: `1BBB016C3C1FAD2062A23370BFC99397CB8D395A14BACE5C634D7D3E66A4B804`; CLI SHA-256: `6B30D61A87A2A164949C35EFB8BEEE2226B40344741CD64BA974365F752F0382`.
- This confirms the Windows build and package integrity at this repository head. It does not close the remaining physical display/DPI and broad manual acceptance gates. ACE DAO 16.0 and mapped SMB/UNC remain **WAIVED / NOT VERIFIED**; PR #2 stays open and unmerged.

## W315 — live packaged workspace resize attempt (2026-10-28)

- Relaunched the freshly rebuilt, hash-verified GUI package and inspected the main workspace at **1182×732**. Attempted native window-border drags toward **1000×650** and **960×600** using the live window screenshot as coordinate reference. Neither attempt changed the captured window dimensions. The app closed afterward and process inventory showed no remaining PowerQueryManager process.
- Source inspection shows the main window declares a 960×600 minimum and no fixed-size or maximized startup call; the full `app/build.ps1` suite passed its Qt resize/layout assertions at 960×600. Those facts do not substitute for a physical live-window check. Both physical target dimensions remain **NOT VERIFIED**; investigate why the desktop resize action has no effect before claiming that gate. Physical 150% DPI and broad manual gates remain open; ACE DAO 16.0 and mapped SMB/UNC remain **WAIVED / NOT VERIFIED**.

## W316 — compact-workspace dialogs and file graph (2026-10-28)

- Exact tested application source commit: `ecfa2b9b35ff79ff33e01e2d16db742ef267f01a`. `app/build.ps1` on Windows completed **544 passed in 41.80 s** and rebuilt the packaged GUI/CLI. `tools/verify_release_package.py app/vector-release` passed with **116 files / 76,331,125 bytes**; packaged CLI `--help` exited 0. GUI SHA-256: `601718AF835F6021AED98254FAEDDED3B2F897A9C8420052F438760EAEDA7819`; CLI SHA-256: `6B30D61A87A2A164949C35EFB8BEEE2226B40344741CD64BA974365F752F0382`.
- On the rebuilt packaged GUI, used the Windows system Size keyboard command to set the main window client area to **960×600**. Main workspace showed its two-column metric cards and vertical scrolling without horizontal overflow. Dependency Finder opened within the same parent bounds. Opened the file dependency graph; its graph area, graph controls and accessible edge summary remained visible at the same window size.
- Repeated the synthetic read-only Deep general-map scenario using disposable `Dashboard.xlsx` and `Inventory.xlsx`. **PASS for this scoped scenario:** 2 files, 1 exact-path Excel-connection edge, 0 unresolved/out-of-scope references; visible lime arrow and accessible summary agreed (`Dashboard.xlsx depends on Inventory.xlsx [exact_path]`). No Excel process was started.
- This verifies only the 960×600 main/Finder/file-graph path and the stated scan scenario. 1366×768, physical 125%/150% DPI, query graph, inspector, dialogs beyond the exercised set, broad workflows, Excel/Access integration and screen-reader acceptance remain open/not verified. ACE DAO 16.0 and mapped SMB/UNC remain **WAIVED / NOT VERIFIED**. PR #2 remains open and unmerged.

## W317 — exact-package keyboard and modal traversal replay (2026-10-28)

- Re-ran `tools/run_packaged_keyboard_smoke.py --exe app/vector-release/PowerQueryManager/PowerQueryManager.exe --dialogs` against the package built from application source commit `ecfa2b9b35ff79ff33e01e2d16db742ef267f01a` (current QA/docs head before this entry: `e407a7c4916915de4030d1fbc3fbe0e98c075565`). **PASS:** main-window Tab traversal completed a full in-window cycle; 43 focusable controls were visited in order, including workbook path, Add files, query table, existing-query policy, backup and import action. History, inspector, query graph and source-mapping dialog Tab cycles reached their expected controls; source-mapping Right Arrow changed category and Escape closed dialogs. No workbook or Excel instance opened.
- Process inventory after the run showed no PowerQueryManager or Excel process. This revalidates keyboard reachability on the latest rebuilt package and includes graph/inspector sizing constructor paths, but does not establish visual layout at 1366×768 or 125%/150% DPI, screen-reader accessibility, or all buttons/workflows. ACE DAO 16.0 and mapped SMB/UNC remain **WAIVED / NOT VERIFIED**; PR #2 stays open and unmerged.

## W318 — latest-package profile-backed mixed-source import (2026-10-28)

- Ran `tools/run_packaged_gui_import_smoke.py --exe app/vector-release/PowerQueryManager/PowerQueryManager.exe --mixed-import` against application source commit `ecfa2b9b35ff79ff33e01e2d16db742ef267f01a` (harness/docs branch head `0194d6e03663f166b406255d0c2a4af94b93959e`) on Windows with Excel 16.0 build 20430.0 and a disposable workbook. **PASS:** imported three numbered `.m`/`.pq` sources through a versioned profile; review order matched Excel query order; exact formulas matched source; the pre-import backup was verified; read-only reopen verification passed; manager exited 0; owned Excel processes exited naturally with no forced termination.
- This extends mixed profile/import evidence to the latest GUI package. It does not qualify other Excel builds, Access, load-mode variants, refresh, 150% DPI, or all remaining workbook/security cases. ACE DAO 16.0 and mapped SMB/UNC remain **WAIVED / NOT VERIFIED**.

## W319 — latest-package all-skip import no-op (2026-10-28)

- Ran `tools/run_packaged_gui_import_smoke.py --exe app/vector-release/PowerQueryManager/PowerQueryManager.exe --skip-existing` on the package built from application source commit `ecfa2b9b35ff79ff33e01e2d16db742ef267f01a`, on Excel 16.0 build 20430.0 with a disposable `.xlsx`. **PASS:** existing query was kept unchanged, the UI reported no workbook changes, exact workbook hash/timestamp and formula were preserved, and no backup/stage was created. Manager exited 0; read-only verification and Excel cleanup completed with no forced termination.
- This verifies the all-skip branch for this disposable workbook and Office build only. Other workbook types, Access, load/refresh modes and remaining security workflows remain separate acceptance gates; ACE DAO 16.0 and mapped SMB/UNC remain **WAIVED / NOT VERIFIED**.

## W320 — latest-package DAO 12.0 local dependency acceptance (2026-10-28)

- Ran `tools/run_dependency_windows_acceptance.py --cli-exe app/vector-release/PowerQueryManager/PowerQueryManagerCLI.exe --dao-progid DAO.DBEngine.120 --performance-repeats 1 --report evidence/W320-dao12-acceptance.json` on Windows 10 build 26300 / Excel 16.0 build 20430.0, using 10 generated disposable files (4 Access, 4 Excel). The harness status is **PASS** on package/source revision `11831c8782a2040f800050240ffcdf51c920210e`.
- Quick and Deep each matched all **12 expected dependency edges**, with zero Quick-only or Deep-only edges; packaged CLI and scanner API were at parity. GUI filter counts matched for incoming/outgoing/both across four targets; the file graph contained 12 edges; JSON and CSV exports each verified all 12 edges. GUI progress was observed and cancellation returned the expected cancelled result.
- Read-only integrity passed: all 10 generated source files retained identical SHA-256, size and modification time; no lock files appeared; Excel process set returned to its initial empty set; fixture cleanup passed; no forced Excel termination or cleanup error occurred. One expected advisory issue and two unresolved references were reported and retained in the raw JSON. One-run scanner timings were Quick 3.08 s and Deep 3.40 s for this fixture only.
- The report is preserved at `evidence/W320-dao12-acceptance.json`. After keeping the report out of the staged release folder, `tools/verify_release_package.py app/vector-release` again passed at **116 files / 76,331,125 bytes**; GUI/CLI hashes remain `601718AF835F6021AED98254FAEDDED3B2F897A9C8420052F438760EAEDA7819` / `6B30D61A87A2A164949C35EFB8BEEE2226B40344741CD64BA974365F752F0382`.
- This covers local synthetic fixtures only. ACE DAO 16.0 and mapped SMB/UNC remain **WAIVED / NOT VERIFIED**. No network roots were exercised.

## W321 — latest-package live query graph and inspector review (2026-10-28)

- Launched the package GUI (SHA-256 `601718AF835F6021AED98254FAEDDED3B2F897A9C8420052F438760EAEDA7819`) built from source commit `ecfa2b9b35ff79ff33e01e2d16db742ef267f01a`. At a captured 1182×732 window, used the workspace scroll area to add two synthetic `.m` files from `%TEMP%` without selecting a workbook or launching Excel.
- **PASS for scoped visual review:** the query table showed both source rows and the detected dependency; the query graph rendered `Final → Source` with the dependency direction matching its text summary. The inspector rendered the selected source M pane, workbook comparison pane, and exact line-diff pane. Since no workbook was selected, the workbook formula was explicitly unavailable and comparison remained unknown, as expected. The graph and inspector fit within the displayed window and their controls were visible. App closed normally; no PowerQueryManager, Excel, or Access processes remained.
- This pass is at 1182×732 and covers the live query graph and inspector only. It does not establish the 960×600 or 1366×768/125%/150% layouts, all graph interactions, or remaining workflows. ACE DAO 16.0 and mapped SMB/UNC remain **WAIVED / NOT VERIFIED**; PR #2 remains open and unmerged.
