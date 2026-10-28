"""Stage the minimal Windows onedir package and create a SHA256SUMS manifest."""
from __future__ import annotations

import hashlib
import argparse
import json
import re
import os
from pathlib import Path
import shutil
import tempfile
import uuid

from package_fs import replace_directory


ROOT = Path(__file__).resolve().parents[1]
PARSER = argparse.ArgumentParser(description=__doc__)
PARSER.add_argument("--source", type=Path,
                    help="PyInstaller onedir folder to package")
PARSER.add_argument("--destination", type=Path, default=ROOT / "dist",
                    help="release package directory (must be inside the handoff root)")
PARSER.add_argument("--docs-only", action="store_true",
                    help="refresh the manual, summary, and SHA256SUMS without touching the packaged executables")
ARGS = PARSER.parse_args()
build_source_sha = os.environ.get("PQM_BUILD_SOURCE_SHA", "").strip().lower()
if build_source_sha and not re.fullmatch(r"[0-9a-f]{40}", build_source_sha):
    PARSER.error("PQM_BUILD_SOURCE_SHA must be a full 40-digit commit SHA.")
FINAL_DIST = ARGS.destination.resolve()
DIST = FINAL_DIST
TARGET = (DIST / "PowerQueryManager").resolve()

if ROOT not in FINAL_DIST.parents or TARGET.parent != DIST:
    raise SystemExit("Package paths do not resolve inside the handoff root.")

if ARGS.docs_only:
    if ARGS.source is not None:
        PARSER.error("--source cannot be used with --docs-only; documentation refresh reads the existing destination package.")
    DIST.mkdir(parents=True, exist_ok=True)
    if not TARGET.is_dir():
        raise SystemExit("Cannot refresh package documentation: packaged PowerQueryManager folder is missing.")
else:
    SOURCE = (ARGS.source or ROOT / "app" / "dist" / "PowerQueryManager").resolve()
    if ROOT not in SOURCE.parents:
        raise SystemExit("Package paths do not resolve inside the handoff root.")
    if (not (SOURCE / "PowerQueryManager.exe").is_file()
            or not (SOURCE / "PowerQueryManagerCLI.exe").is_file()
            or not (SOURCE / "_internal").is_dir()):
        raise SystemExit("Build the PyInstaller onedir package first (app/dist/PowerQueryManager).")
    DIST = Path(tempfile.mkdtemp(prefix=f".{FINAL_DIST.name}.stage-", dir=FINAL_DIST.parent)).resolve()
    TARGET = DIST / "PowerQueryManager"
    shutil.copytree(SOURCE, TARGET)

(DIST / "README.txt").write_text(
    """Power Query Manager — implementation preview

Extract the complete PowerQueryManager folder and run PowerQueryManager.exe. Keep _internal beside the executable. This package contains no source, tests, fixtures, caches, or virtual environment.

Security: before any COM open, Windows Zone.Identifier metadata is checked; Internet/Restricted zone marks and uninterpretable zone streams are refused without removing the mark or changing Office trust settings. This conservative preflight does not certify Protected View UI behavior. Programmatic opens set UpdateLinks=0 and force-disable VBA macros, then restore the owned Excel instance setting during cleanup. Excel 4.0/XLM macro prompts are an Office exception that this app does not approve or bypass; XLM workbook behavior is unverified. A disposable VBA-bearing `.xlsm` with a harmless Workbook_Open event passed GUI import and staged synchronous/background refresh on Excel 16.0 build 20430.0; the event source was preserved and the event did not run. Other VBA project forms and Office builds remain unverified. See the Microsoft AutomationSecurity documentation.

Workbook conflicts: Windows Restart Manager and Excel's Running Object Table check whether another process or Excel instance has the workbook open before import, refresh, and snapshot overwrite; import and refresh repeat the file-use check before commit. Staged writes reject paths where the workbook, stage, or any parent is a symbolic link or Windows reparse point. Recovery discovery and copy-out apply the same path checks to workbook, stage, destination, and parents, and never resolve through a reparse point. The app never closes another application. If Windows cannot verify use, the operation stops. W187 verified both Restart Manager detection of a separate process and exact-path Excel ROT detection for a read-only workbook open without a `~$` lock file on Excel 16.0 build 20430.0. A small race remains if a different app opens the file after the final check.

Workflow: select a workbook; inspect its queries read-only; add .m/.pq sources; review/edit order, names, paths, enabled state, and load destination; inspect formula diffs; edit source-path and M-source environment mappings, then save/load a versioned profile; confirm Replace or Skip; import changes into a staged copy. Definition-only and connection-only import do not evaluate M. Worksheet and Data Model loads require an explicit confirmation because they evaluate the selected query. After inspecting workbook queries, use Refresh selected/all to explicitly refresh existing Power Query connections. These operations warn that external data may be contacted, offer refresh mode and timeout, create a verified backup by default, and commit a staged copy only after save, read-only reopen verification, and concurrent-change checks succeed.

Crash recovery: the candidate list shows age, size, modified time, and a preliminary ZIP/Open XML structure check. This triage does not prove an import committed or validate query contents; Excel still opens a selected candidate read-only before copy-out. Recovery never replaces the selected workbook or deletes its stage file. If Windows prevents cleanup after failure, the operation log reports the retained stage filename so you can inspect or save a recovery copy.

Workbook privacy preservation: before staged import, DataMashup permission metadata is checked for FirewallEnabled and WorkbookGroupType; a changed or unverifiable DPAPI-bound permission record stops commit. Default true/nil was preserved in one Excel 16.0 build 20430.0 smoke. Non-default flags and cross-user bindings remain unverified. The app does not inspect or change source credentials or privacy-level assignments.

Query packages: use Export selected/all in the query inventory to create a `.pqm.zip`. New manifest-v2 packages preserve row order and exact formula hashes and include conservative dependency/ambiguity/cycle metadata. Version-1 packages remain importable and receive dependency metadata in memory. **Compare package versions…** validates two packages and previews added, modified, unchanged, removed, or conflicting query names with exact M diffs; removed queries are never deleted from a workbook. ZIP import rejects compressed archives above 32 MiB and expanded content above 25 MiB, checks size and parses through the same open file handle, then extracts into an isolated temporary folder and adds the sources to the editable grid for review before any workbook operation.

Dependency graph: select Dependency graph… to see confident query-to-query references, disabled dependencies, and cycles. Ambiguous M expressions are warnings without guessed arrows. A read-only text summary is included, and the graph never changes import order.

Manager controls: Undo/Redo (Ctrl+Z/Ctrl+Y) covers editable table fields, drag order, and selected-row removal; bulk enable/order/destination changes are grouped as one action. Dark theme is available from the title row. The window starts within the available display size. Automated UI tests passed at Qt scale factors 1.25 and 1.5 (W47); OS display scaling, screen-reader behavior, and physical laptop acceptance remain pending.

Query inspector: use the title-row toggle or View → Query inspector to show a docked, read-only source/workbook M preview for the selected query. The dock is hidden by default, reports absent queries and unavailable formulas clearly, limits each preview to 100,000 characters, and links to the full compare/editor.

Failure diagnostics: `%LOCALAPPDATA%\PowerQueryManager\logs\errors.jsonl` stores rotated local diagnostics (1 MiB, three backups). Each record includes timestamp, operation, exception type and fixed category only. Raw exception text, workbook paths, M formulas and credentials are omitted. Diagnostic log files, rotation targets and parent directories containing symbolic links or Windows reparse points are refused; logging is skipped if the user-local destination is redirected or unwritable.

Common failures are categorized in the UI and CLI with a next step: Excel timeout, unsupported capability, changed/locked workbook, backup or verification failure, query conflict, source drift, or invalid input. Detailed operation output stays visible to the user, including failed, rolled-back, and not-attempted query outcomes for imports; the local diagnostic file stores no raw error text.

Optional CLI: PowerQueryManagerCLI.exe inspect --workbook target.xlsx reads inventory; validate --profile deployment.json checks a profile and enabled source files offline; export --workbook target.xlsx --output queries.pqm.zip --package exports a verified query package; import --dry-run --package queries.pqm.zip --workbook target.xlsx compares read-only, while import --apply ... requires typing IMPORT and defaults to skip/definition-only; plan --dry-run --profile deployment.json --workbook one.xlsx --workbook two.xlsm creates read-only plans; refresh --apply --all --workbook target.xlsx requires an interactive terminal and typing REFRESH after a source-access warning, stages the workbook, backs it up, and reports per-query results. Use --query NAME to select queries, --synchronous for Excel sync mode, and --timeout SECONDS (1-7200). For deployment use PowerQueryManagerCLI.exe deploy --profile deployment.json --workbook one.xlsx --workbook two.xlsm --apply; it preflights every target, shows the exact actions, then requires typing DEPLOY. Add --output report.json for a new per-target report file. Workbooks are staged and committed independently; the batch is not atomic and partial failures are reported. For profile source-path root rules, add --apply-path-mappings; inspect the exact preview and type APPLY PATH MAPPINGS. For M-code environment rules, add --apply-environment-mappings; inspect the exact unified diff and type APPLY M-CODE MAPPINGS. When both are enabled, path mapping is previewed first and the M diff is computed from the mapped source. Only hash-checked temporary source copies are used.
Portable deployment release: PowerQueryManagerCLI.exe bundle --profile deployment.json --output release.pqmdeploy embeds a versioned profile and exact source bytes with SHA-256 validation; source-path mappings require preview and typed approval. Use the resulting archive with plan --profile release.pqmdeploy or deploy --profile release.pqmdeploy, and deploy requires the exact DEPLOY phrase, entered at the prompt or supplied deliberately with --confirm DEPLOY. Repeat --confirm for each separately reviewed mapping phrase; a confirmation never substitutes for a different phrase; a closed approval input is treated as a decline. Explicit refresh accepts --confirm REFRESH; query-package import accepts --confirm IMPORT. The creator's workbook path is replaced with a placeholder. CLI exit codes: 0 complete, 2 validation/output error, 3 unsupported or declined action, 4 Excel/COM preflight error, 5 failed or partial operation.

Local history: select History… in the operation log to review comparison, import, refresh, export, profile, CLI plan/deploy, and restore event metadata. CLI events contain action/status and aggregate counts only; query names, formulas, source paths, and workbook paths are excluded. Select two snapshots and choose Compare selected snapshots… to inspect query additions/removals, exact M formula changes, and known load destination changes. Both snapshots are hash-verified and opened read-only on a COM worker thread; comparison never refreshes external data. It is stored in %LOCALAPPDATA%\\PowerQueryManager\\history.sqlite3, refuses symbolic links/reparse points in its path, omits formula/code fields, and redacts recognized secrets, and provides retention and deletion controls. Snapshot capture is OFF by default. When enabled, full .xlsx/.xlsm workbooks are stored in a local history_snapshots folder before import and refresh; copied contents are flushed before exclusive hard-link publication (an existing UUID destination is never overwritten), and the snapshot index plus audit event commit in one SQLite transaction. An immediate SQLite writer lock spans publication and indexing so concurrent startup cannot recover a live file. Expired snapshots are pruned before lookup/restore; cleanup only removes stored paths that resolve inside the history snapshot directory. Restores are hash checked, require Excel closed, create a verified backup before overwrite, and default to a new file when the original is unavailable. On Windows, restore publication and verification rollback use the same existing-target ReplaceFileW safety helper as import and refresh; if the selected path disappears after review, the app refuses to recreate it and retains/reports the verified backup. Workbook content may be sensitive. If Windows or the process stops after a snapshot file is published but before its history transaction commits, startup moves the unindexed manager-named file to `%LOCALAPPDATA%\\PowerQueryManager\\history.sqlite3.recovered` without deleting it and reports the folder for manual review; recovered files are not automatically treated as valid snapshots. A crash during hard-link publication can leave a recent `.pqm-snapshot-*` alias; startup preserves recent in-progress names while moving the published UUID file for manual review. If the history database is corrupt or has an incompatible schema, the original SQLite file and any WAL/journal sidecars are preserved as timestamped `history.corrupt-*` files before a fresh store is started; the operation log reports the archive path. Damaged entries are not automatically merged. If the recovery error identifies a stale `.recovery.lock`, close every Power Query Manager instance first; after confirming none is recovering, remove only the named lock file and retry. Never remove or move the history database or its `-wal`, `-shm`, or `-journal` sidecars as lock cleanup.

The query inspector can edit only standalone declared Text, Number, and Logical literal parameters. It changes just the literal and uses the separate source-copy save workflow; comments, UTF-8 BOM, and detected line endings are preserved. Unsupported or ambiguous expressions are disabled in this parameter control.

The inspector includes line-numbered source/workbook panes, basic M syntax colors, added/removed diff colors, and search in the source, workbook, or diff. Syntax coloring is lexical only. Source edits stay in the dialog buffer until Save edited source copy is used; a later import operation is required to change Excel.

Interactive profiles may opt into M-code environment mappings. The UI and CLI preview exact diffs before approval and store edits only in temporary copies. Rewrites are limited to recognized literal source paths in File.Contents, Folder.Files, Folder.Contents, SharePoint.Files, SharePoint.Contents, and Web.Contents; dynamic expressions stay unchanged with warnings. No full M parser is claimed.

Definition only is the default. Connection only, new worksheet, existing worksheet, and Data Model modes are enabled only after their temporary-workbook probe passes for the active Office version/build. Probe queries use inline static M only. App-level create/repeat-import, changed-formula update, read-only reopen, output-value and owned-process cleanup smokes passed for the tested modes on Excel 16.0 builds 20326.0 and 20430.0; the current-source load matrix passed connection-only, new worksheet, existing worksheet, and Data Model on Excel 16.0 build 20430.0 in W129; the packaged GUI Data Model path also passed exact formula, model membership/row-count, backup, and process-cleanup checks. W129 additionally verifies exact DispatchEx PID cleanup after capability-probe Quit errors and fails closed when cleanup is uncertain. Background and synchronous refresh were tested with inline static M; W44 repeats the background path after staged saves and final verification opens were put behind the owned-process timeout wrapper. W46 rebuilt the onedir package after stage-cleanup warnings were added and passed packaged GUI definition-only import on Excel 16.0 build 20430.0 with exact formula readback, verified backup, preserved worksheet marker, clean manager exit and owned-process cleanup. W49 bounds capability-probe evaluation, save, and verification-open calls to 60 seconds per call and disables unverified load modes on timeout. W51 rejects symbolic-link workbook paths at write preflight, stage creation, and final commit. W52 checks the workbook and every parent path for Windows reparse-point attributes at those boundaries; a real disposable junction was rejected during preflight, stage creation and commit, preserving the workbook and staged bytes. W53 repeats the reparse-point and sibling checks at the final practical boundary before replace; this narrows but does not eliminate the OS-level path-swap race. W54 applies reparse checks to recovery discovery and copy-out; a real junctioned destination was rejected before publishing into its target. W55 repeats recovery path checks immediately before exclusive publication; a simulated late path change is refused. OS-level swaps after the last check remain possible. The packaged GUI definition-only smoke passed under a non-elevated user token; broader filesystem-policy combinations remain unverified. The timeout helper's controlled delay and open-password prompt smokes are not certification of arbitrary blocked COM or M evaluation. External connectors, other Office builds, VBA-bearing `.xlsm` workbooks, arbitrary blocked COM timeouts and evaluated packaged GUI loads remain unverified. Release gates are incomplete; use disposable workbooks.
""",
    encoding="utf-8",
)
(DIST / "RELEASE_SUMMARY.txt").write_text(
    """Power Query Manager implementation preview spanning v1.3.1 stability through v2.0.0 advanced tooling

Windows onedir package built from current app source with PyInstaller 6.13.0, Python 3.11.9, and PySide6 6.8.3. This is not a production release. See evidence/releases/v*.md in the source handoff for release gates.

The app supports .m/.pq import, staging and verified backups, manager inspection/export/diff/profiles, conservative declared-literal parameter editing, opt-in previewed M-code environment mapping, probe-gated definition-only/connection/worksheet/Data Model modes, explicit selected/all refresh with background/synchronous modes and timeout, redacted local metadata history with read-only two-snapshot query/M/load comparison, read-only query-package update preview, and a CLI for offline profile validation, read-only inspection/planning, and separately staged multi-workbook deployment from either a local profile or a portable hash-verified release bundle. Static-M load/import/reimport paths passed on Excel 16.0 builds 20326.0 and 20430.0; W129 repeats current-source connection-only, new worksheet, existing worksheet, and Data Model modes on build 20430.0, and verifies packaged GUI Data Model membership and row count. W44 verifies staged save and final read-only verification opens use bounded calls. W46 packages the retained-stage warning and passes GUI definition-only import on Excel 16.0 build 20430.0. W47 records automated Qt UI checks at 125% and 150% scale factors. W49 adds process-aware 60-second bounds around capability-probe evaluation/save/verification calls and disables modes after timeout. W51 rejects symbolic-link workbook paths at write preflight, staging and commit. W52 rejects Windows reparse points in the workbook, stage, or any parent path; an actual disposable junction failed at preflight, stage creation and commit. W53 repeats path checks immediately before replace to reduce (not eliminate) swap exposure. W54 applies the same checks to recovery discovery and copy-out; a real junctioned output parent was rejected. W55 repeats recovery path checks before exclusive publication. W56 verifies related report formulas and names remain attached to replaced worksheet query output, while preserving unrelated workbook objects; one lingering smoke-owned Excel process required exact-PID cleanup after graceful Quit. W57 verifies read-only snapshot comparison across definition-only/connection-only and worksheet/Data Model load states on Excel 16.0 build 20430.0; worksheet and model snapshots use equivalent separate workbooks because mode conversion is fail-closed. W58 adds fake-COM regression coverage proving such a refused conversion preserves the existing formula and connection. W59 repeats packaged GUI definition-only import under PCRN\\Robert with WindowsPrincipal.IsInRole(Administrator)=False; broader ACL and policy combinations remain open. W60 validates DPAPI-bound DataMashup permission metadata and preserves default FirewallEnabled=true/WorkbookGroupType=nil on Excel 16.0 build 20430.0; non-default flags and cross-user bindings remain unverified. The packaged GUI import smoke passed under a non-elevated user token; broader filesystem-policy combinations remain untested. W61 passed staged create/repeat-replace/read-only verification for a disposable macro-enabled .xlsm container on Excel 16.0 build 20430.0; it contained no VBA project and does not certify macro preservation or non-execution. W62 passed the packaged GUI definition-only import smoke for disposable .xlsm and .xlsx containers on the same build under the non-admin token; neither fixture had a VBA project. Blocked COM/M-evaluation timeouts, external connectors, other Office builds, screen-reader user testing, OS scaling, and other VBA project forms remain unverified. W68 checked the packaged history dialog names/roles through Windows UI Automation in an isolated temporary user-data directory. W69 records completed read-only snapshot comparisons by before/after snapshot IDs only; formula, path and diff content stay out of history. W70 fixes before/after ordering to follow recorded UTC time rather than newest-first table order. W64 verified a harmless Workbook_Open event was preserved and did not run during packaged GUI import; W67 repeated this for staged synchronous/background refresh. W71 repeated the packaged VBA-bearing .xlsm GUI import on Excel 16.0 build 20430.0: exact formula and backup verification, VBA source preserved, event marker unchanged, manager exit 0, and smoke-owned Excel cleanup. W72 process hard-exit tests prove complete-file outcomes immediately before and after atomic replacement; they do not simulate Excel COM termination or power loss. W73 packaged GUI all-skip on a read-only workbook on Excel 16.0 build 20430.0 preserves hash, timestamp, existing formula, marker, and read-only attribute, creates no backup/stage, reports no changes, and cleans up its own processes. Mixed skip/create batches and other Office builds remain unverified. Use disposable workbooks; this is not a production release. W74 passed packaged GUI/profile import of three numbered mixed .m/.pq sources on Excel 16.0 build 20430.0: review order and Excel query collection matched numeric prefixes, exact source formulas read back, the original-hash backup verified, and smoke-owned processes exited. Other Office builds and broader workbook fixtures remain unverified. W75 passed packaged GUI replacement of an existing definition-only query on the same build, with exact changed-formula readback, verified original-hash backup and unrelated worksheet marker preservation; loaded table/link associations remain open. W76 fixed capability gating that rejected the preserve-existing UI/profile policy: the policy now remains available while destination modes stay probe-gated. Packaged GUI replacement preserved a connection-only load and exact formula on Excel 16.0 build 20430.0; worksheet/Data Model preservation and other Office builds remain open. W77 repeated the controlled owned-Excel timeout cleanup smoke on Excel 16.0 build 20430.0; arbitrary blocked COM and M-evaluation timeout remain unverified. Recent verification W78-W119 adds packaged existing/new worksheet and connection-only GUI acceptance, safe Unicode load-name reimport, fail-closed Excel lookup errors, a timeout-completion race fix, and disposable table/report/chart/shape/local PivotTable preservation checks and repeated multi-process local-history writer/snapshot plus synchronized corrupt-store recovery checks on Windows 10 build 26300 / Excel 16.0 build 20430.0. The current full portable suite passed 345 tests; the onedir GUI/CLI rebuilt; controlled timeout cleanup, packaged query replacement and connection-only GUI import passed. W113 additionally verifies a full packaged keyboard focus cycle across 34 main-window controls. W114 fixes disabled profile rows with missing sources blocking multi-workbook deployment; full suite 338 passed and real two-workbook apply/repeat smoke passed. W116 prevents false `applied` batch reports: one Created/Replaced/Skipped result is required per enabled query; rollback/not-attempted and incomplete/duplicate/unknown results fail the target. Deployment suite 15 passed; full suite 345 passed; onedir rebuilt. W118 adds portable `.pqmdeploy` profile/source bundles with per-source SHA-256, safe extraction, source-path mapping approval, and CLI plan/deploy support; focused bundle/CLI/manager tests passed (68), and a live two-workbook bundle apply/repeat smoke passed on Excel 16.0 build 20430.0. W119 adds explicit repeated `--confirm` phrases for non-interactive CLI operation; wrong operation phrases do not authorize deployment. W120 verifies packaged CLI background and synchronous static-M refresh with `--confirm REFRESH`, hash-matched backups and no external source on Excel 16.0 build 20430.0. W121 verifies packaged CLI live query-package export/import, exact Unicode-name/formula readback, source-hash preservation, target backup and owned-process cleanup. W123 verifies a local PivotTable slicer and saved Keep/Drop selection survive production staged definition-only import on Excel 16.0 build 20430.0; broader slicer/timeline and Office-build coverage remain open. W124 corrects case-sensitive bounded M reference matching; W125 injects a final transaction-commit failure and verifies all changed query rows are reported rolled back, original workbook bytes stay unchanged, and the failed stage is cleaned up. W126 fixes package comparison history so only added/modified/unchanged/removed/conflict totals are retained; formulas and source paths remain excluded. W128 adds bounded successful-import created/replaced/skipped counts and profile path/environment mapping counts to local audit history, without recording query names, formulas, or source paths. The W129 full Windows build passed 363 tests in 18.81s; packaged GUI Data Model import passed on Excel 16.0 build 20430.0 with formula/model/row-count, backup and owned-process checks. Capability-probe cleanup now waits for and, if needed, terminates only its exact DispatchEx PID after Quit errors, failing closed on uncertain cleanup. W130 holds one kernel process handle through timeout termination and wait so later PID reuse cannot redirect cleanup after handle acquisition; the rebuilt GUI/CLI passed 365 embedded tests and packaged Data Model smoke on Excel 16.0 build 20430.0. W134 rechecks that the captured DispatchEx HWND still belongs to that PID after handle acquisition, failing closed and closing the handle on mismatch. The full build passed 368 embedded tests; packaged Data Model import passed on the same Excel build. W135 also passed packaged CLI background and synchronous static-M refresh with hash-matched backups, read-only query verification and owned-process cleanup on Excel 16.0 build 20430.0. W131 reran packaged GUI connection-only, new worksheet, existing worksheet and Data Model imports on that build; all formula/load-state, backup and owned-process checks passed. W132 added a widget-tree regression for visible enabled main-window controls. W133 replaces Qt’s table-corner affordance with an explicit Select all query rows button because the internal corner control was absent from Windows UI Automation; the packaged check now requires the Button role/name and passes, and a Qt test verifies it selects all rows. The rebuilt UI suite has 61 tests and passes at both 125% and 150% scale factors. W136 adds named interactive controls in the inspector, package update, profile mapping, history comparison and dependency graph dialogs, and accessible select-all mapping actions. W140 fixes history compare/restore UI-thread dispatch; the full suite passes 372 tests, with 65 UI tests at both 125% and 150% Qt scale. Packaged Windows UIA now verifies the query inspector, dependency graph, profile mappings, package update preview and history comparison; combined disposable Excel definition-only import passed on Excel 16.0 build 20430.0. Follow-up W141 packaged UIA restored a disposable workbook to the selected snapshot SHA-256 and verified the exact pre-restore backup hash, followed by a passing GUI import smoke. W142 packaged keyboard traversal completed a 35-control cycle in the expected order without Excel; corrupt local history was preserved byte-for-byte, archived, and reported during isolated-profile recovery. W143 real Excel staged recovery passed with one-query .xlsx and .xlsm-container workbooks, preserving original and stage hashes and reopening the exact formula; the .xlsm had no VBA project. W144 marshals worker-finished UI updates through a typed main-thread slot; 373 tests pass, Qt scale suites pass, and rebuilt packaged GUI Excel import passed on build 20430.0. W145 makes the packaged import acceptance harness reject elevated tokens; non-elevated package import passed on the current user profile. W146 fixes COM shutdown ordering after a real packaged VBA-bearing GUI import exposed an apartment-lifetime issue; on Windows 10 build 26300 / Excel 16.0 build 20430.0 the rebuilt package preserved VBA source without running Workbook_Open, verified exact M, backup and read-only reopen, and returned the app-owned Excel process set to baseline. Packaged CLI background and synchronous refresh also passed with static inline M and verified backups. Full suite: 379 passed; release gates remain UNVERIFIED for other Office builds and remaining Windows/accessibility/security matrix cases. W147 passes packaged history comparison UI Automation on Excel 16.0 build 20430.0, exposing the named changes table, formula diff and close control after two local snapshots; packaged import and exact-PID cleanup also pass, while pre-existing Excel PID 36464 remains untouched. Screen-reader, physical display, alternate Office builds and the wider matrix remain unverified. W148 passes packaged CLI source-path and M-literal mapping with explicit typed approvals on Excel 16.0 build 20430.0; mapped/untouched formulas and original source hashes match, with no external M evaluation. W149 makes CLI end-of-input a safe decline with exit 3; a failing-first regression and 380-test Windows build pass, followed by packaged CLI help and mapping smoke on Excel 16.0 build 20430.0. W151 verifies timeout termination during a genuinely blocked synchronous Excel QueryTable COM refresh against a delayed loopback-only response on build 20430.0; exact PID exit and preservation of pre-existing PID 36464 are confirmed. This does not certify Power Query refresh timeout or other Office builds. W127 hardens the smoke harness against its app process exiting between completion-dialog dismissal and close-window lookup, and the rebuilt GUI replacement then passed again. W125 full build: 359 tests in 15.64s; the rebuilt packaged GUI replacement smoke passed on Excel 16.0 build 20430.0. The packaged CLI deployed/repeated the portable bundle against two disposable Excel workbooks with formula, backup and process-exit verification. W115 explicitly opens a disposable workbook through ProtectedViewWindows.Open and observes the Protected View UI on Excel 16.0 build 20430.0; automatic ZoneId=3 triggering remains unverified; read-only Office 16 registry state has DisableInternetFilesInPV=1, and no security setting was changed. Release gates remain UNVERIFIED: arbitrary blocked COM termination, other Office builds, automatic ZoneId=3 Protected View triggering, OneDrive conflict, broader workbook object variants, physically damaged-history salvage (standard SQLite SELECT cannot read the truncated-page fixtures), screen-reader user testing and physical OS scaling are still open.

W175 binds the compressed ZIP size check and parser to one opened file descriptor; portable focused tests and the 406-test suite pass. W176 reruns the current compact-package GUI on Windows 10 build 26300 / Excel 16.0 build 20430.0 under a non-elevated token: all-skip preserves exact workbook hash/timestamp with no backup or stage, replacement verifies exact M and the original-hash backup while preserving an unrelated marker, and mixed numbered .m/.pq import verifies source/formula/order readback. The smoke-owned Excel processes exited and pre-existing PID 36464 remained. These are one-host results; release gates remain UNVERIFIED for other Office builds and the remaining accessibility, security, and workbook matrix. W177 re-runs packaged connection-only, new-worksheet, existing-worksheet at Output!C4, and Data Model loads on the same Excel build. It verifies connection/sheet/model distinctions, exact formulas, worksheet values 17/29, model membership/row count 2, backups and process cleanup. External providers and other Office builds remain unverified. W185 fails closed when the worksheet collection count, indexed item, or name cannot be read, preventing an uncertain name lookup from being treated as permission to create a sheet. The fake-COM regression passes. The 410-test Windows build and refreshed package passed verification (229 files / 135,790,075 bytes), contracts, 118 traceability references and CLI help. Rebuilt packaged new-worksheet and existing-worksheet GUI imports pass on Excel 16.0 build 20430.0 with values, exact formulas, backups, read-only verification and owned-process cleanup. Real collection-error injection and other Office builds remain unverified. W186 completes current-package keyboard traversal through 35 focusable controls without Excel; screen-reader and physical display checks remain open. W178 re-runs current packaged CLI static-M refresh in background and synchronous modes on Excel 16.0 build 20430.0. Both reports say refreshed, backups match, no external source is contacted, and owned Excel processes exit. W180 re-runs the real workbook preservation smoke on the same host, covering ordinary table formatting/data, names/formulas, chart, shape, PivotTable values, query-backed report references, load-state distinctions and privacy metadata. Other object types and Office builds remain unverified. W181 adds a failing-first fake-COM regression for a failed read-only verification-workbook close: the importer aborts before commit, preserves the original file, discards the stage and reports rollback. Full build passed 407 tests; the 229-file compact package, contracts, 115 traceability references, packaged CLI help and rebuilt GUI definition-only import passed on Excel 16.0 build 20430.0. Real Excel close-failure injection remains unverified. W182 re-runs the rebuilt packaged existing-query replacement on Excel 16.0 build 20430.0: exact new M, original-hash backup, unrelated marker, manager exit and owned-process cleanup all pass. W183 makes optional operation-log UIA extraction resilient to elements without control types; two regressions pass. The full Windows build passed 409 tests; refreshed package, contracts, 117 traceability references and CLI help pass. W184 verifies the refreshed packaged VBA-bearing `.xlsm` import preserved its harmless VBA project without running Workbook_Open, alongside exact M, backup, read-only reopen, and process cleanup on Excel 16.0 build 20430.0.
""",
    encoding="utf-8",
)

# Keep the compact package docs aligned with the latest acceptance corrections.
with (DIST / "README.txt").open("a", encoding="utf-8") as manual:
    manual.write(
        "\n\nLatest acceptance notes (2026-10-23): W153 packaged GUI definition-only import passed on Excel 16.0 "
        "build 20430.0 with exact M, backup, read-only reopen and owned-process checks. W152 found a packaged "
        "Power Query false-success defect. W154/W155 rebuilt and retested the CLI: successful connection-only "
        "refresh contacted a loopback CSV source in background and synchronous modes, while a deliberately delayed "
        "source timed out with exit 5, unchanged original bytes, verified backup and removed stage. The tested "
        "GUI import also passed again as W156. W157 passed packaged connection-only and existing worksheet-loaded "
        "refresh in background and synchronous modes; the worksheet table at D4 updated to 42 from a loopback source. "
        "W158 passed single-table Data Model refresh through Excel's synchronous ModelTable.Refresh API. On a "
        "multi-table model, partial selections were refused before source access; all-selected refresh used the "
        "synchronous Workbook.Model.Refresh API only after exact mapping checks. Both selected loopback sources, "
        "model row-count readback, and exact backups were verified. This was one tested host/build; "
        "other Office builds, external providers and broader matrix cases remain unverified. See source handoff. "
        "W188 also confirmed on the same Excel build that import, refresh, and overwrite-restore each refuse a "
        "read-only open workbook detected through Excel ROT without an owner-lock file; no stage or backup was "
        "created and the workbook hash remained unchanged.\n"
        "W189 reran packaged GUI UI Automation and definition-only import on the current Windows host/build with "
        "exact M, backup, read-only reopen, exit 0 and owned Excel cleanup.\n"
        "W190 completed three sequential packaged GUI import runs; each verified formulas/backups and cleaned its "
        "owned Excel processes. UI Automation did not expose the optional operation-log text in two runs, which "
        "was addressed in W191 by adding stable automation-ID metadata and retry/TextPattern lookup. Three "
        "sequential UIA runs and a strict-marker import then exposed the final operation-log completion record. "
        "W192 reran the 35-control keyboard cycle; critical import controls stayed in order and focus remained "
        "inside the manager. Screen-reader and physical display review remain open. W193 rebuilt and published "
        "the current 413-test source tree and passed the strict operation-log/import smoke on that package. W194 "
        "fixes stage-cleanup false success when file absence cannot be confirmed; the 414-test rebuild and current "
        "packaged GUI import smoke pass. W195 directly exercises Windows delete-sharing denial on a disposable "
        "stage: cleanup reports and preserves the file until the owned blocking handle closes. W196 updates M, "
        "diff, gutter and current-line colors for dark/light themes; automated contrast and scale tests pass.\n"
    )
with (DIST / "RELEASE_SUMMARY.txt").open("a", encoding="utf-8") as summary:
    summary.write(
        "\n\nW152 — FAIL: packaged Power Query refresh reported success without contacting its loopback source. "
        "W154 — PASS: rebuilt packaged CLI connection-only query refresh contacted a loopback CSV source in "
        "background and synchronous modes; exact formulas, backups, read-only reopen, original connection/load "
        "shape and owned Excel cleanup passed on Excel 16.0 build 20430.0. W155 — PASS: the delayed loopback "
        "Power Query refresh returned exit 5 at the 3-second timeout, preserved original bytes and backup, "
        "removed its stage, and terminated only its captured Excel PID. W157 — PASS: packaged worksheet-loaded "
        "refresh updated the existing D4 table to 42 in both modes, retaining query identity, formula and backup. "
        "W158 — PASS: single-table Data Model refresh ran through ModelTable.Refresh; multi-table full selection ran "
        "through Workbook.Model.Refresh only after all model sources mapped to selected queries. A partial selection "
        "was refused before source access or workbook change. Both loopback sources, two-row model readback, and exact "
        "backups were verified. "
        "The Data Model API has no background mode. W153/W156 — PASS: packaged GUI controls and definition-only import "
        "passed UI Automation and disposable-workbook readback. These are one host/build; other Office builds, external "
        "providers and the remaining Excel matrix stay UNVERIFIED, so release gates remain UNVERIFIED. W188 — PASS: "
        "with a disposable workbook held read-only open in Excel and no owner-lock file, import, refresh, and "
        "overwrite-restore refused before stage/backup creation; hash preserved and owned Excel closed. W189 — PASS: "
        "recovered GUI package import and UI Automation on the available Windows/Excel build, exact M, backup, "
        "read-only reopen, exit 0 and owned-process cleanup. W190 — PASS: three sequential GUI import runs each "
        "verified exact M, backup, reopen and exit 0 with owned-process cleanup; operation-log text was unavailable "
        "to UIA in two runs. W191 — PASS: stable operation-log automation ID, retry and TextPattern lookup exposed "
        "the complete final log in three sequential runs and a strict-marker smoke. Full release gates remain "
        "UNVERIFIED. W192 — PASS: packaged 35-control keyboard traversal reached import controls in order and kept "
        "focus inside the manager. Screen-reader and physical display review remain open. W193 — PASS: rebuilt "
        "package passed strict final-log UIA, exact formula, backup, reopen, app exit and owned-process cleanup. "
        "Current package has 229 files; complete release gates remain UNVERIFIED. W194 — PASS: stage cleanup now "
        "fails closed if absence cannot be verified; the 414-test rebuilt package and strict packaged import smoke pass. "
        "W195 — PASS: real Windows delete-sharing denial preserved and reported the disposable stage, then cleanup "
        "succeeded after the test-owned handle closed. This is not an Excel-originated lock test. W196 — PASS: "
        "light/dark editor token, diff and gutter contrast checks and nested inspector theme update pass. The current "
        "UI suite passes at 125% and 150% Qt scale. W197 — PASS: Excel 16.0 build 20430.0 preserved internal CRLF "
        "in Query.Formula but removed terminal CR/LF; rebuilt GUI import accepts only that observed normalization, "
        "preserves the BOM/CRLF source file SHA-256, and passed exact readback under the narrow rule, backup, "
        "read-only reopen, UIA completion log, and owned-process cleanup for create and replace. Full suite: "
        "424 passed. W198 — PASS: final staged commit now uses native ReplaceFileW, refuses to create a "
        "disappeared workbook path, verifies both target/stage hashes, and attempts recovery after partial "
        "replacement errors; the Windows NTFS stream test confirms the original Zone.Identifier survives. "
        "Packaged GUI replacement passed on Excel 16.0 build 20430.0, including a BOM/CRLF source replacement "
        "with unchanged source hash; build suite: 424. W199's second clean onedir build matched the published "
        "W198 payload across all 227 runtime files. W200 tests injected partial API error 1177 recovery using real "
        "MoveFileW; OS-generated partial replacement errors remain unverified. W201 — PASS: the current "
        "packaged GUI ran from a disposable read/execute-only folder; its write probe was denied, and the "
        "non-elevated GUI import, exact formula, verified backup, read-only reopen, manager exit and owned "
        "Excel cleanup passed on Excel 16.0 build 20430.0. Existing Excel PID 36464 remained untouched. "
        "Other ACL/policy variants and Office builds remain unverified. W203 closes a history-restore safety gap: "
        "both snapshot publication and verification rollback now require an existing workbook path; two failing-first "
        "native-boundary regressions pass, and two clean builds each pass 426 tests with identical hashes across all 227 runtime files. Packaged history restore passed on Excel 16.0 build 20430.0 with exact snapshot and pre-restore backup hashes verified. W206 packaged GUI import and UI Automation also passed at Qt scale factors 1.25 and 1.5 with exact M, backup, reopen and process cleanup. W209 packaged load-mode reacceptance passed for new sheet, existing sheet at Output!C4, connection-only, and Data Model on Excel 16.0 build 20430.0 with exact formulas, appropriate object distinctions, worksheet/model values, backups and owned-process cleanup. Other Office builds and external providers remain unverified. W210 current-package acceptance also passed existing-query replace, preserve-existing connection-only replacement, all-skip no-op, and mixed numbered .m/.pq profile import with exact hashes/formulas/order and owned-process cleanup. W211 current-package refresh acceptance passed connection-only and worksheet-loaded Power Query refresh in background and synchronous modes against localhost, the 3-second timeout with exit 5 and exact original/backup preservation, and Data Model refresh with both selected/helper tables returning two rows. Other Office builds and external providers remain unverified. W213 current-package VBA-bearing .xlsm smoke preserved vbaProject.bin and the harmless Workbook_Open marker did not execute; formula, backup, read-only reopen, exit and owned-process cleanup passed on Excel 16.0 build 20430.0. XLM and other Office/VBA forms remain unverified. W214 current packaged CLI deployment also passed a two-workbook .pqmdeploy apply/repeat smoke: first statuses applied/applied, repeated no_changes/no_changes, exact formula/load readback and original-hash backups verified. The disabled missing source was ignored and three owned Excel PIDs exited. An earlier attempt detected PID 2200 disappearing; ownership/cause unknown, and only a stable-baseline retry is counted. W215 mixed-target deployment isolation passed: writable target applied with a verified backup; read-only target blocked before backup and remained byte-identical; owned Excel processes exited. Mid-commit batch failures, sync conflicts and other Office builds remain unverified. W216 packaged source-path/M-literal mappings and W217 packaged CLI query-package export/import passed on Excel 16.0 build 20430.0 with exact previews, typed approvals, unchanged source bytes, exact Unicode/formula readback, verified backup and owned-process cleanup; no external M source was evaluated. Screen-reader, physical display and broader Excel matrix remain unverified.\n"
    )

with (DIST / "README.txt").open("a", encoding="utf-8") as manual:
    manual.write(
        "\nW219 — PASS: packaged GUI declared-parameter editing saved the approved literal to a separate source copy, "
        "preserved the original source SHA-256, imported the edited definition into a disposable workbook, and "
        "verified Excel formula, timestamped backup, read-only reopen, manager exit and owned-process cleanup on "
        "Excel 16.0 build 20430.0. Other Office builds and screen-reader acceptance remain unverified.\n"
    )
with (DIST / "RELEASE_SUMMARY.txt").open("a", encoding="utf-8") as summary:
    summary.write(
        "\nW219 — PASS: packaged GUI declared-parameter editing saved the approved literal to a separate source copy, "
        "preserved the original source SHA-256, imported the edited definition into a disposable workbook, and "
        "verified Excel formula, timestamped backup, read-only reopen, manager exit and owned-process cleanup on "
        "Excel 16.0 build 20430.0. Other Office builds and screen-reader acceptance remain unverified; release gates "
        "remain UNVERIFIED.\n"
    )
with (DIST / "README.txt").open("a", encoding="utf-8") as manual:
    manual.write(
        "W221 — PASS: multi-workbook deployment failure reports now include the target's observed post-failure "
        "SHA-256 and whether it still matches the pre-deployment bytes, plus whether a verified backup is available. "
        "The 427-test Windows build passed; packaged GUI import and two-workbook CLI apply/repeat smokes passed on "
        "Excel 16.0 build 20430.0. These do not certify cross-workbook atomicity or simulate native partial Excel commit failures.\n"
    )
with (DIST / "RELEASE_SUMMARY.txt").open("a", encoding="utf-8") as summary:
    summary.write(
        "\nW221 — PASS: multi-workbook deployment failure reports now include the target's observed post-failure "
        "SHA-256 and whether it still matches the pre-deployment bytes, plus whether a verified backup is available. "
        "The 427-test Windows build passed; packaged GUI import and two-workbook CLI apply/repeat smokes passed on "
        "Excel 16.0 build 20430.0. This does not certify cross-workbook atomicity or simulate native partial Excel "
        "commit failures; release gates remain UNVERIFIED.\n"
    )
with (DIST / "README.txt").open("a", encoding="utf-8") as manual:
    manual.write(
        "W225 — PASS: dependency scanning now ends M line comments at CR, CRLF, LF, NEL, line separator and "
        "paragraph separator boundaries. Unit and dependency graph UI regressions pass; the tokenizer remains "
        "advisory and is not a full M parser.\n"
    )
with (DIST / "RELEASE_SUMMARY.txt").open("a", encoding="utf-8") as summary:
    summary.write(
        "\nW225 — PASS: dependency scanning now ends M line comments at CR, CRLF, LF, NEL, line separator and "
        "paragraph separator boundaries. Unit and dependency graph UI regressions pass; the tokenizer remains "
        "advisory and is not a full M parser. Other Office and remaining Windows matrix gates remain UNVERIFIED.\n"
    )
with (DIST / "README.txt").open("a", encoding="utf-8") as manual:
    manual.write(
        "W227 — PASS: regular M identifiers are now tokenized using the documented Unicode start and part "
        "character classes, including combining, connector and formatting characters. This avoids false-negative "
        "query references; dependency scanning remains advisory and is not a complete parser.\n"
    )
with (DIST / "RELEASE_SUMMARY.txt").open("a", encoding="utf-8") as summary:
    summary.write(
        "\nW227 — PASS: regular M identifiers are now tokenized using the documented Unicode start and part "
        "character classes, including combining, connector and formatting characters. This avoids false-negative "
        "query references; dependency scanning remains advisory and is not a complete parser. Remaining Excel "
        "matrix gates are UNVERIFIED.\n"
    )
with (DIST / "README.txt").open("a", encoding="utf-8") as manual:
    manual.write(
        "W229 — PASS: the advisory dependency scanner decodes M character escapes and doubled quotes in quoted "
        "identifiers, and marks malformed escapes ambiguous. Failing-first scanner and graph UI regressions pass. "
        "The rebuilt 445-test package passed definition-only GUI import/readback on Excel 16.0 build 20430.0 with "
        "verified backup, read-only reopen, manager exit and owned-process cleanup. Broader release gates remain unverified.\n"
    )
with (DIST / "RELEASE_SUMMARY.txt").open("a", encoding="utf-8") as summary:
    summary.write(
        "\nW229 — PASS: quoted-identifier character escapes are decoded for dependency matching and malformed "
        "escapes are marked ambiguous. The 445-test Windows build, package verification and packaged GUI "
        "definition-only import/readback passed on Excel 16.0 build 20430.0; verified backup, read-only reopen, "
        "manager exit and owned-process cleanup were confirmed. Other Office and accessibility/security matrix "
        "gates remain UNVERIFIED.\n"
    )
with (DIST / "README.txt").open("a", encoding="utf-8") as manual:
    manual.write(
        "W230 — PASS: dotted regular identifiers are kept whole during dependency scanning, so `Base.Field` "
        "does not create a false edge to query `Base`; exact dotted names still match, and dynamic `Record.Field` "
        "constructs remain ambiguous. The 449-test Windows build passed; focused failing-first regressions pass. "
        "This bounded scanner is not a full M parser.\n"
    )
with (DIST / "RELEASE_SUMMARY.txt").open("a", encoding="utf-8") as summary:
    summary.write(
        "\nW230 — PASS: the dependency scanner compares complete dotted M regular identifiers and keeps dynamic "
        "Record.Field constructs ambiguous. The 449-test Windows build and packaged GUI definition-only import "
        "passed on Excel 16.0 build 20430.0 with exact formula readback, verified backup, read-only reopen, "
        "manager exit 0 and owned-process cleanup. Tests prevent partial-name and record-field-label edges; "
        "the scanner remains advisory rather than a complete M parser. Other Office and broader accessibility, "
        "security, and workbook gates remain UNVERIFIED.\n"
    )
with (DIST / "README.txt").open("a", encoding="utf-8") as manual:
    manual.write(
        "W231 — PASS: hash-prefixed M keywords are tokenized as reserved constructs instead of query references. "
        "The scanner treats `#shared` and `#sections` as ambiguous/dynamic and flags unknown `#identifier` forms. "
        "The 461-test Windows build passed; regression coverage passes. This scanner remains advisory.\n"
    )
with (DIST / "RELEASE_SUMMARY.txt").open("a", encoding="utf-8") as summary:
    summary.write(
        "\nW231 — PASS: the dependency scanner recognizes M hash-prefixed literal keywords, prevents false edges "
        "to same-named queries, and flags `#shared`, `#sections`, and unknown hash-prefixed identifiers as "
        "ambiguous where reference resolution cannot be established. The 461-test Windows build and packaged GUI "
        "definition-only import passed on Excel 16.0 build 20430.0 with exact formula, verified backup, read-only "
        "reopen, manager exit 0, and owned-process cleanup. Broader release gates remain UNVERIFIED.\n"
    )
with (DIST / "README.txt").open("a", encoding="utf-8") as manual:
    manual.write(
        "W232 — PASS: the disposable Excel preservation smoke now checks data validation, conditional formatting, "
        "and a hyperlink across staged query replacement, in addition to the existing table, chart, shape, PivotTable, "
        "report-link and privacy checks. Observed on Excel 16.0 build 20430.0; broader object/build coverage remains open.\n"
    )
with (DIST / "RELEASE_SUMMARY.txt").open("a", encoding="utf-8") as summary:
    summary.write(
        "\nW232 — PASS: the source-based disposable Excel 16.0 build 20430.0 staged replacement smoke preserved "
        "data validation, conditional formatting, hyperlink, ordinary table/filter/style, defined names, formulas, "
        "chart, shape, PivotTable values, query-backed report links and default workbook privacy metadata. The three "
        "added objects were read back after replacement. Other Excel builds and workbook-object variants remain "
        "UNVERIFIED.\n"
    )

with (DIST / "README.txt").open("a", encoding="utf-8") as manual:
    manual.write(
        "W233 — PASS: current packaged GUI keyboard traversal completed a 35-control Tab cycle, with core import "
        "controls reachable in order and focus staying inside the manager. No workbook or Excel instance was opened. "
        "Modal dialogs, screen-reader use, physical display scaling and other Office builds remain unverified.\n"
    )
with (DIST / "RELEASE_SUMMARY.txt").open("a", encoding="utf-8") as summary:
    summary.write(
        "\nW233 — PASS: packaged GUI keyboard traversal visited 35 focusable controls and completed a full Tab cycle "
        "within the manager. Workbook path, Add files, query table, existing-query policy, backup and import controls "
        "were reachable in order. No workbook or Excel instance was opened; pre-existing Excel PID 36464 remained. "
        "Modal keyboard paths, screen-reader use, physical display scaling and other Office builds remain UNVERIFIED.\n"
    )

with (DIST / "README.txt").open("a", encoding="utf-8") as manual:
    manual.write(
        "W234 — PASS: an M quoted identifier not present in the available source list is shown as ambiguous in "
        "dependency status; the graph adds no guessed edge. This warns about unresolved context without claiming "
        "the query is absent from the workbook. Windows package smoke used a disposable workbook; broader M and Office coverage remains open.\n"
    )
with (DIST / "RELEASE_SUMMARY.txt").open("a", encoding="utf-8") as summary:
    summary.write(
        "\nW234 — PASS: a quoted M identifier outside the available source list appears as an ambiguous dependency "
        "warning and creates no guessed dependency edge. The scanner does not claim that the query is missing from "
        "the workbook because existing workbook queries are not part of this source-list graph. Rebuilt-package "
        "replacement smoke passed on Excel 16.0 build 20430.0 with exact formula, verified backup, read-only reopen "
        "and owned-process cleanup. Other Office builds and broader reference semantics remain UNVERIFIED.\n"
    )

with (DIST / "README.txt").open("a", encoding="utf-8") as manual:
    manual.write(
        "W235 — PASS: the current rebuilt packaged CLI refreshed disposable connection-only and worksheet-loaded "
        "queries in background and synchronous modes against a loopback-only source. A delayed loopback refresh "
        "timed out with exit 5 and preserved the workbook/backup while removing its stage. Other Office builds and external providers remain unverified.\n"
    )
with (DIST / "RELEASE_SUMMARY.txt").open("a", encoding="utf-8") as summary:
    summary.write(
        "\nW235 — PASS: current rebuilt package acceptance on Excel 16.0 build 20430.0 verified all four GUI load modes, "
        "packaged CLI background/synchronous connection-only refresh and worksheet-loaded refresh, and the 3-second "
        "delayed-loopback Power Query timeout. Worksheet refresh read back value 42. Exact formulas, load states, "
        "pre-operation backups, stage cleanup and owned Excel process cleanup passed; no external network was used. "
        "Other Office builds, providers and workbook matrix cases remain UNVERIFIED.\n"
    )

with (DIST / "README.txt").open("a", encoding="utf-8") as manual:
    manual.write(
        "W236 — PASS: current package preserved legacy mixed numbered .m/.pq import order, Skip no-op behavior, and "
        "connection-only load during query replacement on Excel 16.0 build 20430.0. Broader Office and workbook coverage remains open.\n"
    )
with (DIST / "RELEASE_SUMMARY.txt").open("a", encoding="utf-8") as summary:
    summary.write(
        "\nW236 — PASS: rebuilt package mixed `.m`/`.pq` import matched numeric review/Excel order and exact formulas; "
        "all-skip preserved workbook hash/timestamp with no backup or stage; replacing an existing query preserved "
        "its connection-only load. Verified backups, read-only reopen, manager exit and owned-process cleanup passed "
        "on Excel 16.0 build 20430.0; pre-existing PID 36464 remained. Other Office builds remain UNVERIFIED.\n"
    )

with (DIST / "README.txt").open("a", encoding="utf-8") as manual:
    manual.write(
        "W237 — PASS: current package deployment bundle apply/repeat, read-only target isolation, Unicode query-package "
        "roundtrip and packaged GUI comparison dialogs passed on Excel 16.0 build 20430.0. Other Office and policy variants remain unverified.\n"
    )
with (DIST / "RELEASE_SUMMARY.txt").open("a", encoding="utf-8") as summary:
    summary.write(
        "\nW237 — PASS: current packaged CLI applied a portable bundle to two disposable workbooks and repeated as "
        "no_changes/no_changes with exact M, two verified backups and disabled missing source ignored. Current-source "
        "deployment blocked a read-only target before backup, preserving its hash while the writable target applied. "
        "Packaged query-package export/import roundtripped three formulas, including Unicode/punctuation, without "
        "changing the source workbook or evaluating M. Packaged GUI inspector, dependency graph, profile mapping and "
        "package comparison dialogs exposed named controls; mapping edits were cancelled; disposable definition-only "
        "import passed. Excel 16.0 build 20430.0 only; other Office, policy and workbook variants remain UNVERIFIED.\n"
    )

with (DIST / "README.txt").open("a", encoding="utf-8") as manual:
    manual.write(
        "W238 — PASS: current packaged GUI launched and imported from a read/execute-only install directory under a "
        "non-elevated token. Restricted corporate policy variants remain unverified.\n"
    )
with (DIST / "RELEASE_SUMMARY.txt").open("a", encoding="utf-8") as summary:
    summary.write(
        "\nW238 — PASS: a disposable copy of the current one-folder package had only read/execute permissions; a write "
        "probe was denied, then the packaged GUI launched under a non-elevated token and completed definition-only "
        "import to a separate disposable workbook with exact formula, verified backup, read-only reopen and owned "
        "Excel cleanup. The install directory stayed read-only. Restricted corporate policy variants remain UNVERIFIED.\n"
    )

with (DIST / "README.txt").open("a", encoding="utf-8") as manual:
    manual.write(
        "W239 — PASS: fresh portable suite and release verification passed on Windows 10 build 26300; "
        "remaining Excel/accessibility gates stay unverified.\n"
    )
with (DIST / "RELEASE_SUMMARY.txt").open("a", encoding="utf-8") as summary:
    summary.write(
        "\nW239 — PASS: fresh current-source verification on Windows 10 build 26300 / Python 3.11.9 passed "
        "463 tests, schema/contract validation, 134 feature-traceability references, and the 229-file compact "
        "release manifest. All 227 packaged runtime files match the current PyInstaller payload byte-for-byte; "
        "packaged CLI --help exits 0. No new Excel behavior was exercised in this verification pass; the latest "
        "disposable-Excel acceptance remains W238 and earlier W235–W237. Alternate Office builds, external providers, "
        "screen-reader, physical display, and remaining security/workbook matrix gates remain UNVERIFIED.\n"
    )

with (DIST / "README.txt").open("a", encoding="utf-8") as manual:
    manual.write(
        "W240 — PASS: a clean deterministic PyInstaller rebuild matched all 227 current runtime files byte-for-byte "
        "on this Windows/Python/toolchain. Cross-host reproducibility remains unverified.\n"
    )
with (DIST / "RELEASE_SUMMARY.txt").open("a", encoding="utf-8") as summary:
    summary.write(
        "\nW240 — PASS: a clean PyInstaller onedir rebuild into an isolated temporary directory reproduced the current "
        "Windows package exactly. With PYTHONHASHSEED=0 and SOURCE_DATE_EPOCH=1704067200, all 227 runtime files "
        "matched byte-for-byte; there were zero missing, extra or changed files. Host: Windows 10 build 26300, "
        "Python 3.11.9, PyInstaller 6.13.0. This establishes same-host/toolchain repeatability only; cross-host "
        "reproducibility and all remaining Office/security/accessibility gates remain UNVERIFIED.\n"
    )

with (DIST / "README.txt").open("a", encoding="utf-8") as manual:
    manual.write(
        "W241 — PASS: packaged keyboard smoke covered the main window and source-mapping modal; Tab cycles, "
        "Right Arrow tab switching, and Escape dismissal passed without opening Excel.\n"
    )
with (DIST / "RELEASE_SUMMARY.txt").open("a", encoding="utf-8") as summary:
    summary.write(
        "\nW241 — PASS: the current packaged GUI passed UI Automation keyboard traversal through the 35-control "
        "main-window cycle and source-mapping modal. Tab reached both mapping tables and their actions; Right Arrow "
        "switched from Source paths to M source locations; Escape dismissed the modal. No workbook or Excel instance "
        "was opened and the only Excel process remained pre-existing PID 36464. The result covers keyboard navigation "
        "for this dialog, not screen-reader user testing, physical display scaling, or other Office builds.\n"
    )

with (DIST / "README.txt").open("a", encoding="utf-8") as manual:
    manual.write(
        "W242 — PASS: packaged keyboard smoke also traversed operation history controls using a temporary user profile; "
        "no workbook or Excel was opened.\n"
    )
with (DIST / "RELEASE_SUMMARY.txt").open("a", encoding="utf-8") as summary:
    summary.write(
        "\nW242 — PASS: the current packaged GUI operation-history dialog completed a UI Automation Tab cycle through "
        "the history table, retention, snapshot setting, refresh, delete and close controls. The dialog was dismissed "
        "with Escape. `LOCALAPPDATA` pointed to a temporary profile that was removed after manager exit; saved user "
        "history was not read or changed. No workbook or Excel process was opened; pre-existing PID 36464 remained. "
        "This verifies keyboard navigation for the operation-history dialog only, not screen-reader or physical-display use.\n"
    )

with (DIST / "README.txt").open("a", encoding="utf-8") as manual:
    manual.write(
        "W243 — PASS: the dependency-graph modal's text summary and Close action are keyboard reachable; Escape dismisses it.\n"
    )
with (DIST / "RELEASE_SUMMARY.txt").open("a", encoding="utf-8") as summary:
    summary.write(
        "\nW243 — PASS: the current packaged GUI dependency-graph dialog completed a UI Automation Tab cycle through "
        "its text summary and Close action, then closed on Escape. The graph was empty and no workbook or Excel instance "
        "was opened. This validates keyboard access for the empty-graph dialog only; populated graph interaction, "
        "screen-reader use, physical display scaling and other Office builds remain UNVERIFIED.\n"
    )

with (DIST / "README.txt").open("a", encoding="utf-8") as manual:
    manual.write(
        "W244 — PASS: Tab moves focus through the M editors in the query inspector and reaches save-copy/close actions.\n"
    )
with (DIST / "RELEASE_SUMMARY.txt").open("a", encoding="utf-8") as summary:
    summary.write(
        "\nW244 — PASS: the query inspector's source, workbook and diff editors now let Tab move focus to the next "
        "control instead of inserting indentation, making the full dialog keyboard reachable. A failing-first Qt "
        "regression checks all three editor properties; the Windows build passed 467 tests. The rebuilt compact "
        "package passed query-inspector keyboard traversal through search, Find actions, all three editors, save-copy and "
        "Close, followed by the operation-history, empty dependency-graph, and both source-mapping tab paths. The "
        "history check ran first in a temporary LOCALAPPDATA profile because source inspection creates local history. "
        "No workbook or Excel instance was opened; pre-existing Excel PID 36464 remained. Screen-reader, physical "
        "display and alternate Office checks remain UNVERIFIED.\n"
    )

with (DIST / "README.txt").open("a", encoding="utf-8") as manual:
    manual.write(
        "W245 — PASS: the rebuilt packaged GUI definition-only import passed on Excel 16.0 build 20430.0 with exact M readback and a verified backup.\n"
    )
with (DIST / "RELEASE_SUMMARY.txt").open("a", encoding="utf-8") as summary:
    summary.write(
        "\nW245 — PASS: the rebuilt one-folder packaged GUI completed a definition-only import on a disposable `.xlsx` "
        "with Excel 16.0 build 20430.0. UI Automation found the named main window, inspector, dependency graph, "
        "mapping controls and completion log. Read-only reopen verified the exact M formula; the timestamped backup "
        "matched the original workbook SHA-256; the manager exited 0 and the smoke-owned Excel instances exited. "
        "Pre-existing Excel PID 36464 remained untouched. The harness ran non-elevated and contacted no external "
        "source. This does not certify other Office builds, external providers, or remaining workbook/security and "
        "accessibility matrix items, which remain UNVERIFIED.\n"
    )

with (DIST / "README.txt").open("a", encoding="utf-8") as manual:
    manual.write(
        "W246 — PASS: existing-query replacement passed on Excel 16.0 build 20430.0 with exact M and backup verification.\n"
    )
with (DIST / "RELEASE_SUMMARY.txt").open("a", encoding="utf-8") as summary:
    summary.write(
        "\nW246 — PASS: the current rebuilt GUI replaced an existing definition-only query in a disposable `.xlsx` "
        "on Excel 16.0 build 20430.0. Read-only reopen matched the new exact M formula and the timestamped backup "
        "matched the original SHA-256; an unrelated worksheet marker was preserved. Manager exit was 0. Seed PID "
        "26924 and app-owned PID 43216 exited; verification PID 38424 was terminated during cleanup; the process "
        "set returned to `[36464]`, leaving pre-existing Excel PID 36464 untouched. No refresh or external source "
        "was requested. Other Office builds, loaded-query association variants, and broader workbook/security/"
        "accessibility gates remain UNVERIFIED.\n"
    )

with (DIST / "README.txt").open("a", encoding="utf-8") as manual:
    manual.write(
        "W247 — PASS: the current deterministic clean build reproduced all 227 runtime files byte-for-byte.\n"
    )
with (DIST / "RELEASE_SUMMARY.txt").open("a", encoding="utf-8") as summary:
    summary.write(
        "\nW247 — PASS: after the query-inspector keyboard change, a clean PyInstaller onedir rebuild in an isolated "
        "temporary directory reproduced the current runtime payload exactly. With PYTHONHASHSEED=0 and "
        "SOURCE_DATE_EPOCH=1704067200, current and rebuilt payloads each contained 227 files, with zero missing, "
        "extra, or changed files. Host: Windows 10 build 26300 / Python 3.11.9 / PyInstaller 6.13.0. This is "
        "same-host/toolchain repeatability only; cross-host reproducibility and remaining Office, security, and "
        "accessibility gates remain UNVERIFIED.\n"
    )

with (DIST / "README.txt").open("a", encoding="utf-8") as manual:
    manual.write(
        "W248 — PASS: all four packaged GUI load modes passed on Excel 16.0 build 20430.0 using disposable workbooks.\n"
    )
with (DIST / "RELEASE_SUMMARY.txt").open("a", encoding="utf-8") as summary:
    summary.write(
        "\nW248 — PASS: the rebuilt packaged GUI reaccepted new-worksheet, existing-worksheet at Output!C4, explicit "
        "connection-only, and Data Model loads on Excel 16.0 build 20430.0, Windows 10 build 26300, non-elevated. "
        "Each disposable workbook reopened read-only with the exact formula and a matching original-hash backup. "
        "New/existing worksheet tables read back values 17/29; the existing sheet preserved its unrelated marker; "
        "connection-only produced a workbook connection without worksheet/model output; Data Model membership and "
        "row count 2 were verified. Each manager exited 0 and the process set returned to pre-existing Excel PID "
        "36464. Static inline M was used; no external source was contacted. This is one Office build and does not "
        "certify external providers, other builds, or the remaining workbook/security/accessibility matrix, which "
        "remain UNVERIFIED.\n"
    )

with (DIST / "README.txt").open("a", encoding="utf-8") as manual:
    manual.write(
        "W249 — PASS: background and synchronous refresh passed on Excel 16.0 build 20430.0 through the packaged CLI.\n"
    )
with (DIST / "RELEASE_SUMMARY.txt").open("a", encoding="utf-8") as summary:
    summary.write(
        "\nW249 — PASS: the rebuilt packaged CLI refreshed a disposable connection-only static-M query in both "
        "background and synchronous modes on Excel 16.0 build 20430.0. Both JSON reports returned `refreshed`; "
        "each timestamped backup hash matched the exact pre-refresh workbook bytes, and read-only post-refresh "
        "inspection confirmed the exact formula and connection-only load state. `no_external_source` and "
        "`owned_excel_processes_exited` were true. No network source was contacted. This is one Office build and "
        "does not certify external providers, worksheet/Data Model refresh, or other Office versions; remaining "
        "refresh/security/accessibility gates stay UNVERIFIED.\n"
    )

with (DIST / "README.txt").open("a", encoding="utf-8") as manual:
    manual.write(
        "W250 — PASS: packaged CLI background/synchronous refresh passed for connection-only and worksheet-loaded Power Query.\n"
    )
with (DIST / "RELEASE_SUMMARY.txt").open("a", encoding="utf-8") as summary:
    summary.write(
        "\nW250 — PASS: the rebuilt packaged CLI refreshed connection-only and worksheet-loaded Power Query queries "
        "against a temporary 127.0.0.1 CSV endpoint in background and synchronous modes on Excel 16.0 build 20430.0. "
        "All four reports returned `refreshed`; seven loopback HTTP requests were observed. Exact formulas and "
        "connection shape were preserved, four pre-refresh backups matched their exact workbook hashes, the worksheet "
        "output reopened with value 42, and temporary refresh connections/sheets were removed. The Excel process set "
        "returned to `[36464]`; no external network was used. Data Model refresh, other providers/builds and remaining "
        "security/accessibility gates remain UNVERIFIED.\n"
    )

with (DIST / "README.txt").open("a", encoding="utf-8") as manual:
    manual.write(
        "W251 — PASS: 3-second refresh timeout returned exit 5; original workbook and verified backup were preserved.\n"
    )
with (DIST / "RELEASE_SUMMARY.txt").open("a", encoding="utf-8") as summary:
    summary.write(
        "\nW251 — PASS: the current packaged CLI's synchronous Power Query refresh timed out after 3 seconds while "
        "waiting on a delayed loopback `Web.Contents` response, on Excel 16.0 build 20430.0. The request was received; "
        "CLI exit code was 5 and the report identified timeout. Original workbook hash was preserved, the timestamped "
        "backup matched, and the abandoned stage was removed. Excel process set returned to `[36464]`; no external "
        "network was used. Other Office versions, providers and remaining workbook/security/accessibility gates stay "
        "UNVERIFIED.\n"
    )

with (DIST / "README.txt").open("a", encoding="utf-8") as manual:
    manual.write(
        "W252 — CLI path mapping in noninteractive runs requires --confirm \"APPLY PATH MAPPINGS\" after reviewing the preview.\n"
    )
with (DIST / "RELEASE_SUMMARY.txt").open("a", encoding="utf-8") as summary:
    summary.write(
        "\nW252 — CLI bundle, plan, and deploy now report missing noninteractive source-mapping approval as an approval requirement and provide the exact confirmation phrase. The operation remains fail closed.\n"
    )

with (DIST / "README.txt").open("a", encoding="utf-8") as manual:
    manual.write(
        "W253 — Noninteractive deploy/import/refresh use --confirm DEPLOY, IMPORT, or REFRESH. Source and M-code mapping reviews use --confirm \"APPLY PATH MAPPINGS\" and --confirm \"APPLY M-CODE MAPPINGS\". Missing approval is reported separately; a wrong phrase is treated as a decline. Path-only plans also require mapping approval.\n"
    )
with (DIST / "RELEASE_SUMMARY.txt").open("a", encoding="utf-8") as summary:
    summary.write(
        "\nW253 — CLI consequential actions now distinguish missing noninteractive approval from an explicit decline and print the exact --confirm phrase required to continue. All actions remain fail closed.\n"
    )

with (DIST / "README.txt").open("a", encoding="utf-8") as manual:
    manual.write(
        "W254 — Workbook query inventory, export/package/refresh actions, and refresh options now have explicit accessible names for assistive technology.\n"
    )
with (DIST / "RELEASE_SUMMARY.txt").open("a", encoding="utf-8") as summary:
    summary.write(
        "\nW254 — Added explicit accessible names to workbook inventory and refresh option controls, with focused UI coverage.\n"
    )

with (DIST / "README.txt").open("a", encoding="utf-8") as manual:
    manual.write(
        "W256 — Missing required confirmation uses the dedicated approval_required diagnostic and keeps the operation fail closed.\n"
    )
with (DIST / "RELEASE_SUMMARY.txt").open("a", encoding="utf-8") as summary:
    summary.write(
        "\nW256 — Missing noninteractive approval now has a dedicated actionable error category with the exact confirmation phrase; wrong phrases remain explicit declines.\n"
    )

with (DIST / "README.txt").open("a", encoding="utf-8") as manual:
    manual.write(
        "W258/W259 — The Windows interface is English-only. The package omits unused Qt translation catalogs and unused QML/Quick, PDF, virtual-keyboard, and OpenGL components; QtCore, QtGui, QtWidgets, and the required Windows platform plugin remain included.\n"
    )
with (DIST / "RELEASE_SUMMARY.txt").open("a", encoding="utf-8") as summary:
    summary.write(
        "\nW257 — The current packaged GUI imported a definition-only query into a disposable Excel workbook; exact formula, verified backup and read-only reopen passed, and the existing Excel PID 36464 remained untouched.\n"
        "W258/W259 — Package trimming removed 94 unused Qt translation catalogs and unused QML/Quick, PDF, virtual-keyboard and OpenGL runtime components. Package verification rejects their return. The final 124-file package is about 88.5 MB and passed packaged CLI help plus GUI UI/import acceptance on Excel 16.0 build 20430.0; other Qt surfaces and Office builds remain unverified.\n"
    )

with (DIST / "README.txt").open("a", encoding="utf-8") as manual:
    manual.write(
        "W260 — The manager itself does not use Qt networking. Query refresh is performed by desktop Excel and may contact configured sources. The package omits unused Qt Network/TLS bindings, plugins, and Qt OpenSSL copies.\n"
    )
with (DIST / "RELEASE_SUMMARY.txt").open("a", encoding="utf-8") as summary:
    summary.write(
        "\nW260 — Removed unused Qt Network/TLS runtime components; the app contains no QtNetwork imports, while query refresh remains delegated to Excel. Verifier coverage rejects these optional components. The 116-file, approximately 76.2 MB package passed CLI help and packaged GUI import/UIA on Excel 16.0 build 20430.0. External-source refresh and other Office builds remain unverified.\n"
    )

with (DIST / "README.txt").open("a", encoding="utf-8") as manual:
    manual.write(
        "W261 — Packaged connection-only and worksheet query refresh passed in background and synchronous modes using a loopback-only source; Excel, not the manager, made the HTTP requests.\n"
    )
with (DIST / "RELEASE_SUMMARY.txt").open("a", encoding="utf-8") as summary:
    summary.write(
        "\nW261 — After removing Qt Network/TLS, packaged CLI refresh passed on Excel 16.0 build 20430.0 for connection-only and worksheet loads in background and synchronous modes. Four reports returned `refreshed`; ten loopback requests were observed; exact formulas, immediate-pre-refresh backup hashes, worksheet value 42 and cleanup passed. No external network was used. Other providers, Data Model refresh on this exact package and other Office builds remain unverified.\n"
    )

with (DIST / "README.txt").open("a", encoding="utf-8") as manual:
    manual.write(
        "W262 — Current packaged GUI Data Model import passed on Excel 16.0 build 20430.0 after its capability probe; formula, model membership, row count, backup and cleanup were verified with static inline M.\n"
    )
with (DIST / "RELEASE_SUMMARY.txt").open("a", encoding="utf-8") as summary:
    summary.write(
        "\nW262 — Current compact GUI Data Model import passed on Excel 16.0 build 20430.0. Capability probe enabled the mode; read-only reopen verified exact formula, Data Model membership, two-row result and original-hash backup. Manager and owned Excel process exited; pre-existing Excel PID 36464 remained untouched. No external source was used. Data Model refresh and other Office builds remain unverified.\n"
    )

with (DIST / "README.txt").open("a", encoding="utf-8") as manual:
    manual.write(
        "W263 — Data Model refresh passed on Excel 16.0 build 20430.0; Excel exposes this API synchronously, with no background option.\n"
    )
with (DIST / "RELEASE_SUMMARY.txt").open("a", encoding="utf-8") as summary:
    summary.write(
        "\nW263 — Current packaged CLI refreshed a query loaded to the Excel Data Model on build 20430.0. Refresh returned `refreshed` in 3.031 seconds; model connection and two rows, staged verification, original-hash backup and read-only reopen passed. Excel's API is synchronous only; background Data Model refresh is unavailable. No external source was contacted. Other Office builds remain unverified.\n"
    )

with (DIST / "README.txt").open("a", encoding="utf-8") as manual:
    manual.write(
        "W264 — Packaged GUI replacement preserved an existing Data Model association and its one-row contents without implicitly refreshing it on Excel 16.0 build 20430.0.\n"
    )
with (DIST / "RELEASE_SUMMARY.txt").open("a", encoding="utf-8") as summary:
    summary.write(
        "\nW264 — Current packaged GUI replaced a Data Model query definition on Excel 16.0 build 20430.0. Read-only reopen verified the changed formula, retained `InModel` association, and unchanged one-row model contents, proving replacement did not implicitly refresh the model. Backup hash, manager exit, and owned-process cleanup passed; pre-existing Excel PID 36464 remained untouched. Other Office builds and workbook relationship variants remain unverified.\n"
    )

with (DIST / "README.txt").open("a", encoding="utf-8") as manual:
    manual.write(
        "W265 — Packaged GUI replacement preserved an existing worksheet query table and its data without implicitly refreshing it on Excel 16.0 build 20430.0.\n"
    )
with (DIST / "RELEASE_SUMMARY.txt").open("a", encoding="utf-8") as summary:
    summary.write(
        "\nW265 — Current packaged GUI replaced a worksheet-loaded query definition on Excel 16.0 build 20430.0. Read-only reopen verified the changed formula, original table identity and Output!C4 anchor, and unchanged one-row data. Backup hash, manager exit and owned-process cleanup passed; pre-existing Excel PID 36464 remained untouched. Other Office builds and workbook relationship variants remain unverified.\n"
    )

with (DIST / "README.txt").open("a", encoding="utf-8") as manual:
    manual.write(
        "W266 — Source comparison treats Excel-removed terminal line endings as unchanged while preserving exact M text in the diff.\n"
    )
with (DIST / "RELEASE_SUMMARY.txt").open("a", encoding="utf-8") as summary:
    summary.write(
        "\nW266 — Source-to-workbook comparison now treats terminal CR/LF removed by Excel during save as equivalent for status classification, while diffs continue to show exact original text. A failing-first regression reproduced the false MODIFIED state; comparison tests and the full 479-test suite passed, and the Windows onedir GUI/CLI rebuilt successfully.\n"
    )

with (DIST / "README.txt").open("a", encoding="utf-8") as manual:
    manual.write(
        "W267 — Current packaged GUI showed UNCHANGED for a CRLF source whose terminal line ending Excel had removed, then imported it with source bytes preserved.\n"
    )
with (DIST / "RELEASE_SUMMARY.txt").open("a", encoding="utf-8") as summary:
    summary.write(
        "\nW267 — The current compact packaged GUI classified a source ending CRLF as UNCHANGED when a disposable workbook contained the same formula without its terminal line ending, then completed replacement. On Excel 16.0 build 20430.0, read-only reopen verified formula roundtrip, original-hash backup, unchanged source SHA-256, manager exit and owned Excel cleanup; baseline PID 36464 remained untouched.\n"
    )

with (DIST / "README.txt").open("a", encoding="utf-8") as manual:
    manual.write(
        "W269 — Import profiles migrate unversioned/schema-0 documents that match the known v1 structure; ambiguous legacy fields and unknown versions are rejected.\n"
    )
with (DIST / "RELEASE_SUMMARY.txt").open("a", encoding="utf-8") as summary:
    summary.write(
        "\nW269 — Profile loading now migrates only unversioned/schema-0 documents whose fields already match the known v1 profile contract. It adds schema_version=1 in memory without modifying the source file; unknown versions, booleans and extra legacy fields are rejected. Failing-first migration regressions and the full 484-test Windows suite passed; one-folder GUI/CLI rebuilt. The packaged CLI validated a temporary schema-0 profile, returned `schema_version: 1`, exited 0, and left the original file at version 0.\n"
    )

with (DIST / "README.txt").open("a", encoding="utf-8") as manual:
    manual.write(
        "W270 — Profile schema versions must be JSON integers; floating-point `1.0` is rejected so a profile cannot save as v1 and later fail to load.\n"
    )
with (DIST / "RELEASE_SUMMARY.txt").open("a", encoding="utf-8") as summary:
    summary.write(
        "\nW270 — Tightened profile schema-version validation to reject floating-point values such as `1.0`; previously Python numeric equality allowed `save_profile` to emit a document that strict migration loading then rejected. The failing-first regression passed after the fix; focused profile tests and the full Windows suite passed, and GUI/CLI were rebuilt.\n"
    )

with (DIST / "README.txt").open("a", encoding="utf-8") as manual:
    manual.write(
        "W271 — Query exports shorten long filenames safely; the manifest keeps the complete original query name.\n"
    )
with (DIST / "RELEASE_SUMMARY.txt").open("a", encoding="utf-8") as summary:
    summary.write(
        "\nW271 — Folder and query-package exports now cap sanitized filename stems at 60 Unicode code points, keeping UTF-8/UTF-16 components within common filesystem limits while retaining original names in package manifests and collision suffixes in exported folders. A failing-first Windows regression reproduced the 255-character-name failure; three focused export/package tests pass. The refreshed onedir GUI/CLI build, packaged CLI help, packaged GUI keyboard traversal, schema validator, and release-package verifier passed. No Excel behavior changed.\n"
    )

with (DIST / "README.txt").open("a", encoding="utf-8") as manual:
    manual.write(
        "W275 — Failure details redact common Authorization header credentials and signed URL parameters such as `sig` and `X-Amz-Signature`; local diagnostic files continue to omit exception text.\n"
    )
with (DIST / "RELEASE_SUMMARY.txt").open("a", encoding="utf-8") as summary:
    summary.write(
        "\nW275 — Extended error sanitization to redact Bearer/Basic Authorization headers, client secrets, and signed URL parameters including `sig`, `X-Amz-Signature`, and `X-Amz-Credential`. A failing-first regression reproduced token leakage; all 13 focused redaction tests pass. The refreshed onedir GUI/CLI build and package verification passed. Local JSONL diagnostics still record only exception type/category metadata.\n"
    )

with (DIST / "README.txt").open("a", encoding="utf-8") as manual:
    manual.write(
        "W276 — Redaction also covers quoted JSON token fields such as `refresh_token` and `id_token`.\n"
    )
with (DIST / "RELEASE_SUMMARY.txt").open("a", encoding="utf-8") as summary:
    summary.write(
        "\nW276 — Fixed a failing-first regression showing quoted JSON `refresh_token`, `id_token`, and `client_secret` properties remained visible in error text. Redaction now accepts quoted field names and common compound token keys; all 14 focused redaction tests pass. The onedir GUI/CLI package is rebuilt and verified.\n"
    )

with (DIST / "README.txt").open("a", encoding="utf-8") as manual:
    manual.write(
        "W277 — Redaction recognizes compound credential field names such as `AWS_SECRET_ACCESS_KEY`, `x_api_key`, and `private_key`.\n"
    )
with (DIST / "RELEASE_SUMMARY.txt").open("a", encoding="utf-8") as summary:
    summary.write(
        "\nW277 — Expanded assignment redaction to compound credential key names including AWS secret/access keys, prefixed API keys, signing/encryption keys, and code verifiers. A failing-first regression reproduced three compound-field leaks; all 15 focused redaction tests pass. The onedir package is rebuilt and verified.\n"
    )

with (DIST / "README.txt").open("a", encoding="utf-8") as manual:
    manual.write(
        "W278 — Error details redact Authorization and Proxy-Authorization values for schemes including Basic, Bearer, Negotiate, and Digest.\n"
    )
with (DIST / "RELEASE_SUMMARY.txt").open("a", encoding="utf-8") as summary:
    summary.write(
        "\nW278 — A failing-first regression reproduced leakage from Negotiate and Digest authorization headers. Sanitization now redacts complete Authorization and Proxy-Authorization values for supported header forms. The focused redaction module passed 16 tests; the refreshed onedir package and its manifest were rebuilt and verified.\n"
    )

with (DIST / "README.txt").open("a", encoding="utf-8") as manual:
    manual.write(
        "W282 — Authorization values are redacted through the end of the header line, including semicolon-containing Digest parameters.\n"
    )
with (DIST / "RELEASE_SUMMARY.txt").open("a", encoding="utf-8") as summary:
    summary.write(
        "\nW282 — A failing-first regression found that a semicolon inside a Digest parameter could expose the remainder of that authorization value. Redaction now covers the full header line. The focused redaction module passed 17 tests; the refreshed onedir package and its manifest were rebuilt and verified.\n"
    )

with (DIST / "README.txt").open("a", encoding="utf-8") as manual:
    manual.write(
        "W293 — Packaged CLI query-package import verified three exact query definitions, including a Unicode name; the export source remained unchanged and the target backup matched its original hash.\n"
        "W294 — Packaged CLI rejected a ZIP path-traversal entry before workbook inspection; the target stayed unchanged and Excel was not launched.\n"
        "W295 — Packaged GUI import to a disposable VBA-bearing .xlsm preserved the VBA project and unchanged Workbook_Open marker, with exact formula and backup verification.\n"
        "W297 — Packaged CLI deployment applied to a writable disposable workbook and blocked a read-only target; the successful backup and formula were verified, and the blocked target remained unchanged.\n"
        "W298 — Excel workbook lifecycle calls now use an injectable adapter across inspection, import, refresh, and capability probing; focused fake-COM coverage and one rebuilt packaged GUI import on Excel 16.0 build 20430.0 passed.\n"
        "W299 — the packaged CLI refused deployment while Excel held a disposable workbook read-only without an owner-lock file; it returned blocked_preflight before changing bytes, making a backup/stage, or starting another Excel process.\n"
        "W300 — selected source-row removal is a single undoable manager edit; undo restores original row identities/order and redo removes them again. The 40-test UI model suite passed; the rebuilt onedir GUI completed packaged 35-control keyboard traversal and CLI help/package verification passed.\n"
        "W301 — packaged CLI plans and deployments write local audit events with aggregate counts and no query, source, or workbook identifiers.\n"
        "W302 — duplicate source paths are filtered within each file-add batch; the first selected row is preserved. The focused 41-test query-model suite passed. Current packaged definition-only GUI import passed on Excel 16.0 build 20430.0 with exact formula, verified backup, read-only reopen, clean exit and owned-process cleanup. Packaged keyboard traversal did not complete reliably and remains an open acceptance item.\n"
        "These checks used Excel 16.0 build 20430.0 on Windows 10 build 26300. Other Office builds and the broader workbook, provider, security, and accessibility matrix remain unverified.\n"
    )
with (DIST / "RELEASE_SUMMARY.txt").open("a", encoding="utf-8") as summary:
    summary.write(
        "\nW293 verified packaged CLI query-package import with three exact definitions including a Unicode name, unchanged source workbook on export, verified target backup, and no external source evaluation. W294 verified that the packaged CLI rejects a ZIP traversal entry before workbook inspection without changing the target or starting Excel. W295 verified current packaged GUI import to a disposable VBA-bearing .xlsm with exact formula, preserved VBA source, unchanged Workbook_Open marker, verified backup, and owned-process cleanup. W297 verified packaged CLI mixed-batch isolation: one writable target applied and a second read-only target returned blocked_preflight with unchanged bytes and no backup; the command returned the partial-failure exit code and owned Excel processes exited. W298 adds the injectable Excel workbook lifecycle port across inspector, importer, refresher, and load-capability probe; 91 focused fake-COM/service tests passed, the rebuilt package passed definition-only GUI import on Excel 16.0 build 20430.0, CLI help passed, and the compact manifest verified. W299 verified packaged CLI active-use refusal while Excel held a disposable workbook read-only without an owner-lock file; the workbook hash remained unchanged, no backup/stage was created, and no additional Excel process started. W300 makes multi-row source removal one undoable action and verifies original row identity/order across undo and redo. The focused 40-test UI model suite passed; rebuilt packaged GUI keyboard traversal (35 controls), CLI help, and package verification passed. W301 records packaged CLI dry-run plans and deployment outcomes in local history using action/status and aggregate counts only; focused tests confirm sensitive paths are excluded. Broader Windows/Office release gates remain UNVERIFIED.\n"
    )

# The GitHub Windows package workflow supplies the exact source SHA before
# building. Put a verifiable provenance record INSIDE the package, rather than
# shipping an unstamped binary that might be mistaken for current source.
if build_source_sha:
    gui_binary = TARGET / "PowerQueryManager.exe"
    cli_binary = TARGET / "PowerQueryManagerCLI.exe"
    provenance = {
        "source_commit": build_source_sha,
        "ui_design": "Vector Concept 03",
        "package_status": "development-preview",
        "gui_sha256": hashlib.sha256(gui_binary.read_bytes()).hexdigest(),
        "cli_sha256": hashlib.sha256(cli_binary.read_bytes()).hexdigest(),
    }
    (TARGET / "BUILD_PROVENANCE.json").write_text(
        json.dumps(provenance, indent=2, sort_keys=True) + "\n", encoding="utf-8"
    )
    with (DIST / "RELEASE_SUMMARY.txt").open("a", encoding="utf-8") as summary:
        summary.write(
            "\nBuild provenance: Vector Concept 03; "
            f"source commit {build_source_sha}; executable digests in "
            "PowerQueryManager/BUILD_PROVENANCE.json.\n"
        )

files = sorted(path for path in DIST.rglob("*") if path.is_file() and path.name != "SHA256SUMS")
manifest = "".join(
    f"{hashlib.sha256(path.read_bytes()).hexdigest()}  {path.relative_to(DIST).as_posix()}\n"
    for path in files
)
(DIST / "SHA256SUMS").write_text(manifest, encoding="utf-8")
print(f"Staged {len(files)} files; package size {sum(p.stat().st_size for p in files):,} bytes.")

# Keep the handoff package limited to launchable runtime and concise user docs.
from verify_release_package import verify_package
verify_package(DIST)

if not ARGS.docs_only:
    expected = {"PowerQueryManager", "README.txt", "RELEASE_SUMMARY.txt", "SHA256SUMS"}
    backup = FINAL_DIST.parent / f".{FINAL_DIST.name}.previous-{uuid.uuid4().hex}"
    moved_old = False
    if FINAL_DIST.exists():
        if not FINAL_DIST.is_dir() or {item.name for item in FINAL_DIST.iterdir()} != expected:
            shutil.rmtree(DIST)
            raise SystemExit("Refusing to replace a release directory containing unexpected files.")
        try:
            replace_directory(FINAL_DIST, backup)
            moved_old = True
        except OSError as exc:
            shutil.rmtree(DIST)
            raise SystemExit(f"Could not safely replace the current package; it remains untouched: {exc}") from exc
    try:
        replace_directory(DIST, FINAL_DIST)
    except OSError as exc:
        if moved_old:
            try:
                replace_directory(backup, FINAL_DIST)
            except OSError as restore_error:
                raise SystemExit(
                    f"Package publish failed ({exc}) and prior package restore failed ({restore_error}); "
                    f"prior package remains at {backup}."
                ) from restore_error
        shutil.rmtree(DIST, ignore_errors=True)
        raise SystemExit(f"Package publish failed; the prior package was restored: {exc}") from exc
    if moved_old:
        try:
            shutil.rmtree(backup)
        except OSError as exc:
            print(f"NOTE: prior package retained at {backup} because cleanup was blocked: {exc}")
    print(f"Published verified release package to {FINAL_DIST}")
