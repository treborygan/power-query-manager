# Power Query Manager changelog

This changelog records implementation milestones for the successor to Power Query Installer 1.3.0. It does not itself certify a release. See each [release evidence file](../evidence/releases/) for commands, observed results, limitations, and gate status.

## 2026-10-27 — Vector desktop redesign (feature branch, NOT RELEASED)

- Reflowed workspace actions and metrics for narrow windows; corrected a Finder Qt API test and retained inspector theme-toggle coverage for the Vector dark and light palettes.
- Increased contrast of unchecked checkbox outlines in both themes after inspecting the live light interface.
- Reflowed Dependency Finder in both scan modes and added no-horizontal-workspace-overflow coverage at laptop sizes.
- Reflowed the file graph heading, zoom controls and status so its minimum layout fits a 960×600 window.

- Applied the selected Vector concept to the existing PySide6 desktop UI with a graphite canvas, acid-lime primary actions, teal focus markers, compact sidebar, status metrics based on the original query model, and a Ctrl+K action palette that triggers the existing buttons.
- Reorganized workbook/source/import areas without replacing query models, source/target selectors, undo/redo, import safeguards, inspectors or history. Existing query-table and dialog action names remain available.
- Redesigned Dependency Finder with scrollable laptop-height layout, file and target panels, mode/depth controls, clearer results and coverage area, and original JSON/CSV/cancel behavior.
- Restyled both file and query dependency graphs. The file map adds zoom, pan, node selection and repeated-arrow visual deduplication, without altering the underlying graph/report. Accessibility text indexes remain.
- Updated source-side UI assertions for layout and behavior; **the modified Qt source has not been executed in this chat environment** because PySide6 and a Windows host are unavailable. Build, Windows test suite, GUI visual checks and packaged verification remain **PENDING** on the new branch; no release/binary claim is made.

## 2026-10-27 — portable Linux regression repair (Windows gates unchanged)

- Cleared the 11 previously recorded Linux failures without removing or skipping any failing assertions. COM Running Object Table tests can use fake COM on non-Windows hosts; the Windows security-zone preflight now has an injectable stream reader so inspector/import rejection paths are tested offline.
- Made path-preview assertions OS-neutral without loosening path identity checks, and supplied mandatory Shiboken placeholder paths in the isolated packaging test fixture while retaining the production verifier's checks.
- Added pinned Linux source-only test dependencies, including `jsonschema`. Full offscreen Linux suite: **526 passed, 10 skipped, 0 failed**. The 10 skipped Windows-only integrations still require their actual Windows/Excel/DAO QA runs.

## 2026-10-27 — Dependency Finder (post-v2 feature; accepted for integration **with two waivers**)

- Later Windows W296 validation supersedes the initial source-only evidence below: **537 Windows tests passed**; packaged GUI/CLI rebuilt; disposable real Excel Mashup/links and Access DAO 12.0 fixtures, local and OneDrive Quick/Deep, GUI filters/graph/JSON/CSV/cancellation, hash/timestamp/lock/temp cleanup and natural Excel process exit passed on the documented host.
- The owner waived **ACE DAO 16.0** (provider unregistered; NOT VERIFIED) and **mapped SMB/UNC** (no test mappings; NOT RUN) solely for merging PR #1. Neither is a PASS or a compatibility claim. The overall product release gates remain separate. See W296/W297 in `QA_WINDOWS_MATRIX.md`. Earlier bullet points below describe implementation-stage evidence at the time, not the latest final acceptance state.

- Added a read-only *reverse* dependency scan for incoming references to multiple selected targets across `.xlsx`, `.xlsm`, `.accdb` and `.mdb` consumer files.
- Added separate GUI folder/multi-file selection, worker progress/cancellation, confidence/coverage displays, JSON/CSV export, and `dependency-scan` CLI command.
- Access metadata is inspected by DAO on disposable local database copies. No SharePoint-specific integration; `.xls`/`.xlsb` remain explicitly unsupported.
- Portable tests pass; real Access Engine, Excel DataMashup, mapped Windows folders and packaged binaries still require smoke verification.
- Review hardening: moved folder inventory off the GUI thread, added zero-match CSV summaries and spreadsheet-formula escaping, made M query-name attribution token-aware, released DAO proxies before temp cleanup, and capped cumulative ZIP/XML metadata reads.
- Added **General Dependency Map** to index every selected Excel/Access file's incoming/outgoing relationships, allow instant file/direction filtering and show a visual advisory graph. `dependency-map` CLI exports the same edge/unresolved-reference index. Quick/deep modes allow formula scanning to be opt-in for large scans.
- Non-custom Excel ZIP/XML metadata is now incrementally parsed and pruned; streamed references are matched immediately so large worksheet formula lists are not retained in RAM. Bounded DataMashup XML remains an exception. Tests are portable/synthetic; live Windows/Access and packaged validation remain unverified.
- Final review hardened both modes against duplicate target basenames (ambiguous references do not become guessed edges), captured multiple distinct workbook/database references from one formula or saved query, and prevented raw SQL from being exported as a guessed filename. Repeated formula-edge objects are deduplicated incrementally.


## v2.0.0 — advanced tooling (partial implementation; gate **UNVERIFIED**)

- Added multi-workbook dry-run/deployment with per-workbook results, deployment bundles, source-path and M-literal mapping previews, supported parameter editing, query packages, local history, and an optional CLI.
- Hardened explicit CLI approval phrases and redacted operation history/diagnostics.
- Remaining acceptance includes broader workbook/provider behavior, damaged-history salvage, and the full Windows/security/accessibility matrix.
- Evidence: [v2.0.0](../evidence/releases/v2.0.0.md).

## v1.5.0 — Excel loads, dependencies, and refresh (partial implementation; gate **UNVERIFIED**)

- Added capability-probed definition-only, connection-only, new worksheet, existing worksheet/cell, and Data Model loads, plus preserve-existing replacement behavior.
- Added conservative M reference analysis and explicit selected/all refresh with timeout and diagnostics. Data Model refresh is synchronous in the observed Excel API.
- Current-package disposable Excel checks verified all load modes, refresh paths, Data Model refresh, and preservation of existing worksheet and model loads on Excel 16.0 build 20430.0. Other Office versions, complex model relationships, and external providers remain unverified.
- Evidence: [v1.5.0](../evidence/releases/v1.5.0.md).

## v1.4.0 — manager, inspection, and profiles (partial implementation; gate **UNVERIFIED**)

- Added the PySide6 query manager, read-only workbook inventory, source/workbook status and M diff/editor, query export, versioned profiles, dependency display, and package comparison.
- W266 fixed a false `MODIFIED` result when Excel removes terminal CR/LF from a saved formula; W267 verified `UNCHANGED` status and replacement through the packaged GUI on Excel 16.0 build 20430.0.
- W269 added a constrained migration from unversioned/schema-0 profiles matching the v1 contract; profile load does not rewrite source files and refuses unknown versions or ambiguous legacy fields.
- W270 made profile schema-version checks require an integer, preventing `1.0` from being saved as v1 and then rejected on reload.
- W271 bounded query export filenames for long and multibyte names while preserving their full names in manifests.
- Screen-reader, physical display, broader workbook, and other Office-version acceptance remain open.
- Evidence: [v1.4.0](../evidence/releases/v1.4.0.md).

## v1.3.1 — safety and repeatability (gate **UNVERIFIED**)

- Replaced risky direct edits with staged workbook changes, verified backups, read-only verification, guarded rollback, macro security handling, and owned Excel process cleanup.
- Added actionable failure reporting, redacted local diagnostics, all-skip no-op behavior, and repeatable one-folder Windows packaging.
- W275 extends user-facing error redaction to common Authorization headers and signed URL credentials; local diagnostics still omit exception messages.
- W276 adds redaction for quoted JSON refresh, identity, and client-secret token fields.
- W277 adds matching for compound AWS, API, private, signing, and encryption key field names.
- W278 redacts complete Authorization and Proxy-Authorization values for schemes including Negotiate and Digest.
- W279-W281 reaccept the current compact GUI's replace, mixed `.m`/`.pq` import, and skip-existing no-op paths on disposable Excel workbooks.
- W282 extends authorization redaction through semicolons inside Digest parameters.
- W283 reaccepts existing-query replacement on the exact W282 compact package and installed Excel build.
- W284 reaccepts Data Model import on the W282 package, with capability probe and post-save membership readback on the installed Excel build.
- W285 reaccepts explicit packaged Data Model refresh with synchronous completion, verified staged commit, and exact backup readback.
- W286 reaccepts existing-worksheet load to `Output!C4` with capability probing and value readback.
- W287 reaccepts the distinct connection-only destination state on the current compact package.
- W288 reaccepts new-worksheet load with capability probing and table value readback.
- W289 reaccepts definition-only behavior and dialog control discovery in the current packaged GUI.
- W290 reaccepts offline profile validation in the current packaged CLI without modifying the profile or starting Excel.
- W291 reaccepts packaged multi-workbook bundle apply/repeat with per-target backups and no-change results.
- W292 reaccepts packaged multi-workbook dry-run with unchanged target hashes and no workbook staging.
- Remaining acceptance includes additional Windows/Office versions and unresolved lifecycle, security, and accessibility matrix cases.
- Evidence: [v1.3.1](../evidence/releases/v1.3.1.md).


- W293 verified packaged CLI query-package import with three exact query definitions, a Unicode name, unchanged export source, exact target backup, and no external source evaluation on Excel 16.0 build 20430.0. Other Office builds and the broader acceptance matrix remain unverified.
- W295 reaccepted current packaged GUI import to a disposable VBA-bearing `.xlsm`: exact M, preserved VBA source, unchanged `Workbook_Open` marker, verified backup and process cleanup on Excel 16.0 build 20430.0. The broader macro/security matrix remains unverified.
- W296 refreshed the shipped manual and release summary to match the v1.3.1–v2.0.0 implementation scope and W293–W295 evidence. Docs-only package verification passed; broader release gates remain unverified.
- W297 verified current packaged CLI deployment isolation when one of two disposable targets is read-only: writable target applied, read-only target blocked before backup/write, original bytes preserved, and partial-failure exit code 5 returned.
- W298 adds an injectable Excel workbook lifecycle port across inspection, staged import, refresh and capability probing. Focused fake-COM/service tests and a rebuilt packaged GUI import passed; the broader Office matrix remains unverified.
- W300 makes bulk selected-row removal one undoable manager action. Undo restores the original row objects, order and stable IDs; redo removes them again. The focused UI model suite passed 40 tests.

- W301: the packaged CLI records dry-run plan and deployment results in the same local history store as the GUI. Entries contain action, status, workbook/query counts and aggregate outcomes; they exclude names, formulas and paths.
- W302: adding source files now filters duplicate paths within the same batch as well as against existing rows, keeping the first selected entry. Focused query-model suite: 41 passed. Rebuilt package verified; packaged keyboard smoke was inconclusive and is deferred.
