## 2026-10-27 — Vector Windows validation in progress

## 2026-10-28 — Vector responsive layout and graph label repair

- Rebuilt the packaged GUI/CLI after responsive metric/action layout changes and query graph label z-order repair. Windows build 26300; `app/build.ps1`: **541 passed**. Release verification passed (116 payload files / 76,330,010 bytes); CLI `--help` exited 0. GUI SHA-256: `145F60264FD182B96787ED42F405C513FD3A7274155FF7A1C9FA336756D0F2B5`; CLI SHA-256: `6B30D61A87A2A164949C35EFB8BEEE2226B40344741CD64BA974365F752F0382`.
- Exact tested application commit: `e934432e2f8f86b421bf51d2b7869fad5326318c`.
- The packaged app launched and was visually compared with the supplied Vector Concept 03 board. Main workspace, synthetic query table, five-node query graph and Finder empty state were observed; graph labels/status now render above the node cards. Finder accessibility tree exposes scan depth, file/target pickers, run/cancel, results, coverage and exports. This pass did not execute a packaged Finder scan or physically test the 125%/150% DPI and requested display resolutions.
- See `docs/QA_WINDOWS_MATRIX.md` for exact limitations and remaining gates. ACE DAO 16.0 and mapped SMB/UNC remain **WAIVED / NOT VERIFIED**. Merge gate remains open pending non-waived manual checks and review.

- 2026-10-28: exact tested source tree `50f567d4e190aa24f479574810cfd580c1178b1a` passed `app/build.ps1` with 541 tests; its app source/tests are unchanged from `3c7e4173e7515007db49510a2cb1aafe4fbb1c72`. The package verifier passed for 116 files / 76,328,840 bytes; GUI and CLI hashes are recorded in the QA matrix. Packaged CLI `--help` exited 0. The GUI was launched and its main workspace, command palette and Dependency Finder were visually inspected; no workbook was selected or modified. Previously recorded disposable Excel/DAO and packaged import/history checks still apply. The keyboard traversal passed once after two unexplained focus-escape failures. Populated graph/dialog interactions and actual OS DPI/resolution checks remain open; ACE DAO 16.0 and mapped SMB/UNC remain explicitly waived.

- A live light-theme inspection found low-contrast unchecked checkbox outlines. The Vector unchecked outline treatment was visually checked; the final package remains on the tested native checked-state rendering so its Windows check glyph stays visible.
- Resolved in the 2026-10-28 candidate: the Finder reflow fits both modes at 760×560, 960×600 and 1366×768 without horizontal workspace scrolling; the file graph controls fit laptop widths and nodes support Arrow/Home/End selection; the inspector moves below the workspace at 960×600. Focused regressions and the full Windows build passed.
- The initial Windows build stopped at 537 passed, 3 failed (inspector test expected the old default palette, Finder test used an invalid Qt method, and workspace horizontal overflow at laptop width). Tests and implementation were corrected without removing the overflow or theme assertions. Commit `d72c25440231def2c4bd91234523da72064d1456` passed 540 tests; current 541-test result and release verification are recorded above and in `docs/QA_WINDOWS_MATRIX.md`.
- The 91 UI tests passed separately with `QT_SCALE_FACTOR=1.25` and `1.5` using the offscreen platform. This is simulated Qt scaling, not OS-level display scaling or physical-display acceptance.

## W292 — 2026-10-23 — current-package multi-workbook dry run

- W282 packaged CLI `plan --dry-run` inspected two Excel-created disposable `.xlsx` targets with a static definition-only source. Both plans reported `create`; before/after workbook SHA-256 values matched. No backup or stage was created, and the Excel process set remained at baseline PID 36464 after the CLI run.
- Package: 116 files / 76,202,313 payload bytes; manifest SHA-256 `1EF8DD10D6D76D1B66BED9D7F96AD5409DF5C9B235944A11C34F5C4D013D36C3`.

## W291 — 2026-10-23 — current-package portable deployment bundle apply/repeat

- The W282 packaged CLI applied a `.pqmdeploy` bundle to two disposable `.xlsx` workbooks on Excel 16.0 build 20430.0; both targets reported `applied`, then both reported `no_changes` on repeat. Exact formula/load readback, two original-hash backups and disabled-missing-source omission passed. All harness-owned Excel PIDs exited; PID 36464 remained untouched.
- Package: 116 files / 76,202,313 payload bytes; manifest SHA-256 `1EF8DD10D6D76D1B66BED9D7F96AD5409DF5C9B235944A11C34F5C4D013D36C3`.

## W290 — 2026-10-23 — current-package offline CLI profile validation

- W282 packaged CLI `validate --profile examples/profiles/demonstration.profile.json` returned `valid: true`, schema version 1, with all three example source files valid. The profile SHA-256 was unchanged; the Excel process set did not change.
- Package: 116 files / 76,202,313 payload bytes; manifest SHA-256 `1EF8DD10D6D76D1B66BED9D7F96AD5409DF5C9B235944A11C34F5C4D013D36C3`.

## W289 — 2026-10-23 — current-package definition-only and dialog smoke

- The W282 packaged GUI created a definition-only query on a disposable `.xlsx`; exact formula, backup, read-only reopen and manager exit 0 passed. Operation report confirmed no connection or worksheet load. Owned Excel cleanup returned to baseline PID 36464.
- Packaged UI Automation also exposed named controls in query inspector, dependency graph and profile mapping dialogs; mapping edits were cancelled. Package: 116 files / 76,202,313 payload bytes; manifest SHA-256 `1EF8DD10D6D76D1B66BED9D7F96AD5409DF5C9B235944A11C34F5C4D013D36C3`.

- The W282 packaged GUI's capability probe enabled `new_sheet`; a disposable workbook created the query table at `GuiOutput!A1`. Read-only reopen verified exact formula, load state, and values 17/29; backup hash, manager exit 0, and owned Excel cleanup passed on Windows 10 build 26300 / Excel 16.0 build 20430.0. Baseline PID 36464 was untouched.
- Package: 116 files / 76,202,313 payload bytes; manifest SHA-256 `1EF8DD10D6D76D1B66BED9D7F96AD5409DF5C9B235944A11C34F5C4D013D36C3`. Other Office builds remain **UNVERIFIED**.

## W287 — 2026-10-23 — current-package connection-only load

- The W282 packaged GUI's capability probe enabled `connection_only`; a disposable workbook persisted an explicit connection. Read-only reopen verified the exact formula and connection-only state, with no worksheet table, no Data Model membership and no refresh. Backup hash, manager exit 0 and owned Excel cleanup passed on Excel 16.0 build 20430.0; baseline PID 36464 was untouched.
- Package: 116 files / 76,202,313 payload bytes; manifest SHA-256 `1EF8DD10D6D76D1B66BED9D7F96AD5409DF5C9B235944A11C34F5C4D013D36C3`. Other Office builds remain **UNVERIFIED**.

## W286 — 2026-10-23 — current-package existing-worksheet load

- The W282 packaged GUI's capability probe enabled `existing_sheet`; a disposable workbook load to `Output!C4` read back the exact formula, table anchor, and values 17/29. Unrelated marker, backup hash, manager exit 0 and owned Excel cleanup passed on Windows 10 build 26300 / Excel 16.0 build 20430.0. Baseline PID 36464 was untouched.
- Package: 116 files / 76,202,313 payload bytes; manifest SHA-256 `1EF8DD10D6D76D1B66BED9D7F96AD5409DF5C9B235944A11C34F5C4D013D36C3`. Other Office builds remain **UNVERIFIED**.

## W285 — 2026-10-23 — current-package Data Model refresh

- The W282 package imported and explicitly refreshed a static query to the Data Model through the GUI and packaged CLI. Excel reported synchronous completion in 3.047 seconds; connection and two model rows were verified. Staged workbook verification passed; the exact pre-refresh backup hash matched; no stage remained; owned Excel cleanup returned to baseline PID 36464.
- Package: 116 files / 76,202,313 payload bytes; manifest SHA-256 `1EF8DD10D6D76D1B66BED9D7F96AD5409DF5C9B235944A11C34F5C4D013D36C3`. No external source was used. Other builds remain **UNVERIFIED**.

## W284 — 2026-10-23 — current-package Data Model import

- The W282 packaged GUI's capability probe enabled Data Model mode; a disposable workbook imported the static query and read-only reopen verified the exact formula, model membership and row count 2. Backup hash, manager exit 0, and owned Excel cleanup passed on Windows 10 build 26300 / Excel 16.0 build 20430.0; baseline PID 36464 was untouched.
- Package: 116 files / 76,202,313 payload bytes; manifest SHA-256 `1EF8DD10D6D76D1B66BED9D7F96AD5409DF5C9B235944A11C34F5C4D013D36C3`. Other Excel builds and complex models remain **UNVERIFIED**.

## W283 — 2026-10-23 — current W282 package GUI replacement

- The current onedir GUI replaced a seeded query on a disposable `.xlsx`, non-elevated, Windows 10 build 26300 / Excel 16.0 build 20430.0. Exact formula, original-hash backup, unrelated marker, manager exit 0, and owned Excel process cleanup passed. The read-only verification process required forced cleanup; Excel returned to pre-test PID 36464, untouched.
- Package: 116 files / 76,202,313 payload bytes; `SHA256SUMS` SHA-256 `1EF8DD10D6D76D1B66BED9D7F96AD5409DF5C9B235944A11C34F5C4D013D36C3`.

## W282 — 2026-10-23 — redact semicolon-containing authorization values

- A failing-first test reproduced a leak when a quoted Digest authorization parameter contained `;`. The sanitizer now replaces the remainder of the authorization header line. Focused redaction module: **17 passed**.
- Windows 10 build 26300 / Python 3.11.9 / PyInstaller 6.13.0 onedir GUI/CLI rebuild passed. Package verifier, contract validation, packaged CLI help, and GUI keyboard traversal (35 controls) passed; full suite and Excel were not rerun for this sanitizer-only change.
- Compact package: 116 files / 76,202,313 payload bytes; `SHA256SUMS` SHA-256 `1EF8DD10D6D76D1B66BED9D7F96AD5409DF5C9B235944A11C34F5C4D013D36C3`.

## W281 — 2026-10-23 — current-package skip-existing no-op

- W278 compact GUI reported the seeded existing query as skipped and committed no workbook changes on Excel 16.0 build 20430.0. The disposable `.xlsx` hash/timestamp and original formula were unchanged; no backup or stage was created; manager exit 0; owned Excel process cleanup returned to baseline PID 36464.
- Existing Excel PID 36464 was untouched. Other formats and Office builds remain **UNVERIFIED**.

## W280 — 2026-10-23 — current-package mixed M/PQ import

- The W278 packaged GUI imported three numbered `.m`/`.pq` sources in the reviewed order on a disposable workbook under a non-elevated token, Windows 10 build 26300 / Excel 16.0 build 20430.0. Excel query order, exact formulas, backup, read-only reopen, manager exit 0, and owned Excel process cleanup passed. Existing Excel PID 36464 was left untouched.
- Package identity: 116 files / 76,201,871 payload bytes; `SHA256SUMS` SHA-256 `59616718721CF27ACA63CEF5CC80B7BB97C62E23420946C2E9E579FF861821DD`. This covers one ordering fixture and one Office build; broader compatibility remains **UNVERIFIED**.

## W279 — 2026-10-23 — W278 GUI replacement acceptance

- Current packaged GUI replacement passed on a disposable workbook; exact formula, backup, unrelated marker, manager exit 0, and baseline Excel process set verified. See release evidence and `docs/EXCEL_COM_SPIKES.md`.

## W278 — 2026-10-23 — complete authorization header redaction

- A failing-first test reproduced leaked credential material in `Proxy-Authorization: Negotiate …` and `Authorization: Digest …` error text. The sanitizer now removes complete authorization header values while leaving following headers intact.
- Focused redaction module: **16 passed**. PyInstaller 6.13.0 rebuilt GUI/CLI on Windows 10 build 26300 / Python 3.11.9. Contract validation, packaged CLI help, package verification, and GUI keyboard traversal (35 controls) passed. No workbook or Excel instance was opened; the full suite was not rerun.
- Compact package: 116 files / 76,201,871 payload bytes; `SHA256SUMS` SHA-256 `59616718721CF27ACA63CEF5CC80B7BB97C62E23420946C2E9E579FF861821DD`.
- Excel/Office, wider header formats, and overall safety/security release gates remain **UNVERIFIED**.

## W277 — 2026-10-23 — compound credential key redaction

- A failing-first regression reproduced leakage of `AWS_SECRET_ACCESS_KEY`, `x_api_key`, and `private_key` fields. The sanitizer now matches compound secret-bearing key names while retaining common error text outside those assignments.
- Focused redaction module: **15 passed**. PyInstaller GUI/CLI rebuild, contract/package verifiers, CLI help, and current packaged GUI keyboard traversal (35 controls) passed. No workbook or Excel instance was opened; full suite was not rerun.
- Compact package: 116 files / 76,201,426 payload bytes; `SHA256SUMS` SHA-256 `2B34B9FE9AE37DB2FB283A5FFC44B4E47F9DCABEFB2F1ABD3691B5034CA34213`.

## W276 — 2026-10-23 — redact quoted JSON credential fields

- A failing-first regression reproduced leakage of quoted JSON `refresh_token`, `id_token`, and `client_secret` values. The sanitizer now accepts quoted property names and recognizes common compound token fields.
- Focused redaction module: **14 passed**. PyInstaller GUI/CLI rebuild, package/contract verifiers, packaged CLI help, and GUI keyboard traversal (35 controls) passed. No workbook or Excel instance was opened; the full suite was not rerun.
- Compact package: 116 files / 76,200,852 payload bytes; `SHA256SUMS` SHA-256 `A5269C0AE1DE3AD6523DEE04F2D01A1FB3EE9F7385DFEFF9236845639DAB9EF1`.

## W275 — 2026-10-23 — redact authorization and signed-URL secrets

- Failing-first regression reproduced common secret leakage from HTTP Bearer/Basic authorization headers, `client_secret`, `sig`, and `X-Amz-Signature` values in formatted errors. Sanitization now covers those forms plus subscription/shared-access keys and common signed URL credential fields. Structured local diagnostics still omit exception text entirely.
- Focused redaction module: **13 passed**. PyInstaller 6.13.0 rebuild passed on Windows 10 build 26300 / Python 3.11.9; contract validation, packaged CLI `--help`, package verifier, and current packaged GUI keyboard traversal (35 controls) passed. Full suite and Excel COM were not rerun for this sanitizer-only change.
- Compact package: 116 files / 76,200,376 payload bytes; `SHA256SUMS` SHA-256 `0BBEF48C55C8B386A71E225330A497B218F1CC2A1044AD503502B8B003E99711`.

## W274 — 2026-10-23 — query-package long-name manifest coverage

- Added a focused acceptance test for a 255-character Unicode query name in a `.pqm.zip`: the generated filename component stays within 255 UTF-8 bytes and the manifest preserves the full query name.
- Targeted test: **1 passed**. No runtime source or package binary changed; the current W271 package remains verified at 116 files / 76,199,132 payload bytes.

## W273 — 2026-10-23 — current compact GUI Excel replacement

- The W271 compact package replaced a seeded query on a disposable workbook using `tools/run_packaged_gui_import_smoke.py --replace-existing`, under a non-admin token on Windows 10 build 26300 / Excel 16.0 build 20430.0.
- Exact new formula, original-hash backup, unrelated worksheet marker, manager exit 0, and cleanup to baseline Excel PID `[36464]` passed. The manager-owned Excel PID exited gracefully; the separate read-only verification PID required forced cleanup. Existing PID 36464 was untouched; no refresh/external source was used.
- Current package identity remains 116 files / 76,199,132 payload bytes; manifest SHA256 `479F59F9AF3EDFB86720FF518A084F1B66AFA69F51DEE8A2D9FD9678EFB15135`.

## W272 — 2026-10-23 — rebuilt packaged profile migration

- Current compact CLI validated a temporary `schema_version: 0` profile, reported `schema_version: 1` and `valid: true`, exited 0, and left the input profile bytes unchanged. No workbook or Excel instance was opened.
- The package used here is the verified W271 package: 116 files / 76,199,132 payload bytes; manifest SHA256 `479F59F9AF3EDFB86720FF518A084F1B66AFA69F51DEE8A2D9FD9678EFB15135`.
- Overall release gates remain **UNVERIFIED**.

## W271 — 2026-10-23 — portable export filenames

- A failing-first Windows regression reproduced `OSError [Errno 22]` when a valid 255-character query name was exported as a filename. Folder and query-package exports now share a safe 60-code-point stem helper; original query names remain in manifests, and folder collision suffixes still apply.
- Focused export/package selection: **3 passed**. PyInstaller onedir GUI/CLI rebuild, packaged GUI keyboard traversal (35 controls), CLI `--help`, contract validation, and release-package verification passed. No Excel behavior changed.
- Final compact package: 116 files / 76,199,132 payload bytes; `SHA256SUMS` SHA-256 `479F59F9AF3EDFB86720FF518A084F1B66AFA69F51DEE8A2D9FD9678EFB15135`.
- Full release gates remain **UNVERIFIED**.

## W270 — 2026-10-23 — strict integer profile version

- `validate_profile` now rejects floating-point schema versions; Python previously equated `1.0` with integer `1`, allowing save output that migration loading then refused.
- A failing-first regression reproduced the inconsistency. Focused tests: **7 passed** after fix. `app/build.ps1` completed successfully, including its full test suite and PyInstaller GUI/CLI build. Refreshed package: 116 files / 76,198,528 bytes; package verifier, contracts, and packaged CLI `--help` passed. Manifest SHA256: `34F511116CE1F8B02BC6BF34627361DCC6D836A88D3637CCAC8DA59F708A124E`.
- Rebuilt packaged GUI replacement passed on Windows 10 build 26300 / Excel 16.0 build 20430.0 under a non-admin token. Exact formula readback, pre-import backup, marker preservation, manager exit 0, and cleanup to the baseline Excel PID set `[36464]` passed. Existing PID 36464 was untouched.
- No Excel behavior changed; release gates remain **UNVERIFIED**.

## W269 — 2026-10-23 — constrained profile schema migration

- Implemented the missing M-04 compatibility path: loader migrates unversioned or `schema_version: 0` profile documents that already match the v1 shape, without mutating their files. Unknown versions, booleans, and extra/ambiguous legacy fields are rejected; profile save still emits v1.
- Failing-first migration tests reproduced both unversioned and v0 rejection. Focused profile migration/validation selection: **7 passed**. `app/build.ps1` passed 484 tests and rebuilt GUI/CLI; package and contract verification passed.
- No Excel behavior changed. Release gate remains **UNVERIFIED**.

## W268 — 2026-10-23 — milestone changelog

- Added `docs/CHANGELOG.md` covering the approved v1.3.1, v1.4.0, v1.5.0, and v2.0.0 roadmap; it separates implemented work from remaining acceptance and links each milestone to its release evidence. Added a source manual link. Release gates are not promoted, and no Git commit is claimed.
## W267 — 2026-10-23 — packaged GUI terminal-newline comparison acceptance

- Current compact packaged GUI inspected a disposable workbook whose query formula had no terminal line ending while its matching `.m` source ended CRLF. The query grid showed `UNCHANGED`; GUI replacement then completed.
- Excel 16.0 build 20430.0 read-only reopen verified the formula, exact pre-import backup, unchanged source SHA-256, manager exit 0 and cleanup to baseline process set `[36464]`. Existing PID 36464 was untouched; no external data source was contacted.
- Other Office builds, providers and remaining security/accessibility matrix items remain **UNVERIFIED**. No release gate is promoted.

## W266 — 2026-10-23 — comparison status matches Excel terminal-newline roundtrip

- A failing-first regression showed source-to-workbook comparison returned MODIFIED when Excel had removed only the source's terminal CR/LF. Status hashing now ignores terminal line endings after normalizing CRLF/CR; exact M diff/source text remains unchanged.
- Focused comparison tests: 3 passed. `app/build.ps1`: 479 passed; PyInstaller rebuilt the one-folder GUI and CLI on Windows 10 build 26300 / Python 3.11.9 / PyInstaller 6.13.0. Real Excel import/compare smoke was not repeated because this change affects portable status hashing only.
- v1.4.0 and overall release gates remain **UNVERIFIED** pending the complete Windows and accessibility matrix. No Git commit is available in this recovered checkout.

## W265 — 2026-10-23 — preserve worksheet load during GUI replacement

- The current packaged GUI replaced a worksheet-loaded query definition on a disposable workbook in Excel 16.0 build 20430.0. Read-only reopen verified the exact new formula, same `tblGuiSmoke` table at `Output!C4`, and original one-row value 5, showing no implicit refresh or relocation.
- Original-hash backup, unrelated marker, manager exit 0 and smoke-owned Excel cleanup passed. The verification Excel process required forced cleanup; the process set returned to `[36464]`, leaving pre-existing Excel PID 36464 untouched. No external source was contacted.
- Harness syntax check passed. Other Office builds and complex workbook/table relationships remain **UNVERIFIED**; no release gate is promoted.

## W264 — 2026-10-23 — preserve Data Model association during GUI replacement

- The latest packaged GUI replaced the M definition of a Data Model query on a disposable workbook under Excel 16.0 build 20430.0. Read-only reopen verified the new formula, retained `InModel` association and model table, and unchanged one-row model contents, confirming replacement did not refresh the changed definition implicitly.
- Exact original-hash backup, unrelated marker, manager exit 0 and smoke-owned Excel cleanup passed. Process set returned to `[36464]`; pre-existing Excel PID 36464 was untouched. No external source was contacted.
- Harness syntax check passed. Broader Data Model relationships and other Office builds remain **UNVERIFIED**; no release gate is promoted.



## W197 — 2026-10-23 — BOM/CRLF M source round-trip behavior

- A failing-first fake-COM regression showed staged verification rolled back because Excel strips a terminal CR/LF from a saved M formula. Formula comparison now allows only that Excel-observed terminal line-ending normalization while preserving exact checks for internal line endings and M text.
- On Windows 10 build 26300 / Excel 16.0 build 20430.0, a direct disposable workbook probe confirmed `Queries.Add` preserves internal CRLF across add/save/read-only reopen and removes terminal CRLF. The rebuilt packaged GUI imported BOM-prefixed CRLF `.m` through both create and existing-query replacement paths on separate disposable workbooks; source SHA-256, formula readback, backup, read-only reopen, manager exit 0 and owned-process cleanup passed. Create additionally verified the UIA completion log. Existing Excel PID 36464 was untouched.
- `app/build.ps1`: **420 tests passed**, PyInstaller onedir rebuild passed. Wider host/workbook acceptance and all release gates remain unverified.

## W198 — 2026-10-23 — fail closed if workbook vanishes at final commit

- A failing-first Windows test proved the prior `os.replace` path installed the stage as a newly created workbook when the selected target disappeared at the replacement boundary. The commit path now uses native `ReplaceFileW`, which requires an existing target, retains a temporary rollback copy, checks old/new hashes and attempts recovery from a reported partial replacement. Focused transaction/crash tests pass.
- A real Windows named-stream regression confirms the existing `Zone.Identifier` survives when the staged copy lacks that stream. The rebuilt GUI's existing-query replacement passed on disposable Excel 16.0 build 20430.0 with exact formula, backup, marker, app exit and owned-process cleanup. PID 36464 was untouched.
- The exact published package then passed BOM-prefixed CRLF source replacement against a disposable Excel workbook, including unchanged source hash, formula readback under the terminal-newline rule, backup, UIA completion, and owned-process cleanup.
- `app/build.ps1` passed **421 tests** and rebuilt GUI/CLI; a fresh full suite including the stream test passed **422 tests**. Partial native replacement failures and broader Windows policies remain unverified.

## W199 — 2026-10-23 — clean onedir payload comparison

- A second `app/build.ps1` run passed **422 tests** and rebuilt the onedir package using Windows 10 build 26300, Python 3.11.9, PyInstaller 6.13.0 and PySide6 6.8.3. Comparing against the published W198 package found **227 runtime files; 0 missing, 0 extra, 0 hash differences**.
- This establishes same-host/toolchain payload repeatability after W198. It does not establish cross-host reproducibility; Excel matrix and product release gates remain unverified.

## W200 — 2026-10-23 — partial ReplaceFileW recovery coverage

- Added deterministic Windows tests for API error 1177 partial replacement and a subsequent restore denial. Successful recovery confirms the original hash is back at the selected path and the stage remains; failed recovery preserves and names the original-hash copy. The restore uses real `MoveFileW`; partial state is injected.
- Full build: **424 passed**; transaction suite: **18 passed**. Freshly rebuilt GUI passed BOM/CRLF existing-query replacement on disposable Excel 16.0 build 20430.0; formula readback, source hash, backup, UIA operation completion and cleanup verified. Only existing PID 36464 remained.
- OS-generated partial failures and broader Excel matrix remain unverified.

## W127 — 2026-10-22 packaged GUI smoke cleanup race

- Fixed the smoke harness to prioritize its process handle and tolerate a vanished HWND after app exit; it still verifies manager exit code and independent workbook readback. Rebuilt packaged GUI replacement rerun passed on Excel 16.0 build 20430.0 with formula, backup, marker and process cleanup checks.

- Final W126/W127 package docs refresh verified at 229 files / **135,738,985 bytes**.


## W126 — 2026-10-22 package comparison history counts

- Failing-first regression proved five safe aggregate counters were lost by the history metadata allowlist. Added those counts while keeping formulas and source paths filtered. History tests: 34 passed. Full Windows build: **360 passed in 20.46s**; PyInstaller rebuilt GUI/CLI. Packaged GUI replacement smoke passed on Excel 16.0 build 20430.0. Final package: 229 files / **135,738,985 bytes**; package, contract and handoff verification passed.


## W128 — 2026-10-22 import outcome and mapping audit counts

- Successful import history now includes created/replaced/skipped aggregate counts and fixed import action/status; profile mapping history retains path/environment mapping counts. Regression verifies no query names, formulas, or source paths are recorded. Focused history/UI tests: 35 passed. Full build: **361 passed in 19.43s**; GUI/CLI rebuilt. Packaged replacement smoke passed on Excel 16.0 build 20430.0. Package: 229 files / **135,739,551 bytes**; package, contract and handoff checks pass.

- Final package summary/hash refresh after W128: 229 files / **135,739,551 bytes**; package verifier passed.


## W129 — 2026-10-22 — capability probe cleanup and packaged Data Model load

- Repaired capability-probe cleanup so a failed COM `Quit()` is followed by a bounded wait for the exact `DispatchEx` PID, termination of only that PID if it remains alive, and a second exit check. If exact-process cleanup cannot be confirmed, load modes stay disabled. Two failing-first fake-COM regressions cover already-exited and still-running owned processes.
- Windows onedir build completed with the full suite: **363 passed in 18.81s**; rebuilt GUI/CLI. Distribution refresh and verifier: **229 files / 135,742,675 bytes**, all manifest hashes valid.
- Source Excel load matrix passed on Windows 10 build 26300 / Excel 16.0 build 20430.0 for connection-only, new worksheet, existing worksheet, and Data Model; verified create/repeat/changed-formula states, read-only reopen, output values where applicable, and owned-process cleanup.
- Packaged GUI Data Model smoke passed on the same build: capability probe enabled Data Model; static inline-M formula, model connection/table membership and row count 2, verified pre-import backup hash, manager exit 0, and cleanup of smoke-owned Excel processes. Preexisting Excel PID 36464 was untouched. No external source was evaluated.
- Limits: one Office build and disposable workbook. This does not close other Office builds, non-default workbook security settings, broad Data Model/PivotTable relationships, external connectors, or the full v1.5.0/v1.3.1 matrix. Overall release gates remain UNVERIFIED / no release.


## W130 — 2026-10-22 — hold owned Excel process handles through timeout cleanup

- Added an `OwnedExcelProcess` wrapper that opens a Windows process handle with `PROCESS_TERMINATE | SYNCHRONIZE`, waits and terminates using that same kernel handle, and closes the handle exactly once. Bounded calls hold the handle while the callback runs and through the termination wait. Capability probing retains its process handle from DispatchEx PID capture through shutdown; refresh timeout calls also carry the handle to their injected cleanup boundary. This prevents a later PID reuse from redirecting a timeout termination after the process handle has been acquired.
- Failing-first regressions proved the old API rejected a captured-handle target and lacked the handle wrapper. Focused lifetime/capability/refresh suite: **32 passed**. Full portable Windows suite: **365 passed in 15.34s**. `app/build.ps1` passed its embedded suite (**365 passed in 19.36s**) and rebuilt the onedir GUI/CLI using Windows 10 build 26300, Python 3.11.9, PyInstaller 6.13.0 and PySide6 6.8.3.
- Repackaged distribution: **229 files / 135,746,035 bytes**; package verifier passed. Packaged CLI `--help` exited 0. Packaged GUI Data Model smoke and definition-only control smoke passed on Excel 16.0 build 20430.0; formula/model/row-count, original-hash backup, manager exit 0 and process cleanup verified.
- Scope limit: unit tests verify that wait/termination reuse one kernel handle; live Excel smokes verified successful cleanup but did not force termination of a blocked COM call. The brief interval between resolving the Excel window PID and acquiring its process handle, genuinely hung COM calls, other Office builds, and unrelated workbook variants remain unverified. No release gate is green.


## W131 — 2026-10-22 — packaged GUI acceptance for all supported query load modes

- On Windows 10 build 26300 / Excel 16.0 build 20430.0, ran the rebuilt packaged GUI harness on separate disposable workbooks for connection-only, new worksheet, existing worksheet (`Output!C4`), and Data Model loads. All four passed the runtime capability probe and exact post-save formula/load-state readback. Connection-only had no output worksheet or model table; worksheet values 17/29 and destination anchors were verified; Data Model connection/table membership and row count 2 were verified. Each test verified the original-hash backup, manager exit 0 and owned Excel cleanup. Existing Excel PID 36464 remained untouched.
- This reruns the full supported-mode packaged GUI matrix against the W130 build. Static inline M only; no external source evaluation. It verifies one Office build and the tested disposable workbook shapes only; other Office builds and complex workbook/Data Model relationships remain open.

- Distribution verifier after the W131 documentation refresh: **229 files / 135,746,212 bytes**; all manifest entries verified.


## W132 — 2026-10-22 — main-window interactive control accessible names

- Added a widget-tree regression that checks visible main-window buttons, combo boxes, line edits, text editors, spin boxes, and table views for explicit accessible names. The failing-first check found Qt's query-table corner select-all button was unnamed. The button now exposes the accessible name “Select all query source table cells” while preserving its select-all behavior.
- Focused accessibility/layout tests passed. All **60 UI tests** passed at `QT_SCALE_FACTOR=1.25` and again at `1.5`. Full Windows build passed **366 tests in 17.49s** and rebuilt GUI/CLI; packaged GUI UI Automation and disposable `.xlsx` definition-only import passed on Windows 10 build 26300 / Excel 16.0 build 20430.0 with formula, backup, manager exit and owned-process cleanup readback.
- This checks the main-window Qt widget tree and an automated UIA flow. It is not a screen-reader user study, OS-level display scaling, physical laptop review, or verification of every modal dialog.

- Distribution after W132 docs refresh: **229 files / 135,746,568 bytes**; release package verifier passed.


## W133 — 2026-10-22 — packaged UIA select-all action

- Windows UI Automation did not expose Qt's internal table-corner button even after an accessible name was set. Disabled that corner control and added a normal, labeled `Select all query rows` button that calls `QTableView.selectAll`; the widget regression confirms it selects two test query rows. The packaged smoke now requires this UIA name/role and passed against the rebuilt executable.
- Full Windows build passed **367 tests in 20.56s** and rebuilt onedir GUI/CLI. UI suite: **61 passed** at `QT_SCALE_FACTOR=1.25`; **61 passed** at `1.5`. Packaged Windows UI Automation recognized the explicit select-all button; packaged definition-only import passed on Excel 16.0 build 20430.0 with exact formula, verified backup, read-only reopen, manager exit 0 and owned-process cleanup.
- Scope: main-window select-all control and tested import path. This does not establish screen-reader usability, physical laptop layout, OS-level scaling or modal dialog accessibility.

- Final package refresh after W133: **229 files / 135,746,743 bytes**; all hashes verified.


## W134 — 2026-10-22 — validate HWND owner after acquiring process handle

- `open_owned_excel_process` now accepts the DispatchEx window handle, opens the process handle for the captured PID, then re-queries the HWND owner before allowing a bounded COM callback. If the HWND no longer belongs to that PID or its owner cannot be verified, the new handle is closed and the operation fails closed. `call_with_timeout` and capability probing pass the HWND; refresh stores and passes the owned Excel HWND with its PID. Subsequent wait/termination continues to use the held kernel process handle.
- Failing-first test confirmed the old helper did not accept an HWND; a regression now proves changed HWND ownership rejects the operation and closes the just-opened handle. Focused lifetime/capability/refresh/importer checks: **34 passed**. Full suite: **368 passed in 18.77s**; Windows build's embedded suite: **368 passed in 23.35s** and GUI/CLI rebuilt.
- Repackaged distribution verifier passed; packaged CLI `--help` exited 0. Packaged GUI Data Model import passed on Excel 16.0 build 20430.0 with exact formula, model membership/row count 2, matching backup, manager exit 0 and owned Excel cleanup. Preexisting Excel PID 36464 remained untouched; no external source was evaluated.
- Scope limit: Win32 provides no atomic HWND-to-process-handle acquisition. The post-open HWND owner recheck rejects observed owner changes, and a held process handle prevents later PID reuse from retargeting termination. Simultaneous reuse of both PID and HWND before the recheck, forced termination of a genuinely blocked COM call, other Office builds and broader workbook relationships remain unverified.


## W135 — 2026-10-22 — packaged CLI refresh after HWND ownership recheck

- Ran `tools/run_excel_cli_refresh_smoke.py --cli-exe dist/PowerQueryManager/PowerQueryManagerCLI.exe` on Windows 10 build 26300 / Excel 16.0 build 20430.0. Both background and synchronous refresh returned `refreshed` for a disposable workbook with static inline M.
- Each pre-refresh backup matched the exact workbook hash; post-refresh read-only inspection verified the exact formula and connection-only load. The CLI process-set check confirmed every smoke-owned Excel process exited. No external source was contacted.
- This verifies the W134 HWND-to-process ownership check on both CLI refresh paths for one Excel build. It does not simulate process/window ID reuse or a genuinely blocked COM call; other Office builds remain open.

- Final W135 distribution: **229 files / 135,748,386 bytes**; SHA256 manifest verified.


## W136 — 2026-10-22 — accessible names and row actions in modal dialogs

- Added a dialog widget-tree regression covering the query inspector, package update preview, profile mappings, history comparison and dependency graph. It found unnamed M editors/actions, mapping controls/tab category, and corner controls in selected tables. Named the source/workbook/diff editors, inspector actions, close actions, profile mapping tabs/actions, dependency diagram, and dialog button box; disabled inaccessible table-corner controls. Profile mapping tables now include named select-all path and M-source buttons; a regression verifies all rows select.
- Failing-first dialog accessibility test initially failed on the missing accessible names. Focused dialog/profile tests pass. Full suite: **370 passed in 17.64s**; UI suite: **63 passed** at Qt scale 1.25 and again 1.5. Windows build embedded **370 passed in 19.16s**, rebuilt GUI/CLI.
- Package verifier passed after rebuild; packaged CLI `--help` exited 0. Packaged UI Automation and disposable Excel 16.0 build 20430.0 definition-only import passed with formula, backup, reopen, app exit and process cleanup verification.
- Scope: Qt widget-tree naming and row selection are covered for the tested modal dialogs; packaged UIA specifically verified the main window, not each modal. Screen-reader user testing and physical display acceptance remain open.

- Final W136 distribution refresh: **229 files / 135,750,045 bytes**; package verifier passed.


## W137 — 2026-10-22 — packaged modal accessibility verification

- Extended `tools/run_packaged_gui_import_smoke.py` to inspect the packaged query inspector and both profile-mapping categories through Windows UI Automation. It verifies names and roles for the M source/workbook/diff editors, search and actions; mapping tables, tabs, row actions and save/cancel buttons. Profile edits were cancelled.
- Packaged Windows UIA and definition-only import passed on Windows 10 build 26300 / Excel 16.0 build 20430.0. Disposable `.xlsx` formula readback, original backup, read-only reopen, manager exit 0 and owned Excel cleanup all passed. Pre-existing Excel PID 36464 was not touched.
- Scope: package update, history comparison and dependency graph are covered by Qt widget tests but were not inspected through packaged UIA. Screen-reader user testing, physical laptop review and OS-level display scaling remain open; v1.4.0 gate remains **UNVERIFIED**.


## W138 — 2026-10-22 — packaged dependency graph accessibility verification

- Extended the packaged UIA smoke to open Query dependencies and check the named diagram (`Group`), text summary (`Edit`), dialog action group and named Close button. UIA exposed the role as Group, which is now reflected in the smoke assertion.
- Packaged UIA and the full disposable `.xlsx` definition-only import passed on Windows 10 build 26300 / Excel 16.0 build 20430.0. Formula readback, backup, read-only reopen, manager exit and owned Excel cleanup passed. Pre-existing Excel PID 36464 remained untouched.
- The package update and history comparison dialogs remain to be checked through packaged UIA. Screen-reader use, physical laptop review and OS-level display scaling remain unverified; v1.4.0 gate remains **UNVERIFIED**.


## W139 — 2026-10-22 — packaged query-package preview accessibility

- The packaged smoke now generates two local `.pqm.zip` inputs, drives the Compare package versions action, and checks the summary, change table, M diff, no-workbook-changes notice and Close button by Windows UIA. The full packaged disposable `.xlsx` import flow also passed on Excel 16.0 build 20430.0 with exact formula readback, verified backup, read-only reopen and owned-process cleanup.
- The packaged query inspector, dependency graph and profile-mapping dialogs also passed UIA in the same run. Mapping changes were cancelled. Pre-existing Excel PID 36464 was left untouched.
- The local history comparison dialog remains unverified through packaged UIA: the smoke seeded two disposable snapshots but the comparison dialog did not appear within 90 seconds. The manager exited during harness cleanup; no Excel process beyond PID 36464 remained. This is not a pass. The v1.4.0 gate remains **UNVERIFIED**.


## W140 — 2026-10-22 — marshal history worker callbacks to the UI thread

- Root cause: history compare and restore worker signals were connected to local Python closures. A QThread characterization showed plain callable callbacks execute on the emitting thread. The compare completion wrote its audit event from the worker thread and failed to show the comparison dialog; restore callbacks also accessed widgets there.
- Added an owned QObject slot bridge for compare/restore success, failure, and thread completion. Failing-first real-QThread regressions caught both comparison-dialog and restore-visibility callbacks running off the Qt application thread. All focused history tests pass; full portable suite: **372 passed**. UI suite: **65 passed** at Qt scale 1.25 and 1.5. Windows onedir build passed **372 embedded tests** and rebuilt GUI/CLI.
- Combined packaged UIA verified the query inspector, dependency graph, profile mappings, package update preview and workbook history comparison controls. The same run passed disposable `.xlsx` definition-only import on Excel 16.0 build 20430.0 with exact formula readback, verified backup, read-only reopen, manager exit 0 and owned Excel cleanup. Only pre-existing Excel PID 36464 remained.
- Packaged CLI `--help` exits 0. W140 distribution: **229 files / 135,750,972 bytes**; package verifier passed. Handoff and contract checks pass. Other Office builds, screen-reader use, and physical display acceptance remain open; v1.4.0 remains **UNVERIFIED**.
- Follow-up packaged history-restore smoke passed through Windows UI Automation on Excel 16.0 build 20430.0: the disposable workbook matched its selected snapshot SHA-256 after restore, and the generated pre-restore backup matched the exact pre-restore workbook hash. The same run completed definition-only GUI import, formula readback, backup verification, read-only reopen and manager/owned Excel cleanup. Only pre-existing Excel PID 36464 remained.




## W142 — 2026-10-22 — packaged keyboard path and history recovery

- Packaged main-window keyboard smoke passed: Tab traversal remained inside the app, completed a full cycle through 35 focusable controls, and reached workbook path, Add files, query table, replace/skip policy, backup and import action in expected order. It opened no workbook and no Excel process.
- Packaged history UIA smoke passed with intentionally corrupt SQLite bytes under isolated temporary `%LOCALAPPDATA%`: original corrupt bytes were preserved as one `history.corrupt-*` archive, a fresh store was initialized, archive recovery was reported in the UI, and manager exited 0.
- These checks cover the tested packaged UI flows only; they do not replace screen-reader, physical laptop, broader Excel matrix or release-gate acceptance. Release gates remain **UNVERIFIED**.
- Docs-only package refresh and integrity verification passed: 229 files / 135,751,359 bytes. Contract validation and handoff verification also passed.


## W143 — 2026-10-22 — query-bearing staged recovery for both workbook containers

- Improved `tools/run_excel_recovery_smoke.py` with an explicit `.xlsm` container option and matching stage/output extensions. Packaged app source code was unchanged; this is a QA harness extension.
- The query-bearing `.xlsx` smoke passed: one query recovered, exact M formula reopened read-only, original SHA-256 unchanged, stage preserved, output matched stage hash, and all two recovery-owned Excel processes plus seed process exited.
- The `.xlsm` container smoke also passed on Windows 10 build 26300 / Excel 16.0 build 20430.0 with the same one-query/formula/original/stage/output/process checks. The generated `.xlsm` had no VBA project; VBA-bearing preservation, interrupted recovery, and alternate Office builds remain unverified.


## W144 — 2026-10-22 — marshal general UI worker completion to the GUI thread

- Audit found five remaining worker-thread `QThread.finished` lambdas that called `_set_busy(False)` directly for import, workbook inspection, capability probing, refresh and staged recovery. Added a typed `MainWindow` slot and connected each completion signal to it; history operations already use their QObject UI bridge.
- Failing-first QThread regression confirmed a typed completion callback updates the UI on the main Qt thread. Focused worker/history tests pass; full portable suite **373 passed**; UI suite **66 passed** at Qt scale 1.25 and 1.5. Windows build completed with 373 embedded tests.
- Rebuilt package passed verifier for 229 files / 135,751,591 bytes and CLI help. Packaged keyboard traversal passed (35 controls, expected order, no Excel). Disposable packaged GUI `.xlsx` definition-only import passed on Windows 10 build 26300 / Excel 16.0 build 20430.0 with exact formula, verified backup, read-only reopen, manager exit 0 and owned-process cleanup. Existing Excel PID 36464 remained untouched. Remaining Office/accessibility matrix items and release gates remain **UNVERIFIED**.


## W145 — 2026-10-22 — standard-user packaged GUI acceptance guard

- `tools/run_packaged_gui_import_smoke.py` now refuses an elevated token and prints its elevation status before launching the manager. This prevents administrator runs from being represented as standard-user acceptance.
- The updated harness passed on the current `PCRN\Robert` non-elevated Windows token. Rebuilt packaged GUI import passed on Excel 16.0 build 20430.0 with exact formula, verified backup, read-only reopen, manager exit 0 and smoke-owned Excel cleanup. The app was launched directly without elevation. This covers this user profile and test folder ACL only, not all standard-user/group-policy configurations.
## W146 — 2026-10-23 — COM cleanup ordering, packaged VBA import, and refresh

- Root cause fixed from the packaged Excel run: the worker waited for Excel process exit before releasing its COM apartment. That forced termination could report a successful import while leaving a secondary Excel window. The installer now calls Quit, releases COM references and uninitializes the worker apartment, then waits up to 15 seconds on the captured process handle. It terminates only that captured PID if needed. The UI worker logger now accepts both per-query result objects and cleanup status strings; a regression reproduces the earlier `str.query_name` failure.
- `app/build.ps1` passed on Windows 10 build 26300 / Python 3.11.9 with the complete suite: **379 passed**. Focused lifetime/query/inspector/refresh/worker tests: **82 passed**. Profile, plan, and package contract validation passed.
- Published `dist` packaged GUI command `python -B tools/run_packaged_gui_import_smoke.py --exe dist/PowerQueryManager/PowerQueryManager.exe --vba` passed on Excel 16.0 build 20430.0 under a non-elevated token. The disposable VBA-bearing `.xlsm` retained exact M formula and VBA source; `Workbook_Open` marker did not change; pre-import backup hash matched; read-only reopen passed; manager exited 0; no smoke-owned Excel PID remained. The captured app-owned Excel process exited gracefully (`forced termination=False`). Pre-existing Excel PID 36464 was untouched.
- Packaged CLI refresh smoke passed in background and synchronous modes on Excel 16.0 build 20430.0. Both results were `refreshed`, backups were verified, no external source was contacted, and smoke-owned processes exited.
- Final package verifier: **229 files / 135,757,954 bytes**; packaged CLI `--help` exit 0. Contract validation passed. There is no `.git` metadata, so commit SHA is unavailable.
- `verify_handoff.py --allow-modified-app` remains failed because the build/test run created local `app/.venv`, `.pytest_cache`, and `__pycache__` artifacts in the development workspace (9,366 cache/environment findings); the release `dist` verifier confirms these are not shipped. No handoff artifact hash is claimed as clean.
- Other Office builds, blocked-COM termination behavior, XLM, external connectors, OneDrive race/conflict cases, broad workbook relationship preservation, physical display/screen-reader acceptance and portions of the Windows matrix remain open. **v1.3.1 gate remains UNVERIFIED**; dependent release gates remain UNVERIFIED.


## W147 — 2026-10-23 — packaged history comparison UIA

- Published package history comparison rerun passed through Windows UI Automation on non-elevated Windows 10 build 26300 / Excel 16.0 build 20430.0. The dialog exposed the named changes table, formula diff and close controls after comparing two seeded local snapshots. Its disposable import, exact formula, verified backup, read-only reopen and manager exit also passed. Only pre-existing Excel PID 36464 remained after exact-PID cleanup.
- Documentation/package refresh: `python -B tools/package_release.py --docs-only` passed; `python -B tools/verify_release_package.py dist` passed (**229 files / 135,758,916 bytes**); `python -B tools/validate_contracts.py` passed.
- `python -B tools/verify_handoff.py --allow-modified-app` failed with 9,366 findings because the development tree contains generated `.venv`, `.pytest_cache`, and `__pycache__` contents. No cleanup workaround was attempted. These artifacts are excluded by the passing distribution verifier.
- The full 379-test suite was run during W146's build; no application source changed afterward. W147 is a new live packaged GUI acceptance. Release gates remain **UNVERIFIED**.


## W148 — 2026-10-23 — packaged CLI mapping integration

- Fixed the packaged integration harness to supply explicit mapping/deployment approval phrases. Rerun passed with actual package CLI plus desktop Excel 16.0 build 20430.0 on disposable files: previews appeared, two definitions deployed, mapped and untouched formulas matched read-only Excel inspection, source hashes were unchanged, and no smoke-owned Excel process remained. No M evaluation/external data contact occurred. Pre-existing Excel PID 36464 was untouched.
- The in-process CLI smoke also passed, verifying the same two formulas and source hashes; smoke-owned Excel PIDs 33864 and 3484 exited. Documentation-only package refresh and verifier passed at 229 files / 135,759,125 bytes; contract validation passed. Application executables did not change. The v2.0.0 gate remains **UNVERIFIED** for other Office builds and remaining matrix cases.


## W149 — 2026-10-23 — safe CLI approval cancellation on EOF

- Failing-first regression reproduced an uncaught EOF from the mapping-approval prompt; fix turns EOF into a blank response and lets the existing decline handler return 3 without Excel/source/workbook mutation. The new regression passed; CLI tests: **29 passed**; full Windows `app/build.ps1`: **380 passed** and rebuilt onedir GUI/CLI.
- Rebuilt `dist` passed verification at 229 files / 135,759,512 bytes; CLI `--help` exit 0. Packaged source-path + M-code mapping smoke passed on Excel 16.0 build 20430.0, with exact M readback, source hashes preserved and owned Excel cleanup. Process inventory contained only pre-existing PID 36464, untouched. Release gates remain **UNVERIFIED**.


## W151 — 2026-10-23 — blocked COM termination smoke

- New live smoke passed with a real synchronous `QueryTable.Refresh(False)` COM call blocked on a deliberately delayed HTTP response from a loopback-only server. The 3-second timeout terminated only captured Excel PID 7508; its retained process handle signaled exit; PID inventory returned to `[36464]`, leaving the pre-existing user process untouched. No workbook was opened/saved and no external network was contacted. Excel 16.0 build 20430.0 / Windows 10 build 26300 / Python 3.11.9.
- The smoke was repeated after tightening harness setup/cleanup and passed again with owned PID 10700; Excel process set again returned to `[36464]`. The docs-only package refresh and verifier passed at **229 files / 135,759,820 bytes**; contract validation passed.
- This closes the previously untested blocked-COM termination path for this method/build. It does not certify Power Query refresh behavior, other methods/builds, or the complete lifecycle matrix. Full suite remains **380 passed** from W149; this turn changed only the QA harness and documentation.


## W152 — 2026-10-23 — packaged Power Query refresh false-success finding

- `tools/run_excel_pq_blocked_refresh_timeout_smoke.py` failed on Windows 10 build 26300 / Excel 16.0 build 20430.0: packaged CLI exited 0 and reported `refreshed` plus staged commit in 0.094 seconds, but the disposable loopback `Web.Contents` source received no request.
- Direct COM probe returned `None` promptly; `Refreshing` was false while a delayed local request was pending, `AfterCalculate` did not signal completion, and `RefreshDate` was unsupported for this query connection. Another run did not reach the endpoint. Only pre-existing Excel PID 36464 remained after each harness; it was untouched.
- This is a confirmed false-success risk. Power Query refresh execution and timeout are **UNVERIFIED / RELEASE BLOCKING**; the earlier static-M refresh smoke did not prove query evaluation. The independent W151 legacy QueryTable blocked-COM result is unaffected. No application source or executables changed in W152; portable suite remains **380 passed** from W149.
- Recovered package integrity after the probes: `python -B tools/verify_release_package.py dist` passed at **229 files / 135,759,820 bytes**; `python -B tools/validate_contracts.py` passed. The feature-traceability helper was not found at the attempted path and was not rerun. There is no `.git` metadata, so no commit SHA is available.


## W153 — 2026-10-23 — recovered packaged GUI import acceptance

- The recovered `dist/PowerQueryManager/PowerQueryManager.exe` passed `tools/run_packaged_gui_import_smoke.py` on non-elevated Windows 10 build 26300 / Excel 16.0 build 20430.0. UI Automation found named main manager, inspector, dependency graph, and profile mapping controls; mapping edits were cancelled.
- Disposable `.xlsx` definition-only import passed exact M formula readback, backup hash verification, read-only reopen, manager exit 0, and graceful exit of its captured Excel PID. Excel process inventory returned to `[36464]`; the pre-existing user-owned PID was untouched.
- W153 verifies packaged GUI/import only; at that point the newly discovered W152 refresh false-success still blocked v1.5.0. W154/W155 below supersede that refresh status for the tested connection-only query shape.


## W154/W155 — 2026-10-23 — packaged Power Query refresh correction

- Rebuilt app passed `app/build.ps1`: **384 tests passed** and PyInstaller rebuilt GUI/CLI one-folder executables.
- W154: packaged CLI refresh of a disposable connection-only Power Query contacted a loopback CSV source in both background and synchronous modes (three GETs total). Backup hashes matched each pre-refresh workbook; exact M formula, single original query connection and worksheet count read back; temporary query tables/connections were removed; all owned Excel processes exited.
- W155: packaged synchronous refresh against a deliberately delayed loopback source hit its 3-second timeout, returned exit 5, retained original workbook bytes and a matching backup, removed the stage, and terminated only its captured Excel PID. Process inventory returned to the pre-existing PID 36464.
- This corrects W152 for one Office build and a connection-only query: the refresher now uses a worksheet QueryTable result surface instead of trusting `WorkbookConnection.Refresh`. At W154, Data Model and worksheet-loaded refresh were still open; W157/W158 below add their later scoped acceptance. External providers, other Office builds and the remaining matrix are still open; release gates remain **UNVERIFIED**.


## W157 — 2026-10-23 — packaged worksheet-loaded refresh

- The published CLI refresh smoke passed connection-only and existing worksheet-loaded queries in both background and synchronous modes on Excel 16.0 build 20430.0. The worksheet table retained its `D4` location and query connection and contained the loopback response value `42`; exact M formula, per-operation backup hashes, original worksheet count and connection inventory were verified.
- Nine local GETs were observed across four disposable refreshes. All reports said `refreshed`; temporary scratch objects/stages were absent; process inventory returned to `[36464]`. No external network was used.
- W157 closes the worksheet-load refresh case on this one build. At that point Data Model refresh remained refused; W158 below supersedes that status for the tested selected-table loopback case. Alternate Office builds, external connectors, provider error behavior and the remaining matrix remain open. Release gates remain **UNVERIFIED**.

## W158 — 2026-10-23 — packaged Data Model table refresh

- Added single-table Data Model refresh through Excel's `ModelTable.Refresh()` API and all-selected multi-table refresh through `Workbook.Model.Refresh()` only when every model table maps uniquely to a selected Power Query. Partial multi-table selections fail before backup/source access; the staged file is discarded. Both COM APIs are synchronous, time-bounded against the owned Excel PID, and followed by connection and row-count readback; operation reports preserve that limitation.
- `app/build.ps1`: **390 tests passed**; PyInstaller rebuilt GUI and CLI. Published `dist` CLI smoke passed background-requested and synchronous-requested **Refresh all** on a two-table disposable workbook on Windows 10 build 26300 / Excel 16.0 build 20430.0. Both selected local loopback sources were contacted, both model tables reopened with two rows, formulas/connections/worksheet count stayed correct, backups matched each pre-refresh workbook, stages were removed, and Excel process inventory returned to `[36464]`. A partial selection of one of those two queries was refused with the original workbook hash unchanged and no new source request.
- Microsoft documents `ModelTable.Refresh()` as refreshing its model table's source connections and `Model.Refresh()` as refreshing all sources associated with that workbook model. The full-selected-model route is enabled only after the source mapping proves that all model tables are selected. Other Office builds, external providers, credentials/privacy prompts, complex model relationships, and remaining Windows acceptance remain unverified. Release gates remain **UNVERIFIED**.

## W159 — 2026-10-23 — recovered package documentation and integrity refresh

- `python -B tools/package_release.py --docs-only` regenerated the compact package manual, release summary, and SHA256SUMS without changing the packaged executables. The package verifier passed: **229 files / 135,780,354 bytes**. Contract validation passed, and feature traceability resolved all **98** cited test functions.
- Updated W157 historical wording to make clear that W158 supersedes its then-current Data Model refusal, and updated generated package notes to describe all-selected multi-table refresh versus preflight refusal of partial selection.
- W159 is documentation/integrity verification only; it does not change the 390-test result or Windows acceptance. No Git metadata is present, so commit SHA remains unavailable. Release gates remain **UNVERIFIED**.

## W160 — 2026-10-23 — recovered published GUI import acceptance rerun

- Command: `python -B tools/run_packaged_gui_import_smoke.py --exe dist/PowerQueryManager/PowerQueryManager.exe` on Windows 10 build 26300 / Excel 16.0 build 20430.0 under a non-elevated token.
- PASS: UI Automation found named manager, query grid, filter, replace policy, backup, import, operation log, inspector, dependency graph, and profile mapping controls. Mapping changes were cancelled. Disposable definition-only `.xlsx` import verified exact formula, backup, and read-only reopen; manager exited 0 and its Excel process exited gracefully. The verifier's Excel process was cleaned. Process inventory returned to pre-existing PID 36464, which was untouched.
- This acceptance used the recovered published executable; broader Windows display/accessibility and alternate Office build acceptance remain open. Release gates remain **UNVERIFIED**.

## W161 — 2026-10-23 — history snapshot publication crash recovery

- Added a failing-first subprocess regression that terminates with `os._exit(73)` immediately after atomic snapshot-file publication and before its SQLite registration. Before the fix, reopening the history store left the complete workbook bytes unindexed in `history_snapshots` and exposed no recovery path.
- `HistoryStore` now detects unregistered manager-named UUID workbook snapshots and moves them without deleting bytes to `history.sqlite3.recovered`. The UI reports the location for manual review; recovered files do not become trusted/verified history snapshots automatically. Recent in-progress `.pqm-snapshot-*` copies (under 24 hours), unrelated filenames, indexed files, and reparse points are not moved.
- Verification: the hard-exit regression and adjacent live-temp/unrelated-file preservation test pass; `tests/test_history.py`: **37 passed**. Packaged rebuild and full suite are pending. No Excel interaction occurred for this filesystem-only change. Overall gates remain **UNVERIFIED**.

## W162 — 2026-10-23 — rebuilt package and history crash-recovery verification

- `app/build.ps1`: **392 passed** and PyInstaller rebuilt the GUI/CLI on Windows 10 build 26300 / Python 3.11.9. A subsequent full `pytest -q` run after adding UI-report coverage passed **393 tests**. Root package generation/verifier passed at **229 files / 135,782,556 bytes**; contract validation passed and all **101** feature-traceability tests resolve.
- The rebuilt published GUI passed `python -B tools/run_packaged_gui_import_smoke.py --exe dist/PowerQueryManager/PowerQueryManager.exe` on Excel 16.0 build 20430.0 under a non-elevated token. Definition-only import formula, backup, read-only reopen and exit checks passed; process inventory returned to pre-existing Excel PID 36464.
- History power-loss recovery was exercised in the portable subprocess test; Excel was not involved in that scenario. The packaged GUI smoke does not simulate this history crash path inside the executable. Other power-loss boundaries and all broader release gates remain **UNVERIFIED**.

## W163 — 2026-10-23 — flush snapshot contents before publication

- Added a regression that first failed because `create_snapshot()` did not call `fsync` before the final rename. The implementation now flushes the copied temporary snapshot file before verifying/publishing its UUID filename.
- `app/build.ps1`: **394 passed** and PyInstaller rebuilt GUI/CLI. Package generation and verification passed at **229 files / 135,782,806 bytes**; contract validation passed; all **102** cited test functions resolve; packaged CLI `--help` exited 0.
- Rebuilt GUI smoke passed on Excel 16.0 build 20430.0 under a non-elevated token: UI Automation, disposable definition-only import, exact formula, backup, read-only reopen, exit 0 and owned-process cleanup passed. Process inventory returned to pre-existing PID 36464.
- The file content flush is verified; atomic rename directory-entry durability during physical power loss has not been verified. Other crash phases and release gates remain **UNVERIFIED**.

## W164 — 2026-10-23 — stale partial snapshot recovery coverage

- Added a 48-hour-old disposable `.pqm-snapshot-*` partial-copy fixture and confirmed startup preserves it in the recovered folder without indexing it as a valid snapshot. A companion test confirms the recent-copy threshold leaves a current temp untouched.
- Both focused tests pass; the full suite now passes **395 tests**. This test-only update does not change the already rebuilt executable or the verified 229-file package. Physical power-loss and directory-entry durability remain unverified.

## W165 — 2026-10-23 — diagnostic log reparse-point refusal

- `record_failure()` now checks the user-local diagnostics directory, active log, and all rotation targets for symbolic links/Windows reparse points before directory creation, rotation, and append. If redirected or unwritable, logging remains best-effort and returns `None` without changing the operation outcome.
- Added regressions for a redirected log directory and log file; both verify the log is not created/rotated/changed. `tests/test_redaction.py`: **9 passed**. Package rebuild and full suite are pending. Overall v1.3.1 gate remains **UNVERIFIED**.

## W166 — 2026-10-23 — rebuilt and verified diagnostic-log guard

- `app/build.ps1`: **397 passed**, PyInstaller rebuilt GUI and CLI. Published package verification passed: **229 files / 135,783,793 bytes**; contract validation passed; all **105** traceability references resolve; packaged CLI `--help` exited 0.
- Rebuilt GUI import smoke passed on Windows 10 build 26300 / Excel 16.0 build 20430.0 as a non-elevated user, with exact formula, backup, read-only reopen and owned-process cleanup. Only pre-existing Excel PID 36464 remained.
- The new reparse tests use an injected reparse-point detector; a physical redirected log-folder acceptance was not run. This change does not close broader ACL, OneDrive or kernel race cases. v1.3.1 remains **UNVERIFIED**.

## W167 — 2026-10-23 — actual Windows directory symlink rejection

- A disposable Windows directory symlink under the test temp directory redirected the user-local log root to a separate temp target. `record_failure()` returned `None`; no `logs/errors.jsonl` was created at the target. Temporary files were cleaned by the test context. Existing Excel process PID 36464 was not touched.
- Added `test_error_log_refuses_actual_directory_symlink`; `tests/test_redaction.py`: **10 passed** and the full suite passed **398 tests**. The test skips only when the host cannot create a directory symlink. This tests one actual NTFS reparse path; file-link, ACL, OneDrive and concurrent path-swap cases remain open. Package binaries were unchanged from W166.

## W168 — 2026-10-23 — actual Windows log-file symlink rejection

- A disposable `errors.jsonl` symlink pointed to an external temp file. `record_failure()` refused the path; the target text remained byte-for-byte unchanged and no rotation backup appeared.
- Added `test_error_log_refuses_actual_file_symlink_before_append_or_rotation`; the full suite passes **399 tests**. Package binaries remain unchanged from W166. ACL/OneDrive behavior, other reparse types and concurrent path swaps remain open.


## W156 — 2026-10-23 — rebuilt published packaged GUI acceptance

- After publishing the rebuilt GUI/CLI package, `tools/run_packaged_gui_import_smoke.py --exe dist/PowerQueryManager/PowerQueryManager.exe` passed under a non-elevated token on Windows 10 build 26300 / Excel 16.0 build 20430.0. Manager, inspector, dependency graph and profile mapping UIA checks passed; disposable definition-only import preserved exact M, backup hash and read-only reopen; manager exited 0 and owned Excel processes were cleaned.
- Process inventory returned to `[36464]`; pre-existing user-owned Excel PID 36464 was untouched. This does not complete screen-reader, display, alternate Office or remaining release matrix items.


## W169 — 2026-10-23 — atomic snapshot index and audit event

- A failing-first subprocess regression terminates immediately after snapshot metadata insertion. Before the fix, the child returned normally because snapshot index and audit event used separate connections. `create_snapshot()` now inserts both through one SQLite transaction; `record()` shares the same transaction helper. On hard exit, SQLite rolls back the metadata/event pair and startup discovers the published workbook as unindexed for manual recovery.
- `app/build.ps1` passed **400 tests** and rebuilt GUI/CLI with PyInstaller on Windows 10 build 26300 / Python 3.11.9. The refreshed compact package passed verification at **229 files / 135,784,108 bytes**; contract validation passed; all **108** traceability references resolved; packaged CLI `--help` exited 0. The rebuilt GUI passed packaged UI Automation checks and disposable `.xlsx` definition-only import on Excel 16.0 build 20430.0: exact formula, verified backup, read-only reopen, manager exit 0, and smoke-owned Excel cleanup. Only pre-existing Excel PID 36464 remained. Snapshot crash simulation itself is portable SQLite/filesystem coverage and did not interact with Excel. Excel release gates remain **UNVERIFIED**. A post-build default `verify_handoff.py` run is not a release-package verifier: it reported the expected modified-baseline hashes and local `.venv`/test cache paths created by the Windows build. The dedicated release-package verifier passed, and those development files are absent from `dist`.


## W170 — 2026-10-23 — no-overwrite snapshot publication

- A failing-first UUID collision regression showed `os.replace` overwrote bytes already present at the generated snapshot filename. Snapshot publication now uses an exclusive hard link after flushing and hash verification; it fails closed if the destination exists and cleans up only a destination published by this call. The collision test preserves the sentinel bytes, leaves the snapshot index empty, and removes the temporary copy.
- Crash-boundary coverage now exercises hard-link publication: a forced exit in the link-to-temp-unlink interval leaves both names with the same workbook bytes; startup recovers the UUID-named copy and leaves the recent temporary alias intact. Focused history suite: **41 passed**. `app/build.ps1`: **401 passed** and rebuilt GUI/CLI with PyInstaller. The compact package verifier passed at **229 files / 135,784,395 bytes**; contract validation passed; all **109** traceability references resolved; packaged CLI `--help` exited 0. Rebuilt packaged GUI smoke passed UI Automation and disposable definition-only import on Excel 16.0 build 20430.0 with exact formula, verified backup, read-only reopen, manager exit 0 and owned-process cleanup; only pre-existing Excel PID 36464 remained. Excel was not used to test the filesystem/database-only collision behavior.


## W171 — 2026-10-23 — repeatable PyInstaller onedir payload

- Two consecutive clean `app/build.ps1` builds each passed **401 tests**. Before the build fix, clean outputs differed in three files (GUI/CLI PE timestamps and `base_library.zip` module ordering). `build.ps1` now pins `PYTHONHASHSEED=0` and `SOURCE_DATE_EPOCH=1704067200` only around PyInstaller, restoring prior environment values afterward.
- After the fix, the two 227-file onedir payloads had identical relative paths and SHA-256 for every file (0 missing, 0 extra, 0 changed) on Windows 10 build 26300 / Python 3.11.9 / PyInstaller 6.13.0. This demonstrates same-host/toolchain repeatability only, not cross-host reproducibility. The deterministic build was published into root `dist`; package verification passed at **229 files / 135,784,395 bytes**. Contract validation passed; all **109** feature-traceability references resolved; packaged CLI `--help` exited 0. The rebuilt packaged GUI smoke passed on Excel 16.0 build 20430.0 under a non-elevated token: UI Automation, disposable definition-only import, exact formula, verified backup, read-only reopen, manager exit 0 and owned-process cleanup. Only pre-existing Excel PID 36464 remained.


## W172 — 2026-10-23 — serialize snapshot publication and recovery

- A barrier-based concurrent-startup regression failed against the current behavior: a second `HistoryStore` instance completed orphan recovery while the creator was paused immediately after hard-link publication. The creator could then commit an index pointing at a moved file.
- Snapshot creation now starts an immediate SQLite write transaction before publishing the final hard link and holds it through snapshot metadata plus history-event commit. Startup orphan recovery acquires the same write lock before reading indexed paths and moving candidates. The regression passes; `tests/test_history.py`: **42 passed**. `app/build.ps1`: **402 passed** and rebuilt GUI/CLI. Compact package verification passed at **229 files / 135,784,626 bytes**; contract validation passed; all **110** traceability references resolved; packaged CLI `--help` exited 0. Rebuilt packaged GUI smoke passed UI Automation and disposable definition-only import on Excel 16.0 build 20430.0 with exact formula, verified backup, read-only reopen, manager exit 0 and owned-process cleanup; only pre-existing PID 36464 remained. No Excel interaction occurred for the concurrency regression itself. A second clean build after this code change again passed **402 tests** and matched the prior 227-file onedir payload exactly (0 missing, extra or changed hashes), confirming W171 reproducibility controls still work with W172.


## W173 — 2026-10-23 — reject redirected history database paths

- Two failing-first Windows regressions showed `HistoryStore` followed a file symlink to an existing SQLite database and created a database through a symlinked parent. The constructor now checks the complete lexical history path before parent creation/open, then rechecks after creating a normal missing parent. Both actual NTFS symlink tests pass; the existing target database hash is unchanged and the linked directory receives no database file. Focused history suite: **44 passed**.
- The current `%LOCALAPPDATA%\PowerQueryManager\history.sqlite3` path has no reparse-point component. `app/build.ps1`: **404 passed** and rebuilt GUI/CLI. Compact package verification passed at **229 files / 135,784,777 bytes**; contract validation passed; all **112** traceability references resolved; packaged CLI `--help` exited 0. The rebuilt packaged GUI smoke passed UI Automation and disposable definition-only import on Excel 16.0 build 20430.0 with exact formula, verified backup, read-only reopen, manager exit 0 and owned-process cleanup; only pre-existing PID 36464 remained. No workbook or Excel process was used in the symlink regression itself.


## W174 — 2026-10-23 — cap compressed query-package size before parsing

- A failing-first test showed `extract_query_zip()` had no upper bound on the compressed archive itself, despite the entry and expanded-size limits. It now rejects archives larger than 32 MiB before opening `ZipFile` or creating a staging directory; callers may set a smaller positive byte limit. The new oversized/trailing-data ZIP regression passes; `tests/test_manager.py`: **35 passed**.
- `app/build.ps1`: **405 passed** and rebuilt GUI/CLI. Compact package verification passed at **229 files / 135,785,193 bytes**; contract validation passed; all **113** traceability references resolved; packaged CLI `--help` exited 0. Packaged GUI smoke passed UI Automation and disposable definition-only import on Excel 16.0 build 20430.0 with exact formula, verified backup, read-only reopen, manager exit 0 and owned-process cleanup; only pre-existing PID 36464 remained. Existing path traversal, symlink, reserved-name, duplicate, expanded-size and compression-ratio protections remain covered. No Excel actions occurred in the archive-limit regression itself.

## W175 — 2026-10-23 — bind ZIP size check to parsed file handle

- A regression models a path replacement at the parser boundary and requires `ZipFile` to parse the already-open file that was size-checked. The extractor now opens first, checks the opened file descriptor with `fstat`, rejects non-regular/empty/oversized inputs, and passes that same stream to the parser. `tests/test_manager.py` and `tests/test_deployment_bundle.py`: **44 passed**; `app/build.ps1`: **406 passed** and rebuilt GUI/CLI.
- Windows does not allow the test's original path-replacement method while the file is open without delete sharing; the regression was adjusted to simulate the parser boundary, not to claim an OS-level rename race. Compact package verification passed: **229 files / 135,786,012 bytes**; contract validation passed; all **114** traceability references resolved; packaged CLI `--help` exited 0. Packaged GUI import smoke passed UI Automation and disposable definition-only import on Windows 10 build 26300 / Excel 16.0 build 20430.0 under a non-elevated token. The exact M formula, verified backup, read-only reopen and manager exit 0 were confirmed; the harness cleaned its own automation processes. Only pre-existing Excel PID 36464 remained. Screen-reader, physical-display, alternate Office builds and the complete Excel acceptance matrix remain unverified; v1.4.0 gate remains **UNVERIFIED**.

## W176 — 2026-10-23 — expanded current-package Excel acceptance

- Using the current compact package on Windows 10 build 26300 / Excel 16.0 build 20430.0 under a non-elevated token, packaged GUI `--skip-existing` preserved the disposable workbook hash and timestamp, created no backup/stage, retained the existing formula, reported no change, and exited 0. `--replace-existing` read back the new exact formula, verified the original-hash backup, preserved an unrelated worksheet marker, and exited 0. `--mixed-import` imported three numbered `.m`/`.pq` files in expected review and Excel collection order with exact formula readback and backup verification. All smoke-owned Excel processes exited; only pre-existing Excel PID 36464 remained.
- These are current-package results on one Office build. Other Office builds, screen-reader testing, physical OS scaling, and the complete Windows matrix remain unverified; v1.3.1/v1.4.0 release gates remain **UNVERIFIED**. No code change was made in this acceptance-only increment.

## W177 — 2026-10-23 — current-package load-mode GUI acceptance

- On Windows 10 build 26300 / Excel 16.0 build 20430.0 under a non-elevated token, current packaged GUI connection-only mode read back the exact formula and explicit workbook connection with no worksheet or model table and no refresh. New-worksheet capability probe/import verified exact M, values 17/29, worksheet load state and backup hash. Existing-worksheet capability probe/import verified exact M and Output!C4 table anchor/values 17/29, unrelated marker and backup hash. Data Model probe/import verified exact M, model membership and row count 2, plus backup hash. Each manager exited 0; every harness-owned Excel process was cleaned and pre-existing PID 36464 remained.
- These are static disposable-query probes on one Excel build, not provider or cross-version certification. Other Office builds, screen-reader/physical display and remaining Windows matrix cases remain unverified; v1.5.0 gate remains **UNVERIFIED**.

## W178 — 2026-10-23 — packaged CLI static-M refresh acceptance

- The current packaged CLI refreshed a disposable inline-static-M query in background and synchronous modes on Excel 16.0 build 20430.0. Both reports returned `refreshed`; each staged workbook had a verified backup hash; the harness confirmed no external source was contacted and all owned Excel processes exited. Existing Excel PID 36464 was the only process left.
- This validates the current packaged refresh path for inline static M on one Office build. External providers, blocked Power Query refreshes, other Office builds and remaining matrix cases are not covered; v1.5.0 remains **UNVERIFIED**.

## W179 — 2026-10-23 — package-summary generator regression checks

- `app/.venv/Scripts/python.exe -m pytest tests/test_release_package_staging.py tests/test_release_package_verifier.py -q`: **8 passed** after the current release-summary generator updates. Final compact package remained at **229 files / 135,787,309 bytes**; package verifier, contract validator and all **114** traceability references passed. This check covers docs-only publication/staging behavior and package integrity, not Excel compatibility.

## W180 — 2026-10-23 — broad workbook-object preservation rerun

- `python -B tools/run_excel_preserve_smoke.py` passed on Windows 10 build 26300 / Excel 16.0 build 20430.0. Query replacement preserved the ordinary table's identity/range/data/style/filter, defined names, cell formulas, `ForceFullCalculation`, chart and shape objects, PivotTable values 7/54/61, and report formulas/names linked to a query-backed table. Distinct connection-only and worksheet-loaded states and workbook privacy metadata (`FirewallEnabled=True`, `WorkbookGroupType=None`) were read back. The harness reports all owned Excel processes exited; the pre-existing Excel PID inventory was checked separately.
- This is one workbook fixture and one Office build; other object families and Office versions remain unverified. v1.5.0 remains **UNVERIFIED**.

## W181 — 2026-10-23 — verification workbook close failure blocks commit

- Failing-first regression showed the importer swallowed an error closing Excel's read-only staged-workbook verification view and continued to `commit_stage()`. The importer now raises an actionable sanitized failure when that close is not confirmed. The stage is discarded, the original workbook bytes stay unchanged, and the affected query is reported `ROLLED_BACK`. The regression passed after the fix; `tests/test_excel_queries.py`: **42 passed**. `app/build.ps1`: **407 passed** and rebuilt GUI/CLI. Compact package verification passed at **229 files / 135,788,307 bytes**; contracts, all **115** traceability references, and packaged CLI help passed. Packaged GUI definition-only import passed on Excel 16.0 build 20430.0 with exact M, backup, read-only reopen, manager exit 0 and smoke-owned process cleanup; only pre-existing PID 36464 remained. Real Excel close-failure injection remains unverified.

## W182 — 2026-10-23 — rebuilt-package replacement reacceptance

- After the 409-test build, the rebuilt compact package passed `tools/run_packaged_gui_import_smoke.py --replace-existing` on Windows 10 build 26300 / Excel 16.0 build 20430.0 under a non-elevated token. The new formula read back exactly; original-hash backup and unrelated worksheet marker were verified; the GUI reported replacement, manager exited 0, and the harness cleaned owned Excel processes. Only pre-existing PID 36464 remained.
- This exercises the normal successful replace path, not real Office verification-close failure injection or other Office builds; v1.3.1 remains **UNVERIFIED**.

## W183 — 2026-10-23 — tolerate incomplete UI Automation nodes in smoke harness

- A packaged GUI smoke reached successful import completion, then the optional operation-log print enumerated all descendants and crashed when pywinauto encountered a UIA node with `control_type=None`. Moved this diagnostic to a best-effort helper; unknown node types now make the log unavailable rather than fail the acceptance run. Test-first regressions cover the `KeyError(None)` path and successful named-edit extraction; both pass. Full `app/build.ps1`: **409 passed**, GUI/CLI rebuilt. The refreshed compact package passed at **229 files / 135,788,815 bytes**; contracts, **117** traceability references and packaged CLI help passed. The initial VBA smoke was not counted as accepted because post-import verification did not run; the refreshed W184 rerun passed.

## W185 — 2026-10-23 — fail closed on unknown worksheet inventory

- A failing-first fake-COM regression showed new-sheet creation swallowed arbitrary errors from `Worksheets.Item(name)` and proceeded as if the destination were absent. Import now enumerates indexed worksheets and compares names; enumeration/count/name failures produce an actionable inventory error, with no worksheet creation. Existing-sheet selection uses the same helper and never assumes an unverified destination. The regression plus `tests/test_excel_queries.py`: **44 passed**. `app/build.ps1`: **410 passed** and rebuilt GUI/CLI. Compact package passed at **229 files / 135,790,075 bytes**; contracts, all **118** traceability references and packaged CLI help passed. Packaged new-worksheet and existing-worksheet imports both passed on Excel 16.0 build 20430.0: capability probes, exact M, values 17/29, Output!C4 anchor, unrelated marker, backup hashes, read-only reopen, manager exit 0 and owned Excel cleanup. Real Excel injection of worksheet collection failures and other Office builds remain unverified.

## W186 — 2026-10-23 — current-package keyboard traversal

- `tools/run_packaged_keyboard_smoke.py --exe dist/PowerQueryManager/PowerQueryManager.exe` passed against the current compact package. Tab traversal visited **35 focusable controls**, reached workbook path, Add files, query grid, existing-query policy, backup and import controls in order, and completed a full cycle without leaving the manager. No workbook or Excel instance was opened. Pywinauto emitted its STA-mode warning and reset to STA; the harness exited 0.

## W187 — 2026-10-23 — open-workbook preflight detection without lock file

- A focused Windows subprocess regression (`test_windows_restart_manager_detects_an_open_file_without_excel_owner_lock`) passed: Restart Manager found another process holding the disposable `.xlsx` path without a `~$` file, and write preflight refused it. The packaged environment's real Excel 16.0 build 20430.0 smoke opened a valid disposable `.xlsx` read-only with no owner-lock file; the exact-path ROT check returned `Microsoft Excel (open workbook)`, preflight refused, workbook hash stayed unchanged, and the owned Excel process exited. Only pre-existing PID 36464 remained after both checks.
- This completes the two W22 active-use detection subcases on one Windows/Office host. It does not exercise full import/refresh/restore against the lock or OneDrive synchronization conflicts; broader matrix remains open.

## W188 — 2026-10-23 — mutation services refuse read-only Excel open without owner-lock file

- `python -B tools/run_excel_open_without_lock_smoke.py` passed on Windows 10 build 26300 / Excel 16.0 build 20430.0 / Python 3.11.9 x64. A valid disposable `.xlsx` was held read-only open in Excel without a `~$` owner-lock file. Import, refresh, and history overwrite-restore each refused the active target before creating a stage, backup, or Excel automation instance. The workbook SHA-256 stayed unchanged. The harness closed only its owned Excel process; pre-existing PID 36464 remained.
- This verifies the current services' preflight on one Office build. It does not cover OneDrive conflicts, alternate Office builds, or every filesystem race. v1.3.1 and later release gates remain **UNVERIFIED**.

## W190 — 2026-10-23 — three sequential packaged GUI imports and process cleanup

- Ran `python -B tools/run_packaged_gui_import_smoke.py --exe dist/PowerQueryManager/PowerQueryManager.exe` three times sequentially. Each run used a separate disposable `.xlsx`, passed UIA checks for the main manager and key dialogs, imported a definition-only query, verified exact M and the pre-import backup hash by read-only reopen, and exited 0.
- Each run cleaned its owned Excel processes and returned inventory to the pre-existing Excel PID 36464. No unrelated Excel process was closed. Runs 2 and 3 completed import successfully, but Windows UI Automation did not expose the optional operation-log text; the log-text assertion is therefore not established for all three runs.
- Host: Windows 10 build 26300 / Excel 16.0 build 20430.0 / non-elevated. Other Office builds, physical display and screen-reader acceptance remain unverified; release gates remain **UNVERIFIED**.

## W192 — 2026-10-23 — keyboard traversal on the W191 package

- `python -B tools/run_packaged_keyboard_smoke.py --exe dist/PowerQueryManager/PowerQueryManager.exe` passed. UI Automation traversed 35 controls without focus leaving the manager; workbook path, Add files, query table, replace policy, backup and import action appeared in the expected order. No workbook or Excel process was opened.
- The host's pywinauto emitted its known COM STA reset warning; the smoke exited 0. This checks control traversal, not screen-reader use or physical display scaling.

## W193 — 2026-10-23 — final 413-test build and published-package smoke

- `app/build.ps1` rebuilt the one-folder GUI/CLI from the current source tree. The full current portable suite is **413 passed**; two package staging/verifier tests also pass. Package generation published only runtime files and concise docs; package verifier: **229 files / 135,793,203 bytes**. Contracts and **118** feature-traceability references passed; packaged CLI `--help` exited 0.
- The exact published `dist/PowerQueryManager/PowerQueryManager.exe` passed strict operation-log UIA lookup on Windows 10 build 26300 / Excel 16.0 build 20430.0, non-elevated. The final “Excel automation finished” entry was present. Disposable definition-only import verified exact M, original-hash backup, read-only reopen and manager exit 0. Only harness-owned Excel exited; pre-existing PID 36464 remained untouched.
- Same-host package verification is complete for this build. Screen-reader, physical display, alternate Office builds and remaining Excel matrix cases remain unverified; release gates remain **UNVERIFIED**.

## W194 — 2026-10-23 — fail closed when stage absence cannot be verified

- A failing-first regression injected ten stage-unlink failures and made path existence appear false while `stat()` returned access denied. Old cleanup reported success and omitted the retained-stage warning. `cleanup_stage()` now reports success only for a confirmed `FileNotFoundError`; a verified-present or unreadable path returns false so the caller logs the stage filename.
- Focused transaction/import/refresh cleanup tests: **15 passed**. Full suite after the fix: **414 passed**; package staging/verifier tests: **8 passed**. `app/build.ps1` rebuilt GUI/CLI; compact package verifier passed at **229 files / 135,793,988 bytes**; contracts, 118 traceability references and CLI help passed.
- Exact published package GUI import smoke passed on Windows 10 build 26300 / Excel 16.0 build 20430.0, non-elevated: strict final operation-log UIA, exact M, verified original-hash backup, read-only reopen, manager exit 0 and owned-process cleanup. Pre-existing Excel PID 36464 remained untouched. This normal-path smoke does not inject a real Windows sharing violation; that case is covered through the filesystem regression.
- Alternate Office builds, OneDrive conflicts, screen-reader and physical display acceptance remain open; v1.3.1 and later release gates remain **UNVERIFIED**.

## W195 — 2026-10-23 — real Windows delete-sharing denial during stage cleanup

- A Windows-only test opened a disposable staged workbook with `CreateFileW` using `FILE_SHARE_READ` but omitting `FILE_SHARE_DELETE`. Cleanup encountered the actual Windows sharing denial, returned false, and reported the retained stage filename; the staged bytes remained exact. After the test closed its own handle, a second cleanup confirmed deletion.
- Focused sharing test: **1 passed**; full app suite: **415 passed**. This uses an owned test handle only and did not open Excel or touch a workbook. The production source was already included in W194's verified build; this increment adds direct OS-level regression coverage.
- This proves Windows file-sharing behavior, not Excel-originated lock behavior. Real Excel failure injection, alternate Office builds, OneDrive, screen-reader and physical display cases remain open; release gates remain **UNVERIFIED**.

## W196 — 2026-10-23 — contrast-safe M/diff editor themes

- A failing-first contrast regression measured the existing light-theme comment color at **4.02:1** on white and dark-theme M token colors at **2.60–4.32:1** on the dark editor background. Dark-theme diffs and line-number gutters also retained light palettes.
- M syntax, diff, line-number gutter and current-line highlight now follow the inherited manager theme. Unit UI checks enforce at least **4.5:1** for M tokens, diff text/background pairs and gutter labels in light/dark modes; a nested Query Inspector regression confirms editors update when the parent theme toggles. The selected palette pairs measure 5.02:1 or higher for light M syntax and 6.85:1 or higher for dark diff text; the contrast assertions pass for all tested tokens and roles.
- `app/build.ps1` rebuilt the one-folder GUI/CLI. Full app suite: **418 passed**; UI suite at `QT_SCALE_FACTOR=1.25`: **70 passed**; at `1.5`: **70 passed**. Compact package verifier: **229 files / 135,795,382 bytes** before the documentation refresh; contracts, 118 traceability references and CLI help passed. Published GUI import smoke passed on Windows 10 build 26300 / Excel 16.0 build 20430.0 with exact M, backup, read-only reopen, strict operation-log UIA and owned-process cleanup.
- This is automated contrast and scaled-layout coverage, not screen-reader or physical OS display acceptance. Other Office builds and remaining Windows matrix items remain **UNVERIFIED**.

## W201 — read-only install-folder acceptance

- The exact published GUI package ran from a disposable current-user read/execute-only folder. A write probe was denied; packaged UIA, definition-only import, exact formula readback, backup, read-only reopen, manager exit 0 and smoke-owned Excel cleanup passed on Windows 10 build 26300 / Excel 16.0 build 20430.0. Existing Excel PID 36464 remained untouched; the temp tree was removed.
- Repeatable harness: `powershell -NoProfile -ExecutionPolicy Bypass -File tools/run_packaged_readonly_install_smoke.ps1`. Other ACL/policy variants and Office builds remain unverified.

## W202 — final package and gate accounting

- `dist` docs/summary/checksums refreshed and verified: 229 payload files, 135,801,796 bytes. All 227 packaged runtime files match `app/dist/PowerQueryManager`; no missing, extra, or hash-different files. Package staging/verifier tests: 8 passed; contracts and 118 traceability references pass; packaged CLI help exits 0.
- `verify_handoff.py --allow-modified-app` continues to flag workspace-local venv/cache/bytecode/build files during its recursive hygiene scan, while milestone references and baseline accounting pass. These remain outside `dist`; no development environment files were deleted.
- Remaining release acceptance is explicitly **UNVERIFIED** in release reports; no production release gate was forced green.

## W203 — history restore no longer recreates a missing workbook path

- Two failing-first Windows boundary tests cover target deletion immediately before snapshot restore publication and immediately before rollback after verification failure. Both use the shared `ReplaceFileW` helper; no target path is recreated and the verified timestamped backup remains available.
- `app/build.ps1` passed **426 tests** and rebuilt the GUI/CLI. The packaged GUI history-restore smoke on Excel 16.0 build 20430.0 restored a disposable workbook to its snapshot hash and verified the exact pre-restore backup. The follow-on definition-only import passed formula readback, backup, read-only reopen and owned-process cleanup. Existing Excel PID 36464 remained untouched.
- Rebuilt-package read/execute-only installation smoke also passed: write probe denied, UI Automation passed, import/backup/reopen passed, manager exited 0, and the temporary ACL tree was removed.

## W204 — published-package documentation and checksum refresh

- Refreshed the manual, release summary and SHA-256 manifest after W203. The compact package passed verification; release staging/verifier tests passed **8**. Full 426-test build, contract validation, 118 feature traceability references, packaged CLI help and the GUI/Excel W203/W201 acceptance results are recorded above.
- This host/build does not cover generated native partial `ReplaceFileW` failures, cloud-sync races, other Office versions, screen-reader use, or physical display acceptance. Release gates remain **UNVERIFIED**.

## W205 — final published package record

- `dist` verifies at 229 payload files / 135,802,657 bytes. Manifest SHA-256: `FFA83A238E4EB8EE66E293E4D1B7BD08E39C9DE6E5C13C22193988E11A163830`. All 227 runtime files match the rebuilt app payload.
- Final package staging/verifier tests: **8 passed**; contract validation, 118 traceability references and packaged CLI help pass. W203 history-restore and W201 read-only-folder packaged GUI smokes pass on Excel 16.0 build 20430.0. The docs-only refresh did not change executables.
- Release gates remain **UNVERIFIED** pending broader Office, sync/security, accessibility and workbook matrix acceptance.

## W206 — packaged scaled-layout acceptance

- GUI acceptance passed at Qt scale factors **1.25** and **1.5**, each with UIA dialog checks, exact import/readback, verified backup, read-only reopen and process cleanup. PID 36464 was untouched.
- This does not certify physical monitor scaling or screen-reader use. Those gates remain open.

## W207 — repeated clean build

- Second clean Windows build passed **426 tests** and matched all **227** published runtime file hashes (0 missing/extra/different). Toolchain: Windows 10 build 26300, Python 3.11.9, PyInstaller 6.13.0.
- This is same-host repeatability, not cross-host reproducibility.

## W208 — final compact package manifest

- The current `dist` package contains 229 payload files / 135,802,851 bytes. Manifest SHA-256: `2DC92E739E1528C69A81362FC12FD1D7A0C0B3DDD08E1A9C0DD81CD0404893A3`. Package verifier passed; all 227 runtime files match both clean W203 builds.
- Packaging tests (8), contracts, traceability (118 refs), and CLI help passed. History restore, read-only install-folder, and scaled GUI import acceptance passed on the available Excel build. Remaining matrix gates are unverified.

## W209 — rebuilt package load-mode reacceptance

- Current published package passed separate disposable workbook smokes for new worksheet, existing worksheet (`Output!C4`), connection-only, and Data Model on Excel 16.0 build 20430.0. It verified values 17/29 for worksheets, no sheet/model object for connection-only, and Data Model membership with two rows.
- Exact M, backup, read-only reopen, manager exit and smoke-owned process cleanup passed in each run; existing PID 36464 remained untouched. These use static inline M on one Office build; external sources, other builds and the broader matrix remain unverified.

## W210 — current-package compatibility checks

- Replace, preserve-existing connection-only, all-skip, and mixed numbered `.m`/`.pq` profile import passed on the current package and Excel 16.0 build 20430.0. All checks used disposable workbooks, exact readback/hashes, and smoke-owned process cleanup; PID 36464 was untouched.
- No-change skip produced no backup/stage. Reverse-serialized profile rows imported in numeric-prefix order. Other Office versions and wider loaded-query replacement cases remain unverified.

### W211 — current packaged refresh behavior

- Current CLI package passed connection-only and worksheet-loaded refresh in background and synchronous modes against a local loopback CSV service, plus timeout (exit 5, original hash/backup preserved, stage removed) and selected Data Model refresh. Both model tables reopened with two rows. Pre-existing PID 36464 remained untouched.
- This verifies one Office build and localhost source execution; external providers and other Office versions remain unverified.

## W212 — final compact package record after W209-W211

- `dist` verifies at 229 payload files / 135,803,776 bytes; manifest SHA-256 `EF6898E685BE15DBC228A717C6F6965C41964A0F71C1E21F19B5D8BEED3B16DF`. Docs-only refresh left all runtime binaries unchanged.
- The current package has passed load-mode, replace/skip/order, history restore, scaled UI, loopback refresh/timeout and Data Model refresh acceptance on Excel 16.0 build 20430.0. Other Office builds, external providers and remaining policy/accessibility/workbook matrix cases remain unverified.

## W213 — macro-preservation acceptance

- The rebuilt packaged GUI preserved VBA bytes in a disposable `.xlsm` and did not run its harmless `Workbook_Open` marker during definition-only import. Exact formula, backup, read-only reopen, exit and owned-process cleanup passed on Excel 16.0 build 20430.0.
- XLM and other VBA/Office variants remain unverified.

## W214 — packaged multi-workbook deployment

- Packaged `.pqmdeploy` applied to two disposable workbooks with exact formula/load readback and verified original backups. The repeat batch was a no-op on both; disabled missing source was skipped. All three smoke-owned Excel PIDs exited; pre-existing PID 36464 remained.
- An earlier attempt stopped on a process-set change naming PID 2200; its ownership and cause could not be established. No ownership claim is made for that attempt. The accepted retry started with stable baseline `[36464]`.

## W215 — mixed-target batch isolation

- The engine applied the writable disposable workbook and safely blocked the read-only workbook before backup. Exact backup and unchanged read-only target hash verified; both smoke-owned Excel processes exited; PID 36464 remained.
- Mid-commit failure recovery, OneDrive/cloud-sync and other Office builds remain open.

## W216/W217 — packaged v2 mappings and query packages

- Packaged CLI source-path and M-literal mappings passed with exact preview, typed approvals, unchanged source bytes and exact formula readback. Query-package export/import passed with three live workbook definitions, Unicode/punctuation name, unchanged export source hash, verified import backup and exact formula readback.
- No external source evaluated; Excel 16.0 build 20430.0. Other Office builds and broader M/provider behavior remain unverified.

## W218 — final compact package after v2 reacceptance

- Package verifier passed: 229 payload files / 135,805,053 bytes. Manifest SHA-256 `547F225AA0BAD2D961690CC16905FE678B1F77E9A1B17A134C6E9B00DFB88460`. Staging/verifier tests 8 passed; schemas, 118 traceability refs and CLI help passed.
- Current package reacceptance now includes v2 deployment bundle apply/repeat, mixed read-only batch isolation, path/M mapping and query-package export/import. All were disposable, used one Office build, and preserved PID 36464.
- Remaining v2/release matrix work stays unverified; no release gate is declared complete.

## W219 — packaged GUI declared-parameter editor

- The packaged inspector changed only the supported declared Text literal and saved a separate `.m` copy. Its exact bytes matched the intended edit; the original source hash stayed unchanged.
- Import of that copy into a disposable workbook passed formula readback, timestamped backup verification, read-only reopen with unrelated marker preserved, manager exit 0, and cleanup of the owned Excel process on Windows 10 build 26300 / Excel 16.0 build 20430.0. The run was non-elevated; pre-existing PID 36464 remained untouched.
- Other Office builds and screen-reader acceptance remain unverified. No release gate is declared complete.

## W220 — current package verification after W219

- Final compact `dist` verifies at **229 files / 135,805,910 bytes**; `SHA256SUMS` SHA-256 is `CAE221086423FAB1B05E82F828E7A5F8CD832AC53AAA97B20B868DAB57DAB294`. All 227 app runtime files match the published package (0 missing, differing, or extra files).
- Package staging/verifier tests: **8 passed**. Parameter/core inspector tests: **13 passed**. Profile, plan, and query-package schemas validate; all **118** feature traceability references resolve; packaged CLI `--help` exits 0. Package verifier passes.
- `verify_handoff.py --allow-modified-app` remains failed because it recursively counts the existing development `.venv`, pytest caches, and Python bytecode (9,799 findings). No workspace cleanup was attempted; those files are absent from compact `dist`.
- Excel process inventory returned only pre-existing PID 36464. Release gates remain **UNVERIFIED** for other Office builds and the remaining accessibility/security/workbook matrix.

## W221 — deployment failure integrity reporting

- A failing-first regression reproduced a per-workbook installer exception after the fake adapter changed target bytes. `deploy_multi_workbook` now reports the target's observed post-failure SHA-256 and explicitly says whether the bytes still match the pre-deployment hash. If changed, the result states whether a verified pre-deployment backup is available. A no-change failure regression confirms the unchanged hash is reported. Focused deployment suite: **16 passed**.
- `app/build.ps1` passed **427 tests** and rebuilt the onedir GUI/CLI with Python 3.11.9 / PyInstaller 6.13.0 on Windows 10 build 26300. The packaged CLI two-workbook deployment bundle apply/repeat smoke passed on Excel 16.0 build 20430.0 (applied/applied, then no_changes/no_changes, exact formulas and two verified backups); the packaged GUI disposable definition-only import also passed formula/backup/read-only reopen/exit/process checks. Pre-existing Excel PID 36464 remained untouched.
- These do not simulate native Excel failure after workbook publication and do not establish cross-workbook atomicity. Other Office versions, sync conflict and remaining desktop acceptance stay **UNVERIFIED**.

## W222 — published package after W221

- The rebuilt compact package verifies at **229 files / 135,807,226 bytes**. Manifest SHA-256: `2FF2975295C17EA2EF97BD80873F578B2AAD038639B0B10F3DDEDE8256F7CA1D`. The 227 runtime files match `app/dist` exactly; no missing, differing or extra runtime files.
- Full Windows build: **427 passed**. Deployment, package staging, and package verifier suites: **24 passed**. Contracts and 119 feature traceability references pass. Packaged CLI help exits 0; packaged GUI and two-workbook CLI Excel acceptance passed on the available build.
- This checkout has no Git metadata; release evidence files now state that commit SHAs are unavailable. Other Office builds, native partial Excel commit failure, cross-workbook atomicity, cloud-sync conflicts, screen-reader and physical display acceptance remain **UNVERIFIED**.

## W223 — parameter editor reaccepted on the rebuilt package

- Re-ran `tools/run_packaged_gui_import_smoke.py --parameter-edit` against the current one-folder GUI after W221's runtime rebuild. The packaged inspector saved only the approved literal into a separate source file and preserved the original source SHA-256. Excel read back the expected query formula from a disposable `.xlsx`; verified backup, read-only reopen, unrelated-cell marker, manager exit 0 and smoke-owned Excel cleanup passed on Excel 16.0 build 20430.0. Existing PID 36464 remained untouched.
- This repeats one Office build only; other Office builds and the remaining release matrix stay **UNVERIFIED**.

## W224 — vanished-target failure reporting regression

- Added and ran a disposable fake-installer case that removes the target and raises after a verified backup exists. Deployment continues to report `failed`, does not recreate the target path, states the final workbook state could not be verified, and identifies the verified pre-deployment backup. Focused deployment suite: **17 passed**; complete portable suite: **428 passed**.
- This regression does not simulate filesystem deletion by Excel or a native partial COM commit. The runtime executable is unchanged from W221, whose packaged CLI bundle apply/repeat and GUI import acceptance passed on Excel 16.0 build 20430.0.

## W225 — M line-comment tokenization and dependency graph

- Microsoft’s M lexical rules recognize CR, CRLF, LF, NEL, U+2028 and U+2029 line breaks. A failing-first parameterized regression reproduced four false-negative dependency results because the tokenizer ended `//` comments only at LF. The scanner now consumes through any documented line break, handling CRLF as one terminator. All six cases pass; the dependency graph dialog regression confirms a real `Consumer -> Base` edge remains visible after a U+2028 comment.
- Focused dependency scanner tests: **13 passed**; dependency graph UI tests: **5 passed**; complete portable suite: **434 passed** before the added graph UI case (**435 expected after that test; full build pending**).
- This scanner is still advisory and does not claim full M grammar support. No Excel workbook was opened for tokenizer validation; other Office builds and remaining Windows acceptance stay **UNVERIFIED**.

## W226 — rebuilt M dependency scanner package

- `app/build.ps1` passed **435 tests** and rebuilt both Windows onedir executables on Windows 10 build 26300 / Python 3.11.9 / PyInstaller 6.13.0. The compact package verifies at **229 files / 135,808,039 bytes**; manifest SHA-256 `07A9CFFC7CDCA64F3B2BDA3A5D6E3FF31662D4A72BB3CFFF2B49FAC19B57E16C`. All 227 runtime files hash-match `app/dist` (0 missing, different, or extra).
- Package staging/verifier tests: **8 passed**; contracts pass; all **122** traceability references resolve; packaged CLI `--help` exits 0. Rebuilt packaged GUI import passed with exact formula, verified backup, read-only reopen, manager exit 0 and owned-process cleanup on Excel 16.0 build 20430.0. Pre-existing PID 36464 was preserved.
- W225 scanner tests cover the six M newline forms; the dependency graph UI test verifies the detected edge and summary. The scanner remains advisory, not a full M parser. Other Office builds and remaining Windows/security/accessibility matrix cases stay **UNVERIFIED**.

## W227 — M Unicode regular identifiers

- Microsoft’s lexical specification allows regular-identifier starts in Unicode categories Lu, Ll, Lt, Lm, Lo, Nl and underscore; part characters also include Nd, Pc, Mn, Mc and Cf. Four failing-first dependency regressions reproduced false negatives for a combining mark, connector punctuation, formatting character, and Nl start. The tokenizer now follows those category sets. Scanner tests: **17 passed**; dependency graph dialog UI tests: **5 passed**.
- The scanner remains bounded/advisory and does not implement all M syntax. No query was refreshed or evaluated. A full Windows build/package follows this change.

## W228 — final package for M dependency lexical fixes

- The full Windows build passed **439 tests** and rebuilt GUI/CLI on Windows 10 build 26300 / Python 3.11.9 / PyInstaller 6.13.0. Full suite covers all M line terminators, Unicode regular-identifier categories, and the dependency graph UI summary.
- Compact `dist` verifies at **229 files / 135,809,314 bytes**; `SHA256SUMS` SHA-256 `94CA3503CEEA22A095932127E7A9283615AEA611D6DE5CD5A23275FB9AA791C2`. Its 227 runtime files match `app/dist` exactly. Package tests: **8 passed**; contracts and **123** traceability refs pass; packaged CLI help exits 0.
- The rebuilt packaged GUI import passed UI Automation, exact formula readback, verified backup, read-only reopen, manager exit 0, and owned Excel cleanup on Excel 16.0 build 20430.0. Only pre-existing Excel PID 36464 remained.
- The graph scanner remains advisory, not a full M parser. Other Office builds, provider-specific behavior, screen-reader/physical display use, and remaining workbook/security matrix cases remain **UNVERIFIED**.

## W229 — quoted identifier escape handling

- Microsoft’s M lexical specification allows character escape sequences and doubled quotes in quoted identifiers. Five failing-first cases reproduced missed query references for `#(lf)`, 4/8 digit Unicode escapes, escaped `#`, and doubled quotes; a sixth invalid-codepoint case showed that malformed escapes were incorrectly treated as a literal identifier. The tokenizer now decodes valid escapes and marks malformed escapes ambiguous without emitting a guessed dependency.
- Focused scanner tests: **23 passed**; dependency graph UI tests: **5 passed**; complete suite: **445 passed**. Windows build succeeded on Windows 10 build 26300 / Python 3.11.9 / PyInstaller 6.13.0. Contracts pass; all **125** cited traceability functions resolve. Package staging/verifier/smoke-support tests: **13 passed**.
- The refreshed compact `dist` verifies at **229 files / 135,812,357 bytes**; `SHA256SUMS` SHA-256 is `248D1A565530505A9B252CBCD12EC905CCCDE5859145F1D6D1E15B2C69A54E31`. Its 227 runtime files match `app/dist` exactly; packaged CLI `--help` exits 0.
- Packaged GUI import passed UI Automation and disposable `.xlsx` formula readback, verified backup, read-only reopen, manager exit 0, and owned Excel cleanup on Excel 16.0 build 20430.0. The pre-existing PID 36464 was present at the start and remained at smoke completion; no other Office builds were exercised.
- This extends lexical coverage but does not make the scanner a complete M parser. Excel evaluation is not used to validate tokenizer behavior; remaining Office, provider, accessibility, and security/workbook matrix gates remain **UNVERIFIED**.

## W230 — dotted M identifier matching

- Microsoft’s M grammar defines a regular identifier as an available identifier optionally followed by dot-separated regular identifiers. A failing-first regression reproduced a false edge from `Base.Field` to query `Base`; the tokenizer now keeps dotted regular identifiers whole. `Expression.Evaluate` and `Record.Field` remain dynamic/ambiguous when unquoted, quoted names with those characters remain ordinary query references, and dotted record field labels are not treated as query references.
- Focused dependency and graph tests: **32 passed**; full app suite: **449 passed**. Windows build succeeded on Windows 10 build 26300 / Python 3.11.9 / PyInstaller 6.13.0. Package staging/verifier/smoke-support tests: **13 passed**; contracts pass and all **129** cited traceability functions resolve.
- The compact package verifies at **229 files / 135,813,787 bytes**; `SHA256SUMS` SHA-256 `EB101082C6A45AACB41127C06A17C374ECE6A5B5E753285A42BBE86AF4B9FE6F`. Its 227 runtime files match `app/dist`; packaged CLI help exits 0. Package docs were refreshed after acceptance to include the W230 build and Excel smoke result.
- Rebuilt packaged GUI import passed UI Automation, disposable `.xlsx` exact formula readback, verified backup, read-only reopen, manager exit 0, and owned-process cleanup on Excel 16.0 build 20430.0. Pre-existing PID 36464 was present at the start and at smoke completion.
- This remains conservative lexical analysis, not full M parsing or external-query validation. Other Office builds and the remaining accessibility/security/workbook matrix remain **UNVERIFIED**.

## W231 — hash-prefixed M keywords

- Microsoft’s lexical grammar defines reserved `#` keywords including `#date`, `#table`, `#shared`, and `#sections`. Eleven failing-first cases reproduced false dependency edges when these tokens were split into `#` and a query-name identifier. The tokenizer now consumes hash keywords as a single token, suppresses edges for literal keywords, marks `#shared` / `#sections` ambiguous, and fails closed on unknown `#identifier` forms.
- Focused dependency and graph tests: **44 passed**; full app suite: **461 passed**. Windows build succeeded on Windows 10 build 26300 / Python 3.11.9 / PyInstaller 6.13.0. Package staging/verifier/smoke-support tests: **13 passed**; contracts pass and all **132** cited traceability functions resolve.
- The compact package verifies at **229 files / 135,815,185 bytes**; `SHA256SUMS` SHA-256 is `06F410D8BED362AB7E4487B82DCB6D8991CC200312B0AC635A18185F979BB8F3`. All 227 runtime files match `app/dist`; packaged CLI `--help` exits 0.
- Packaged GUI import passed UI Automation, disposable `.xlsx` exact formula readback, verified backup, read-only reopen, manager exit 0, and owned-process cleanup on Excel 16.0 build 20430.0. Pre-existing PID 36464 was present before and at smoke completion.
- This narrows false edges but does not provide full M grammar parsing. Dynamic `#shared`/`#sections` resolution remains advisory and ambiguous by design; other Office builds and broader accessibility/security/workbook gates remain **UNVERIFIED**.

## W232 — non-query workbook object preservation

- Extended the disposable real-Excel replacement smoke to create and verify list validation on `F2` (type and range formula), one conditional format on `G2:G4` (type/operator/formula/fill color), and a hyperlink on `H2` (address/display text). All three survived staged replacement and read-only reopen alongside the pre-existing ordinary table/filter/style, chart, shape, PivotTable, defined names, formulas, query-backed report links, load-state distinctions, and default Power Query privacy metadata.
- Source-based smoke passed on Excel 16.0 build 20430.0. The workbook uses static inline M tables and contacts no external source. Every smoke-owned Excel process exited; pre-existing PID 36464 remained.
- Compact package documentation was refreshed without rebuilding or modifying the runtime executables. Package verification passes at **229 files / 135,815,952 bytes**; `SHA256SUMS` SHA-256 is `F74868674D687D158C0E205B88B57BF12607F01AA01BFF4DBA030BFF7518E7B1`. All 227 package runtime files match `app/dist`; package staging/verifier/smoke-support tests: **13 passed**; contracts and 132 traceability references pass; packaged CLI help exits 0.
- This improves T150-03 evidence for the tested object types. Other Office builds and untested workbook relationships remain **UNVERIFIED**.

## W233 — current package keyboard navigation

- Packaged `tools/run_packaged_keyboard_smoke.py` passed against `dist/PowerQueryManager/PowerQueryManager.exe`: UI Automation visited 35 focusable controls in a complete Tab cycle, reached workbook path, Add files, query table, existing-query policy, backup and import action in order, and stayed inside the manager. It opened no workbook or Excel instance and exited 0. Pywinauto emitted its STA-mode warning before resetting to STA.
- Excel process inventory before and after contained only pre-existing PID 36464. This closes the current-build main-window keyboard traversal check; modal keyboard paths, screen-reader use, physical display scaling, other Office builds and remaining workbook/security matrix tests remain **UNVERIFIED**.
- Docs-only package refresh after W233 passed verification at 229 files / 135,816,704 bytes; `SHA256SUMS` SHA-256 is `E672AD92671DA8B58B7330F6639EEC60D2AD49EA6E7BEB42AD2DF34EBDA74D03`. All packaged runtime files match `app/dist/PowerQueryManager`; package staging tests: 3 passed; CLI `--help` exits 0. No executable payload changed.

## W234 — unresolved quoted M reference visibility

- A failing-first core regression showed `#"Missing Query"` outside the source list was displayed as “No direct query references detected.” The scanner now preserves quoted-identifier token identity and reports unmatched quoted identifiers as ambiguous, with no guessed dependency edge. A Qt graph regression verifies the ambiguous node and empty edge set. It does not claim that the query is missing from the workbook; the graph has no workbook-wide query inventory. Scanner and graph tests: **46 passed**.
- Full app suite: **463 passed**. `app/build.ps1` rebuilt GUI/CLI on Windows 10 build 26300 / Python 3.11.9 / PyInstaller 6.13.0, with **463 tests passed**. Packaged GUI replacement smoke passed on Excel 16.0 build 20430.0: exact formula readback, verified backup, read-only reopen, manager exit 0, and owned Excel cleanup; pre-existing PID 36464 remained. Rebuilt packaged keyboard traversal passed (35 controls). These Excel runs used a separate ordinary static formula and did not evaluate the unresolved reference.
- The release smoke establishes package integrity and adjacent import behavior, while the new ambiguous graph behavior is covered by core/Qt tests. Dynamic M parsing remains advisory. Full W14 missing-query/cycle effect, other Office builds, screen-reader use and physical display review remain **UNVERIFIED**.
- Final W234 docs-only refresh of the rebuilt package passed: 229 files / 135,817,933 bytes; `SHA256SUMS` SHA-256 `1841EAC79B0B159019B3B0219D3B010DFF661514021D5AEC5EB8349DC1215AD7`. All runtime files match `app/dist/PowerQueryManager`; package staging tests: 3 passed; packaged CLI help exits 0. The executable payload is unchanged by the docs refresh.
- W235’s final docs-only package refresh verifies at 229 files / 135,818,797 bytes; `SHA256SUMS` SHA-256 `8BBD7A9302F3DB4EEE54F3EAC437E76E031D658993B369E4F4FF4330407C5746`. All runtime files hash-match `app/dist/PowerQueryManager`; package staging tests passed (3), packaged CLI help exited 0, and no executable bytes changed.

## W236 — rebuilt package legacy import and replace/skip acceptance

- On Windows 10 build 26300 / Excel 16.0 build 20430.0, the current package mixed numbered `.m`/`.pq` import smoke matched review and Excel query order and exact source formulas. All-skip preserved workbook SHA-256/timestamp and created no backup/stage. Replacing the seeded connection-only query changed the formula and preserved the connection-only load association.
- All three packaged GUI runs passed verified backup where applicable, read-only reopen, manager exit 0, and smoke-owned Excel cleanup. Pre-existing PID 36464 remained. These acceptance tests used generated disposable `.xlsx` workbooks; other Office builds and broader workbook cases remain **UNVERIFIED**.
- Final W236 docs-only package refresh verifies at 229 files / 135,819,461 bytes; `SHA256SUMS` SHA-256 `3611CA0A803972996DA4FCB2170B509038107A1F47B24EA1171A49135813792C`. All runtime files match `app/dist/PowerQueryManager`; package staging tests passed (3); CLI help exited 0. No executable payload changed.

## W237 — current packaged v2 deployment, package and dialog acceptance

- Packaged CLI deployment-bundle smoke applied a hash-verified portable bundle to two disposable workbooks (`applied`, `applied`) and repeated as (`no_changes`, `no_changes`). Exact M/load state, two original-hash backups and ignoring the disabled missing source were verified; owned Excel processes exited.
- Current source deployment mixed-target smoke applied the writable workbook and returned `blocked_preflight` for the read-only target; the read-only file hash stayed unchanged and received no backup; the writable target backup/hash and formula readback passed. Packaged CLI query-package import/export roundtripped three formulas, including Unicode/punctuation, preserved the source workbook hash and did not evaluate M.
- Packaged UI Automation found named query inspector, dependency graph, profile mapping and package comparison controls; mapping edits were cancelled. The same packaged GUI completed a disposable definition-only import with exact M, verified backup, read-only reopen and owned-process cleanup.
- All tests ran on Excel 16.0 build 20430.0 with generated disposable workbooks and non-elevated app. Pre-existing PID 36464 remained. Alternate Office builds, sync conflicts, external providers and wider security matrix remain **UNVERIFIED**.
- W237 docs-only package refresh verified 229 files / 135,820,464 bytes; manifest SHA-256 `29D7C2652A53E592C6AB7A0546107B3ED82319961039914F266DC86E520F881A`. All 227 runtime files hash-match `app/dist/PowerQueryManager`; package staging tests passed (3); CLI help exited 0. Runtime binaries were unchanged.

## W238 — current package read-only install-directory acceptance

- The current one-folder package was copied to a disposable directory, granted the caller only read/execute permissions, and refused a write probe. Under a non-elevated token, the packaged GUI launched from that directory and imported a definition-only query to a separate disposable workbook. Exact formula readback, timestamped backup, read-only reopen, manager exit 0, and owned Excel cleanup passed. The package directory remained read-only and the temporary copy was removed.
- Excel 16.0 build 20430.0, Windows 10 build 26300; pre-existing PID 36464 remained. This verifies local read/execute-only directory behavior, not execution under enterprise application-control, roaming-profile, or restricted network policies; those remain **UNVERIFIED**.
- W238 docs-only package refresh verifies at 229 files / 135,821,085 bytes; manifest SHA-256 `226A69978C008FB35D223995EA8130F43A5BABA30F36A9DF14A61FFC5C083725`. All 227 runtime files hash-match `app/dist/PowerQueryManager`; package staging tests passed (3); CLI help exited 0. Runtime executables were unchanged.

## W239 — fresh full-suite and release integrity audit

- Current source suite: `python -m pytest -q` — **463 passed** on Windows 10 build 26300 / Python 3.11.9.
- Contract/schema validator passed; all **134** feature-traceability references resolve; the compact package verifier passed at **229 files / 135,821,849 bytes**; all **227** packaged runtime files match `app/dist/PowerQueryManager` byte-for-byte. Packaged CLI `--help` exited 0.
- This turn made no runtime code changes and performed no new Excel automation. Latest live Excel acceptance remains the separately recorded W235–W238 disposable-workbook tests on Excel 16.0 build 20430.0. Alternate Office builds, external providers, screen-reader and physical display use, enterprise policies, and remaining workbook/security matrix cases remain **UNVERIFIED**.
- The default `verify_handoff.py` check is for the pristine handoff and now reports post-implementation baseline changes and generated caches. This is separate from the passing compact package verifier. Git commit SHAs are unavailable because this checkout has no `.git` metadata.
- Final W239 docs-only package refresh manifest SHA-256: E7EC2D147E97064FB2638C4384905E37911FCB944A63D3BFA7A78A6EE157C0B1; verified payload remains 229 entries / 135,821,849 bytes.


## W240 — latest clean Windows package rebuild reproducibility

- Rebuilt `PowerQueryInstaller.spec` with PyInstaller `--clean` into an isolated temporary dist/work directory using `PYTHONHASHSEED=0` and `SOURCE_DATE_EPOCH=1704067200`. The current payload and clean output each contained **227 files**; missing: **0**, extra: **0**, SHA-256 differences: **0**.
- Host: Windows 10 build 26300, Python 3.11.9, PyInstaller 6.13.0. Temporary build output was removed automatically. The shipped `app/dist` and `dist` were not modified by the reproducibility check; runtime hashes remain equal. This proves same-host/toolchain repeatability for the current sources, not cross-host reproducibility.
- No Excel automation was run; Excel PID 36464 was left untouched. Office-version, connector, enterprise-policy, screen-reader, physical-display and remaining workbook/security acceptance stays **UNVERIFIED**.
- W240 docs-only refresh verifies the final package at 229 manifest entries / 135,822,550 payload bytes; SHA256SUMS SHA-256 258FFAC12F6C6E745310FD0607BDA81D645CC7C4C192C9370817FD2DDBB0EBF7. Runtime files were unchanged.

## W241 — current packaged modal keyboard navigation

- Extended the packaged keyboard smoke with optional `--dialogs` coverage and regressions for complete Tab cycles and focus escape. `app/.venv/Scripts/python.exe -m pytest -q`: **465 passed**; focused keyboard/package helper tests: **10 passed**.
- On the current compact Windows package, UI Automation traversed the 35-control main-window Tab cycle; the source-mapping dialog Tab cycles reached both mapping tables and actions; Right Arrow switched categories; Escape closed the modal. Manager exit was 0. No workbook or Excel instance was opened, and pre-existing PID 36464 remained untouched.
- This covers main-window and source-mapping modal keyboard operation only. Other modal paths, screen-reader user testing, physical display scaling and alternate Office builds remain **UNVERIFIED**. Runtime executables were not rebuilt.
- W241 docs-only package refresh verifies 229 manifest entries / 135,823,259 payload bytes; SHA256SUMS SHA-256 E728A92F157C0C47A0BFAF08C5102048E40FD450A25E52192C36F6905E1C7223. Runtime executables were unchanged.

## W242 — operation-history modal keyboard acceptance

- Extended the current packaged GUI keyboard smoke to the local operation-history modal. UI Automation completed a Tab cycle through the history table, retention period, snapshot toggle, refresh, retention apply, delete and close controls, then Escape dismissed the dialog.
- The smoke set `LOCALAPPDATA` to a temporary profile and removed it after manager exit, so user history was not read or modified. No workbook or Excel instance was opened; pre-existing PID 36464 remained. This adds keyboard evidence for the history dialog only; screen-reader, physical-display and other modal paths remain **UNVERIFIED**.
- W242 docs-only package refresh verifies 229 manifest entries / 135,823,957 payload bytes; SHA256SUMS SHA-256 FEDCA7BB01516FCDADCE923AF3F581BC453A0FC8E7ECBA03A3338984A6C841CD. Runtime files were unchanged.

## W243 — dependency-graph modal keyboard acceptance

- Extended `run_packaged_keyboard_smoke.py --dialogs`; the current packaged dependency-graph dialog's keyboard cycle reached the text summary and Close action, and Escape dismissed it. The graph had no query sources. No workbook or Excel instance was opened.
- The complete `--dialogs` smoke still passed the 35-control main cycle, source mapping path/M tab cycles, and operation-history controls using the disposable local profile. Populated graph interaction, screen-reader and physical display acceptance remain **UNVERIFIED**.
- W243 docs-only package refresh verifies 229 manifest entries / 135,824,492 payload bytes; SHA256SUMS SHA-256 FD1E91708B3451677F91F04EA43A53911A7F03959F6BD0750D1E99F24A31A686. Runtime executables were unchanged.

## W244 — packaged query-inspector keyboard navigation

- Regression-first repair: the M source, workbook and diff editors now set `tabChangesFocus=True`; Tab navigates out of each editor to the next dialog control rather than inserting indentation. A Qt regression first failed, then passed for all three editors.
- `app/build.ps1` passed the full suite (**467 passed in 23.34s**) and rebuilt the PyInstaller GUI/CLI on Windows 10 build 26300 / Python 3.11.9 / PyInstaller 6.13.0. The rebuilt compact package passed the combined Windows UI Automation smoke: main-window 35-control cycle, operation history, query inspector search/Find/editor/save-copy/close, empty dependency graph, and both source-mapping tabs.
- The smoke used an isolated temporary `LOCALAPPDATA`, tested history before inspection because inspection writes a local history entry, opened no workbook or Excel, and left pre-existing Excel PID 36464 untouched. Packaged UI process exited 0. Remaining screen-reader, physical-display, alternate Office, and broad security/workbook gates remain **UNVERIFIED**.
- Final W244 compact-package verification: **229 manifest entries / 135,825,519 payload bytes**, `SHA256SUMS` SHA-256 `372081D56B27568BA6F85EF0562D9416DCBCD25D7F9C59F6DC1702789B96409A`; all 227 runtime files match the rebuilt source payload, and packaged CLI `--help` exits 0.

## W245 — rebuilt packaged GUI Excel import acceptance

- Ran `python -B tools/run_packaged_gui_import_smoke.py --exe dist/PowerQueryManager/PowerQueryManager.exe` under a non-elevated token. The rebuilt package imported one static inline-M query in a disposable `.xlsx` through the visible GUI using definition-only mode on Excel 16.0 build 20430.0.
- UI Automation found the named manager/inspector/dependency/mapping controls and operation log. Read-only Excel reopen verified the exact formula; the timestamped backup SHA-256 matched the original workbook. Manager exit was 0. Smoke-owned Excel PIDs 34508, 43900 and 35764 exited/cleaned; before and after process sets contained only pre-existing PID 36464, which was untouched. Temporary workbook and source were removed.
- No external source or query refresh was used. This reaccepts definition-only packaged GUI import on one Office build only; alternate Office builds, providers, other load modes under this specific rebuilt executable, and remaining workbook/security/accessibility gates remain **UNVERIFIED**.

## W246 — rebuilt packaged replacement acceptance

- Command: `python -B tools/run_packaged_gui_import_smoke.py --exe dist/PowerQueryManager/PowerQueryManager.exe --replace-existing` on Windows 10 build 26300 / Excel 16.0 build 20430.0 under a non-elevated token. A disposable `.xlsx` seeded with a definition-only `GuiSmoke` query was replaced through the rebuilt GUI.
- Read-only reopen verified the exact new formula and original-hash timestamped backup; the unrelated worksheet marker remained intact. No refresh was requested. Manager exited 0. Seed PID 26924 and app-owned PID 43216 exited; verification PID 38424 was cleaned up; process set returned to `[36464]`. Existing Excel PID 36464 was untouched. Temporary workbook/source removed.
- This confirms replacement on one current package/Office build; loaded-query preservation variants and broader Office/workbook/security/accessibility gates remain **UNVERIFIED**.
- Current W246 package refresh verifies **229 manifest entries / 135,827,136 payload bytes**; `SHA256SUMS` SHA-256 `9BB0A72CD81637088DAC027E98F1C2CA73D24A4373B0CE01A751413074AD7A12`.

## W247 — current-source PyInstaller reproducibility

- After the inspector accessibility change, a clean isolated PyInstaller `--clean` build matched the current source payload byte-for-byte. Both contained 227 files; missing 0, extra 0, different 0. Environment: Windows 10 build 26300, Python 3.11.9, PyInstaller 6.13.0, `PYTHONHASHSEED=0`, `SOURCE_DATE_EPOCH=1704067200`. Temporary build output was automatically removed.
- This confirms same-host/toolchain reproducibility for the current source, not cross-host or cross-toolchain reproducibility.
- Final W247 compact package: 229 manifest entries / 135,827,789 payload bytes; `SHA256SUMS` SHA-256 `47A8848B716DBF4F68E0CA1A03DB60E8005F568232C8FB492A91AB1827269295`.

## W248 — current rebuilt package load-mode acceptance

- The rebuilt packaged GUI passed four separate disposable `.xlsx` smokes on Windows 10 build 26300 / Excel 16.0 build 20430.0 under a non-elevated token: `--new-sheet`, `--connection-only`, `--existing-sheet`, and `--data-model`. Exact formulas and original-hash backups matched after read-only reopen. Worksheet values 17/29 were verified; the existing destination was `Output!C4` and preserved its unrelated marker; connection-only created a connection with no sheet/model output; Data Model membership and row count 2 were verified.
- All four manager processes exited 0 and smoke-owned Excel processes were cleaned up. Each process set returned to `[36464]`; that pre-existing Excel process remained untouched. Temporary files were removed. No external source was contacted; inline static M was used.
- These are fresh current-package results on one Office build. Other Office versions, providers, and remaining workbook/security/accessibility cases remain **UNVERIFIED**.
- W248 compact package: 229 manifest entries / 135,828,764 payload bytes; `SHA256SUMS` SHA-256 `459A7E0B8DD70D371A960A54562BC4113A919354A908D1B04FB710A16195B6DF`.

## W249 — current packaged CLI refresh modes

- `python -B tools/run_excel_cli_refresh_smoke.py --cli-exe dist/PowerQueryManager/PowerQueryManagerCLI.exe` passed on Windows 10 build 26300 / Excel 16.0 build 20430.0. The smoke created a disposable workbook with a static inline-M query and explicit Mashup connection; it then refreshed through the packaged CLI in background and synchronous modes.
- Both JSON results reported `refreshed`. Each backup SHA-256 matched the exact pre-refresh workbook bytes; read-only post-refresh inspection confirmed the exact M formula and connection-only state. `no_external_source=true` and `owned_excel_processes_exited=true`; no external network source was contacted.
- This reaccepts packaged CLI refresh on one Office build. External providers, worksheet/Data Model refresh on this exact CLI build, and other Office versions remain **UNVERIFIED**.
- W249 compact package verifies at 229 manifest entries / 135,829,542 payload bytes; `SHA256SUMS` SHA-256 `4BBCEA17C88937471BB7CA5F4D08F345973CA34C1B2300D23D03F87A36936F63`.

## W250 — current packaged Power Query refresh and worksheet readback

- `python -B tools/run_excel_pq_refresh_smoke.py --cli-exe dist/PowerQueryManager/PowerQueryManagerCLI.exe` passed on Windows 10 build 26300 / Excel 16.0 build 20430.0. The rebuilt CLI refreshed a disposable query backed by a temporary loopback CSV endpoint in four cases: connection-only background/synchronous and worksheet background/synchronous.
- All four reports returned `refreshed`; seven requests reached the loopback server. Four backups matched their exact pre-refresh workbook hashes, the M formula and connection shape persisted, worksheet output reopened with value 42, and temporary refresh objects were removed. The only Excel process before/after was PID 36464. No external network was used.
- This covers one host/build and temporary HTTP CSV behavior; other Office builds/providers, Data Model refresh, and the remaining accessibility/security matrix stay **UNVERIFIED**.
- W250 compact package: 229 manifest entries / 135,830,358 payload bytes; `SHA256SUMS` SHA-256 `0EC3A59B09A30FF2A76178C7B35C8C57BF457948F01B2C7466BC804FF8AB20C6`.

## W251 — current packaged CLI timeout and rollback

- `python -B tools/run_excel_pq_blocked_refresh_timeout_smoke.py --cli-exe dist/PowerQueryManager/PowerQueryManagerCLI.exe` passed on Excel 16.0 build 20430.0. A disposable workbook's synchronous query waited on a delayed localhost `Web.Contents` response with a 3-second timeout.
- The request reached the loopback server; the packaged CLI exited 5 and reported timeout. Original workbook bytes were preserved, the single backup matched the original SHA-256, no stage remained, and process set returned to `[36464]`. `external_network_used=false`.
- This verifies timeout rollback for one current packaged path/build; other Office versions/providers and remaining security/accessibility matrix cases remain **UNVERIFIED**.
- W251 compact package: 229 manifest entries / 135,831,024 payload bytes; `SHA256SUMS` SHA-256 `69288982BAB19D796F3AF51E722E22D8301976AED4AC08E6981630E6FCE390DD`.

## W252 — noninteractive source-mapping approval diagnostics

- A focused regression first reproduced that `bundle --apply-path-mappings` without a supplied approval phrase was reported as unsupported. The CLI now reports `Approval required`, explains the exact `--confirm "APPLY PATH MAPPINGS"` phrase, and exits 3 for the existing fail-closed path. The same classification is applied to `plan` and `deploy`.
- Targeted command: `python -B -m pytest -q app/tests/test_cli.py -k "bundle_without_noninteractive_mapping_approval or path_mapping_decline or combines_path_and_m_code_mappings"` — 3 passed.
- Rebuilt the PyInstaller onedir package on Windows 10 build 26300 / Python 3.11.9 / PyInstaller 6.13.0 without running the broad test suite. The rebuilt CLI `--help` exited 0; `tools/verify_release_package.py` passed at 229 files / 135,831,629 payload bytes after refreshing package docs with W252. `dist/SHA256SUMS` SHA-256: `35B59CB96C893782D877B979A8924A7438B8A9BEF8B7FE00016923B4DB121D3C`.
- Real Excel mapping deployment was not repeated because this diagnostic-only change does not change mapping application. Release gates remain subject to the existing Windows matrix and unverified cases.

## W253 — noninteractive CLI approval boundaries

- `deploy`, package `import`, `refresh`, M-code mapping, and source-path mapping now distinguish a missing noninteractive approval from an explicit wrong-phrase decline. Missing approval reports the exact `--confirm` value and exits 3 before any Excel operation that depends on the approval. Wrong phrases remain fail-closed declines.
- Targeted checks: `python -B -m pytest -q tests/test_cli.py -k "package_import_without_noninteractive_confirmation or without_noninteractive_confirmation or without_noninteractive_mapping_approval or confirmation_argument_does_not_approve_wrong or refresh_decline"` — 6 passed, 27 deselected; `python -B -m py_compile src/pq_installer/cli.py` passed.
- The path-only plan regression and combined path/M-code mapping regression also passed in a focused 8-test selection (26 deselected), after first reproducing the missing approval. Rebuilt on Windows 10 build 26300 / Python 3.11.9 / PyInstaller 6.13.0 without the full suite. Packaged CLI help exited 0. Packaged bundle and path-only plan smokes both returned 3, printed the exact required phrase, and created no output file. Final compact package verification passed at 229 files / 135,832,476 payload bytes after documentation refresh; `SHA256SUMS` SHA-256 `8E101F969582DDBE0BCE956BFD3DBF44B95E3208D0CB4630510D51926E8022E8`.
- Current rebuilt GUI passed `tools/run_packaged_keyboard_smoke.py --exe dist/PowerQueryManager/PowerQueryManager.exe --dialogs`: main window (35 controls), history, inspector, dependency graph, and source mappings completed keyboard traversal; no workbook or Excel was opened. Manager exited cleanly. Process readback showed no Power Query Manager process; the only Excel PID remained 36464. The first UIA attempt hit a transient COM subscriber error, then the rerun passed.
- The complete CLI test module passed: `python -B -m pytest -q tests/test_cli.py` — 34 passed in 1.49 seconds. The full application suite was not rerun.

## W254 — workbook inventory and refresh accessibility

- Added explicit accessible names for the workbook query inventory table, export/package/refresh actions, and refresh mode, timeout, backup, continue, and cancel controls. A focused UI regression first reproduced the unnamed inventory table; targeted coverage now asserts the inventory actions and refresh options.
- `python -B -m pytest -q tests/ui/test_inventory_accessibility.py tests/ui/test_history_accessibility.py` — 5 passed. Rebuilt on Windows 10 build 26300 / Python 3.11.9 / PyInstaller 6.13.0 and verified the compact package at 229 files / 135,833,330 payload bytes. `SHA256SUMS` SHA-256 `D76A562712B7900BFEA330DFECC876B8F82F89151C456C7310FE516F87B9E193`.
- The current packaged GUI's keyboard smoke passed the main 35-control cycle plus history, inspector, dependency graph, and source-mapping dialogs. No workbook or Excel was opened; manager cleanup was verified and Excel process inventory remained `[36464]`. The inventory-specific controls are covered by Qt UI tests rather than live Excel UI Automation.
- Additional focused accessibility selection passed: 8 tests, 37 deselected, covering inventory/history names, common dialog names, and manager interactive widgets.
- Screen-reader use, OS display scaling, physical display review, and remaining release matrix items stay unverified.

## W255 — latest packaged GUI replacement acceptance

- `python -B tools/run_packaged_gui_import_smoke.py --exe dist/PowerQueryManager/PowerQueryManager.exe --replace-existing` passed on Windows 10 build 26300 / Excel 16.0 build 20430.0. The current GUI package replaced the definition of a seeded query; read-only reopen confirmed the exact new formula, verified backup matched the original, and the unrelated worksheet marker remained intact.
- The packaged manager exited 0. Owned Excel PID 41352 closed without forced termination; verification PID 9168 was cleaned up. Process set returned to the pre-existing `[36464]`. A first invocation's live command handle was not retained, so its exit status is not counted; the subsequent directly captured run exited 0.
- Compact package verification remains 229 files / 135,833,330 payload bytes; `SHA256SUMS` SHA-256 `D76A562712B7900BFEA330DFECC876B8F82F89151C456C7310FE516F87B9E193`. This reaccepts one replacement path on one Office build, not the remaining workbook/security matrix.
- Wrong phrase and missing phrase tests use injected Excel adapters and verify no installer/refresher is called. No additional Excel instance was launched for this change; existing Windows/Excel acceptance matrix gates remain as previously recorded.


## W256 — actionable missing-approval category

- Missing noninteractive approvals now use the dedicated `approval_required` diagnostic category, with guidance to review the preview and provide the exact confirmation phrase. This separates an approval boundary from malformed input in both command-line detail and local diagnostic records.
- The regression first failed because `ApprovalRequired` was not defined in the shared diagnostic module. After implementation, `python -B -m pytest -q tests/test_cli.py` passed (34 tests), the focused failure-category selection passed (3 passed, 9 deselected), and a follow-up regression verified the persisted log category (`python -B -m pytest -q tests/test_redaction.py -k missing_user_approval` — 1 passed, 11 deselected). `py_compile` passed for CLI and diagnostics.
- Rebuilt the one-folder PyInstaller package on Windows 10 build 26300 / Python 3.11.9 / PyInstaller 6.13.0. Packaged CLI `--help` and `bundle --help` exited 0. A packaged path-mapping bundle smoke exited 3 with `approval_required`, printed `APPLY PATH MAPPINGS`, and created no bundle; source and mapped input contents remained unchanged.
- Package verifier passed after documentation refresh: 229 files / 135,834,097 payload bytes. Current manifest SHA-256 is recorded with this release package after the final evidence update. No Excel was opened for this diagnostic-only change; baseline Excel PID 36464 was not touched. Full application suite and full Windows acceptance matrix remain unrun for this update; release gates remain **UNVERIFIED**.
- Final compact package: 229 manifest entries / 135,834,097 payload bytes; SHA256SUMS SHA-256 `710D3758DEB336305043A22172161789A9FBAEEA931086576C37B2FE55FB957B`.


## W257 — current packaged GUI import after diagnostic rebuild

- `python -B tools/run_packaged_gui_import_smoke.py --exe dist/PowerQueryManager/PowerQueryManager.exe` passed on Windows 10 build 26300 / Excel 16.0 build 20430.0, non-elevated. The current GUI package imported a definition-only query into a disposable `.xlsx`; read-only reopen confirmed the exact M formula and original-hash timestamped backup.
- The packaged manager exited 0. App-owned Excel PID 21856 exited without forced termination; verification PID 20092 was cleaned up. The process set returned to `[36464]`; the pre-existing Excel PID 36464 was untouched. No external source was contacted; the temporary workbook and source were removed.
- This reaccepts definition-only GUI import on one package/build after W256. Replacement, load-mode, other Office and broader security matrix results remain separately scoped as documented.

## W258 — compact Windows package

- PyInstaller collection now omits all 94 Qt translation catalogs because this application does not load Qt translators and its UI is English-only. The package verifier rejects accidental catalog inclusion. The regression first failed with an existing valid-manifest test package containing `PySide6/translations/qtbase_de.qm`; the focused verifier test now passes.
- A clean Windows 10 build 26300 / Python 3.11.9 / PyInstaller 6.13.0 onedir build produced 0 Qt translation files. The packaged CLI `--help` exited 0. The packaged GUI smoke passed main-window/dialog UI Automation and a definition-only query import into a disposable `.xlsx`; exact M, original-hash backup, read-only reopen, manager exit 0, and owned Excel cleanup passed on Excel 16.0 build 20430.0. Process set returned to `[36464]`, leaving the pre-existing Excel instance untouched.
- The Qt removal reduced runtime payload by 6,277,341 bytes (about 6.3 MB) while retaining required QtCore/Gui/Widgets DLLs and plugins. This is English-only UI packaging; localization is not supported. Remaining workbook, Office-version and security matrix items stay **UNVERIFIED**.

## W259 — second-stage Qt runtime trim

- The PyInstaller spec now filters Qt QML/Quick, PDF, virtual-keyboard, and OpenGL modules/plugins plus `opengl32sw.dll`; none is used by the app's QtWidgets/QGraphicsView interface. The package verifier rejects these runtime files. A failing-first package regression reproduced shipping `Qt6Qml.dll`; package-verifier suite passed 7 tests.
- Rebuilt with Windows 10 build 26300 / Python 3.11.9 / PyInstaller 6.13.0. The compact package has 124 manifest files / 88,553,103 payload bytes. CLI `--help` passed. Packaged UI Automation and a disposable definition-only Excel import passed on Office 16.0 build 20430.0, with exact formula and original-hash backup after read-only reopen; manager and owned Excel cleanup passed, baseline PID 36464 remained.
- Runtime payload decreased by an additional 41,004,504 bytes relative to W258 (about 41 MB). This excludes optional Qt components and catalogs; the tested use case is the packaged manager's current Windows UI and import path. Other Qt surfaces, Office builds, and remaining acceptance matrix items are **UNVERIFIED**.

- Final W259 compact package: 124 manifest files / 88,553,103 payload bytes; SHA256SUMS SHA-256 `36A3776805249AA6666A5639AA12D86F3E204A9332E14B5C3D206C3C1939EF7F`.

## W260 — remove unused Qt networking runtime

- Source inspection found no QtNetwork imports, sockets, or app-managed HTTP clients; query refresh is delegated to desktop Excel. The PyInstaller spec now omits QtNetwork and its Python binding, network-information/TLS plugins, and duplicate Qt OpenSSL DLLs. The release verifier rejects accidental inclusion, and the new regression first failed with a fake package containing `Qt6Network.dll`; package verifier suite passed 8 tests.
- Rebuilt Windows 10 build 26300 / Python 3.11.9 / PyInstaller 6.13.0 package contains 116 manifest files / 76,193,358 payload bytes. CLI `--help` passed. Packaged GUI UI Automation and definition-only import passed on Excel 16.0 build 20430.0 with exact formula, verified backup, read-only reopen and owned-process cleanup; pre-existing Excel PID 36464 remained untouched.
- Runtime payload fell another 12,361,616 bytes; package is approximately 76.2 MB. External-source refresh, other Office builds and remaining workbook/security matrix items remain **UNVERIFIED**.

- Final W260 compact package: 116 manifest files / 76,193,358 payload bytes; SHA256SUMS SHA-256 `5DDC03E6FB531348E2F842AB8E515641F92D697E725E73947851A025D927836A`.

## W261 — refresh acceptance after Qt network removal

- `python -B tools/run_excel_pq_refresh_smoke.py --cli-exe dist/PowerQueryManager/PowerQueryManagerCLI.exe` passed on Windows 10 build 26300 / Excel 16.0 build 20430.0 with the current 116-file package. A local server bound to `127.0.0.1` served CSV to Excel's Power Query engine; the application CLI itself had no Qt Network/TLS stack.
- Connection-only and worksheet refresh each passed in background and synchronous modes. All four CLI reports returned `refreshed`; ten loopback requests were received; the exact M formula persisted; four timestamped backups matched their pre-refresh hashes; worksheet output reopened with value 42; temporary workbook connections/sheets were removed. Process set returned to `[36464]`, and `external_network_used=false`.
- This confirms Excel-owned loopback refresh with the network-trimmed CLI on one Office build. External providers, authentication/privacy prompts, Data Model refresh on this exact binary, other Office builds and the remaining security/workbook matrix remain **UNVERIFIED**.

- Final W261 package: 116 manifest files / 76,193,358 payload bytes; SHA256SUMS SHA-256 `26EABD1707BE77F1B151C553D25343A3B9DD07C04812810A2F5408AE258D4273`.

## W262 — current compact-package Data Model load

- `python -B tools/run_packaged_gui_import_smoke.py --exe dist/PowerQueryManager/PowerQueryManager.exe --data-model` passed on Windows 10 build 26300 / Excel 16.0 build 20430.0, non-elevated. The current compact GUI's live capability probe enabled `data_model`; a disposable `.xlsx` with static inline M was loaded into the Data Model.
- Read-only reopen confirmed exact M, model membership, row count 2, and timestamped backup SHA-256 equal to the original workbook hash. Manager exited 0; the owned Excel PID exited without forced termination; process set returned to `[36464]`. No external source was used.
- This verifies Data Model import on one Excel build/package. Data Model refresh on this exact package, other Office versions and remaining workbook/security acceptance remain **UNVERIFIED**.

- Final W262 compact package: 116 manifest files / 76,193,358 payload bytes; SHA256SUMS SHA-256 `8647F7D5E6E4E21DEF5F7974F8263C4C3F45779C8997ED8C2466C38C84B45A94`.

## W263 — current-package Data Model refresh

- The acceptance harness now supports `--data-model --refresh-data-model`. Its first attempt exposed and fixed a missing `json` import in the harness before producing evidence; the full rerun passed: `python -B tools/run_packaged_gui_import_smoke.py --exe dist/PowerQueryManager/PowerQueryManager.exe --data-model --refresh-data-model`.
- Host/package: Windows 10 build 26300; desktop Excel 16.0 build 20430.0; current compact package; non-elevated. The app imported static inline M to the Data Model, then the packaged CLI refreshed `GuiSmoke` with exact `--confirm REFRESH`. No external source was contacted.
- CLI report returned `refreshed` in 3.031 seconds; Excel verified the model connection and 2 rows; the staged workbook passed read-only verification and committed. The backup matched exact pre-refresh workbook SHA-256; no stage remained. Read-only reopen confirmed formula, model membership and row count 2. Manager exited 0; owned Excel processes exited; process set returned to `[36464]`.
- Excel's Data Model table API exposes synchronous refresh only. Background Data Model refresh is unavailable on this build. Other Office builds/providers and remaining workbook/security matrix cases remain **UNVERIFIED**.
- Final W263 compact package: 116 manifest files / 76,193,889 payload bytes; SHA256SUMS SHA-256 `01F8798055DBB528E027AF43BE7D630B663B4892A6F2BA775152B54FF26995DF`.

- Final W264 package manifest: 116 files / 76,194,516 payload bytes; SHA256SUMS SHA-256 `2A5936442E18D270C70F57A0C5919D3BC42C04B4D387318B423A45B5BFA88CE8`.

- Final W265 package manifest: 116 files / 76,195,081 payload bytes; SHA256SUMS SHA-256 `AA7640EA027E460955F7609A02BDB335698342B93C7EE58DBEDF5AE446191E75`.

- W266 package: 116 manifest entries / 76,195,628 payload bytes; SHA256SUMS SHA-256 `AB1066B5F9AF608CBBE7937362EC5AC91B7AC747CCB0047C92B6F4A7FDC2BF2B`. Package verifier and CLI `--help` passed.

- W267 documentation-refresh package: 116 manifest entries / 76,196,191 payload bytes; SHA256SUMS SHA-256 `AE8B594BA13313AC71116DBA22C79C96574CB522B311D3A5AF96802FC115623D`. Verifier passed; executable payload was unchanged from W266.


- W269: profile migration/validation selection 7 passed; full Windows suite 484 passed; PyInstaller rebuilt both launchers. Contract validator and package verifier passed; packaged CLI `--help` exited 0. Compact package: 116 files / 76,197,823 payload bytes; SHA256SUMS SHA-256 `F176BE40041F4589B8B689014E6C75270E28E2327CE263051D606A0870B1B2FC`. Profile/Office fixture coverage and release gates remain **UNVERIFIED**.

- Packaged CLI smoke: `validate --profile` accepted a temporary schema-0 profile and relative `.pq`, returned `valid: true`/`schema_version: 1`, exited 0 and did not rewrite the original profile. Excel was not started.
- Final W269 documentation-refresh package: 116 manifest entries / 76,197,970 payload bytes; SHA256SUMS SHA-256 `84B2EADADBD5D758D52CA6CAFBA49550225E142024FB7F7FFBE021C20ECDC65C`; executable payload unchanged.


## W293 — current packaged CLI query-package import

- Ran `python -B tools/run_excel_cli_package_import_smoke.py` against the current packaged CLI and disposable source/target workbooks on Windows 10 build 26300 / Excel 16.0 build 20430.0, non-elevated.
- Three exported query definitions imported and read back with exact names/formulas, including `Équipe – Résultats 2026 (Q1)`. The source workbook hash was unchanged by export; the target backup matched its pre-import SHA-256. No external source evaluation occurred.
- Owned Excel processes exited and the process set returned to the existing baseline PID 36464. Package and release scope remain as recorded below; this single-build scenario does not close the broader acceptance matrix.
## W294 — packaged CLI rejects ZIP path traversal

- Ran the current packaged CLI `import --dry-run` against a disposable archive containing `../escape.pq` and a disposable sentinel target file.
- The CLI exited 2 with an actionable unsafe-path validation error before workbook inspection. Target SHA-256 remained unchanged, no escaped file was created, and Excel was not launched.
- This verifies one malicious archive shape through the packaged CLI; duplicate entries, symlinks, oversized archives and other hostile bundle cases remain covered by unit tests and/or **UNVERIFIED** for packaged acceptance as listed in traceability.
## W295 — current packaged GUI VBA-bearing `.xlsm`

- Ran `python -B tools/run_packaged_gui_import_smoke.py --exe dist/PowerQueryManager/PowerQueryManager.exe --xlsm --vba` on Windows 10 build 26300 / Excel 16.0 build 20430.0, non-elevated.
- The current packaged GUI imported a definition-only query to a disposable VBA-bearing `.xlsm`. Read-only reopen verified the exact M formula and VBA source preservation; the `Workbook_Open` marker remained unchanged. The timestamped backup was verified, manager exited 0, and only pre-existing Excel PID 36464 remained after cleanup.
- This covers one disposable VBA project and one Excel build; other VBA/macro-security policies and Office builds remain **UNVERIFIED**.
## W296 — refresh shipped docs to current package evidence

- Updated the package documentation generator to describe the v1.3.1–v2.0.0 implementation scope and include W293–W295 acceptance results. Ran `python -B tools/package_release.py --docs-only`; this refreshed the package manual, release summary and SHA256SUMS without rebuilding the executables.
- `python -B tools/verify_release_package.py`: **PASS**, 116 files / 76,203,691 payload bytes. `dist/SHA256SUMS` SHA-256: `946CE3E511A48F74B7D2F58C793FB3F717EB5A86508CDC37F46850E1FA64DD7D`.
- GUI/CLI executable timestamps remained 2026-10-23 08:51:49; latest evidence describes the same W282 binary. Broader release gates remain **UNVERIFIED**.
## W297 — current packaged CLI mixed read-only batch

- Extended `tools/run_excel_deployment_mixed_smoke.py` with `--cli-exe` to exercise the packaged CLI against two disposable workbooks, one writable and one marked read-only. Two harness attempts exposed stdout preamble parsing and seed-Excel baseline timing assumptions; both were corrected before the passing run.
- Command: `python -B tools/run_excel_deployment_mixed_smoke.py --cli-exe dist/PowerQueryManager/PowerQueryManagerCLI.exe`.
- Windows 10 build 26300 / Excel 16.0 build 20430.0, non-elevated: results were `applied` and `blocked_preflight`; CLI returned its partial-failure exit code 5. The writable target's backup hash and deployed formula verified; the read-only target received no backup and remained hash-identical. Packaged CLI deploy and inspect returned to the initial Excel PID set `[36464]`.
- Refreshed package documentation via `python -B tools/package_release.py --docs-only`; verifier passed: 116 files / 76,204,147 payload bytes. `SHA256SUMS` SHA-256: `33D45AEAF496D75C9F02D930B8368614E74D55BB03A3758ECC315969466CD6DC`. GUI/CLI executables were not rebuilt.
## W298 — injectable Excel workbook lifecycle protocol

- Added `ExcelWorkbookOperations` / `ComExcelWorkbookOperations` for workbook add/open/save/save-as/close. Inspector, installer, refresher and load-capability probe accept the port and route workbook lifecycle calls through it while retaining existing `excel_factory` and `excel_adapter` injection APIs. Added a fake-COM installer test and updated the source README/roadmap.
- Failing-first adapter regression initially failed because the inspector had no `workbook_operations` parameter. Fixed the test fixture's required protocol surface and added installer lifecycle assertions. Focused tests: `python -B -m pytest -q app/tests/test_excel_adapter.py app/tests/test_excel_inspector.py app/tests/test_excel_queries.py app/tests/test_excel_refresh.py app/tests/test_capabilities.py` — **91 passed**.
- Rebuilt with PyInstaller 6.13.0 / Python 3.11.9 on Windows 10 build 26300 using reproducible build environment variables. Current packaged GUI definition-only import passed on Excel 16.0 build 20430.0: exact formula, verified backup, read-only reopen, exit 0; owned Excel PIDs exited and only pre-existing PID 36464 remained. Packaged CLI `--help` and release-package verifier passed.
- Refreshed docs-only package. Final package: 116 files / 76,208,339 payload bytes; `SHA256SUMS` SHA-256 `AE23240EC59E1DBE2127849923AB025FEFA39106D0A10F9C289005E8663B44E2`. Full test suite and remaining Excel/Office matrix were not run for this change; release gates remain **UNVERIFIED**.
## W299 — packaged CLI refuses workbook open read-only without owner-lock file

- Extended `tools/run_excel_open_without_lock_smoke.py` with `--cli-exe` to test the shipped deployment CLI while Excel holds a disposable workbook open read-only and does not create a `~$` lock file.
- Command: `app/.venv/Scripts/python.exe -B tools/run_excel_open_without_lock_smoke.py --cli-exe dist/PowerQueryManager/PowerQueryManagerCLI.exe`.
- Result: **PASS** on Windows 10 build 26300 / Excel 16.0 build 20430.0, non-elevated. The CLI returned `blocked_preflight` (exit 5) with an active-use reason. Workbook SHA-256 remained unchanged; no backup/stage was created; Excel process set remained at the baseline plus the one harness-owned instance. Import, refresh and overwrite-restore service paths were also refused before stage/backup or Excel factory use.
- Cleanup: harness-owned Excel process exited; only pre-existing PID 36464 remained. Release gates for other Office builds and the wider matrix remain **UNVERIFIED**.
- Shipped documentation refreshed with `python -B tools/package_release.py --docs-only`; package verifier **PASS**, 116 files / 76,208,804 payload bytes. `SHA256SUMS` SHA-256: `9A75B59C9BA1E3A4B06865608B5556E81FEE5B83179E5A097A96DAEF026DCF28`. GUI SHA-256 `025275CE6394682FB184D53616FB0F4569BFA6ACA69E03AA5E1BDB65CEB13905`; CLI SHA-256 `1965301E28B523DD3EA6119B24A780E3437BD320A2A7BBFDE4A629E096FEB028` (both unchanged by docs-only refresh).

## W300 — undoable selected-row removal

- Added a single undo command for bulk source-row removal. Undo restores the removed `QueryRow` objects at their original positions with stable IDs and metadata; redo removes those same rows.
- Regression first failed because the prior removal cleared the undo stack. Focused UI model suite: `python -B -m pytest -q tests/ui/test_query_table_model.py` — **40 passed**.
- Updated the v1.4 roadmap and user manual. The feature is portable UI behavior; no Excel automation was needed. Full application suite and physical accessibility checks were not run for this change.
- Packaged W300 build: PyInstaller 6.13.0 / Python 3.11.9 with deterministic build environment; packaged manager keyboard cycle visited 35 controls and exited cleanly without opening Excel; packaged CLI `--help` exited 0. Docs-only package refresh verifier **PASS**, 116 files / 76,210,287 payload bytes; manifest SHA-256 `259DF541C740C1C7AAF28619EB74BB2381F39308A26142B3F700A6712DFAFF4B`. GUI SHA-256 `A87F4D4A19830E12B2141F62C4BBE0832013FA7AA8CA7977B100342133C75039`; CLI SHA-256 `1965301E28B523DD3EA6119B24A780E3437BD320A2A7BBFDE4A629E096FEB028`.

## W301 — local audit for CLI plans and deployments

- The CLI now records dry-run plan completions and deployment outcomes in the shared local `HistoryStore`. Metadata is restricted to action/status and aggregate workbook/query/action counts. Source paths, workbook paths, query names, formulas, and credentials are not included.
- The red test verified that plans previously produced no history event. Focused regression command: `python -B -m pytest -q tests/test_cli.py::test_cli_multitarget_plan_uses_inspector_worker_and_writes_no_overwrite_json tests/test_cli.py::test_cli_explicit_confirmation_argument_allows_noninteractive_deployment tests/test_cli.py::test_cli_multi_workbook_deploy_reports_per_target_results_and_partial_exit` — **3 passed**.
- Existing application-wide test behavior was not run. Both one-folder launchers were rebuilt with PyInstaller 6.13.0 / Python 3.11.9 and deterministic build variables. Packaged CLI mixed-workbook acceptance passed on Windows 10 build 26300 / Excel 16.0 build 20430.0: a disposable writable target applied with a verified backup; a read-only target was blocked before backup and stayed hash-identical; exit 5 represented the partial result; owned Excel processes returned to baseline PID 36464. The isolated temporary `LOCALAPPDATA` contained exactly one `cli_deployment_completed` event with aggregate counts and no query/formula/path data.
- Rebuilt packaged GUI definition-only import passed on the same host/build: exact formula, verified backup, read-only reopen, manager exit 0, and owned Excel cleanup; baseline PID 36464 remained. Package verifier passed at 116 files / 76,240,748 payload bytes before this evidence refresh. Broader Office, workbook, accessibility, security and full-suite gates remain **UNVERIFIED**.
- Final docs-only package refresh and verifier passed: 116 files / 76,240,748 payload bytes. `dist/SHA256SUMS` SHA-256 `DDD3A2AA93CB226B9B2F98C4FC111AB95C76CEE7D03D267724770A6B4A71E6A2`; GUI SHA-256 `9A48550167EAC70C7F84F86C03E80498F54B13E85C1B60F2E2D67AA05FB0CA70`; CLI SHA-256 `A631F4256483B10EA921CBBD463A28C88318C1C7E7361AC7AD17E6BFA9255C4E`. Packaged CLI `--help` exited 0.

## W302 — ignore duplicate source paths within an add batch

- `QueryTableModel.append_sources()` now deduplicates paths against existing rows and within the incoming batch, case-insensitively, retaining the first selected source. The regression first failed because two rows were appended for one path.
- Focused verification: `python -B -m pytest -q tests/ui/test_query_table_model.py` — **41 passed**. PyInstaller 6.13.0 rebuilt both launchers; compact package verifier passed at 116 files / 76,240,786 payload bytes.
- The rebuilt packaged keyboard smoke did not complete reliably: one run found two matching title elements; the next observed focus leave the manager during Tab traversal. A one-off window inventory observed one visible main window. This remains inconclusive, not a pass, and is deferred while core import workflows remain the priority.
- After adding the current-package Excel import result to the manual, final docs-only package refresh and verification passed: 116 files / 76,241,206 bytes. `dist/SHA256SUMS` SHA-256 `A543D52F5855F2DD8BBB83DF5B94598825561792849FFDC5F9630B3CF907E2FB`; GUI SHA-256 `EB53833FC96CBE8FBE3AD40280F6569799D1D4D5FAB223C8713CF67DD509A93C`; CLI SHA-256 `A631F4256483B10EA921CBBD463A28C88318C1C7E7361AC7AD17E6BFA9255C4E`. Only pre-existing Excel PID 36464 remained active.
- Core workflow confirmation on the W302 package: `python -B tools/run_packaged_gui_import_smoke.py --exe dist/PowerQueryManager/PowerQueryManager.exe` passed on Windows 10 build 26300 / desktop Excel 16.0 build 20430.0, non-elevated. The disposable `.xlsx` received the exact formula; timestamped backup hash matched the original workbook; read-only reopen passed; manager exited 0; only PID 36464 remained. This does not cover load modes other than definition-only or broader Office builds.

## 2026-10-27 — Dependency Finder implementation evidence

- New source-only reverse dependency engine, Access-copy DAO metadata reader,
  PySide6 dialog/worker and CLI; no SharePoint-specific API. Source code is on
  a feature branch and has not been packaged or Windows-validated.
- Focused pytest on synthetic files/fake DAO and existing dependency/CLI
  regressions: **74 passed**, including literal precision, missing DAO,
  empty scan scope, COM apartment/cleanup, CSV coverage/injection protection,
  async mapped-folder inventory and XML budget handling.
- Linux baseline prior to feature changes: **475 passed, 11 skipped, 11 failed**
  due to existing Windows/path/package assumptions. Do not interpret those
  failures as passing or silently skip them in a Windows release gate.
- Full Linux regression after this feature: **498 passed, 11 skipped, 11 failed**;
  the failing test names were identical to the pre-change baseline.
- Windows `DAO.DBEngine`, real Access files, original before/after hashes,
  Excel DataMashup and packaged GUI/CLI: **NOT VERIFIED**.

## 2026-10-27 — General Dependency Map and bounded-stream scan evidence

- Added general file-to-file dependency indexing across selected Excel and
  Access files, an outgoing/incoming GUI filter backed by one in-memory index,
  a visual advisory graph, and the `dependency-map` CLI. Find Dependents
  remains available. Quick and Deep modes control worksheet-formula coverage.
- Office XML metadata now uses incremental parsing with completed-node
  removal, and matches/unresolved filename hints are deduplicated as each
  reference arrives. A failing-first 60,000-formula stress test initially
  peaked at 9,899,400 traced bytes; the incremental-index fix passed the
  6 MiB peak threshold on that same synthetic sample. This does **not** claim
  a strict peak-RAM bound for every file or actual Excel workbook.
- The map CSV summary checks both incoming and outgoing links; a failing-first
  UI regression caught its earlier misleading `no_matches` label for a
  file with only outgoing links. Diagram arrowheads are also covered by a
  portable GUI regression.
- A fresh read-only reviewer found reverse duplicate-basename false positives
  and lost second filenames inside a single formula/Access query; three
  failing-first regressions reproduced them. Extraction now checks every
  safe file mention and resolves exact-path priority or emits ambiguity.
  A fourth failing-first regression excluded raw saved-SQL text from the
  unresolved-filename export.
- Final focused dependency/CLI/UI regression: **39 passed**. Full Linux
  offscreen suite: **514 passed, 11 skipped, 11 failed**. The 11 failing names
  are identical to the pre-feature Linux baseline and concern Windows/host
  path/stream or existing release-package assumptions; they are still failures,
  not skipped or claimed passed. Real Windows mapped folders,
  genuine Excel DataMashup, DAO Access Engine, performance on large workbooks,
  and the rebuilt packaged GUI/CLI are still **NOT VERIFIED**.

## 2026-10-27 — resolved the 11 inherited Linux regression failures

- Source environment: Linux, Python 3.12, `QT_QPA_PLATFORM=offscreen`,
  `PYTHONPATH=/workspace/scratch/ea9b0abc6b01/.deps:app/src`; `jsonschema==4.26.0` installed alongside the
  portable test dependencies. This was a **source-only** regression run, not
  a Windows packaged build or Excel/Access integration run.
- Two COM Running Object Table fake tests now avoid probing WinDLL on POSIX;
  the Windows COM path itself is unchanged.
- Four failed Internet/Restricted-zone preflight tests now exercise the full
  policy through an injectable simulated `Zone.Identifier` reader on POSIX.
  The intranet-zone acceptance test also uses the seam. Actual Windows ADS
  reading and Office Protected View remain separate, unverified integration
  gates; POSIX colon-named files are **not** treated as Windows ADS at runtime.
- Three path-mapping failures used hard-coded Windows separators in expected
  strings; their assertions now compare exact resolved paths on the host.
- The versioned query-package schema test no longer fails for lack of the
  declared `jsonschema` dependency. An additional formerly skipped test is
  exercised when the dependency is installed.
- The isolated release-package fixture now includes inert files at the two
  mandatory Shiboken paths. The **production verifier remains unchanged**
  and requires these files in a real release package; fake DLLs are never
  shipped or executed.
- Final command: `PYTHONPATH=/workspace/scratch/ea9b0abc6b01/.deps:app/src QT_QPA_PLATFORM=offscreen python -m pytest -q app/tests --tb=short`
  — **526 passed, 10 skipped, 0 failed in 17.39 s** (post-review rerun). This closes the inherited
  Linux failures but does not advance the Windows/Excel/DAO release gates.

## 2026-10-27 — Dependency Finder Windows build check

- Exact source commit tested: `c5009db52d9c321f86d5e5bd9b11e7b73fa93aeb`.
- On Windows 10 build 26300 / Python 3.11.9, `app/build.ps1` confirmed the pinned Windows dependencies, passed the full suite (**536 passed in 37.79s**) and rebuilt both onedir executables with PyInstaller 6.13.0.
- Release staging and verification passed: **116 files / 76,310,185 bytes**. Packaged CLI help exposes `dependency-scan`, `dependency-map`, and Quick/Deep depth options. Packaged GUI launched; its accessibility tree showed the Dependency Finder and Dependency graph controls.
- GUI scan behavior was not run. No `.accdb` or `.mdb` fixtures were found in the checkout or inspected fixture locations, so actual DAO linked-table/saved-query scanning cannot be claimed. Genuine workbook Mashup and external-link detection, scan filters, graph/export accuracy, large real-workbook Quick/Deep timing/RSS, local/SMB/UNC/synchronized folders, and scan-level unchanged-file/lock/temp checks remain **NOT VERIFIED**. Required acceptance is **FAIL / OPEN** until these integrations are exercised.
- No source failure was found in the Windows suite; no repair was made. Portable tests for scanner behavior do not establish Office/DAO acceptance. Package executable hashes are recorded in generated `dist/SHA256SUMS`.

## 2026-10-27 — Dependency Finder harness and integration diagnostic

- Added a repeatable Windows harness at `tools/run_dependency_windows_acceptance.py` for real ACE DAO databases, desktop Excel Mashup/external links, expected graph assertions, Quick/Deep scans, packaged CLI multi-target scans, GUI filtering/graph/exports, progress/cancellation, hashes/timestamps, lock/temp files, process cleanup, and optional mapped SMB/UNC/synchronized roots.
- Fixed Office customXml Mashup detection for UTF-16LE/BE marker encodings. Focused regression module: **12 passed**. Windows `app/build.ps1` full suite: **537 passed in 39.46 s**; rebuilt packaged GUI and CLI. Package verifier: **116 files / 76,310,327 bytes**. Packaged CLI help shows `dependency-scan`, `dependency-map`, Quick and Deep options.
- Harness diagnostic ran against base revision `7c9d94b5f3d7502b0204de7873042604fb36ee24` with then-uncommitted changes, using available `DAO.DBEngine.120`. It is not evidence for a clean commit. Required `DAO.DBEngine.160` was absent, so ACE 16.0 acceptance is **FAIL / OPEN**.
- Excel 16.0 build 20430.0 created genuine Mashup and external workbook fixtures, plus a 31,200-cell external-formula workbook. Quick found 12 dependency records in 6.43 s with 189,259,776-byte peak process-tree working set; Deep found the same 12 in 3.63 s with 186,060,800-byte peak. Both returned the eight expected pairs, one ambiguous reference and two unresolved references. GUI directional filters, 12-edge graph, JSON/CSV exports, progress and cancellation passed. Multi-target CLI scan passed with six relationships.
- Source SHA-256 and timestamps stayed unchanged; no Access lock files, leftover DAO temporary directories, or post-run Excel processes remained; disposable fixtures were removed. Excel did require forced termination after failing to exit naturally for 21.33 s, making cleanup status **FAIL**. Mapped SMB, UNC and synchronized folders were **NOT RUN**.
- Full acceptance remains **FAIL / OPEN** pending ACE DAO 16.0, natural Excel process exit and mapped/UNC/synchronized repetitions. See `docs/QA_WINDOWS_MATRIX.md` for the detailed matrix and exact harness invocation. Final committed revision and package hashes follow after commit.
- Clean-revision acceptance preflight on commit `f542a72f5c9087b44dae85cff614f5b0ed07b89f` confirmed the default required-provider gate: **FAIL**, because `DAO.DBEngine.160` is not registered (only `DAO.DBEngine.120` is available). No fixtures were created in this ACE 16.0 preflight. Full diagnostic details from the working-tree build are in `%TEMP%\\pqm-dependency-harness-final.json`; its initial dirty-tree SHA is explicitly not treated as the tested commit.

## 2026-10-27 — Excel COM cleanup repair and repeated path/performance evidence

- Supersedes the prior forced-cleanup diagnostic. The harness no longer initializes a parent COM apartment just to inspect the Excel PID set; it creates DAO fixtures in a short-lived worker before starting a separate Excel automation lifecycle. A full sequential DAO→Excel fixture exercise including a 31,200-formula workbook exited naturally in **3.63 s**, with no forced termination or Excel process leak. `Excel.Quit()` is unconditional and cleanup exceptions are recorded.
- Three interleaved Quick/Deep repetitions passed for local and OneDrive synchronized locations using the packaged CLI, with DAO 12.0 diagnostic fixtures. Local Quick median **3.842 s**, range **2.922–8.667 s**, peak working set **120,893,440 bytes**; Deep median **3.841 s**, range **3.646–3.996 s**, peak **118,755,328 bytes**. OneDrive Quick median **3.241 s**, range **2.821–7.259 s**, peak **120,717,312 bytes**; Deep median **3.546 s**, range **3.362–3.615 s**, peak **118,595,584 bytes**. Each mode returned 12 dependency records; both detected the external-link edge for the 31,200-cell formula workbook through external-link relationship metadata. No mode-only edges appeared in this fixture.
- Actual reads passed source SHA-256/size/mtime preservation, no Access lock files, temporary directory cleanup, Excel process-set restoration and removal of disposable fixtures on both roots. Local GUI filter/graph/export/progress/cancellation checks passed. OneDrive path scan passed.
- Default ACE 16.0 preflight still fails before creating fixtures: only DAO 12.0 is registered. The session is non-administrator; no Office/Access runtime installation was attempted. `net use` and `Get-SmbMapping` show no mapped SMB connection, so mapped SMB/UNC are **NOT RUN**. OneDrive is the only synchronized-path acceptance in this run. Required Windows integration acceptance remains **FAIL / OPEN**.
- The independent GitHub AI_Context repository review dated 2026-10-26 still classified Power Query Manager's packaged Windows acceptance as incomplete; this run advances that evidence only for the DAO 12 diagnostic and OneDrive path. ACE 16.0 and network share claims remain open.

## 2026-10-27 — clean-revision W296 confirmation

- Exact tested commit: `a13ff507e76e4cca80fd204c69781a033ffbed39`. The complete DAO 12.0 diagnostic harness passed on local and OneDrive synchronized roots. Excel exited normally in 3.23 s, no forced termination, no process leak. All source hashes/sizes/timestamps were unchanged; there were no Access lock files or scan temp directories; both disposable trees were removed.
- Three interleaved packaged Quick/Deep runs on local: Quick median **3.060 s**, range **2.910–3.307 s**, peak working set **190,595,072 bytes**; Deep median **3.501 s**, range **3.388–3.682 s**, peak **192,208,896 bytes**. OneDrive: Quick median **3.174 s**, range **3.112–3.318 s**, peak **194,142,208 bytes**; Deep median **3.610 s**, range **3.565–3.733 s**, peak **191,352,832 bytes**. Each mode returned 12 dependency records for the 31,200-formula fixture, plus two unresolved references and one ambiguity issue. Local GUI checks passed.
- ACE 16.0 preflight on this same SHA failed closed because the provider is not registered; DAO 12.0 is the only available provider. This session is non-administrator, so no runtime/Office install was attempted. No SMB mappings or `net use` connections exist; mapped SMB/UNC remain **NOT RUN**. Overall Windows acceptance remains **FAIL / OPEN**.

## 2026-10-27 — explicit two-gate waiver for Dependency Finder PR #1

- Repository owner authorized **waiving** the ACE DAO 16.0 and mapped SMB/UNC acceptance checks and allowed merge if appropriate. These are **WAIVED / NOT VERIFIED**, explicitly **not PASS**. The clean ACE 16.0 preflight was blocked by unregistered `DAO.DBEngine.160`; DAO 12.0 tests do not prove 16.0 compatibility. No mapped SMB/UNC test environment was available; OneDrive is not a substitute.
- Evidence retained unchanged: clean tested source `a13ff507e76e4cca80fd204c69781a033ffbed39`; 537 Windows tests reported for the build; repeatable local and OneDrive DAO 12.0 integration passed; genuine Excel Mashup/external links, Quick/Deep, packaged CLI, GUI filters/graph/exports/cancellation and file-integrity checks passed; Excel exited naturally in 3.23 s with no residual process. Refer to W296 for precise benchmark values and fixtures.
- **Feature PR disposition:** accepted **with the two named waivers** for integration; this documentation change does not rerun Windows tests or declare the full application's broader release gates complete. Supported tested paths: local and OneDrive-synchronized folders on the documented Windows host with DAO 12.0. ACE DAO 16.0 and mapped SMB/UNC remain unvalidated and must not be represented as verified.
- Production scan behavior, fail-closed diagnostics, and test harness requirements are unchanged. This waiver does not authorize editing, refreshing or deleting source files.

## 2026-10-27 — Vector UI implementation candidate (not release accepted)

- Selected visual reference: Vector / Precision Engineering, using the three-concept design studio supplied in this conversation.
- Working GitHub branch: `feature/vector-ui-redesign`, based on merged Dependency Finder main. The implementation preserves the existing query model, search/filter, undo/redo, import approvals, workbook security paths, worker wiring, DAO scanning, CSV/JSON exports and graph/index objects. No scanner or Excel/Access service source changed.
- Replaced presentation with shared Vector Qt stylesheet supporting dark default and light alternative. Main window now has a navigational rail, compact cards, live file/source/enabled/plan dashboard metrics, prominent review/import, and Ctrl+K searchable commands dispatched through existing enabled buttons. Finder gets a scrollable structured layout, readable table, persistent coverage and scan controls. File graph adds zoom/pan/selection, visible legends, high contrast nodes and deduplicated visual arrows (full underlying edge records retained); query graph colors are aligned.
- Regression changes cover KPI updates, original action presence, inspector and target behavior, scrollable finder, graph selection, zoom, summary and legacy accessibility/theme assertions. Existing checks for worker cancellation, source integrity, no-refresh scanning and query import remain unchanged.
- **VALIDATION STATUS: NOT YET EXECUTED.** Neither PySide6 nor the Windows/Excel/DAO runtime is available in the current coding environment. No claim of successful pytest, screenshot verification, packaged EXE rebuild or Windows smoke testing. Before merging: run `python -m pytest -q app/tests/ui` and full Windows `app/build.ps1`; verify packaged Qt startup, dialog scanning, Ctrl+K, 960×600 / 1366×768 / 125% / 150% DPI, and that all original actions are reachable. Repeat disposable workbook import/readback and local/OneDrive Dependency Finder read-only checks on the compiled build.
- ACE DAO 16.0 and mapped SMB/UNC remain explicitly WAIVED/NOT VERIFIED under W297 for the earlier feature; this redesign does not convert those waivers into passes.

## 2026-10-28 — Vector graph viewport correction and Windows package

- Fixed initial graph panning by fitting the dependency diagram after the first dialog layout; added a 960×600 scene-bounds regression assertion. App-source commit `536ee6ce5d901eabbe7345c8c2ffc287cc31d372`.
- `app/build.ps1`: 541 tests passed (41.26 s); PyInstaller build succeeded. `tools/verify_release_package.py app/vector-release`: 116 files, 76,330,146 bytes, PASS. GUI and CLI EXE hashes are recorded in W298 of `docs/QA_WINDOWS_MATRIX.md`. CLI help passed; packaged GUI launched.
- Disposable file-map scan detected the expected exact `Dashboard.xlsx` → `Inventory.xlsx` edge; pre/post SHA-256 hashes matched. The rebuilt package now passes a fresh first-show graph visual check with both nodes and edge visible; see the follow-up entry below.
- PR #2 remains open pending broader interactive workflow and physical display/DPI checks. ACE DAO 16.0 and mapped SMB/UNC retain explicit WAIVED / NOT VERIFIED status.

## 2026-10-28 — fresh packaged Vector graph check

- On the freshly rebuilt package from app-code commit 536ee6ce5d901eabbe7345c8c2ffc287cc31d372, a Quick scan of synthetic Dashboard.xlsx and Inventory.xlsx returned one exact internal edge. The graph auto-fit on first show with both nodes and edge visible; accessible text confirmed Dashboard.xlsx depends on Inventory.xlsx [exact_path]. GUI process exited and was absent in the follow-up process check. See W298 for package hashes and test context.
- The test was performed at the current 1180×730 app-window size. 960×600 and 1366×768 display layouts, 125%/150% DPI, and remaining manual workflows are still open. ACE DAO 16.0 and mapped SMB/UNC remain WAIVED / NOT VERIFIED, never PASS.
- Further package interaction at 1180×730 confirmed Ctrl+K opens the command palette with its existing actions, the inspector checkbox exposes the Source M / Workbook M panel, and the light theme keeps workspace and Finder controls legible with a visible focus outline. Finder sections and Quick/Deep selectors were visible; no scan was run in this theme check. Full keyboard-only navigation and broader workflows remain unverified.

## 2026-10-28 — packaged Vector source and profile round-trip

- On package rebuilt from app-source commit `536ee6ce5d901eabbe7345c8c2ffc287cc31d372`, Windows GUI added synthetic `StaticTable.m`, showed its source row and metrics, and opened the read-only source inspector. The workbook comparison was correctly unavailable/unknown without an inspected workbook query.
- The packaged UI saved and loaded a JSON profile and reported one source row loaded. Synthetic source and workbook hashes matched their pre-check hashes; the workbook was selected as a target path but was not opened or modified. App exited normally and the process was absent afterward.
- Live window measured 1182×732. Attempts to set 960×600 did not resize it; physical 960×600 / 1366×768 and OS DPI verification remain open. The Enabled table heading looked clipped although the accessibility label was complete; investigate before release acceptance. History, full import, Finder scan/graphs, filters, exports, progress and cancellation remain open for live UI validation. Full test/build/package results and EXE hashes are recorded in W298. ACE DAO 16.0 and mapped SMB/UNC remain WAIVED / NOT VERIFIED. PR #2 stays open and unmerged.

## 2026-10-28 — Vector table heading fix and packaged acceptance

- App-source commit `3e45bc450482b5311a6a30938c27ea19fca50a45` fixes the clipped Enabled column heading with an initial 132 px resizable section and a UI regression assertion. `app/build.ps1` passed **542 tests in 44.40 s**; UI suite passed **92 tests** at both simulated 1.25 and 1.5 scale factors. Rebuilt package verification passed: **116 files / 76,330,196 bytes**. GUI SHA-256 `CFD598065D357484D3C79935E2465587858FBAD4D7930A051F706A48EDD646AB`; CLI SHA-256 `6B30D61A87A2A164949C35EFB8BEEE2226B40344741CD64BA974365F752F0382`.
- Fresh packaged UI screenshot at 1182×732 shows the complete Enabled heading and synthetic source row. Packaged Excel GUI create, skip, replace and mixed three-source import passed with exact readback and safety checks; snapshot compare/restore and package compare dialogs passed. See W300 for exact evidence.
- Current packaged CLI/GUI DAO 12.0 acceptance passed on 10 local synthetic Excel/Access fixtures: unchanged file hashes/sizes/timestamps, no lock/temp artifacts, Quick/Deep 12-edge parity over 31,200 formula cells, verified JSON/CSV exports and filters, progress/cancel, natural Excel exit, no residual processes. This does not verify DAO 16.0.
- Physical 960×600 / 1366×768 and OS-level 125%/150% DPI checks remain open; CUA resize did not change the window from 1182×732. Full keyboard-only coverage and remaining manual controls remain open. ACE DAO 16.0 and mapped SMB/UNC remain WAIVED / NOT VERIFIED. PR #2 remains open and unmerged.

## 2026-10-28 — physical Vector display observation

- On packaged app-code commit `3e45bc450482b5311a6a30938c27ea19fca50a45`, inspected the maximized app at physical 1366×768/100%; workspace fit without visible clipping. At 1366×768/125%, metric cards reflowed to two columns and the observed workspace remained scrollable without visible clipping.
- Windows scale selector offered 100% and 125% only. Physical 150% is NOT VERIFIED; Qt scale-factor suites passed at 1.25 and 1.5 but are not physical DPI evidence. Physical 960×600 is NOT VERIFIED because it was not offered as a display mode; Qt layout checks cover it. Settings restored and confirmed at 1920×1080/100%.
- ACE DAO 16.0 and mapped SMB/UNC remain WAIVED / NOT VERIFIED. PR #2 remains open and unmerged pending open acceptance gates.

## 2026-10-28 — live package Finder and keyboard accessibility

- Verified the staged release again: 116 files / 76,330,196 bytes; GUI SHA-256 `CFD598065D357484D3C79935E2465587858FBAD4D7930A051F706A48EDD646AB`; CLI SHA-256 `6B30D61A87A2A164949C35EFB8BEEE2226B40344741CD64BA974365F752F0382`.
- Packaged GUI Quick general map on disposable Dashboard.xlsx / Inventory.xlsx returned one exact edge; file graph showed both nodes and directed edge. Node selection and Fit graph worked. Incoming scan targeting Inventory.xlsx found Dashboard.xlsx as one dependent. Targeting both files intentionally produced the documented partial `no_consumers` result because no candidate files remained to inspect.
- Keyboard smoke passed with 43 main controls and dialog cycles for history, inspector, query graph, and both mapping tabs. Smoke used temporary profile/source and no Excel instance. App and smoke processes exited cleanly.
- Physical 960×600 and 150% DPI remain NOT VERIFIED; ACE DAO 16.0 and mapped SMB/UNC remain WAIVED / NOT VERIFIED. PR #2 remains open and unmerged.

## 2026-10-28 — Finder long-path layout repair and refreshed package

- Exact tested app-code commit: `a6b32e733fff3f48daf42cf1f90e1a93d840a5bf`; change keeps full path text and supplies full-path tooltips while eliding long table cells within fixed column widths. Automatic content sizing was removed because it caused horizontal overflow.
- Regression: `test_map_results_keep_long_paths_without_horizontal_table_overflow` passed at 960×600, asserting full item text and tooltips, no table horizontal scroll, and column minimums within the window. Targeted UI tests: 2 passed. Full Windows build/test suite: 543 passed in 52.81 s; `app/build.ps1` completed.
- Refreshed `app/vector-release` from that build; release verifier PASS, 116 files / 76,330,151 bytes. Packaged GUI launched successfully; packaged CLI `--help` listed all supported commands. GUI SHA-256: `1BBB016C3C1FAD2062A23370BFC99397CB8D395A14BACE5C634D7D3E66A4B804`; CLI SHA-256: `6B30D61A87A2A164949C35EFB8BEEE2226B40344741CD64BA974365F752F0382`.
- Existing physical layout notes are in W302; Finder and 43-control/dialog keyboard smoke in W303. The refreshed release has not yet had the prior disposable-data graph scenario replayed after rebuild. ACE DAO 16.0 and mapped SMB/UNC remain WAIVED / NOT VERIFIED. Physical 960×600 and 150% remain NOT VERIFIED. PR #2 stays open and unmerged.

## 2026-10-28 — refreshed package live GUI map replay

- Tested the untracked refreshed release package built from app-code commit `a6b32e733fff3f48daf42cf1f90e1a93d840a5bf` (current feature branch HEAD is documentation-only commit `97f6ebabc2291df12bc8c4f667fd218e127a3f8b`). Package verifier previously passed at 116 files / 76,330,151 bytes; GUI SHA-256 `1BBB016C3C1FAD2062A23370BFC99397CB8D395A14BACE5C634D7D3E66A4B804`; CLI SHA-256 `6B30D61A87A2A164949C35EFB8BEEE2226B40344741CD64BA974365F752F0382`.
- Live packaged Dependency Finder used only disposable `Dashboard.xlsx` and `Inventory.xlsx`. “Map all selected files and their dependencies” in Quick mode finished read-only with `Mapped 2 file(s); 1 internal dependency edge(s); 0 outside-scope or unresolvable file reference(s)`. Results table showed the Excel connection edge, source `SyntheticSource`, confidence `exact_path`.
- Refreshed package graph screen showed both workbook nodes and the exact directed Dashboard.xlsx → Inventory.xlsx edge; accessible text repeated `Dashboard.xlsx depends on Inventory.xlsx [exact_path]`. Clicking both nodes updated the selected-file label; Fit Graph responded. Finder’s result cells elided the long visible paths on this run. Full-text and tooltip preservation is covered by the automated regression in the tested app-code commit; tooltip behavior was not separately rechecked during this interaction.
- This was a live GUI read-only scan over the existing synthetic fixtures; it did not open Excel or change them. This observation does not establish the remaining workflows, Excel integration gates, physical resolution/DPI matrix, or clean shutdown. ACE DAO 16.0 and mapped SMB/UNC remain WAIVED / NOT VERIFIED; physical 960×600 and 150% DPI remain NOT VERIFIED. PR #2 remains open and unmerged.

## 2026-10-28 — refreshed package query graph and inspector

- On the refreshed packaged GUI built from app-code commit `a6b32e733fff3f48daf42cf1f90e1a93d840a5bf`, loaded only disposable `Base.pq` and `Final.pq` (`Final` references `Base`). Main query table showed both source rows and computed order `1. Base → 2. Final`.
- Query dependency graph showed both nodes, the connecting edge, and accessible text `Final depends on Base`; graph opened at the current window size. Inspector opened for selected `Final`, displayed source expression `let X = Base in X`, and marked the workbook formula comparison unknown because no workbook was selected. Closed without saving. No workbook was opened or changed.
- This confirms these two screens on the refreshed package only; it does not close remaining acceptance gates for physical DPI/resolution, broad workflow keyboard coverage, Excel/Access integration or clean shutdown. ACE DAO 16.0 and mapped SMB/UNC remain WAIVED / NOT VERIFIED; physical 960×600 and 150% DPI remain NOT VERIFIED. PR #2 remains open and unmerged.

## 2026-10-28 — refreshed package Finder, profile, import and no-op acceptance

- Current feature branch head `de619c62ad02cb4abae76855112168c94330ee25` is documentation-only after app-code commit `a6b32e733fff3f48daf42cf1f90e1a93d840a5bf`; QA harness revision was `de619c62ad02cb4abae76855112168c94330ee25`. `app/vector-release` verifier **PASS**: 116 files / 76,330,151 bytes. GUI SHA-256 `1BBB016C3C1FAD2062A23370BFC99397CB8D395A14BACE5C634D7D3E66A4B804`; CLI SHA-256 `6B30D61A87A2A164949C35EFB8BEEE2226B40344741CD64BA974365F752F0382`. Packaged CLI `--help` succeeds. The full Windows build/test result for this app-code revision remains 543 passed in 52.81 s (recorded in the preceding 2026-10-28 Finder long-path repair entry).
- W306 DAO 12.0 local disposable dependency diagnostic **PASS**: 10 fixtures; 12 expected edges in Quick and Deep; hashes/sizes/timestamps unchanged; no Access locks or DAO temp folders; Excel PID set restored. Packaged GUI exercised incoming/outgoing/both filters, graph, verified JSON/CSV exports, progress and cancellation after one inspected file; multi-target reverse scan matched six relationships. DAO 12.0 is diagnostic only; ACE DAO 16.0 remains explicitly WAIVED / NOT VERIFIED.
- W307 packaged existing-query replacement **PASS** on Windows 10 build 26300 / Excel 16.0 build 20430.0: exact formula readback, verified pre-import backup, unrelated worksheet marker preserved, manager exit 0, Excel closed naturally. W308 profile-backed 3-source mixed `.m`/`.pq` import **PASS**: review order matched Excel order, formulas read back exactly, backup/read-only reopen verified, no forced Excel termination. W309 all-skip **PASS**: workbook hash/timestamp/formula unchanged, no backup/stage, no-change result; manager/Excel cleanup passed. W305 history dialog/recovery smoke **PASS** with isolated corrupt-history fixture and clean exit.
- These are scoped Windows passes on synthetic data, not full product certification. Physical 960×600 and 150% DPI remain NOT VERIFIED; physical 1366×768 at 100% and 125% are recorded in W302 on the prior package revision. Broader workbook/load/security/accessibility matrix remains open. ACE DAO 16.0 and mapped SMB/UNC are waived. PR #2 is open/draft; it stays unmerged pending remaining non-waived acceptance and review.
