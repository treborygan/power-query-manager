# Local helper scripts

- `verify_handoff.py`: validates documentation, gate references and pristine `app/` SHA-256 hashes. After Codex changes the source, use `--allow-modified-app`. Optional `--run-tests` requires pytest.
- `validate_contracts.py`: optional full JSON Schema Draft 2020-12 check for contracts and sample; requires `jsonschema` installed.
- `repro_unsafe_replace.py`: proves the baseline delete-before-safe-replace issue using fake COM; use `--expect-fixed` after implementing safety. This does **not** change real workbooks.
- `run_tests.ps1`: cross-project pytest wrapper using local venv or system Python, no GitHub Actions.
- `check_windows_env.ps1`: readonly preflight for Windows commands and registered Excel.Application ProgID.
- `run_excel_refresh_smoke.py --mode background|synchronous`: opt-in real-Excel smoke that creates and refreshes only a disposable workbook containing a static inline M table; it contacts no external source and checks backup, reopen, and owned-process cleanup.
- `run_excel_import_smoke.py`: exercises capability-probed connection-only staged create/reimport and read-only verification on a disposable workbook.
- `run_excel_deployment_mixed_smoke.py`: deploys to two disposable `.xlsx` workbooks with one set read-only; verifies the writable target applies with a correct backup while the read-only target is blocked before inspection/backup and remains byte-identical.
- `run_excel_load_smoke.py`: probes and exercises worksheet/Data Model load create, repeat import, changed-M update, values, unrelated-cell preservation, reopen verification, and owned-process cleanup on disposable static-M workbooks.
- `run_excel_preserve_smoke.py`: replaces query formulas in a disposable workbook and verifies table/filter/style, data validation, conditional formatting, hyperlink, chart, shape, PivotTable, report links, defined names, calculation state, query load states, privacy metadata, and owned-process cleanup.
- `run_excel_password_safety_smoke.py`: creates an open-password disposable workbook and verifies inspection/import time out, preserve the original hash, and stop only their owned Excel processes.
- `run_excel_cli_environment_mapping_smoke.py`: combines reviewed path-root and M-code mapping over two disposable definition-only queries through CLI deployment; verifies the M diff uses the path-mapped source, formulas after reopen, unchanged source bytes, backup, and owned Excel process exit. Pass `--packaged` to exercise the interactive packaged CLI. Its `File.Contents` literal is not evaluated.
- `run_excel_cli_package_import_smoke.py`: exports a live disposable workbook inventory to a versioned package, imports it through the packaged CLI into a second `.xlsx`, and verifies Unicode/punctuation query-name readback, exact formulas, source/backup hashes, and process cleanup.
- `run_excel_history_compare_smoke.py`: captures two local snapshots around static-M formula/addition and capability-probed definition-only → connection-only changes in a disposable `.xlsx`, runs the real snapshot comparison worker, and verifies status/diff/load results, unchanged snapshot hashes, and owned-process cleanup without refresh.
- `run_excel_recovery_smoke.py`: creates a query-bearing disposable `.xlsx`, copies it to the abandoned-stage filename pattern, runs the recovery worker, and verifies the separate recovered copy reopens read-only with the exact query/formula while source/stage hashes remain unchanged and owned Excel instances exit.
- `run_excel_open_without_lock_smoke.py`: opens a disposable workbook read-only in real Excel without an owner-lock file, then confirms `validate_workbook_for_write` finds the exact ROT document path and refuses before writes; it checks unchanged bytes and owned-process exit.
- `build_and_smoke.ps1`: runs existing Windows build, verifies file paths, optionally starts EXE for manual startup acceptance. Does not make live COM tests automatic.
- `package_release.py`: stages the built `app/dist/PowerQueryManager` onedir payload into the handoff `dist/`, writes the compact README/release summary, and regenerates `SHA256SUMS`. Invoked by `build_and_smoke.ps1` after a successful app build.
- `verify_release_package.py`: confirms the staged package has the required GUI/CLI launchers and user docs, contains no development/cache/fixture/log directories, and has exact SHA256 manifest coverage with no missing, extra, duplicate, unsafe, or mismatched paths. Run with `python tools/verify_release_package.py`.
- `run_packaged_keyboard_smoke.py --exe dist/PowerQueryManager/PowerQueryManager.exe`: verifies a main-window Tab cycle without opening Excel. Add `--dialogs` to also traverse the packaged source-mapping modal, switch mapping categories with Right Arrow, and dismiss with Escape.

All scripts deliberately avoid opening/editing the user's workbooks. Real Excel acceptance requires disposable test workbooks and explicit user operation. The worksheet/model and password smokes intentionally run only built-in static fixtures; they do not validate external data sources or user M.

- With --dialogs, 
un_packaged_keyboard_smoke.py also traverses operation-history controls using a temporary LOCALAPPDATA profile and deletes that profile after the manager exits.

- --dialogs also checks the dependency-graph modal's text summary and Close action; Escape dismisses it. The graph is empty and no workbook or Excel process is opened.

- `run_packaged_keyboard_smoke.py --exe dist/PowerQueryManager/PowerQueryManager.exe --dialogs` also traverses operation-history, query-inspector, dependency-graph, and source-mapping dialogs in a temporary local profile. The inspector's source/workbook/diff editors transfer focus on Tab; history is checked before inspection because inspecting a source creates a local history row. This opens no workbook or Excel.

- Current package repeatability (W247) was checked by a clean PyInstaller `--clean` rebuild into an isolated temporary directory using `PYTHONHASHSEED=0` and `SOURCE_DATE_EPOCH=1704067200`; compare all relative SHA-256 entries against `app/dist/PowerQueryManager`. Current result: 227 files, zero missing/extra/different. Same Windows/toolchain only.


The PyInstaller build omits Qt translation catalogs and unused QML/Quick, PDF, virtual-keyboard, OpenGL, and Qt Network/TLS components. The application interface is English-only. The manager does not use Qt networking; explicit query refresh is performed by desktop Excel and may contact sources configured in M. `verify_release_package.py` rejects accidental inclusion of these unused runtime components.
