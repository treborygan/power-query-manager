# Paste everything below into Codex

You are the implementation lead/orchestrator for **Power Query Manager**, the successor to Power Query Installer v1.3.0. A self-contained source + specification package exists in this working directory. **Execute the implementation; do not deliver another plan alone.** The complete original Python/PySide6/Excel-COM project is in `app/`. The user approved the 4-release roadmap: v1.3.1 stability, v1.4.0 manager UI/inspection/export/diff/profiles, v1.5.0 Excel loading/dependencies/refresh, v2.0.0 advanced tooling. Aim to finish as much as the environment permits, with honest milestone gates.

Before editing:
1. Read `AGENTS.md`, `START_HERE.md`, **every** `docs/*.md`, `agents/*.md`, `contracts/*.json`, and `app/README.md`; list real app files and understand the current tests; verify hashes with `python tools/verify_handoff.py`.
2. Run the baseline unit suite from `app/` and record actual baseline result and environment in `evidence/baseline.md`. Do not use historical claims as proof.
3. Produce a short architecture/ownership plan with phase-specific tasks, acceptance tests and risks. Then **write code, tests, build scripts, release documentation, and evidence**, not just a new checklist.

Execution model: use an orchestrator plus up to three specialist agents if the execution environment supports true parallel agents/worktrees. If not, emulate the sequence in one agent while preserving ownership boundaries. Start independent engine/UI/QA tracks after the orchestrator freezes `contracts/profile.schema.json`, `contracts/plan.schema.json` and `docs/ARCHITECTURE.md`. See `AGENTS.md` and `agents/` for file ownership, review protocols and phase gates. Never let agents edit the same files concurrently; assign branch/worktree ownership, merge/rebase with tests. No GitHub Actions.

Nonnegotiable functionality:
- Preserve v1.3.0 `.m` and `.pq` import, deterministic filename prefix ordering, manually editable order/names/paths/enabled state, and replace/skip behavior. Preserve original M bytes/content unless explicit user editing or parameter changes occur.
- Repair unsafe query deletion/replacement, stage changes without risking source workbook data, robust backup/restore and read-only/macro/security checks; do not edit non-query Excel objects implicitly.
- Inspect existing workbook queries + loads and compare source-to-workbook M code; provide clear new/modified/unchanged/conflict/unknown status; export queries and save/load versioned import profiles.
- Create a maintainable `QAbstractTableModel`/`QTableView` manager-style PySide6 UI including drag/drop, sorting/filtering, selected-row and bulk edits, query inspector, M diff/editor, destination configuration and operation report; responsive at standard laptop resolutions, keyboard navigable and accessible.
- Implement supported loads: definition-only, connection-only (if genuinely distinct/available), new worksheet, existing worksheet at starting cell, Excel Data Model, and preserve existing load. Probe capabilities by Office version/build; unsupported variants must be disabled with explanation. Never conflate definition-only, connection-only and loaded modes.
- Analyze query references conservatively (M language tokenizer or explicit ambiguity status), show dependency issues, preserve user order unless approved, detect cycles; do not falsely claim full M-language parsing.
- Refresh selected/all queries with controllable sync/async behavior, timeout and diagnostics; never run external refresh just to validate M without explicit consent.
- v2: multi-workbook dry-run/deployment, source-path profile mapping with preview, optional parameter editing, import/export packages, version/history, optional CLI with dry-run, dependency view.
- Secure-by-default: no exfiltration/telemetry, secrets redacted in logs, input bundles treated untrusted, macro setting guarded, no shell injection/SQL string concatenation hazards, no silent destructive overwrites, no admin required.

Quality expectations:
- Use dependency-free core models and an injectable Excel adapter for fake-COM unit tests; on real Windows, Excel automation must execute on a COM-initialized worker thread and guarantee cleanup of **its own** Excel process only.
- For each phase, write failing test(s) first for known defects, add regression/edge/UI tests and log failures meaningfully; run static checks and local tests. Follow `docs/QA_WINDOWS_MATRIX.md` on disposable `.xlsx` / `.xlsm` with installed desktop Excel for true acceptance.
- Any COM loading/refresh/data-model approach must be verified on real Excel and cited in `docs/EXCEL_COM_SPIKES.md` with precise observed version/build, method and limitations; sample code from Microsoft docs is a starting point, not a certification.
- Build and launch a one-folder PyInstaller EXE, check process exit and Qt components, import on disposable Excel workbook, absence of orphaned processes, and repeatable package on supported Windows. Preserve one-file EXE as optional only if it passes real smoke tests.
- Use concise, actionable errors; never report a staged workbook as 'successfully installed' if final verification/commit failed.

Deliverables:
1. Updated `app/` source, migrations (if applicable), tests, fixtures, build scripts and user documentation for releases advanced.
2. Per-release `evidence/releases/vX.Y.Z.md` containing features, git commit(s), test commands/outcomes, Windows Excel matrix results, known limitations and release gate result.
3. Compact `dist` package with only necessary runtime files, user manual/summary, and `SHA256SUMS`; do not bundle virtual environments, caches, fixture workbooks, `.git`, source, raw logs or redundant binaries.
4. Final report listing completed vs incomplete release features and explaining any Windows-dependent checks impossible in the current environment.

Release sequence: complete v1.3.1 safety before v1.4.0; v1.4.0 before v1.5.0 load functionality; v1.5.0 before v2.0.0. If verification cannot be done on a real Windows Excel host, still implement and test in portable simulation but mark Excel-specific features **unverified**, not production-ready. Never force release gates to green.

Start by exploring `app/` and executing its existing tests now, then proceed into implementation.
