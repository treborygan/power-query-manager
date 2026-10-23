# Power Query Manager — Codex implementation handoff

**Source baseline:** uploaded `Power_Query_Installer_1.3.0_Source_Build.zip` (2026-10-21), extracted **without changes** under `app/`.  
**Implementation goal:** evolve the v1.3.0 PySide6/Excel-COM Windows program through v1.3.1, v1.4.0, v1.5.0, and v2.0.0.  
**Status of this archive:** implementation **handoff only**, not a completed upgrade and **not** a verified Windows executable.  
**First action:** read `CODEX_KICKOFF_PROMPT.md` and `AGENTS.md`, then inspect the code in `app/`.

## Quick start for Codex

1. Unpack into a writable repository/worktree. Keep `app/` as the project working directory and retain `docs/`, `contracts/`, `agents/`, `examples/`, and `tools/` alongside it.
2. Paste the **entire** contents of `CODEX_KICKOFF_PROMPT.md` into Codex. The orchestrator must read `AGENTS.md` and the specifications before writing application code.
3. Check pristine baseline: `python tools/verify_handoff.py`; from `app/` run `PYTHONPATH=src python -m pytest -q` (Linux/macOS) or `..\tools\run_tests.ps1` (Windows).
4. Start in v1.3.1; deliver each phase with evidence and stop/go signoff. Do not declare releases complete based only on fake COM tests.
5. On Windows with desktop Excel, run the manual integration matrix in `docs/QA_WINDOWS_MATRIX.md` **on disposable fixtures only** and build/test the EXE via `tools/build_and_smoke.ps1`.
6. At final delivery, keep the **release** distribution minimal: EXE plus required `_internal` files (for onedir), a short README/summary and any required licenses. Do not ship source/tests/cache as release files. Preserve full source and evidence in the development repository.

## Directory map

| Path | Purpose |
|---|---|
| `app/` | Real, unchanged v1.3.0 application code, tests, and original build settings — Codex edits here |
| `CODEX_KICKOFF_PROMPT.md` | Single paste-ready implementation prompt |
| `AGENTS.md` | Mandatory multi-agent execution rules, ownership, nonnegotiables |
| `agents/` | Four agent role briefs and reporting templates |
| `docs/` | source audit, functional specification, roadmap, architecture, decision flowcharts, UX, testing, deployment, risks and acceptance matrix |
| `contracts/` | JSON schemas, release gate definitions, behavior contracts |
| `examples/` | Non-sensitive demonstration `.m` / `.pq` sources and profile |
| `tools/` | Handoff and JSON-schema validation, defect reproduction, environment/test helpers and Windows smoke procedure |
| `evidence/` | Templates for run logs and release status; no fabricated results |
| `BASELINE_SHA256SUMS.txt` | SHA-256 hashes for unmodified baseline source files |

## Important scope boundaries

- Keep **Python + PySide6 + Windows Excel COM**, not a browser rewrite or Excel JS replacement.
- No admin rights must be necessary for normal launch. Enterprise endpoint restrictions may block unsigned executables; report accurately; do not bypass corporate policy.
- **No GitHub Actions** or other CI automation unless the user explicitly approves it. Local build/test scripts are allowed.
- Do not include user's workbooks, credentials, SharePoint tokens, tenant details, personal paths, or network connections in fixtures or telemetry.
- Do not claim the product can reliably load Data Model connections, preserve complex pre-existing loads, or parse arbitrary M code until verified on actual supported Excel installations.
- Unsupported feature/capability must be visible and actionable, never silently substituted.

## Baseline observations

From the supplied source: supports multi-file `.m` and `.pq`, numeric-prefix ordering, editable rows (enabled, order, name, path), `Workbook.Queries` definition create/update, skip existing, timestamped backup and simple rollback. One load mode only: `QUERY_ONLY`, without worksheet/Data Model operations. Current UI is a single `QTableWidget` with hardcoded stylesheet and a `QThread` worker. Source includes 18 pytest tests (18 passed in the handoff environment), but no actual Windows/Excel acceptance evidence. A high-risk replacement fallback deletes an existing query before successful replacement/recovery when old `Formula` is inaccessible. Read `docs/BASELINE_AUDIT.md` for exact details.

## Definition of complete

All critical criteria in `docs/ACCEPTANCE_CRITERIA.md` and phase gates in `contracts/release_gates.json` must pass, with reproducible logs/screenshots, verified Windows EXE startup/COM behavior and backup integrity. Explicitly list unmet items; do not present an unverified build as working.
