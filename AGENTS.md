# Codex agent coordination contract

**All agents must read this file.** This is a code implementation task, not a planning-only task. Preserve baseline user project and do not fake validation.

## Roles and ownership

**Orchestrator / integration lead**: final owner of `contracts/`, public Python interfaces, architecture, merges, release versions, changelog, gating and cross-agent issues. Assign single ownership of each changed file. Maintains `evidence/STATUS.md` and `evidence/DECISIONS.md` with short timestamped entries; decides unresolved conflicts; runs full local tests after merges. Never merge without review.

**Agent A — Excel COM & domain engine**: owns `app/src/pq_installer/core/`, Excel adapter implementation, inventory, safe edit strategy, load/refresh/rollback and model contracts once orchestrator approves API. Adds isolated unit/fake-COM tests under `app/tests/core/` or uniquely named engine tests. Does not edit UI files or packaging configs. Crucial: COM worker thread, no Excel live interaction on QA-only/Linux host; document unsupported capabilities.

**Agent B — UI/UX and profiles**: owns `app/src/pq_installer/ui/`, QAbstractTableModel, preview/diff UI, styles, validation indicators, accessibility and integration with orchestrator-approved service API. Adds GUI tests under `app/tests/ui/`. May implement profile view layer but not core schema/parsing. Does not directly issue COM calls from the GUI thread.

**Agent C — QA/security/release**: owns `app/tests/integration/`, `app/tests/fixtures/`, release scripts/spec, smoke harness, performance/release evidence and security review. May add tests to other namespaces only by coordination. No GitHub Actions. Ensures package contains only needed files. Never mark Windows test PASS without actually running Excel installed locally.

## Branch/parallel execution rules

1. Orchestrator freezes dataclass/API/schema contracts before parallel changes. Prefer separate worktrees or branches (`agent/core`, `agent/ui`, `agent/qa`), merging serially after review.
2. No shared mutable file ownership, no simultaneous edits to `pyproject.toml`, `app/src/pq_installer/core/models.py`, or release docs. Orchestrator merges these on request.
3. Each agent reports: changed files, reasons, tests run and exact output, assumptions, blockers, interfaces impacted and next steps. Record material blockers and decisions.
4. Agents can propose contract changes via an RFC in `evidence/DECISIONS.md`, but cannot silently change public contracts. Document version and migration decisions.
5. When tests fail: isolate cause, add reproducer, fix within ownership; do not delete tests, lower assertions, skip real acceptance, or add placeholders as finished work.
6. Before each merge, run relevant unit suite, then orchestrator runs full suite with COM fakes. After each release run the manual Windows matrix if a Windows Excel host is present.
7. If tools do not support subagents, the orchestrator executes sequentially. Do not spawn nominal agents without real isolation.

## Cross-cutting invariants

- M code may contain credentials or private URLs; prevent leaking formulas in logs unless user intentionally opens the editor; structured logs redact secrets and paths where appropriate.
- Query load/inventory/status mapping must never silently guess: `unknown/unsupported` is a valid state.
- Dry-run MUST NOT write to original, open Excel with external link updates, refresh data, or make unexpected workbook changes.
- GUI edits disabled during execution; worker receives a frozen immutable plan; unsafe cancellation not allowed mid-commit.
- Preserve references to existing tables/loads when updating M. No destructive recreate fallback if restoring the original formula/loads is not proven.
- No corporate policy bypass; do not include real customer/workbook/tenant data in test fixtures.
- No GitHub Actions; standalone/local reproducible PowerShell + pytest preferred.

## Phase releases

### v1.3.1 (required gate)
Fix unsafe fallback, read-only checks, security/macro settings, protected backup strategy, error/log consistency, immutable worker plan, packaged startup. Baseline behaviors unchanged. See `docs/ROADMAP.md`.

### v1.4.0
Read-only workbook inventory, changed status/diff, export, profiles, JSON schema migrations, modern UI. Write methods not allowed in inspect mode.

### v1.5.0
Worksheet/Data Model loads + preservation, dependency warnings/sorting opt-in, refresh. Every load mode gated by real Excel acceptance.

### v2.0.0
Multi-workbook deployment, path mapping/parameters, query packages, history and optional CLI; keep automated tests off workbooks containing real data.

### Definition of DONE
Feature-specific acceptance in `docs/ACCEPTANCE_CRITERIA.md`; reproducible evidence saved to repo; no silent destructive modifications; release package minimal and functioning on target Windows desktop Excel. If environmental tests unavailable: explicit NOT VERIFIED gate.
