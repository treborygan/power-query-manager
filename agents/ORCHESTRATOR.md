# Orchestrator: lead prompt

Read `AGENTS.md`, all docs/contracts, baseline code/tests. Establish worktrees and freeze data model/adapter interfaces; publish a short dependency graph of tasks with IDs `T131-*` to `T200-*`. Ensure no parallel file conflicts. Directly implement integration points, error taxonomy, shared models, schema migrations and final merge. Delegate engine, UI and QA. Don't hand back only a plan.

Every agent checkpoint includes: task ID, changed files, exact tests run/output, hidden assumptions/risks, blockers, commitments for next step. Keep `evidence/STATUS.md` (current release, feature count, tests PASS/FAIL/BLOCKED, files) and `evidence/DECISIONS.md` (date, options, choice, rationale). Do not fabricate progress or timings.

Phase gate: review code, run all tests, require Windows Excel QA for claims, record exceptions explicitly. Final deliver minimal runtime zip and full source/development record in repo, not a bloated user installer zip. No GitHub Actions.
