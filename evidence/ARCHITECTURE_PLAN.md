# Architecture and ownership plan

Execution is sequential because this extracted handoff has no isolated agent worktrees available. The orchestrator owns contracts, public interfaces, architecture, integration and gates.

| Release | Tasks | Acceptance | Main risks |
|---|---|---|---|
| v1.3.1 | Add immutable source snapshots, conservative backup, preflight file safety, non-destructive query replacement, COM security/open options and clear errors; regression tests first. | S-01..S-06, LEG-01..05 via portable/fake COM. S-07 requires Windows build/Excel. | Excel COM object behavior and atomic commit cannot be certified without Excel. |
| v1.4.0 | Pure inventory/comparison/profile/export/diff services and manager table model/UI. | M-01..M-07; PySide UI tests when available. | GUI toolkit/package and Excel inventory mappings. |
| v1.5.0 | Capability-gated load modes, conservative dependencies, refresh controls. | L-01..L-09. | All load/refresh behavior requires real Office evidence; disabled until certified. |
| v2.0.0 | Safe packages, multi-workbook dry-run, mapping/parameter preview, history/CLI. | A-01..A-07. | Broad scope and installed Excel/package acceptance. |

Contracts: preserve supplied `contracts/profile.schema.json`, `contracts/plan.schema.json`, and `docs/ARCHITECTURE.md` as frozen v1 contract definitions during the initial implementation. Core remains dependency-free and COM-injected; UI communicates through immutable plans/services. Automated tests use disposable temp files and fake adapters only. Windows workbook matrix is never inferred from unit results.
