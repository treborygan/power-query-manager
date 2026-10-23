# Documentation

This index groups the project documentation by task. For the application overview and packaged quick start, begin at the [repository README](../README.md). The [application guide](../app/README.md) covers detailed behavior, requirements, development, and packaging.

## Understand the project

- [Roadmap](ROADMAP.md) — release sequence and planned scope.
- [Architecture](ARCHITECTURE.md) — application boundaries and component responsibilities.
- [Acceptance criteria](ACCEPTANCE_CRITERIA.md) — behavior required for the supported workflows.
- [Feature traceability](FEATURE_TRACEABILITY.md) — map requirements to implementation and evidence.
- [Changelog](CHANGELOG.md) — recorded implementation changes.
- [Open design decisions](OPEN_DESIGN_DECISIONS.md) — decisions that remain open or need future review.

## Work safely with Excel

- [Security and failure policy](SECURITY_AND_FAILURE_POLICY.md) — staging, backups, safe failure, and data boundaries.
- [Decision flows](DECISION_FLOWS.md) — user-visible operation and failure paths.
- [Windows acceptance matrix](QA_WINDOWS_MATRIX.md) — disposable-workbook procedures and remaining manual gates.
- [Excel COM observations](EXCEL_COM_SPIKES.md) — observed Office versions, methods, and limitations.
- [Release operations](RELEASE_OPERATIONS.md) — packaging and release checks.

## Review implementation evidence

- [Final implementation report](../evidence/FINAL_REPORT.md) — delivered scope, measured checks, and current limitations.
- [Milestone release records](../evidence/releases/) — v1.3.1, v1.4.0, v1.5.0, and v2.0.0 evidence and gate status.
- [Baseline](../evidence/baseline.md) — measured baseline environment and result.
- [Contracts](../contracts/) — schemas for profiles and deployment plans.

Excel-dependent behavior is supported only to the extent documented by the observed Windows/Office evidence. A successful portable test or package build does not certify every Excel version, workbook type, connector, or security policy.
