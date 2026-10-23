# Agent A: Excel integration / core

Own `app/src/pq_installer/core/` and compatible engine tests. Read original code in `core/excel_queries.py`, `core/models.py`, `core/backup.py`, source discovery and validation. First write a failing regression for fallback `existing.Delete(); Queries.Add()` when `Formula` unreadable; fix by refusing unsafe mutation. Handle secure Workbooks.Open, COM apartment, ReadOnly, unique validated backups, immutable plans, source hashing, staging, transaction verification. No GUI direct COM access.

After v1.3.1, implement read-only workbook inventory/export, classifier and schema loader; then prototype load modes with real Excel and document `docs/EXCEL_COM_SPIKES.md`. Preserve original connection/tables/model loads when replacing query formulas; when unknown, block. Write fake-COM tests for all failure injection points and refresh statuses; mark any real Excel feature unverified until QA supplies evidence. Use conservative M scanner. Keep interface frozen or RFC proposed to orchestrator before edit.

Report: changed file inventory, Python unit results, manual Excel spike outcome/Office versions, capability grid, unsupported cases, risks, next task ID.
