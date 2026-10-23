# Explicit decision log to resolve by implementation / Excel spikes

Do not block all work on these questions; prototype default, log evidence and choose safe behavior. Escalate to user only for material tradeoffs.

| Decision | Initial safe default | Decision authority / gate |
|---|---|---|
| Connection only vs definition only | distinct option disabled until real Excel confirms | Agent A + QA; v1.5 spike |
| Data Model support across Excel versions | capability detection; unsupported message | Agent A + Windows QA |
| Existing load mapping where several connections/tables match | UNKNOWN; disallow destructive replacement | Agent A |
| Source file encoding beyond UTF-8 BOM | UTF-8-sig; actionable error; no guessed fallback | Orchestrator |
| Workbook may contain external links, macros and sensitive sources | no auto-refresh, force-disable programmatic macros, no external link updates | Security reviewer |
| Query editor edits original file or imported payload | in-memory draft; separate Save to File and Apply actions | Agent B |
| Backup retention policy | keep backups; opt-in cleanup | Orchestrator |
| Staging in OneDrive/network share | abort if no proven safe strategy; offer Save As new output | Agent A + QA |
| Query package format versioning | JSON manifest schema v1 + .pq/.m, sha256 | Orchestrator |
| Dependency parser completeness | conservative scanner + UNKNOWN; no false complete graph | Agent A |
| Real tests inaccessible | deliver code + unit tests but gate as unverified | Orchestrator |
| Release naming | keep launcher `PowerQueryInstaller.exe` during transition; UI display `Power Query Manager` | Orchestrator |
