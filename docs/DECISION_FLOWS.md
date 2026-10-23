# Decision flows for agents and importer

These are operational flowcharts. Do not use them as substitutes for source code tests.

## Agent issue / implementation flow

```mermaid
flowchart TD
  A[Pick next release-gated task] --> B[Read source, contract, tests and ownership]
  B --> C{API or ownership conflict?}
  C -- Yes --> D[Write RFC and request orchestrator decision]
  D --> B
  C -- No --> E[Create failing regression or acceptance test]
  E --> F[Implement minimal safe change]
  F --> G[Run local tests and review diff]
  G --> H{Tests pass?}
  H -- No --> I[Log repro, isolate cause and fix]
  I --> F
  H -- Yes --> J{Requires real Excel or packaged GUI?}
  J -- Yes --> K[Run Windows disposable acceptance]
  J -- No --> L[Submit evidence and changes for review]
  K --> M{Windows acceptance available and passing?}
  M -- No --> N[Mark BLOCKED or UNVERIFIED, disable unsupported mode]
  N --> L
  M -- Yes --> L
  L --> O[Orchestrator merges and runs full tests]
  O --> P{Release gate satisfied?}
  P -- No --> A
  P -- Yes --> Q[Tag/report release and advance milestone]
```

## Workbook operation safety flow

```mermaid
flowchart TD
  A[User selects workbook and query sources] --> B[Validate input, profile and source text]
  B --> C{Valid and trusted scope?}
  C -- No --> X[Show actionable validation error, no changes]
  C -- Yes --> D[Inspect workbook in isolated Excel process]
  D --> E{Read-only, protected, in use or unsupported?}
  E -- Yes --> X
  E -- No --> F[Compute exact actions, destinations, conflicts and file hashes]
  F --> G{Unknown/destructive conflict?}
  G -- Yes --> H[Require explicit resolution or cancel]
  H --> F
  G -- No --> I[Present dry-run and user confirmation]
  I --> J{Confirmed actual changes?}
  J -- No --> Y[Exit without changes]
  J -- Yes --> K[Create collision-safe verified backup and staged copy]
  K --> L{Backup/stage valid?}
  L -- No --> Z[Abort, original unchanged, preserve evidence]
  L -- Yes --> M[Re-check hashes, locks and permissions]
  M --> N{Target changed since preflight?}
  N -- Yes --> Z
  N -- No --> O[Apply query definitions to staged copy]
  O --> P[Configure only verified destinations]
  P --> Q{Refresh explicitly authorized?}
  Q -- Yes --> R[Refresh and observe completion or timeout]
  Q -- No --> S[Save, close and reopen staged workbook]
  R --> S
  S --> T{Queries, loads and workbook reopen verified?}
  T -- No --> Z
  T -- Yes --> U[Recheck original fingerprint and commit method]
  U --> V{Safe final commit possible?}
  V -- No --> Z
  V -- Yes --> W[Commit and verify, report success with backup]
```

**Important:** `Z` is not 'rollback succeeded' unless confirmed. If failure occurs after partial commit, record the exact persisted state and do not overwrite a possibly independently modified workbook. Cases that may evaluate M (load to sheet/model or refresh) should visibly disclose possible file/network/database access before execution.
