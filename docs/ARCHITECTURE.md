# Target architecture and interface boundaries

Follow this as the starting architecture; fine-tune based on actual tests, but never couple Excel COM directly to UI table widgets.

```text
UI / CLI                 Application services           Pure domain / adapters
┌───────────────────┐    ┌───────────────────────┐       ┌───────────────────────────────┐
│ MainWindow        │───>│ Catalog/PlanBuilder    │──────>│ SourceDiscovery/Reader        │
│ QueryTableModel   │    │ Inspector/Comparator  │──────>│ MReferenceScanner (bounded)   │
│ Inspector+MEditor │    │ Profile/Package svc   │──────>│ ProfileSchema/Validation      │
│ DestinationEditor │    │ Import/Export/Refresh │──────>│ ExcelAdapter protocol         │
│ Progress+Logs     │    └───────────────────────┘       │   COM adapter / fake adapter  │
└──────┬────────────┘                                     └───────────────────────────────┘
       │ sealed immutable JobPlan                                       │
       └─────────> Worker/QThread + COM apartment ────────> Workbook staging/transaction
```

## Contract objects (versioned)

- `QuerySource`: stable ID (not `id()`), path, display/query name, enabled, user order, prefix metadata, source hash, encoding; M content loaded exactly once at plan generation.
- `QueryPlanItem`: source ID/hash/name/action/destination/dependencies and preflight checks; immutable; source contents included only in protected in-memory plan and not logged by default.
- `WorkbookInventory`: workbook metadata; existing WorkbookQuery entries, known/unknown connection(s), linked sheets/tables/model and capabilities; snapshot before changes, read-only scan.
- `LoadDestination`: explicit discriminated mode (`definition_only`, `connection_only`, `new_sheet`, `existing_sheet`, `data_model`, `preserve_existing`) and sheet/table/cell properties; cannot instantiate invalid combinations.
- `InstallPlan`: versioned immutable snapshot with workbook fingerprint, selected actions, conflict resolution, refresh policy, expected diffs and backup policy. Preflight must re-check workbook fingerprint immediately before applying plan (TOCTOU guard).
- `OperationResult`: per query and per stage, planned vs actual action/status, started/ended timestamps, errors, fallback reasons, verifications, backup/staging/output paths; persist redacted summary.

JSON shapes in `contracts/profile.schema.json` and `contracts/plan.schema.json` are external serialized forms. Code may have richer Python types, but serialization/migrations must be tested and explicit. The `schema_version` is independent of application version.

## COM adapter and threads

The current injectable application boundary is `ExcelApplicationAdapter` in `core/excel_adapter.py`: `create_application()` is called on the operation worker after its COM apartment is initialized, and `requires_com_initialization` determines whether that apartment is owned by the service. The production adapter always uses `DispatchEx`; `FactoryExcelAdapter` preserves the legacy `excel_factory` test seam, and explicit fake adapters can model creation/COM policy. The inspector, importer, refresher and capability probe share this boundary. `ExcelWorkbookOperations` centralizes workbook add/open/save/save-as/close calls behind a default COM adapter and an injectable fake port. Import and refresh also share `WorkbookTransactionAdapter` in `core/excel_transaction.py` for verified sibling staging, pre-commit fingerprint/active-user checks, atomic replace, and cleanup. Query discovery/mutation/load/refresh logic remains in its domain services; the larger workflow protocol (`inspect_workbook`, `preflight`, `apply_definitions`, `configure_load`, `refresh`, `verify`, `close`) and blocked-COM lifecycle matrix remain future work. Keep COM proxies inside the initialized apartment/work thread. UI workers establish/clean their own COM instances; no COM proxies pass through Qt signals, only plain dataclasses/JSON.

### Transaction + staged workbook approach

1. Validate all sources, constraints and permissions without creating side effects. Hash original workbook, record metadata, ensure it is not already open/locked. Snapshot existing loads and plans; user confirms.
2. Create unique immutable backup and verify bytes/hash. Create a **working copy** with same extension and supported metadata; track original → working → final safely (especially OneDrive-synced/network paths). Assess whether a local staging copy can preserve external data location semantics before enabling.
3. Open the working copy in an isolated Excel COM process with explicit restricted macro/links policy. Abort if Excel reports read-only. Keep one workbook handle; do not involve user's other Excel process.
4. Apply all query definitions, then configure requested loads only when capability supported and destination preflight green. Avoid changing tables in situ without documenting and verifying preservation rules.
5. Optional user-authorized refresh; observe completion. Save and close cleanly. Reopen working workbook for verification: expected M formulas, connections/destinations/table names and workbook readability.
6. Commit the staged file to requested output/target **only with explicit strategy**. On same-device rename/replace may be atomic but cloud sync and network shares are different. Pre-commit recheck original fingerprint; if changed, preserve working copy and report conflict; NEVER force overwrite.
7. Preserve backup + logs; success only after commit and final validation. On failure quarantine stage and original unchanged, or clearly report post-commit recovery outcome. No fake 'all rolled back' claims.

If direct in-place editing is retained as fallback, it must require explicit opt-in and have stronger proven recovery/backup tests. It cannot be default.

## Load-mode semantics

- **Definition only**: workbook query formula exists; no explicit request to create connection/table/model.
- **Connection only**: explicit connection object without worksheet/model load **if the Excel build exposes a reproducible distinct state**. Otherwise mark unavailable and point user to definition only.
- **New sheet / Existing sheet**: creates or reuses appropriate Power Query-backed table connection and verified `ListObject` at chosen location; never overwrite unrelated cells/tables.
- **Data Model**: create/maintain actual in-model connection/table and verify through Excel's `Workbook.Model`, not just a named query.
- **Preserve existing**: only for existing query with known load objects; maintain formula/connection/worksheet table/model links as currently configured. If ambiguous mapping, prevent automated replacement or ask targeted confirmation.

Ensure all load modes are idempotent and have detection semantics (already loaded, conflict, unsupported, unknown). Do not call query-definition-only a connection-only load.

## Excel version feasibility

First prototype against throwaway workbooks and capture `Application.Version`, `Application.Build`, architecture/bitness and Excel edition (only if accessible without private identifiers). Provider strings / connection APIs may vary; test Unicode/spaces/punctuation query names, external data privacy settings and Power Query firewall errors. Implement capability detection and graceful unavailable UI.

## M dependency analyzer

M is a complete language: avoid naive regex substitution or false certainty. Implement conservative lexical scanning with strings (escaped double quotes), `//` and `/* */` comments, nested comments where applicable, quoted identifiers `#"..."`, records, `let` locals, and case-sensitive names consistent with M. Match only known external query names in expression contexts. Expose dependencies with confidence and blockers for unresolved/ambiguous references; do not reorder automatically without consent. Detect cycles in confident graph and stable topological sorts preserving user's tie order.

## Portability/performance

Only PySide6/pywin32 and deliberate small dependencies. Keep non-COM validation/pattern parsing available offline and deterministic. Avoid heavyweight runtime/runtime admin requirements. No network access needed for ordinary installed query definitions unless they themselves refresh from external sources with explicit user action. Build on Windows with onedir PyInstaller first; optional onefile only after parity.
