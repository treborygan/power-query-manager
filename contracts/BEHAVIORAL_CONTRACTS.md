# Core behavioral contracts (agreed before parallel changes)

## Input discovery

Each standalone `.m` and `.pq` file yields exactly one source with full M expression, preserving Unicode/newlines/encoding semantics. Do not assume extension implies distinct syntax. Numeric prefix accepted if file stem begins with digits followed by `_`, `-`, whitespace; strip only that prefix from initial query name. Identical prefix numbers cause deterministic tie warning. Auto-assigned unprefixed orders remain deterministic when later batches imported. User manual order overrides automation. ZIP path traversal/symlink/bomb defense required before enabling archives.

## Plan identity

Stable UUID or deterministic derived query-source ID must not depend on Python object `id()` or view row index; every edit and view sort selects by ID. Plan generated from an immutable snapshot containing formula SHA-256, workbook hash, options, destination and action. Revalidate file hash/lock before commit. Snapshot and source files must not mutate under running job.

## Diff and statuses

`NEW`: no match. `MODIFIED`: differing formulas and no unhandled collision. `UNCHANGED`: matching exact formula content/hash. `SKIP`: existing kept by explicit policy. `CONFLICT`: ambiguous match, load collision or incompatible action requiring resolution. `UNKNOWN`: can inspect query but load/dependency state uncertain. For comparison, normalized whitespace view may aid users but must never suppress real changes without explicit policy; use exact content/hash for safe skip decisions.

## Excel outputs and idempotency

A destination (sheet/table/model) is not considered installed until inspected after save/reopen. Existing connections and tables are not duplicated on reruns. For preservation of unknown load linkage, stop instead of destructive recreate. Changes to original only after staging verification + hash recheck; conflict creates recoverable staged artifact with truthful error. Refresh may execute external M and needs explicit authorization.

## Validation

Sheet names distinct from query names; sheet <=31 chars, Excel sheet invalid `[ ] : * ? / \\`, nonempty, not conflicting case-insensitively. Cells A1 addresses within Excel grid (`XFD1048576` bounds), not merely matching regex. Table names comply with Excel rules and must not collide with names or cell references. Profiles schema v1 is stable; future additions require schema migrations and tests. Plan schema represents intents, not proof that COM can perform an operation.

## Capability and uncertainty

Distinguish feature requested, technically implemented, Windows Excel supported and tested, packaged verified. If backend cannot determine mapping/capability, UI shows UNKNOWN/unavailable; no silent fallback to a mode with different semantics.

## Query package manifests

- `query-package.schema.json` accepts manifest versions 1 and 2. New exports use v2; v1 imports derive v2 dependency metadata in memory and never rewrite the input archive.
- Query order and formula UTF-8 bytes are preserved; the manifest hash is SHA-256 of each exact formula byte sequence. Extraction validates safe paths, count/size/ratio limits, duplicate names, symlinks, manifest shape, hashes, and unlisted source files before sources reach the UI.
- V2 dependencies are advisory only. The bounded M scanner records edges only when analysis is unambiguous; otherwise the row is marked `ambiguous` with no asserted edges. Cycle paths are derived from known edges. Import rechecks dependency metadata against the formula bytes and does not reorder query rows automatically.
- Imported package rows enter the editable source table for review. Workbook compare, diff, and the existing explicit replace/skip choice remain the only path to workbook mutation.
