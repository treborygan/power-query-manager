# UI specification and interaction contract

The app should remain a native, resizable PySide6 desktop tool, designed for 1366×768 laptop and larger monitors (support high-DPI/scaling and keyboard navigation).

## Page structure

1. Left navigation: **Workbook**, **Queries**, **Review**, **History**, **Settings**. Selected page persists per session; no page navigation allowed that would mutate a running job.
2. Header: current workbook name/path (truncated with tooltip), status badge, read-only/lock/Excel capability indicators, **Inspect** action. Avoid showing company paths in exported screenshots/logs by default.
3. Main **Query grid**: enabled, order, name, source path, source type, M hash/change badge, destination mode, target sheet/table/cell, dependency warnings, action. `QAbstractTableModel` + sort/filter proxy. Row sort **must not** silently alter configured import order; order column controls execution.
4. Side **Inspector**: current Excel formula vs imported source, line diff, metadata, dependencies, inferred connection/load, capabilities, safety warnings. Editor has syntax highlighting and Unsaved/Save/Apply Explicit confirmation.
5. Bottom operation/progress panel: counts by status, current phase, structured per-item logs; progress determinate when step counts known; cancel only at safe pre-commit points.
6. Review screen: summary of additions/updates/skips/conflicts, load destinations, dependencies, backup path, refresh side effects, commit mode and explicit confirmation. Batch apply disabled until preflight passes.

## Required interactions

- Drag/drop `.m`, `.pq` or `.zip` onto source grid; Add files/folder/recursive folder. Explain why unsupported sources skipped.
- Bulk edit chosen query load modes, replace/skip policy, order (preserve relative ordering), query name mapping, and enable/disable; multi-select never mutates non-selected rows.
- Changing a query name updates the current plan only, not its source file unless Export/Save explicitly invoked.
- Editing M shows a dirty marker; closing an edited inspector presents Save/Discard/Cancel choices. Ctrl+Z for grid edits and code editing.
- Ability to filter/search queries and sort grid visually without changing the import plan. Drag row to reorder only if specifically in 'Reorder' mode, then renumber deterministically.
- Any unresolved conflict has clear explanation + 'keep existing', 'replace definition only', 'compare', or 'cancel' choices. Show destructive side effects by object.
- Dark/light/system mode, icons, validation tooltips, accessible names, sensible focus order/keyboard shortcuts and consistent error colors without relying on color alone.
- For load destinations, sheet picker populated from inspected workbook; new sheet names validated (Excel 31-char and forbidden characters) and unique; cell editor validates A1 absolute/relative mapping; table name validated separately.

## Empty/loading/error states

No workbook: show workbook picker and a brief explanation. Workbook inspection unavailable: show actionable Excel not installed/COM access denied/read-only/protected view error. Empty import: show Add files. Unsupported target modes: disable controls **and** state reason. Import failed: show phase, object, backup/staging file locations and a 'View details' action.

## UI QA

Qt offscreen unit tests cover editable model validations, column states, sorting is view-only, table row identity survives filtering, async busy lock, close/cancel behavior, and time-saving keyboard flows. Real desktop screenshots for 1366x768, 1920x1080, 125%/150% scaling, light/dark themes, popup positioning, overflow, high contrast. Accessibility checklist in `docs/QA_WINDOWS_MATRIX.md`.
