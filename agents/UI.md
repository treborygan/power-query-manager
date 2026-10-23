# Agent B: PySide6 UX

Own `app/src/pq_installer/ui/` and UI tests. Inspect original `main_window.py`/`install_worker.py`; retain all legacy features. Convert `QTableWidget` data flow into `QAbstractTableModel` plus proxy. The domain immutable plan is supplied by orchestrator/engine; use it without duplicate business rules.

Implement manager navigation, workbook inspector, source grid, validation badges, load destination dialog, M code/diff viewer, profiles, drag/drop and bulk edits, keyboard shortcuts, dark/light/system, progress logs and safe cancellation. Distinguish view sort from import sequence. Freeze all mutating widgets during jobs. For sensitive M expressions, do not print entire code in operation logs by default. Test via Qt offscreen where possible; evaluate packaged GUI for readability on Windows at 1366×768 and 125% scaling. Don't add buttons with no functional action.

Report: changed screens/files, UI tests/screenshots, any platform-specific bugs, interface needs and exact next task IDs.
