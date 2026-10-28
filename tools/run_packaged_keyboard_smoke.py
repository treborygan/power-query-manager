"""Verify keyboard focus traversal through the packaged manager and optional dialog.

Requires the optional Windows UI smoke dependencies in
``tools/requirements_ui_smoke.txt``. This does not open Excel or a workbook.
"""
from __future__ import annotations

import argparse
import csv
import io
import os
import subprocess
import sys
import tempfile
import time
from pathlib import Path


def _excel_pids() -> set[int]:
    result = subprocess.run(
        ["tasklist.exe", "/FI", "IMAGENAME eq EXCEL.EXE", "/FO", "CSV", "/NH"],
        check=True,
        capture_output=True,
        text=True,
        encoding="mbcs",
        errors="replace",
    )
    pids: set[int] = set()
    for row in csv.reader(io.StringIO(result.stdout)):
        if len(row) >= 2 and row[0].casefold() == "excel.exe" and row[1].isdigit():
            pids.add(int(row[1]))
    return pids


def _tab_cycle(window, automation, process_id: int, label: str, limit: int = 100) -> list[str]:
    """Return accessible names visited in one Tab cycle, failing on focus escape."""
    focused = automation.GetFocusedElement()
    if focused.CurrentProcessId != process_id:
        raise AssertionError(f"Initial {label} keyboard focus is outside the packaged manager.")
    first_runtime_id = tuple(focused.GetRuntimeId())
    visited = [focused.CurrentName or ""]
    for _ in range(limit):
        window.type_keys("{TAB}", set_foreground=True)
        time.sleep(0.06)
        focused = automation.GetFocusedElement()
        if focused.CurrentProcessId != process_id:
            raise AssertionError(
                f"Tab moved focus outside the packaged manager during {label} traversal; "
                f"visited={visited!r}, focused_name={focused.CurrentName!r}, "
                f"focused_pid={focused.CurrentProcessId}, manager_pid={process_id}."
            )
        visited.append(focused.CurrentName or "")
        if tuple(focused.GetRuntimeId()) == first_runtime_id and len(visited) > 1:
            return visited[:-1]
    raise AssertionError(
        f"Keyboard focus did not complete a {label} Tab cycle; "
        f"visited={visited!r}"
    )


def _wait_child_window(parent, title: str, timeout: float = 10):
    deadline = time.monotonic() + timeout
    while time.monotonic() < deadline:
        try:
            child = parent.child_window(title=title, control_type="Window")
            child.element_info.rectangle
            return child
        except Exception:
            time.sleep(0.2)
    raise TimeoutError(f"Packaged UI dialog {title!r} did not appear.")


def _accept_file_picker(dialog) -> None:
    for role in ("SplitButton", "Button"):
        try:
            button = dialog.child_window(auto_id="1", control_type=role)
            if button.exists(timeout=0.3):
                button.invoke()
                return
        except Exception:
            continue
    dialog.set_focus()
    dialog.type_keys("{ENTER}", set_foreground=True)


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--exe", type=Path, required=True, help="Packaged manager executable")
    parser.add_argument(
        "--dialogs", action="store_true",
        help="also traverse source-mapping and operation-history modals without using the real user profile",
    )
    args = parser.parse_args()
    executable = args.exe.resolve()
    if not executable.is_file():
        parser.error(f"Packaged executable not found: {executable}")

    try:
        from comtypes.client import CreateObject
        from comtypes.gen.UIAutomationClient import CUIAutomation
        from pywinauto import Application
        import win32con
        import win32gui
    except ImportError as exc:
        raise SystemExit(
            "Install tools/requirements_ui_smoke.txt in the QA environment first."
        ) from exc

    excel_before = _excel_pids()
    isolated_user_data = tempfile.TemporaryDirectory(prefix="pqm-keyboard-user-data-")
    environment = os.environ.copy()
    environment["LOCALAPPDATA"] = isolated_user_data.name
    source_root = Path(isolated_user_data.name) / "sources"
    source_root.mkdir()
    source_file = source_root / "KeyboardSmoke.pq"
    source_file.write_text("let\n    Source = 1\nin\n    Source\n", encoding="utf-8")
    manager = subprocess.Popen([str(executable)], cwd=str(executable.parent), env=environment)
    window = None
    application = None
    try:
        application = Application(backend="uia").connect(process=manager.pid, timeout=20)
        window = application.window(title="Power Query Manager")
        window.wait("visible", timeout=20)
        window.set_focus()
        time.sleep(0.35)
        automation = CreateObject(CUIAutomation)

        visited = _tab_cycle(window, automation, manager.pid, "main-window")

        required = [
            "Target workbook path",
            "Add files",
            "Query source settings table",
            "Existing query behavior",
            "Create a timestamped workbook backup",
            "Review and import selected queries",
        ]
        missing = [name for name in required if name not in visited]
        if missing:
            raise AssertionError(f"Keyboard traversal skipped required controls: {missing!r}")
        positions = [visited.index(name) for name in required]
        if positions != sorted(positions):
            raise AssertionError(f"Unexpected keyboard order for main controls: {positions!r}")

        print("PASS: packaged main-window Tab traversal stayed in the manager and completed a full cycle.")
        print("PASS: workbook path, Add files, query table, existing-query policy, backup and import action were reachable in order.")
        print(f"PASS: {len(visited)} focusable controls visited; no workbook or Excel instance was opened.")

        if args.dialogs:
            # Inspecting a source records a local history entry. Exercise the
            # empty history table before that dialog can add a row to it.
            window.child_window(title="Open operation history", control_type="Button").invoke()
            history = _wait_child_window(window, "Local operation history")
            history.set_focus()
            history_visited = _tab_cycle(history, automation, manager.pid, "operation-history dialog")
            history_required = [
                "Local operation history entries",
                "History retention period in days",
                "Capture optional full workbook snapshots before changes",
                "Refresh history entries",
                "Apply history retention period",
                "Delete selected history entry",
                "Delete all local history",
                "Close operation history",
            ]
            missing_history = [name for name in history_required if name not in history_visited]
            if missing_history:
                raise AssertionError(
                    f"Operation-history dialog keyboard traversal skipped controls: {missing_history!r}; "
                    f"visited={history_visited!r}"
                )
            history.type_keys("{ESC}", set_foreground=True)
            history.wait_not("visible", timeout=5)
            window.wait("visible", timeout=5)
            print(
                "PASS: operation-history dialog Tab cycle reached retention, snapshot, refresh, delete and close controls; "
                "history data stayed in the disposable user profile."
            )

            window.child_window(title="Add files", control_type="Button").invoke()
            picker = _wait_child_window(window, "Select Power Query files")
            filename = picker.child_window(auto_id="1148", control_type="ComboBox").child_window(
                control_type="Edit"
            )
            filename.set_edit_text(str(source_file))
            _accept_file_picker(picker)
            table = window.child_window(title="Query source settings table", control_type="Table")
            deadline = time.monotonic() + 10
            source_row = None
            while time.monotonic() < deadline:
                source_row = next((item for item in table.descendants(control_type="DataItem")
                                   if item.window_text() == "KeyboardSmoke"), None)
                if source_row is not None:
                    break
                time.sleep(0.15)
            if source_row is None:
                raise AssertionError("Temporary KeyboardSmoke source did not appear in the manager grid.")
            source_row.select()
            window.child_window(title="Inspect / compare selected", control_type="Button").invoke()
            inspector = _wait_child_window(window, "Query inspector — KeyboardSmoke")
            inspector.set_focus()
            editor_names = {
                "Editable source M expression",
                "Read-only workbook M expression",
                "Line-by-line workbook and source M diff",
            }
            inspector_visited = _tab_cycle(inspector, automation, manager.pid, "query-inspector dialog")
            inspector_required = [
                "Search M expressions and diff",
                "Find in source",
                "Find in workbook",
                "Find in diff",
                *sorted(editor_names),
                "Save edited M as a separate source copy",
                "Close query inspector",
            ]
            missing_inspector = [name for name in inspector_required if name not in inspector_visited]
            if missing_inspector:
                raise AssertionError(
                    f"Query-inspector keyboard traversal skipped controls: {missing_inspector!r}; "
                    f"visited={inspector_visited!r}"
                )
            inspector.type_keys("{ESC}", set_foreground=True)
            inspector.wait_not("visible", timeout=5)
            window.wait("visible", timeout=5)
            print(
                "PASS: query-inspector keyboard traversal reached search, M/diff editors, save-copy and close controls; "
                "Tab moved through the editor panes and Escape closed the unchanged dialog."
            )

            window.child_window(title="Dependency graph", control_type="Button").invoke()
            graph = _wait_child_window(window, "Query dependencies")
            graph.set_focus()
            graph_visited = _tab_cycle(graph, automation, manager.pid, "dependency-graph dialog")
            graph_required = [
                "Text summary of query dependencies",
                "Close query dependency graph",
            ]
            missing_graph = [name for name in graph_required if name not in graph_visited]
            if missing_graph:
                raise AssertionError(
                    f"Dependency-graph dialog keyboard traversal skipped controls: {missing_graph!r}; "
                    f"visited={graph_visited!r}"
                )
            graph.type_keys("{ESC}", set_foreground=True)
            graph.wait_not("visible", timeout=5)
            window.wait("visible", timeout=5)
            print("PASS: dependency-graph dialog Tab cycle reached its text summary and Close action; Escape closed it.")

            mapping_button = window.child_window(
                title="Edit source environment mappings", control_type="Button"
            )
            mapping_button.invoke()
            dialog = _wait_child_window(window, "Edit source environment mappings")
            dialog.set_focus()
            dialog_visited = _tab_cycle(dialog, automation, manager.pid, "source-mapping dialog")
            dialog_required = [
                "Source paths",
                "Profile source path mappings",
                "Add source path mapping",
                "Select all source path mappings",
                "Remove selected source path mappings",
                "Cancel mapping edits",
                "Save profile mappings",
            ]
            missing_dialog = [name for name in dialog_required if name not in dialog_visited]
            if missing_dialog:
                raise AssertionError(
                    f"Source-mapping dialog keyboard traversal skipped controls: {missing_dialog!r}; "
                    f"visited={dialog_visited!r}"
                )
            dialog.type_keys("{RIGHT}", set_foreground=True)
            time.sleep(0.1)
            focused = automation.GetFocusedElement()
            if focused.CurrentProcessId != manager.pid or focused.CurrentName != "M source locations":
                raise AssertionError(
                    "Right-arrow from the Source paths tab did not reach M source locations; "
                    f"focused={focused.CurrentName!r}"
                )
            m_dialog_visited = _tab_cycle(dialog, automation, manager.pid, "M-source mapping tab")
            m_required = [
                "M source locations",
                "M source location mappings",
                "Add M source location mapping",
                "Select all M source location mappings",
                "Remove selected M source location mappings",
                "Cancel mapping edits",
                "Save profile mappings",
            ]
            missing_m = [name for name in m_required if name not in m_dialog_visited]
            if missing_m:
                raise AssertionError(
                    f"M-source mapping tab keyboard traversal skipped controls: {missing_m!r}; "
                    f"visited={m_dialog_visited!r}"
                )
            dialog.type_keys("{ESC}", set_foreground=True)
            dialog.wait_not("visible", timeout=5)
            window.wait("visible", timeout=5)
            print(
                "PASS: source-mapping dialog Tab cycles reached both tables and their actions; "
                "Right Arrow switched categories and Escape closed the dialog."
            )

        return 0
    finally:
        if window is not None:
            try:
                win32gui.SendMessage(window.handle, win32con.WM_CLOSE, 0, 0)
            except Exception:
                pass
        try:
            manager.wait(timeout=10)
        except subprocess.TimeoutExpired:
            manager.terminate()
            try:
                manager.wait(timeout=5)
            except subprocess.TimeoutExpired:
                manager.kill()
                manager.wait(timeout=5)
        excel_after = _excel_pids()
        unexpected = excel_after - excel_before
        if unexpected:
            raise RuntimeError(f"Keyboard UI smoke unexpectedly created Excel process(es): {sorted(unexpected)}")
        if manager.returncode != 0:
            raise RuntimeError(f"Packaged manager exited with code {manager.returncode}.")
        isolated_user_data.cleanup()


if __name__ == "__main__":
    sys.exit(main())
