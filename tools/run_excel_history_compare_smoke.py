"""Compare two locally captured query snapshots using disposable desktop Excel workbooks."""
from __future__ import annotations

import hashlib
import gc
from pathlib import Path
import sys
import tempfile
import time

import pythoncom
import win32api
import win32con
import win32com.client
import win32event
import win32process

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "app" / "src"))
from pq_installer.core.history import HistoryStore
from pq_installer.core.excel_inspector import ExcelQueryInspector
from pq_installer.core.excel_capabilities import probe_excel_capabilities
from pq_installer.core.excel_queries import ExcelQueryInstaller
from pq_installer.core.models import ExistingQueryMode, LoadMode, QuerySource
from pq_installer.ui.history_worker import SnapshotCompareWorker


OLD_M = "let\n    Source = #table({\"Value\"}, {{1}})\nin\n    Source"
NEW_M = "let\n    Source = #table({\"Value\"}, {{2}})\nin\n    Source"
STABLE_M = "let\n    Source = #table({\"Value\"}, {{\"stable\"}})\nin\n    Source"


def sha256(path: Path) -> str:
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def process_running(pid: int) -> bool:
    try:
        handle = win32api.OpenProcess(win32con.SYNCHRONIZE, False, pid)
    except Exception:
        return False
    try:
        return win32event.WaitForSingleObject(handle, 0) == win32event.WAIT_TIMEOUT
    finally:
        win32api.CloseHandle(handle)


def assert_snapshot_load_transition(store, before, after, expected_before, expected_after, pids):
    before_hashes = (sha256(before["snapshot_path"]), sha256(after["snapshot_path"]))
    inspector_pids = []

    def tracked_excel_factory():
        app = win32com.client.DispatchEx("Excel.Application")
        inspector_pids.append(win32process.GetWindowThreadProcessId(app.Hwnd)[1])
        return app

    worker = SnapshotCompareWorker(
        store, before["snapshot_id"], after["snapshot_id"],
        inspector_factory=lambda: ExcelQueryInspector(excel_factory=tracked_excel_factory),
    )
    results, failures = [], []
    worker.completed.connect(results.append)
    worker.failed.connect(failures.append)
    worker.run()
    assert not failures, f"snapshot compare failed: {failures}"
    assert len(results) == 1, f"expected one comparison result, got {len(results)}"
    changes = {change.name: change for change in results[0]["changes"]}
    change = changes["LoadSwitch"]
    assert change.status == "MODIFIED", change
    if expected_before.endswith(":"):
        assert change.before.load_state.startswith(expected_before), change.before
    else:
        assert change.before.load_state == expected_before, change.before
    if expected_after.endswith(":"):
        assert change.after.load_state.startswith(expected_after), change.after
    else:
        assert change.after.load_state == expected_after, change.after
    assert "Load destination changed" in change.reason
    assert before_hashes == (sha256(before["snapshot_path"]), sha256(after["snapshot_path"])), \
        "comparison changed immutable snapshot bytes"
    pids.extend(inspector_pids)
    return change


def create_workbook(path: Path) -> tuple[str, str, int]:
    app = win32com.client.DispatchEx("Excel.Application")
    app.Visible = False
    app.DisplayAlerts = False
    pid = win32process.GetWindowThreadProcessId(app.Hwnd)[1]
    book = None
    previous = None
    try:
        previous = app.AutomationSecurity
        app.AutomationSecurity = 3
        book = app.Workbooks.Add()
        book.Queries.Add("Changed", OLD_M, "")
        book.Queries.Add("Stable", STABLE_M, "")
        book.Queries.Add("LoadSwitch", STABLE_M, "")
        book.SaveAs(str(path), FileFormat=51)
        version, build = str(app.Version), str(app.Build)
        book.Close(SaveChanges=False)
        book = None
        return version, build, pid
    finally:
        if book is not None:
            book.Close(SaveChanges=False)
        if previous is not None:
            app.AutomationSecurity = previous
        app.Quit()
        del app


def update_workbook(path: Path) -> int:
    app = win32com.client.DispatchEx("Excel.Application")
    app.Visible = False
    app.DisplayAlerts = False
    pid = win32process.GetWindowThreadProcessId(app.Hwnd)[1]
    book = None
    previous = None
    try:
        previous = app.AutomationSecurity
        app.AutomationSecurity = 3
        book = app.Workbooks.Open(str(path), UpdateLinks=0, ReadOnly=False,
                                  IgnoreReadOnlyRecommended=True, AddToMru=False)
        book.Queries.Item("Changed").Formula = NEW_M
        book.Queries.Add("Added", STABLE_M, "")
        book.Save()
        book.Close(SaveChanges=False)
        book = None
        return pid
    finally:
        if book is not None:
            book.Close(SaveChanges=False)
        if previous is not None:
            app.AutomationSecurity = previous
        app.Quit()
        del app


def main() -> None:
    pythoncom.CoInitialize()
    try:
        with tempfile.TemporaryDirectory(prefix="pqm-history-compare-") as temp:
            root = Path(temp)
            workbook = root / "history-compare.xlsx"
            version, build, first_pid = create_workbook(workbook)
            store = HistoryStore(root / "history.sqlite3")
            store.set_snapshots_enabled(True)
            first = store.create_snapshot(workbook, event="workbook_snapshot_before_import")
            second_pid = update_workbook(workbook)

            owned_pids = [first_pid, second_pid]
            original_dispatch = win32com.client.DispatchEx

            def capture_dispatch(*args, **kwargs):
                app = original_dispatch(*args, **kwargs)
                owned_pids.append(win32process.GetWindowThreadProcessId(app.Hwnd)[1])
                return app

            # Capability probing and the actual load transition each own an isolated Excel instance.
            win32com.client.DispatchEx = capture_dispatch
            try:
                capabilities = probe_excel_capabilities()
            finally:
                win32com.client.DispatchEx = original_dispatch
            if not capabilities.supports(LoadMode.CONNECTION_ONLY.value):
                raise RuntimeError(f"Excel build does not support connection-only mode: {capabilities.reasons}")
            source_path = root / "LoadSwitch.m"
            source_path.write_text(STABLE_M, encoding="utf-8")
            load_source = QuerySource(source_path, "LoadSwitch", load_mode=LoadMode.CONNECTION_ONLY)
            win32com.client.DispatchEx = capture_dispatch
            try:
                load_result = ExcelQueryInstaller().install(
                    workbook, [load_source], ExistingQueryMode.REPLACE, capabilities=capabilities,
                )
            finally:
                win32com.client.DispatchEx = original_dispatch
            if not load_result or load_result[0].status.value != "Replaced":
                raise AssertionError(f"connection-only transition was not applied: {load_result!r}")
            second = store.create_snapshot(workbook, event="workbook_snapshot_after_import")
            hashes_before = (sha256(first["snapshot_path"]), sha256(second["snapshot_path"]))
            results, failures = [], []
            inspector_pids = []

            def tracked_excel_factory():
                app = win32com.client.DispatchEx("Excel.Application")
                inspector_pids.append(win32process.GetWindowThreadProcessId(app.Hwnd)[1])
                return app

            worker = SnapshotCompareWorker(
                store, first["snapshot_id"], second["snapshot_id"],
                inspector_factory=lambda: ExcelQueryInspector(excel_factory=tracked_excel_factory),
            )
            worker.completed.connect(results.append)
            worker.failed.connect(failures.append)
            worker.run()
            assert not failures, f"snapshot compare failed: {failures}"
            assert len(results) == 1, f"expected one comparison result, got {len(results)}"
            changes = {change.name: change for change in results[0]["changes"]}
            assert changes["Changed"].status == "MODIFIED", changes["Changed"]
            assert "{{1}}" in changes["Changed"].diff and "{{2}}" in changes["Changed"].diff
            assert changes["Stable"].status == "UNCHANGED", changes["Stable"]
            assert changes["Added"].status == "ADDED", changes["Added"]
            assert changes["LoadSwitch"].status == "MODIFIED", changes["LoadSwitch"]
            assert "Load destination changed" in changes["LoadSwitch"].reason
            assert changes["LoadSwitch"].before.load_state == "definition_only"
            assert changes["LoadSwitch"].after.load_state == "connection_only"
            assert hashes_before == (sha256(first["snapshot_path"]), sha256(second["snapshot_path"])), \
                "comparison changed immutable snapshot bytes"

            # Compare snapshots from equivalent disposable workbook shapes for
            # worksheet and Data Model states. The production importer correctly
            # refuses implicit load-mode conversions on an existing workbook.
            transition_results = [
                f"{changes['LoadSwitch'].before.load_state} -> {changes['LoadSwitch'].after.load_state}"
            ]
            loaded_snapshots = {}
            for mode in (LoadMode.NEW_SHEET, LoadMode.DATA_MODEL):
                if not capabilities.supports(mode.value):
                    raise RuntimeError(f"Excel build did not certify {mode.value}: {capabilities.reasons}")
                mode_workbook = root / f"history-{mode.value}.xlsx"
                mode_version, mode_build, seed_pid = create_workbook(mode_workbook)
                assert (mode_version, mode_build) == (version, build)
                owned_pids.append(seed_pid)
                mode_source_path = root / f"LoadSwitch-{mode.value}.m"
                mode_source_path.write_text(STABLE_M, encoding="utf-8")
                load_source = QuerySource(
                    mode_source_path, "LoadSwitch", load_mode=mode,
                    sheet_name="HistoryOutput" if mode == LoadMode.NEW_SHEET else None,
                    table_name="tblHistoryLoadSwitch" if mode == LoadMode.NEW_SHEET else None,
                )
                win32com.client.DispatchEx = capture_dispatch
                try:
                    mode_result = ExcelQueryInstaller().install(
                        mode_workbook, [load_source], ExistingQueryMode.REPLACE, capabilities=capabilities,
                    )
                finally:
                    win32com.client.DispatchEx = original_dispatch
                if not mode_result or mode_result[0].status.value != "Replaced":
                    raise AssertionError(f"{mode.value} load was not created: {mode_result!r}")
                loaded_snapshots[mode] = store.create_snapshot(
                    mode_workbook, event=f"load_state_{mode.value}",
                )
            transition = assert_snapshot_load_transition(
                store, loaded_snapshots[LoadMode.NEW_SHEET], loaded_snapshots[LoadMode.DATA_MODEL],
                "worksheet:", "data_model", owned_pids,
            )
            transition_results.append(
                f"{transition.before.load_state} -> {transition.after.load_state}"
            )
            for owned_pid in (*owned_pids, *inspector_pids):
                for _ in range(50):
                    if not process_running(owned_pid):
                        break
                    time.sleep(0.1)
                assert not process_running(owned_pid), f"smoke-owned Excel process did not exit: {owned_pid}"
            print(f"PASS Excel {version} build {build}; snapshot compare found formula and load Modified, "
                  "Unchanged, and Added; observed load comparisons: " + ", ".join(transition_results) + "; "
                  f"exact M diff present; snapshot hashes unchanged; {len(inspector_pids) + len(owned_pids)} owned Excel "
                  "processes exited; no refresh")
    finally:
        pythoncom.CoUninitialize()


if __name__ == "__main__":
    main()
