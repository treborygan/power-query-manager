"""Build disposable Office/DAO dependency fixtures and verify the Windows scanner.

Run from a Windows desktop with Excel and ACE DAO 16.0 installed:

    python tools/run_dependency_windows_acceptance.py

Optional existing mapped SMB, UNC, and synchronized roots can be supplied to
repeat read-only scans there. The harness creates a uniquely named child under
each supplied root and removes only that child unless --keep-fixtures is used.
It never creates shares, maps drives, changes sync settings, refreshes queries,
or opens source files for editing during a scan.
"""
from __future__ import annotations

import argparse
import csv
import ctypes
import hashlib
import json
import os
from pathlib import Path
import shutil
import subprocess
import sys
import tempfile
import time
import uuid

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "app" / "src"))


def _sha(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def _snapshot(paths):
    result = {}
    for path in sorted(paths):
        stat = path.stat()
        result[path.name] = {
            "sha256": _sha(path), "size": stat.st_size,
            "mtime_ns": stat.st_mtime_ns,
        }
    return result


def _excel_pids() -> set[int]:
    output = subprocess.run(
        ["tasklist.exe", "/FI", "IMAGENAME eq EXCEL.EXE", "/FO", "CSV", "/NH"],
        check=True, capture_output=True, text=True, encoding="mbcs", errors="replace",
    ).stdout
    import io
    return {int(row[1]) for row in csv.reader(io.StringIO(output))
            if len(row) > 1 and row[1].isdigit()}


def _peak_working_set(pid: int) -> int:
    """Read the Windows process peak RSS counter, or return zero if it exited."""
    try:
        import win32api
        import win32con
        import win32process
        handle = win32api.OpenProcess(
            win32con.PROCESS_QUERY_INFORMATION | win32con.PROCESS_VM_READ, False, pid
        )
        try:
            return int(win32process.GetProcessMemoryInfo(handle)["PeakWorkingSetSize"])
        finally:
            win32api.CloseHandle(handle)
    except Exception:
        return 0


def _process_tree(root_pid: int) -> set[int]:
    """Return the launcher PID and current descendants using Toolhelp32."""
    from ctypes import wintypes

    class PROCESSENTRY32W(ctypes.Structure):
        _fields_ = [
            ("dwSize", wintypes.DWORD), ("cntUsage", wintypes.DWORD),
            ("th32ProcessID", wintypes.DWORD), ("th32DefaultHeapID", ctypes.c_size_t),
            ("th32ModuleID", wintypes.DWORD), ("cntThreads", wintypes.DWORD),
            ("th32ParentProcessID", wintypes.DWORD), ("pcPriClassBase", ctypes.c_long),
            ("dwFlags", wintypes.DWORD), ("szExeFile", wintypes.WCHAR * 260),
        ]

    kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
    kernel32.CreateToolhelp32Snapshot.argtypes = [wintypes.DWORD, wintypes.DWORD]
    kernel32.CreateToolhelp32Snapshot.restype = wintypes.HANDLE
    kernel32.Process32FirstW.argtypes = [wintypes.HANDLE, ctypes.POINTER(PROCESSENTRY32W)]
    kernel32.Process32FirstW.restype = wintypes.BOOL
    kernel32.Process32NextW.argtypes = [wintypes.HANDLE, ctypes.POINTER(PROCESSENTRY32W)]
    kernel32.Process32NextW.restype = wintypes.BOOL
    kernel32.CloseHandle.argtypes = [wintypes.HANDLE]
    snapshot = kernel32.CreateToolhelp32Snapshot(0x00000002, 0)
    if snapshot == wintypes.HANDLE(-1).value:
        return {root_pid}
    try:
        entry = PROCESSENTRY32W()
        entry.dwSize = ctypes.sizeof(entry)
        parents = {}
        if kernel32.Process32FirstW(snapshot, ctypes.byref(entry)):
            while True:
                parents[int(entry.th32ProcessID)] = int(entry.th32ParentProcessID)
                entry.dwSize = ctypes.sizeof(entry)
                if not kernel32.Process32NextW(snapshot, ctypes.byref(entry)):
                    break
        result = {root_pid}
        changed = True
        while changed:
            changed = False
            for pid, parent in parents.items():
                if parent in result and pid not in result:
                    result.add(pid)
                    changed = True
        return result
    finally:
        kernel32.CloseHandle(snapshot)


def _new_dao_database(engine, path: Path, version: int):
    # DAO dbVersion40=64 for Jet 4 .mdb; dbVersion120=1280 for ACE .accdb.
    database = engine.CreateDatabase(
        str(path), ";LANGID=0x0409;CP=1252;COUNTRY=0", version
    )
    table = database.CreateTableDef("Inventory")
    field = table.CreateField("ItemID", 4)  # dbLong
    table.Fields.Append(field)
    database.TableDefs.Append(table)
    query = database.CreateQueryDef(
        "LocalInventoryQuery", "SELECT ItemID FROM Inventory"
    )
    query = None
    database.Close()


def _link_table(engine, database_path: Path, table_name: str, source: Path,
                *, query_name: str | None = None):
    database = engine.OpenDatabase(str(database_path), False, False)
    table = database.CreateTableDef(table_name)
    table.Connect = ";DATABASE=" + str(source)
    table.SourceTableName = "Inventory"
    table.Attributes = int(table.Attributes) | -0x80000000  # signed DAO dbAttachedTable flag
    database.TableDefs.Append(table)
    query = database.CreateQueryDef(
        query_name or "LinkedInventoryQuery",
        f"SELECT ItemID FROM [{table_name}] IN '{source}'",
    )
    query = None
    table = None
    database.Close()


def _excel_fixture(folder: Path, *, large_rows: int):
    import pythoncom
    import win32process
    import win32com.client

    pythoncom.CoInitialize()
    excel = None
    workbook = None
    owned_pid = None
    fixture_result = None
    cleanup_started = None
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.AskToUpdateLinks = False
        excel.AutomationSecurity = 3
        owned_pid = int(win32process.GetWindowThreadProcessId(int(excel.Hwnd))[1])

        external_target = folder / "ExternalTarget.xlsx"
        workbook = excel.Workbooks.Add()
        excel.Calculation = -4135  # xlCalculationManual; never evaluate external links during fixture creation
        workbook.Worksheets(1).Name = "Data"
        workbook.Worksheets(1).Cells(1, 1).Value = 42
        workbook.SaveAs(str(external_target), 51)
        workbook.Close(False)
        workbook = None

        report = folder / "Report.xlsx"
        workbook = excel.Workbooks.Add()
        workbook.Worksheets(1).Cells(1, 1).Formula = (
            f"='{external_target.parent}\\[{external_target.name}]Data'!$A$1"
        )
        workbook.SaveAs(str(report), 51)
        workbook.Close(False)
        workbook = None

        mashup = folder / "Queries.xlsm"
        workbook = excel.Workbooks.Add()
        query_specs = (
            ("PQM_Accdb_Source", f'let Source = Access.Database(File.Contents("{folder / "Warehouse.accdb"}"), [CreateNavigationProperties=false]) in Source'),
            ("PQM_Mdb_Source", f'let Source = Access.Database(File.Contents("{folder / "Stock.mdb"}"), [CreateNavigationProperties=false]) in Source'),
            ("PQM_Ambiguous_Source", 'let Source = File.Contents("Ambiguous.accdb") in Source'),
            ("PQM_Unresolved_Source", 'let Source = File.Contents("OutsideScope.xlsx") in Source'),
        )
        for query_name, m in query_specs:
            workbook.Queries.Add(query_name, m)
            connection_string = (
                "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;"
                f'Location={query_name};Extended Properties=""'
            )
            workbook.Connections.Add2(
                "Query - " + query_name, "Disposable dependency acceptance query",
                connection_string, f"SELECT * FROM [{query_name}]", 2, False, False,
            )
        workbook.SaveAs(str(mashup), 52)  # genuine macro-enabled Excel package
        workbook.Close(False)
        workbook = None

        large = folder / "LargeExternalFormulas.xlsx"
        workbook = excel.Workbooks.Add()
        sheet = workbook.Worksheets(1)
        formula = f"='{external_target.parent}\\[{external_target.name}]Data'!$A$1"
        # Fill a real worksheet range with external-workbook formulas in one COM call.
        sheet.Range(f"A1:Z{large_rows}").Formula = formula
        workbook.SaveAs(str(large), 51)
        workbook.Close(False)
        workbook = None
        fixture_result = {"excel_version": str(excel.Version), "excel_build": str(excel.Build),
                          "seed_excel_pid": owned_pid, "large_formula_cells": large_rows * 26}
        return fixture_result
    finally:
        if workbook is not None:
            try:
                workbook.Close(False)
            except Exception:
                pass
        workbook = None
        if "sheet" in locals():
            sheet = None
        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                pass
        excel = None
        pythoncom.CoUninitialize()
        if owned_pid is not None:
            import gc
            gc.collect()
            cleanup_started = time.monotonic()
            deadline = cleanup_started + 20
            while owned_pid in _excel_pids() and time.monotonic() < deadline:
                time.sleep(0.1)
            if owned_pid in _excel_pids():
                # This PID came from DispatchEx in this harness, and was not
                # present in the captured baseline. Contain only that owned
                # automation process; record this as a failing cleanup result.
                forced = subprocess.run(
                    ["taskkill.exe", "/PID", str(owned_pid), "/T", "/F"],
                    capture_output=True, text=True, encoding="mbcs", errors="replace",
                )
                deadline = time.monotonic() + 10
                while owned_pid in _excel_pids() and time.monotonic() < deadline:
                    time.sleep(0.1)
                if owned_pid in _excel_pids():
                    raise RuntimeError(f"Owned Excel PID {owned_pid} resisted Quit() and cleanup.")
                if fixture_result is not None:
                    fixture_result["excel_forced_termination"] = True
                    fixture_result["excel_force_cleanup_exit_code"] = forced.returncode
            if fixture_result is not None:
                fixture_result["excel_quit_to_process_exit_seconds"] = time.monotonic() - cleanup_started


def _create_fixtures(folder: Path, *, large_rows: int, dao_progid: str):
    import pythoncom
    import win32com.client

    folder.mkdir(parents=True, exist_ok=False)
    pythoncom.CoInitialize()
    engine = None
    try:
        try:
            engine = win32com.client.Dispatch(dao_progid)
        except Exception as exc:
            available = []
            for candidate in ("DAO.DBEngine.160", "DAO.DBEngine.120", "DAO.DBEngine.36"):
                try:
                    probe = win32com.client.Dispatch(candidate)
                    available.append(candidate)
                    probe = None
                except Exception:
                    pass
            raise RuntimeError(
                f"Required DAO provider {dao_progid} is not registered; available DAO providers: "
                f"{available or 'none'}. Install the matching Microsoft Access Database Engine."
            ) from exc
        _new_dao_database(engine, folder / "Warehouse.accdb", 1280)
        _new_dao_database(engine, folder / "Stock.mdb", 64)
        _new_dao_database(engine, folder / "Archive.mdb", 64)
        _new_dao_database(engine, folder / "Broker.accdb", 1280)
        for duplicate_folder in (folder / "DuplicateA", folder / "DuplicateB"):
            duplicate_folder.mkdir()
            _new_dao_database(engine, duplicate_folder / "Ambiguous.accdb", 1280)
        _link_table(engine, folder / "Warehouse.accdb", "LinkedStock", folder / "Stock.mdb")
        _link_table(engine, folder / "Stock.mdb", "LinkedArchive", folder / "Archive.mdb")
        _link_table(engine, folder / "Broker.accdb", "WarehouseLink", folder / "Warehouse.accdb",
                    query_name="WarehouseSavedQuery")
        _link_table(engine, folder / "Broker.accdb", "StockLink", folder / "Stock.mdb",
                    query_name="StockSavedQuery")
    finally:
        engine = None
        pythoncom.CoUninitialize()
    excel = _excel_fixture(folder, large_rows=large_rows)
    return excel


def _scan_cli(cli_exe: Path, files: list[Path], depth: str):
    command = [str(cli_exe), "dependency-map", "--depth", depth]
    for path in files:
        command += ["--file", str(path)]
    process = subprocess.Popen(command, stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                               text=True, encoding="utf-8", errors="replace")
    peak = 0
    start = time.perf_counter()
    deadline = start + 300
    while True:
        try:
            # Drain both pipes while the process runs; otherwise a large JSON
            # report can fill stdout and deadlock the child before it exits.
            stdout, stderr = process.communicate(timeout=0.05)
            break
        except subprocess.TimeoutExpired:
            pids = _process_tree(process.pid)
            peak = max(peak, sum(_peak_working_set(pid) for pid in pids))
            if time.perf_counter() >= deadline:
                subprocess.run(["taskkill.exe", "/PID", str(process.pid), "/T", "/F"],
                               capture_output=True, text=True)
                process.kill()
                process.communicate()
                raise TimeoutError(f"Packaged CLI {depth} scan exceeded 300 seconds.")
    elapsed = time.perf_counter() - start
    if process.returncode not in {0, 5}:
        raise RuntimeError(f"Packaged CLI {depth} scan exited {process.returncode}: {stderr or stdout}")
    report = json.loads(stdout)
    if (process.returncode == 5) != bool(report.get("issues") or report.get("cancelled")):
        raise AssertionError(f"Packaged CLI {depth} exit code disagrees with its coverage report.")
    return report, {"elapsed_seconds": elapsed, "peak_working_set_bytes": peak}


def _assert_expected(report, files):
    expected = {
        ("Warehouse.accdb", "Stock.mdb"),
        ("Stock.mdb", "Archive.mdb"),
        ("Broker.accdb", "Warehouse.accdb"),
        ("Broker.accdb", "Stock.mdb"),
        ("Queries.xlsm", "Warehouse.accdb"),
        ("Queries.xlsm", "Stock.mdb"),
        ("Report.xlsx", "ExternalTarget.xlsx"),
        ("LargeExternalFormulas.xlsx", "ExternalTarget.xlsx"),
    }
    names = {str(path): path.name for path in files}
    actual = {(names.get(edge["consumer_path"], Path(edge["consumer_path"]).name),
               names.get(edge["target_path"], Path(edge["target_path"]).name))
              for edge in report["edges"]}
    missing = expected - actual
    if missing:
        raise AssertionError(f"Expected graph edges were not detected: {sorted(missing)}; got {sorted(actual)}")
    access_sources = {(Path(edge["consumer_path"]).name, Path(edge["target_path"]).name,
                       edge["source_kind"], edge["source_name"]) for edge in report["edges"]}
    for consumer, target, kind in (
        ("Warehouse.accdb", "Stock.mdb", "Access linked table"),
        ("Warehouse.accdb", "Stock.mdb", "Access saved query"),
        ("Stock.mdb", "Archive.mdb", "Access linked table"),
        ("Stock.mdb", "Archive.mdb", "Access saved query"),
        ("Broker.accdb", "Warehouse.accdb", "Access linked table"),
        ("Broker.accdb", "Warehouse.accdb", "Access saved query"),
        ("Broker.accdb", "Stock.mdb", "Access linked table"),
        ("Broker.accdb", "Stock.mdb", "Access saved query"),
    ):
        if not any(a == consumer and b == target and c == kind for a, b, c, _ in access_sources):
            raise AssertionError(f"Missing Access {kind} dependency: {consumer} -> {target}.")
    mashup_sources = {(Path(edge["consumer_path"]).name, Path(edge["target_path"]).name,
                       edge["source_kind"]) for edge in report["edges"]}
    if not {("Queries.xlsm", "Warehouse.accdb", "Power Query"),
            ("Queries.xlsm", "Stock.mdb", "Power Query")}.issubset(mashup_sources):
        raise AssertionError("Genuine Excel Mashup source dependencies were not both discovered.")
    unresolved_names = {item["reference_filename"] for item in report["unresolved"]}
    if not {"Ambiguous.accdb", "OutsideScope.xlsx"}.issubset(unresolved_names):
        raise AssertionError(f"Ambiguous or unresolved Power Query references were lost: {unresolved_names}")
    if not any(issue["status"] == "ambiguous_reference" for issue in report["issues"]):
        raise AssertionError("Duplicate-basename reference did not produce an ambiguity warning.")
    return {"expected_edges": sorted([list(pair) for pair in expected]),
            "actual_edges": sorted([list(pair) for pair in actual]),
            "dependency_edge_count": len(report["edges"])}


def _exercise_dialog(report, files):
    from dataclasses import asdict
    from PySide6.QtWidgets import QApplication, QFileDialog, QMessageBox
    from pq_installer.core.reverse_dependencies import filter_dependency_map
    from pq_installer.ui.dependency_finder_dialog import DependencyFinderDialog
    from pq_installer.ui.file_dependency_graph_dialog import FileDependencyGraphDialog

    app = QApplication.instance() or QApplication([])
    dialog = DependencyFinderDialog()
    dialog.mode_combo.setCurrentIndex(dialog.mode_combo.findData("map"))
    dialog.latest_report = report
    dialog._show_report(report)
    report_data = asdict(report)
    file_by_name = {Path(path).name: path for path in report.files}
    checks = {}
    for name in ("Warehouse.accdb", "Stock.mdb", "Queries.xlsm", "ExternalTarget.xlsx"):
        selected = file_by_name[name]
        dialog.map_filter_combo.setCurrentIndex(dialog.map_filter_combo.findData(selected))
        for direction in ("incoming", "outgoing", "both"):
            dialog.direction_combo.setCurrentIndex(dialog.direction_combo.findData(direction))
            app.processEvents()
            rows = dialog.results_table.rowCount()
            wanted = len(filter_dependency_map(report, selected, direction=direction))
            if rows != wanted:
                raise AssertionError(f"GUI {direction} filter for {name}: {rows} rows, expected {wanted}.")
            checks[f"{name}:{direction}"] = rows

    graph = FileDependencyGraphDialog(report)
    if {(Path(edge.consumer_path).name, Path(edge.target_path).name) for edge in graph.report.edges} != {
        (Path(edge["consumer_path"]).name, Path(edge["target_path"]).name)
        for edge in report_data["edges"]
    }:
        raise AssertionError("Visual dependency graph edge model differs from exported report.")
    graph.show()
    app.processEvents()
    if not graph.isVisible() or not graph.graph_view.scene().items():
        raise AssertionError("Visual dependency graph did not render its nodes/relationships.")
    graph.close()

    exports = {}
    original_picker = QFileDialog.getSaveFileName
    original_info = QMessageBox.information
    try:
        for kind in ("json", "csv"):
            output = Path(tempfile.gettempdir()) / f"pqm-dependency-acceptance-{uuid.uuid4().hex}.{kind}"
            QFileDialog.getSaveFileName = staticmethod(lambda *_, target=str(output): (target, ""))
            QMessageBox.information = staticmethod(lambda *_, **__: QMessageBox.StandardButton.Ok)
            dialog._export(kind)
            if kind == "json":
                rendered = json.loads(output.read_text(encoding="utf-8"))
                if len(rendered["edges"]) != len(report.edges):
                    raise AssertionError("GUI JSON export lost dependency edges.")
            else:
                with output.open(encoding="utf-8", newline="") as stream:
                    rows = list(csv.DictReader(stream))
                count = sum(row["row_type"] == "dependency" for row in rows)
                if count != len(report.edges):
                    raise AssertionError("GUI CSV export lost dependency edges.")
            exports[kind] = {"verified": True, "edge_count": len(report.edges)}
            output.unlink()
    finally:
        QFileDialog.getSaveFileName = original_picker
        QMessageBox.information = original_info
        dialog.close()
        app.processEvents()
    return {"filter_row_counts": checks, "graph_edges": len(report.edges),
            "exports": exports}


def _gui_cancel_progress(files):
    """Exercise the dialog's threaded progress/cancel wiring on real workbook scans."""
    from PySide6.QtWidgets import QApplication
    from pq_installer.core.reverse_dependencies import scan_dependency_map
    from pq_installer.ui.dependency_finder_dialog import DependencyFinderDialog

    app = QApplication.instance() or QApplication([])
    def slower_scan(paths, *, depth, progress, cancelled):
        def paced(done, total, path):
            progress(done, total, path)
            time.sleep(0.12)
        return scan_dependency_map(paths, depth=depth, progress=paced, cancelled=cancelled)

    dialog = DependencyFinderDialog(map_runner=slower_scan)
    dialog.mode_combo.setCurrentIndex(dialog.mode_combo.findData("map"))
    dialog.add_scan_paths(files, recursive=False)
    dialog.consumer_list.selectAll()
    dialog._start_scan()
    deadline = time.monotonic() + 30
    cancelled_at = None
    while dialog._running and time.monotonic() < deadline:
        app.processEvents()
        if cancelled_at is None and dialog.progress_bar.value() >= 1:
            cancelled_at = time.monotonic()
            dialog._cancel_scan()
        time.sleep(0.01)
    while dialog._scan_thread is not None and time.monotonic() < deadline:
        app.processEvents()
        time.sleep(0.01)
    report = dialog.latest_report
    result = {"progress_observed": cancelled_at is not None,
              "cancelled_result": bool(report and report.cancelled),
              "inspected_before_cancel": len(report.inspected) if report else 0}
    dialog.close()
    app.processEvents()
    if not result["progress_observed"] or not result["cancelled_result"]:
        raise AssertionError(f"GUI progress/cancel check failed: {result}")
    return result


def _run_location(label: str, fixture_source: Path, root: Path, cli_exe: Path,
                  *, keep: bool, large_rows: int):
    run_dir = root / f"pqm-dependency-acceptance-{uuid.uuid4().hex}"
    shutil.copytree(fixture_source, run_dir)
    files = sorted(path for path in run_dir.rglob("*")
                   if path.is_file() and path.suffix.casefold() in {".xlsx", ".xlsm", ".accdb", ".mdb"})
    before = _snapshot(files)
    lock_before = sorted(path.name for path in run_dir.glob("*.laccdb")) + sorted(path.name for path in run_dir.glob("*.ldb"))
    temporary_before = {path.name for path in Path(tempfile.gettempdir()).glob("pqm-access-scan-*")}
    excel_before = _excel_pids()
    modes = {}
    cli_reports = {}
    for depth in ("quick", "deep"):
        report, metrics = _scan_cli(cli_exe, files, depth)
        cli_reports[depth] = report
        graph = _assert_expected(report, files)
        modes[depth] = {**metrics, **graph, "issue_count": len(report["issues"]),
                        "unresolved_count": len(report["unresolved"])}
        if report["depth"] != depth:
            raise AssertionError(f"CLI returned depth {report['depth']!r} for requested {depth!r}.")
    reverse_command = [str(cli_exe), "dependency-scan", "--target",
                       str(run_dir / "Warehouse.accdb"), "--target", str(run_dir / "Stock.mdb"),
                       "--folder", str(run_dir), "--recursive", "--depth", "deep"]
    reverse_process = subprocess.run(reverse_command, capture_output=True, text=True,
                                     encoding="utf-8", errors="replace")
    if reverse_process.returncode not in {0, 5}:
        raise RuntimeError(f"Packaged multi-target dependency-scan failed: {reverse_process.stderr}")
    reverse_report = json.loads(reverse_process.stdout)
    reverse_pairs = {(Path(hit["consumer_path"]).name, Path(hit["target_path"]).name)
                     for hit in reverse_report["hits"]}
    if not {("Broker.accdb", "Warehouse.accdb"), ("Broker.accdb", "Stock.mdb"),
            ("Queries.xlsm", "Warehouse.accdb"), ("Queries.xlsm", "Stock.mdb")}.issubset(reverse_pairs):
        raise AssertionError(f"Packaged multi-target reverse scan missed expected Access/Mashup hits: {reverse_pairs}")
    after = _snapshot(files)
    lock_after = sorted(path.name for path in run_dir.glob("*.laccdb")) + sorted(path.name for path in run_dir.glob("*.ldb"))
    temporary_after = {path.name for path in Path(tempfile.gettempdir()).glob("pqm-access-scan-*")}
    excel_after = _excel_pids()
    if before != after:
        raise AssertionError(f"Source hashes/timestamps changed during {label} scan.")
    if lock_before != lock_after:
        raise AssertionError(f"Access lock files changed during {label}: {lock_before} -> {lock_after}.")
    if temporary_after - temporary_before:
        raise AssertionError(f"Access scan temporary directories remain after {label}: {sorted(temporary_after - temporary_before)}.")
    if excel_after != excel_before:
        raise AssertionError(f"Excel process set changed during {label} scan: {excel_before} -> {excel_after}.")
    result = {"status": "PASS", "root_type": label, "file_count": len(files),
              "source_snapshot_unchanged": True, "lock_files_before": lock_before,
              "lock_files_after": lock_after, "excel_pids_unchanged": True,
              "access_temp_directories_cleaned": True,
              "source_snapshot_before": before, "source_snapshot_after": after,
              "modes": modes,
              "multi_target_reverse_scan": {"status": "PASS", "target_count": 2,
                                             "matched_relationships": len(reverse_report["hits"])}}
    if label == "local":
        from pq_installer.core.reverse_dependencies import scan_dependency_map
        ui_report = scan_dependency_map(files, depth="deep")
        cli_edges = {(edge["consumer_path"], edge["target_path"], edge["source_kind"], edge["source_name"])
                     for edge in cli_reports["deep"]["edges"]}
        api_edges = {(edge.consumer_path, edge.target_path, edge.source_kind, edge.source_name)
                     for edge in ui_report.edges}
        if cli_edges != api_edges:
            raise AssertionError("Packaged CLI and in-process scan results differ.")
        result["cli_api_parity"] = True
        result["gui"] = _exercise_dialog(ui_report, files)
        quick_formula_edges = [edge for edge in cli_reports["quick"]["edges"]
                               if edge["source_kind"] == "Excel formula"]
        deep_formula_edges = [edge for edge in cli_reports["deep"]["edges"]
                              if edge["source_kind"] == "Excel formula"]
        quick_edge_ids = {(edge["consumer_path"], edge["target_path"], edge["source_kind"], edge["source_name"])
                          for edge in cli_reports["quick"]["edges"]}
        deep_edge_ids = {(edge["consumer_path"], edge["target_path"], edge["source_kind"], edge["source_name"])
                         for edge in cli_reports["deep"]["edges"]}
        if quick_edge_ids - deep_edge_ids:
            raise AssertionError("Deep scan lost a dependency detected by Quick scan.")
        if not any(edge["consumer_path"].endswith("LargeExternalFormulas.xlsx")
                   and edge["target_path"].endswith("ExternalTarget.xlsx")
                   for edge in cli_reports["quick"]["edges"]):
            raise AssertionError("Quick scan missed the large workbook's genuine external workbook link.")
        if not any(edge["consumer_path"].endswith("LargeExternalFormulas.xlsx")
                   and edge["target_path"].endswith("ExternalTarget.xlsx")
                   for edge in cli_reports["deep"]["edges"]):
            raise AssertionError("Deep scan missed the large workbook's genuine external workbook link.")
        result["quick_deep_formula_coverage"] = {
            "quick_excel_formula_edges": len(quick_formula_edges),
            "deep_excel_formula_edges": len(deep_formula_edges),
            "quick_dependency_edge_count": len(quick_edge_ids),
            "deep_dependency_edge_count": len(deep_edge_ids),
            "deep_only_edges": len(deep_edge_ids - quick_edge_ids),
            "quick_only_edges": len(quick_edge_ids - deep_edge_ids),
            "large_formula_cells": large_rows * 26,
        }
        result["gui_cancel_progress"] = _gui_cancel_progress(files)
    if not keep:
        shutil.rmtree(run_dir)
        result["fixture_cleanup"] = not run_dir.exists()
    else:
        result["fixture_directory"] = str(run_dir)
    return result


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--cli-exe", type=Path, default=ROOT / "dist" / "PowerQueryManager" / "PowerQueryManagerCLI.exe")
    parser.add_argument("--report", type=Path, help="Write the final JSON evidence to this new path.")
    parser.add_argument("--mapped-root", type=Path, help="Existing mapped SMB drive/folder to exercise.")
    parser.add_argument("--unc-root", type=Path, help="Existing UNC share path to exercise independently.")
    parser.add_argument("--sync-root", type=Path, help="Existing synchronized test folder; no sync settings are changed.")
    parser.add_argument("--dao-progid", default="DAO.DBEngine.160",
                        help="DAO COM provider used to create fixtures (default requires ACE DAO 16.0).")
    parser.add_argument("--large-rows", type=int, default=1200,
                        help="Rows of 26 external-link formulas used for the Quick/Deep performance comparison.")
    parser.add_argument("--keep-fixtures", action="store_true", help="Keep disposable input folders for inspection.")
    args = parser.parse_args()
    if os.name != "nt":
        parser.error("This acceptance harness requires Windows, desktop Excel, and ACE DAO 16.0.")
    if not args.cli_exe.is_file():
        parser.error(f"Packaged CLI not found: {args.cli_exe}")
    if args.large_rows < 1 or args.large_rows > 20000:
        parser.error("--large-rows must be in 1..20000.")

    import pythoncom
    pythoncom.CoInitialize()
    baseline_excel = _excel_pids()
    pythoncom.CoUninitialize()
    temp_root = Path(tempfile.mkdtemp(prefix="pqm-dependency-acceptance-"))
    fixture_root = temp_root / "seed"
    output = {"schema_version": 1, "status": "FAIL", "source_commit": None,
              "host": {"platform": sys.platform, "python": sys.version.split()[0]},
              "fixture_counts": {}, "locations": [], "baseline_excel_pids": sorted(baseline_excel)}
    try:
        output["fixture_creation"] = _create_fixtures(fixture_root, large_rows=args.large_rows,
                                                       dao_progid=args.dao_progid)
        output["dao_provider_requested"] = args.dao_progid
        seed_files = list(fixture_root.iterdir())
        output["fixture_counts"] = {"files": len(seed_files),
                                     "access": sum(p.suffix.casefold() in {".accdb", ".mdb"} for p in seed_files),
                                     "excel": sum(p.suffix.casefold() in {".xlsx", ".xlsm"} for p in seed_files)}
        cli_version = subprocess.run([str(args.cli_exe), "--help"], check=True, capture_output=True,
                                     text=True, encoding="utf-8", errors="replace").stdout
        output["packaged_cli_exposes_features"] = all(name in cli_version for name in ("dependency-scan", "dependency-map"))
        locations = [("local", temp_root)]
        for label, root in (("mapped_smb", args.mapped_root), ("unc", args.unc_root), ("synchronized", args.sync_root)):
            if root is not None:
                if not root.is_dir():
                    raise ValueError(f"The supplied {label} root is not an accessible directory.")
                locations.append((label, root))
        for label, root in locations:
            if label == "local":
                # Fixture was already built here; test directly in a dedicated copy.
                source = fixture_root
                local_root = temp_root / "local-scans"
                local_root.mkdir()
            else:
                source = fixture_root
                local_root = root
            output["locations"].append(_run_location(label, source, local_root, args.cli_exe,
                                                       keep=args.keep_fixtures, large_rows=args.large_rows))
        output["status"] = ("FAIL" if output["fixture_creation"].get("excel_forced_termination")
                            or any(not row.get("excel_pids_unchanged", False) for row in output["locations"])
                            else "PASS")
        if output["status"] == "FAIL":
            output["failure"] = {"type": "CleanupAcceptanceFailure",
                                 "detail": "Fixture-owned Excel required forced termination; cleanup gate is not a pass."}
    except Exception as exc:
        output["failure"] = {"type": type(exc).__name__, "detail": str(exc)}
    finally:
        # Seed Excel is explicitly closed by the fixture creator. Confirm it did not leak.
        try:
            output["final_excel_pids"] = sorted(_excel_pids())
            output["excel_process_set_restored"] = output["final_excel_pids"] == sorted(baseline_excel)
        except Exception as exc:
            output["process_check_error"] = type(exc).__name__
        if not args.keep_fixtures:
            shutil.rmtree(temp_root, ignore_errors=True)
            output["fixture_cleanup"] = not temp_root.exists()
        output["source_commit"] = subprocess.run(
            ["git", "rev-parse", "HEAD"], cwd=ROOT, check=True, capture_output=True, text=True
        ).stdout.strip()
        rendered = json.dumps(output, indent=2)
        if args.report:
            with args.report.open("x", encoding="utf-8", newline="\n") as stream:
                stream.write(rendered + "\n")
        print(rendered)
    return 0 if output["status"] == "PASS" and output.get("excel_process_set_restored") else 1


if __name__ == "__main__":
    raise SystemExit(main())
