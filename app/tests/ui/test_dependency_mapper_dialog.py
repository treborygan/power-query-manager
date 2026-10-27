import time
import csv

from PySide6.QtWidgets import QApplication, QFileDialog, QGraphicsPolygonItem, QMessageBox

from pq_installer.core.reverse_dependencies import DependencyHit, DependencyMapReport
from pq_installer.ui.dependency_finder_dialog import DependencyFinderDialog
from pq_installer.ui.file_dependency_graph_dialog import FileDependencyGraphDialog


def test_map_mode_discovers_all_edges_and_filters_without_rescan(tmp_path):
    app = QApplication.instance() or QApplication([])
    warehouse = tmp_path / "Warehouse.accdb"
    inventory = tmp_path / "Inventory.xlsx"
    dashboard = tmp_path / "Dashboard.xlsx"
    for path in (warehouse, inventory, dashboard):
        path.write_bytes(b"fixture")
    calls = []

    def map_runner(files, *, depth, progress, cancelled):
        calls.append((tuple(files), depth))
        return DependencyMapReport(tuple(map(str, files)), tuple(map(str, files)), (
            DependencyHit(str(warehouse), str(inventory), "Power Query", "Inventory", "exact_path", warehouse.name),
            DependencyHit(str(inventory), str(dashboard), "Excel connection", "Source", "exact_path", inventory.name),
        ), (), (), depth)

    dialog = DependencyFinderDialog(map_runner=map_runner)
    dialog.add_scan_paths((warehouse, inventory, dashboard))
    dialog.mode_combo.setCurrentIndex(dialog.mode_combo.findData("map"))
    assert dialog.run_button.isEnabled()  # no manually selected targets needed
    assert dialog.scan_depth_combo.currentData() == "quick"
    dialog.run_button.click()
    deadline = time.monotonic() + 3
    while dialog.latest_report is None and time.monotonic() < deadline:
        app.processEvents()
        time.sleep(.01)

    assert len(calls) == 1
    assert dialog.results_table.rowCount() == 2
    dialog.map_filter_combo.setCurrentIndex(dialog.map_filter_combo.findData(str(warehouse)))
    dialog.direction_combo.setCurrentIndex(dialog.direction_combo.findData("incoming"))
    assert dialog.results_table.rowCount() == 1
    assert dialog.results_table.item(0, 1).text() == str(inventory)
    dialog.map_filter_combo.setCurrentIndex(dialog.map_filter_combo.findData(str(inventory)))
    dialog.direction_combo.setCurrentIndex(dialog.direction_combo.findData("outgoing"))
    assert dialog.results_table.rowCount() == 1
    assert dialog.results_table.item(0, 0).text() == str(warehouse)
    assert len(calls) == 1  # in-memory graph view, no file re-read
    deadline = time.monotonic() + 3
    while dialog._running and time.monotonic() < deadline:
        app.processEvents();time.sleep(.01)
    dialog.close()


def test_general_map_graph_has_accessible_visual_and_summary(tmp_path):
    app = QApplication.instance() or QApplication([])
    a = tmp_path / "Warehouse.accdb"
    b = tmp_path / "Dashboard.xlsx"
    report = DependencyMapReport((str(a), str(b)), (str(a), str(b)), (
        DependencyHit(str(a), str(b), "Power Query", "Load", "filename_only", a.name),
    ), (), (), "quick")

    dialog = FileDependencyGraphDialog(report)

    assert dialog.graph_view.scene().items()
    assert "Dashboard.xlsx" in dialog.summary.toPlainText()
    assert "Warehouse.accdb" in dialog.summary.toPlainText()
    assert "filename_only" in dialog.summary.toPlainText()
    assert any(isinstance(item, QGraphicsPolygonItem) for item in dialog.graph_view.scene().items())
    dialog.close()


def test_general_map_csv_summary_considers_outgoing_links_and_unresolved(tmp_path, monkeypatch):
    app = QApplication.instance() or QApplication([])
    warehouse = tmp_path / "Warehouse.accdb"
    dashboard = tmp_path / "Dashboard.xlsx"
    report = DependencyMapReport((str(warehouse), str(dashboard)),
        (str(warehouse), str(dashboard)), (
            DependencyHit(str(warehouse), str(dashboard), "Power Query", "Load", "exact_path", warehouse.name),
        ), (), (), "quick")
    destination = tmp_path / "map.csv"
    monkeypatch.setattr(QFileDialog, "getSaveFileName", lambda *a, **kw: (str(destination), "CSV"))
    monkeypatch.setattr(QMessageBox, "information", lambda *a, **kw: None)
    dialog = DependencyFinderDialog()
    dialog.latest_report = report
    dialog._export("csv")

    with destination.open(encoding="utf-8", newline="") as file:
        rows = list(csv.DictReader(file))
    summaries = {row["target_path"]: row["status"] for row in rows if row["row_type"] == "scan_summary"}
    assert summaries == {str(warehouse): "links_found", str(dashboard): "links_found"}
    dialog.close()
