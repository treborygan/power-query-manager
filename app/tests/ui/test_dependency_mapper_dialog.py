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


def test_map_results_keep_long_paths_without_horizontal_table_overflow(tmp_path):
    app = QApplication.instance() or QApplication([])
    root = tmp_path
    for index in range(4):
        root = root / f"source-root-{index}-with-a-long-disambiguating-name"
        root.mkdir()
    target = root / "Inventory.xlsx"
    consumer = root / "Dashboard.xlsx"
    target.write_bytes(b"fixture")
    consumer.write_bytes(b"fixture")
    report = DependencyMapReport((str(target), str(consumer)), (str(target), str(consumer)), (
        DependencyHit(str(target), str(consumer), "Excel connection", "Inventory", "exact_path", target.name),
    ), (), (), "quick")

    dialog = DependencyFinderDialog(map_runner=lambda *args, **kwargs: report)
    dialog.add_scan_paths((target, consumer))
    dialog.mode_combo.setCurrentIndex(dialog.mode_combo.findData("map"))
    dialog.resize(960, 600)
    dialog.show()
    app.processEvents()
    dialog.run_button.click()
    deadline = time.monotonic() + 3
    while dialog.latest_report is None and time.monotonic() < deadline:
        app.processEvents()
        time.sleep(.01)
    app.processEvents()

    target_item = dialog.results_table.item(0, 0)
    consumer_item = dialog.results_table.item(0, 1)
    assert target_item.text() == str(target)
    assert consumer_item.text() == str(consumer)
    assert target_item.toolTip() == str(target)
    assert consumer_item.toolTip() == str(consumer)
    assert dialog.results_table.horizontalScrollBar().maximum() == 0
    assert dialog.minimumSizeHint().width() <= 960
    assert dialog.minimumSizeHint().height() <= 600
    dialog.close()


def test_general_map_graph_has_accessible_visual_and_summary(tmp_path):
    app = QApplication.instance() or QApplication([])
    a = tmp_path / "Warehouse.accdb"
    b = tmp_path / "Dashboard.xlsx"
    report = DependencyMapReport((str(a), str(b)), (str(a), str(b)), (
        DependencyHit(str(a), str(b), "Power Query", "Load", "filename_only", a.name),
    ), (), (), "quick")

    dialog = FileDependencyGraphDialog(report)

    dialog.resize(960, 600)
    dialog.show()
    app.processEvents()
    app.processEvents()
    assert dialog.minimumSizeHint().width() <= 960
    assert dialog.minimumSizeHint().height() <= 600
    assert dialog.graph_view.width() > 0 and dialog.graph_view.height() > 0
    mapped_scene = dialog.graph_view.mapFromScene(dialog.graph_view.scene().sceneRect()).boundingRect()
    assert dialog.graph_view.viewport().rect().contains(mapped_scene)

    assert dialog.graph_view.scene().items()
    assert "Dashboard.xlsx" in dialog.summary.toPlainText()
    assert "Warehouse.accdb" in dialog.summary.toPlainText()
    assert "filename_only" in dialog.summary.toPlainText()
    assert any(isinstance(item, QGraphicsPolygonItem) for item in dialog.graph_view.scene().items())
    dialog.close()


def test_vector_graph_node_selection_and_zoom_keep_accessible_edge_index(tmp_path):
    from PySide6.QtCore import Qt
    from PySide6.QtTest import QTest
    from PySide6.QtWidgets import QGraphicsItem
    app = QApplication.instance() or QApplication([])
    consumer = tmp_path / "Report.xlsx"
    target = tmp_path / "Warehouse.accdb"
    report = DependencyMapReport((str(consumer), str(target)),
        (str(consumer), str(target)), (
            DependencyHit(str(target), str(consumer), "Power Query", "Source",
                          "exact_path", target.name),
        ), (), (), "quick")
    dialog = FileDependencyGraphDialog(report)
    dialog.show()
    app.processEvents()
    dialog.graph_view.setFocus()
    QTest.keyClick(dialog.graph_view, Qt.Key_Home)
    assert len(dialog.graph_view.scene().selectedItems()) == 1
    first = dialog.graph_view.scene().selectedItems()[0]
    QTest.keyClick(dialog.graph_view, Qt.Key_Right)
    assert len(dialog.graph_view.scene().selectedItems()) == 1
    assert dialog.graph_view.scene().selectedItems()[0] is not first
    nodes = [item for item in dialog.graph_view.scene().items()
             if item.data(0) == str(consumer)]
    assert len(nodes) == 1
    assert nodes[0].flags() & QGraphicsItem.ItemIsSelectable
    dialog.graph_view.scene().clearSelection()
    nodes[0].setSelected(True)
    app.processEvents()
    assert "Report.xlsx" in dialog.selection_status.text()
    assert "depends on" in dialog.summary.toPlainText()
    dialog.graph_view.scale(1.2, 1.2)
    assert dialog.graph_view.transform().m11() > 1
    dialog._reset_zoom()
    assert abs(dialog.graph_view.transform().m11() - 1) < 0.01
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
