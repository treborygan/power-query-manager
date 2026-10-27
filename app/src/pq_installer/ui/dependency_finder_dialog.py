"""Read-only reverse dependency workspace, kept separate from query import."""
from __future__ import annotations

import csv
from dataclasses import asdict
import json
from pathlib import Path
from threading import Event

from PySide6.QtCore import QObject, QThread, Signal, Slot
from PySide6.QtWidgets import (
    QAbstractItemView, QCheckBox, QDialog, QFileDialog, QGridLayout,
    QHBoxLayout, QLabel, QListWidget, QListWidgetItem, QMessageBox, QPlainTextEdit,
    QProgressBar, QPushButton, QTableWidget, QTableWidgetItem, QVBoxLayout,
)

from pq_installer.core.reverse_dependencies import DependencyHit, discover_scan_files, scan_reverse_dependencies


_FILE_FILTER = "Office files (*.xlsx *.xlsm *.accdb *.mdb *.xlsb *.xls)"


def _safe_csv_row(values):
    """Prevent Excel from treating user-controlled labels as CSV formulas."""
    def protect(value):
        if isinstance(value, str) and value.lstrip(" \t\r\n").startswith(("=", "+", "-", "@")):
            return "'" + value
        return value
    return {key: protect(value) for key, value in values.items()}


class DependencyScanWorker(QObject):
    progressed = Signal(int, int, str)
    completed = Signal(object)
    failed = Signal(str)
    finished = Signal()

    def __init__(self, targets, consumers, scan_runner=None):
        super().__init__()
        self.targets = tuple(targets)
        self.consumers = tuple(consumers)
        self.scan_runner = scan_runner or scan_reverse_dependencies
        self._cancel_event = Event()

    def cancel(self):
        self._cancel_event.set()

    @Slot()
    def run(self):
        try:
            self.completed.emit(self.scan_runner(
                self.targets, self.consumers, progress=self.progressed.emit,
                cancelled=self._cancel_event.is_set,
            ))
        except Exception:
            # Raw M, SQL and connection strings could contain passwords.
            self.failed.emit("Dependency scanning failed. Inspect the selected paths and permissions.")
        finally:
            self.finished.emit()


class InventoryWorker(QObject):
    completed = Signal(object)
    failed = Signal(str)
    finished = Signal()

    def __init__(self, paths, recursive):
        super().__init__()
        self.paths = tuple(paths)
        self.recursive = recursive
        self._cancel_event = Event()

    def cancel(self):
        self._cancel_event.set()

    @Slot()
    def run(self):
        try:
            paths = discover_scan_files(self.paths, recursive=self.recursive,
                                        cancelled=self._cancel_event.is_set)
            self.completed.emit(() if self._cancel_event.is_set() else paths)
        except Exception:
            self.failed.emit("Folder discovery failed; check permissions and connection to the mapped folder.")
        finally:
            self.finished.emit()


class DependencyFinderDialog(QDialog):
    def __init__(self, parent=None, *, scan_runner=None):
        super().__init__(parent)
        self.setWindowTitle("Dependency Finder — incoming references")
        self.resize(1050, 740)
        self.setMinimumSize(690, 540)
        self.scan_runner = scan_runner
        self.latest_report = None
        self._scan_thread = None
        self._scan_worker = None
        self._running = False
        self._inventory_thread = None
        self._inventory_worker = None
        self._inventory_running = False
        self._build_ui()

    def _build_ui(self):
        layout = QVBoxLayout(self)
        intro = QLabel("Find which Excel or Access files depend on your selected file(s). "
                       "Only definitions and connection metadata are inspected; nothing is refreshed or changed.")
        intro.setWordWrap(True)
        layout.addWidget(intro)

        selection = QGridLayout()
        selection.addWidget(QLabel("1. Select files to inspect (possible consumers)"), 0, 0)
        selection.addWidget(QLabel("2. Select dependency targets (e.g. Warehouse.accdb)"), 0, 1)
        self.consumer_list = QListWidget()
        self.consumer_list.setAccessibleName("Files to inspect for incoming dependency references")
        self.consumer_list.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.consumer_list.itemSelectionChanged.connect(self._update_enabled)
        self.target_list = QListWidget()
        self.target_list.setAccessibleName("Files whose dependents you want to find")
        self.target_list.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.target_list.itemSelectionChanged.connect(self._update_enabled)
        selection.addWidget(self.consumer_list, 1, 0)
        selection.addWidget(self.target_list, 1, 1)
        source_actions = QHBoxLayout()
        self.add_files_button = QPushButton("Add files…")
        self.add_files_button.clicked.connect(self._browse_scan_files)
        self.add_folder_button = QPushButton("Add folder…")
        self.add_folder_button.clicked.connect(self._browse_scan_folder)
        self.recursive_check = QCheckBox("Include subfolders")
        self.recursive_check.setChecked(True)
        self.select_all_button = QPushButton("Select all")
        self.select_all_button.clicked.connect(self.consumer_list.selectAll)
        for widget in (self.add_files_button, self.add_folder_button, self.recursive_check, self.select_all_button):
            source_actions.addWidget(widget)
        selection.addLayout(source_actions, 2, 0)
        target_actions = QHBoxLayout()
        self.add_targets_button = QPushButton("Add target file(s)…")
        self.add_targets_button.clicked.connect(self._browse_targets)
        self.from_consumers_button = QPushButton("Use selected files as targets")
        self.from_consumers_button.clicked.connect(self._copy_selected_targets)
        self.remove_targets_button = QPushButton("Remove selected targets")
        self.remove_targets_button.clicked.connect(self._remove_targets)
        for widget in (self.add_targets_button, self.from_consumers_button, self.remove_targets_button):
            target_actions.addWidget(widget)
        selection.addLayout(target_actions, 2, 1)
        layout.addLayout(selection, 3)

        run_row = QHBoxLayout()
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        self.progress_label = QLabel("Ready. Choose targets and files to inspect.")
        self.progress_label.setWordWrap(True)
        self.run_button = QPushButton("Run read-only dependency scan")
        self.run_button.clicked.connect(self._start_scan)
        self.cancel_button = QPushButton("Cancel scan")
        self.cancel_button.clicked.connect(self._cancel_scan)
        self.cancel_button.setEnabled(False)
        run_row.addWidget(self.progress_label, 2)
        run_row.addWidget(self.progress_bar, 1)
        run_row.addWidget(self.run_button)
        run_row.addWidget(self.cancel_button)
        layout.addLayout(run_row)

        layout.addWidget(QLabel("3. Incoming dependencies (one row per source/query and target)"))
        self.results_table = QTableWidget(0, 5)
        self.results_table.setAccessibleName("Reverse dependency results")
        self.results_table.setHorizontalHeaderLabels((
            "Target", "File depending on target", "Reference type", "Query / connection", "Match confidence"
        ))
        self.results_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.results_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.results_table.horizontalHeader().setStretchLastSection(True)
        layout.addWidget(self.results_table, 3)
        layout.addWidget(QLabel("Coverage and limitations (check this before concluding a target has no dependents)"))
        self.coverage_text = QPlainTextEdit()
        self.coverage_text.setReadOnly(True)
        self.coverage_text.setMaximumHeight(110)
        self.coverage_text.setAccessibleName("Dependency scanner coverage and warnings")
        layout.addWidget(self.coverage_text)
        export_row = QHBoxLayout()
        self.export_json_button = QPushButton("Export JSON…")
        self.export_csv_button = QPushButton("Export CSV…")
        self.export_json_button.clicked.connect(lambda: self._export("json"))
        self.export_csv_button.clicked.connect(lambda: self._export("csv"))
        export_row.addStretch()
        export_row.addWidget(self.export_json_button)
        export_row.addWidget(self.export_csv_button)
        layout.addLayout(export_row)
        self._update_enabled()

    def add_scan_paths(self, paths, *, recursive=True):
        self._merge_scan_candidates(discover_scan_files(paths, recursive=recursive))

    def _merge_scan_candidates(self, paths):
        existing = {self.consumer_list.item(i).data(256) for i in range(self.consumer_list.count())}
        for path in paths:
            if str(path) not in existing:
                item = QListWidgetItem(str(path))
                item.setData(256, str(path))
                self.consumer_list.addItem(item)
                item.setSelected(True)
                existing.add(str(path))
        self._update_enabled()

    def add_target_paths(self, paths):
        existing = {self.target_list.item(i).data(256) for i in range(self.target_list.count())}
        for candidate in paths:
            path = Path(candidate).expanduser().resolve()
            if path.is_file() and str(path) not in existing:
                item = QListWidgetItem(str(path))
                item.setData(256, str(path))
                self.target_list.addItem(item)
                existing.add(str(path))
        self._update_enabled()

    def _browse_scan_files(self):
        paths, _ = QFileDialog.getOpenFileNames(self, "Select Excel/Access files to inspect", "", _FILE_FILTER)
        if paths:
            self._start_inventory(paths, recursive=False)

    def _browse_scan_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "Select folder to inventory")
        if folder:
            self._start_inventory((folder,), recursive=self.recursive_check.isChecked())

    def _start_inventory(self, paths, *, recursive):
        if self._running or self._inventory_running:
            return
        self._inventory_running = True
        self.progress_label.setText("Discovering Excel and Access files…")
        self.cancel_button.setText("Cancel discovery")
        self.cancel_button.setEnabled(True)
        self._update_enabled()
        self._inventory_thread = QThread(self)
        self._inventory_worker = InventoryWorker(paths, recursive)
        self._inventory_worker.moveToThread(self._inventory_thread)
        self._inventory_thread.started.connect(self._inventory_worker.run)
        self._inventory_worker.completed.connect(self._inventory_completed)
        self._inventory_worker.failed.connect(self._show_error)
        self._inventory_worker.finished.connect(self._inventory_thread.quit)
        self._inventory_worker.finished.connect(self._inventory_worker.deleteLater)
        self._inventory_thread.finished.connect(self._inventory_thread.deleteLater)
        self._inventory_thread.finished.connect(self._inventory_finished)
        self._inventory_thread.start()

    @Slot(object)
    def _inventory_completed(self, paths):
        self._merge_scan_candidates(paths)
        self.progress_label.setText(f"Discovered {self.consumer_list.count()} candidate file(s). Select files to inspect.")

    @Slot()
    def _inventory_finished(self):
        self._inventory_running = False
        self._inventory_thread = None
        self._inventory_worker = None
        self.cancel_button.setEnabled(False)
        self.cancel_button.setText("Cancel scan")
        self._update_enabled()

    def _browse_targets(self):
        paths, _ = QFileDialog.getOpenFileNames(self, "Select dependency target files", "", "All files (*)")
        self.add_target_paths(paths)

    def _copy_selected_targets(self):
        self.add_target_paths(item.data(256) for item in self.consumer_list.selectedItems())

    def _remove_targets(self):
        for item in self.target_list.selectedItems():
            self.target_list.takeItem(self.target_list.row(item))
        self._update_enabled()

    def _update_enabled(self):
        ready = bool(self.consumer_list.selectedItems()) and self.target_list.count() > 0
        busy = self._running or self._inventory_running
        self.run_button.setEnabled(ready and not busy)
        for button in (self.add_files_button, self.add_folder_button, self.add_targets_button,
                       self.from_consumers_button, self.remove_targets_button, self.select_all_button):
            button.setEnabled(not busy)
        self.recursive_check.setEnabled(not busy)
        self.consumer_list.setEnabled(not busy)
        self.target_list.setEnabled(not busy)
        self.export_json_button.setEnabled(self.latest_report is not None and not busy)
        self.export_csv_button.setEnabled(self.latest_report is not None and not busy)

    def _start_scan(self):
        if self._running or not self.run_button.isEnabled():
            return
        targets = tuple(Path(self.target_list.item(i).data(256)) for i in range(self.target_list.count()))
        consumers = tuple(Path(item.data(256)) for item in self.consumer_list.selectedItems())
        self.latest_report = None
        self.results_table.setRowCount(0)
        self.coverage_text.clear()
        self._running = True
        self.cancel_button.setText("Cancel scan")
        self.progress_bar.setRange(0, len(consumers))
        self.progress_bar.setValue(0)
        self.progress_bar.setVisible(True)
        self.progress_label.setText("Scanning selected files offline…")
        self.cancel_button.setEnabled(True)
        self._update_enabled()
        self._scan_thread = QThread(self)
        self._scan_worker = DependencyScanWorker(targets, consumers, self.scan_runner)
        self._scan_worker.moveToThread(self._scan_thread)
        self._scan_thread.started.connect(self._scan_worker.run)
        self._scan_worker.progressed.connect(self._on_progress)
        self._scan_worker.completed.connect(self._show_report)
        self._scan_worker.failed.connect(self._show_error)
        self._scan_worker.finished.connect(self._scan_thread.quit)
        self._scan_worker.finished.connect(self._scan_worker.deleteLater)
        self._scan_thread.finished.connect(self._scan_thread.deleteLater)
        self._scan_thread.finished.connect(self._thread_finished)
        self._scan_thread.start()

    @Slot(int, int, str)
    def _on_progress(self, completed, total, path):
        self.progress_bar.setRange(0, total)
        self.progress_bar.setValue(completed)
        self.progress_label.setText(f"Inspected {completed}/{total}: {Path(path).name}")

    @Slot(object)
    def _show_report(self, report):
        self.latest_report = report
        self.results_table.setRowCount(len(report.hits))
        for index, hit in enumerate(report.hits):
            values = (hit.target_path, hit.consumer_path, hit.source_kind, hit.source_name, hit.confidence)
            for column, value in enumerate(values):
                self.results_table.setItem(index, column, QTableWidgetItem(str(value)))
        self.results_table.resizeColumnsToContents()
        lines = [f"Attempted {len(report.inspected)} file(s); found {len(report.hits)} dependency reference(s)."]
        if report.cancelled:
            lines.append("CANCELLED: Remaining files were not inspected.")
        for issue in report.issues:
            label = Path(issue.path).name if issue.path else "Scan scope"
            lines.append(f"{label}: {issue.status} — {issue.detail}")
        if not report.hits:
            lines.append("No references were found in the inspected definitions; this is not proof that no dependencies exist.")
        lines.append("Dynamic VBA, macros and external scripts may contain additional dependencies.")
        self.coverage_text.setPlainText("\n".join(lines))
        self.progress_label.setText("Read-only scan finished" + (" (partial)" if report.issues or report.cancelled else ""))

    @Slot(str)
    def _show_error(self, message):
        self.coverage_text.setPlainText(message)
        self.progress_label.setText("Scan failed")

    @Slot()
    def _thread_finished(self):
        self._running = False
        self.cancel_button.setEnabled(False)
        self.progress_bar.setVisible(False)
        self._scan_thread = None
        self._scan_worker = None
        self._update_enabled()

    def _cancel_scan(self):
        if self._inventory_running and self._inventory_worker is not None:
            self._inventory_worker.cancel()
            self.cancel_button.setEnabled(False)
            self.progress_label.setText("Cancelling folder discovery…")
            return
        if self._running and self._scan_worker is not None:
            self._scan_worker.cancel()
            self.cancel_button.setEnabled(False)
            self.progress_label.setText("Cancelling after current file…")

    def reject(self):
        if self._running or self._inventory_running:
            self._cancel_scan()
            return
        super().reject()

    def closeEvent(self, event):
        if self._running or self._inventory_running:
            self._cancel_scan()
            event.ignore()
            return
        super().closeEvent(event)

    def _export(self, kind):
        if self.latest_report is None:
            return
        filename, _ = QFileDialog.getSaveFileName(
            self, "Export dependency report", f"dependency-report.{kind}",
            "JSON (*.json)" if kind == "json" else "CSV (*.csv)"
        )
        if not filename:
            return
        try:
            with open(filename, "x", encoding="utf-8", newline="") as output:
                if kind == "json":
                    json.dump({"schema_version": 1, **asdict(self.latest_report)}, output, indent=2)
                    output.write("\n")
                else:
                    fields = ["row_type", *DependencyHit.__dataclass_fields__, "status", "detail"]
                    writer = csv.DictWriter(output, fieldnames=fields)
                    writer.writeheader()
                    for hit in self.latest_report.hits:
                        writer.writerow(_safe_csv_row({"row_type": "dependency", **asdict(hit)}))
                    for target in self.latest_report.targets:
                        writer.writerow(_safe_csv_row({"row_type": "scan_summary", "target_path": target,
                                         "status": "matches_found" if any(hit.target_path == target for hit in self.latest_report.hits) else "no_matches",
                                         "detail": f"Attempted {len(self.latest_report.inspected)} consumer file(s)."}))
                    for issue in self.latest_report.issues:
                        writer.writerow(_safe_csv_row({"row_type": "coverage_issue", "consumer_path": issue.path,
                                         "status": issue.status, "detail": issue.detail}))
                    if self.latest_report.cancelled:
                        writer.writerow({"row_type": "coverage_issue", "status": "cancelled",
                                         "detail": "Scan was cancelled before all files could be inspected."})
            QMessageBox.information(self, "Report exported", "Dependency report saved to the selected new file.")
        except FileExistsError:
            QMessageBox.warning(self, "File exists", "Choose a new filename; existing reports will not be overwritten.")
        except OSError:
            QMessageBox.warning(self, "Export failed", "The report could not be saved at that location.")
