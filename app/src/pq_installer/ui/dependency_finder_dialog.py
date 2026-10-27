"""Read-only reverse dependency workspace, kept separate from query import."""
from __future__ import annotations

import csv
from dataclasses import asdict
import json
from pathlib import Path
from threading import Event

from PySide6.QtCore import QObject, QThread, Signal, Slot
from PySide6.QtWidgets import (
    QAbstractItemView, QCheckBox, QComboBox, QDialog, QFileDialog, QGridLayout,
    QHBoxLayout, QLabel, QListWidget, QListWidgetItem, QMessageBox, QPlainTextEdit,
    QProgressBar, QPushButton, QTableWidget, QTableWidgetItem, QVBoxLayout,
    QFrame, QHeaderView,
)

from pq_installer.core.reverse_dependencies import (
    DependencyHit, DependencyMapReport, discover_scan_files, filter_dependency_map,
    scan_dependency_map, scan_reverse_dependencies,
)
from pq_installer.ui.file_dependency_graph_dialog import FileDependencyGraphDialog
from pq_installer.ui import apply_vector_theme


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

    def __init__(self, targets, consumers, scan_runner=None, *, mode="reverse", depth="deep", map_runner=None):
        super().__init__()
        self.targets = tuple(targets)
        self.consumers = tuple(consumers)
        self.scan_runner = scan_runner or scan_reverse_dependencies
        self.map_runner = map_runner or scan_dependency_map
        self.mode = mode
        self.depth = depth
        self._cancel_event = Event()

    def cancel(self):
        self._cancel_event.set()

    @Slot()
    def run(self):
        try:
            if self.mode == "map":
                result = self.map_runner(self.consumers, depth=self.depth,
                    progress=self.progressed.emit, cancelled=self._cancel_event.is_set)
            else:
                result = self.scan_runner(self.targets, self.consumers, depth=self.depth,
                    progress=self.progressed.emit, cancelled=self._cancel_event.is_set)
            self.completed.emit(result)
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
    def __init__(self, parent=None, *, scan_runner=None, map_runner=None):
        super().__init__(parent)
        self.setWindowTitle("Dependency Finder — incoming references")
        self.resize(1180, 800)
        self.setMinimumSize(760, 560)
        self.scan_runner = scan_runner
        self.map_runner = map_runner
        self.latest_report = None
        self._scan_thread = None
        self._scan_worker = None
        self._running = False
        self._inventory_thread = None
        self._inventory_worker = None
        self._inventory_running = False
        self._build_ui()
        apply_vector_theme(self, getattr(parent, "dark_theme_enabled", True))

    def _build_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(22, 16, 22, 18)
        layout.setSpacing(12)
        eyebrow = QLabel("DEPENDENCY INTELLIGENCE    /    READ-ONLY")
        eyebrow.setObjectName("vectorEyebrow")
        layout.addWidget(eyebrow)
        title = QLabel("TRACE FILE RELATIONSHIPS")
        title.setObjectName("vectorHero")
        layout.addWidget(title)
        intro = QLabel("Map which Excel and Access files reference a target, or inspect the entire file-to-file network. "
                       "Only definitions and connection metadata are read; nothing is refreshed or changed.")
        intro.setWordWrap(True)
        intro.setObjectName("secondaryText")
        layout.addWidget(intro)

        mode_row = QHBoxLayout()
        mode_row.addWidget(QLabel("Mode"))
        self.mode_combo = QComboBox()
        self.mode_combo.setAccessibleName("Dependency scanning mode")
        self.mode_combo.addItem("Find dependents of selected target(s)", "reverse")
        self.mode_combo.addItem("Map all selected files and their dependencies", "map")
        self.mode_combo.currentIndexChanged.connect(self._on_mode_changed)
        mode_row.addWidget(self.mode_combo, 2)
        mode_row.addWidget(QLabel("Scan depth"))
        self.scan_depth_combo = QComboBox()
        self.scan_depth_combo.setAccessibleName("Quick or deep dependency scanning")
        self.scan_depth_combo.addItem("Quick — queries, connections and links", "quick")
        self.scan_depth_combo.addItem("Deep — also worksheet formulas and names", "deep")
        self.scan_depth_combo.setCurrentIndex(1)  # retain previous reverse scan scope
        mode_row.addWidget(self.scan_depth_combo, 2)
        layout.addLayout(mode_row)

        selection_row = QHBoxLayout()
        selection_row.setSpacing(12)
        consumer_panel = QFrame()
        consumer_panel.setObjectName("vectorListPanel")
        consumer_layout = QVBoxLayout(consumer_panel)
        consumer_layout.setContentsMargins(14, 14, 14, 14)
        consumer_layout.setSpacing(10)
        consumer_heading = QLabel("01  /  FILES TO INSPECT")
        consumer_heading.setObjectName("vectorEyebrow")
        consumer_layout.addWidget(consumer_heading)
        self.consumer_count_label = QLabel("Possible dependency consumers")
        self.consumer_count_label.setObjectName("vectorHint")
        consumer_layout.addWidget(self.consumer_count_label)
        self.consumer_list = QListWidget()
        self.consumer_list.setAccessibleName("Files to inspect for incoming dependency references")
        self.consumer_list.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.consumer_list.itemSelectionChanged.connect(self._update_enabled)
        self.consumer_list.setMinimumHeight(150)
        consumer_layout.addWidget(self.consumer_list, 1)

        source_actions = QGridLayout()
        source_actions.setHorizontalSpacing(6)
        source_actions.setVerticalSpacing(6)
        self.add_files_button = QPushButton("Add files…")
        self.add_files_button.clicked.connect(self._browse_scan_files)
        self.add_folder_button = QPushButton("Add folder…")
        self.add_folder_button.clicked.connect(self._browse_scan_folder)
        self.recursive_check = QCheckBox("Include subfolders")
        self.recursive_check.setChecked(True)
        self.select_all_button = QPushButton("Select all")
        self.select_all_button.clicked.connect(self.consumer_list.selectAll)
        for index, widget in enumerate((self.add_files_button, self.add_folder_button,
                                         self.recursive_check, self.select_all_button)):
            source_actions.addWidget(widget, index // 2, index % 2)
        source_actions.setColumnStretch(0, 1)
        source_actions.setColumnStretch(1, 1)
        consumer_layout.addLayout(source_actions)
        selection_row.addWidget(consumer_panel, 1)

        target_panel = QFrame()
        target_panel.setObjectName("vectorListPanel")
        target_layout = QVBoxLayout(target_panel)
        target_layout.setContentsMargins(14, 14, 14, 14)
        target_layout.setSpacing(10)
        target_heading = QLabel("02  /  DEPENDENCY TARGETS")
        target_heading.setObjectName("vectorEyebrow")
        target_layout.addWidget(target_heading)
        self.target_count_label = QLabel("Which files should be traced?")
        self.target_count_label.setObjectName("vectorHint")
        target_layout.addWidget(self.target_count_label)
        self.target_list = QListWidget()
        self.target_list.setAccessibleName("Files whose dependents you want to find")
        self.target_list.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.target_list.itemSelectionChanged.connect(self._update_enabled)
        self.target_list.setMinimumHeight(150)
        target_layout.addWidget(self.target_list, 1)
        target_actions = QGridLayout()
        target_actions.setHorizontalSpacing(6)
        target_actions.setVerticalSpacing(6)
        self.add_targets_button = QPushButton("Add target file(s)…")
        self.add_targets_button.clicked.connect(self._browse_targets)
        self.from_consumers_button = QPushButton("Use selected files as targets")
        self.from_consumers_button.clicked.connect(self._copy_selected_targets)
        self.remove_targets_button = QPushButton("Remove selected targets")
        self.remove_targets_button.clicked.connect(self._remove_targets)
        for index, widget in enumerate((self.add_targets_button, self.from_consumers_button,
                                         self.remove_targets_button)):
            target_actions.addWidget(widget, index // 2, index % 2)
        target_actions.setColumnStretch(0, 1)
        target_actions.setColumnStretch(1, 1)
        target_layout.addLayout(target_actions)
        selection_row.addWidget(target_panel, 1)
        layout.addLayout(selection_row, 2)

        run_row = QHBoxLayout()
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        self.progress_label = QLabel("Ready. Choose targets and files to inspect.")
        self.progress_label.setWordWrap(True)
        self.run_button = QPushButton("Run read-only dependency scan")
        self.run_button.setObjectName("vectorPrimary")
        self.run_button.clicked.connect(self._start_scan)
        self.cancel_button = QPushButton("Cancel scan")
        self.cancel_button.clicked.connect(self._cancel_scan)
        self.cancel_button.setEnabled(False)
        run_row.addWidget(self.progress_label, 2)
        run_row.addWidget(self.progress_bar, 1)
        run_row.addWidget(self.run_button)
        run_row.addWidget(self.cancel_button)
        layout.addLayout(run_row)

        map_filter_row = QHBoxLayout()
        self.map_filter_label = QLabel("Explore file")
        self.map_filter_combo = QComboBox()
        self.map_filter_combo.setAccessibleName("Filter dependency map by file")
        self.map_filter_combo.addItem("All mapped files", None)
        self.direction_combo = QComboBox()
        self.direction_combo.setAccessibleName("Incoming or outgoing dependencies")
        self.direction_combo.addItem("Both directions", "both")
        self.direction_combo.addItem("Files depending on selected file", "incoming")
        self.direction_combo.addItem("Files selected file depends on", "outgoing")
        self.graph_button = QPushButton("Visual file graph…")
        self.graph_button.setAccessibleName("Show visual file dependency graph")
        self.graph_button.clicked.connect(self._show_file_graph)
        self.map_filter_combo.currentIndexChanged.connect(self._render_map_results)
        self.direction_combo.currentIndexChanged.connect(self._render_map_results)
        for widget in (self.map_filter_label, self.map_filter_combo, self.direction_combo, self.graph_button):
            map_filter_row.addWidget(widget)
        layout.addLayout(map_filter_row)

        result_heading = QLabel("03  /  DETECTED RELATIONSHIPS")
        result_heading.setObjectName("vectorEyebrow")
        layout.addWidget(result_heading)
        self.results_table = QTableWidget(0, 5)
        self.results_table.setAccessibleName("Reverse dependency results")
        self.results_table.setHorizontalHeaderLabels((
            "Target", "File depending on target", "Reference type", "Query / connection", "Match confidence"
        ))
        self.results_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.results_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.results_table.setAlternatingRowColors(True)
        self.results_table.verticalHeader().setVisible(False)
        self.results_table.setMinimumHeight(185)
        header = self.results_table.horizontalHeader()
        header.setMinimumSectionSize(100)
        header.setSectionResizeMode(QHeaderView.Interactive)
        for column, width in enumerate((175, 225, 155, 195, 150)):
            self.results_table.setColumnWidth(column, width)
        header.setStretchLastSection(True)
        layout.addWidget(self.results_table, 3)
        warning_heading = QLabel("COVERAGE  /  LIMITATIONS & WARNINGS")
        warning_heading.setObjectName("vectorEyebrow")
        layout.addWidget(warning_heading)
        self.coverage_text = QPlainTextEdit()
        self.coverage_text.setReadOnly(True)
        self.coverage_text.setMaximumHeight(115)
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
        self._on_mode_changed()
        self._update_enabled()

    def _on_mode_changed(self, *_):
        mapping = self.mode_combo.currentData() == "map"
        self.scan_depth_combo.setCurrentIndex(0 if mapping else 1)
        for widget in (self.map_filter_label, self.map_filter_combo, self.direction_combo, self.graph_button):
            widget.setVisible(mapping)
        self.run_button.setText("Build read-only dependency map" if mapping else "Run read-only dependency scan")
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
        mapping = self.mode_combo.currentData() == "map"
        ready = bool(self.consumer_list.selectedItems()) and (mapping or self.target_list.count() > 0)
        busy = self._running or self._inventory_running
        self.run_button.setEnabled(ready and not busy)
        for button in (self.add_files_button, self.add_folder_button, self.select_all_button):
            button.setEnabled(not busy)
        for button in (self.add_targets_button, self.from_consumers_button, self.remove_targets_button):
            button.setEnabled(not busy and not mapping)
        self.mode_combo.setEnabled(not busy)
        self.scan_depth_combo.setEnabled(not busy)
        self.recursive_check.setEnabled(not busy)
        self.consumer_list.setEnabled(not busy)
        self.target_list.setEnabled(not busy and not mapping)
        is_map = isinstance(self.latest_report, DependencyMapReport)
        self.map_filter_combo.setEnabled(is_map and not busy)
        self.direction_combo.setEnabled(is_map and not busy)
        self.graph_button.setEnabled(is_map and not busy)
        self.consumer_count_label.setText(f"{self.consumer_list.count()} candidate file(s)  ·  {len(self.consumer_list.selectedItems())} selected")
        self.target_count_label.setText(f"{self.target_list.count()} selected target(s)" if not mapping else "General map indexes every inspected file")
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
        self._scan_worker = DependencyScanWorker(targets, consumers, self.scan_runner,
            mode=self.mode_combo.currentData(), depth=self.scan_depth_combo.currentData(),
            map_runner=self.map_runner)
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
        if isinstance(report, DependencyMapReport):
            self.map_filter_combo.blockSignals(True)
            self.map_filter_combo.clear()
            self.map_filter_combo.addItem("All mapped files", None)
            for path in report.files:
                self.map_filter_combo.addItem(str(path), str(path))
            self.map_filter_combo.blockSignals(False)
            self._render_map_results()
            total = len(report.edges)
            lines = [f"Mapped {len(report.inspected)} file(s); {total} internal dependency edge(s); "
                     f"{len(report.unresolved)} outside-scope or unresolvable file reference(s)."]
            for item in report.unresolved[:20]:
                lines.append(f"{Path(item.consumer_path).name}: unresolved reference to {item.reference_filename}")
            if len(report.unresolved) > 20:
                lines.append(f"…and {len(report.unresolved) - 20} further unresolved references in the exported index.")
        else:
            self._fill_results_table(report.hits)
            lines = [f"Attempted {len(report.inspected)} file(s); found {len(report.hits)} dependency reference(s)."]
            if not report.hits:
                lines.append("No references found in inspected definitions; this is not proof that none exist.")
        self._update_enabled()
        if report.cancelled:
            lines.append("CANCELLED: Remaining files were not inspected.")
        for issue in report.issues[:40]:
            label = Path(issue.path).name if issue.path else "Scan scope"
            lines.append(f"{label}: {issue.status} — {issue.detail}")
        if len(report.issues) > 40:
            lines.append(f"…plus {len(report.issues) - 40} additional issues in the JSON/CSV export.")
        lines.append("Dynamic VBA, macros and external scripts may contain additional dependencies.")
        self.coverage_text.setPlainText("\n".join(lines))
        self.progress_label.setText("Read-only scan finished" + (" (partial)" if report.issues or report.cancelled else ""))

    def _fill_results_table(self, hits):
        self.results_table.setUpdatesEnabled(False)
        self.results_table.setRowCount(len(hits))
        for index, hit in enumerate(hits):
            values = (hit.target_path, hit.consumer_path, hit.source_kind, hit.source_name, hit.confidence)
            for column, value in enumerate(values):
                self.results_table.setItem(index, column, QTableWidgetItem(str(value)))
        self.results_table.setUpdatesEnabled(True)
        if len(hits) < 400:
            self.results_table.resizeColumnsToContents()
        else:
            for column in range(4):
                self.results_table.setColumnWidth(column, 185)

    def _render_map_results(self, *_):
        if isinstance(self.latest_report, DependencyMapReport):
            selected = self.map_filter_combo.currentData()
            direction = self.direction_combo.currentData() or "both"
            self._fill_results_table(filter_dependency_map(self.latest_report, selected, direction=direction))

    def _show_file_graph(self):
        if isinstance(self.latest_report, DependencyMapReport):
            FileDependencyGraphDialog(self.latest_report, self).exec()

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
                    kind_name = "dependency-map" if isinstance(self.latest_report, DependencyMapReport) else "reverse-dependencies"
                    json.dump({"schema_version": 1, "bundle_type": kind_name, **asdict(self.latest_report)}, output, indent=2)
                    output.write("\n")
                else:
                    fields = ["row_type", *DependencyHit.__dataclass_fields__, "status", "detail"]
                    writer = csv.DictWriter(output, fieldnames=fields)
                    writer.writeheader()
                    is_map = isinstance(self.latest_report, DependencyMapReport)
                    hits = self.latest_report.edges if is_map else self.latest_report.hits
                    for hit in hits:
                        writer.writerow(_safe_csv_row({"row_type": "dependency", **asdict(hit)}))
                    for target in self.latest_report.files if is_map else self.latest_report.targets:
                        has_links = any(hit.target_path == target or (is_map and hit.consumer_path == target)
                                        for hit in hits)
                        writer.writerow(_safe_csv_row({"row_type": "scan_summary", "target_path": target,
                                         "status": ("links_found" if has_links else "no_links") if is_map else
                                                   ("matches_found" if has_links else "no_matches"),
                                         "detail": f"Attempted {len(self.latest_report.inspected)} consumer file(s)."}))
                    if is_map:
                        for item in self.latest_report.unresolved:
                            writer.writerow(_safe_csv_row({"row_type": "unresolved", "consumer_path": item.consumer_path,
                                             "source_kind": item.source_kind, "source_name": item.source_name,
                                             "evidence": item.reference_filename, "status": "outside_scope_or_unknown"}))
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
