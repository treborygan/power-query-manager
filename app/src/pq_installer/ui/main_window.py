from pathlib import Path
import os
import shutil
import json

from PySide6.QtCore import Qt, QThread, QObject, Slot
from PySide6.QtGui import QFont, QKeySequence
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QGridLayout, QLabel,
    QLineEdit, QPushButton, QFileDialog, QTableView, QStyledItemDelegate,
    QHeaderView, QComboBox, QCheckBox, QPlainTextEdit, QMessageBox,
    QGroupBox, QSplitter, QAbstractItemView, QProgressBar, QDialog, QInputDialog, QSpinBox,
    QTableWidget, QTableWidgetItem, QDockWidget, QTabWidget, QScrollArea, QFrame
)

from pq_installer.core.models import ExistingQueryMode, LoadMode
from pq_installer.core.mode_parser import parse_existing_query_mode
from pq_installer.core.manager import (export_queries, export_query_package, build_import_profile, save_profile,
    load_profile, profile_to_sources, preview_profile_path_mappings,
    preview_profile_environment_mappings, materialize_profile_environment_mapping_copies, diff_m_code,
    compare_query_packages)
from pq_installer.core.excel_capabilities import ExcelCapabilities
from pq_installer.core.packages import extract_query_zip, apply_query_package_manifest
from pq_installer.core.history import HistoryStore
from pq_installer.core.error_log import record_failure
from pq_installer.core.query_loader import discover_query_files
from pq_installer.core.validation import validate_workbook, validate_query_sources, ordered_sources, order_warnings
from pq_installer.core.recovery import find_staged_workbooks, diagnose_staged_workbook
from pq_installer.ui.instructions import INSTRUCTIONS
from pq_installer.ui.install_worker import InstallWorker
from pq_installer.ui.query_table_model import QueryTableModel, QuerySortFilterProxyModel
from pq_installer.ui.inventory_model import QueryInventoryModel
from pq_installer.ui.inspect_worker import InspectWorker
from pq_installer.ui.inspect_worker import CapabilityProbeWorker
from pq_installer.ui.query_inspector_dialog import QueryInspectorDialog
from pq_installer.ui.source_table_view import SourceTableView
from pq_installer.ui.refresh_worker import RefreshWorker
from pq_installer.ui.history_worker import SnapshotRestoreWorker, SnapshotCompareWorker
from pq_installer.ui.history_comparison_dialog import HistoryComparisonDialog
from pq_installer.ui.dependency_graph_dialog import DependencyGraphDialog
from pq_installer.ui.dependency_finder_dialog import DependencyFinderDialog
from pq_installer.ui.package_update_dialog import QueryPackageUpdateDialog
from pq_installer.ui.profile_mappings_dialog import ProfileMappingsDialog
from pq_installer.ui.recovery_worker import StagedRecoveryWorker
from pq_installer.ui import apply_vector_theme


class _HistoryWorkerUiBridge(QObject):
    """Marshal history worker outcomes and completion to the GUI thread."""

    def __init__(self, completed, failed, finished, parent):
        super().__init__(parent)
        self._completed=completed
        self._failed=failed
        self._finished=finished

    @Slot(object)
    def completed(self, outcome):
        self._completed(outcome)

    @Slot(str)
    def failed(self, message):
        self._failed(message)

    @Slot()
    def thread_finished(self):
        self._finished()


class LoadModeDelegate(QStyledItemDelegate):
    def __init__(self,parent=None,supported_modes=('definition_only',)):
        super().__init__(parent);self.supported_modes=tuple(supported_modes)
    def createEditor(self, parent, option, index):
        editor = QComboBox(parent)
        labels={'definition_only':'Definition only (no connection or worksheet load)',
                'connection_only':'Connection only (no worksheet/Data Model load)',
                'new_sheet':'New worksheet table',
                'existing_sheet':'Existing worksheet table',
                'data_model':'Excel Data Model',
                'preserve_existing':'Preserve existing load'}
        for mode in self.supported_modes:editor.addItem(labels.get(mode,mode),mode)
        return editor
    def setEditorData(self, editor, index):
        editor.setCurrentIndex(max(0, editor.findData(index.data(Qt.EditRole))))
    def setModelData(self, editor, model, index):
        model.setData(index, editor.currentData(), Qt.EditRole)

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle('Power Query Manager')
        screen=QApplication.primaryScreen()
        if screen is not None:
            available=screen.availableGeometry()
            width=max(760,min(1180,available.width()-40))
            height=max(560,min(700,available.height()-40))
        else:
            width,height=1180,700
        self.setMinimumSize(min(960,width),min(600,height))
        self.resize(width,height)
        self.sources = []
        self.query_model = QueryTableModel(parent=self)
        self.query_proxy = QuerySortFilterProxyModel(self)
        self.query_proxy.setSourceModel(self.query_model)
        self._thread = None
        self._worker = None
        self._inspect_thread = None
        self._inspect_worker = None
        self._refresh_thread = None
        self._recovery_thread=None
        self._refresh_worker = None
        self._history_restore_thread=None
        self._history_restore_worker=None
        self._history_restore_bridge=None
        self._history_compare_thread=None
        self._history_compare_worker=None
        self._history_compare_bridge=None
        self._operation_busy=False
        self.inventory = ()
        self.capabilities=None
        self._staged_import_dirs=[]
        self._loaded_profile_path_mappings=[]
        self._loaded_profile_environment_mappings=[]
        self._build_ui()
        self._build_inspector_dock()
        self.table.selectionModel().selectionChanged.connect(self._update_docked_inspector)
        self.query_model.dataChanged.connect(self._update_docked_inspector)
        self._apply_theme(True)

    def _apply_theme(self, dark=True):
        """Change only presentation; the application's action state is retained."""
        self.dark_theme_enabled = bool(dark)
        apply_vector_theme(self, dark)

    def _scroll_to_vector(self, section):
        """Sidebar is an anchor navigator, not a replacement for existing widgets."""
        self.workspace_scroll.ensureWidgetVisible(section, 16, 24)

    def _update_vector_metrics(self, *_):
        """Derive dashboard metrics from the existing query model, never a duplicate."""
        if not hasattr(self, "vector_kpis"):
            return
        count = self.query_model.rowCount()
        enabled = sum(bool(self.query_model.source_at(i).enabled) for i in range(count))
        self.vector_kpis["Sources"].setText(f"{count:02d}")
        self.vector_kpis["Enabled"].setText(f"{enabled:02d}")
        selected = Path(self.workbook_edit.text()).is_file()
        self.vector_kpis["Workbook"].setText("01" if selected else "00")
        self.vector_kpis["Mode"].setText("PLAN" if self.dry_run_check.isChecked() else "STAGED")

    def _build_ui(self):
        shell = QWidget(self)
        shell.setObjectName("vectorCanvas")
        shell_layout = QHBoxLayout(shell)
        shell_layout.setContentsMargins(0, 0, 0, 0)
        shell_layout.setSpacing(0)
        sidebar = QFrame(shell)
        sidebar.setObjectName("vectorSidebar")
        sidebar.setFixedWidth(70)
        side = QVBoxLayout(sidebar)
        side.setContentsMargins(9, 15, 9, 15)
        side.setSpacing(11)
        mark = QLabel("◈")
        mark.setAlignment(Qt.AlignCenter)
        mark.setObjectName("vectorEyebrow")
        mark.setToolTip("Power Query Manager")
        mark.setAccessibleName("Power Query Manager")
        mark.setMinimumHeight(48)
        side.addWidget(mark)
        side.addSpacing(15)
        navigation = (
            ("⌂", "Workspace and workbook", lambda: self._scroll_to_vector(self.vector_workbook_panel)),
            ("≡", "Query sources", lambda: self._scroll_to_vector(self.vector_sources_panel)),
            ("◇", "Dependency Finder", self._open_dependency_finder),
            ("◷", "Operation history", self._open_history),
            ("⚙", "Import settings", lambda: self._scroll_to_vector(self.vector_options_panel)),
        )
        for index, (glyph, name, callback) in enumerate(navigation):
            nav_button = QPushButton(glyph)
            nav_button.setObjectName("vectorNavActive" if index == 0 else "vectorNav")
            nav_button.setAccessibleName(name)
            nav_button.setToolTip(name)
            nav_button.clicked.connect(callback)
            side.addWidget(nav_button)
        side.addStretch()
        help_link = QPushButton("?")
        help_link.setObjectName("vectorNav")
        help_link.setAccessibleName("Workflow and safety instructions")
        help_link.setToolTip("Workflow and safety instructions")
        help_link.clicked.connect(lambda: self._scroll_to_vector(self.vector_notes))
        side.addWidget(help_link)
        shell_layout.addWidget(sidebar)
        scroll = QScrollArea()
        self.workspace_scroll = scroll
        scroll.setObjectName("vectorScroll")
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.NoFrame)
        scroll.setAccessibleName('Power Query Manager workspace')
        root = QWidget()
        root.setObjectName("vectorRoot")
        scroll.setWidget(root)
        shell_layout.addWidget(scroll, 1)
        self.setCentralWidget(shell)
        layout = QVBoxLayout(root)
        layout.setContentsMargins(28, 20, 28, 22)
        layout.setSpacing(14)

        status_row = QHBoxLayout()
        status_eyebrow = QLabel("●  PQM  /  LOCAL WORKSPACE   /   MODULE 01")
        status_eyebrow.setObjectName("vectorEyebrow")
        status_row.addWidget(status_eyebrow)
        status_row.addStretch()
        self.theme_toggle = QCheckBox('Dark theme')
        self.theme_toggle.setAccessibleName('Toggle dark theme')
        self.theme_toggle.setChecked(True)
        self.theme_toggle.toggled.connect(self._apply_theme)
        status_row.addWidget(self.theme_toggle)
        self.inspector_toggle = QCheckBox('Query inspector')
        self.inspector_toggle.setAccessibleName('Show query inspector panel')
        status_row.addWidget(self.inspector_toggle)
        layout.addLayout(status_row)

        eyebrow = QLabel("WORKSPACE   /   EXCEL AUTOMATION")
        eyebrow.setObjectName("vectorEyebrow")
        layout.addWidget(eyebrow)
        hero = QHBoxLayout()
        title = QLabel('QUERY OPERATIONS')
        title.setObjectName("vectorHero")
        title.setAccessibleName("Power Query Manager – query operations")
        hero.addWidget(title)
        hero.addStretch()
        hero_actions = QHBoxLayout()
        hero.addLayout(hero_actions)
        layout.addLayout(hero)
        subtitle = QLabel('Inspect, compare and safely import Power Query definitions into desktop Excel.')
        subtitle.setWordWrap(True)
        subtitle.setObjectName('secondaryText')
        layout.addWidget(subtitle)

        metrics = QHBoxLayout()
        metrics.setSpacing(10)
        self.vector_kpis = {}
        for label, value, hint in (
            ("Workbook", "00", "Active workbook"),
            ("Sources", "00", "Source definitions"),
            ("Enabled", "00", "Import queue"),
            ("Mode", "STAGED", "No automatic refresh"),
        ):
            card = QFrame()
            card.setObjectName("vectorMetric")
            cell = QVBoxLayout(card)
            cell.setContentsMargins(14, 12, 14, 12)
            cell.setSpacing(5)
            caption = QLabel(label.upper())
            caption.setObjectName("vectorMetricLabel")
            cell.addWidget(caption)
            number = QLabel(value)
            number.setObjectName("vectorMetricValue")
            number.setAccessibleName(label + " status")
            cell.addWidget(number)
            hint_label = QLabel(hint)
            hint_label.setObjectName("vectorMetricCaption")
            cell.addWidget(hint_label)
            self.vector_kpis[label] = number
            metrics.addWidget(card, 1)
        layout.addLayout(metrics)

        workbook_box = QGroupBox('01  /  WORKBOOK     Target workbook')
        self.vector_workbook_panel = workbook_box
        wb_layout = QGridLayout(workbook_box)
        self.workbook_edit = QLineEdit(); self.workbook_edit.setPlaceholderText('Choose an .xlsx or .xlsm workbook')
        self.workbook_edit.setAccessibleName('Target workbook path')
        browse_wb = QPushButton('Browse…');browse_wb.setAccessibleName('Browse for target workbook'); browse_wb.clicked.connect(self._browse_workbook)
        self.open_btn = QPushButton('Open workbook');self.open_btn.setAccessibleName('Open selected workbook in Excel'); self.open_btn.setEnabled(False); self.open_btn.clicked.connect(self._open_workbook)
        self.inspect_btn = QPushButton('Inspect queries');self.inspect_btn.setAccessibleName('Inspect selected workbook queries read-only'); self.inspect_btn.setEnabled(False); self.inspect_btn.clicked.connect(self._inspect_workbook)
        self.probe_btn=QPushButton('Probe Excel capabilities');self.probe_btn.setAccessibleName('Probe Excel load capabilities');self.probe_btn.clicked.connect(self._probe_excel)
        self.capability_label=QLabel('Extra load modes stay disabled until an Excel capability probe passes.')
        self.workbook_edit.textChanged.connect(lambda text: (self.open_btn.setEnabled(Path(text).is_file()),self.inspect_btn.setEnabled(Path(text).is_file()),self._refresh_recovery_state()))
        self.workbook_edit.textChanged.connect(self._update_vector_metrics)
        wb_layout.addWidget(QLabel('Target workbook'),0,0);wb_layout.addWidget(self.workbook_edit,0,1,1,2)
        workbook_actions=QGridLayout()
        for index, button in enumerate((browse_wb, self.open_btn, self.inspect_btn, self.probe_btn)):
            workbook_actions.addWidget(button, 0, index)
            workbook_actions.setColumnStretch(index, 1)
        wb_layout.addLayout(workbook_actions,1,0,1,3)
        workbook_note=QLabel('Inspect opens read-only. Capability probes use temporary workbooks and never refresh data. Imports stage changes and do not refresh queries.')
        workbook_note.setWordWrap(True);workbook_note.setMinimumWidth(0);workbook_note.setObjectName('secondaryText')
        wb_layout.addWidget(workbook_note,2,0,1,3)
        self.capability_label.setWordWrap(True);self.capability_label.setMinimumWidth(0)
        wb_layout.addWidget(self.capability_label,3,0,1,3)
        self.recover_stage_btn=QPushButton('Recover staged copy…')
        self.recover_stage_btn.setAccessibleName('Recover a verified staged workbook copy')
        self.recover_stage_btn.setToolTip('Open an abandoned staged workbook read-only in Excel and save it to a new file. The selected workbook is never overwritten.')
        self.recover_stage_btn.setEnabled(False);self.recover_stage_btn.clicked.connect(self._recover_staged_copy)
        wb_layout.addWidget(self.recover_stage_btn,4,0)
        recovery_note=QLabel('Crash recovery verifies a leftover staged copy in Excel and writes a separate new file; it never replaces the selected workbook.')
        recovery_note.setWordWrap(True);recovery_note.setMinimumWidth(0);recovery_note.setObjectName('secondaryText')
        wb_layout.addWidget(recovery_note,4,1,1,2)
        layout.addWidget(workbook_box)

        finder_row = QHBoxLayout()
        self.dependency_finder_button = QPushButton('Dependency Finder…')
        self.dependency_finder_button.setObjectName('vectorFinderButton')
        self.dependency_finder_button.setAccessibleName('Find files that depend on selected Excel or Access databases and workbooks')
        self.dependency_finder_button.clicked.connect(self._open_dependency_finder)
        finder_row.addWidget(self.dependency_finder_button)
        finder_note = QLabel('Find incoming references to .xlsx, .xlsm, .accdb or .mdb files in local/mapped folders (read-only).')
        finder_note.setObjectName('secondaryText')
        finder_note.setWordWrap(True)
        finder_row.addWidget(finder_note, 1)
        finder_band = QFrame()
        finder_band.setObjectName('vectorFinderBand')
        finder_band_layout = QVBoxLayout(finder_band)
        finder_band_layout.setContentsMargins(12, 9, 12, 9)
        finder_band_layout.addLayout(finder_row)
        layout.addWidget(finder_band)

        query_box = QGroupBox('02  /  SOURCES     Queries & source files')
        self.vector_sources_panel = query_box
        q_layout = QVBoxLayout(query_box)
        controls=QGridLayout()
        source_actions=[('Add files…',self._add_files),('Add folder…',self._add_folder),
            ('Add folder recursively…',self._add_folder_recursive),('Add ZIP…',self._add_zip),
            ('Inspect / compare selected…',self._inspect_selected),('Dependency graph…',self._show_dependency_graph),
            ('Browse selected path…',self._browse_source),('Remove selected',self._remove_selected),
            ('Clear list',self._clear_sources)]
        for index,(label,callback) in enumerate(source_actions):
            button=QPushButton(label);button.setAccessibleName(label.rstrip('…'));button.clicked.connect(callback)
            controls.addWidget(button,index//4,index%4)
        for column in range(4): controls.setColumnStretch(column,1)
        q_layout.addLayout(controls)
        bulk=QGridLayout()
        bulk_actions=[('Enable selected',lambda:self._set_selected_enabled(True)),
            ('Disable selected',lambda:self._set_selected_enabled(False)),
            ('Set selected order…',self._set_selected_order),
            ('Configure selected destinations…',self._configure_destinations),
            ('Compare package versions…',self._preview_query_package_update)]
        for index,(label,callback) in enumerate(bulk_actions):
            button=QPushButton(label);button.setAccessibleName(label.rstrip('…'));button.clicked.connect(callback)
            bulk.addWidget(button,index//3,index%3)
        self.undo_action=self.query_model.undo_stack.createUndoAction(self,'Undo table edit')
        self.undo_action.setShortcut(QKeySequence.StandardKey.Undo);self.addAction(self.undo_action)
        self.redo_action=self.query_model.undo_stack.createRedoAction(self,'Redo table edit')
        self.redo_action.setShortcut(QKeySequence.StandardKey.Redo);self.addAction(self.redo_action)
        self.undo_button=QPushButton('Undo');self.undo_button.setAccessibleName('Undo last table edit');self.undo_button.setToolTip('Undo the last table edit (Ctrl+Z)')
        self.undo_button.clicked.connect(self.undo_action.trigger);bulk.addWidget(self.undo_button,2,0)
        self.redo_button=QPushButton('Redo');self.redo_button.setAccessibleName('Redo last table edit');self.redo_button.setToolTip('Redo the last table edit (Ctrl+Y)')
        self.redo_button.clicked.connect(self.redo_action.trigger);bulk.addWidget(self.redo_button,2,1)
        self.query_model.undo_stack.canUndoChanged.connect(self._update_undo_actions)
        self.query_model.undo_stack.canRedoChanged.connect(self._update_undo_actions)
        self._update_undo_actions()
        for column in range(3): bulk.setColumnStretch(column,1)
        q_layout.addLayout(bulk)
        self.search_edit=QLineEdit(); self.search_edit.setPlaceholderText('Filter queries by name, path, mode, or status…')
        self.search_edit.setAccessibleName('Filter query source rows')
        self.search_edit.textChanged.connect(self.query_proxy.setFilterFixedString)
        q_layout.addWidget(self.search_edit)
        self.select_all_button=QPushButton('Select all query rows')
        self.select_all_button.setAccessibleName('Select all query rows')
        q_layout.addWidget(self.select_all_button)
        self.table = SourceTableView()
        self.table.setAccessibleName('Query source settings table')
        self.table.setCornerButtonEnabled(False)
        self.table.setMinimumHeight(260)
        self.table.filesDropped.connect(self._handle_dropped_paths)
        self.table.setModel(self.query_proxy)
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.table.setAlternatingRowColors(True)
        self.table.setSortingEnabled(True)
        self.table.setDragDropMode(QAbstractItemView.InternalMove)
        self.table.setDefaultDropAction(Qt.MoveAction)
        self.table.setDragDropOverwriteMode(False)
        self.table.setTabKeyNavigation(True)
        self.select_all_button.clicked.connect(self.table.selectAll)
        self.load_mode_delegate=LoadModeDelegate(self.table,self.query_model.supported_load_modes)
        self.table.setItemDelegateForColumn(QueryTableModel.LOAD,self.load_mode_delegate)
        self.table.horizontalHeader().setSectionResizeMode(QueryTableModel.ENABLED,QHeaderView.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(QueryTableModel.ORDER,QHeaderView.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(QueryTableModel.NAME,QHeaderView.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(QueryTableModel.PATH,QHeaderView.Stretch)
        self.table.horizontalHeader().setSectionResizeMode(QueryTableModel.LOAD,QHeaderView.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(QueryTableModel.STATUS,QHeaderView.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(QueryTableModel.DEPENDENCIES,QHeaderView.Stretch)
        self.query_model.dataChanged.connect(self._model_changed)
        q_layout.addWidget(self.table)
        self.sequence_label = QLabel('Add .m or .pq files to see the import sequence.')
        self.sequence_label.setObjectName('secondaryText');self.sequence_label.setWordWrap(True)
        q_layout.addWidget(self.sequence_label)
        layout.addWidget(query_box, 3)

        options = QGroupBox('03  /  IMPORT CONTROL     Options & safety')
        self.vector_options_panel = options
        options_layout=QGridLayout(options)
        mode_row=QHBoxLayout();mode_row.addWidget(QLabel('If query exists:'))
        self.mode_combo=QComboBox();self.mode_combo.setAccessibleName('Existing query behavior'); self.mode_combo.addItem('Replace definition',ExistingQueryMode.REPLACE.value); self.mode_combo.addItem('Skip existing query',ExistingQueryMode.SKIP.value); mode_row.addWidget(self.mode_combo)
        options_layout.addLayout(mode_row,0,0,1,2)
        self.backup_check=QCheckBox('Create timestamped workbook backup before changes');self.backup_check.setAccessibleName('Create a timestamped workbook backup'); self.backup_check.setChecked(True); options_layout.addWidget(self.backup_check,1,0,1,2)
        self.dry_run_check=QCheckBox('Plan only (no Excel changes)');self.dry_run_check.setAccessibleName('Plan only without changing Excel');options_layout.addWidget(self.dry_run_check,2,0)
        self.dry_run_check.toggled.connect(self._update_vector_metrics)
        save_profile_btn=QPushButton('Save profile…');save_profile_btn.setAccessibleName('Save import profile');save_profile_btn.clicked.connect(self._save_profile);options_layout.addWidget(save_profile_btn,2,1)
        load_profile_btn=QPushButton('Load profile…');load_profile_btn.setAccessibleName('Load import profile');load_profile_btn.clicked.connect(self._load_profile);options_layout.addWidget(load_profile_btn,3,0)
        self.edit_mappings_btn=QPushButton('Edit source mappings…');self.edit_mappings_btn.setAccessibleName('Edit source environment mappings');self.edit_mappings_btn.clicked.connect(self._edit_source_mappings);options_layout.addWidget(self.edit_mappings_btn,3,1)
        options_layout.setColumnStretch(0,1);options_layout.setColumnStretch(1,1)
        layout.addWidget(options)
        bottom=QHBoxLayout()
        self.progress=QProgressBar(); self.progress.setRange(0,0); self.progress.setVisible(False); self.progress.setTextVisible(False)
        self.install_btn=QPushButton('Review and import queries')
        self.install_btn.setAccessibleName('Review and import selected queries')
        self.install_btn.setObjectName('primary')
        self.install_btn.setToolTip('Review before committing any staged Excel workbook changes')
        self.install_btn.clicked.connect(self._install)
        hero_actions.addWidget(self.install_btn)
        bottom.addWidget(self.progress,1)
        safety_status=QLabel('STAGED / REVIEW REQUIRED / NO AUTOMATIC REFRESH')
        safety_status.setObjectName('vectorEyebrow')
        bottom.addWidget(safety_status)
        layout.addLayout(bottom)
        splitter=QSplitter(Qt.Horizontal)
        self.vector_notes = splitter
        instruction_box=QGroupBox('Workflow and limits'); il=QVBoxLayout(instruction_box); instructions=QPlainTextEdit(); instructions.setAccessibleName('Workflow and safety instructions');instructions.setReadOnly(True); instructions.setPlainText(INSTRUCTIONS); il.addWidget(instructions)
        log_box=QGroupBox('Operation log'); ll=QVBoxLayout(log_box)
        log_controls=QHBoxLayout();self.history_button=QPushButton('History…');self.history_button.setAccessibleName('Open operation history');self.history_button.clicked.connect(self._open_history);log_controls.addStretch();log_controls.addWidget(self.history_button);ll.addLayout(log_controls)
        self.log=QPlainTextEdit();self.log.setObjectName('operationLogOutput');self.log.setAccessibleName('Operation log output');self.log.setAccessibleDescription('Live read-only progress and result details for the current operation.');self.log.setReadOnly(True);self.log.setPlaceholderText('Validation and Excel progress will appear here.');ll.addWidget(self.log)
        splitter.addWidget(instruction_box); splitter.addWidget(log_box); splitter.setSizes([500,680]); layout.addWidget(splitter,2)
        splitter.setMinimumHeight(180)
        self.statusBar().showMessage('Ready · Local operations · Source files remain protected')
        self.query_model.rowsInserted.connect(self._update_vector_metrics)
        self.query_model.rowsRemoved.connect(self._update_vector_metrics)
        self.query_model.modelReset.connect(self._update_vector_metrics)
        self.query_model.dataChanged.connect(self._update_vector_metrics)
        self._update_vector_metrics()

    def _build_inspector_dock(self):
        self._inspector_source_row=-1
        self.inspector_dock=QDockWidget('Query inspector',self)
        self.inspector_dock.setObjectName('queryInspectorDock')
        self.inspector_dock.setAccessibleName('Query inspector dock')
        self.inspector_dock.setAllowedAreas(Qt.DockWidgetArea.RightDockWidgetArea|Qt.DockWidgetArea.BottomDockWidgetArea)
        self.inspector_dock.setFeatures(QDockWidget.DockWidgetFeature.DockWidgetClosable|
                                        QDockWidget.DockWidgetFeature.DockWidgetMovable|
                                        QDockWidget.DockWidgetFeature.DockWidgetFloatable)
        panel=QWidget(self.inspector_dock);layout=QVBoxLayout(panel)
        self.inspector_status=QLabel('Select a query row to inspect its source and workbook definition.')
        self.inspector_status.setWordWrap(True);self.inspector_status.setObjectName('secondaryText')
        layout.addWidget(self.inspector_status)
        self.inspector_tabs=QTabWidget(panel)
        self.inspector_source_text=QPlainTextEdit(panel);self.inspector_source_text.setReadOnly(True)
        self.inspector_source_text.setAccessibleName('Selected source M code preview')
        self.inspector_workbook_text=QPlainTextEdit(panel);self.inspector_workbook_text.setReadOnly(True)
        self.inspector_workbook_text.setAccessibleName('Workbook M code preview')
        self.inspector_tabs.addTab(self.inspector_source_text,'Source M')
        self.inspector_tabs.addTab(self.inspector_workbook_text,'Workbook M')
        layout.addWidget(self.inspector_tabs,1)
        self.inspector_open_button=QPushButton('Open compare and editor…',panel)
        self.inspector_open_button.setAccessibleName('Open full query comparison and editor')
        self.inspector_open_button.clicked.connect(self._inspect_docked_query)
        layout.addWidget(self.inspector_open_button)
        self.inspector_dock.setWidget(panel)
        self.addDockWidget(Qt.DockWidgetArea.RightDockWidgetArea,self.inspector_dock)
        self.inspector_dock.hide()
        self.inspector_toggle.toggled.connect(self.inspector_dock.setVisible)
        self.inspector_dock.visibilityChanged.connect(self.inspector_toggle.setChecked)
        view_menu=self.menuBar().addMenu('&View');view_menu.addAction(self.inspector_dock.toggleViewAction())
        self._update_docked_inspector()

    def _record_diagnostic(self,operation,error):
        path=record_failure(operation,error)
        if path and hasattr(self,'log'):self.log.appendPlainText(f'Local diagnostic record saved: {path}')

    def _update_docked_inspector(self,*_):
        selected=self.table.selectionModel().selectedRows() if hasattr(self,'table') else []
        if not selected:
            self._inspector_source_row=-1
            self.inspector_status.setText('Select a query row to inspect its source and workbook definition.')
            self.inspector_source_text.clear();self.inspector_workbook_text.clear()
            self.inspector_open_button.setEnabled(False)
            return
        row=self.query_proxy.mapToSource(selected[0]).row()
        if row<0 or row>=self.query_model.rowCount():return
        self._inspector_source_row=row
        source=self.query_model.source_at(row)
        existing=next((q for q in self.inventory if q.name.casefold()==source.name.casefold()),None)
        try:
            text=Path(source.path).read_text(encoding='utf-8-sig')
            if len(text)>100_000:text=text[:100_000]+'\n\n[Preview truncated at 100,000 characters.]'
            self.inspector_source_text.setPlainText(text)
        except (OSError,UnicodeError) as exc:
            self.inspector_source_text.setPlainText(f'[Source unavailable: {exc}]')
        if existing is None:
            self.inspector_workbook_text.setPlainText('[Query not present in the inspected workbook; comparison is unknown.]')
            self.inspector_status.setText(f'{source.name} — not present in the inspected workbook')
        elif existing.formula is None:
            self.inspector_workbook_text.setPlainText('[Workbook formula unavailable; comparison is unknown.]')
            self.inspector_status.setText(f'{source.name} — workbook query found; formula unavailable')
        else:
            formula=existing.formula
            if len(formula)>100_000:formula=formula[:100_000]+'\n\n[Preview truncated at 100,000 characters.]'
            self.inspector_workbook_text.setPlainText(formula)
            self.inspector_status.setText(f'{source.name} — workbook query found; compare the definitions in the tabs')
        self.inspector_open_button.setEnabled(True)

    def _inspect_docked_query(self):
        row=self._inspector_source_row
        if row>=0:self._inspect_source_row(row)

    def _browse_workbook(self):
        path,_=QFileDialog.getOpenFileName(self,'Select Excel workbook','','Excel Workbooks (*.xlsx *.xlsm)')
        if path: self.workbook_edit.setText(path); self.open_btn.setEnabled(True)

    def _open_dependency_finder(self):
        if self._operation_busy:
            return
        dialog = DependencyFinderDialog(self)
        dialog.exec()
    def _add_files(self):
        paths,_=QFileDialog.getOpenFileNames(self,'Select Power Query files','','Power Query sources (*.m *.pq)')
        if paths: self._merge_sources(discover_query_files(paths))
    def _add_folder(self):
        self._add_folder_mode(False)
    def _add_folder_recursive(self):
        self._add_folder_mode(True)
    def _add_folder_mode(self,recursive):
        folder=QFileDialog.getExistingDirectory(self,'Select folder containing Power Query files')
        if folder:
            found=discover_query_files([folder],recursive=recursive)
            if not found: QMessageBox.information(self,'No query files','No .m or .pq source files were found in the selected folder.'); return
            self._merge_sources(found)
    def _add_zip(self):
        path,_=QFileDialog.getOpenFileName(self,'Select Power Query ZIP','','ZIP archives (*.zip)')
        if path:self._import_zip(path)
    def _import_zip(self,path):
        root=None
        try:
            root=extract_query_zip(path)
            sources=discover_query_files([root],recursive=True)
            sources=apply_query_package_manifest(sources,root)
            if not sources:
                shutil.rmtree(root,ignore_errors=True)
                QMessageBox.information(self,'No query files','The ZIP contains no usable .m or .pq sources.');return
            self._staged_import_dirs.append(root);self._merge_sources(sources)
        except Exception as exc:
            if root is not None and root not in self._staged_import_dirs:shutil.rmtree(root,ignore_errors=True)
            self._record_diagnostic('package_import',exc)
            QMessageBox.critical(self,'ZIP import rejected',f'{type(exc).__name__}: {exc}')

    def _preview_query_package_update(self):
        package_filter='Power Query packages (*.pqm.zip *.zip)'
        base,_=QFileDialog.getOpenFileName(self,'Select current query package','','Power Query packages (*.pqm.zip *.zip)')
        if not base:return
        updated,_=QFileDialog.getOpenFileName(self,'Select updated query package','',package_filter)
        if not updated:return
        try:
            changes=compare_query_packages(base,updated)
        except Exception as exc:
            self._record_diagnostic('package_update_preview',exc)
            QMessageBox.critical(self,'Package comparison failed',f'{type(exc).__name__}: {exc}')
            return
        dialog=QueryPackageUpdateDialog(base,updated,changes,self)
        dialog.exec()
        counts={status:sum(row.status==status for row in changes)
                for status in ('ADDED','MODIFIED','UNCHANGED','REMOVED','CONFLICT')}
        self._record_history('query_package_update_previewed',counts)
        self.statusBar().showMessage('Package update preview completed; no workbook changes were made.')
    def _handle_dropped_paths(self,paths):
        source_paths=[p for p in paths if Path(p).suffix.casefold() in {'.m','.pq'}]
        archives=[p for p in paths if Path(p).suffix.casefold()=='.zip']
        if source_paths:self._merge_sources(discover_query_files(source_paths))
        for archive in archives:self._import_zip(archive)
    def _merge_sources(self, new_sources):
        paths={str(s.path).casefold() for s in self.sources}
        self.sources.extend(s for s in new_sources if str(s.path).casefold() not in paths)
        self.query_model.append_sources(new_sources)
        # Keep automatically assigned no-prefix files after the largest explicit prefix, even
        # when users add files in separate selection batches. Explicit user edits are preserved.
        prefixed=[s.order for s in self.sources if s.order_from_prefix]
        next_order=max(prefixed,default=0)+1
        auto=[s for s in self.sources if not s.order_from_prefix and s.order_automatic]
        auto.sort(key=lambda s:(s.path.name.casefold(),str(s.path).casefold()))
        for offset,source in enumerate(auto): source.order=next_order+offset
        self._refresh_table()
    def _refresh_table(self):
        self.query_model.refresh_all()
        self._update_source_list(); self._update_sequence(); self._validate_rows()
    def _update_source_list(self):
        self.sources=[self.query_model.source_at(i) for i in range(self.query_model.rowCount())]
    def _source_by_uid(self,uid):
        row=self.query_model.row_for_id(str(uid))
        return self.query_model.source_at(row) if row >= 0 else None
    def _model_changed(self,top_left=None,bottom_right=None,*_):
        if top_left is not None and top_left.column()==bottom_right.column() and top_left.column() in (QueryTableModel.STATUS,QueryTableModel.DEPENDENCIES):
            return
        self._update_source_list(); self._update_sequence()
        if self.inventory:
            for row in range(self.query_model.rowCount()):
                self.query_model.set_status(self.query_model.stable_id_at(row),'Re-inspect after source edits')
        self._validate_rows()
    def _validate_rows(self):
        names={}
        for source in self.sources:
            if source.enabled: names.setdefault(source.name.casefold(), []).append(source)
        issues={}
        for row,source in enumerate(self.sources):
            if not source.enabled: continue
            reason=''
            try:
                from pq_installer.core.validation import validate_query_name
                validate_query_name(source.name)
            except ValueError as exc: reason=str(exc)
            if not reason and len(names.get(source.name.casefold(), []))>1: reason=f'Duplicate query name: {source.name}'
            if not reason and (not source.path.is_file()): reason=f'Source file not found: {source.path}'
            if not reason and source.path.suffix.lower() not in {'.m','.pq'}: reason='Use a .m or .pq source file.'
            issues[self.query_model.stable_id_at(row)]=reason
        for row in range(self.query_model.rowCount()):
            stable_id=self.query_model.stable_id_at(row); reason=issues.get(stable_id,'')
            current=self.query_model.data(self.query_model.index(row,QueryTableModel.STATUS)) or ''
            if reason:
                self.query_model.set_status(stable_id,'Invalid: '+reason)
            elif current.startswith('Invalid: '):
                self.query_model.set_status(stable_id,'Re-inspect after source edits' if self.inventory else 'Not inspected')
        self.query_model.apply_dependency_analysis()
    def _update_sequence(self):
        seq=ordered_sources(self.sources); warnings=order_warnings(self.sources)
        text='Import order: '+'  →  '.join(f'{s.order}. {s.name}' for s in seq) if seq else 'Add .m or .pq files to see the import sequence.'
        if warnings: text+='\nOrder ties: '+'; '.join(warnings)
        self.sequence_label.setText(text)
    def _browse_source(self):
        rows=sorted({self.query_proxy.mapToSource(i).row() for i in self.table.selectedIndexes()})
        if not rows: return
        row=rows[0]; old=Path(self.query_model.source_at(row).path); path,_=QFileDialog.getOpenFileName(self,'Select Power Query source',str(old.parent),'Power Query sources (*.m *.pq)')
        if path:self.query_model.setData(self.query_model.index(row,QueryTableModel.PATH),path,Qt.EditRole)
    def _show_dependency_graph(self):
        sources=[self.query_model.source_at(row) for row in range(self.query_model.rowCount())]
        DependencyGraphDialog(sources,self).exec()
    def _inspect_selected(self):
        selected=self.table.selectionModel().selectedRows()
        if not selected:
            QMessageBox.information(self,'Select a query','Select one query row to inspect its source and workbook formula.');return
        row=self.query_proxy.mapToSource(selected[0]).row()
        self._inspect_source_row(row)
    def _inspect_source_row(self,row):
        if row<0 or row>=self.query_model.rowCount():return
        source=self.query_model.source_at(row)
        existing=next((q for q in self.inventory if q.name.casefold()==source.name.casefold()),None)
        stable_id=self.query_model.stable_id_at(row)
        def saved(destination):
            current=self.query_model.row_for_id(stable_id)
            if current>=0:self.query_model.setData(self.query_model.index(current,QueryTableModel.PATH),str(destination),Qt.EditRole)
        try:dialog=QueryInspectorDialog(source,existing,self,saved)
        except (OSError,ValueError) as exc:QMessageBox.warning(self,'Cannot inspect query source',str(exc));return
        self._record_history('query_compared',{'query_name':source.name})
        dialog.exec()
    def _remove_selected(self):
        rows=sorted({self.query_proxy.mapToSource(i).row() for i in self.table.selectionModel().selectedRows()})
        self.query_model.remove_rows(rows); self._update_source_list(); self._update_sequence(); self._validate_rows()
    def _selected_source_rows(self):
        return sorted({self.query_proxy.mapToSource(index).row() for index in self.table.selectionModel().selectedRows()})
    def _set_selected_enabled(self,enabled):
        rows=self._selected_source_rows()
        if not rows:return
        self.query_model.undo_stack.beginMacro('Enable selected queries' if enabled else 'Disable selected queries')
        try:
            for row in rows:
                self.query_model.setData(self.query_model.index(row,QueryTableModel.ENABLED),Qt.Checked if enabled else Qt.Unchecked,Qt.CheckStateRole)
        finally:self.query_model.undo_stack.endMacro()
    def _set_selected_order(self):
        rows=self._selected_source_rows()
        if not rows:QMessageBox.information(self,'Select queries','Select one or more rows to set their order.');return
        start,ok=QInputDialog.getInt(self,'Set selected order','Starting order:',min(self.query_model.source_at(r).order for r in rows),0,2147483647)
        if not ok:return
        rows.sort(key=lambda r:(self.query_model.source_at(r).order,self.query_model.source_at(r).path.name.casefold()))
        self.query_model.undo_stack.beginMacro('Set selected query order')
        try:
            for offset,row in enumerate(rows):self.query_model.setData(self.query_model.index(row,QueryTableModel.ORDER),start+offset,Qt.EditRole)
        finally:self.query_model.undo_stack.endMacro()
    def _configure_destinations(self):
        rows=self._selected_source_rows()
        if not rows:QMessageBox.information(self,'Select queries','Select one or more rows to configure their load destinations.');return
        updates=[]
        for row in rows:
            source=self.query_model.source_at(row)
            if source.load_mode not in {LoadMode.NEW_SHEET,LoadMode.EXISTING_SHEET}:
                continue
            sheet_name=source.sheet_name;cell=source.cell;table_name=source.table_name
            if source.load_mode==LoadMode.NEW_SHEET:
                name,ok=QInputDialog.getText(self,'New worksheet destination',
                    f'New worksheet name for {source.name} (leave blank for an automatic name):',
                    text=source.sheet_name or '')
                if not ok:return
                sheet_name=name.strip() or None
            else:
                name,ok=QInputDialog.getText(self,'Existing worksheet destination',
                    f'Existing worksheet name for {source.name}:',text=source.sheet_name or '')
                if not ok:return
                sheet_name=name
                cell,ok=QInputDialog.getText(self,'Starting cell',
                    f'Upper-left starting cell for {source.name}:',text=source.cell or 'A1')
                if not ok:return
                cell=cell
            table,ok=QInputDialog.getText(self,'Excel table name',
                f'Excel table name for {source.name} (blank uses an automatic name):',
                text=source.table_name or '')
            if not ok:return
            table_name=table.strip() or None
            updates.append((self.query_model.stable_id_at(row),sheet_name,cell,table_name))
        if updates:
            self.query_model.undo_stack.beginMacro('Configure query destinations')
            try:
                for stable_id,sheet_name,cell,table_name in updates:
                    self.query_model.set_destination(stable_id,sheet_name=sheet_name,cell=cell,table_name=table_name)
            finally:self.query_model.undo_stack.endMacro()
        self._validate_rows()
    def _clear_sources(self): self.query_model.clear(); self._update_source_list(); self._update_sequence(); self._validate_rows()
    def _update_undo_actions(self,*_):
        self.undo_action.setEnabled(self.query_model.undo_stack.canUndo() and not self._operation_busy)
        self.redo_action.setEnabled(self.query_model.undo_stack.canRedo() and not self._operation_busy)
    def _set_busy(self,busy,activity='Importing…'):
        self._operation_busy=bool(busy)
        self.centralWidget().setEnabled(not busy)
        # Dock widgets and menu actions are siblings of the central widget in
        # QMainWindow, so disabling the central widget alone leaves them live.
        if hasattr(self,'inspector_dock'):
            self.inspector_dock.setEnabled(not busy)
        self.menuBar().setEnabled(not busy)
        self._update_undo_actions()
        self.install_btn.setEnabled(not busy); self.install_btn.setText(activity if busy else 'Review and import queries'); self.progress.setVisible(busy)
        self._refresh_recovery_state()

    @Slot()
    def _on_background_thread_finished(self):
        self._set_busy(False)

    def _refresh_recovery_state(self):
        button=getattr(self,'recover_stage_btn',None)
        if button is None:return
        workbook=Path(self.workbook_edit.text().strip()) if self.workbook_edit.text().strip() else None
        try:candidates=find_staged_workbooks(workbook) if workbook is not None and workbook.is_file() else ()
        except OSError:candidates=()
        button.setEnabled(bool(candidates) and not self._operation_busy)
        button.setToolTip((f'{len(candidates)} abandoned staged copy/copies found. Recovery verifies in Excel and saves a separate copy without changing the selected workbook.'
                           if candidates else 'No matching staged copies were found beside the selected workbook.'))

    def _recover_staged_copy(self):
        workbook=Path(self.workbook_edit.text().strip())
        candidates=find_staged_workbooks(workbook)
        if not candidates:
            self._refresh_recovery_state()
            QMessageBox.information(self,'No staged copies','No abandoned staged workbook copies were found beside the selected workbook.')
            return
        try:diagnoses=[diagnose_staged_workbook(workbook,path) for path in candidates]
        except (OSError,ValueError):
            self._refresh_recovery_state()
            QMessageBox.information(self,'Stage list changed',
                'A staged candidate changed or disappeared while it was being checked. Review the refreshed list and try again.')
            return
        status_labels={'structure_present':'Open XML parts present; needs Excel check',
                       'missing_workbook_parts':'incomplete Open XML structure',
                       'unreadable_container':'unreadable ZIP container','empty':'empty file'}
        labels=[f"{path.name} — {status_labels[diagnosis['status']]}; "
                f"{diagnosis['size_bytes']:,} bytes; age {diagnosis['age_seconds']/3600:.1f} h; "
                f"modified {diagnosis['modified_at']}"
                for path,diagnosis in zip(candidates,diagnoses)]
        selected,ok=QInputDialog.getItem(self,'Select staged workbook','Candidate:',labels,0,False)
        if not ok:return
        selected_index=labels.index(selected)
        staged=candidates[selected_index]
        diagnosis=diagnoses[selected_index]
        QMessageBox.information(self,'Recovery safety',
            f"Preliminary stage check: {diagnosis['message']}\n\n"
            'This check does not prove that an import committed or validate query contents. '
            'The selected candidate must still open read-only in Excel before its bytes can be saved to a new file. '
            'The selected workbook and staged candidate are preserved.')
        output,_=QFileDialog.getSaveFileName(self,'Save verified recovery copy',
            str(workbook.with_name(f'{workbook.stem}_recovered{workbook.suffix}')),
            'Excel Workbooks (*.xlsx *.xlsm)')
        if not output:return
        self._set_busy(True,'Verifying recovery copy…')
        self.statusBar().showMessage('Opening staged workbook read-only for recovery verification…')
        self._recovery_thread=QThread(self)
        self._recovery_worker=StagedRecoveryWorker(workbook,staged,output)
        self._recovery_worker.moveToThread(self._recovery_thread)
        self._recovery_thread.started.connect(self._recovery_worker.run)
        self._recovery_worker.completed.connect(self._on_recovery_completed)
        self._recovery_worker.failed.connect(self._on_recovery_failed)
        self._recovery_worker.finished.connect(self._recovery_thread.quit)
        self._recovery_worker.finished.connect(self._recovery_worker.deleteLater)
        self._recovery_thread.finished.connect(self._recovery_thread.deleteLater)
        self._recovery_thread.finished.connect(self._on_background_thread_finished)
        self._recovery_thread.start()

    def _on_recovery_completed(self,outcome):
        self.log.appendPlainText(f"RECOVERY COPY SAVED: {outcome['path']}")
        self.log.appendPlainText(f"Excel read-only verification found {outcome['query_count']} query definition(s); SHA-256: {outcome['sha256']}.")
        self.statusBar().showMessage('Verified staged copy saved separately; original workbook unchanged')
        self._record_history('staged_copy_recovered',{'query_count':outcome['query_count'],'sha256':outcome['sha256']})
        QMessageBox.information(self,'Recovery copy saved',f"Excel verified the staged workbook read-only. A separate copy was saved to:\n\n{outcome['path']}\n\nThe selected workbook was not modified.")

    def _on_recovery_failed(self,message):
        self.log.appendPlainText(f'RECOVERY FAILED: {message}')
        self.statusBar().showMessage('Staged copy recovery failed; original workbook unchanged')
        QMessageBox.warning(self,'Could not recover staged copy',f'{message}\n\nThe selected workbook and candidate staged file were not intentionally modified.')
    def _sync_table_edits(self):
        self._update_source_list()
    def _install(self):
        try:
            self._sync_table_edits()
            workbook=validate_workbook(self.workbook_edit.text().strip()); validate_query_sources(self.sources)
            sequence=ordered_sources(self.sources); warns=order_warnings(self.sources)
            summary='Final import sequence:\n'+'\n'.join(
                f'{s.order}. {s.name} — {s.path.name} [load: {s.load_mode.value}'
                +(f', {s.sheet_name or "new sheet"}!{s.cell or "A1"}' if s.load_mode in {LoadMode.NEW_SHEET,LoadMode.EXISTING_SHEET} else '')
                +']' for s in sequence)
            if warns: summary+='\n\n'+ '\n'.join(warns)
            summary+='\n\nExcel desktop will edit a staged copy and commit it only after verification.'
            if any(s.load_mode in {LoadMode.NEW_SHEET,LoadMode.EXISTING_SHEET,LoadMode.DATA_MODEL} for s in sequence):
                summary+='\n\nSelected worksheet/Data Model loads evaluate query results and may contact external data sources. Continue only if you intend to run these queries.'
            elif all(s.load_mode==LoadMode.DEFINITION_ONLY for s in sequence):
                summary+='\n\nDefinition-only mode adds or updates query definitions without refreshing query results.'
            if self.dry_run_check.isChecked():summary+='\n\nPLAN ONLY: Excel will not be opened and no workbook or backup file will be written.'
            answer=QMessageBox.question(self,'Review import sequence',summary+'\n\nContinue?',QMessageBox.Yes|QMessageBox.Cancel,QMessageBox.Cancel)
            if answer!=QMessageBox.Yes:return
            self.log.clear(); self.log.appendPlainText(f'Workbook: {workbook}'); self.log.appendPlainText(f'Enabled queries: {len(sequence)}'); self.log.appendPlainText('Validating sources and starting Excel COM…')
            if self.dry_run_check.isChecked():
                self.log.appendPlainText('DRY RUN: offline path and source checks passed. Excel was not opened; no backup or workbook write occurred.')
                self.statusBar().showMessage('Dry-run review complete; workbook unchanged')
                return
            mode=parse_existing_query_mode(self.mode_combo.currentData()); self._set_busy(True)
            history_store=self._get_history_store();snapshot_store=history_store if history_store.snapshots_enabled else None
            self._thread=QThread(self); self._worker=InstallWorker(workbook,sequence,mode,self.backup_check.isChecked(),self.capabilities,snapshot_store)
            self._worker.moveToThread(self._thread); self._thread.started.connect(self._worker.run); self._worker.log.connect(self.log.appendPlainText)
            self._worker.completed.connect(self._on_completed); self._worker.failed.connect(self._on_failed); self._worker.finished.connect(self._thread.quit); self._worker.finished.connect(self._worker.deleteLater); self._thread.finished.connect(self._thread.deleteLater); self._thread.finished.connect(self._on_background_thread_finished); self._thread.start()
        except ValueError as exc:
            self._record_diagnostic('import_validation',exc)
            self.log.appendPlainText(f'VALIDATION ERROR: {exc}'); QMessageBox.warning(self,'Cannot import queries',str(exc))
        except Exception as exc:
            self._record_diagnostic('import_start',exc)
            self._set_busy(False); message=f'{type(exc).__name__}: {exc}'; self.log.appendPlainText(f'ERROR: {message}'); QMessageBox.critical(self,'Could not start import',message)
    def _on_completed(self,outcome):
        if isinstance(outcome,dict):results=outcome.get('results',());snapshot_id=outcome.get('snapshot_id')
        else:results=outcome;snapshot_id=None
        all_skipped = bool(results) and all(
            getattr(result.status,'value',result.status)=='Skipped' for result in results)
        statuses=[getattr(result.status,'value',result.status) for result in results]
        details={'query_count':len(results),'workbook':Path(self.workbook_edit.text()).name,
            'action':'import','status':'skipped' if all_skipped else 'completed',
            'created_count':sum(status=='Created' for status in statuses),
            'replaced_count':sum(status=='Replaced' for status in statuses),
            'skipped_count':sum(status=='Skipped' for status in statuses)}
        if snapshot_id:details['snapshot_id']=snapshot_id
        self._record_history('import_skipped' if all_skipped else 'import_completed',details)
        if all_skipped:
            self.statusBar().showMessage('No changes: all selected queries were skipped')
            QMessageBox.information(self,'No changes',
                f'All {len(results)} selected queries were skipped. The workbook was not changed.')
            return
        self.statusBar().showMessage('Import completed successfully')
        QMessageBox.information(self,'Import complete',f'Processed {len(results)} enabled query file(s). Review the operation log and Data > Queries & Connections in Excel.')
    def _on_failed(self,message):
        self._record_history('import_failed',{'workbook':Path(self.workbook_edit.text()).name})
        self.statusBar().showMessage('Import failed; review the operation log'); self.log.appendPlainText(f'FAILED: {message}')
        QMessageBox.critical(self,'Import failed',f'Excel could not complete the operation.\n\n{message}\n\nIf enabled, the timestamped backup is available beside the workbook.')
    def _inspect_workbook(self):
        try: workbook=validate_workbook(self.workbook_edit.text().strip())
        except ValueError as exc: QMessageBox.warning(self,'Cannot inspect workbook',str(exc)); return
        self._set_busy(True)
        self.statusBar().showMessage('Inspecting workbook read-only…')
        self._inspect_thread=QThread(self);self._inspect_worker=InspectWorker(workbook)
        self._inspect_worker.moveToThread(self._inspect_thread)
        self._inspect_thread.started.connect(self._inspect_worker.run)
        self._inspect_worker.completed.connect(self._on_inspected)
        self._inspect_worker.failed.connect(self._on_inspect_failed)
        self._inspect_worker.finished.connect(self._inspect_thread.quit)
        self._inspect_worker.finished.connect(self._inspect_worker.deleteLater)
        self._inspect_thread.finished.connect(self._inspect_thread.deleteLater)
        self._inspect_thread.finished.connect(self._on_background_thread_finished)
        self._inspect_thread.start()
    def _probe_excel(self):
        if self._inspect_thread is not None and self._inspect_thread.isRunning():return
        self._set_busy(True);self.statusBar().showMessage('Probing Excel on a disposable temporary workbook…')
        self._inspect_thread=QThread(self);self._inspect_worker=CapabilityProbeWorker()
        self._inspect_worker.moveToThread(self._inspect_thread);self._inspect_thread.started.connect(self._inspect_worker.run)
        self._inspect_worker.completed.connect(self._on_capabilities)
        self._inspect_worker.failed.connect(self._on_capability_probe_failed)
        self._inspect_worker.finished.connect(self._inspect_thread.quit);self._inspect_worker.finished.connect(self._inspect_worker.deleteLater)
        self._inspect_thread.finished.connect(self._inspect_thread.deleteLater);self._inspect_thread.finished.connect(self._on_background_thread_finished);self._inspect_thread.start()
    def _on_capabilities(self,capabilities):
        self.capabilities=capabilities
        self.query_model.set_supported_load_modes(capabilities.supported_load_modes)
        self.load_mode_delegate.supported_modes=tuple(capabilities.supported_load_modes)+('preserve_existing',)
        self.capability_label.setText(f'Office {capabilities.office_version}, build {capabilities.build}; supported: {", ".join(capabilities.supported_load_modes) or "none"}.')
        if capabilities.reasons:
            self.capability_label.setToolTip('\n'.join(f'{mode}: {reason}' for mode,reason in capabilities.reasons.items()))
        self.statusBar().showMessage('Excel capability probe complete; unsupported modes remain disabled.')
    def _on_capability_probe_failed(self,message):
        self.capability_label.setText('Excel capability probe failed; additional modes are disabled.')
        self.capability_label.setToolTip(message);self.statusBar().showMessage('Excel capability probe failed')
        QMessageBox.critical(self,'Excel capability probe failed',message)
    def _on_inspected(self,inventory):
        self.inventory=tuple(inventory)
        self._record_history('workbook_inspected',{'query_count':len(self.inventory)})
        self.query_model.apply_inventory(self.inventory,skip_existing=self.mode_combo.currentData()==ExistingQueryMode.SKIP.value)
        dialog=QDialog(self);dialog.setWindowTitle(f'Workbook query inventory ({len(self.inventory)})');dialog.resize(1000,600)
        layout=QVBoxLayout(dialog);view=QTableView(dialog);view.setAccessibleName('Workbook query inventory');view.setModel(QueryInventoryModel(self.inventory,view));view.setAlternatingRowColors(True)
        view.setSelectionBehavior(QAbstractItemView.SelectRows);view.setWordWrap(False)
        view.horizontalHeader().setSectionResizeMode(0,QHeaderView.ResizeToContents);view.horizontalHeader().setSectionResizeMode(1,QHeaderView.Stretch)
        view.horizontalHeader().setSectionResizeMode(2,QHeaderView.ResizeToContents);view.horizontalHeader().setSectionResizeMode(3,QHeaderView.ResizeToContents)
        layout.addWidget(view)
        actions=QHBoxLayout();export_selected=QPushButton('Export selected…',dialog);export_selected.setAccessibleName('Export selected workbook queries');export_all=QPushButton('Export all…',dialog);export_all.setAccessibleName('Export all workbook queries')
        export_selected_package=QPushButton('Package selected…',dialog);export_selected_package.setAccessibleName('Package selected workbook queries');export_all_package=QPushButton('Package all…',dialog);export_all_package.setAccessibleName('Package all workbook queries')
        refresh_selected=QPushButton('Refresh selected…',dialog);refresh_selected.setAccessibleName('Refresh selected workbook queries');refresh_all=QPushButton('Refresh all…',dialog);refresh_all.setAccessibleName('Refresh all workbook queries')
        export_selected.clicked.connect(lambda:self._export_inventory(view,selected_only=True));export_all.clicked.connect(lambda:self._export_inventory(view,selected_only=False))
        export_selected_package.clicked.connect(lambda:self._export_inventory_package(view,selected_only=True))
        export_all_package.clicked.connect(lambda:self._export_inventory_package(view,selected_only=False))
        refresh_selected.clicked.connect(lambda:self._refresh_inventory(dialog,view,selected_only=True))
        refresh_all.clicked.connect(lambda:self._refresh_inventory(dialog,view,selected_only=False))
        actions.addWidget(export_selected);actions.addWidget(export_all);actions.addWidget(export_selected_package);actions.addWidget(export_all_package)
        actions.addWidget(refresh_selected);actions.addWidget(refresh_all);actions.addStretch()
        close=QPushButton('Close',dialog);close.setAccessibleName('Close workbook query inventory');close.clicked.connect(dialog.accept);actions.addWidget(close);layout.addLayout(actions)
        self.statusBar().showMessage(f'Inspected {len(self.inventory)} queries; review each observed load status before refresh.')
        dialog.exec()

    def _refresh_inventory(self,dialog,view,selected_only):
        if selected_only:
            rows=sorted({index.row() for index in view.selectionModel().selectedRows()})
            if not rows:
                QMessageBox.information(view,'Select queries','Select one or more workbook query rows to refresh.')
                return
            names=[self.inventory[row].name for row in rows]
            scope=f'{len(names)} selected query(s)'
        else:
            names=None
            scope=f'all {len(self.inventory)} workbook query(s)'
        options=QDialog(dialog);options.setWindowTitle('Refresh options')
        layout=QVBoxLayout(options)
        notice=QLabel('Refreshing may access external data sources and change query results. It runs only after you confirm. The workbook is refreshed in a staged copy and committed only if every selected refresh succeeds.')
        notice.setWordWrap(True);layout.addWidget(notice)
        mode_row=QHBoxLayout();mode_row.addWidget(QLabel('Excel query mode:'))
        mode=QComboBox(options);mode.setAccessibleName('Excel refresh behavior');mode.addItem('Background refresh (wait for completion)',True);mode.addItem('Synchronous refresh',False);mode_row.addWidget(mode);mode_row.addStretch();layout.addLayout(mode_row)
        timeout_row=QHBoxLayout();timeout_row.addWidget(QLabel('Timeout per query (seconds):'))
        timeout=QSpinBox(options);timeout.setAccessibleName('Refresh timeout per query in seconds');timeout.setRange(1,7200);timeout.setValue(300);timeout_row.addWidget(timeout);timeout_row.addStretch();layout.addLayout(timeout_row)
        backup=QCheckBox('Create timestamped backup before refresh',options);backup.setAccessibleName('Create a timestamped workbook backup before refresh');backup.setChecked(True);layout.addWidget(backup)
        buttons=QHBoxLayout();start=QPushButton('Continue to confirmation',options);start.setAccessibleName('Continue to refresh confirmation');cancel=QPushButton('Cancel',options);cancel.setAccessibleName('Cancel refresh options')
        start.clicked.connect(options.accept);cancel.clicked.connect(options.reject);buttons.addStretch();buttons.addWidget(start);buttons.addWidget(cancel);layout.addLayout(buttons)
        if options.exec()!=QDialog.Accepted:return
        asynchronous=bool(mode.currentData());timeout_seconds=timeout.value();make_backup=backup.isChecked()
        answer=QMessageBox.question(dialog,'Confirm external data refresh',
            f'Refresh {scope} in this workbook now? Excel may contact external data sources. A successful operation replaces the workbook with the staged refreshed copy.',
            QMessageBox.Yes|QMessageBox.Cancel,QMessageBox.Cancel)
        if answer!=QMessageBox.Yes:return
        dialog.accept()
        try:
            workbook=validate_workbook(self.workbook_edit.text().strip())
            self._set_busy(True,'Refreshing…')
            self.statusBar().showMessage('Refreshing selected workbook queries…')
            self._refresh_thread=QThread(self)
            history_store=self._get_history_store();snapshot_store=history_store if history_store.snapshots_enabled else None
            self._refresh_worker=RefreshWorker(workbook,names,timeout_seconds=timeout_seconds,
                                               asynchronous=asynchronous,make_backup=make_backup,
                                               history_store=snapshot_store)
            self._refresh_worker.moveToThread(self._refresh_thread)
            self._refresh_thread.started.connect(self._refresh_worker.run)
            self._refresh_worker.log.connect(self.log.appendPlainText)
            self._refresh_worker.completed.connect(self._on_refresh_completed)
            self._refresh_worker.failed.connect(self._on_refresh_failed)
            self._refresh_worker.finished.connect(self._refresh_thread.quit)
            self._refresh_worker.finished.connect(self._refresh_worker.deleteLater)
            self._refresh_thread.finished.connect(self._refresh_thread.deleteLater)
            self._refresh_thread.finished.connect(self._on_background_thread_finished)
            self._refresh_thread.start()
        except Exception as exc:
            self._record_diagnostic('refresh_start',exc)
            self._set_busy(False)
            QMessageBox.critical(self,'Could not start refresh',f'{type(exc).__name__}: {exc}')

    def _on_refresh_completed(self,outcome):
        if isinstance(outcome,dict):results=outcome.get('results',());snapshot_id=outcome.get('snapshot_id')
        else:results=outcome;snapshot_id=None
        details={'query_count':len(results),'refreshed':sum(result.status=='refreshed' for result in results)}
        if snapshot_id:details['snapshot_id']=snapshot_id
        self._record_history('refresh_completed',details)
        refreshed=sum(result.status=='refreshed' for result in results)
        skipped=sum(result.status=='skipped' for result in results)
        self.statusBar().showMessage('Refresh completed; review the operation log.')
        QMessageBox.information(self,'Refresh complete',
            f'Refreshed {refreshed} query(s); skipped {skipped}. Review per-query details in the operation log.')

    def _on_refresh_failed(self,message,results):
        self._record_history('refresh_failed',{'query_count':len(results)})
        self.statusBar().showMessage('Refresh failed; original workbook retained.')
        for result in results:
            self.log.appendPlainText(f'{result.query_name}: {result.status} - {result.message}')
        self.log.appendPlainText(f'REFRESH FAILED: {message}')
        QMessageBox.critical(self,'Refresh failed',
            f'Excel could not complete the refresh. The original workbook was retained.\n\n{message}')
    def _export_inventory(self,view,selected_only):
        if selected_only:
            rows=sorted({index.row() for index in view.selectionModel().selectedRows()})
            if not rows:QMessageBox.information(view,'Select queries','Select one or more query rows to export.');return
            selected=[self.inventory[row] for row in rows]
        else:selected=list(self.inventory)
        folder=QFileDialog.getExistingDirectory(view,'Choose export folder')
        if not folder:return
        try:written=export_queries(selected,folder)
        except Exception as exc:self._record_diagnostic('export',exc);QMessageBox.critical(view,'Export failed',f'{type(exc).__name__}: {exc}');return
        self._record_history('queries_exported',{'query_count':len(written)})
        QMessageBox.information(view,'Queries exported',f'Exported {len(written)} formula(s) and a query-name manifest to:\n{folder}')
    def _export_inventory_package(self,view,selected_only):
        if selected_only:
            rows=sorted({index.row() for index in view.selectionModel().selectedRows()})
            if not rows:QMessageBox.information(view,'Select queries','Select one or more query rows to package.');return
            selected=[self.inventory[row] for row in rows]
        else:selected=list(self.inventory)
        path,_=QFileDialog.getSaveFileName(view,'Export versioned query package','','Power Query package (*.pqm.zip)')
        if not path:return
        if not path.casefold().endswith('.zip'):path += '.pqm.zip'
        try:package=export_query_package(selected,path)
        except Exception as exc:self._record_diagnostic('package_export',exc);QMessageBox.critical(view,'Package export failed',f'{type(exc).__name__}: {exc}');return
        exported_count=sum(query.formula is not None for query in selected)
        self._record_history('query_package_exported',{'query_count':exported_count})
        QMessageBox.information(view,'Package exported',f'Created a versioned, hash-verified package with {exported_count} query formula(s):\n{package}')
    def _on_inspect_failed(self,message):
        self.statusBar().showMessage('Inspection failed')
        QMessageBox.critical(self,'Workbook inspection failed',message)
    def _save_profile(self):
        self._update_source_list()
        if not self.sources:QMessageBox.information(self,'No queries','Add at least one query before saving a profile.');return
        if not self.workbook_edit.text().strip():QMessageBox.warning(self,'Choose a workbook','Select a workbook path before saving a profile.');return
        path,_=QFileDialog.getSaveFileName(self,'Save import profile','','Power Query profiles (*.json)')
        if not path:return
        if Path(path).exists() and QMessageBox.question(self,'Replace profile?','This profile file exists. Replace it?',QMessageBox.Yes|QMessageBox.Cancel,QMessageBox.Cancel)!=QMessageBox.Yes:return
        try:
            profile=build_import_profile(Path(path).stem,self.workbook_edit.text().strip(),self.sources,
                                         [self.query_model.stable_id_at(i) for i in range(self.query_model.rowCount())],
                                         base_dir=Path(path).parent,existing_query_policy=self.mode_combo.currentData(),
                                         backup=self.backup_check.isChecked(),dry_run=self.dry_run_check.isChecked())
            if self._loaded_profile_path_mappings:profile['path_mappings']=self._loaded_profile_path_mappings
            if self._loaded_profile_environment_mappings:profile['environment_mappings']=self._loaded_profile_environment_mappings
            save_profile(profile,path)
        except Exception as exc:self._record_diagnostic('profile_save',exc);QMessageBox.critical(self,'Profile save failed',f'{type(exc).__name__}: {exc}');return
        self._record_history('profile_saved',{'profile_name':profile['profile_name'],'query_count':len(profile['queries'])})
        QMessageBox.information(self,'Profile saved',f'Saved profile to:\n{path}')

    def _edit_source_mappings(self):
        dialog=ProfileMappingsDialog(
            [dict(row) for row in self._loaded_profile_path_mappings],
            [dict(row) for row in self._loaded_profile_environment_mappings],self)
        if dialog.exec()!=QDialog.DialogCode.Accepted:return
        try:
            path_mappings,environment_mappings=dialog.mappings()
        except ValueError as exc:
            QMessageBox.warning(self,'Source mappings are invalid',str(exc));return
        self._loaded_profile_path_mappings=path_mappings
        self._loaded_profile_environment_mappings=environment_mappings
        self._record_history('profile_source_mappings_edited',{
            'path_mapping_count':len(path_mappings),
            'environment_mapping_count':len(environment_mappings)})
        self.statusBar().showMessage('Source mappings updated in memory; save a profile to persist them.')
    def _load_profile(self):
        path,_=QFileDialog.getOpenFileName(self,'Load import profile','','Power Query profiles (*.json)')
        if not path:return
        try:
            profile=load_profile(path)
            apply_mappings=False
            apply_environment_mappings=False
            environment_previews=[]
            if any(mapping['enabled'] for mapping in profile.get('path_mappings',[])):
                previews=preview_profile_path_mappings(profile,path)
                if any(row.ambiguous for row in previews):
                    raise ValueError('Path mappings are ambiguous for one or more sources. Resolve equal-priority mapping conflicts and try again.')
                changed=[row for row in previews if row.changed]
                details='\n'.join(
                    f'{row.query_id}: {row.source_path} -> {row.mapped_path} '
                    f'({"target exists" if row.target_exists else "target missing"})'
                    for row in changed
                ) or 'No source paths match the enabled mapping roots; all paths will stay unchanged.'
                prompt=QMessageBox(self);prompt.setWindowTitle('Preview profile path mappings')
                prompt.setText(f'{len(changed)} source path(s) will change. Apply these mappings?')
                prompt.setDetailedText(details);prompt.setIcon(QMessageBox.Question)
                apply_button=prompt.addButton('Apply mappings',QMessageBox.AcceptRole)
                keep_button=prompt.addButton('Keep original paths',QMessageBox.RejectRole)
                prompt.setDefaultButton(keep_button)
                prompt.exec()
                apply_mappings=prompt.clickedButton() is apply_button
                if not apply_mappings:return
            if any(mapping['enabled'] for mapping in profile.get('environment_mappings',[])):
                environment_previews=preview_profile_environment_mappings(profile,path)
                conflicts=[item for item in environment_previews if item.ambiguous]
                if conflicts:
                    raise ValueError('M-code environment mappings are ambiguous. Resolve conflicting mappings before loading this profile.')
                changed=[item for item in environment_previews if item.changed]
                details=[]
                for item in environment_previews:
                    if item.changed:
                        details.append(f'--- {item.query_name}: mapped M code ---\n'+diff_m_code(
                            item.original_formula,item.mapped_formula,source_name='Profile source',workbook_name='Mapped preview'))
                    for warning in item.warnings:details.append(f'--- {item.query_name}: warning ---\n{warning}')
                if not details:details=['No recognized literal source paths matched the enabled mappings.']
                prompt=QMessageBox(self);prompt.setWindowTitle('Preview M-code environment mappings')
                prompt.setText(f'{len(changed)} query formula(s) will be rewritten in temporary copies. Apply these changes?')
                prompt.setDetailedText('\n\n'.join(details));prompt.setIcon(QMessageBox.Question)
                apply_button=prompt.addButton('Apply M mappings',QMessageBox.AcceptRole)
                keep_button=prompt.addButton('Keep original paths',QMessageBox.RejectRole)
                prompt.setDefaultButton(keep_button);prompt.exec()
                apply_environment_mappings=prompt.clickedButton() is apply_button
                if not apply_environment_mappings:return
            workbook,sources,stable_ids=profile_to_sources(
                profile,path,supported_load_modes=self.query_model.supported_load_modes,
                apply_path_mappings=apply_mappings,apply_environment_mappings=apply_environment_mappings)
            if apply_environment_mappings:
                mapped_paths,staging=materialize_profile_environment_mapping_copies(environment_previews)
                if staging is not None:self._staged_import_dirs.append(staging)
                for source,stable_id in zip(sources,stable_ids):
                    if stable_id in mapped_paths:
                        source.profile_source_path=source.path
                        source.path=mapped_paths[stable_id]
        except Exception as exc:self._record_diagnostic('profile_load',exc);QMessageBox.critical(self,'Profile rejected',f'{type(exc).__name__}: {exc}');return
        self._loaded_profile_path_mappings=list(profile.get('path_mappings',[]))
        self._loaded_profile_environment_mappings=list(profile.get('environment_mappings',[]))
        self.query_model.replace_sources(sources,stable_ids);self._update_source_list()
        self._record_history('profile_loaded',{'profile_name':profile['profile_name'],'query_count':len(sources)})
        self.workbook_edit.setText(str(workbook))
        self.backup_check.setChecked(profile['options']['backup'])
        self.dry_run_check.setChecked(profile['options']['dry_run'])
        self.mode_combo.setCurrentIndex(self.mode_combo.findData(profile['options']['existing_query_policy']))
        self.inventory=()
        self._update_sequence();self._validate_rows()
        QMessageBox.information(self,'Profile loaded',f'Loaded {len(sources)} source row(s). Review the paths and query names before applying.')
    def _open_workbook(self):
        path=self.workbook_edit.text().strip()
        if not path or not Path(path).exists(): QMessageBox.warning(self,'Workbook not found','Select a valid workbook first.'); return
        try: os.startfile(path)
        except Exception as exc:self._record_diagnostic('open_workbook',exc);QMessageBox.critical(self,'Cannot open workbook',str(exc))

    def _get_history_store(self):
        store=getattr(self,'history_store',None)
        if store is None:
            root=Path(os.environ.get('LOCALAPPDATA',Path.home()/'AppData'/'Local'))/'PowerQueryManager'
            store=HistoryStore(root/'history.sqlite3')
            self.history_store=store
            if store.recovery_archive is not None:
                self.log.appendPlainText(
                    'The local history database was unreadable. Its original bytes were preserved at '
                    f'{store.recovery_archive}; a new local history has been started.')
            if store.recovered_snapshot_files:
                self.log.appendPlainText(
                    'Unindexed workbook snapshots from an interrupted history write were preserved for manual '
                    f'review in {store.recovered_snapshot_files[0].parent} '
                    f'({len(store.recovered_snapshot_files)} file(s)); they were not deleted or added to history.')
        return store

    def _record_history(self,event,details):
        try:
            self._get_history_store().record(event,details)
        except Exception as exc:
            from pq_installer.core.redaction import safe_error
            self.log.appendPlainText(f'History could not be saved: {safe_error(exc)}')

    def _open_history(self):
        try:store=self._get_history_store()
        except Exception as exc:
            QMessageBox.warning(self,'History unavailable',f'Local history could not be opened: {exc}')
            return
        dialog=QDialog(self);dialog.setWindowTitle('Local operation history');dialog.resize(850,460)
        layout=QVBoxLayout(dialog)
        hint=QLabel('History metadata omits formulas and redacts recognized secrets. Optional workbook snapshots contain the full local workbook, may include sensitive data, and stay in this Windows user profile. Retention and deletion controls apply to snapshots too.')
        hint.setAccessibleName('History privacy and snapshot information')
        hint.setWordWrap(True);layout.addWidget(hint)
        table=QTableWidget(0,4,dialog);table.setHorizontalHeaderLabels(['Version ID','UTC time','Action','Details'])
        table.setAccessibleName('Local operation history entries')
        table.setEditTriggers(QAbstractItemView.NoEditTriggers);table.setSelectionBehavior(QAbstractItemView.SelectRows)
        table.setSelectionMode(QAbstractItemView.ExtendedSelection);table.setWordWrap(True);layout.addWidget(table,1)
        retention=QSpinBox(dialog);retention.setRange(1,3650);retention.setSuffix(' days');retention.setValue(store.retention_days)
        retention.setAccessibleName('History retention period in days')
        snapshot_toggle=QCheckBox('Capture local snapshots before import and refresh',dialog)
        snapshot_toggle.setAccessibleName('Capture optional full workbook snapshots before changes')
        snapshot_toggle.setToolTip('Copies the full workbook into local history before changes. No snapshot is made unless enabled.')
        snapshot_toggle.setChecked(store.snapshots_enabled)
        refresh=QPushButton('Refresh');refresh.setAccessibleName('Refresh history entries')
        apply_retention=QPushButton('Apply retention');apply_retention.setAccessibleName('Apply history retention period')
        compare=QPushButton('Compare selected snapshots…');compare.setAccessibleName('Compare selected workbook versions');compare.setEnabled(False)
        restore=QPushButton('Restore selected version…');restore.setAccessibleName('Restore selected workbook version')
        delete=QPushButton('Delete selected');delete.setAccessibleName('Delete selected history entry')
        clear=QPushButton('Delete all history');clear.setAccessibleName('Delete all local history')
        close=QPushButton('Close');close.setAccessibleName('Close operation history')
        layout.addWidget(snapshot_toggle)
        controls=QHBoxLayout();controls.addWidget(QLabel('Keep entries for:'));controls.addWidget(retention);controls.addWidget(apply_retention);controls.addStretch()
        controls.addWidget(refresh);controls.addWidget(compare);controls.addWidget(restore);controls.addWidget(delete);controls.addWidget(clear);controls.addWidget(close);layout.addLayout(controls)
        def reload_rows():
            rows=store.list(limit=1000);table.setRowCount(len(rows))
            for row_index,row in enumerate(rows):
                values=(row['version_id'],row['occurred_at'],row['event'],json.dumps(row['details'],ensure_ascii=False,sort_keys=True))
                for column,value in enumerate(values):
                    item=QTableWidgetItem(value)
                    if column==0:item.setData(Qt.UserRole,row['version_id'])
                    table.setItem(row_index,column,item)
            table.resizeColumnsToContents();table.horizontalHeader().setStretchLastSection(True)
            update_restore_button()
        def selected_snapshot_id():
            row=table.currentRow()
            if row<0:return None
            try:return json.loads(table.item(row,3).text()).get('snapshot_id')
            except (AttributeError,TypeError,json.JSONDecodeError):return None
        def selected_snapshot_ids():
            ids=[]
            for row in sorted({index.row() for index in table.selectionModel().selectedRows()}):
                try:
                    snapshot_id=json.loads(table.item(row,3).text()).get('snapshot_id')
                    occurred_at=table.item(row,1).text()
                    if snapshot_id:ids.append((occurred_at,snapshot_id))
                except Exception:continue
            # History rows are displayed newest first. Compare always means
            # earlier -> later, independent of visual row or click order.
            return [snapshot_id for _occurred_at,snapshot_id in sorted(ids)]
        def update_compare_button():compare.setEnabled(len(selected_snapshot_ids())==2)
        def update_restore_button():
            snapshot_id=selected_snapshot_id()
            if not snapshot_id:restore.setEnabled(False);return
            try:store.get_snapshot(snapshot_id);restore.setEnabled(True)
            except Exception:restore.setEnabled(False)
        def toggle_snapshots(enabled):
            try:store.set_snapshots_enabled(bool(enabled))
            except Exception as exc:
                QMessageBox.warning(dialog,'Snapshot setting not changed',str(exc))
                snapshot_toggle.setChecked(store.snapshots_enabled)
        def restore_selected():
            snapshot_id=selected_snapshot_id()
            if not snapshot_id:return
            try:
                snapshot=store.get_snapshot(snapshot_id);target=Path(snapshot['workbook_path'])
                exists=target.exists();expected=None
                prompt=(f"Restore the saved version to {target}?\n\n"
                    +(f'A verified timestamped backup of the current workbook will be created. Close the workbook in Excel before continuing.' if exists else 'This will create a workbook at the recorded original path.')
                    +'\n\nThe snapshot is hash-verified. Continue?')
                if QMessageBox.question(dialog,'Restore workbook version',prompt,QMessageBox.Yes|QMessageBox.Cancel,QMessageBox.Cancel)!=QMessageBox.Yes:return
                restore.setEnabled(False)
                thread=QThread(self);worker=SnapshotRestoreWorker(store,snapshot_id);worker.moveToThread(thread)
                self._history_restore_thread=thread;self._history_restore_worker=worker
                def restored(outcome):
                    if dialog.isVisible():
                        reload_rows()
                        message=f"Restored {Path(outcome['restored_path']).name}."
                        if outcome['backup_path']:message+=f"\nBackup: {outcome['backup_path']}"
                        if outcome.get('history_warning'):message+='\nHistory event warning: '+outcome['history_warning']
                        QMessageBox.information(dialog,'Version restored',message)
                def restore_failed(message):
                    if dialog.isVisible():QMessageBox.warning(dialog,'Restore stopped',message)
                def restore_thread_finished():
                    if dialog.isVisible():restore.setEnabled(True)
                    self._history_restore_thread=None
                    self._history_restore_worker=None
                    self._history_restore_bridge=None
                bridge=_HistoryWorkerUiBridge(restored,restore_failed,restore_thread_finished,dialog)
                self._history_restore_bridge=bridge
                worker.completed.connect(bridge.completed,Qt.ConnectionType.AutoConnection)
                worker.failed.connect(bridge.failed,Qt.ConnectionType.AutoConnection)
                worker.finished.connect(thread.quit);worker.finished.connect(worker.deleteLater)
                thread.finished.connect(bridge.thread_finished,Qt.ConnectionType.AutoConnection)
                thread.finished.connect(thread.deleteLater)
                thread.started.connect(worker.run);thread.start()
            except Exception as exc:QMessageBox.warning(dialog,'Restore stopped',str(exc))
        def compare_selected():
            ids=selected_snapshot_ids()
            if len(ids)!=2:return
            try:
                compare.setEnabled(False)
                thread=QThread(self);worker=SnapshotCompareWorker(store,ids[0],ids[1])
                worker.moveToThread(thread);self._history_compare_thread=thread;self._history_compare_worker=worker
                def compared(outcome):
                    self._record_history('workbook_snapshots_compared',{
                        'action':'snapshot_comparison_completed',
                        'before_snapshot_id':ids[0],
                        'after_snapshot_id':ids[1],
                        'comparison':'completed',
                    })
                    if dialog.isVisible():
                        result=HistoryComparisonDialog(outcome['before'],outcome['after'],outcome['changes'],dialog)
                        result.exec()
                def compare_failed(message):
                    if dialog.isVisible():QMessageBox.warning(dialog,'Version comparison stopped',message)
                def compare_thread_finished():
                    if dialog.isVisible():update_compare_button()
                    self._history_compare_thread=None
                    self._history_compare_worker=None
                    self._history_compare_bridge=None
                bridge=_HistoryWorkerUiBridge(compared,compare_failed,compare_thread_finished,dialog)
                self._history_compare_bridge=bridge
                worker.completed.connect(bridge.completed,Qt.ConnectionType.AutoConnection)
                worker.failed.connect(bridge.failed,Qt.ConnectionType.AutoConnection)
                worker.finished.connect(thread.quit);worker.finished.connect(worker.deleteLater)
                thread.finished.connect(bridge.thread_finished,Qt.ConnectionType.AutoConnection)
                thread.finished.connect(thread.deleteLater)
                thread.started.connect(worker.run);thread.start()
            except Exception as exc:
                compare.setEnabled(True);QMessageBox.warning(dialog,'Version comparison stopped',str(exc))
        def apply_days():
            try:store.set_retention_days(retention.value());reload_rows()
            except Exception as exc:QMessageBox.warning(dialog,'Retention not changed',str(exc))
        def delete_selected():
            row=table.currentRow()
            if row<0:return
            version=table.item(row,0).data(Qt.UserRole)
            if QMessageBox.question(dialog,'Delete history entry?','Delete the selected history entry?',QMessageBox.Yes|QMessageBox.Cancel,QMessageBox.Cancel)!=QMessageBox.Yes:return
            store.delete(version);reload_rows()
        def clear_all():
            if QMessageBox.question(dialog,'Delete all history?','Permanently delete all local history entries?',QMessageBox.Yes|QMessageBox.Cancel,QMessageBox.Cancel)!=QMessageBox.Yes:return
            store.clear();reload_rows()
        refresh.clicked.connect(reload_rows);apply_retention.clicked.connect(apply_days)
        snapshot_toggle.toggled.connect(toggle_snapshots);table.itemSelectionChanged.connect(update_restore_button)
        table.itemSelectionChanged.connect(update_compare_button)
        compare.clicked.connect(compare_selected)
        restore.clicked.connect(restore_selected)
        delete.clicked.connect(delete_selected);clear.clicked.connect(clear_all);close.clicked.connect(dialog.accept)
        reload_rows();dialog.exec()

    def closeEvent(self,event):
        active=[thread for thread in (self._thread,self._inspect_thread,self._refresh_thread,self._recovery_thread,self._history_restore_thread,
            getattr(self,'_history_compare_thread',None)) if thread is not None and thread.isRunning()]
        if active:
            QMessageBox.warning(self,'Operation still running','Wait for the current Excel operation to finish before closing the manager.')
            event.ignore();return
        for staged in self._staged_import_dirs:
            try:
                target=Path(staged).resolve()
                if target.name.startswith(('pq-import-','pq-map-')) and target.parent==Path(__import__('tempfile').gettempdir()).resolve():
                    shutil.rmtree(target,ignore_errors=True)
            except OSError:pass
        event.accept()
