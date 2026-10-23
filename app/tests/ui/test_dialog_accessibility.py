import os
os.environ.setdefault('QT_QPA_PLATFORM','offscreen')

from pathlib import Path

from PySide6.QtWidgets import (QApplication,QAbstractButton,QComboBox,QGraphicsView,
    QLineEdit,QPlainTextEdit,QSpinBox,QTableView,QTabWidget)

from pq_installer.core.models import QuerySource
from pq_installer.ui.dependency_graph_dialog import DependencyGraphDialog
from pq_installer.ui.history_comparison_dialog import HistoryComparisonDialog
from pq_installer.ui.package_update_dialog import QueryPackageUpdateDialog
from pq_installer.ui.profile_mappings_dialog import ProfileMappingsDialog
from pq_installer.ui.query_inspector_dialog import QueryInspectorDialog

app=QApplication.instance() or QApplication([])


def test_visible_enabled_controls_in_query_dialogs_have_accessible_names(tmp_path):
    source_path=tmp_path/'query.pq';source_path.write_text('let Source = 1 in Source')
    source=QuerySource(source_path,'Query')
    dialogs=(
        QueryInspectorDialog(source,None),
        QueryPackageUpdateDialog('current.pqm.zip','updated.pqm.zip',()),
        ProfileMappingsDialog(),
        HistoryComparisonDialog({'occurred_at':'before','workbook_name':'Book'},
                                {'occurred_at':'after','workbook_name':'Book'},()),
        DependencyGraphDialog(()),
    )
    widget_types=(QAbstractButton,QComboBox,QGraphicsView,QLineEdit,QPlainTextEdit,
                  QSpinBox,QTableView,QTabWidget)
    missing=[]
    for dialog in dialogs:
        dialog.show();app.processEvents()
        for widget_type in widget_types:
            for widget in dialog.findChildren(widget_type):
                if widget.isVisibleTo(dialog) and widget.isEnabled() and not widget.accessibleName().strip():
                    missing.append((dialog.windowTitle(),widget.metaObject().className(),
                                    widget.text() if hasattr(widget,'text') else ''))
        dialog.close()
    assert not missing, f'Visible enabled dialog controls lack accessible names: {missing}'
