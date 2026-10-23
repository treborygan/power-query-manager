import os
os.environ.setdefault('QT_QPA_PLATFORM','offscreen')

from PySide6.QtWidgets import (QApplication,QAbstractButton,QCheckBox,QComboBox,
    QDialog,QSpinBox,QTableView)

from pq_installer.core.manager import QueryInfo
from pq_installer.ui.main_window import MainWindow

app=QApplication.instance() or QApplication([])


def test_workbook_inventory_dialog_has_accessible_table_and_actions(tmp_path,monkeypatch):
    monkeypatch.setenv('LOCALAPPDATA',str(tmp_path/'local-app-data'))
    captured={}
    monkeypatch.setattr(QDialog,'exec',lambda dialog:(captured.setdefault('dialog',dialog) and 0))
    window=MainWindow()
    try:
        window._on_inspected((QueryInfo('Query A','1',load_state='connection_only'),))
        dialog=captured['dialog']
        table=dialog.findChild(QTableView)
        assert table.accessibleName()=='Workbook query inventory'
        buttons={button.text():button for button in dialog.findChildren(QAbstractButton)}
        expected={
            'Export selected…':'Export selected workbook queries',
            'Export all…':'Export all workbook queries',
            'Package selected…':'Package selected workbook queries',
            'Package all…':'Package all workbook queries',
            'Refresh selected…':'Refresh selected workbook queries',
            'Refresh all…':'Refresh all workbook queries',
            'Close':'Close workbook query inventory',
        }
        for label,name in expected.items():
            assert buttons[label].accessibleName()==name
    finally:
        if captured.get('dialog') is not None:captured['dialog'].close()
        window.close()


def test_refresh_options_have_accessible_names(tmp_path,monkeypatch):
    monkeypatch.setenv('LOCALAPPDATA',str(tmp_path/'local-app-data'))
    captured={}
    monkeypatch.setattr(QDialog,'exec',lambda dialog:(captured.setdefault('dialog',dialog) and QDialog.Rejected))
    window=MainWindow()
    try:
        window.inventory=(QueryInfo('Query A','1',load_state='connection_only'),)
        parent=QDialog(window)
        window._refresh_inventory(parent,None,selected_only=False)
        options=captured['dialog']
        widgets=[]
        for widget_type in (QComboBox,QSpinBox,QCheckBox,QAbstractButton):
            widgets.extend(options.findChildren(widget_type))
        controls={widget.accessibleName():widget for widget in widgets if widget.accessibleName()}
        expected={'Excel refresh behavior','Refresh timeout per query in seconds',
          'Create a timestamped workbook backup before refresh','Continue to refresh confirmation',
          'Cancel refresh options'}
        assert expected.issubset(controls)
    finally:
        if captured.get('dialog') is not None:captured['dialog'].close()
        window.close()
