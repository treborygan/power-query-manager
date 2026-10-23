import os
os.environ.setdefault('QT_QPA_PLATFORM','offscreen')

from PySide6.QtWidgets import QApplication, QMessageBox
from pq_installer.core.history import HistoryStore
from pq_installer.core.models import InstallResult, QueryStatus
from pq_installer.ui.main_window import MainWindow

app=QApplication.instance() or QApplication([])


def test_completed_import_history_records_outcome_counts_without_formula_or_path(tmp_path,monkeypatch):
    store=HistoryStore(tmp_path/'history.sqlite3')
    window=MainWindow();window.history_store=store
    window.workbook_edit.setText(str(tmp_path/'report.xlsx'))
    monkeypatch.setattr(QMessageBox,'information',lambda *args:None)
    results=(
        InstallResult('CreatedQuery',QueryStatus.CREATED,'created'),
        InstallResult('UpdatedQuery',QueryStatus.REPLACED,'replaced'),
        InstallResult('SkippedQuery',QueryStatus.SKIPPED,'skipped'),
    )

    try:
        window._on_completed({'results':results,'snapshot_id':'snapshot-id'})
        event=next(row for row in store.list() if row['event']=='import_completed')
        assert event['details']=={
            'query_count':3,'workbook':'report.xlsx','snapshot_id':'snapshot-id',
            'action':'import','status':'completed',
            'created_count':1,'replaced_count':1,'skipped_count':1,
        }
        assert str(tmp_path) not in str(event['details'])
    finally:
        window.close()
