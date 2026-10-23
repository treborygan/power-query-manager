import os
os.environ.setdefault('QT_QPA_PLATFORM', 'offscreen')

from PySide6.QtWidgets import QApplication, QDialog, QTableWidget, QSpinBox, QCheckBox, QPushButton, QLabel
from pq_installer.core.history import HistoryStore
from pq_installer.ui.main_window import MainWindow

app = QApplication.instance() or QApplication([])


def test_history_dialog_has_accessible_names_for_table_and_controls(tmp_path, monkeypatch):
    window = MainWindow()
    window.history_store = HistoryStore(tmp_path / 'history.sqlite3')
    captured = {}

    def capture(dialog):
        captured['dialog'] = dialog
        return 0

    monkeypatch.setattr(QDialog, 'exec', capture)
    try:
        window._open_history()
        dialog = captured['dialog']
        hint = dialog.findChild(QLabel)
        table = dialog.findChild(QTableWidget)
        retention = dialog.findChild(QSpinBox)
        snapshots = dialog.findChild(QCheckBox)
        buttons = {button.text(): button for button in dialog.findChildren(QPushButton)}
        assert hint.accessibleName() == 'History privacy and snapshot information'
        assert table.accessibleName() == 'Local operation history entries'
        assert retention.accessibleName() == 'History retention period in days'
        assert snapshots.accessibleName() == 'Capture optional full workbook snapshots before changes'
        expected = {
            'Refresh': 'Refresh history entries',
            'Apply retention': 'Apply history retention period',
            'Compare selected snapshots…': 'Compare selected workbook versions',
            'Restore selected version…': 'Restore selected workbook version',
            'Delete selected': 'Delete selected history entry',
            'Delete all history': 'Delete all local history',
            'Close': 'Close operation history',
        }
        for label, accessible_name in expected.items():
            assert buttons[label].accessibleName() == accessible_name
    finally:
        if 'dialog' in captured:
            captured['dialog'].close()
        window.close()

def test_corrupt_history_recovery_is_reported_in_ui(tmp_path, monkeypatch):
    from pq_installer.ui.main_window import MainWindow

    root = tmp_path / 'local-app-data'
    app_dir = root / 'PowerQueryManager'
    app_dir.mkdir(parents=True)
    (app_dir / 'history.sqlite3').write_bytes(b'corrupt history')
    monkeypatch.setenv('LOCALAPPDATA', str(root))
    window = MainWindow()
    try:
        store = window._get_history_store()
        assert store.recovery_archive is not None
        assert str(store.recovery_archive) in window.log.toPlainText()
        assert 'new local history' in window.log.toPlainText().casefold()
    finally:
        window.close()

def test_unindexed_snapshot_recovery_is_reported_in_ui_without_trusting_file(tmp_path, monkeypatch):
    import uuid

    root = tmp_path / 'local-app-data'
    app_dir = root / 'PowerQueryManager'
    snapshots = app_dir / 'history_snapshots'
    snapshots.mkdir(parents=True)
    orphan = snapshots / f'{uuid.uuid4()}.xlsx'
    original = b'disposable snapshot bytes for recovery'
    orphan.write_bytes(original)
    monkeypatch.setenv('LOCALAPPDATA', str(root))
    window = MainWindow()
    try:
        store = window._get_history_store()
        recovered = store.recovered_snapshot_files
        assert len(recovered) == 1
        assert recovered[0].read_bytes() == original
        assert str(recovered[0].parent) in window.log.toPlainText()
        assert 'manual review' in window.log.toPlainText().casefold()
        assert 'not deleted or added to history' in window.log.toPlainText().casefold()
        assert not orphan.exists()
        assert store.list_snapshots() == []
    finally:
        window.close()
