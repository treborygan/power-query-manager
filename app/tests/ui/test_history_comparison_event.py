import os
os.environ.setdefault('QT_QPA_PLATFORM', 'offscreen')

from PySide6.QtCore import QThread, QObject, Signal, QItemSelectionModel, QEventLoop, QTimer
from PySide6.QtWidgets import QApplication, QDialog, QTableWidget, QPushButton
from pq_installer.core.history import HistoryStore
from pq_installer.ui import main_window as main_window_module
from pq_installer.ui.main_window import MainWindow

app = QApplication.instance() or QApplication([])


class ImmediateThread(QThread):
    def start(self, priority=QThread.InheritPriority):
        self.started.emit()
        self.finished.emit()


class FakeCompareWorker(QObject):
    completed = Signal(object)
    failed = Signal(str)
    finished = Signal()

    def __init__(self, store, before_id, after_id):
        super().__init__()
        self.before_id = before_id
        self.after_id = after_id

    def moveToThread(self, thread):
        self.test_thread = thread

    def run(self):
        self.completed.emit({
            'before': {'snapshot_id': self.before_id, 'occurred_at': 'before', 'workbook_name': 'report.xlsx'},
            'after': {'snapshot_id': self.after_id, 'occurred_at': 'after', 'workbook_name': 'report.xlsx'},
            'changes': (),
        })
        self.finished.emit()


def test_completed_ui_snapshot_comparison_records_both_snapshot_ids(tmp_path, monkeypatch):
    store = HistoryStore(tmp_path / 'history.sqlite3')
    before_file = tmp_path / 'before.xlsx'
    after_file = tmp_path / 'after.xlsx'
    before_file.write_bytes(b'before snapshot')
    after_file.write_bytes(b'after snapshot')
    before = store.create_snapshot(before_file)
    after = store.create_snapshot(after_file)

    monkeypatch.setattr(main_window_module, 'QThread', ImmediateThread)
    monkeypatch.setattr(main_window_module, 'SnapshotCompareWorker', FakeCompareWorker)
    opened = []

    def show_dialog(dialog):
        dialog.show()
        opened.append(dialog)
        return 0

    monkeypatch.setattr(QDialog, 'exec', show_dialog)
    window = MainWindow()
    window.history_store = store
    try:
        window._open_history()
        history_dialog = next(dialog for dialog in opened
                              if dialog.windowTitle() == 'Local operation history')
        table = history_dialog.findChild(QTableWidget)
        for row in range(table.rowCount()):
            table.selectionModel().select(
                table.model().index(row, 0),
                QItemSelectionModel.Select | QItemSelectionModel.Rows,
            )
        compare = next(button for button in history_dialog.findChildren(QPushButton)
                       if button.text() == 'Compare selected snapshots…')
        assert compare.isEnabled()
        compare.click()
        event = next(item for item in store.list() if item['event'] == 'workbook_snapshots_compared')
        assert event['details'] == {
            'action': 'snapshot_comparison_completed',
            'before_snapshot_id': before['snapshot_id'],
            'after_snapshot_id': after['snapshot_id'],
            'comparison': 'completed',
        }
    finally:
        for dialog in opened:
            dialog.close()
        window.close()


def test_snapshot_comparison_result_opens_dialog_on_ui_thread(tmp_path, monkeypatch):
    store = HistoryStore(tmp_path / 'history.sqlite3')
    before_file = tmp_path / 'before.xlsx'
    after_file = tmp_path / 'after.xlsx'
    before_file.write_bytes(b'before snapshot')
    after_file.write_bytes(b'after snapshot')
    before = store.create_snapshot(before_file)
    after = store.create_snapshot(after_file)
    opened = []
    loop = QEventLoop()

    class ThreadedCompareWorker(QObject):
        completed = Signal(object)
        failed = Signal(str)
        finished = Signal()

        def __init__(self, _store, before_id, after_id):
            super().__init__()
            self.before_id = before_id
            self.after_id = after_id

        def run(self):
            self.completed.emit({
                'before': {'snapshot_id': self.before_id, 'occurred_at': 'before', 'workbook_name': 'report.xlsx'},
                'after': {'snapshot_id': self.after_id, 'occurred_at': 'after', 'workbook_name': 'report.xlsx'},
                'changes': (),
            })
            self.finished.emit()

    class ComparisonDialogProbe:
        def __init__(self, *_args):
            pass

        def exec(self):
            opened.append(QThread.currentThread())
            loop.quit()
            return 0

    monkeypatch.setattr(main_window_module, 'SnapshotCompareWorker', ThreadedCompareWorker)
    monkeypatch.setattr(main_window_module, 'HistoryComparisonDialog', ComparisonDialogProbe)
    history_dialogs = []

    def show_history_dialog(dialog):
        dialog.show()
        if dialog.windowTitle() == 'Local operation history':
            history_dialogs.append(dialog)
        return 0

    monkeypatch.setattr(QDialog, 'exec', show_history_dialog)
    window = MainWindow()
    window.history_store = store
    try:
        window._open_history()
        history_dialog = history_dialogs[0]
        table = history_dialog.findChild(QTableWidget)
        for row in range(table.rowCount()):
            table.selectionModel().select(
                table.model().index(row, 0),
                QItemSelectionModel.Select | QItemSelectionModel.Rows,
            )
        compare = next(button for button in history_dialog.findChildren(QPushButton)
                       if button.text() == 'Compare selected snapshots…')
        assert compare.isEnabled()
        compare.click()
        comparison_thread = window._history_compare_thread
        QTimer.singleShot(5000, loop.quit)
        loop.exec()
        assert opened == [app.thread()], 'comparison dialog callback must run on the Qt application thread'
    finally:
        if 'comparison_thread' in locals() and comparison_thread is not None and comparison_thread.isRunning():
            cleanup_loop = QEventLoop()
            comparison_thread.finished.connect(cleanup_loop.quit)
            QTimer.singleShot(5000, cleanup_loop.quit)
            cleanup_loop.exec()
        for widget in app.topLevelWidgets():
            if widget.windowTitle() == 'Local operation history':
                widget.close()
        window.close()


def test_snapshot_restore_result_checks_history_dialog_on_ui_thread(tmp_path, monkeypatch):
    store = HistoryStore(tmp_path / 'history.sqlite3')
    workbook = tmp_path / 'report.xlsx'
    workbook.write_bytes(b'disposable workbook snapshot')
    snapshot = store.create_snapshot(workbook)
    loop = QEventLoop()
    callback_threads = []

    class ThreadedRestoreWorker(QObject):
        completed = Signal(object)
        failed = Signal(str)
        finished = Signal()

        def __init__(self, _store, snapshot_id):
            super().__init__()
            self.snapshot_id = snapshot_id

        def run(self):
            self.completed.emit({
                'restored_path': str(workbook),
                'backup_path': None,
                'history_warning': None,
            })
            self.finished.emit()

    original_is_visible = QDialog.isVisible

    def track_visibility(dialog):
        if dialog.windowTitle() == 'Local operation history':
            callback_threads.append(QThread.currentThread())
            loop.quit()
            return False
        return original_is_visible(dialog)

    monkeypatch.setattr(main_window_module, 'SnapshotRestoreWorker', ThreadedRestoreWorker)
    monkeypatch.setattr(main_window_module.QMessageBox, 'question', lambda *_args, **_kwargs: main_window_module.QMessageBox.Yes)
    history_dialogs = []
    monkeypatch.setattr(QDialog, 'isVisible', track_visibility)
    monkeypatch.setattr(QDialog, 'exec', lambda dialog: (history_dialogs.append(dialog)
                                                         if dialog.windowTitle() == 'Local operation history' else None) or 0)
    window = MainWindow()
    window.history_store = store
    try:
        window._open_history()
        history_dialog = history_dialogs[0]
        table = history_dialog.findChild(QTableWidget)
        table.selectRow(0)
        restore = next(button for button in history_dialog.findChildren(QPushButton)
                       if button.text() == 'Restore selected version…')
        assert restore.isEnabled()
        restore.click()
        restore_thread = window._history_restore_thread
        QTimer.singleShot(5000, loop.quit)
        loop.exec()
        assert callback_threads and callback_threads[0] == app.thread(), (
            'restore completion must check dialog visibility on the Qt application thread')
    finally:
        if 'restore_thread' in locals() and restore_thread is not None and restore_thread.isRunning():
            cleanup_loop = QEventLoop()
            restore_thread.finished.connect(cleanup_loop.quit)
            QTimer.singleShot(5000, cleanup_loop.quit)
            cleanup_loop.exec()
        window.close()


def test_background_thread_completion_is_dispatched_to_main_window_thread():
    window = MainWindow()
    observed = []
    window._set_busy = lambda busy, activity='Importing…': observed.append(QThread.currentThread())
    class FinishingThread(QThread):
        def run(self):
            pass

    thread = FinishingThread()
    thread.finished.connect(window._on_background_thread_finished)
    thread.start()
    assert thread.wait(5000)
    app.processEvents()
    assert observed == [app.thread()]
    window.close()
