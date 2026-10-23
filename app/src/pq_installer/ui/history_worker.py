"""Background worker for potentially large local workbook version restores."""
import hashlib
from pathlib import Path

from PySide6.QtCore import QObject,Signal,Slot
from pq_installer.core.error_log import format_failure, record_failure
from pq_installer.core.excel_inspector import ExcelQueryInspector
from pq_installer.core.history_comparison import compare_query_inventories


class SnapshotRestoreWorker(QObject):
    completed=Signal(object)
    failed=Signal(str)
    finished=Signal()

    def __init__(self,history_store,snapshot_id):
        super().__init__()
        self.history_store=history_store
        self.snapshot_id=snapshot_id

    @Slot()
    def run(self):
        try:
            snapshot=self.history_store.get_snapshot(self.snapshot_id)
            target=Path(snapshot['workbook_path'])
            if target.exists():
                digest=hashlib.sha256()
                with target.open('rb') as stream:
                    for chunk in iter(lambda:stream.read(1024*1024),b''):digest.update(chunk)
                outcome=self.history_store.restore_snapshot(self.snapshot_id,target,
                    overwrite=True,expected_current_sha256=digest.hexdigest())
            else:
                outcome=self.history_store.restore_snapshot(self.snapshot_id,target,overwrite=False)
            self.completed.emit(outcome)
        except Exception as exc:
            record_failure('history_restore',exc)
            self.failed.emit(format_failure(exc))
        finally:
            self.finished.emit()


class SnapshotCompareWorker(QObject):
    """Inspect two verified history snapshots on a COM worker thread, without refresh."""
    completed=Signal(object)
    failed=Signal(str)
    finished=Signal()

    def __init__(self,history_store,before_id,after_id,inspector_factory=None):
        super().__init__()
        self.history_store=history_store
        self.before_id=before_id
        self.after_id=after_id
        self.inspector_factory=inspector_factory or ExcelQueryInspector

    @Slot()
    def run(self):
        try:
            before=self.history_store.get_snapshot(self.before_id)
            after=self.history_store.get_snapshot(self.after_id)
            if (after['occurred_at'],after['snapshot_id']) < (before['occurred_at'],before['snapshot_id']):
                before,after=after,before
            if Path(before['snapshot_path']).suffix.casefold()!=Path(after['snapshot_path']).suffix.casefold():
                raise ValueError('Workbook snapshots must use the same Excel file type to compare query inventories.')
            before_queries=self.inspector_factory().inspect(before['snapshot_path'])
            after_queries=self.inspector_factory().inspect(after['snapshot_path'])
            self.completed.emit({'before':before,'after':after,
                'changes':compare_query_inventories(before_queries,after_queries)})
        except Exception as exc:
            record_failure('history_compare',exc)
            self.failed.emit(format_failure(exc))
        finally:
            self.finished.emit()
