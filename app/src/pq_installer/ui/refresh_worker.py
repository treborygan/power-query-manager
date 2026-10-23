from pathlib import Path
from pq_installer.core.error_log import format_failure, record_failure

from PySide6.QtCore import QObject, Signal, Slot

from pq_installer.core.excel_refresh import ExcelQueryRefresher, RefreshOperationError, RefreshResult


class RefreshWorker(QObject):
    log = Signal(str)
    completed = Signal(object)
    failed = Signal(str, object)
    finished = Signal()

    def __init__(self, workbook: Path, query_names, *, timeout_seconds: int,
                 asynchronous: bool, make_backup: bool,history_store=None):
        super().__init__()
        self.workbook = Path(workbook)
        self.query_names = tuple(query_names) if query_names is not None else None
        self.timeout_seconds = timeout_seconds
        self.asynchronous = asynchronous
        self.make_backup = make_backup
        self.history_store=history_store

    @Slot()
    def run(self):
        try:
            mode = 'Excel background' if self.asynchronous else 'synchronous'
            self.log.emit(f'Starting explicitly requested {mode} refresh; timeout per query: {self.timeout_seconds}s.')
            snapshot_id=None
            if self.history_store is not None and self.history_store.snapshots_enabled:
                snapshot=self.history_store.create_snapshot(self.workbook,event='workbook_snapshot_before_refresh')
                snapshot_id=snapshot['snapshot_id']
                self.log.emit(f'Local pre-refresh workbook version captured: {snapshot_id}')

            def report(item):
                if isinstance(item, RefreshResult):
                    self.log.emit(f'{item.query_name}: {item.status} - {item.message} ({item.duration_seconds:.1f}s)')
                else:
                    self.log.emit(str(item))

            results = ExcelQueryRefresher().refresh(
                self.workbook,
                self.query_names,
                timeout_seconds=self.timeout_seconds,
                asynchronous=self.asynchronous,
                make_backup=self.make_backup,
                logger=report,
            )
            self.completed.emit({'results':tuple(results),'snapshot_id':snapshot_id})
        except RefreshOperationError as exc:
            diagnostic=record_failure('refresh',exc)
            if diagnostic:self.log.emit(f'Local diagnostic record saved: {diagnostic}')
            self.failed.emit(format_failure(exc), exc.results)
        except Exception as exc:
            details = format_failure(exc)
            diagnostic=record_failure('refresh',exc)
            if diagnostic:self.log.emit(f'Local diagnostic record saved: {diagnostic}')
            self.failed.emit(details, ())
        finally:
            self.finished.emit()
