from pathlib import Path
from dataclasses import replace
import hashlib
from pq_installer.core.error_log import format_failure, record_failure
from pq_installer.core.excel_inspector import ExcelQueryInspector
from pq_installer.core.models import ExistingQueryMode, InstallResult, QueryStatus
from pq_installer.core.query_loader import read_m_code
from pq_installer.core.validation import validate_query_sources

from PySide6.QtCore import QObject, Signal, Slot
from pq_installer.core.backup import create_timestamped_backup
from pq_installer.core.excel_queries import ExcelQueryInstaller

class InstallWorker(QObject):
    log = Signal(str)
    completed = Signal(object)
    failed = Signal(str)
    finished = Signal()

    def __init__(self, workbook: Path, sources, mode, make_backup: bool, capabilities=None,history_store=None):
        super().__init__()
        self.workbook = workbook
        # Capture independent source objects and hashes before the job crosses threads.
        self.sources = [replace(source, source_sha256=hashlib.sha256(Path(source.path).read_bytes()).hexdigest()) for source in sources]
        self.mode = mode
        self.make_backup = make_backup
        self.capabilities = capabilities
        self.history_store=history_store
        self.workbook_sha256 = hashlib.sha256(Path(workbook).read_bytes()).hexdigest()

    def _all_enabled_queries_will_be_skipped(self):
        if self.mode != ExistingQueryMode.SKIP or not self.sources:
            return False
        validate_query_sources(self.sources)
        for source in self.sources:
            read_m_code(source)
        if hashlib.sha256(Path(self.workbook).read_bytes()).hexdigest() != self.workbook_sha256:
            raise RuntimeError('Workbook changed after import review; reload the workbook and review again.')
        inventory = ExcelQueryInspector().inspect(self.workbook)
        if hashlib.sha256(Path(self.workbook).read_bytes()).hexdigest() != self.workbook_sha256:
            raise RuntimeError('Workbook changed during skip preflight; reload the workbook and review again.')
        existing_names = {str(query.name).casefold() for query in inventory}
        return all(source.name.casefold() in existing_names for source in self.sources)

    @Slot()
    def run(self):
        try:
            self.log.emit('Worker thread started.')
            skip_only = self._all_enabled_queries_will_be_skipped()
            if skip_only:
                if hashlib.sha256(Path(self.workbook).read_bytes()).hexdigest() != self.workbook_sha256:
                    raise RuntimeError('Workbook changed after skip preflight; reload the workbook and review again.')
                results = tuple(InstallResult(source.name, QueryStatus.SKIPPED,
                    'Existing query kept unchanged; no workbook write was needed.') for source in self.sources)
                for result in results:
                    self.log.emit(f'{result.query_name}: {result.status.value} - {result.message}')
                self.log.emit('No workbook changes were committed; all selected queries were skipped.')
                self.completed.emit({'results': results, 'snapshot_id': None})
                return
            snapshot_id=None
            if self.history_store is not None and self.history_store.snapshots_enabled:
                snapshot=self.history_store.create_snapshot(self.workbook,event='workbook_snapshot_before_import')
                snapshot_id=snapshot['snapshot_id']
                self.log.emit(f'Local pre-import workbook version captured: {snapshot_id}')
            if self.make_backup:
                self.log.emit('Creating workbook backup...')
                backup = create_timestamped_backup(self.workbook)
                if hashlib.sha256(Path(backup).read_bytes()).hexdigest() != self.workbook_sha256:
                    raise RuntimeError('Workbook changed after import review; the verified backup does not match the reviewed workbook. Reload and review again.')
                self.log.emit(f'Backup created: {backup}')

            self.log.emit('Initializing Excel COM automation...')
            installer = ExcelQueryInstaller()
            def log_install_event(event):
                if isinstance(event, str):
                    self.log.emit(event)
                else:
                    self.log.emit(f'{event.query_name}: {event.status.value} - {event.message}')
            results = installer.install(
                self.workbook,
                self.sources,
                self.mode,
                logger=log_install_event,
                capabilities=self.capabilities,
                expected_workbook_sha256=self.workbook_sha256,
            )
            if results and all(getattr(result.status, 'value', result.status) == 'Skipped'
                               for result in results):
                self.log.emit('No workbook changes were committed; all selected queries were skipped.')
            else:
                self.log.emit('Excel automation finished. Workbook saved.')
            self.completed.emit({'results':tuple(results),'snapshot_id':snapshot_id})
        except Exception as exc:
            details = format_failure(exc)
            self.log.emit(f'ERROR: {details}')
            for result in getattr(exc,'results',()):
                status=getattr(result.status,'value',result.status)
                self.log.emit(f'{result.query_name}: {status} - {result.message}')
            diagnostic=record_failure('import',exc)
            if diagnostic:self.log.emit(f'Local diagnostic record saved: {diagnostic}')
            self.failed.emit(details)
        finally:
            self.finished.emit()
