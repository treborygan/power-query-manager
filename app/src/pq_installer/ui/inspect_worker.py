from pathlib import Path
from PySide6.QtCore import QObject, Signal, Slot
from pq_installer.core.error_log import format_failure, record_failure
from pq_installer.core.excel_inspector import ExcelQueryInspector
from pq_installer.core.excel_capabilities import probe_excel_capabilities


class InspectWorker(QObject):
    completed=Signal(object)
    failed=Signal(str)
    finished=Signal()
    def __init__(self,workbook:Path):super().__init__();self.workbook=Path(workbook)
    @Slot()
    def run(self):
        try:self.completed.emit(ExcelQueryInspector().inspect(self.workbook))
        except Exception as exc:
            diagnostic=record_failure('inspect',exc)
            suffix=f'\nDiagnostic log: {diagnostic}' if diagnostic else ''
            self.failed.emit(f'{format_failure(exc)}{suffix}')
        finally:self.finished.emit()


class CapabilityProbeWorker(QObject):
    completed=Signal(object)
    failed=Signal(str)
    finished=Signal()
    @Slot()
    def run(self):
        try:self.completed.emit(probe_excel_capabilities())
        except Exception as exc:
            diagnostic=record_failure('capability_probe',exc)
            suffix=f'\nDiagnostic log: {diagnostic}' if diagnostic else ''
            self.failed.emit(f'{format_failure(exc)}{suffix}')
        finally:self.finished.emit()
