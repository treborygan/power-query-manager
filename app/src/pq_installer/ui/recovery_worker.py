"""Read-only Excel verification and copy-out of an abandoned staged workbook."""
from pathlib import Path
import hashlib

from PySide6.QtCore import QObject, Signal, Slot

from pq_installer.core.error_log import format_failure, record_failure
from pq_installer.core.excel_inspector import ExcelQueryInspector
from pq_installer.core.recovery import find_staged_workbooks, save_verified_recovery_copy
from pq_installer.core.windows_file_users import ensure_no_active_file_users


class StagedRecoveryWorker(QObject):
    completed=Signal(object)
    failed=Signal(str)
    finished=Signal()

    def __init__(self,workbook,staged,output):
        super().__init__()
        self.workbook=Path(workbook)
        self.staged=Path(staged)
        self.output=Path(output)

    @Slot()
    def run(self):
        try:
            resolved=self.staged.resolve()
            if resolved not in tuple(path.resolve() for path in find_staged_workbooks(self.workbook)):
                raise ValueError('The selected file is no longer a matching staged workbook.')
            digest=hashlib.sha256()
            with self.staged.open('rb') as stream:
                for chunk in iter(lambda:stream.read(1024*1024),b''):
                    digest.update(chunk)
            digest=digest.hexdigest()
            ensure_no_active_file_users(self.staged)
            inventory=ExcelQueryInspector().inspect(self.staged)
            copied=save_verified_recovery_copy(self.workbook,self.staged,self.output,expected_sha256=digest)
            self.completed.emit({'path':str(copied),'query_count':len(inventory),'sha256':digest})
        except Exception as exc:
            diagnostic=record_failure('staged_recovery',exc)
            suffix=f' Diagnostic log: {diagnostic}' if diagnostic else ''
            self.failed.emit(f'{format_failure(exc)}{suffix}')
        finally:
            self.finished.emit()
