"""Core models, services, and injectable workbook application adapters."""

from .excel_adapter import (
    DispatchExExcelAdapter,
    ComExcelWorkbookOperations,
    ExcelApplicationAdapter,
    ExcelWorkbookOperations,
    FactoryExcelAdapter,
    resolve_excel_application_adapter,
    resolve_excel_workbook_operations,
)
from .excel_transaction import LocalWorkbookTransactionAdapter, WorkbookTransactionAdapter

__all__ = [
    'DispatchExExcelAdapter',
    'ComExcelWorkbookOperations',
    'ExcelApplicationAdapter',
    'ExcelWorkbookOperations',
    'FactoryExcelAdapter',
    'resolve_excel_application_adapter',
    'resolve_excel_workbook_operations',
    'LocalWorkbookTransactionAdapter',
    'WorkbookTransactionAdapter',
]
