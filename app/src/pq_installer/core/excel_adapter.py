"""Injectable Excel application creation port shared by COM services and fake-COM tests.

The adapter creates an Excel Application in the calling thread. Production workers
must call it only after initializing their COM apartment; fake adapters can declare
that COM initialization is not required.
"""
from __future__ import annotations

from typing import Any, Callable, Protocol


class ExcelApplicationAdapter(Protocol):
    """Factory boundary for the Excel Application proxy used by core services."""

    requires_com_initialization: bool

    def create_application(self) -> Any: ...


class ExcelWorkbookOperations(Protocol):
    """Workbook-level COM operations shared by inspection, import and refresh."""

    def add_workbook(self, application: Any) -> Any: ...
    def open_workbook(self, application: Any, path, **options) -> Any: ...
    def save_workbook(self, workbook: Any) -> None: ...
    def save_workbook_as(self, workbook: Any, path, file_format: int) -> None: ...
    def close_workbook(self, workbook: Any, save_changes: bool = False) -> None: ...


class ComExcelWorkbookOperations:
    """Production implementation of the small workbook lifecycle port."""

    def add_workbook(self, application: Any) -> Any:
        return application.Workbooks.Add()

    def open_workbook(self, application: Any, path, **options) -> Any:
        return application.Workbooks.Open(str(path), **options)

    def save_workbook(self, workbook: Any) -> None:
        workbook.Save()

    def save_workbook_as(self, workbook: Any, path, file_format: int) -> None:
        workbook.SaveAs(str(path), file_format)

    def close_workbook(self, workbook: Any, save_changes: bool = False) -> None:
        workbook.Close(save_changes)


def resolve_excel_workbook_operations(operations=None) -> ExcelWorkbookOperations:
    """Resolve a workbook lifecycle port and reject incomplete fake adapters."""
    operations = operations if operations is not None else ComExcelWorkbookOperations()
    required = ('add_workbook', 'open_workbook', 'save_workbook', 'save_workbook_as', 'close_workbook')
    if any(not callable(getattr(operations, name, None)) for name in required):
        raise TypeError('Excel workbook operations must implement add/open/save/save-as/close methods.')
    return operations


class DispatchExExcelAdapter:
    """Create an isolated desktop Excel process using out-of-process COM."""

    requires_com_initialization = True

    def create_application(self) -> Any:
        import win32com.client
        return win32com.client.DispatchEx('Excel.Application')


class FactoryExcelAdapter:
    """Wrap a test or compatibility factory; COM ownership remains caller-defined."""

    requires_com_initialization = False

    def __init__(self, factory: Callable[[], Any]):
        if not callable(factory):
            raise TypeError('Excel application factory must be callable.')
        self._factory = factory

    def create_application(self) -> Any:
        return self._factory()


def resolve_excel_application_adapter(*, adapter=None, factory=None) -> ExcelApplicationAdapter:
    """Resolve the shared adapter while preserving the existing ``excel_factory`` API."""
    if adapter is not None and factory is not None:
        raise ValueError('Provide either excel_adapter or excel_factory, not both.')
    if adapter is None:
        adapter = FactoryExcelAdapter(factory) if factory is not None else DispatchExExcelAdapter()
    if not callable(getattr(adapter, 'create_application', None)):
        raise TypeError('Excel adapter must provide create_application().')
    if not isinstance(getattr(adapter, 'requires_com_initialization', None), bool):
        raise TypeError('Excel adapter must declare requires_com_initialization as a bool.')
    return adapter
