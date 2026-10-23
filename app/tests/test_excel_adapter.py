from __future__ import annotations

from types import SimpleNamespace
import sys
import pytest

from pq_installer.core.excel_adapter import DispatchExExcelAdapter, resolve_excel_application_adapter
from pq_installer.core.excel_inspector import ExcelQueryInspector


class _Collection:
    Count = 0

    def Item(self, _index):
        raise IndexError(_index)


class _Workbook:
    ReadOnly = True
    Queries = _Collection()
    Connections = _Collection()
    Worksheets = _Collection()

    def __init__(self):
        self.closed = []

    def Close(self, save_changes):
        self.closed.append(save_changes)


class _Excel:
    AutomationSecurity = 1
    Visible = True
    DisplayAlerts = True

    def __init__(self, workbook):
        self.Workbooks = SimpleNamespace(Open=lambda *args, **kwargs: workbook)
        self.quit = False

    def Quit(self):
        self.quit = True


def test_inspector_uses_shared_fake_adapter_without_initializing_com(tmp_path, monkeypatch):
    workbook_path = tmp_path / "sample.xlsx"
    workbook_path.write_bytes(b"disposable test placeholder")
    workbook = _Workbook()
    excel = _Excel(workbook)
    calls = []

    class FakeAdapter:
        requires_com_initialization = False

        def create_application(self):
            calls.append("create")
            return excel

    monkeypatch.setitem(sys.modules, "pythoncom", SimpleNamespace(
        CoInitialize=lambda: (_ for _ in ()).throw(AssertionError("fake adapter must not initialize COM")),
        CoUninitialize=lambda: (_ for _ in ()).throw(AssertionError("fake adapter must not uninitialize COM")),
    ))

    inventory = ExcelQueryInspector(excel_adapter=FakeAdapter()).inspect(workbook_path)

    assert inventory == ()
    assert calls == ["create"]
    assert workbook.closed == [False]
    assert excel.quit


def test_inspector_routes_workbook_lifecycle_through_injected_operations(tmp_path, monkeypatch):
    workbook_path = tmp_path / 'lifecycle.xlsx'
    workbook_path.write_bytes(b'disposable test placeholder')
    workbook = _Workbook()
    excel = _Excel(workbook)
    lifecycle = []

    class FakeApplicationAdapter:
        requires_com_initialization = False

        def create_application(self):
            return excel

    class FakeWorkbookOperations:
        def add_workbook(self, application):
            return application.Workbooks.Add()

        def open_workbook(self, application, path, **options):
            lifecycle.append(('open', application, path, options))
            return application.Workbooks.Open(str(path), **options)

        def save_workbook(self, workbook):
            lifecycle.append(('save', workbook))
            workbook.Save()

        def save_workbook_as(self, workbook, path, file_format):
            workbook.SaveAs(str(path), file_format)

        def close_workbook(self, workbook, save_changes=False):
            lifecycle.append(('close', workbook, save_changes))
            workbook.Close(save_changes)

    monkeypatch.setitem(sys.modules, 'pythoncom', SimpleNamespace(
        CoInitialize=lambda: (_ for _ in ()).throw(AssertionError('fake adapter must not initialize COM')),
        CoUninitialize=lambda: (_ for _ in ()).throw(AssertionError('fake adapter must not uninitialize COM')),
    ))

    inventory = ExcelQueryInspector(
        excel_adapter=FakeApplicationAdapter(), workbook_operations=FakeWorkbookOperations()).inspect(workbook_path)

    assert inventory == ()
    assert lifecycle[0][0] == 'open'
    assert lifecycle[0][1] is excel
    assert lifecycle[0][2] == workbook_path
    assert lifecycle[0][3]['ReadOnly'] is True
    assert lifecycle[-1] == ('close', workbook, False)


def test_adapter_and_legacy_factory_cannot_both_be_configured():
    adapter = SimpleNamespace(requires_com_initialization=False, create_application=lambda: object())
    with pytest.raises(ValueError, match='either excel_adapter or excel_factory'):
        resolve_excel_application_adapter(adapter=adapter, factory=lambda: object())


def test_dispatch_adapter_uses_dispatchex_and_requires_com(monkeypatch):
    excel = object()
    calls = []
    client = SimpleNamespace(DispatchEx=lambda name: (calls.append(name) or excel))
    monkeypatch.setitem(sys.modules, 'win32com', SimpleNamespace(client=client))
    monkeypatch.setitem(sys.modules, 'win32com.client', client)

    adapter = DispatchExExcelAdapter()

    assert adapter.requires_com_initialization is True
    assert adapter.create_application() is excel
    assert calls == ['Excel.Application']

