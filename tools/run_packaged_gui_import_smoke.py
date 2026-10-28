"""Drive a packaged definition-only import through the visible Qt UI and Excel."""
from __future__ import annotations

import argparse
import ctypes
from ctypes import wintypes
import hashlib
import json
import os
import re
import stat
import subprocess
import sys
import tempfile
import time
import zipfile
from pathlib import Path

import pythoncom
import win32api
import win32con
import win32gui
import win32process
import win32com.client


def _sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open('rb') as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b''):
            digest.update(chunk)
    return digest.hexdigest()


def _excel_pids() -> set[int]:
    """Enumerate Excel process IDs without opening or controlling any instance."""
    psapi = ctypes.WinDLL('psapi', use_last_error=True)
    kernel32 = ctypes.WinDLL('kernel32', use_last_error=True)
    psapi.EnumProcesses.argtypes = [ctypes.POINTER(wintypes.DWORD), wintypes.DWORD,
                                    ctypes.POINTER(wintypes.DWORD)]
    psapi.EnumProcesses.restype = wintypes.BOOL
    kernel32.OpenProcess.argtypes = [wintypes.DWORD, wintypes.BOOL, wintypes.DWORD]
    kernel32.OpenProcess.restype = wintypes.HANDLE
    kernel32.QueryFullProcessImageNameW.argtypes = [wintypes.HANDLE, wintypes.DWORD,
                                                    wintypes.LPWSTR, ctypes.POINTER(wintypes.DWORD)]
    kernel32.QueryFullProcessImageNameW.restype = wintypes.BOOL
    kernel32.CloseHandle.argtypes = [wintypes.HANDLE]
    kernel32.CloseHandle.restype = wintypes.BOOL
    pid_array = (wintypes.DWORD * 4096)()
    bytes_needed = wintypes.DWORD()
    if not psapi.EnumProcesses(pid_array, ctypes.sizeof(pid_array), ctypes.byref(bytes_needed)):
        raise ctypes.WinError(ctypes.get_last_error())
    found: set[int] = set()
    for pid in pid_array[:bytes_needed.value // ctypes.sizeof(wintypes.DWORD)]:
        handle = kernel32.OpenProcess(0x1000, False, pid)  # PROCESS_QUERY_LIMITED_INFORMATION
        if not handle:
            continue
        try:
            size = wintypes.DWORD(32768)
            image = ctypes.create_unicode_buffer(size.value)
            if kernel32.QueryFullProcessImageNameW(handle, 0, image, ctypes.byref(size)):
                if Path(image.value).name.casefold() == 'excel.exe':
                    found.add(int(pid))
        finally:
            kernel32.CloseHandle(handle)
    return found


def _wait_child(window, title: str, timeout: float = 15):
    deadline = time.monotonic() + timeout
    while time.monotonic() < deadline:
        try:
            child = window.child_window(title=title, control_type='Window')
            child.element_info.rectangle
            return child
        except Exception:
            time.sleep(0.2)
    try:
        visible = [(child.element_info.control_type, child.window_text())
                   for child in window.descendants(control_type='Window')]
        statuses = [child.window_text() for child in window.descendants(control_type='StatusBar')]
    except Exception as exc:
        visible, statuses = [], [f'{type(exc).__name__}: {exc}']
    raise TimeoutError(f'Packaged UI dialog {title!r} did not appear; visible={visible!r}; status={statuses!r}.')


def _accept_file_dialog(dialog) -> None:
    """Accept the Windows picker across UIA's Button/SplitButton role variants."""
    for role in ('SplitButton', 'Button'):
        try:
            button = dialog.child_window(auto_id='1', control_type=role)
            if button.exists(timeout=0.4):
                button.invoke()
                return
        except Exception:
            continue
    # On some Windows builds the default Open action is exposed only as the
    # dialog's keyboard default, even though the file-name edit remains visible.
    try:
        dialog.set_focus()
        dialog.type_keys('{ENTER}', set_foreground=True)
        return
    except Exception as exc:
        observed = [(item.element_info.name, item.element_info.control_type,
                     item.element_info.automation_id)
                    for item in dialog.descendants()]
        raise RuntimeError(f'Windows file picker has no operable Open action: {observed!r}.') from exc


def _assert_main_window_accessible(window) -> None:
    observed = {(item.element_info.name, item.element_info.control_type)
                for item in window.descendants()}
    required = {
        ('Target workbook path', 'Edit'),
        ('Query source settings table', 'Table'),
        ('Select all query rows', 'Button'),
        ('Filter query source rows', 'Edit'),
        ('Existing query behavior', 'ComboBox'),
        ('Create a timestamped workbook backup', 'CheckBox'),
        ('Review and import selected queries', 'Button'),
        ('Operation log output', 'Edit'),
    }
    missing = sorted(required - observed)
    if missing:
        raise AssertionError(f'Packaged main window is missing UI Automation names/roles: {missing!r}.')


def _assert_named_controls(dialog, required) -> None:
    observed = {(item.element_info.name, item.element_info.control_type)
                for item in dialog.descendants()}
    missing = sorted(set(required) - observed)
    if missing:
        raise AssertionError(f'Packaged modal is missing UI Automation names/roles: {missing!r}; observed={sorted(observed)!r}.')


def _exercise_packaged_modal_accessibility(window) -> None:
    source_rows = [item for item in window.descendants(control_type='DataItem')
                   if item.window_text() == 'GuiSmoke']
    if not source_rows:
        raise AssertionError('Packaged source row GuiSmoke is not available for inspector selection.')
    source_rows[0].select()
    window.child_window(title='Inspect / compare selected', control_type='Button').invoke()
    inspector = _wait_child(window, 'Query inspector — GuiSmoke')
    _assert_named_controls(inspector, [
        ('Search M expressions and diff', 'Edit'),
        ('Editable source M expression', 'Edit'),
        ('Read-only workbook M expression', 'Edit'),
        ('Line-by-line workbook and source M diff', 'Edit'),
        ('Edit supported declared M parameter', 'Button'),
        ('Save edited M as a separate source copy', 'Button'),
        ('Close query inspector', 'Button'),
    ])
    inspector.child_window(title='Close query inspector', control_type='Button').invoke()

    window.child_window(title='Dependency graph', control_type='Button').invoke()
    dependencies = _wait_child(window, 'Query dependencies')
    _assert_named_controls(dependencies, [
        ('Visual query dependency diagram', 'Group'),
        ('Text summary of query dependencies', 'Edit'),
        ('Query dependency dialog actions', 'Group'),
        ('Close query dependency graph', 'Button'),
    ])
    dependencies.child_window(title='Close query dependency graph', control_type='Button').invoke()

    window.child_window(title='Edit source environment mappings', control_type='Button').invoke()
    mappings = _wait_child(window, 'Edit source environment mappings')
    _assert_named_controls(mappings, [
        ('Profile mapping categories', 'Group'),
        ('Source paths', 'TabItem'),
        ('M source locations', 'TabItem'),
        ('Profile source path mappings', 'Table'),
        ('Add source path mapping', 'Button'),
        ('Select all source path mappings', 'Button'),
        ('Remove selected source path mappings', 'Button'),
        ('Cancel mapping edits', 'Button'),
        ('Save profile mappings', 'Button'),
    ])
    mappings.child_window(title='M source locations', control_type='TabItem').select()
    _assert_named_controls(mappings, [
        ('M source location mappings', 'Table'),
        ('Add M source location mapping', 'Button'),
        ('Select all M source location mappings', 'Button'),
        ('Remove selected M source location mappings', 'Button'),
    ])
    mappings.child_window(title='Cancel mapping edits', control_type='Button').invoke()
    print('PASS: packaged query inspector, dependency graph and profile mapping dialogs expose named controls through Windows UI Automation; mapping changes cancelled.')


def _exercise_packaged_parameter_editor(window, root: Path, source: Path,
                                        original_source_hash: str) -> tuple[Path, bytes]:
    """Edit a declared parameter via the packaged GUI and save only a new source copy."""
    source_rows = [item for item in window.descendants(control_type='DataItem')
                   if item.window_text() == 'GuiSmoke']
    if not source_rows:
        raise AssertionError('Packaged parameter source row GuiSmoke is not available.')
    source_rows[0].select()
    window.child_window(title='Inspect / compare selected', control_type='Button').invoke()
    inspector = _wait_child(window, 'Query inspector — GuiSmoke')
    parameter_button = inspector.child_window(title='Edit supported declared M parameter', control_type='Button')
    if not parameter_button.is_enabled():
        raise AssertionError('Packaged GUI did not recognize the supported declared text parameter.')
    parameter_button.invoke()
    editor = _wait_child(window, 'Edit declared parameter')
    edits = editor.descendants(control_type='Edit')
    if len(edits) != 1:
        observed = [(item.element_info.name, item.element_info.automation_id)
                    for item in edits]
        raise AssertionError(f'Parameter input dialog has unexpected edit controls: {observed!r}.')
    edits[0].set_edit_text('US')
    editor.type_keys('{ENTER}', set_foreground=True)

    # Editing adds a trailing asterisk to the Qt dialog title; reacquire the
    # accessible window rather than using the now-stale exact-title wrapper.
    inspector = window.child_window(title_re=r'Query inspector.*GuiSmoke.*', control_type='Window')
    inspector.child_window(title='Save edited M as a separate source copy', control_type='Button').invoke()
    picker = _wait_child(inspector, 'Save edited source copy')
    edited_source = root / 'GuiSmoke_edited.m'
    try:
        filename = picker.child_window(auto_id='1148', control_type='ComboBox').child_window(control_type='Edit')
        filename.set_edit_text(str(edited_source))
    except Exception as open_dialog_shape:
        # Open and Save dialogs expose different automation IDs on some Windows
        # builds; accept the unique filename Edit when that variant is present.
        filename_edits = picker.descendants(control_type='Edit')
        filename_combo = picker.child_window(title='File name:', control_type='ComboBox')
        if filename_combo.exists(timeout=0.2):
            filename_combo.child_window(auto_id='1001', control_type='Edit').set_edit_text(str(edited_source))
        elif len(filename_edits) == 1:
            filename_edits[0].set_edit_text(str(edited_source))
        else:
            observed = [(item.element_info.name, item.element_info.automation_id,
                         item.element_info.control_type)
                        for item in picker.descendants()]
            raise RuntimeError(f'Save picker has no unambiguous filename edit; controls={observed!r}.') from open_dialog_shape
    _accept_file_dialog(picker)
    saved = _wait_child(inspector, 'Source copy saved')
    saved.child_window(title='OK', control_type='Button').invoke()
    inspector.child_window(title='Close query inspector', control_type='Button').invoke()
    expected = b'"US" meta [IsParameterQuery=true, Type="Text"]\r\n'
    if not edited_source.is_file() or edited_source.read_bytes() != expected:
        actual = edited_source.read_bytes() if edited_source.is_file() else None
        raise AssertionError(f'Parameter source copy bytes differ from the approved edit: {actual!r}.')
    if _sha256(source) != original_source_hash:
        raise AssertionError('Packaged parameter editor changed the original source file.')
    print('PASS: packaged parameter editor saved the edited literal to a separate source copy and preserved the exact original source bytes.')
    return edited_source, expected


def _create_workbook(path: Path, *, macro_enabled: bool = False,
                     with_vba: bool = False, existing_sheet_destination: bool = False) -> tuple[str, str, int]:
    excel = win32com.client.DispatchEx('Excel.Application')
    pid = int(win32process.GetWindowThreadProcessId(int(excel.Hwnd))[1])
    version, build = str(excel.Version), str(excel.Build)
    book = None
    try:
        excel.Visible = False
        excel.DisplayAlerts = False
        book = excel.Workbooks.Add()
        if existing_sheet_destination:
            output = book.Worksheets(1)
            output.Name = 'Output'
            marker = book.Worksheets.Add(After=book.Worksheets.Item(book.Worksheets.Count))
            marker.Name = 'PQM_Marker'
            marker.Range('A1').Value = 'Power Query Manager GUI smoke marker'
        else:
            book.Worksheets(1).Range('A1').Value = 'Power Query Manager GUI smoke marker'
        book.SaveAs(str(path), 52 if macro_enabled else 51)
        if with_vba:
            module = book.VBProject.VBComponents.Item('ThisWorkbook').CodeModule
            module.AddFromString(
                'Private Sub Workbook_Open()\r\n'
                '    Worksheets(1).Range("A1").Value = "MACRO_RAN"\r\n'
                'End Sub')
            book.Save()
        book.Close(False)
        book = None
    finally:
        if book is not None:
            try:
                book.Close(False)
            except Exception:
                pass
        excel.Quit()
        excel = None
    return version, build, pid


def _change_workbook_cell(path: Path, value: str) -> int:
    """Change one cell in a smoke-owned disposable workbook and close owned Excel."""
    excel = win32com.client.DispatchEx('Excel.Application')
    pid = int(win32process.GetWindowThreadProcessId(int(excel.Hwnd))[1])
    book = None
    try:
        excel.Visible = False
        excel.DisplayAlerts = False
        book = excel.Workbooks.Open(str(path), UpdateLinks=0, ReadOnly=False,
                                    IgnoreReadOnlyRecommended=True, AddToMru=False)
        book.Worksheets(1).Range('A1').Value = value
        book.Save()
        book.Close(False)
        book = None
    finally:
        if book is not None:
            try:
                book.Close(False)
            except Exception:
                pass
        excel.Quit()
        excel = None
    deadline = time.monotonic() + 15
    while pid in _excel_pids() and time.monotonic() < deadline:
        time.sleep(0.2)
    if pid in _excel_pids():
        raise RuntimeError(f'Smoke-owned workbook mutation Excel PID {pid} did not exit.')
    return pid


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--exe', type=Path,
                        default=Path(__file__).resolve().parents[1] / 'dist' / 'PowerQueryManager' / 'PowerQueryManager.exe')
    parser.add_argument('--xlsm', action='store_true',
                        help='use a disposable macro-enabled workbook container without a VBA project')
    parser.add_argument('--bom-crlf-source', action='store_true',
                        help='import UTF-8-BOM and CRLF M source and verify exact Excel formula readback and unchanged source bytes')
    parser.add_argument('--vba', action='store_true',
                        help='create a harmless Workbook_Open macro in a disposable .xlsm; requires existing Excel VBA project access')
    parser.add_argument('--skip-existing', action='store_true',
                        help='seed an existing query, choose Skip, and verify packaged GUI performs no workbook write or backup')
    parser.add_argument('--replace-existing', action='store_true',
                        help='seed an existing query and verify packaged GUI replaces its formula with a verified backup')
    parser.add_argument('--compare-terminal-newline', action='store_true',
                        help='verify packaged GUI reports UNCHANGED when Excel removed only the source terminal CRLF')
    parser.add_argument('--replace-preserve-load', action='store_true',
                        help='seed a connection-only query and verify packaged GUI replacement preserves that load state')
    parser.add_argument('--replace-preserve-model-load', action='store_true',
                        help='seed a Data Model query and verify packaged GUI definition replacement preserves the model association without refreshing it')
    parser.add_argument('--replace-preserve-worksheet-load', action='store_true',
                        help='seed a worksheet-loaded query and verify packaged GUI definition replacement preserves its table and rows without refreshing it')
    parser.add_argument('--mixed-import', action='store_true',
                        help='import numbered .m/.pq sources through a versioned profile and verify exact formula/order readback')
    parser.add_argument('--new-sheet', action='store_true',
                        help='probe Excel capabilities, import a static query to a new worksheet through the packaged GUI, and verify values')
    parser.add_argument('--connection-only', action='store_true',
                        help='import a static query as a workbook connection through the packaged GUI and verify no worksheet or model table is created')
    parser.add_argument('--data-model', action='store_true',
                        help='probe Excel capabilities, import a static query to the Data Model through the packaged GUI, and verify model membership')
    parser.add_argument('--refresh-data-model', action='store_true',
                        help='after --data-model import, explicitly refresh that static query with the packaged CLI and verify backup and model readback')
    parser.add_argument('--existing-sheet', action='store_true',
                        help='probe Excel capabilities, import a static query to Output!C4 through the packaged GUI, and verify placement and values')
    parser.add_argument('--package-dialog', action='store_true',
                        help='compare two generated query packages and verify the packaged preview dialog through Windows UI Automation')
    parser.add_argument('--history-dialog', action='store_true',
                        help='seed two local snapshots and verify the packaged history comparison dialog through Windows UI Automation')
    parser.add_argument('--history-restore', action='store_true',
                        help='restore a hash-verified local snapshot through the packaged GUI and verify its backup on a disposable workbook')
    parser.add_argument('--parameter-edit', action='store_true',
                        help='edit a declared text parameter in the packaged inspector, save a separate source copy, import it, and verify Excel formula/backup')
    args = parser.parse_args()
    if args.compare_terminal_newline:
        args.replace_existing = True
    if ctypes.windll.shell32.IsUserAnAdmin():
        parser.error('Run this acceptance smoke from a non-elevated Windows session; administrator runs do not verify standard-user behavior.')
    if args.vba:
        args.xlsm = True
    if args.history_dialog and args.history_restore:
        parser.error('Choose either --history-dialog or --history-restore for a history smoke.')
    if args.bom_crlf_source and any((args.skip_existing,
                                      args.replace_preserve_load, args.mixed_import,
                                      args.new_sheet, args.connection_only,
                                      args.data_model, args.existing_sheet,
                                      args.replace_preserve_model_load,
                                      args.replace_preserve_worksheet_load)):
                    parser.error('--bom-crlf-source is a create or existing-query replacement smoke; do not combine it with other modes.')
    if sum((args.new_sheet, args.existing_sheet, args.connection_only, args.data_model)) > 1:
        parser.error('Choose only one load-mode smoke option.')
    if args.refresh_data_model and not args.data_model:
        parser.error('--refresh-data-model requires --data-model.')
    sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'app' / 'src'))
    executable = args.exe.resolve()
    if not executable.is_file():
        parser.error(f'Packaged executable not found: {executable}')

    try:
        from pywinauto import Application  # noqa: F811
    except ImportError as exc:
        raise SystemExit('Install tools/requirements_ui_smoke.txt in the QA environment before running this smoke.') from exc

    pythoncom.CoInitialize()
    manager = None
    manager_hwnd = None
    excel_before = _excel_pids()
    print('PASS: acceptance harness is running with a non-elevated Windows token; packaged manager will launch without elevation.')
    try:
        with tempfile.TemporaryDirectory(prefix='pqm-packaged-gui-smoke-') as temporary:
            root = Path(temporary)
            workbook = root / ('gui-import.xlsm' if args.xlsm else 'gui-import.xlsx')
            source = root / 'GuiSmoke.m'
            existing_formula = 'let Source = #table({"Value"}, {{5}}) in Source'
            formula = (existing_formula + '\r\n' if args.compare_terminal_newline else
                       '"EU" meta [IsParameterQuery=true, Type="Text"]\r\n'
                       if args.parameter_edit else
                       'let Source = #table({"Value"}, {{17}, {29}}) in Source')
            if args.bom_crlf_source:
                formula = '// UTF-8 BOM and CRLF source acceptance\r\n' + formula + '\r\n'
            source_bytes = (b'\xef\xbb\xbf' + formula.encode('utf-8')
                            if args.bom_crlf_source else formula.encode('utf-8'))
            source.write_bytes(source_bytes)
            source_hash = _sha256(source)
            package_pair = None
            if args.package_dialog:
                from pq_installer.core.manager import QueryInfo, export_query_package
                package_pair = (root / 'current.pqm.zip', root / 'updated.pqm.zip')
                export_query_package([QueryInfo('GuiSmoke', 'let Source = 1 in Source')], package_pair[0])
                export_query_package([
                    QueryInfo('GuiSmoke', 'let Source = 2 in Source'),
                    QueryInfo('NewQuery', 'let Value = 3 in Value'),
                ], package_pair[1])
            history_root = root / 'localappdata'
            restore_snapshot_hash = None
            restore_expected_backup_hash = None
            if args.history_dialog or args.history_restore:
                import shutil
            office_version, office_build, seed_pid = _create_workbook(
                workbook, macro_enabled=args.xlsm, with_vba=args.vba,
                existing_sheet_destination=(args.existing_sheet or args.replace_preserve_worksheet_load))
            seed_deadline = time.monotonic() + 15
            while seed_pid in _excel_pids() and time.monotonic() < seed_deadline:
                time.sleep(0.2)
            if seed_pid in _excel_pids():
                raise RuntimeError(f'Smoke-owned seed Excel PID {seed_pid} did not exit.')
            print(f'INFO: seed Excel PID {seed_pid} exited; process set={sorted(_excel_pids())}.')
            if args.history_dialog:
                from pq_installer.core.history import HistoryStore
                history_workbook = root / 'history-second.xlsx'
                shutil.copyfile(workbook, history_workbook)
                history_store = HistoryStore(history_root / 'PowerQueryManager' / 'history.sqlite3')
                history_store.create_snapshot(workbook, event='packaged_ui_smoke_before')
                history_store.create_snapshot(history_workbook, event='packaged_ui_smoke_after')
            elif args.history_restore:
                from pq_installer.core.history import HistoryStore
                history_store = HistoryStore(history_root / 'PowerQueryManager' / 'history.sqlite3')
                snapshot = history_store.create_snapshot(workbook, event='packaged_ui_smoke_restore_before')
                restore_snapshot_hash = snapshot['sha256']
                _change_workbook_cell(workbook, 'Changed after saved history snapshot')
                restore_expected_backup_hash = _sha256(workbook)
            mixed_rows = []
            if args.mixed_import:
                from pq_installer.core.models import LoadMode, QuerySource
                mixed_specs = [
                    ('01_First.m', 'First', 'let Source = #table({"Value"}, {{1}}) in Source'),
                    ('02_Second.pq', 'Second', 'let Source = #table({"Value"}, {{2}}) in Source'),
                    ('03_Third.m', 'Third', 'let Source = #table({"Value"}, {{3}}) in Source'),
                ]
                for order, (filename, name, code) in enumerate(mixed_specs, 1):
                    path = root / filename
                    path.write_text(code, encoding='utf-8')
                    mixed_rows.append((QuerySource(path, name, order=order), f'mixed-{name.lower()}'))

            if args.replace_preserve_load:
                from pq_installer.core.excel_capabilities import probe_excel_capabilities
                from pq_installer.core.excel_queries import ExcelQueryInstaller
                from pq_installer.core.models import ExistingQueryMode, LoadMode, QuerySource
                source.write_text(existing_formula, encoding='utf-8')
                capabilities = probe_excel_capabilities(staging_parent=root)
                if not capabilities.supports('connection_only'):
                    raise RuntimeError(f'Connection-only seed is unsupported: {capabilities.reasons}')
                seed = QuerySource(source, 'GuiSmoke', load_mode=LoadMode.CONNECTION_ONLY)
                ExcelQueryInstaller().install(workbook, [seed], ExistingQueryMode.REPLACE,
                                              capabilities=capabilities)
                source.write_text(formula, encoding='utf-8')
            elif args.replace_preserve_model_load:
                from pq_installer.core.excel_capabilities import probe_excel_capabilities
                from pq_installer.core.excel_queries import ExcelQueryInstaller
                from pq_installer.core.models import ExistingQueryMode, LoadMode, QuerySource
                source.write_text(existing_formula, encoding='utf-8')
                capabilities = probe_excel_capabilities(staging_parent=root)
                if not capabilities.supports('data_model'):
                    raise RuntimeError(f'Data Model seed is unsupported: {capabilities.reasons}')
                seed = QuerySource(source, 'GuiSmoke', load_mode=LoadMode.DATA_MODEL)
                ExcelQueryInstaller().install(workbook, [seed], ExistingQueryMode.REPLACE,
                                              capabilities=capabilities)
                source.write_text(formula, encoding='utf-8')
            elif args.replace_preserve_worksheet_load:
                from pq_installer.core.excel_capabilities import probe_excel_capabilities
                from pq_installer.core.excel_queries import ExcelQueryInstaller
                from pq_installer.core.models import ExistingQueryMode, LoadMode, QuerySource
                source.write_text(existing_formula, encoding='utf-8')
                capabilities = probe_excel_capabilities(staging_parent=root)
                if not capabilities.supports('existing_sheet'):
                    raise RuntimeError(f'Existing worksheet seed is unsupported: {capabilities.reasons}')
                seed = QuerySource(source, 'GuiSmoke', load_mode=LoadMode.EXISTING_SHEET,
                                   sheet_name='Output', cell='C4', table_name='tblGuiSmoke')
                ExcelQueryInstaller().install(workbook, [seed], ExistingQueryMode.REPLACE,
                                              capabilities=capabilities)
                source.write_text(formula, encoding='utf-8')
            elif args.skip_existing or args.replace_existing:
                seed_excel = win32com.client.DispatchEx('Excel.Application')
                seed_pid = int(win32process.GetWindowThreadProcessId(int(seed_excel.Hwnd))[1])
                seed_book = None
                try:
                    seed_excel.Visible = False
                    seed_excel.DisplayAlerts = False
                    seed_excel.AutomationSecurity = 3
                    seed_book = seed_excel.Workbooks.Open(str(workbook), UpdateLinks=0, ReadOnly=False)
                    seed_book.Queries.Add('GuiSmoke', existing_formula)
                    seed_book.Save()
                    seed_book.Close(False)
                    seed_book = None
                finally:
                    if seed_book is not None:
                        seed_book.Close(False)
                    seed_excel.Quit()
                    seed_excel = None
                seed_deadline = time.monotonic() + 15
                while seed_pid in _excel_pids() and time.monotonic() < seed_deadline:
                    time.sleep(0.2)
                if seed_pid in _excel_pids():
                    raise RuntimeError(f'Smoke-owned query seed Excel PID {seed_pid} did not exit.')
                if args.skip_existing:
                    os.chmod(workbook, stat.S_IREAD)
            original_hash = _sha256(workbook)
            original_mtime = workbook.stat().st_mtime_ns

            if args.skip_existing or args.replace_existing or args.replace_preserve_load or args.replace_preserve_model_load or args.replace_preserve_worksheet_load or args.mixed_import or args.new_sheet or args.existing_sheet or args.connection_only or args.data_model:
                sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'app' / 'src'))
                from pq_installer.core.manager import build_import_profile, save_profile
                from pq_installer.core.models import LoadMode, QuerySource
                profile_path = root / 'skip-profile.json'
                if args.mixed_import:
                    # Store rows in reverse order so acceptance proves numeric prefixes
                    # determine the execution sequence, not profile serialization order.
                    profile_sources = [row[0] for row in reversed(mixed_rows)]
                    profile_ids = [row[1] for row in reversed(mixed_rows)]
                    profile_name = 'Packaged mixed import smoke'
                    existing_policy = 'replace'
                else:
                    if args.new_sheet:
                        profile_sources = [QuerySource(source, 'GuiSmoke', load_mode=LoadMode.NEW_SHEET,
                                                        sheet_name='GuiOutput', table_name='tblGuiSmoke')]
                    elif args.existing_sheet:
                        profile_sources = [QuerySource(source, 'GuiSmoke', load_mode=LoadMode.EXISTING_SHEET,
                                                        sheet_name='Output', cell='C4', table_name='tblGuiSmoke')]
                    elif args.connection_only:
                        profile_sources = [QuerySource(source, 'GuiSmoke', load_mode=LoadMode.CONNECTION_ONLY)]
                    elif args.data_model:
                        profile_sources = [QuerySource(source, 'GuiSmoke', load_mode=LoadMode.DATA_MODEL)]
                    else:
                        source_load_mode = (LoadMode.PRESERVE_EXISTING if (args.replace_preserve_load or args.replace_preserve_model_load or args.replace_preserve_worksheet_load)
                                            else LoadMode.DEFINITION_ONLY)
                        profile_sources = [QuerySource(source, 'GuiSmoke', load_mode=source_load_mode)]
                    profile_ids = ['gui-smoke']
                    profile_name = ('Packaged new-sheet smoke' if args.new_sheet else
                                    'Packaged existing-sheet smoke' if args.existing_sheet else
                                    'Packaged connection-only smoke' if args.connection_only else
                                    'Packaged Data Model smoke' if args.data_model else
                                    'Packaged skip smoke' if args.skip_existing else
                                    'Packaged preserve-worksheet-load replacement smoke' if args.replace_preserve_worksheet_load else
                                    'Packaged preserve-Data-Model-load replacement smoke' if args.replace_preserve_model_load else
                                    'Packaged preserve-load replacement smoke' if args.replace_preserve_load else
                                    'Packaged replace smoke')
                    existing_policy = 'skip' if args.skip_existing else 'replace'
                profile = build_import_profile(profile_name, workbook,
                    profile_sources, profile_ids, base_dir=root,
                    existing_query_policy=existing_policy, backup=True, dry_run=False)
                save_profile(profile, profile_path)

            manager_env = os.environ.copy()
            if args.history_dialog or args.history_restore:
                manager_env['LOCALAPPDATA'] = str(history_root)
            manager = subprocess.Popen([str(executable)], cwd=str(executable.parent), env=manager_env)
            window = Application(backend='uia').connect(process=manager.pid, timeout=20).window(title='Power Query Manager')
            window.wait('visible', timeout=20)
            _assert_main_window_accessible(window)
            print('PASS: packaged main-window UI Automation exposes the workbook, query grid, filter, replace policy, backup, import action and operation log names/roles.')
            manager_hwnd = window.handle
            if args.new_sheet or args.existing_sheet or args.connection_only or args.data_model:
                scrollbars = window.descendants(control_type='ScrollBar')
                if scrollbars:
                    scrollbars[-1].iface_range_value.SetValue(
                        scrollbars[-1].iface_range_value.CurrentMaximum)
                    time.sleep(0.3)
                window.child_window(title='Probe Excel load capabilities', control_type='Button').invoke()
                probe_deadline = time.monotonic() + 120
                capability_text = ''
                while time.monotonic() < probe_deadline:
                    texts = [item.window_text() for item in window.descendants(control_type='Text')]
                    capability_text = next((item for item in texts if item.startswith('Office ') and 'supported:' in item), '')
                    if capability_text:
                        break
                    time.sleep(0.3)
                if not capability_text:
                    raise TimeoutError('Packaged UI capability probe did not finish within 120 seconds.')
                required_mode = ('new_sheet' if args.new_sheet else
                                 'existing_sheet' if args.existing_sheet else
                                 'data_model' if args.data_model else 'connection_only')
                if required_mode not in capability_text:
                    raise RuntimeError(f'Excel capability probe did not support {required_mode}: {capability_text!r}.')
            if args.skip_existing or args.replace_existing or args.replace_preserve_load or args.replace_preserve_model_load or args.replace_preserve_worksheet_load or args.mixed_import or args.new_sheet or args.existing_sheet or args.connection_only or args.data_model:
                scrollbars = window.descendants(control_type='ScrollBar')
                if scrollbars:
                    scrollbars[-1].iface_range_value.SetValue(
                        scrollbars[-1].iface_range_value.CurrentMaximum)
                    time.sleep(0.3)
                window.child_window(title='Load import profile', control_type='Button').invoke()
                picker = _wait_child(window, 'Load import profile')
                filename = picker.child_window(auto_id='1148', control_type='ComboBox').child_window(control_type='Edit')
                filename.set_edit_text(str(profile_path))
                _accept_file_dialog(picker)
                profile_loaded = _wait_child(window, 'Profile loaded')
                profile_loaded.child_window(title='OK', control_type='Button').invoke()
            else:
                window.descendants(control_type='Edit')[0].set_edit_text(str(workbook))
                # Invoke the accessible action even when the scrollable Vector
                # workspace initially positions it outside the visible viewport.
                window.child_window(title='Add files', control_type='Button').invoke()
                picker = _wait_child(window, 'Select Power Query files')
                filename = picker.child_window(auto_id='1148', control_type='ComboBox').child_window(control_type='Edit')
                filename.set_edit_text(str(source))
                _accept_file_dialog(picker)

            table = window.child_window(title='Query source settings table', control_type='Table')
            deadline = time.monotonic() + 10
            cells: list[str] = []
            while time.monotonic() < deadline:
                cells = [cell.window_text() for cell in table.descendants() if cell.element_info.control_type == 'DataItem']
                expected_rows = (['First', 'Second', 'Third'] if args.mixed_import else ['GuiSmoke'])
                if all(name in cells for name in expected_rows):
                    break
                time.sleep(0.2)
            else:
                raise RuntimeError(f'GUI source row did not load: {cells}')

            if args.compare_terminal_newline:
                window.child_window(title='Inspect selected workbook queries read-only', control_type='Button').invoke()
                inventory_dialog = _wait_child(window, 'Workbook query inventory (1)', timeout=30)
                inventory_dialog.child_window(title='Close workbook query inventory', control_type='Button').invoke()
                status_cells = [item.window_text() for item in table.descendants(control_type='DataItem')]
                if 'UNCHANGED' not in status_cells:
                    raise RuntimeError(f'Packaged GUI did not classify the terminal-CRLF roundtrip as UNCHANGED: {status_cells!r}.')
                print('PASS: packaged GUI status is UNCHANGED when Excel removed only the source terminal CRLF.')

            if not (args.skip_existing or args.replace_existing or args.replace_preserve_load or args.replace_preserve_model_load or args.replace_preserve_worksheet_load or args.mixed_import or args.new_sheet or args.existing_sheet or args.connection_only or args.data_model):
                if args.parameter_edit:
                    parameter_copy, expected_source_bytes = _exercise_packaged_parameter_editor(
                        window, root, source, source_hash)
                    source = parameter_copy
                    formula = expected_source_bytes.decode('utf-8').rstrip('\r\n')
                else:
                    _exercise_packaged_modal_accessibility(window)
                if args.package_dialog:
                    window.child_window(title='Compare package versions', control_type='Button').invoke()
                    picker = _wait_child(window, 'Select current query package')
                    picker.child_window(auto_id='1148', control_type='ComboBox').child_window(control_type='Edit').set_edit_text(str(package_pair[0]))
                    _accept_file_dialog(picker)
                    picker = _wait_child(window, 'Select updated query package')
                    picker.child_window(auto_id='1148', control_type='ComboBox').child_window(control_type='Edit').set_edit_text(str(package_pair[1]))
                    _accept_file_dialog(picker)
                    preview = _wait_child(window, 'Query package update preview')
                    _assert_named_controls(preview, [
                        ('Package update summary', 'Text'),
                        ('Query package changes', 'Table'),
                        ('Selected query package diff', 'Edit'),
                        ('No workbook changes safety notice', 'Text'),
                        ('Close query package comparison', 'Button'),
                    ])
                    preview.child_window(title='Close query package comparison', control_type='Button').invoke()
                    print('PASS: packaged query package update preview exposes named summary, changes, diff, safety note and close controls through Windows UI Automation.')
                if args.history_dialog:
                    window.child_window(title='Open operation history', control_type='Button').invoke()
                    history = _wait_child(window, 'Local operation history')
                    rows = history.child_window(title='Local operation history entries', control_type='Table').descendants(control_type='DataItem')
                    snapshot_rows = [index for index in range(0, len(rows), 4)
                                     if index + 3 < len(rows)
                                     and rows[index + 2].window_text() in {'packaged_ui_smoke_before', 'packaged_ui_smoke_after'}]
                    if len(snapshot_rows) != 2:
                        raise RuntimeError(f'Packaged history table did not show two seeded snapshots: {snapshot_rows!r}.')
                    rows[snapshot_rows[0]].click_input()
                    win32api.keybd_event(win32con.VK_SHIFT, 0, 0, 0)
                    try:
                        rows[snapshot_rows[1]].click_input()
                    finally:
                        win32api.keybd_event(win32con.VK_SHIFT, 0, win32con.KEYEVENTF_KEYUP, 0)
                    compare = history.child_window(title='Compare selected workbook versions', control_type='Button')
                    if not compare.is_enabled():
                        state = [(row.window_text(), getattr(row.iface_selection_item, 'CurrentIsSelected', None))
                                 for row in rows]
                        raise RuntimeError(f'Packaged history UIA did not select two snapshot rows: {state!r}.')
                    compare.click_input()
                    deadline = time.monotonic() + 3
                    while compare.is_enabled() and time.monotonic() < deadline:
                        time.sleep(0.1)
                    if compare.is_enabled():
                        raise RuntimeError('Packaged history comparison click did not start the comparison worker.')
                    try:
                        history_compare = _wait_child(window, 'Workbook version comparison', timeout=30)
                    except TimeoutError as exc:
                        completed_events = [entry for entry in history_store.list(limit=100)
                                            if entry['event'] == 'workbook_snapshots_compared']
                        raise TimeoutError(f'{exc} Snapshot comparison history events={completed_events!r}.') from exc
                    _assert_named_controls(history_compare, [
                        ('Query changes between workbook snapshots', 'Table'),
                        ('Selected query formula diff', 'Edit'),
                        ('Close workbook version comparison', 'Button'),
                    ])
                    history_compare.child_window(title='Close workbook version comparison', control_type='Button').invoke()
                    history.child_window(title='Close operation history', control_type='Button').invoke()
                    print('PASS: packaged snapshot comparison exposes named changes table, formula diff and close controls through Windows UI Automation.')
                if args.history_restore:
                    window.child_window(title='Open operation history', control_type='Button').invoke()
                    history = _wait_child(window, 'Local operation history')
                    history_table = history.child_window(title='Local operation history entries', control_type='Table')
                    rows = history_table.descendants(control_type='DataItem')
                    snapshot_rows = [index for index in range(0, len(rows), 4)
                                     if index + 3 < len(rows)
                                     and rows[index + 2].window_text() == 'packaged_ui_smoke_restore_before']
                    if len(snapshot_rows) != 1:
                        raise RuntimeError(f'Packaged restore history did not show its seeded snapshot: {snapshot_rows!r}.')
                    rows[snapshot_rows[0]].click_input()
                    restore = history.child_window(title='Restore selected workbook version', control_type='Button')
                    if not restore.is_enabled():
                        raise RuntimeError('Packaged history snapshot row did not enable restore.')
                    restore.invoke()
                    confirmation = _wait_child(window, 'Restore workbook version')
                    confirmation_text = '\n'.join(item.window_text() for item in confirmation.descendants(control_type='Text'))
                    if 'verified timestamped backup' not in confirmation_text or 'hash-verified' not in confirmation_text:
                        raise RuntimeError(f'Packaged restore confirmation omitted its backup/hash safeguards: {confirmation_text!r}.')
                    confirmation.child_window(title='Yes', control_type='Button').invoke()
                    restored = _wait_child(window, 'Version restored', timeout=30)
                    restored_text = '\n'.join(item.window_text() for item in restored.descendants(control_type='Text'))
                    backup_match = re.search(r'Backup:\s*(.+)', restored_text)
                    if not backup_match:
                        raise RuntimeError(f'Packaged restore did not report its verified backup path: {restored_text!r}.')
                    backup_path = Path(backup_match.group(1).strip())
                    if _sha256(workbook) != restore_snapshot_hash:
                        raise RuntimeError('Packaged snapshot restore did not restore the original disposable workbook hash.')
                    if not backup_path.is_file() or _sha256(backup_path) != restore_expected_backup_hash:
                        raise RuntimeError('Packaged snapshot restore backup did not match the workbook bytes before restore.')
                    restored.child_window(title='OK', control_type='Button').invoke()
                    restore_deadline = time.monotonic() + 10
                    while not restore.is_enabled() and time.monotonic() < restore_deadline:
                        time.sleep(0.1)
                    if not restore.is_enabled():
                        raise TimeoutError('Packaged snapshot restore worker did not return the history UI to idle.')
                    history.child_window(title='Close operation history', control_type='Button').invoke()
                    print('PASS: packaged history restore returned the disposable workbook to its snapshot hash and preserved an exact pre-restore backup through Windows UI Automation.')

            scrollbars = window.descendants(control_type='ScrollBar')
            if scrollbars:
                main_scrollbar = scrollbars[-1]
                main_scrollbar.iface_range_value.SetValue(main_scrollbar.iface_range_value.CurrentMaximum)
                time.sleep(0.3)
            review_button = window.child_window(title='Review and import selected queries', control_type='Button')
            if not review_button.is_enabled():
                raise RuntimeError('Packaged UI import action is unexpectedly disabled after source review.')
            review_button.invoke()
            confirmation = _wait_child(window, 'Review import sequence')
            if args.mixed_import:
                confirmation_text = '\n'.join(
                    child.window_text() for child in confirmation.descendants(control_type='Text'))
                expected_order = ['1. First — 01_First.m', '2. Second — 02_Second.pq', '3. Third — 03_Third.m']
                positions = [confirmation_text.find(item) for item in expected_order]
                if any(position < 0 for position in positions) or positions != sorted(positions):
                    raise RuntimeError(f'Packaged UI review did not show the expected numbered mixed-source order: {confirmation_text!r}')
            confirmation.child_window(title='Yes', control_type='Button').invoke()

            deadline = time.monotonic() + 120
            final_status = ''
            completed_dialog = False
            expected_dialog = 'No changes' if args.skip_existing else 'Import complete'
            while time.monotonic() < deadline:
                time.sleep(0.4)
                final_status = window.child_window(control_type='StatusBar').window_text()
                children = [child.window_text() for child in window.descendants(control_type='Window')]
                if expected_dialog in children:
                    # Read the terminal status after the completion dialog is
                    # visible. A sample taken just before the worker's
                    # completion signal can still contain the startup status.
                    final_status = window.child_window(control_type='StatusBar').window_text()
                    completed_dialog = True
                    window.child_window(title=expected_dialog, control_type='Window').child_window(
                        title='OK', control_type='Button').invoke()
                    break
                if 'Import failed' in children or final_status == 'Import failed; review the operation log':
                    failure_details = []
                    for child in window.descendants(control_type='Window'):
                        if child.window_text() == 'Import failed':
                            for detail in child.descendants():
                                try:
                                    value = detail.window_text()
                                    if value:
                                        failure_details.append(value)
                                except Exception:
                                    pass
                    operation_log = '<operation log unavailable>'
                    for item in window.descendants():
                        if item.element_info.name != 'Operation log output':
                            continue
                        values = [item.window_text()]
                        for pattern in ('value', 'text'):
                            try:
                                values.append(str(getattr(item, f'iface_{pattern}').CurrentValue))
                            except Exception:
                                pass
                        operation_log = ' | '.join(value for value in values if value)
                        break
                    raise RuntimeError(
                        'Packaged GUI reported an import failure: ' + final_status
                        + '; dialog: ' + repr(failure_details)
                        + '; operation log: ' + operation_log)
            else:
                raise TimeoutError(f'Packaged GUI did not complete the import; status={final_status!r}.')

            if not completed_dialog:
                raise RuntimeError(f'Packaged GUI did not display its completion result; status={final_status!r}.')
            if args.skip_existing and 'No changes' not in final_status:
                raise RuntimeError(f'Packaged GUI skip status was not truthful: {final_status!r}.')
            from packaged_smoke_support import read_optional_operation_log
            operation_log = read_optional_operation_log(
                window,
                until_contains=None if args.skip_existing else 'Excel automation finished.',
            )
            if operation_log is None:
                print('INFO: packaged operation log unavailable through UI Automation.')
                if not args.skip_existing:
                    raise RuntimeError('Packaged operation log was not accessible through UI Automation after a successful import.')
            elif not args.skip_existing and 'Excel automation finished.' not in operation_log:
                raise RuntimeError('Packaged operation log did not expose its final completion record through UI Automation.')
            else:
                print('INFO: packaged operation log: ' + operation_log)
            time.sleep(0.5)
            if manager.poll() is None:
                try:
                    handle=window.handle
                except Exception:
                    # The manager may exit between closing its completion dialog
                    # and resolving the top-level window. The process handle is the
                    # authoritative completion signal; a vanished HWND is benign.
                    handle=None
                if handle and win32gui.IsWindow(handle):
                    win32gui.SendMessage(handle, win32con.WM_CLOSE, 0, 0)
            try:
                manager.wait(timeout=2)
            except subprocess.TimeoutExpired:
                # The completion dialog can close just before the worker's finished
                # signal releases the UI busy state; retry only while our process lives.
                time.sleep(0.5)
                if manager.poll() is None:
                    try:
                        handle=window.handle
                    except Exception:
                        handle=None
                    if handle and win32gui.IsWindow(handle):
                        win32gui.SendMessage(handle, win32con.WM_CLOSE, 0, 0)
                manager.wait(timeout=10)
            if manager.returncode != 0:
                raise RuntimeError(f'Packaged GUI exited with code {manager.returncode}.')
            manager = None
            print(f'INFO: packaged manager exited 0; process set={sorted(_excel_pids())}.')

            if args.refresh_data_model:
                cli_executable = executable.with_name('PowerQueryManagerCLI.exe')
                if not cli_executable.is_file():
                    raise RuntimeError(f'Packaged CLI executable not found beside GUI: {cli_executable}')
                before_refresh_hash = _sha256(workbook)
                prior_backups = {item.resolve() for item in workbook.parent.glob(f'{workbook.stem}_backup_*{workbook.suffix}')}
                refresh_command = [str(cli_executable), 'refresh', '--apply', '--query', 'GuiSmoke',
                    '--workbook', str(workbook), '--timeout', '30', '--confirm', 'REFRESH']
                refresh_env = os.environ.copy()
                refresh_env['LOCALAPPDATA'] = str(root / 'localappdata')
                refresh_env['APPDATA'] = str(root / 'localappdata')
                refreshed = subprocess.run(refresh_command, capture_output=True, text=True,
                    encoding='utf-8', errors='replace', env=refresh_env, timeout=90)
                if refreshed.returncode != 0:
                    raise RuntimeError(f'Packaged Data Model refresh returned {refreshed.returncode}:\n'
                        f'{refreshed.stdout}\n{refreshed.stderr}')
                json_start = refreshed.stdout.find('{')
                if json_start < 0:
                    raise RuntimeError(f'Packaged Data Model refresh returned no JSON: {refreshed.stdout!r}')
                refresh_report = json.loads(refreshed.stdout[json_start:])
                results = refresh_report.get('results', [])
                if len(results) != 1 or results[0].get('status') != 'refreshed':
                    raise RuntimeError(f'Packaged Data Model refresh did not confirm success: {refresh_report!r}')
                created_backups = [item for item in workbook.parent.glob(f'{workbook.stem}_backup_*{workbook.suffix}')
                    if item.resolve() not in prior_backups]
                if len(created_backups) != 1 or _sha256(created_backups[0]) != before_refresh_hash:
                    raise RuntimeError('Data Model refresh backup did not match exact pre-refresh workbook bytes.')
                if _sha256(workbook) == before_refresh_hash:
                    raise RuntimeError('Data Model refresh did not commit a staged workbook change.')
                if list(workbook.parent.glob(f'.{workbook.stem}.pqstage-*')):
                    raise RuntimeError('Data Model refresh left a staged workbook behind.')
                print(f'PASS: packaged Data Model refresh report: {results[0]!r}; backup hash verified.')

            sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'app' / 'src'))
            from pq_installer.core.excel_inspector import ExcelQueryInspector
            from pq_installer.core.excel_queries import _formula_matches_excel_roundtrip
            inventory = ExcelQueryInspector().inspect(workbook)
            matches = [query for query in inventory if query.name == 'GuiSmoke']
            if args.mixed_import:
                expected = {name: code for _source, name, code in [
                    (root / '01_First.m', 'First', 'let Source = #table({"Value"}, {{1}}) in Source'),
                    (root / '02_Second.pq', 'Second', 'let Source = #table({"Value"}, {{2}}) in Source'),
                    (root / '03_Third.m', 'Third', 'let Source = #table({"Value"}, {{3}}) in Source'),
                ]}
                observed_names = [item.name for item in inventory]
                if observed_names != list(expected):
                    raise RuntimeError(f'Excel query collection order differs from numbered import order: {observed_names!r}.')
                observed = {item.name: item.formula for item in inventory}
                if observed != expected:
                    raise RuntimeError(f'Excel readback formulas differ from mixed source bytes: {observed!r}.')
            else:
                expected_formula = existing_formula if args.skip_existing else formula
                if len(matches) != 1 or not _formula_matches_excel_roundtrip(
                        expected_formula, matches[0].formula):
                    raise RuntimeError('Read-only Excel reopen did not verify the expected query formula.')
                if (args.bom_crlf_source or args.compare_terminal_newline) and _sha256(source) != source_hash:
                    raise RuntimeError('Source file bytes changed during GUI import.')
                if args.replace_preserve_load and matches[0].load_state != 'connection_only':
                    raise RuntimeError(f'Existing connection-only load changed after GUI replacement: {matches[0].load_state!r}.')
                if args.replace_preserve_model_load and matches[0].load_state != 'data_model':
                    raise RuntimeError(f'Existing Data Model load changed after GUI replacement: {matches[0].load_state!r}.')
                if args.replace_preserve_worksheet_load and not matches[0].load_state.startswith('worksheet: Output!$C$4:'):
                    raise RuntimeError(f'Existing worksheet load changed after GUI replacement: {matches[0].load_state!r}.')
                if args.new_sheet and not matches[0].load_state.startswith('worksheet: GuiOutput!'):
                    raise RuntimeError(f'New worksheet load was not observed after GUI import: {matches[0].load_state!r}.')
                if args.existing_sheet and not matches[0].load_state.startswith('worksheet: Output!$C$4:'):
                    raise RuntimeError(f'Existing worksheet load was not observed at Output!C4: {matches[0].load_state!r}.')
                if args.connection_only and matches[0].load_state != 'connection_only':
                    raise RuntimeError(f'Connection-only load was not observed after GUI import: {matches[0].load_state!r}.')
                if args.data_model and matches[0].load_state != 'data_model':
                    raise RuntimeError(f'Data Model load was not observed after GUI import: {matches[0].load_state!r}.')
            verify_excel = win32com.client.DispatchEx('Excel.Application')
            verify_pid = int(win32process.GetWindowThreadProcessId(int(verify_excel.Hwnd))[1])
            print(f'INFO: read-only verification Excel PID {verify_pid}; process set={sorted(_excel_pids())}.')
            verify_book = None
            try:
                verify_excel.Visible = False
                verify_excel.DisplayAlerts = False
                verify_excel.AutomationSecurity = 3
                verify_book = verify_excel.Workbooks.Open(str(workbook), UpdateLinks=0, ReadOnly=True)
                marker = next((sheet.Range('A1').Value2 for sheet in verify_book.Worksheets if sheet.Range('A1').Value2 == 'Power Query Manager GUI smoke marker'), None)
                if marker != 'Power Query Manager GUI smoke marker':
                    raise RuntimeError(f'Unrelated worksheet marker changed: {marker!r}.')
                if args.new_sheet or args.existing_sheet:
                    output_sheet = verify_book.Worksheets.Item('GuiOutput' if args.new_sheet else 'Output')
                    output_table = output_sheet.ListObjects.Item('tblGuiSmoke')
                    table_anchor = str(output_table.Range.Cells(1, 1).Address)
                    expected_anchor = '$A$1' if args.new_sheet else '$C$4'
                    if table_anchor != expected_anchor:
                        raise RuntimeError(f'Worksheet query table anchor differs: expected {expected_anchor}, got {table_anchor}.')
                    values = [int(output_table.DataBodyRange.Cells(row, 1).Value2)
                              for row in range(1, int(output_table.DataBodyRange.Rows.Count) + 1)]
                    if values != [17, 29]:
                        raise RuntimeError(f'New worksheet table values differ from static M source: {values!r}.')
                if args.connection_only:
                    connection = verify_book.Connections.Item('Query - GuiSmoke')
                    if str(connection.Name).casefold() != 'query - guismoke':
                        raise RuntimeError(f'Connection-only import created an unexpected connection: {connection.Name!r}.')
                    if any(str(item.SourceWorkbookConnection.Name).casefold() == 'query - guismoke'
                           for item in (verify_book.Model.ModelTables.Item(index)
                                        for index in range(1, int(verify_book.Model.ModelTables.Count) + 1))):
                        raise RuntimeError('Connection-only import created a Data Model table.')
                    for sheet in verify_book.Worksheets:
                        for item in sheet.ListObjects:
                            if str(item.QueryTable.WorkbookConnection.Name).casefold() == 'query - guismoke':
                                raise RuntimeError('Connection-only import created a worksheet output table.')
                if args.data_model:
                    connection=verify_book.Connections.Item('Query - GuiSmoke')
                    if not bool(connection.InModel):
                        raise RuntimeError('Data Model import did not mark its workbook connection as InModel.')
                    model_tables=tuple(verify_book.Model.ModelTables.Item(index)
                        for index in range(1,int(verify_book.Model.ModelTables.Count)+1))
                    model_table=next((item for item in model_tables
                        if str(item.SourceWorkbookConnection.Name).casefold()=='query - guismoke'),None)
                    if model_table is None or int(model_table.RecordCount)!=2:
                        observed=None if model_table is None else int(model_table.RecordCount)
                        raise RuntimeError(f'Data Model table membership/row count is incorrect: {observed!r}.')
                if args.replace_preserve_model_load:
                    connection = verify_book.Connections.Item('Query - GuiSmoke')
                    if not bool(connection.InModel):
                        raise RuntimeError('Existing query replacement removed its Data Model association.')
                    model_tables = tuple(verify_book.Model.ModelTables.Item(index)
                        for index in range(1, int(verify_book.Model.ModelTables.Count) + 1))
                    model_table = next((item for item in model_tables
                        if str(item.SourceWorkbookConnection.Name).casefold() == 'query - guismoke'), None)
                    if model_table is None or int(model_table.RecordCount) != 1:
                        observed = None if model_table is None else int(model_table.RecordCount)
                        raise RuntimeError(f'Replacement implicitly changed or removed the existing Data Model rows: {observed!r}.')
                if args.replace_preserve_worksheet_load:
                    output_sheet = verify_book.Worksheets.Item('Output')
                    output_table = output_sheet.ListObjects.Item('tblGuiSmoke')
                    if str(output_table.Range.Cells(1, 1).Address).replace('$', '').upper() != 'C4':
                        raise RuntimeError('Existing worksheet table moved during definition replacement.')
                    if int(output_table.DataBodyRange.Rows.Count) != 1 or int(output_table.DataBodyRange.Cells(1, 1).Value2) != 5:
                        raise RuntimeError('Existing worksheet rows changed during definition replacement; an implicit refresh may have occurred.')
                if args.vba:
                    with zipfile.ZipFile(workbook) as archive:
                        if 'xl/vbaProject.bin' not in archive.namelist():
                            raise RuntimeError('The packaged GUI import dropped xl/vbaProject.bin.')
                    code_module = verify_book.VBProject.VBComponents.Item('ThisWorkbook').CodeModule
                    source_lines = code_module.Lines(1, int(code_module.CountOfLines))
                    if 'Workbook_Open' not in source_lines or 'MACRO_RAN' not in source_lines:
                        raise RuntimeError('The packaged GUI import changed the Workbook_Open VBA source.')
                verify_book.Close(False)
                verify_book = None
            finally:
                if verify_book is not None:
                    try:
                        verify_book.Close(False)
                    except Exception:
                        pass
                output_sheet = output_table = connection = sheet = item = model_table = model_tables = None
                import gc
                gc.collect()
                from pq_installer.core.excel_lifetime import quit_owned_excel
                apartment_released = False
                def release_verify_apartment():
                    nonlocal verify_excel, apartment_released
                    verify_excel = None
                    gc.collect()
                    pythoncom.CoUninitialize()
                    apartment_released = True
                try:
                    needed_termination = quit_owned_excel(verify_excel, after_quit=release_verify_apartment)
                finally:
                    if apartment_released:
                        pythoncom.CoInitialize()
                print(f'INFO: verification Excel PID {verify_pid} cleanup termination={needed_termination}; process set={sorted(_excel_pids())}.')
                if needed_termination:
                    raise RuntimeError('Read-only verification Excel required forced termination.')
            backups = workbook.parent.glob(f'{workbook.stem}_backup_*{workbook.suffix}')
            if args.skip_existing:
                if _sha256(workbook) != original_hash:
                    raise RuntimeError('All-skip GUI operation changed the target workbook bytes.')
                if workbook.stat().st_mtime_ns != original_mtime:
                    raise RuntimeError('All-skip GUI operation changed the target workbook timestamp.')
                readonly_flag = getattr(stat, 'FILE_ATTRIBUTE_READONLY', 0x1)
                if not workbook.stat().st_file_attributes & readonly_flag:
                    raise RuntimeError('All-skip GUI operation cleared the workbook read-only attribute.')
                if list(backups):
                    raise RuntimeError('All-skip GUI operation created an unnecessary workbook backup.')
            else:
                if _sha256(workbook) == original_hash:
                    raise RuntimeError('Target workbook hash did not change after import.')
                if not any(_sha256(backup) == original_hash for backup in backups):
                    raise RuntimeError('No timestamped backup matched the pre-import workbook hash.')
            cleanup_deadline = time.monotonic() + 15
            final_excel_pids = _excel_pids()
            while final_excel_pids != excel_before and time.monotonic() < cleanup_deadline:
                time.sleep(0.25)
                final_excel_pids = _excel_pids()
            if final_excel_pids != excel_before:
                raise RuntimeError(f'Excel process set changed; initial={sorted(excel_before)}, final={sorted(final_excel_pids)}.')
            if args.skip_existing:
                os.chmod(workbook, stat.S_IREAD | stat.S_IWRITE)
            vba_result = 'VBA source preserved and Workbook_Open marker unchanged, ' if args.vba else ''
            if args.skip_existing:
                print(f'PASS: packaged GUI all-skip on Excel {office_version} build {office_build}; '
                      f'{workbook.suffix} exact workbook hash/timestamp and existing formula preserved, '
                      'no backup or stage, no-change UI result, manager exit 0, and Excel process cleanup.')
            elif args.mixed_import:
                print(f'PASS: packaged GUI mixed .m/.pq import on Excel {office_version} build {office_build}; '
                      'numbered review order and Excel query order match; exact source formulas, verified backup, '
                      'read-only reopen, manager exit 0, and Excel process cleanup.')
            elif args.replace_existing:
                if args.compare_terminal_newline:
                    print(f'PASS: packaged GUI terminal-newline comparison/import on Excel {office_version} build {office_build}; '
                          'UI reported UNCHANGED; Excel formula matched source after terminal-CRLF normalization, '
                          'original-hash backup, read-only reopen, manager exit 0, and Excel process cleanup passed.')
                else:
                    print(f'PASS: packaged GUI existing-query replacement on Excel {office_version} build {office_build}; '
                          'new exact formula read back, verified pre-import backup, unrelated worksheet marker, '
                          'manager exit 0, and Excel process cleanup.')
            elif args.replace_preserve_load:
                print(f'PASS: packaged GUI existing-query replacement on Excel {office_version} build {office_build}; '
                      'exact changed formula read back and existing connection-only load preserved, '
                      'verified pre-import backup, manager exit 0, and Excel process cleanup.')
            elif args.replace_preserve_model_load:
                print(f'PASS: packaged GUI existing-query replacement on Excel {office_version} build {office_build}; '
                      'exact changed formula read back; existing Data Model association and original one-row model data preserved without implicit refresh, '
                      'verified pre-import backup, read-only reopen and Excel process cleanup.')
            elif args.replace_preserve_worksheet_load:
                print(f'PASS: packaged GUI existing-query replacement on Excel {office_version} build {office_build}; '
                      'exact changed formula read back; worksheet table identity, Output!C4 anchor and original one-row data preserved without implicit refresh, '
                      'verified pre-import backup, read-only reopen and Excel process cleanup.')
            elif args.new_sheet:
                print(f'PASS: packaged GUI new-worksheet load on Excel {office_version} build {office_build}; '
                      'capability probe enabled new_sheet, exact formula and worksheet load state read back, '
                      'table values 17/29 verified, backup hash matched, manager exit 0, and Excel process cleanup.')
            elif args.connection_only:
                print(f'PASS: packaged GUI connection-only load on Excel {office_version} build {office_build}; '
                      'exact formula and connection-only state read back, workbook connection exists without a worksheet or model table, '
                      'backup hash matched, manager exit 0, and Excel process cleanup.')
            elif args.data_model:
                print(f'PASS: packaged GUI Data Model load on Excel {office_version} build {office_build}; '
                      'capability probe enabled data_model, exact formula and model membership/row count 2 verified, '
                      'backup hash matched, manager exit 0, and Excel process cleanup.')
            elif args.existing_sheet:
                print(f'PASS: packaged GUI existing-worksheet load on Excel {office_version} build {office_build}; '
                      'capability probe enabled existing_sheet, exact formula and Output!C4 load state read back, '
                      'table anchor and values 17/29 verified, unrelated marker preserved, backup hash matched, '
                      'manager exit 0, and Excel process cleanup.')
            else:
                source_detail = ('UTF-8 BOM/CRLF M bytes and exact formula readback, '
                                 if args.bom_crlf_source else '')
                print(f'PASS: packaged GUI definition-only import on Excel {office_version} build {office_build}; '
                      f'{workbook.suffix} {vba_result}{source_detail}exact formula, verified backup, read-only reopen, '
                      'manager exit 0, and Excel process cleanup.')
            return 0
    finally:
        if manager is not None and manager.poll() is None:
            try:
                if manager_hwnd:
                    win32gui.SendMessage(manager_hwnd, win32con.WM_CLOSE, 0, 0)
                manager.wait(timeout=5)
            except Exception:
                manager.terminate()
        pythoncom.CoUninitialize()


if __name__ == '__main__':
    raise SystemExit(main())
