"""Verify a real Excel all-skip import leaves the original workbook byte-identical."""
from __future__ import annotations

import hashlib
import os
from pathlib import Path
import tempfile
import time

import pythoncom
import win32api
import win32con
import win32event
import win32process
import win32com.client

from pq_installer.core.excel_inspector import ExcelQueryInspector
from pq_installer.core.excel_queries import ExcelQueryInstaller
from pq_installer.core.models import ExistingQueryMode, QuerySource, QueryStatus


def _sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def _running(pid: int) -> bool:
    try:
        handle = win32api.OpenProcess(win32con.SYNCHRONIZE, False, pid)
    except Exception:
        return False
    try:
        return win32event.WaitForSingleObject(handle, 0) == win32event.WAIT_TIMEOUT
    finally:
        win32api.CloseHandle(handle)


def _wait_owned_processes(pids: set[int]) -> None:
    deadline = time.monotonic() + 15
    while any(_running(pid) for pid in pids) and time.monotonic() < deadline:
        time.sleep(0.2)
    remaining = sorted(pid for pid in pids if _running(pid))
    if remaining:
        raise RuntimeError(f'Smoke-owned Excel processes remained after cleanup: {remaining}')


def main() -> int:
    pythoncom.CoInitialize()
    owned_pids: set[int] = set()
    original_dispatch = win32com.client.DispatchEx

    def tracked_dispatch(*args, **kwargs):
        application = original_dispatch(*args, **kwargs)
        owned_pids.add(int(win32process.GetWindowThreadProcessId(int(application.Hwnd))[1]))
        return application

    win32com.client.DispatchEx = tracked_dispatch
    try:
        with tempfile.TemporaryDirectory(prefix='pqm-skip-noop-smoke-') as temporary:
            root = Path(temporary)
            workbook_path = root / 'skip-noop.xlsx'
            source_path = root / 'Query.m'
            original_formula = 'let Source = #table({"Value"}, {{17}}) in Source'
            replacement_formula = 'let Source = #table({"Value"}, {{29}}) in Source'
            source_path.write_text(original_formula, encoding='utf-8')

            excel = win32com.client.DispatchEx('Excel.Application')
            workbook = None
            version, build = str(excel.Version), str(excel.Build)
            try:
                excel.Visible = False
                excel.DisplayAlerts = False
                workbook = excel.Workbooks.Add()
                workbook.Worksheets(1).Range('A1').Value = 'Skip no-op marker'
                workbook.SaveAs(str(workbook_path), 51)
                workbook.Close(False)
                workbook = None
            finally:
                if workbook is not None:
                    workbook.Close(False)
                excel.Quit()
                excel = None

            created = ExcelQueryInstaller().install(
                workbook_path, [QuerySource(source_path, 'Query')], ExistingQueryMode.REPLACE)
            if len(created) != 1 or created[0].status != QueryStatus.CREATED:
                raise RuntimeError(f'Initial disposable query creation failed: {created!r}')
            before_skip_hash = _sha256(workbook_path)
            before_skip_mtime = workbook_path.stat().st_mtime_ns
            source_path.write_text(replacement_formula, encoding='utf-8')

            skipped = ExcelQueryInstaller().install(
                workbook_path, [QuerySource(source_path, 'Query')], ExistingQueryMode.SKIP)
            if len(skipped) != 1 or skipped[0].status != QueryStatus.SKIPPED:
                raise RuntimeError(f'Existing query was not reported as skipped: {skipped!r}')
            if _sha256(workbook_path) != before_skip_hash:
                raise RuntimeError('All-skip operation changed the workbook bytes.')
            if workbook_path.stat().st_mtime_ns != before_skip_mtime:
                raise RuntimeError('All-skip operation changed the original workbook timestamp.')
            if list(root.glob('*.pqstage-*')):
                raise RuntimeError('All-skip operation left a staged workbook behind.')

            inventory = ExcelQueryInspector().inspect(workbook_path)
            query = next((row for row in inventory if row.name == 'Query'), None)
            if query is None or query.formula != original_formula:
                raise RuntimeError('Read-only verification found that the skipped query formula changed.')

            _wait_owned_processes(owned_pids)
            print(f'PASS: Excel {version} build {build}; all-skip import preserved workbook SHA-256 and timestamp, kept the existing formula, removed its stage, and exited all smoke-owned Excel processes.')
            return 0
    finally:
        win32com.client.DispatchEx = original_dispatch
        pythoncom.CoUninitialize()


if __name__ == '__main__':
    raise SystemExit(main())
