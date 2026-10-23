"""Verify local snapshot restore against a disposable workbook in desktop Excel."""
from __future__ import annotations

import hashlib
import os
from pathlib import Path
import subprocess
import sys
import tempfile
import time

import pythoncom
import win32api
import win32con
import win32event
import win32process
import win32com.client

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / 'app' / 'src'))
from pq_installer.core.history import HistoryStore


def _hash(path):
    digest = hashlib.sha256()
    with Path(path).open('rb') as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b''):
            digest.update(chunk)
    return digest.hexdigest()


def _running(pid):
    try:
        handle = win32api.OpenProcess(win32con.SYNCHRONIZE, False, pid)
    except Exception:
        return False
    try:
        return win32event.WaitForSingleObject(handle, 0) == win32event.WAIT_TIMEOUT
    finally:
        win32api.CloseHandle(handle)


def _excel_with_cell(path, value, *, read_only=False):
    app = win32com.client.DispatchEx('Excel.Application')
    app.Visible = False
    app.DisplayAlerts = False
    pid = win32process.GetWindowThreadProcessId(app.Hwnd)[1]
    book = None
    try:
        if Path(path).exists():
            book = app.Workbooks.Open(str(path), UpdateLinks=0, ReadOnly=read_only,
                                      IgnoreReadOnlyRecommended=True, AddToMru=False)
        else:
            book = app.Workbooks.Add()
        if value is not None:
            book.Worksheets(1).Range('A1').Value = value
            book.SaveAs(str(path), FileFormat=51)
        result = book.Worksheets(1).Range('A1').Value
        book.Close(SaveChanges=False)
        book = None
        return result, pid, str(app.Version), str(app.Build)
    finally:
        if book is not None:
            book.Close(SaveChanges=False)
        app.Quit()
        del app
        for _ in range(50):
            if not _running(pid):
                break
            time.sleep(.1)
        if _running(pid):
            raise AssertionError(f'Excel process created by smoke did not exit: {pid}')


def main():
    pythoncom.CoInitialize()
    try:
        with tempfile.TemporaryDirectory(prefix='pqm-history-restore-') as temp:
            root = Path(temp)
            workbook = root / 'snapshot-restore.xlsx'
            _, _, version, build = _excel_with_cell(workbook, 'before')
            store = HistoryStore(root / 'history.sqlite3')
            store.set_snapshots_enabled(True)
            snapshot = store.create_snapshot(workbook, event='workbook_snapshot_before_import')
            _excel_with_cell(workbook, 'after')
            current_hash = _hash(workbook)
            outcome = store.restore_snapshot(snapshot['snapshot_id'], workbook, overwrite=True,
                                             expected_current_sha256=current_hash)
            value, _, _, _ = _excel_with_cell(workbook, None, read_only=True)
            assert value == 'before', f'restored cell value mismatch: {value!r}'
            assert _hash(workbook) == snapshot['sha256'], 'restored workbook hash mismatch'
            assert outcome['backup_path'] and _hash(outcome['backup_path']) == current_hash, 'pre-restore backup mismatch'
            assert store.list(limit=10), 'restore event missing from audit history'
            print(f'PASS Excel {version} build {build}; restored A1={value!r}; '
                  f'snapshot={snapshot["sha256"]}; backup={current_hash}; owned Excel processes exited')
    finally:
        pythoncom.CoUninitialize()


if __name__ == '__main__':
    main()
