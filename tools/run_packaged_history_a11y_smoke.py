"""Inspect history-dialog accessibility names from the packaged Windows GUI via UI Automation."""
from __future__ import annotations

import argparse
import os
from pathlib import Path
import shutil
import subprocess
import tempfile
import time

from pywinauto import Application


def _wait_child(parent, title: str, control_type: str, timeout: float = 12):
    deadline = time.monotonic() + timeout
    child = parent.child_window(title=title, control_type=control_type)
    while time.monotonic() < deadline:
        if child.exists():
            return child
        time.sleep(0.2)
    names = [(item.element_info.name, item.element_info.control_type)
             for item in parent.descendants()]
    raise TimeoutError(f'UIA control {title!r}/{control_type} not found; observed={names!r}')


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--exe', type=Path, required=True)
    parser.add_argument('--corrupt-history', action='store_true',
                        help='Seed corrupt local history and verify recoverable archive behavior.')
    args = parser.parse_args()
    executable = args.exe.resolve()
    if not executable.is_file():
        parser.error(f'Packaged executable not found: {executable}')

    manager = None
    with tempfile.TemporaryDirectory(prefix='pqm-uia-history-') as temporary:
        local_app_data = Path(temporary) / 'LocalAppData'
        local_app_data.mkdir()
        environment = os.environ.copy()
        environment['LOCALAPPDATA'] = str(local_app_data)
        corrupt_bytes = b'packaged history recovery smoke: corrupt SQLite bytes\x00'
        history_dir = local_app_data / 'PowerQueryManager'
        if args.corrupt_history:
            history_dir.mkdir(parents=True)
            (history_dir / 'history.sqlite3').write_bytes(corrupt_bytes)
        manager = subprocess.Popen([str(executable)], cwd=str(executable.parent), env=environment)
        try:
            window = Application(backend='uia').connect(process=manager.pid, timeout=20).window(
                title='Power Query Manager')
            window.wait('visible', timeout=20)
            _wait_child(window, 'Open operation history', 'Button').invoke()
            dialog = _wait_child(window, 'Local operation history', 'Window')
            dialog.wait('visible', timeout=10)
            expected = {
                'Local operation history entries': 'Table',
                'History retention period in days': 'Spinner',
                'Capture optional full workbook snapshots before changes': 'CheckBox',
                'Refresh history entries': 'Button',
                'Apply history retention period': 'Button',
                'Compare selected workbook versions': 'Button',
                'Restore selected workbook version': 'Button',
                'Delete selected history entry': 'Button',
                'Delete all local history': 'Button',
                'Close operation history': 'Button',
            }
            observed = {(item.element_info.name, item.element_info.control_type)
                        for item in dialog.descendants()}
            missing = [name for name, role in expected.items() if (name, role) not in observed]
            if missing:
                raise AssertionError(f'History dialog UIA names/roles missing: {missing!r}; observed={sorted(observed)!r}')
            if args.corrupt_history:
                import sqlite3
                archives = list(history_dir.glob('history.corrupt-*'))
                if len(archives) != 1 or archives[0].read_bytes() != corrupt_bytes:
                    raise AssertionError(f'Corrupt history archive missing or changed: {archives!r}')
                database = sqlite3.connect(history_dir / 'history.sqlite3')
                try:
                    settings = database.execute("SELECT value FROM settings WHERE key='retention_days'").fetchone()
                    if settings is None:
                        raise AssertionError('Fresh local history store was not initialized.')
                finally:
                    database.close()
                exposed_text = []
                for item in window.descendants():
                    try:
                        exposed_text.extend((item.window_text(), item.element_info.name))
                    except Exception:
                        continue
                if not any(archives[0].name in value for value in exposed_text):
                    raise AssertionError(f'UI did not report the preserved history archive filename; exposed={[(item.element_info.control_type, item.element_info.name) for item in window.descendants() if item.element_info.name]!r}')
            dialog.child_window(title='Close operation history', control_type='Button').invoke()
            window.close()
            manager.wait(timeout=10)
            if manager.returncode != 0:
                raise RuntimeError(f'Packaged GUI exited with {manager.returncode}.')
            manager = None
            if list(local_app_data.rglob('*')) == []:
                raise AssertionError('History dialog did not use the isolated LOCALAPPDATA test directory.')
            suffix = ' Corrupt history was preserved and recovery was reported.' if args.corrupt_history else ''
            print('PASS: packaged history dialog exposed the table and snapshot/retention actions through Windows UI Automation; manager exited 0; history stayed under the temporary test profile.' + suffix)
            return 0
        finally:
            if manager is not None and manager.poll() is None:
                manager.terminate()
                try:
                    manager.wait(timeout=5)
                except subprocess.TimeoutExpired:
                    manager.kill()
                    manager.wait(timeout=5)


if __name__ == '__main__':
    raise SystemExit(main())
