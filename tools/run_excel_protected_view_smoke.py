"""Open a blank disposable workbook in Excel's explicit Protected View UI."""
from __future__ import annotations

import hashlib
import gc
import tempfile
import time
from pathlib import Path

import pythoncom
import win32api
import win32con
import win32event
import win32process
import win32com.client
from pywinauto import Application


def _sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def main() -> int:
    pythoncom.CoInitialize()
    excel = book = protected = None
    owned_pid = None
    excel_version = excel_build = "unknown"
    try:
        with tempfile.TemporaryDirectory(prefix="pqm-protected-view-") as temporary:
            path = Path(temporary) / "disposable.xlsx"
            excel = win32com.client.DispatchEx("Excel.Application")
            owned_pid = int(win32process.GetWindowThreadProcessId(int(excel.Hwnd))[1])
            excel_version, excel_build = str(excel.Version), str(excel.Build)
            excel.Visible = False
            excel.DisplayAlerts = False
            excel.AutomationSecurity = 3
            book = excel.Workbooks.Add()
            book.SaveAs(str(path), 51)
            book.Close(False)
            book = None
            digest = _sha256(path)
            excel.Visible = True
            protected = excel.ProtectedViewWindows.Open(str(path), None, False, False)
            time.sleep(1.0)

            if int(excel.ProtectedViewWindows.Count) != 1:
                raise AssertionError("Excel did not retain the disposable workbook in Protected View.")
            caption = str(protected.Caption)
            application = Application(backend="uia").connect(process=owned_pid, timeout=15)
            visible_windows = application.windows(visible_only=True)
            ui_text = []
            for window in visible_windows:
                ui_text.append(window.window_text())
                try:
                    ui_text.extend(child.window_text() for child in window.descendants())
                except Exception:
                    pass
            matching = [text for text in ui_text if "protected view" in text.casefold()]
            if not matching:
                raise AssertionError(
                    f"No visible Excel UI Automation text identified Protected View: {ui_text[:40]!r}"
                )

            # Close the explicitly opened, owned Protected View window without enabling editing.
            protected.Close()
            protected = None
            excel.Quit()
            excel = None
            handle = win32api.OpenProcess(win32con.SYNCHRONIZE, False, owned_pid)
            try:
                if win32event.WaitForSingleObject(handle, 15000) != win32event.WAIT_OBJECT_0:
                    raise RuntimeError(f"Smoke-owned Excel PID {owned_pid} did not exit after Quit().")
            finally:
                win32api.CloseHandle(handle)
            if _sha256(path) != digest:
                raise AssertionError("The disposable workbook changed during Protected View inspection.")
            print({
                "status": "PASS",
                "excel_version": str(excel_version),
                "excel_build": str(excel_build),
                "protected_view_caption": caption,
                "protected_view_ui_text": matching[:8],
                "workbook_sha256_unchanged": True,
                "editing_enabled": False,
                "owned_excel_pid": owned_pid,
                "owned_excel_exited": True,
            })
            return 0
    finally:
        if book is not None:
            try:
                book.Close(False)
            except Exception:
                pass
        if protected is not None:
            try:
                protected.Close()
            except Exception:
                pass
        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                pass
        excel = book = protected = None
        gc.collect()
        if owned_pid is not None:
            try:
                handle = win32api.OpenProcess(win32con.SYNCHRONIZE, False, owned_pid)
            except Exception:
                handle = None
            if handle is not None:
                try:
                    if win32event.WaitForSingleObject(handle, 15000) != win32event.WAIT_OBJECT_0:
                        raise RuntimeError(f"Smoke-owned Excel PID {owned_pid} did not exit.")
                finally:
                    win32api.CloseHandle(handle)
        pythoncom.CoUninitialize()


if __name__ == "__main__":
    raise SystemExit(main())
