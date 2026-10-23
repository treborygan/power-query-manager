"""Verify open Excel workbook detection when a read-only open has no owner-lock file."""
from __future__ import annotations

import hashlib
import argparse
import csv
import json
from pathlib import Path
import subprocess
import sys
import tempfile

import pythoncom
import win32api
import win32con
import win32event
import win32process
import win32com.client

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "app" / "src"))
from pq_installer.core.windows_file_users import active_file_users
from pq_installer.core.validation import validate_workbook_for_write
from pq_installer.core.models import ExistingQueryMode, QuerySource
from pq_installer.core.excel_queries import ExcelQueryInstaller
from pq_installer.core.excel_refresh import ExcelQueryRefresher
from pq_installer.core.history import HistoryStore


def hold_process(pid: int):
    return win32api.OpenProcess(win32con.SYNCHRONIZE, False, pid)


def excel_pids() -> set[int]:
    result = subprocess.run(["tasklist", "/FI", "IMAGENAME eq EXCEL.EXE", "/FO", "CSV", "/NH"],
                            check=True, capture_output=True, text=True)
    return {int(row[1]) for row in csv.reader(result.stdout.splitlines())
            if len(row) >= 2 and row[1].isdigit()}


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--cli-exe", type=Path,
                        help="also exercise the packaged CLI against the open read-only workbook")
    args = parser.parse_args()
    if args.cli_exe is not None and not args.cli_exe.is_file():
        parser.error(f"Packaged CLI does not exist: {args.cli_exe}")
    pythoncom.CoInitialize()
    app = book = None
    process_handle = None
    try:
        with tempfile.TemporaryDirectory(prefix="pqm-excel-open-no-lock-") as temp:
            root = Path(temp)
            workbook_path = root / "active.xlsx"
            lock_path = root / "~$active.xlsx"
            baseline_pids = excel_pids()
            app = win32com.client.DispatchEx("Excel.Application")
            app.Visible = False
            app.DisplayAlerts = False
            previous_security = app.AutomationSecurity
            app.AutomationSecurity = 3
            version, build = str(app.Version), str(app.Build)
            pid = int(win32process.GetWindowThreadProcessId(int(app.Hwnd))[1])
            if pid in baseline_pids:
                raise RuntimeError(f"DispatchEx reused pre-existing Excel process {pid}.")
            process_handle = hold_process(pid)
            book = app.Workbooks.Add()
            book.Worksheets.Item(1).Range("A1").Value = "disposable active workbook"
            book.SaveAs(str(workbook_path), FileFormat=51)
            book.Close(SaveChanges=False)
            book = None
            source_path = root / "query.m"
            source_path.write_text("let Value = 1 in Value", encoding="utf-8")
            history = HistoryStore(root / "history.sqlite3")
            snapshot = history.create_snapshot(workbook_path)
            book = app.Workbooks.Open(str(workbook_path), UpdateLinks=0, ReadOnly=True,
                                      IgnoreReadOnlyRecommended=True, AddToMru=False)
            if not bool(book.ReadOnly):
                raise RuntimeError("Excel did not honor the read-only open requested by this smoke.")
            if lock_path.exists():
                raise RuntimeError("Read-only Excel open created an owner-lock file; the no-lock case was not exercised.")

            users = active_file_users(workbook_path)
            if not users:
                raise AssertionError("Open Excel workbook was not detected without an owner-lock file.")
            try:
                validate_workbook_for_write(workbook_path)
            except ValueError as exc:
                if "currently open or in use" not in str(exc):
                    raise
            else:
                raise AssertionError("Write preflight accepted a workbook held open by Excel without its lock file.")
            expected_hash = hashlib.sha256(workbook_path.read_bytes()).hexdigest()

            cli_report = None
            if args.cli_exe is not None:
                profile_path = root / "profile.json"
                profile_path.write_text(json.dumps({
                    "schema_version": 1,
                    "profile_name": "Read-only open workbook refusal smoke",
                    "workbook": {"path": str(workbook_path)},
                    "options": {"existing_query_policy": "replace", "backup": True,
                                "dry_run": False, "refresh": "none"},
                    "queries": [{"id": "active-use-query", "source_path": str(source_path),
                                 "query_name": "PQM_ActiveUse", "enabled": True,
                                 "order": 0, "load": {"mode": "definition_only"}}],
                }, ensure_ascii=False, indent=2), encoding="utf-8")
                completed = subprocess.run(
                    [str(args.cli_exe.resolve()), "deploy", "--profile", str(profile_path),
                     "--workbook", str(workbook_path), "--apply", "--confirm", "DEPLOY"],
                    capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=120)
                if completed.returncode != 5:
                    raise RuntimeError(f"Expected packaged CLI refusal exit 5; got {completed.returncode}: "
                                       f"{completed.stderr}\n{completed.stdout}")
                try:
                    report_start = completed.stdout.index("{")
                    cli_report, _ = json.JSONDecoder().raw_decode(completed.stdout, report_start)
                except (ValueError, json.JSONDecodeError) as exc:
                    raise RuntimeError(f"Packaged CLI did not return a JSON deployment report: "
                                       f"{completed.stdout}\n{completed.stderr}") from exc
                results = cli_report.get("results", [])
                if len(results) != 1 or results[0].get("status") != "blocked_preflight":
                    raise AssertionError(f"Packaged CLI did not block the active workbook: {results!r}.")
                if "currently open or in use" not in results[0].get("message", "").casefold():
                    raise AssertionError(f"Packaged CLI refusal did not explain active workbook use: {results[0]!r}.")
                expected_pids = baseline_pids | {pid}
                observed_pids = excel_pids()
                if observed_pids != expected_pids:
                    raise AssertionError(f"Packaged CLI changed the Excel process set: "
                                         f"{sorted(observed_pids)} != {sorted(expected_pids)}")
                if hashlib.sha256(workbook_path.read_bytes()).hexdigest() != expected_hash:
                    raise AssertionError("Packaged CLI changed the read-only active workbook.")
                if list(root.glob("active_backup_*.xlsx")) or list(root.glob("*.pqstage-*")):
                    raise AssertionError("Packaged CLI created a backup or stage before active-use refusal.")

            def unexpected_excel_factory():
                raise AssertionError("Operation started Excel before active-workbook preflight refused the target.")

            operations = (
                ("import", lambda: ExcelQueryInstaller(excel_factory=unexpected_excel_factory).install(
                    workbook_path, [QuerySource(source_path, "PQM_ActiveUse")], ExistingQueryMode.REPLACE)),
                ("refresh", lambda: ExcelQueryRefresher(excel_factory=unexpected_excel_factory).refresh(
                    workbook_path)),
                ("restore", lambda: history.restore_snapshot(
                    snapshot["snapshot_id"], workbook_path, overwrite=True,
                    expected_current_sha256=expected_hash)),
            )
            refused = []
            for name, operation in operations:
                try:
                    operation()
                except Exception as exc:
                    if "currently open or in use" not in str(exc).casefold():
                        raise
                    refused.append(name)
                else:
                    raise AssertionError(f"{name} accepted a workbook held open by Excel without its lock file.")
            if refused != ["import", "refresh", "restore"]:
                raise AssertionError(f"Unexpected active-use operation results: {refused!r}.")
            if list(root.glob("*.pqstage-*")) or list(root.glob(".pqm-restore-*")):
                raise AssertionError("An operation created a stage before active-use refusal.")
            if list(root.glob("active_backup_*.xlsx")):
                raise AssertionError("An operation created a backup before active-use refusal.")
            book.Close(SaveChanges=False)
            book = None
            app.AutomationSecurity = previous_security
            app.Quit()
            app = None
            if win32event.WaitForSingleObject(process_handle, 10000) != win32event.WAIT_OBJECT_0:
                raise RuntimeError(f"Smoke-owned Excel process did not exit: {pid}")
            if hashlib.sha256(workbook_path.read_bytes()).hexdigest() != expected_hash:
                raise AssertionError("The disposable workbook changed during the read-only active-use check.")
            if lock_path.exists():
                raise AssertionError("Excel recreated an owner-lock file after the test closed the workbook.")
            cli_clause = (" packaged CLI deployment also returned blocked_preflight (exit 5) without "
                          "changing the workbook, creating a stage/backup, or starting Excel;"
                          if cli_report is not None else "")
            print(f"PASS Excel {version} build {build}; file-use check found {users!r} with no owner-lock file; "
                  "preflight blocked import, refresh, and overwrite-restore before stage/backup or Excel factory;" +
                  cli_clause + " workbook hash unchanged; owned Excel process exited.")
    finally:
        if book is not None:
            try:
                book.Close(SaveChanges=False)
            except Exception:
                pass
        if app is not None:
            try:
                app.Quit()
            except Exception:
                pass
        if process_handle is not None:
            win32api.CloseHandle(process_handle)
        pythoncom.CoUninitialize()


if __name__ == "__main__":
    main()
