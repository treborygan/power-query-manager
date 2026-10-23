from __future__ import annotations

import hashlib
import subprocess
import sys
import tempfile
from pathlib import Path

import pytest


ROOT = Path(__file__).resolve().parents[2]


def test_package_stager_can_build_and_verify_an_isolated_release() -> None:
    with tempfile.TemporaryDirectory(dir=ROOT, prefix=".package-test-") as temporary:
        scratch = Path(temporary)
        source = scratch / "source" / "PowerQueryManager"
        destination = scratch / "release"
        (source / "_internal").mkdir(parents=True)
        (source / "PowerQueryManager.exe").write_bytes(b"fake-gui-launcher")
        (source / "PowerQueryManagerCLI.exe").write_bytes(b"fake-cli-launcher")
        (source / "_internal" / "runtime.bin").write_bytes(b"runtime")

        result = subprocess.run(
            [
                sys.executable,
                str(ROOT / "tools" / "package_release.py"),
                "--source",
                str(source),
                "--destination",
                str(destination),
            ],
            cwd=ROOT,
            capture_output=True,
            text=True,
            check=False,
        )

        assert result.returncode == 0, result.stdout + result.stderr
        manual = (destination / "README.txt").read_text(encoding="utf-8")
        summary = (destination / "RELEASE_SUMMARY.txt").read_text(encoding="utf-8")
        assert "stage filename" in manual
        assert "symbolic link" in manual
        assert "W47" in manual
        assert "W49" in manual
        assert "W52" in manual
        assert "W53" in manual
        assert "W54" in manual
        assert "W55" in manual
        assert "W229" in manual
        assert "W230" in manual
        assert "W231" in manual
        assert "W232" in manual
        assert "W234" in manual
        assert "W235" in manual
        assert "W236" in manual
        assert "W237" in manual
        assert "W238" in manual
        assert "W239" in manual
        assert "W240" in manual
        assert "W241" in manual
        assert "W242" in manual
        assert "W243" in manual
        assert "W244" in manual
        assert "Tab moves focus through the M editors" in manual
        assert "W245" in manual
        assert "definition-only import passed on Excel" in manual
        assert "W246" in manual
        assert "existing-query replacement passed on Excel" in manual
        assert "W247" in manual
        assert "227 runtime files" in manual
        assert "W248" in manual
        assert "all four packaged GUI load modes passed" in manual
        assert "461-test Windows build" in manual
        assert "449-test Windows build" in manual
        assert "W46" in summary
        assert "W47" in summary
        assert "W49" in summary
        assert "W52" in summary
        assert "W53" in summary
        assert "W54" in summary
        assert "W55" in summary
        assert "W229" in summary
        assert "W230" in summary
        assert "W231" in summary
        assert "data validation, conditional formatting, hyperlink" in summary
        assert "W234" in summary
        assert "W235" in summary
        assert "W236" in summary
        assert "W237" in summary
        assert "W238" in summary
        assert "W239" in summary
        assert "W240" in summary
        assert "byte-identical" in summary
        assert "W241" in summary
        assert "keyboard traversal" in summary
        assert "W242" in summary
        assert "operation-history dialog" in summary
        assert "W243" in summary
        assert "dependency-graph dialog" in summary
        assert "W244" in summary
        assert "query-inspector keyboard traversal" in summary
        assert "W245" in summary
        assert "Excel 16.0 build 20430.0" in summary
        assert "W246" in summary
        assert "existing-query replacement" in summary
        assert "W247" in summary
        assert "zero missing, extra, or changed" in summary
        assert "W248" in summary
        assert "Data Model membership and row count 2" in summary
        assert "W249" in manual
        assert "background and synchronous refresh passed on Excel" in manual
        assert "W249" in summary
        assert "no_external_source" in summary
        assert "W250" in manual
        assert "worksheet-loaded Power Query" in manual
        assert "W250" in summary
        assert "seven loopback HTTP requests" in summary
        assert "W251" in manual
        assert "3-second refresh timeout returned exit 5" in manual
        assert "W251" in summary
        assert "Original workbook hash was preserved" in summary
        assert "463 tests" in summary
        assert "134 feature-traceability references" in summary
        assert "Excel 16.0 build 20430.0" in summary
        assert "Excel 16.0 build 20430.0" in summary
        assert (destination / "PowerQueryManager" / "_internal" / "runtime.bin").read_bytes() == b"runtime"
        assert (destination / "SHA256SUMS").is_file()

        gui_exe = destination / "PowerQueryManager" / "PowerQueryManager.exe"
        original_hash = hashlib.sha256(gui_exe.read_bytes()).hexdigest()
        docs_only = subprocess.run(
            [
                sys.executable,
                str(ROOT / "tools" / "package_release.py"),
                "--destination",
                str(destination),
                "--docs-only",
            ],
            cwd=ROOT,
            capture_output=True,
            text=True,
            check=False,
        )
        assert docs_only.returncode == 0, docs_only.stdout + docs_only.stderr
        assert hashlib.sha256(gui_exe.read_bytes()).hexdigest() == original_hash

        replacement = subprocess.run(
            [
                sys.executable,
                str(ROOT / "tools" / "package_release.py"),
                "--source",
                str(source),
                "--destination",
                str(destination),
            ],
            cwd=ROOT,
            capture_output=True,
            text=True,
            check=False,
        )
        assert replacement.returncode == 0, replacement.stdout + replacement.stderr
        assert (destination / "PowerQueryManager" / "PowerQueryManagerCLI.exe").is_file()

        sentinel = destination / "keep-user-file.txt"
        sentinel.write_text("must remain", encoding="utf-8")
        refused = subprocess.run(
            [
                sys.executable,
                str(ROOT / "tools" / "package_release.py"),
                "--source",
                str(source),
                "--destination",
                str(destination),
            ],
            cwd=ROOT,
            capture_output=True,
            text=True,
            check=False,
        )
        assert refused.returncode != 0
        assert sentinel.read_text(encoding="utf-8") == "must remain"


def test_release_directory_replace_retries_transient_windows_lock_errors(tmp_path):
    import sys
    sys.path.insert(0, str(ROOT / 'tools'))
    from package_fs import replace_directory

    calls=[];delays=[]
    def replace(source,destination):
        calls.append((source,destination))
        if len(calls)==1:
            error=PermissionError('directory temporarily locked')
            error.winerror=5
            raise error
        return 'published'

    result=replace_directory(tmp_path/'stage',tmp_path/'release',replace=replace,
                             sleep=delays.append,windows=True)
    assert result=='published'
    assert len(calls)==2
    assert delays==[0.05]


def test_release_directory_replace_does_not_retry_unrelated_windows_error(tmp_path):
    import sys
    sys.path.insert(0, str(ROOT / 'tools'))
    from package_fs import replace_directory

    calls=[];delays=[]
    def replace(source,destination):
        calls.append((source,destination))
        error=OSError('invalid name')
        error.winerror=123
        raise error

    with pytest.raises(OSError,match='invalid name'):
        replace_directory(tmp_path/'stage',tmp_path/'release',replace=replace,
                          sleep=delays.append,windows=True)
    assert len(calls)==1 and delays==[]
