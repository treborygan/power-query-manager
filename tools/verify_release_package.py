"""Verify the staged Windows release package and its SHA256SUMS manifest."""
from __future__ import annotations

import hashlib
from pathlib import Path
import sys


FORBIDDEN_DIRS = {".git", ".venv", "__pycache__", "build", "tests", "fixtures", "logs"}
REQUIRED_FILES = {
    "README.txt",
    "RELEASE_SUMMARY.txt",
    "SHA256SUMS",
    "PowerQueryManager/PowerQueryManager.exe",
    "PowerQueryManager/PowerQueryManagerCLI.exe",
    "PowerQueryManager/_internal/shiboken6/Shiboken.pyd",
    "PowerQueryManager/_internal/shiboken6/shiboken6.abi3.dll",
}
UNUSED_QT_RUNTIME_FILES = {
    "qt6opengl.dll",
    "qt6pdf.dll",
    "qt6qml.dll",
    "qt6qmlmeta.dll",
    "qt6qmlmodels.dll",
    "qt6qmlworkerscript.dll",
    "qt6quick.dll",
    "qt6virtualkeyboard.dll",
    "qt6network.dll",
    "qtnetwork.pyd",
    "qnetworklistmanager.dll",
    "qopensslbackend.dll",
    "qschannelbackend.dll",
    "qcertonlybackend.dll",
    "libcrypto-3-x64.dll",
    "libssl-3-x64.dll",
    "opengl32sw.dll",
    "qpdf.dll",
    "qtvirtualkeyboardplugin.dll",
}


def verify_package(root: Path) -> tuple[int, int]:
    root = root.resolve()
    if not root.is_dir():
        raise ValueError(f"Release package directory does not exist: {root}")

    files = {path.relative_to(root).as_posix(): path for path in root.rglob("*") if path.is_file()}
    missing = sorted(REQUIRED_FILES - files.keys())
    if missing:
        raise ValueError(f"Required release files are missing: {', '.join(missing)}")

    forbidden = sorted(
        path.relative_to(root).as_posix()
        for path in root.rglob("*")
        if path.is_dir() and path.name.casefold() in FORBIDDEN_DIRS
    )
    if forbidden:
        raise ValueError(f"Development or transient directories found in package: {', '.join(forbidden)}")

    unused_qt_translations = sorted(
        name for name in files
        if name.casefold().startswith("powerquerymanager/_internal/pyside6/translations/")
    )
    if unused_qt_translations:
        raise ValueError(
            f"Unused Qt translation catalogs found in package ({len(unused_qt_translations)} files; "
            f"first: {unused_qt_translations[0]})"
        )
    unused_qt_runtime = sorted(
        name for name in files
        if name.casefold().startswith("powerquerymanager/_internal/")
        and Path(name).name.casefold() in UNUSED_QT_RUNTIME_FILES
    )
    if unused_qt_runtime:
        raise ValueError(
            f"Unused Qt runtime components found in package ({len(unused_qt_runtime)} files; "
            f"first: {unused_qt_runtime[0]})"
        )

    entries: dict[str, str] = {}
    for line_number, line in enumerate(files["SHA256SUMS"].read_text(encoding="utf-8").splitlines(), 1):
        if not line:
            continue
        parts = line.split("  ", 1)
        if len(parts) != 2 or len(parts[0]) != 64 or any(c not in "0123456789abcdef" for c in parts[0]):
            raise ValueError(f"Invalid SHA256SUMS entry at line {line_number}")
        name = parts[1]
        normalized = Path(name)
        if normalized.is_absolute() or ".." in normalized.parts or name in entries:
            raise ValueError(f"Unsafe or duplicate manifest path at line {line_number}")
        entries[name] = parts[0]

    payload = set(files) - {"SHA256SUMS"}
    if payload != set(entries):
        missing_manifest = sorted(payload - set(entries))
        missing_file = sorted(set(entries) - payload)
        raise ValueError(f"Manifest coverage mismatch; unlisted={missing_manifest}, absent={missing_file}")

    mismatches = [
        name for name, digest in entries.items()
        if hashlib.sha256(files[name].read_bytes()).hexdigest() != digest
    ]
    if mismatches:
        raise ValueError(f"SHA256 mismatch: {', '.join(mismatches)}")
    return len(entries), sum(files[name].stat().st_size for name in payload)


def main() -> int:
    root = Path(sys.argv[1]) if len(sys.argv) > 1 else Path(__file__).resolve().parents[1] / "dist"
    try:
        count, size = verify_package(root)
    except (OSError, ValueError) as exc:
        print(f"FAIL: {exc}", file=sys.stderr)
        return 1
    print(f"PASS: {count} files verified; package payload {size:,} bytes.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
