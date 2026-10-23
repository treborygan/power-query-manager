import hashlib
import importlib.util
from pathlib import Path

import pytest

_VERIFIER_PATH = Path(__file__).resolve().parents[2] / "tools" / "verify_release_package.py"
_SPEC = importlib.util.spec_from_file_location("verify_release_package", _VERIFIER_PATH)
_VERIFIER = importlib.util.module_from_spec(_SPEC)
_SPEC.loader.exec_module(_VERIFIER)
verify_package = _VERIFIER.verify_package


def make_package(root):
    for relative in (
        "README.txt",
        "RELEASE_SUMMARY.txt",
        "PowerQueryManager/PowerQueryManager.exe",
        "PowerQueryManager/PowerQueryManagerCLI.exe",
        "PowerQueryManager/_internal/runtime.bin",
    ):
        path = root / relative
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_bytes(relative.encode("utf-8"))
    files = sorted(path for path in root.rglob("*") if path.is_file())
    manifest = "".join(
        f"{hashlib.sha256(path.read_bytes()).hexdigest()}  {path.relative_to(root).as_posix()}\n"
        for path in files
    )
    (root / "SHA256SUMS").write_text(manifest, encoding="utf-8")


def test_release_package_verifier_accepts_exact_manifest(tmp_path):
    make_package(tmp_path)
    count, _size = verify_package(tmp_path)
    assert count == 5


@pytest.mark.parametrize("change", ["tamper", "extra", "cache", "optional_qt", "optional_network"])
def test_release_package_verifier_rejects_tampering_and_dev_artifacts(tmp_path, change):
    make_package(tmp_path)
    if change == "tamper":
        (tmp_path / "README.txt").write_text("changed", encoding="utf-8")
        message = "SHA256 mismatch"
    elif change == "extra":
        (tmp_path / "unlisted.bin").write_bytes(b"extra")
        message = "Manifest coverage mismatch"
    elif change == "cache":
        (tmp_path / "PowerQueryManager" / "__pycache__").mkdir()
        message = "Development or transient directories"
    else:
        runtime_name = "Qt6Network.dll" if change == "optional_network" else "Qt6Qml.dll"
        optional = tmp_path / "PowerQueryManager" / "_internal" / "PySide6" / runtime_name
        optional.parent.mkdir(parents=True)
        optional.write_bytes(b"unused Qt module")
        files = sorted(path for path in tmp_path.rglob("*") if path.is_file() and path.name != "SHA256SUMS")
        manifest = "".join(
            f"{hashlib.sha256(path.read_bytes()).hexdigest()}  {path.relative_to(tmp_path).as_posix()}\n"
            for path in files
        )
        (tmp_path / "SHA256SUMS").write_text(manifest, encoding="utf-8")
        message = "(?i)unused Qt runtime components"
    with pytest.raises(ValueError, match=message):
        verify_package(tmp_path)


def test_release_package_verifier_rejects_path_traversal_in_manifest(tmp_path):
    make_package(tmp_path)
    manifest = tmp_path / "SHA256SUMS"
    manifest.write_text("0" * 64 + "  ../outside\n", encoding="utf-8")
    with pytest.raises(ValueError, match="Unsafe or duplicate manifest path"):
        verify_package(tmp_path)


def test_release_package_rejects_unused_qt_translation_catalogs(tmp_path):
    make_package(tmp_path)
    translation = tmp_path / "PowerQueryManager" / "_internal" / "PySide6" / "translations" / "qtbase_de.qm"
    translation.parent.mkdir(parents=True)
    translation.write_bytes(b"unused Qt catalog")
    files = sorted(path for path in tmp_path.rglob("*") if path.is_file() and path.name != "SHA256SUMS")
    manifest = "".join(
        f"{hashlib.sha256(path.read_bytes()).hexdigest()}  {path.relative_to(tmp_path).as_posix()}\n"
        for path in files
    )
    (tmp_path / "SHA256SUMS").write_text(manifest, encoding="utf-8")
    with pytest.raises(ValueError, match="(?i)unused Qt translation catalogs"):
        verify_package(tmp_path)
