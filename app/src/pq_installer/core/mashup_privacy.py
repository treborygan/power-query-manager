"""Read-only parsing of workbook-scoped Power Query privacy permissions."""
from __future__ import annotations

from dataclasses import dataclass
import base64
import hashlib
import struct
from pathlib import Path
import zipfile
from xml.etree import ElementTree


_DATA_MASHUP_NS = "http://schemas.microsoft.com/DataMashup"
_DPAPI_ENTROPY = b"DataExplorer Package Components"
_MAX_CUSTOM_XML_BYTES = 64 * 1024 * 1024
_MAX_PACKAGE_ENTRIES = 10000


@dataclass(frozen=True)
class WorkbookPrivacy:
    firewall_enabled: bool
    workbook_group_type: str | None


def _local_name(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def _permission_values(data: bytes) -> WorkbookPrivacy:
    try:
        root = ElementTree.fromstring(data)
    except ElementTree.ParseError as exc:
        raise ValueError("Power Query privacy permissions are malformed.") from exc
    root_namespace = root.tag[1:].split("}", 1)[0] if root.tag.startswith("{") else ""
    if _local_name(root.tag) != "PermissionList" or root_namespace not in {"", _DATA_MASHUP_NS}:
        raise ValueError("Power Query privacy permission stream has an unknown root element.")
    if any((child.tag[1:].split("}", 1)[0] if child.tag.startswith("{") else "") != root_namespace
           for child in root):
        raise ValueError("Power Query privacy permission stream has an unknown namespace.")
    names = [_local_name(child.tag) for child in root]
    if len(names) != len(set(names)):
        raise ValueError("Power Query privacy permission stream has duplicate fields.")
    values = dict(zip(names, root))
    firewall = values.get("FirewallEnabled")
    group = values.get("WorkbookGroupType")
    if firewall is None or (firewall.text or "").strip().casefold() not in {"true", "false", "1", "0"}:
        raise ValueError("Power Query FirewallEnabled permission is missing or invalid.")
    firewall_value = (firewall.text or "").strip().casefold() in {"true", "1"}
    if group is None:
        group_value = None
    elif group.attrib.get("{http://www.w3.org/2001/XMLSchema-instance}nil", "false").casefold() in {"true", "1"}:
        group_value = None
    else:
        group_value = (group.text or "").strip()
        if group_value not in {"None", "Public", "Organizational", "SeparatePrivate"}:
            raise ValueError("Power Query WorkbookGroupType permission is invalid.")
        if group_value == "None":
            group_value = None
    return WorkbookPrivacy(firewall_value, group_value)


def _expected_binding(package_parts: bytes, permissions: bytes) -> bytes:
    package_hash = hashlib.sha256(package_parts).digest()
    permissions_hash = hashlib.sha256(permissions).digest()
    return (struct.pack("<I", len(package_hash)) + package_hash
            + struct.pack("<I", len(permissions_hash)) + permissions_hash)


def _dpapi_unprotect(payload: bytes, entropy: bytes) -> bytes:
    try:
        import win32crypt
        _description, plaintext = win32crypt.CryptUnprotectData(payload, entropy)
        return bytes(plaintext)
    except Exception as exc:
        raise ValueError("Power Query privacy binding could not be verified for this Windows user.") from exc


def _parse_data_mashup(raw: bytes, unprotect) -> WorkbookPrivacy:
    if len(raw) < 20:
        raise ValueError("Power Query DataMashup stream is truncated.")
    version = struct.unpack_from("<I", raw, 0)[0]
    if version != 0:
        raise ValueError("Power Query DataMashup version is not supported for privacy verification.")
    offset = 4
    fields = []
    for _ in range(4):
        if offset + 4 > len(raw):
            raise ValueError("Power Query DataMashup stream is truncated.")
        size = struct.unpack_from("<I", raw, offset)[0]
        offset += 4
        end = offset + size
        if end > len(raw):
            raise ValueError("Power Query DataMashup field length is invalid.")
        fields.append(raw[offset:end])
        offset = end
    if offset != len(raw):
        raise ValueError("Power Query DataMashup stream has unexpected trailing bytes.")
    package_parts, permissions, _metadata, binding = fields
    if binding == b"\x00":
        # The format permits this unbound representation. Settings are still
        # compared before commit, but no integrity binding can be asserted.
        pass
    else:
        plaintext = unprotect(binding, _DPAPI_ENTROPY)
        if plaintext != _expected_binding(package_parts, permissions):
            raise ValueError("Power Query privacy permission binding does not match its package parts.")
    return _permission_values(permissions)


def read_workbook_privacy(path, *, dpapi_unprotect=None) -> WorkbookPrivacy | None:
    """Return workbook-level privacy flags, or None when no DataMashup exists.

    The parser reads only custom XML parts and never extracts package members.
    Invalid or unsupported privacy streams raise ValueError so callers can fail
    closed before replacing the original workbook.
    """
    path = Path(path)
    if not path.is_file() or not zipfile.is_zipfile(path):
        # Non-ZIP fake workbook adapters and invalid workbooks are left for the
        # normal Excel open path to reject; they carry no readable privacy part.
        return None
    unprotect = dpapi_unprotect or _dpapi_unprotect
    found = []
    try:
        with zipfile.ZipFile(path, "r") as archive:
            infos = archive.infolist()
            if len(infos) > _MAX_PACKAGE_ENTRIES:
                raise ValueError("Workbook package has too many parts for privacy verification.")
            for info in infos:
                name = info.filename.replace("\\", "/")
                if not name.casefold().startswith("customxml/") or not name.casefold().endswith(".xml"):
                    continue
                if info.file_size > _MAX_CUSTOM_XML_BYTES:
                    raise ValueError("Workbook custom XML part exceeds the privacy verification size limit.")
                data = archive.read(info)
                if b"<!DOCTYPE" in data.upper() or b"<!ENTITY" in data.upper():
                    raise ValueError("Workbook custom XML uses prohibited document declarations.")
                try:
                    root = ElementTree.fromstring(data)
                except ElementTree.ParseError:
                    if b"DataMashup" in data:
                        raise ValueError("Power Query DataMashup XML is malformed.")
                    continue
                for element in root.iter():
                    if (_local_name(element.tag) == "DataMashup"
                            and element.tag.startswith("{" + _DATA_MASHUP_NS + "}")):
                        text = (element.text or "").strip()
                        try:
                            raw = base64.b64decode(text, validate=True)
                        except Exception as exc:
                            raise ValueError("Power Query DataMashup content is not valid base64.") from exc
                        found.append(_parse_data_mashup(raw, unprotect))
    except (zipfile.BadZipFile, OSError) as exc:
        raise ValueError("Workbook package could not be inspected for Power Query privacy settings.") from exc
    if len(found) > 1:
        raise ValueError("Workbook contains ambiguous Power Query privacy permission streams.")
    return found[0] if found else None
