"""Filesystem-first, read-only incoming dependency discovery for Excel/Access.

Source definitions are examined offline. Never refresh M, evaluate formulas,
execute Access queries, connect to external sources, or edit the input files.
"""
from __future__ import annotations

import base64
from bisect import bisect_right
from dataclasses import dataclass
from io import BytesIO
import ntpath
from pathlib import Path
import re
import struct
from urllib.parse import unquote, urlparse
from xml.etree import ElementTree
import zipfile

from pq_installer.core.m_source_mapping import _tokens


SUPPORTED = frozenset({".xlsx", ".xlsm", ".accdb", ".mdb"})
KNOWN_UNSUPPORTED = frozenset({".xls", ".xlsb"})
_MAX_ENTRIES = 10000
_MAX_XML_SIZE = 64 * 1024 * 1024
_MAX_MASHUP_SIZE = 128 * 1024 * 1024
_MAX_TOTAL_METADATA_BYTES = 256 * 1024 * 1024
_SOURCE_FUNCTIONS = frozenset({
    ("file", "contents"), ("folder", "files"), ("folder", "contents"),
    ("odbc", "datasource"), ("odbc", "query"), ("oledb", "datasource"),
})
_DATABASE_KEYS = frozenset({"data source", "datasource", "database", "dbq", "source db", "sourcefile"})
_QUOTED_PATHS = re.compile(r"(?i)(?:[a-z]:[\\/]|\\\\|/)[^\s;'\"<>,)]+")


@dataclass(frozen=True)
class DependencyReference:
    source_kind: str
    source_name: str
    raw: str
    direct: bool = False


@dataclass(frozen=True)
class DependencyHit:
    target_path: str
    consumer_path: str
    source_kind: str
    source_name: str
    confidence: str  # exact_path | filename_only | possible_literal
    evidence: str  # intentionally never includes raw credentials/SQL/M


@dataclass(frozen=True)
class ScanIssue:
    path: str
    status: str  # unreadable | unsupported | engine_unavailable | dynamic_reference
    detail: str


@dataclass(frozen=True)
class DependencyReport:
    targets: tuple[str, ...]
    inspected: tuple[str, ...]
    hits: tuple[DependencyHit, ...]
    issues: tuple[ScanIssue, ...]
    cancelled: bool = False


def discover_scan_files(inputs, *, recursive=True, cancelled=None):
    """Inventory supported and known-unsupported Office files; skip lock files."""
    discovered = {}
    for input_path in inputs:
        if cancelled is not None and cancelled():
            break
        root = Path(input_path).expanduser()
        if root.is_file():
            candidates = (root,)
        elif root.is_dir():
            candidates = root.rglob("*") if recursive else root.iterdir()
        else:
            continue
        for path in candidates:
            if cancelled is not None and cancelled():
                break
            if (path.suffix.casefold() not in SUPPORTED | KNOWN_UNSUPPORTED
                    or path.name.startswith("~$") or not path.is_file()):
                continue
            full = path.resolve()
            discovered[str(full).casefold()] = full
    return tuple(sorted(discovered.values(), key=lambda path: str(path).casefold()))


def _local_name(tag):
    return tag.rsplit("}", 1)[-1]


def _parse_xml(data):
    if b"<!DOCTYPE" in data.upper() or b"<!ENTITY" in data.upper():
        raise ValueError("Prohibited XML declaration in workbook.")
    return ElementTree.fromstring(data)


def _mashup_scripts(data):
    root = _parse_xml(data)
    for node in root.iter():
        if (_local_name(node.tag) != "DataMashup"
                or not node.tag.startswith("{http://schemas.microsoft.com/DataMashup}")):
            continue
        payload = base64.b64decode((node.text or "").strip(), validate=True)
        if len(payload) > _MAX_MASHUP_SIZE or len(payload) < 20:
            raise ValueError("Power Query DataMashup is too large or truncated.")
        if struct.unpack_from("<I", payload, 0)[0] != 0:
            raise ValueError("Unsupported Power Query DataMashup format.")
        offset = 4
        fields = []
        for _ in range(4):
            if offset + 4 > len(payload):
                raise ValueError("Truncated Power Query DataMashup.")
            size = struct.unpack_from("<I", payload, offset)[0]
            offset += 4
            if size > _MAX_MASHUP_SIZE or offset + size > len(payload):
                raise ValueError("Invalid Power Query DataMashup length.")
            fields.append(payload[offset:offset + size])
            offset += size
        if offset != len(payload):
            raise ValueError("Unexpected Power Query DataMashup bytes.")
        # The first field is an OPC ZIP package. No DPAPI/private binding is
        # required to inspect the Section1.m definition (unlike modification).
        with zipfile.ZipFile(BytesIO(fields[0])) as package:
            infos = package.infolist()
            if len(infos) > _MAX_ENTRIES:
                raise ValueError("Too many Power Query package parts.")
            for info in infos:
                if (info.filename.replace("\\", "/").casefold().startswith("formulas/")
                        and info.filename.casefold().endswith(".m")):
                    if info.file_size > _MAX_XML_SIZE:
                        raise ValueError("Power Query definition exceeds scan limit.")
                    yield (info.filename, package.read(info).decode("utf-8-sig"))


def _m_references(text):
    refs = []
    dynamic = False
    tokens = _tokens(text)
    declarations = []
    for index in range(len(tokens) - 2):
        keyword, name, equals = tokens[index:index + 3]
        if (keyword.kind == "identifier" and keyword.value.casefold() == "shared"
                and name.kind in {"identifier", "quoted_identifier"} and equals.value == "="):
            display = name.value if name.kind == "identifier" else text[name.start + 2:name.end - 1].replace('""', '"')
            declarations.append((keyword.start, display))
    positions = [entry[0] for entry in declarations]

    def source_name(start):
        position = bisect_right(positions, start) - 1
        return declarations[position][1] if position >= 0 else "Power Query definition"

    literal_indices = set()
    for index in range(len(tokens) - 3):
        a, dot, b, opening = tokens[index:index + 4]
        if (a.kind != "identifier" or dot.value != "." or b.kind != "identifier"
                or opening.value != "(" or (a.value.casefold(), b.value.casefold()) not in _SOURCE_FUNCTIONS):
            continue
        argument_index = index + 4
        if argument_index < len(tokens) and tokens[argument_index].kind == "string":
            arg = tokens[argument_index]
            literal_indices.add(argument_index)
            refs.append(DependencyReference("Power Query", source_name(a.start), arg.value, True))
        else:
            dynamic = True
    for index, token in enumerate(tokens):
        if (token.kind == "string" and index not in literal_indices
                and re.search(r"(?i)\.(?:xlsx|xlsm|accdb|mdb)\b", token.value)):
            refs.append(DependencyReference("Power Query literal", source_name(token.start), token.value))
    return refs, dynamic


def _excel_references(path):
    refs = []
    issues = []
    try:
        with zipfile.ZipFile(path) as workbook:
            if len(workbook.infolist()) > _MAX_ENTRIES:
                raise ValueError("Workbook contains too many package parts.")
            metadata_bytes_read = 0
            for info in workbook.infolist():
                name = info.filename.replace("\\", "/")
                lowered = name.casefold()
                if not (
                    lowered == "xl/connections.xml"
                    or (lowered.startswith("xl/externallinks/_rels/") and lowered.endswith(".rels"))
                    or lowered == "xl/workbook.xml"
                    or (lowered.startswith("xl/worksheets/") and lowered.endswith(".xml"))
                    or (lowered.startswith("customxml/") and lowered.endswith(".xml"))
                ):
                    continue
                if info.file_size > _MAX_XML_SIZE:
                    issues.append(ScanIssue(str(path), "unreadable", "Workbook metadata part exceeds scan size limit."))
                    continue
                if metadata_bytes_read + info.file_size > _MAX_TOTAL_METADATA_BYTES:
                    issues.append(ScanIssue(str(path), "unreadable", "Workbook metadata scan budget exceeded; remaining parts were not inspected."))
                    break
                metadata_bytes_read += info.file_size
                data = workbook.read(info)
                if lowered.startswith("customxml/"):
                    if b"DataMashup" not in data:
                        continue
                    try:
                        for script_name, text in _mashup_scripts(data):
                            found, dynamic = _m_references(text)
                            refs.extend(found)
                            if dynamic:
                                issues.append(ScanIssue(str(path), "dynamic_reference", "Power Query has a dynamic source expression that cannot be resolved statically."))
                    except (ValueError, zipfile.BadZipFile, UnicodeError, ElementTree.ParseError) as exc:
                        issues.append(ScanIssue(str(path), "unreadable", "An embedded Power Query definition could not be decoded."))
                    continue
                root = _parse_xml(data)
                if lowered == "xl/connections.xml":
                    for connection in root.iter():
                        if _local_name(connection.tag) != "connection":
                            continue
                        connection_name = connection.attrib.get("name", "Excel connection")
                        for part in connection.iter():
                            for key in ("connection", "sourceFile", "url"):
                                if key in part.attrib:
                                    refs.append(DependencyReference("Excel connection", connection_name, part.attrib[key], True))
                elif lowered.startswith("xl/externallinks/"):
                    for relation in root.iter():
                        if _local_name(relation.tag) == "Relationship" and relation.attrib.get("TargetMode", "").casefold() == "external":
                            refs.append(DependencyReference("Excel external link", name, relation.attrib.get("Target", ""), True))
                elif lowered == "xl/workbook.xml":
                    for item in root.iter():
                        if _local_name(item.tag) == "definedName" and item.text:
                            refs.append(DependencyReference("Excel defined name", item.attrib.get("name", "Defined name"), item.text))
                else:
                    for formula in root.iter():
                        if _local_name(formula.tag) == "f" and formula.text:
                            refs.append(DependencyReference("Excel formula", name, formula.text))
    except (zipfile.BadZipFile, OSError, ValueError, ElementTree.ParseError) as exc:
        issues.append(ScanIssue(str(path), "unreadable", "Workbook ZIP/XML metadata could not be inspected."))
    return tuple(refs), tuple(issues)


def _normalize_path(value):
    value = value.strip().strip('"\'')
    if value.casefold().startswith("file:"):
        parsed = urlparse(value)
        if parsed.scheme.casefold() != "file":
            return ""
        value = ("//" + parsed.netloc if parsed.netloc else "") + unquote(parsed.path)
        if re.match(r"^/[A-Za-z]:/", value):
            value = value[1:]
    return ntpath.normcase(ntpath.normpath(value)) if value else ""


def _possible_paths(text):
    """Find file paths in connection strings, SQL and M literals, not SQL data."""
    value = text.strip()
    yield value
    for item in value.split(";"):
        if "=" in item:
            key, _, candidate = item.partition("=")
            if key.strip().casefold() in _DATABASE_KEYS:
                yield candidate.strip()
    for match in _QUOTED_PATHS.finditer(value):
        yield match.group(0)
    for match in re.finditer(r"['\"]([^'\"]+\.(?:xlsx|xlsm|accdb|mdb))['\"]", value, re.I):
        yield match.group(1)


def _match_reference(reference, target):
    target_norm = _normalize_path(str(target))
    name = target.name.casefold()
    for candidate in _possible_paths(reference.raw):
        normalized = _normalize_path(candidate)
        if normalized == target_norm:
            return "exact_path" if reference.direct else "possible_literal"
        # Relative paths containing directories may be resolvable against the
        # consumer folder, but not reliably for all Office connection types.
    if re.search(r"(?i)(?<![\w.\-])" + re.escape(name) + r"(?![\w.\-])", reference.raw):
        return "filename_only" if reference.direct else "possible_literal"
    return None


def scan_reverse_dependencies(targets, consumers, *, access_reader=None, progress=None, cancelled=None):
    """Find files which reference selected targets, with explicit coverage gaps.

    Never open the target for analysis. On Windows Access metadata is read
    from a disposable database copy by DAO; Excel files are ZIP/XML only.
    """
    from pq_installer.core.access_dependency_reader import AccessEngineUnavailable, read_access_references

    chosen = tuple(dict.fromkeys(Path(path).expanduser().resolve() for path in targets))
    if not chosen:
        raise ValueError("Select at least one target file.")
    if any(not path.is_file() for path in chosen):
        raise ValueError("Every dependency target must be an existing file.")
    requested = tuple(dict.fromkeys(Path(path).expanduser().resolve() for path in consumers))
    target_keys = {str(path).casefold() for path in chosen}
    candidates = [p for p in requested if str(p).casefold() not in target_keys]
    inspected = []
    hits = []
    issues = []
    if not candidates:
        issues.append(ScanIssue("", "no_consumers", "No candidate files remain after excluding the target files; nothing was inspected."))
    reader = access_reader or read_access_references
    is_cancelled = False
    for index, path in enumerate(candidates, start=1):
        if cancelled is not None and cancelled():
            is_cancelled = True
            break
        inspected.append(str(path))
        suffix = path.suffix.casefold()
        references = ()
        try:
            if suffix in {".xlsx", ".xlsm"}:
                references, problems = _excel_references(path)
                issues.extend(problems)
            elif suffix in {".accdb", ".mdb"}:
                references = reader(path)
            else:
                issues.append(ScanIssue(str(path), "unsupported", "This file type is not supported by the read-only scanner."))
        except AccessEngineUnavailable:
            issues.append(ScanIssue(str(path), "engine_unavailable", "Microsoft Access DAO engine is unavailable; install a compatible Access Database Engine."))
        except Exception:
            issues.append(ScanIssue(str(path), "unreadable", "This source could not be inspected; it may be locked, encrypted, damaged, or inaccessible."))
        for reference in references:
            for target in chosen:
                match = _match_reference(reference, target)
                if match:
                    hits.append(DependencyHit(str(target), str(path), reference.source_kind,
                        reference.source_name, match, target.name))
        if progress:
            progress(index, len(candidates), str(path))
    # Exact-path evidence takes priority over basename-only evidence for the
    # same query/connection and target; repeated metadata parts are deduped.
    chosen_hits = {}
    for row in hits:
        key = (row.target_path, row.consumer_path, row.source_kind, row.source_name)
        if key not in chosen_hits or (chosen_hits[key].confidence != "exact_path" and row.confidence == "exact_path"):
            chosen_hits[key] = row
    return DependencyReport(
        tuple(map(str, chosen)), tuple(inspected), tuple(chosen_hits.values()), tuple(issues), is_cancelled
    )
