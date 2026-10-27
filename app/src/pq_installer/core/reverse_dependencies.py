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


@dataclass(frozen=True)
class UnresolvedReference:
    consumer_path: str
    source_kind: str
    source_name: str
    reference_filename: str  # safe basename only; never raw SQL/connection text


@dataclass(frozen=True)
class DependencyMapReport:
    files: tuple[str, ...]
    inspected: tuple[str, ...]
    edges: tuple[DependencyHit, ...]
    unresolved: tuple[UnresolvedReference, ...]
    issues: tuple[ScanIssue, ...]
    depth: str
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


class _GuardedXmlStream:
    """Bounded incremental reader that rejects entity/DTD declarations.

    ElementTree does not need a whole-sheet DOM to find `<f>` elements. The
    32-byte look-behind detects declarations split across ZIP read chunks.
    """

    def __init__(self, stream, maximum):
        self._stream = stream
        self._maximum = maximum
        self._read = 0
        self._previous = b""

    def read(self, count=-1):
        chunk = self._stream.read(min(65536, count) if count is not None and count > 0 else 65536)
        self._read += len(chunk)
        if self._read > self._maximum:
            raise ValueError("XML metadata exceeds the streaming size budget.")
        probe = self._previous + chunk
        if (b"<!DOCTYPE" in probe.upper() or b"<!ENTITY" in probe.upper()
                or b"\x00" in probe):
            raise ValueError("Unsupported XML declaration or encoding.")
        self._previous = probe[-32:]
        return chunk


def _stream_xml_references(workbook, info, path, kind):
    """Stream Excel connections/links/defined names/formulas with subtree cleanup."""
    connection_name = "Excel connection"
    parents = []
    with workbook.open(info) as source:
        stream = _GuardedXmlStream(source, _MAX_TOTAL_METADATA_BYTES)
        for event, element in ElementTree.iterparse(stream, events=("start", "end")):
            local = _local_name(element.tag)
            if event == "start":
                parents.append(element)
                if kind == "connections":
                    if local == "connection":
                        connection_name = element.attrib.get("name", "Excel connection")
                    for key in ("connection", "sourceFile", "url"):
                        if key in element.attrib:
                            yield DependencyReference(
                                "Excel connection", connection_name, element.attrib[key], True
                            )
                elif kind == "links" and local == "Relationship":
                    if element.attrib.get("TargetMode", "").casefold() == "external":
                        yield DependencyReference(
                            "Excel external link", info.filename, element.attrib.get("Target", ""), True
                        )
            else:
                if kind == "names" and local == "definedName" and element.text:
                    yield DependencyReference(
                        "Excel defined name", element.attrib.get("name", "Defined name"), element.text
                    )
                elif kind == "formulas" and local == "f" and element.text:
                    yield DependencyReference("Excel formula", info.filename, element.text)
                parents.pop()
                if parents:
                    # Clearing alone leaves millions of empty `row` objects
                    # attached to sheetData; unlink completed nodes as well.
                    parents[-1].remove(element)
                element.clear()


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


def _excel_references(path, *, depth="deep", on_reference=None):
    if depth not in {"quick", "deep"}:
        raise ValueError("Scan depth must be quick or deep.")
    refs = []
    emit = on_reference or refs.append
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
                    or (depth == "deep" and lowered == "xl/workbook.xml")
                    or (depth == "deep" and lowered.startswith("xl/worksheets/") and lowered.endswith(".xml"))
                    or (lowered.startswith("customxml/") and lowered.endswith(".xml"))
                ):
                    continue
                if lowered.startswith("customxml/") and info.file_size > _MAX_XML_SIZE:
                    issues.append(ScanIssue(str(path), "unreadable", "Workbook metadata part exceeds scan size limit."))
                    continue
                if metadata_bytes_read + info.file_size > _MAX_TOTAL_METADATA_BYTES:
                    issues.append(ScanIssue(str(path), "unreadable", "Workbook metadata scan budget exceeded; remaining parts were not inspected."))
                    break
                metadata_bytes_read += info.file_size
                if lowered.startswith("customxml/"):
                    data = workbook.read(info)
                    # Excel writes customXml/item*.xml as UTF-16LE on current
                    # Office builds. Search both the UTF-8-compatible ASCII
                    # spelling and UTF-16 before parsing the package part.
                    if (b"DataMashup" not in data
                            and "DataMashup".encode("utf-16le") not in data
                            and "DataMashup".encode("utf-16be") not in data):
                        continue
                    try:
                        for script_name, text in _mashup_scripts(data):
                            found, dynamic = _m_references(text)
                            for reference in found:
                                emit(reference)
                            if dynamic:
                                issues.append(ScanIssue(str(path), "dynamic_reference", "Power Query has a dynamic source expression that cannot be resolved statically."))
                    except (ValueError, zipfile.BadZipFile, UnicodeError, ElementTree.ParseError) as exc:
                        issues.append(ScanIssue(str(path), "unreadable", "An embedded Power Query definition could not be decoded."))
                    continue
                kind = ("connections" if lowered == "xl/connections.xml" else
                        "links" if lowered.startswith("xl/externallinks/") else
                        "names" if lowered == "xl/workbook.xml" else "formulas")
                try:
                    for reference in _stream_xml_references(workbook, info, path, kind):
                        emit(reference)
                except (OSError, ValueError, ElementTree.ParseError, zipfile.BadZipFile):
                    issues.append(ScanIssue(str(path), "unreadable", "A workbook XML metadata part could not be inspected safely."))
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


def _explicit_reference_paths(reference):
    """Normalized candidate paths, independent of direct/literal confidence."""
    return {_normalize_path(candidate) for candidate in _possible_paths(reference.raw)}


def scan_reverse_dependencies(targets, consumers, *, access_reader=None, progress=None, cancelled=None,
                              depth="deep"):
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
    hits = {}
    issues = []
    ambiguous_sources = set()
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
        def collect_reference(reference):
            exact_paths = _explicit_reference_paths(reference)
            matches_by_name = {}
            for target in chosen:
                match = _match_reference(reference, target)
                if match:
                    matches_by_name.setdefault(target.name.casefold(), []).append((target, match))
            for filename, possible in matches_by_name.items():
                exact = [(target, match) for target, match in possible
                         if _normalize_path(str(target)) in exact_paths]
                if exact:
                    possible = exact
                elif len(possible) > 1:
                    key = (str(path), reference.source_kind, reference.source_name, filename)
                    if key not in ambiguous_sources:
                        ambiguous_sources.add(key)
                        issues.append(ScanIssue(str(path), "ambiguous_reference",
                            "Multiple targets share a referenced basename; no dependency was guessed."))
                    continue
                for target, match in possible:
                    key = (str(target), str(path), reference.source_kind, reference.source_name)
                    if key not in hits or (hits[key].confidence != "exact_path" and match == "exact_path"):
                        hits[key] = DependencyHit(str(target), str(path), reference.source_kind,
                            reference.source_name, match, target.name)
        try:
            if suffix in {".xlsx", ".xlsm"}:
                references, problems = _excel_references(path, depth=depth, on_reference=collect_reference)
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
            collect_reference(reference)
        if progress:
            progress(index, len(candidates), str(path))
    return DependencyReport(
        tuple(map(str, chosen)), tuple(inspected), tuple(hits.values()), tuple(issues), is_cancelled
    )


def _safe_reference_filename(candidate):
    """Accept a single filename/path; do not allow SQL/M or secret attributes."""
    if not candidate:
        return None
    clean = candidate.strip().strip('"\'')
    if clean.casefold().startswith("file:"):
        parsed = urlparse(clean)
        if parsed.scheme.casefold() != "file":
            return None
        clean = ("//" + parsed.netloc if parsed.netloc else "") + unquote(parsed.path)
    filename = ntpath.basename(clean.replace("/", "\\"))
    # A SQL statement can end with an Office-looking token. Its complete
    # text is not a filename and must never appear in the unresolved index.
    if re.search(r"(?i)\b(?:select\s+.+\s+from|delete\s+.+\s+from|union\s+select|insert\s+into|update\s+\S+\s+set)\b",
                 filename):
        return None
    if re.fullmatch(r"(?i)[^<>|?*\"'=;\[\](){}\r\n\t]+\.(?:xlsx|xlsm|accdb|mdb|xlsb|xls)", filename):
        return filename
    return None


def _reference_filenames(raw):
    """Extract every referenced Office basename, never arbitrary raw SQL/M."""
    seen = set()
    candidates = list(_possible_paths(raw))
    candidates.extend(match.group(1) for match in re.finditer(
        r"\[([^\[\]]+\.(?:xlsx|xlsm|accdb|mdb|xlsb|xls))\]", raw, re.I))
    for candidate in candidates:
        filename = _safe_reference_filename(candidate)
        if filename and filename.casefold() not in seen:
            seen.add(filename.casefold())
            yield filename


def scan_dependency_map(consumers, *, depth="quick", access_reader=None, progress=None, cancelled=None):
    """Inspect each selected file once and index incoming/outgoing links.

    The returned immutable report contains only matched edges and safe file
    basenames for unmatched sources; raw M/SQL/credentials are discarded.
    """
    from pq_installer.core.access_dependency_reader import AccessEngineUnavailable, read_access_references

    if depth not in {"quick", "deep"}:
        raise ValueError("Scan depth must be quick or deep.")
    chosen = tuple(dict.fromkeys(Path(value).expanduser().resolve() for value in consumers))
    by_name = {}
    for item in chosen:
        by_name.setdefault(item.name.casefold(), []).append(item)
    read_access = access_reader or read_access_references
    inspected = []
    edges = {}
    unresolved = {}
    issues = []
    ambiguous_sources = set()
    priorities = {"exact_path": 3, "filename_only": 2, "possible_literal": 1}
    was_cancelled = False
    if not chosen:
        issues.append(ScanIssue("", "no_consumers", "No files were selected for dependency mapping."))
    for number, source in enumerate(chosen, start=1):
        if cancelled and cancelled():
            was_cancelled = True
            break
        inspected.append(str(source))
        references = ()
        suffix = source.suffix.casefold()
        def collect_reference(reference):
            exact_paths = _explicit_reference_paths(reference)
            for filename in _reference_filenames(reference.raw):
                matches = [(target, _match_reference(reference, target))
                           for target in by_name.get(filename.casefold(), ())]
                matches = [(target, confidence) for target, confidence in matches if confidence]
                exact = [(target, confidence) for target, confidence in matches
                         if _normalize_path(str(target)) in exact_paths]
                if exact:
                    matches = exact
                if len(matches) == 1:
                    target, confidence = matches[0]
                    key = (str(target), str(source), reference.source_kind, reference.source_name)
                    if key not in edges or priorities[confidence] > priorities[edges[key].confidence]:
                        edges[key] = DependencyHit(str(target), str(source), reference.source_kind,
                                                   reference.source_name, confidence, target.name)
                else:
                    reference_key = (str(source), reference.source_kind, reference.source_name, filename)
                    if reference_key not in unresolved:
                        unresolved[reference_key] = UnresolvedReference(*reference_key)
                    if len(matches) > 1 and reference_key not in ambiguous_sources:
                        ambiguous_sources.add(reference_key)
                        issues.append(ScanIssue(str(source), "ambiguous_reference",
                                                 "Multiple files share a referenced basename; no edge was guessed."))
        try:
            if suffix in {".xlsx", ".xlsm"}:
                references, problems = _excel_references(source, depth=depth, on_reference=collect_reference)
                issues.extend(problems)
            elif suffix in {".accdb", ".mdb"}:
                references = read_access(source)
            else:
                issues.append(ScanIssue(str(source), "unsupported", "This file type is not supported by the read-only scanner."))
        except AccessEngineUnavailable:
            issues.append(ScanIssue(str(source), "engine_unavailable", "Microsoft Access DAO engine is unavailable; install a compatible Access Database Engine."))
        except Exception:
            issues.append(ScanIssue(str(source), "unreadable", "This source could not be inspected; it may be locked, encrypted, damaged, or inaccessible."))
        for reference in references:
            collect_reference(reference)
        if progress:
            progress(number, len(chosen), str(source))
    return DependencyMapReport(tuple(map(str, chosen)), tuple(inspected),
                               tuple(edges.values()), tuple(unresolved.values()), tuple(issues), depth, was_cancelled)


def filter_dependency_map(report, file_path=None, *, direction="both"):
    """Filter an existing in-memory map without reopening source files."""
    if direction not in {"incoming", "outgoing", "both"}:
        raise ValueError("Direction must be incoming, outgoing or both.")
    if file_path is None:
        return report.edges
    name = str(Path(file_path).expanduser().resolve()).casefold()
    return tuple(edge for edge in report.edges if (
        (direction in {"incoming", "both"} and edge.target_path.casefold() == name)
        or (direction in {"outgoing", "both"} and edge.consumer_path.casefold() == name)
    ))
