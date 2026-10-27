"""Behavioral tests for filesystem-based, read-only incoming dependencies."""
import base64
import io
import struct
import sys
from types import SimpleNamespace
import zipfile

from pq_installer.core.reverse_dependencies import (
    DependencyReference,
    discover_scan_files,
    scan_reverse_dependencies,
)
from pq_installer.core.access_dependency_reader import read_access_references
from pq_installer.core.access_dependency_reader import AccessEngineUnavailable


def _workbook(path, *, m=None, connection=None, external=None, formula=None):
    with zipfile.ZipFile(path, "w") as workbook:
        workbook.writestr("[Content_Types].xml", "<Types/>")
        if m is not None:
            formula_zip = io.BytesIO()
            with zipfile.ZipFile(formula_zip, "w") as package:
                package.writestr("Formulas/Section1.m", m)
            fields = (formula_zip.getvalue(), b"", b"", b"\x00")
            payload = struct.pack("<I", 0) + b"".join(
                struct.pack("<I", len(part)) + part for part in fields
            )
            workbook.writestr(
                "customXml/item1.xml",
                '<DataMashup xmlns="http://schemas.microsoft.com/DataMashup">'
                + base64.b64encode(payload).decode("ascii") + "</DataMashup>",
            )
        if connection is not None:
            workbook.writestr("xl/connections.xml", (
                '<connections xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
                '<connection name="AccessSource"><dbPr connection="'
                + connection.replace("&", "&amp;").replace('"', '&quot;')
                + '" /></connection></connections>'
            ))
        if external is not None:
            workbook.writestr("xl/externalLinks/_rels/externalLink1.xml.rels", (
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                '<Relationship Id="rId1" TargetMode="External" Target="'
                + external + '" /></Relationships>'
            ))
        if formula is not None:
            workbook.writestr("xl/worksheets/sheet1.xml", (
                '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
                '<sheetData><row r="1"><c r="A1"><f>'
                + formula + '</f></c></row></sheetData></worksheet>'
            ))


def test_discovery_includes_access_and_excel_files_and_can_recurse(tmp_path):
    root = tmp_path / "MappedDrive"
    root.mkdir()
    (root / "Warehouse.accdb").write_bytes(b"db")
    (root / "Orders.mdb").write_bytes(b"db")
    (root / "Report.xlsx").write_bytes(b"zip")
    (root / "Book.xlsb").write_bytes(b"old")
    (root / "~$Report.xlsx").write_bytes(b"lock")
    nested = root / "Nested"
    nested.mkdir()
    (nested / "Report.xlsm").write_bytes(b"zip")

    flat = discover_scan_files([root], recursive=False)
    deep = discover_scan_files([root], recursive=True)

    assert {p.name for p in flat} == {"Warehouse.accdb", "Orders.mdb", "Report.xlsx", "Book.xlsb"}
    assert {p.name for p in deep} == {"Warehouse.accdb", "Orders.mdb", "Report.xlsx", "Report.xlsm", "Book.xlsb"}


def test_reverse_scan_reads_excel_power_query_connections_and_links(tmp_path):
    target = tmp_path / "Warehouse.accdb"
    target.write_bytes(b"target")
    pq = tmp_path / "Queries.xlsx"
    _workbook(pq, m='section Section1; shared Inventory = Access.Database(File.Contents("'
              + str(target) + '"));')
    connection = tmp_path / "Connected.xlsm"
    _workbook(connection, connection="Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + str(target) + ";")
    missed = tmp_path / "Unrelated.xlsx"
    _workbook(missed, m='section Section1; shared Other = File.Contents("another.accdb");')
    unrelated_comment = tmp_path / "Comment.xlsx"
    _workbook(unrelated_comment, m='section Section1; // Warehouse.accdb\nshared Other = 1;')
    link = tmp_path / "Link.xlsx"
    _workbook(link, external=target.as_uri())

    result = scan_reverse_dependencies([target], [pq, connection, missed, unrelated_comment, link])

    assert {row.consumer_path for row in result.hits} == {str(pq), str(connection), str(link)}
    assert all(row.target_path == str(target) for row in result.hits)
    assert all(row.confidence == "exact_path" for row in result.hits)
    assert next(hit.source_name for hit in result.hits if hit.consumer_path == str(pq)) == "Inventory"
    assert not result.issues


def test_multiple_targets_match_access_records_and_exclude_self(tmp_path):
    warehouse = tmp_path / "Warehouse.accdb"
    replenishment = tmp_path / "Replenishment v1.0.xlsx"
    for path in (warehouse, replenishment):
        path.write_bytes(b"target")
    consumer = tmp_path / "Orders.mdb"
    consumer.write_bytes(b"fake")

    def access_reader(path):
        assert path == consumer
        return (
            DependencyReference("Access linked table", "WarehouseStock", str(warehouse), True),
            DependencyReference("Access saved query", "Pending", "Replenishment v1.0.xlsx", False),
        )

    result = scan_reverse_dependencies(
        [warehouse, replenishment], [warehouse, consumer, replenishment], access_reader=access_reader
    )
    assert len(result.hits) == 2
    assert {hit.target_path: hit.confidence for hit in result.hits} == {
        str(warehouse): "exact_path", str(replenishment): "possible_literal"
    }
    assert {hit.consumer_path for hit in result.hits} == {str(consumer)}


def test_m_dynamic_path_is_reported_as_incomplete_not_as_a_confirmed_match(tmp_path):
    target = tmp_path / "Warehouse.accdb"
    consumer = tmp_path / "Dynamic.xlsx"
    target.write_bytes(b"target")
    _workbook(consumer, m='section Section1; shared Source = File.Contents(SourcePath);')

    report = scan_reverse_dependencies([target], [consumer])

    assert not report.hits
    assert any(issue.status == "dynamic_reference" for issue in report.issues)


def test_m_query_name_ignores_fake_declarations_inside_comments_and_literals(tmp_path):
    target = tmp_path / "Warehouse.accdb"
    target.write_bytes(b"target")
    workbook = tmp_path / "Queries.xlsx"
    m_code = ('section Section1; shared Real = let '
              'SomeText = "shared FakeString =", '
              '// shared FakeComment =\n'
              'Source = File.Contents("' + str(target) + '") in Source;')
    _workbook(workbook, m=m_code)

    report = scan_reverse_dependencies([target], [workbook])

    assert len(report.hits) == 1
    assert report.hits[0].source_name == "Real"


def test_unsupported_and_unreadable_candidates_are_visible_in_report(tmp_path):
    target = tmp_path / "Warehouse.accdb"
    target.write_bytes(b"target")
    unsupported = tmp_path / "Old.xls"
    unsupported.write_bytes(b"binary")
    broken = tmp_path / "Broken.xlsx"
    broken.write_bytes(b"not zip")

    result = scan_reverse_dependencies([target], [unsupported, broken])

    assert not result.hits
    assert {issue.status for issue in result.issues} == {"unsupported", "unreadable"}


def test_excel_metadata_decompression_has_a_total_scan_budget(tmp_path, monkeypatch):
    from pq_installer.core import reverse_dependencies as module

    target = tmp_path / "Warehouse.accdb"
    target.write_bytes(b"db")
    workbook = tmp_path / "LargeMetadata.xlsx"
    with zipfile.ZipFile(workbook, "w") as archive:
        archive.writestr("xl/worksheets/sheet1.xml", "<worksheet>" + " " * 60 + "</worksheet>")
        archive.writestr("xl/worksheets/sheet2.xml", "<worksheet>" + " " * 60 + "</worksheet>")
    monkeypatch.setattr(module, "_MAX_TOTAL_METADATA_BYTES", 100, raising=False)

    report = scan_reverse_dependencies([target], [workbook])

    assert any(issue.status == "unreadable" and "budget" in issue.detail.casefold() for issue in report.issues)


def test_scan_can_cancel_between_files(tmp_path):
    target = tmp_path / "Warehouse.accdb"
    target.write_bytes(b"target")
    workbooks = [tmp_path / f"Book{index}.xlsx" for index in range(3)]
    for workbook in workbooks:
        _workbook(workbook)
    completed = []

    result = scan_reverse_dependencies(
        [target], workbooks,
        progress=lambda done, total, path: completed.append(done),
        cancelled=lambda: bool(completed),
    )

    assert result.cancelled
    assert completed == [1]
    assert len(result.inspected) == 1


def test_target_only_selection_reports_no_consumers_instead_of_zero_dependencies(tmp_path):
    target = tmp_path / "Warehouse.accdb"
    target.write_bytes(b"database")

    result = scan_reverse_dependencies([target], [target])

    assert not result.hits
    assert any(issue.status == "no_consumers" for issue in result.issues)


def test_missing_access_engine_is_visible_as_partial_coverage(tmp_path):
    target = tmp_path / "Warehouse.accdb"
    consumer = tmp_path / "Report.mdb"
    target.write_bytes(b"database")
    consumer.write_bytes(b"database")

    def unavailable(path):
        raise AccessEngineUnavailable("Not installed")

    result = scan_reverse_dependencies([target], [consumer], access_reader=unavailable)

    assert not result.hits
    assert len(result.issues) == 1
    assert result.issues[0].status == "engine_unavailable"


def test_access_inspector_reads_temporary_copy_using_injected_dao(tmp_path):
    source = tmp_path / "Orders.accdb"
    source.write_bytes(b"example database")
    opened = []

    class Obj:
        def __init__(self, **kwargs):
            self.__dict__.update(kwargs)

    class Database:
        TableDefs = (Obj(Name="WarehouseStock", Connect=";DATABASE=C:\\Data\\Warehouse.accdb"),)
        QueryDefs = (Obj(Name="Q1", SQL="SELECT * FROM [Stock] IN 'C:\\Data\\Warehouse.accdb'", Connect=""),)
        def Close(self):
            pass

    class Engine:
        def OpenDatabase(self, name, exclusive, read_only):
            opened.append((name, exclusive, read_only))
            assert name != str(source)
            assert open(name, "rb").read() == source.read_bytes()
            return Database()

    refs = read_access_references(source, engine_factory=Engine)

    assert len(refs) == 2
    assert {row.source_kind for row in refs} == {"Access linked table", "Access saved query"}
    assert opened[0][1:] == (False, True)
    assert not (tmp_path / "Orders.laccdb").exists()


def test_m_unused_absolute_path_literal_is_not_claimed_as_confirmed_dependency(tmp_path):
    target = tmp_path / "Warehouse.accdb"
    target.write_bytes(b"target")
    workbook = tmp_path / "UnusedLiteral.xlsx"
    _workbook(workbook, m='section Section1; shared Unused = "' + str(target) + '";')

    report = scan_reverse_dependencies([target], [workbook])

    assert len(report.hits) == 1
    assert report.hits[0].confidence == "possible_literal"


def test_access_default_engine_initializes_and_uninitializes_com(tmp_path, monkeypatch):
    from pq_installer.core import access_dependency_reader

    source = tmp_path / "Orders.mdb"
    source.write_bytes(b"database fixture")
    events = []

    class FakeDatabase:
        TableDefs = ()
        QueryDefs = ()
        def Close(self):
            events.append("close")

    class FakeEngine:
        def OpenDatabase(self, path, exclusive, readonly):
            events.append("open")
            assert exclusive is False and readonly is True
            return FakeDatabase()

    monkeypatch.setitem(sys.modules, "pythoncom", SimpleNamespace(
        CoInitialize=lambda: events.append("initialize"),
        CoUninitialize=lambda: events.append("uninitialize"),
    ))
    monkeypatch.setattr(access_dependency_reader, "_default_engine", lambda: FakeEngine())

    assert read_access_references(source) == ()
    assert events == ["initialize", "open", "close", "uninitialize"]


def test_access_releases_dao_proxies_before_temp_directory_cleanup(tmp_path, monkeypatch):
    from pq_installer.core import access_dependency_reader as module

    source = tmp_path / "Warehouse.accdb"
    source.write_bytes(b"database")
    observed = []
    actual_temporary = module.tempfile.TemporaryDirectory

    class ObservedTemporary:
        def __init__(self, *args, **kwargs):
            self._wrapped = actual_temporary(*args, **kwargs)
        def __enter__(self):
            return self._wrapped.__enter__()
        def __exit__(self, *args):
            observed.append("temp_cleanup")
            return self._wrapped.__exit__(*args)

    class Database:
        TableDefs = ()
        QueryDefs = ()
        def Close(self):
            observed.append("db_close")
        def __del__(self):
            observed.append("db_released")

    class Engine:
        def OpenDatabase(self, path, exclusive, read_only):
            return Database()
        def __del__(self):
            observed.append("engine_released")

    monkeypatch.setattr(module.tempfile, "TemporaryDirectory", ObservedTemporary)
    assert read_access_references(source, engine_factory=Engine) == ()
    assert observed.index("db_close") < observed.index("db_released") < observed.index("temp_cleanup")
    assert observed.index("engine_released") < observed.index("temp_cleanup")
