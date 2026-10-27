"""General file-to-file dependency index and streaming Excel inspection."""
import base64
import io
import struct
import zipfile

from pq_installer.core.reverse_dependencies import (
    DependencyReference,
    filter_dependency_map,
    scan_dependency_map,
)


def _book(path, *, connection=None, formula=None, defined=None):
    with zipfile.ZipFile(path, "w") as archive:
        archive.writestr("[Content_Types].xml", "<Types/>")
        if connection:
            archive.writestr("xl/connections.xml", (
                '<connections><connection name="WarehouseSource"><dbPr connection="'
                + connection + '"/></connection></connections>'
            ))
        if formula:
            archive.writestr("xl/worksheets/sheet1.xml", (
                '<worksheet><sheetData><row><c><f>' + formula
                + '</f></c></row></sheetData></worksheet>'
            ))
        if defined:
            archive.writestr("xl/workbook.xml", (
                '<workbook><definedNames><definedName name="Remote">' + defined
                + '</definedName></definedNames></workbook>'
            ))


def test_general_map_scans_all_selected_files_including_mutual_targets_once(tmp_path):
    warehouse = tmp_path / "Warehouse.accdb"
    inventory = tmp_path / "Inventory.xlsx"
    dashboard = tmp_path / "Dashboard.xlsx"
    orders = tmp_path / "Orders.mdb"
    warehouse.write_bytes(b"disposable")
    orders.write_bytes(b"disposable")
    _book(inventory, connection="DATABASE=" + str(warehouse) + ";")
    _book(dashboard, connection="DATABASE=" + str(inventory) + ";")
    calls = []

    def access_reader(path):
        calls.append(path)
        if path == orders:
            return (DependencyReference("Access linked table", "Stock", str(warehouse), True),)
        return ()

    report = scan_dependency_map([warehouse, inventory, dashboard, orders], access_reader=access_reader)

    assert set(report.inspected) == {str(warehouse), str(inventory), str(dashboard), str(orders)}
    assert {(edge.consumer_path, edge.target_path) for edge in report.edges} == {
        (str(inventory), str(warehouse)),
        (str(dashboard), str(inventory)),
        (str(orders), str(warehouse)),
    }
    assert calls == [warehouse, orders]
    assert not report.issues

    incoming = filter_dependency_map(report, warehouse, direction="incoming")
    outgoing = filter_dependency_map(report, inventory, direction="outgoing")
    assert {edge.consumer_path for edge in incoming} == {str(inventory), str(orders)}
    assert {(edge.consumer_path, edge.target_path) for edge in outgoing} == {(str(inventory), str(warehouse))}
    assert calls == [warehouse, orders]  # filtering must not reopen any source


def test_mapper_reports_ambiguous_duplicate_basenames_without_guessing_edges(tmp_path):
    one = tmp_path / "a" / "Warehouse.accdb"
    two = tmp_path / "b" / "Warehouse.accdb"
    one.parent.mkdir(); two.parent.mkdir()
    one.write_bytes(b"db"); two.write_bytes(b"db")
    consumer = tmp_path / "Report.xlsx"
    _book(consumer, connection="DATABASE=Warehouse.accdb;")

    report = scan_dependency_map([one, two, consumer], access_reader=lambda _: ())

    assert not report.edges
    assert any(issue.status == "ambiguous_reference" for issue in report.issues)
    assert any(item.consumer_path == str(consumer) and item.reference_filename == "Warehouse.accdb"
               for item in report.unresolved)


def test_general_map_retains_external_file_dependencies_as_unresolved(tmp_path):
    consumer = tmp_path / "Reporting.xlsx"
    _book(consumer, connection="DATABASE=C:\\External\\Warehouse.accdb;")

    report = scan_dependency_map([consumer])

    assert not report.edges
    assert [(ref.consumer_path, ref.reference_filename) for ref in report.unresolved] == [
        (str(consumer), "Warehouse.accdb")
    ]


def test_general_map_detects_genuine_office_utf16_data_mashup(tmp_path):
    target = tmp_path / "Warehouse.accdb"
    target.write_bytes(b"disposable")
    consumer = tmp_path / "Queries.xlsm"
    formula = f'shared AccessSource = Access.Database(File.Contents("{target}"), [CreateNavigationProperties=false]);'
    mashup = io.BytesIO()
    with zipfile.ZipFile(mashup, "w") as package:
        package.writestr("Formulas/Section1.m", formula)
    fields = (mashup.getvalue(), b"", b"", b"\x00")
    payload = struct.pack("<I", 0) + b"".join(
        struct.pack("<I", len(part)) + part for part in fields
    )
    xml = ('<?xml version="1.0" encoding="utf-16"?>'
           '<DataMashup xmlns="http://schemas.microsoft.com/DataMashup">'
           + base64.b64encode(payload).decode("ascii") + "</DataMashup>")
    with zipfile.ZipFile(consumer, "w") as workbook:
        workbook.writestr("customXml/item1.xml", xml.encode("utf-16"))

    report = scan_dependency_map([target, consumer], depth="quick", access_reader=lambda _: ())

    assert [(edge.consumer_path, edge.target_path, edge.source_kind, edge.source_name)
            for edge in report.edges] == [
        (str(consumer), str(target), "Power Query", "AccessSource")
    ]
    assert not report.issues


def test_general_map_finds_multiple_files_mentioned_in_one_saved_query(tmp_path):
    a = tmp_path / "A.accdb"
    b = tmp_path / "B.accdb"
    consumer = tmp_path / "Combination.mdb"
    for file in (a, b, consumer):
        file.write_bytes(b"disposable")
    query = "SELECT * FROM Alpha IN 'A.accdb' UNION SELECT * FROM Beta IN 'B.accdb'"

    def access_reader(path):
        if path == consumer:
            return (DependencyReference("Access saved query", "Merge", query),)
        return ()

    report = scan_dependency_map([a, b, consumer], access_reader=access_reader)

    assert {(edge.target_path, edge.consumer_path) for edge in report.edges} == {
        (str(a), str(consumer)), (str(b), str(consumer))
    }


def test_unresolved_filename_does_not_contain_raw_sql_or_credentials(tmp_path):
    consumer = tmp_path / "Consumer.mdb"
    consumer.write_bytes(b"disposable")
    query = "SELECT secret_column FROM Unmapped.accdb"

    report = scan_dependency_map([consumer], access_reader=lambda _: (
        DependencyReference("Access saved query", "Load", query),
    ))

    assert all("SELECT" not in row.reference_filename for row in report.unresolved)
    assert all("secret_column" not in row.reference_filename for row in report.unresolved)


def test_deep_formula_can_reference_two_external_workbooks(tmp_path):
    a = tmp_path / "A.xlsx"
    b = tmp_path / "B.xlsx"
    consumer = tmp_path / "Both.xlsx"
    _book(a); _book(b)
    _book(consumer, formula="'[A.xlsx]Sheet1'!A1+'[B.xlsx]Sheet1'!A1")

    report = scan_dependency_map([a, b, consumer], depth="deep")

    assert {edge.target_path for edge in report.edges} == {str(a), str(b)}


def test_quick_scan_skips_large_formula_parts_but_deep_scan_includes_them(tmp_path):
    target = tmp_path / "Warehouse.accdb"
    target.write_bytes(b"disposable")
    consumer = tmp_path / "Dashboard.xlsx"
    _book(consumer, formula="'[Warehouse.accdb]Sheet1'!A1", defined="[Warehouse.accdb]Sheet1!A1")

    quick = scan_dependency_map([target, consumer], depth="quick", access_reader=lambda _: ())
    deep = scan_dependency_map([target, consumer], depth="deep", access_reader=lambda _: ())

    assert not quick.edges
    assert len(deep.edges) == 2
    assert {row.source_kind for row in deep.edges} == {"Excel formula", "Excel defined name"}


def test_deep_scanner_streams_large_worksheet_xml_without_zipfile_read(tmp_path, monkeypatch):
    import pq_installer.core.reverse_dependencies as module

    target = tmp_path / "Warehouse.accdb"
    target.write_bytes(b"disposable")
    consumer = tmp_path / "Big.xlsx"
    _book(consumer, formula="'" + " " * 500 + "[Warehouse.accdb]Sheet'!A1")
    monkeypatch.setattr(module, "_MAX_XML_SIZE", 100)
    original_read = zipfile.ZipFile.read

    def reject_full_metadata_read(archive, name, *args, **kwargs):
        part = name.filename if isinstance(name, zipfile.ZipInfo) else str(name)
        if part.startswith("xl/"):
            raise AssertionError("Excel XML metadata must be parsed incrementally")
        return original_read(archive, name, *args, **kwargs)

    monkeypatch.setattr(zipfile.ZipFile, "read", reject_full_metadata_read)
    report = scan_dependency_map([target, consumer], depth="deep", access_reader=lambda _: ())

    assert len(report.edges) == 1
    assert not report.issues


def test_streaming_parser_rejects_unsafe_doctype_in_excel_xml(tmp_path):
    consumer = tmp_path / "Unsafe.xlsx"
    with zipfile.ZipFile(consumer, "w") as archive:
        archive.writestr("xl/connections.xml", (
            '<!DOCTYPE foo [<!ENTITY leak "secret">]>'
            '<connections><connection name="Source"><dbPr connection="&leak;"/></connection></connections>'
        ))

    report = scan_dependency_map([consumer])

    assert any(issue.status == "unreadable" for issue in report.issues)
    assert not report.edges


def test_streamed_excel_references_are_consumed_without_buffering_raw_formulas(tmp_path):
    from pq_installer.core.reverse_dependencies import _excel_references

    consumer = tmp_path / "LargeFormulaGrid.xlsx"
    with zipfile.ZipFile(consumer, "w") as archive:
        archive.writestr("xl/worksheets/sheet1.xml", (
            "<worksheet><sheetData>" + "".join(
                f"<row><c><f>'[External{index}.xlsx]Sheet'!A1</f></c></row>" for index in range(1500)
            ) + "</sheetData></worksheet>"
        ))
    observed = []

    buffered, issues = _excel_references(consumer, depth="deep", on_reference=observed.append)

    assert not buffered
    assert not issues
    assert len(observed) == 1500
    assert observed[0].source_kind == "Excel formula"


def test_repeated_formula_references_do_not_accumulate_duplicate_edges_in_memory(tmp_path):
    import tracemalloc

    target = tmp_path / "Warehouse.accdb"
    target.write_bytes(b"disposable")
    consumer = tmp_path / "Repeated.xlsx"
    with zipfile.ZipFile(consumer, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        archive.writestr("xl/worksheets/sheet1.xml", (
            "<worksheet><sheetData>" +
            "<row><c><f>'[Warehouse.accdb]Sheet'!A1</f></c></row>" * 60000 +
            "</sheetData></worksheet>"
        ))

    tracemalloc.start()
    try:
        report = scan_dependency_map([target, consumer], depth="deep", access_reader=lambda _: ())
        _, peak = tracemalloc.get_traced_memory()
    finally:
        tracemalloc.stop()

    assert len(report.edges) == 1
    assert not report.issues
    assert peak < 6 * 1024 * 1024  # no 60,000-object intermediate match list
