import json
import zipfile

from pq_installer.cli import _dispatch_parsed, _make_parser


def test_cli_reverse_dependency_scan_supports_folder_and_multiple_targets(tmp_path, capsys):
    target = tmp_path / "Warehouse.accdb"
    extra = tmp_path / "Replenishment.xlsx"
    target.write_bytes(b"db")
    with zipfile.ZipFile(extra, "w") as workbook:
        workbook.writestr("[Content_Types].xml", "<Types/>")
    consumer = tmp_path / "Report.xlsx"
    with zipfile.ZipFile(consumer, "w") as workbook:
        workbook.writestr("xl/connections.xml", (
            '<connections><connection name="WarehouseSource">'
            '<dbPr connection="DATABASE=' + str(target) + ';"/>'
            '</connection></connections>'
        ))
    parser = _make_parser()
    args = parser.parse_args([
        "dependency-scan", "--target", str(target), "--target", str(extra),
        "--folder", str(tmp_path), "--recursive",
    ])

    code = _dispatch_parsed(args, parser)
    output = json.loads(capsys.readouterr().out)

    assert code == 0
    assert len(output["targets"]) == 2
    assert len(output["hits"]) == 1
    assert output["hits"][0]["consumer_path"] == str(consumer)
    assert output["hits"][0]["target_path"] == str(target)


def test_cli_scan_reports_partial_coverage_in_exit_code(tmp_path, capsys):
    target = tmp_path / "Warehouse.accdb"
    target.write_bytes(b"db")
    unsupported = tmp_path / "Legacy.xls"
    unsupported.write_bytes(b"old")
    parser = _make_parser()
    args = parser.parse_args(["dependency-scan", "--target", str(target), "--file", str(unsupported)])

    code = _dispatch_parsed(args, parser)
    output = json.loads(capsys.readouterr().out)

    assert code == 5
    assert output["issues"][0]["status"] == "unsupported"
