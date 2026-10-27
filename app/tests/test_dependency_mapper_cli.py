import json
import zipfile

from pq_installer.cli import _dispatch_parsed, _make_parser


def test_cli_general_map_indexes_connections_and_emits_json(tmp_path, capsys):
    source = tmp_path / "Source.xlsx"
    consumer = tmp_path / "Dashboard.xlsx"
    with zipfile.ZipFile(source, "w") as archive:
        archive.writestr("[Content_Types].xml", "<Types/>")
    with zipfile.ZipFile(consumer, "w") as archive:
        archive.writestr("xl/connections.xml", (
            '<connections><connection name="Link"><dbPr connection="DATABASE='
            + str(source) + ';"/></connection></connections>'
        ))
    parser = _make_parser()
    args = parser.parse_args(["dependency-map", "--folder", str(tmp_path), "--depth", "quick"])

    code = _dispatch_parsed(args, parser)
    result = json.loads(capsys.readouterr().out)

    assert code == 0
    assert result["bundle_type"] == "dependency-map"
    assert result["depth"] == "quick"
    assert [(edge["consumer_path"], edge["target_path"]) for edge in result["edges"]] == [
        (str(consumer), str(source))
    ]
