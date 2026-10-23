from pq_installer.core.m_source_mapping import preview_m_source_mapping


def test_maps_only_recognized_source_literal_spans_and_preserves_other_m_code():
    source = '''// File.Contents("C:\\old\\comment.csv")
let
    Note = "File.Contents(""C:\\old\\text.csv"")",
    Data = Csv.Document(File.Contents("C:\\old\\folder\\data.csv")),
    Web = Web.Contents("https://example.test/sites/old/items")
in Data'''
    mappings = [
        {'key': 'disk', 'from': 'C:\\old', 'to': 'D:\\new'},
        {'key': 'site', 'from': 'https://example.test/sites/old', 'to': 'https://example.test/sites/new'},
    ]
    result = preview_m_source_mapping(source, mappings)
    assert len(result.changes) == 2
    assert 'File.Contents("D:\\new\\folder\\data.csv")' in result.mapped
    assert 'Web.Contents("https://example.test/sites/new/items")' in result.mapped
    assert 'C:\\old\\comment.csv' in result.mapped
    assert 'C:\\old\\text.csv' in result.mapped
    assert not result.warnings


def test_dynamic_source_expressions_are_reported_and_left_unchanged():
    source = 'File.Contents(root & "/daily.csv")'
    result = preview_m_source_mapping(source, [{'key':'disk','from':'C:\\data','to':'D:\\data'}])
    assert result.mapped == source and not result.changes
    assert len(result.warnings) == 1 and 'File.Contents' in result.warnings[0]


def test_mapping_prefix_requires_a_path_boundary_and_rejects_tied_destinations():
    source = 'File.Contents("C:\\database\\a.csv")'
    unmatched = preview_m_source_mapping(source, [{'key':'disk','from':'C:\\data','to':'D:\\data'}])
    assert not unmatched.changes and not unmatched.warnings
    ambiguous = preview_m_source_mapping('File.Contents("C:\\data\\a.csv")', [
        {'key':'one','from':'C:\\data','to':'D:\\data'},
        {'key':'two','from':'C:/data','to':'E:\\data'},
    ])
    assert not ambiguous.changes and any('ambiguous' in message for message in ambiguous.warnings)


def test_mapped_values_are_escaped_as_m_strings_and_only_matching_query_calls_change():
    source = 'File.Contents("C:\\old\\a.csv") + "C:\\old\\a.csv"'
    result = preview_m_source_mapping(source, [{'key':'disk','from':'C:\\old','to':'D:\\with "quote"'}])
    assert result.mapped == 'File.Contents("D:\\with ""quote""\\a.csv") + "C:\\old\\a.csv"'
    assert len(result.changes) == 1
