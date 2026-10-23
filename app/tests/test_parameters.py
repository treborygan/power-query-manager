from pathlib import Path

import pytest

from pq_installer.core.parameters import edit_parameter, inspect_parameter, edit_parameter_file


def test_declared_text_parameter_edit_preserves_comments_and_m_expression():
    source = '// deployment date\n"Café ""shop""" meta [IsParameterQuery=true, Type="Text", IsParameterQueryRequired=true] // keep this\n'
    parameter = inspect_parameter(source)
    assert parameter.kind == 'Text'
    assert parameter.value == 'Café "shop"'
    edited = edit_parameter(source, 'O\'Neil "North"')
    assert edited == '// deployment date\n"O\'Neil ""North"""' + source[source.index(' meta'):]


@pytest.mark.parametrize(('kind', 'literal', 'value', 'expected'), [
    ('Number', '-12.50', '3.125', '3.125'),
    ('Logical', 'true', 'false', 'false'),
])
def test_declared_numeric_and_logical_parameters(kind, literal, value, expected):
    source = f'{literal} meta [IsParameterQuery = true, Type = "{kind}"]'
    assert inspect_parameter(source).value == (float(literal) if kind == 'Number' else literal == 'true')
    assert edit_parameter(source, value).startswith(expected + ' meta')


@pytest.mark.parametrize('source', [
    'let x = 1 in x',
    '"not declared" meta [Type="Text"]',
    'Date.FromText("2024-01-01") meta [IsParameterQuery=true, Type="Date"]',
    '"value" meta [IsParameterQuery=true, Type="Text"] + "more"',
])
def test_unsupported_or_ambiguous_parameter_expressions_are_read_only(source):
    assert inspect_parameter(source) is None
    with pytest.raises(ValueError):
        edit_parameter(source, 'new')


def test_metadata_lookalikes_inside_description_strings_do_not_enable_parameter_editing():
    source='"value" meta [Description="IsParameterQuery=true, Type=""Text"""]'
    assert inspect_parameter(source) is None


def test_parameter_type_metadata_is_case_insensitive_but_returns_supported_kind():
    source='"old" meta [isparameterquery=true, type="text"]'
    assert inspect_parameter(source).kind == 'Text'
    assert edit_parameter(source,'new').startswith('"new" meta')


def test_parameter_file_edit_preserves_utf8_bom_and_never_overwrites_without_opt_in(tmp_path):
    path = tmp_path / 'region.pq'
    original = b'\xef\xbb\xbf// keep\r\n"EU" meta [IsParameterQuery=true, Type="Text"]'
    path.write_bytes(original)
    with pytest.raises(FileExistsError):
        edit_parameter_file(path, 'US', destination=path)
    assert path.read_bytes() == original
    output = edit_parameter_file(path, 'US', destination=tmp_path / 'edited.pq')
    assert output.read_bytes() == b'\xef\xbb\xbf// keep\r\n"US" meta [IsParameterQuery=true, Type="Text"]'
