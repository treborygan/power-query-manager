import pytest

from pq_installer.core.dependencies import analyze_query_references, find_dependency_cycles, analyze_import_dependencies
from pq_installer.core.models import QuerySource


def test_tokenizer_finds_bare_and_quoted_query_refs_but_ignores_strings_comments_and_fields():
    formula='let X = Base, Y = #"Other Query", S = "Base", R = [Base = 1], F = [Base] // Other Query\n in X'
    result=analyze_query_references(formula,['Base','Other Query'])
    assert result.dependencies==('Base','Other Query')
    assert not result.ambiguous


@pytest.mark.parametrize(('escaped_name', 'query_name'), [
    ('A#(lf)B', 'A\nB'),
    ('A#(000A)B', 'A\nB'),
    ('A#(0000000A)B', 'A\nB'),
    ('A#(#)(lf)B', 'A#(lf)B'),
    ('A""B', 'A"B'),
])
def test_quoted_identifier_character_escapes_are_decoded_for_query_matching(escaped_name, query_name):
    result=analyze_query_references(f'let X = #"{escaped_name}" in X',[query_name])
    assert result.dependencies==(query_name,)
    assert not result.ambiguous


def test_invalid_quoted_identifier_character_escape_is_ambiguous():
    result=analyze_query_references('let X = #"A#(110000)B" in X',['A#(110000)B'])
    assert result.dependencies==()
    assert result.ambiguous


def test_quoted_query_identifier_outside_available_source_list_is_ambiguous():
    result=analyze_query_references('let X = #"Missing Query" in X',['Available'])
    assert result.dependencies==()
    assert result.ambiguous
    assert 'unresolved quoted identifier' in result.reason


@pytest.mark.parametrize('newline',['\r','\r\n','\n','\u0085','\u2028','\u2029'])
def test_line_comment_ends_at_every_m_newline_form(newline):
    formula=f'let X = 1 // Base is only a comment{newline} in Base'
    result=analyze_query_references(formula,['Base'])
    assert result.dependencies==('Base',)
    assert not result.ambiguous


@pytest.mark.parametrize('query_name',['A\u0301','A\u203f','A\u200c','ⅫBase'])
def test_regular_identifier_tokenizer_matches_m_unicode_identifier_classes(query_name):
    result=analyze_query_references(f'let X = {query_name} in X',[query_name])
    assert result.dependencies==(query_name,)
    assert not result.ambiguous


def test_local_binding_collision_and_dynamic_reference_are_flagged_ambiguous():
    shadow=analyze_query_references('let Base = 1 in Base',['Base'])
    assert shadow.dependencies==() and shadow.ambiguous
    dynamic=analyze_query_references('Expression.Evaluate("Base")',['Base'])
    assert dynamic.dependencies==() and dynamic.ambiguous


def test_query_and_local_identifier_matching_respects_m_case_sensitivity():
    exact=analyze_query_references('let x = Base in x',['Base'])
    assert exact.dependencies==('Base',) and not exact.ambiguous
    differently_cased=analyze_query_references('let x = base in x',['Base'])
    assert differently_cased.dependencies==() and not differently_cased.ambiguous
    local_different_case=analyze_query_references('let base = 1, x = Base in x',['Base'])
    assert local_different_case.dependencies==('Base',) and not local_different_case.ambiguous


def test_dynamic_reference_detection_respects_m_identifier_case():
    actual=analyze_query_references('Expression.Evaluate("Base")',['Base'])
    assert actual.ambiguous and 'dynamic' in actual.reason
    different_case=analyze_query_references('expression.evaluate("Base")',['Base'])
    assert not different_case.ambiguous and different_case.dependencies==()


def test_dotted_regular_identifier_matches_only_the_complete_query_name():
    result=analyze_query_references('let X = Base.Field in X',['Base','Base.Field','Field'])
    assert result.dependencies==('Base.Field',)
    assert not result.ambiguous


def test_dotted_generalized_identifier_field_label_is_not_a_query_reference():
    result=analyze_query_references('let X = [Base.Field = 1] in X',['Base.Field'])
    assert result.dependencies==()
    assert not result.ambiguous


def test_dotted_quoted_identifier_is_not_treated_as_a_dynamic_builtin():
    result=analyze_query_references('let X = #"Expression.Evaluate" in X',['Expression.Evaluate'])
    assert result.dependencies==('Expression.Evaluate',)
    assert not result.ambiguous


def test_qualified_record_field_function_is_marked_ambiguous():
    result=analyze_query_references('Record.Field([Base = 1], "Base")',['Base'])
    assert result.dependencies==()
    assert result.ambiguous and 'dynamic' in result.reason


@pytest.mark.parametrize(('keyword','literal'),[
    ('binary','#binary("AQID")'),
    ('date','#date(2020,1,1)'),
    ('datetime','#datetime(2020,1,1,0,0,0)'),
    ('datetimezone','#datetimezone(2020,1,1,0,0,0,0,0)'),
    ('duration','#duration(0,1,0,0)'),
    ('infinity','#infinity'),
    ('nan','#nan'),
    ('table','#table({"Column1"},{{1}})'),
    ('time','#time(12,0,0)'),
])
def test_hash_prefixed_m_literal_keyword_is_not_a_query_reference(keyword,literal):
    result=analyze_query_references(f'let X = {literal} in X',[keyword])
    assert result.dependencies==()
    assert not result.ambiguous


@pytest.mark.parametrize('keyword',['sections','shared'])
def test_hash_prefixed_dynamic_m_keyword_is_ambiguous(keyword):
    result=analyze_query_references(f'let X = #{keyword} in X',[keyword])
    assert result.dependencies==()
    assert result.ambiguous and 'dynamic' in result.reason


def test_unknown_hash_prefixed_identifier_is_ambiguous_without_dependency_guess():
    result=analyze_query_references('let X = #mystery in X',['mystery'])
    assert result.dependencies==()
    assert result.ambiguous


def test_disabled_query_dependency_matching_respects_m_case_sensitivity(tmp_path):
    disabled=tmp_path/'disabled.m'; enabled=tmp_path/'enabled.m'
    disabled.write_text('1'); enabled.write_text('let x = Base in x')
    rows=[QuerySource(disabled,'base',enabled=False),QuerySource(enabled,'Consumer')]
    result=analyze_import_dependencies(rows)
    assert result['Consumer']=='No direct query references detected'
    rows[0]=QuerySource(disabled,'Base',enabled=False)
    result=analyze_import_dependencies(rows)
    assert result['Consumer']=='Disabled dependency: Base'


def test_dependency_cycles_report_path_not_suggested_reorder():
    result=find_dependency_cycles({'A':['B'],'B':['A'],'C':['A']})
    assert result==(('A','B','A'),)


def test_import_dependency_report_flags_disabled_references_and_cycles(tmp_path):
    a=tmp_path/'a.m';b=tmp_path/'b.m';c=tmp_path/'c.m'
    a.write_text('let x = B in x');b.write_text('let x = A in x');c.write_text('let x = A in x')
    sources=[QuerySource(a,'A'),QuerySource(b,'B'),QuerySource(c,'C',enabled=False)]
    results=analyze_import_dependencies(sources)
    assert results['A'].startswith('Cycle: A -> B -> A')
    assert results['B'].startswith('Cycle: A -> B -> A')
    assert results['C']=='Disabled query; dependent imports should review this source.'
