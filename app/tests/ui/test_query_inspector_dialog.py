import os
os.environ.setdefault('QT_QPA_PLATFORM','offscreen')
from pathlib import Path
from PySide6.QtWidgets import QApplication
from pq_installer.core.manager import QueryInfo
from pq_installer.core.models import QuerySource
from pq_installer.ui.query_inspector_dialog import QueryInspectorDialog

app=QApplication.instance() or QApplication([])


def test_inspector_searches_source_workbook_and_exact_diff(tmp_path):
    path=tmp_path/'search.pq';path.write_text('let sourceValue = 1 in sourceValue',encoding='utf-8')
    dialog=QueryInspectorDialog(QuerySource(path,'Search'),QueryInfo('Search','let workbookValue = 2 in workbookValue'))
    dialog.search_edit.setText('sourceValue');dialog._find_in('source')
    assert dialog.source_editor.textCursor().selectedText()=='sourceValue'
    dialog.search_edit.setText('workbookValue');dialog._find_in('workbook')
    assert dialog.workbook_view.textCursor().selectedText()=='workbookValue'
    dialog.search_edit.setText('sourceValue');dialog._find_in('diff')
    assert dialog.diff_view.textCursor().selectedText()=='sourceValue'
    dialog.close()


def test_query_inspector_editors_allow_tab_to_move_focus(tmp_path):
    path=tmp_path/'focus.pq';path.write_text('let Source = 1 in Source',encoding='utf-8')
    dialog=QueryInspectorDialog(QuerySource(path,'Focus'),QueryInfo('Focus','let Source = 1 in Source'))

    assert dialog.source_editor.tabChangesFocus()
    assert dialog.workbook_view.tabChangesFocus()
    assert dialog.diff_view.tabChangesFocus()

def test_unchanged_source_copy_preserves_exact_bytes_and_edits_write_a_separate_file(tmp_path):
    original=tmp_path/'input.m';original.write_bytes(b'\xef\xbb\xbflet x = "caf\xc3\xa9" in x\r\n')
    source=QuerySource(original,'input')
    dialog=QueryInspectorDialog(source,QueryInfo('input','let x = 1 in x'))
    unchanged=tmp_path/'unchanged.pq';dialog.save_copy(unchanged)
    assert unchanged.read_bytes()==original.read_bytes()
    dialog.source_editor.insertPlainText('\n// user edit')
    edited=tmp_path/'edited.pq';dialog.save_copy(edited)
    assert original.read_bytes().startswith(b'\xef\xbb\xbf')
    assert b'user edit' in edited.read_bytes() and edited!=original
    dialog.close()


def test_declared_parameter_edit_updates_only_literal_and_unsupported_query_is_read_only(tmp_path,monkeypatch):
    from PySide6.QtWidgets import QInputDialog,QMessageBox
    path=tmp_path/'region.pq';path.write_bytes(b'\xef\xbb\xbf// keep\r\n"EU" meta [IsParameterQuery=true, Type="Text"] // tail\r\n')
    dialog=QueryInspectorDialog(QuerySource(path,'Region'),None)
    assert dialog.parameter_button.isEnabled()
    monkeypatch.setattr(QInputDialog,'getText',lambda *args,**kwargs:('US',True))
    dialog.parameter_button.click()
    assert dialog.source_editor.toPlainText()=='// keep\n"US" meta [IsParameterQuery=true, Type="Text"] // tail\n'
    saved=tmp_path/'region_edited.pq';dialog.save_copy(saved)
    assert saved.read_bytes()==b'\xef\xbb\xbf// keep\r\n"US" meta [IsParameterQuery=true, Type="Text"] // tail\r\n'
    unsupported=tmp_path/'complex.pq';unsupported.write_text('let x = 1 in x',encoding='utf-8')
    second=QueryInspectorDialog(QuerySource(unsupported,'Complex'),None)
    assert not second.parameter_button.isEnabled()
    monkeypatch.setattr(QMessageBox,'warning',lambda *args,**kwargs:QMessageBox.Discard)
    dialog.close();second.close()
