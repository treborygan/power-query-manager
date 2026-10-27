import os
os.environ.setdefault('QT_QPA_PLATFORM','offscreen')

from PySide6.QtWidgets import QApplication
from pq_installer.ui.code_editor import CodeEditor, MCodeHighlighter, DiffHighlighter
from pq_installer.ui.main_window import MainWindow

app=QApplication.instance() or QApplication([])


def test_code_editor_has_line_number_gutter_and_m_highlighting():
    editor=CodeEditor(mode='m')
    assert editor.lineNumberAreaWidth()>0
    assert isinstance(editor.highlighter,MCodeHighlighter)
    editor.setPlainText('let\n    x = "hello"\nin x')
    assert editor.blockCount()==3
    editor.close()


def test_diff_editor_uses_diff_highlighter():
    editor=CodeEditor(mode='diff')
    assert isinstance(editor.highlighter,DiffHighlighter)
    editor.close()


def test_highlighters_distinguish_m_tokens_and_diff_changes():
    editor=CodeEditor(mode='m');editor.setPlainText('let "in" // let');app.processEvents()
    ranges=editor.document().firstBlock().layout().formats()
    def color_at(position):
        return next(item.format.foreground().color().name() for item in ranges
                    if item.start<=position<item.start+item.length)
    assert color_at(0)=='#7c3aed'
    assert color_at(4)=='#087443'
    assert color_at(10)=='#596579'
    diff=CodeEditor(mode='diff');diff.setPlainText('+added\n-removed');app.processEvents()
    added_block=diff.document().findBlockByNumber(0);removed_block=diff.document().findBlockByNumber(1)
    added_layout=added_block.layout();removed_layout=removed_block.layout()
    added_ranges=added_layout.formats();removed_ranges=removed_layout.formats()
    added=added_ranges[0].format.background().color().name()
    removed=removed_ranges[0].format.background().color().name()
    assert added=='#e8f5ec' and removed=='#fceaea'
    editor.close();diff.close()


def _luminance(color):
    channels=[int(color[index:index+2],16)/255 for index in (1,3,5)]
    linear=[value/12.92 if value<=0.04045 else ((value+0.055)/1.055)**2.4 for value in channels]
    return sum(value*weight for value,weight in zip(linear,(0.2126,0.7152,0.0722)))


def _contrast(foreground,background):
    first,second=sorted((_luminance(foreground),_luminance(background)),reverse=True)
    return (first+0.05)/(second+0.05)


def _format_at(editor,line,position):
    app.processEvents()
    block=editor.document().findBlockByNumber(line)
    ranges=block.layout().formats()
    selected=next(item for item in ranges if item.start<=position<item.start+item.length)
    formatting=selected.format
    return (formatting.foreground().color().name(),formatting.background().color().name())


def test_m_syntax_colors_meet_contrast_in_light_and_dark_manager_themes():
    window=MainWindow()
    editor=CodeEditor(parent=window,mode='m')
    source='let "text" // note\n    value = 42\n    Table.SelectRows(value)\nin value'
    editor.setPlainText(source)
    tokens=((0,0),(0,4),(0,12),(1,12),(2,4))
    for dark in (False,True):
        window.theme_toggle.setChecked(dark);app.processEvents()
        background=editor.palette().base().color().name()
        for line,position in tokens:
            foreground,_background=_format_at(editor,line,position)
            assert _contrast(foreground,background)>=4.5, (dark,line,position,foreground,background)
    editor.close();window.close()


def test_diff_and_line_number_colors_follow_dark_theme():
    window=MainWindow()
    editor=CodeEditor(parent=window,mode='diff')
    editor.setPlainText('+added\n-removed\n@@ header')
    for dark in (False,True):
        window.theme_toggle.setChecked(dark);app.processEvents()
        for line in range(3):
            foreground,background=_format_at(editor,line,0)
            assert _contrast(foreground,background)>=4.5, (dark,line,foreground,background)
        gutter_foreground,gutter_background=editor.lineNumberArea._text_colors()
        assert _contrast(gutter_foreground,gutter_background)>=4.5
    editor.close();window.close()


def test_query_inspector_editors_follow_parent_theme_toggle(tmp_path):
    from pq_installer.core.models import QuerySource
    from pq_installer.ui.query_inspector_dialog import QueryInspectorDialog

    source_path=tmp_path/'theme.m';source_path.write_text('let text = "value" in text')
    window=MainWindow()
    dialog=QueryInspectorDialog(QuerySource(source_path,'Theme'),None,parent=window)
    assert window.theme_toggle.isChecked()
    assert dialog.source_editor.palette().base().color().name()=='#0b1418'
    window.theme_toggle.setChecked(False);app.processEvents()
    assert dialog.source_editor.palette().base().color().name()=='#f6f9f8'
    window.theme_toggle.setChecked(True);app.processEvents()
    assert dialog.source_editor.palette().base().color().name()=='#0b1418'
    assert dialog.source_editor.highlighter.string_format.foreground().color().name()=='#67dfa0'
    assert dialog.diff_view.highlighter.added.background().color().name()=='#193a2a'
    dialog.close();window.close()


def test_editor_search_advances_wraps_and_searches_backward():
    editor=CodeEditor(mode='m');editor.setPlainText('alpha beta alpha')
    assert editor.find_text('alpha')
    assert editor.textCursor().selectedText()=='alpha'
    assert editor.find_text('alpha')
    assert editor.textCursor().selectionStart()>0
    assert editor.find_text('alpha')
    assert editor.textCursor().selectionStart()==0
    assert editor.find_text('alpha',backward=True)
    assert editor.textCursor().selectionStart()>0
    assert not editor.find_text('missing')
    editor.close()
