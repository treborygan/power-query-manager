import os
os.environ.setdefault('QT_QPA_PLATFORM','offscreen')

from pathlib import Path
from PySide6.QtWidgets import QApplication, QGraphicsView, QPushButton

from pq_installer.core.models import QuerySource
from pq_installer.ui.dependency_graph_dialog import DependencyGraphDialog, analyze_dependency_graph

app=QApplication.instance() or QApplication([])


def _source(root,name,formula,order=0,enabled=True):
    path=Path(root)/(name+'.pq');path.write_text(formula,encoding='utf-8')
    return QuerySource(path,name,order=order,enabled=enabled)


def test_dependency_graph_shows_known_edges_and_detected_cycles(tmp_path):
    sources=[_source(tmp_path,'Sales','let S = Base in S',1),
             _source(tmp_path,'Base','let A = Sales in A',2)]
    graph=analyze_dependency_graph(sources)
    assert graph.edges == (('Sales','Base'),('Base','Sales'))
    assert graph.cycles == (('Base','Sales','Base'),)
    assert all(node.status.startswith('Cycle:') for node in graph.nodes)


def test_dependency_graph_marks_ambiguous_and_disabled_dependencies(tmp_path):
    sources=[_source(tmp_path,'Consumer','let X = Base in X',1),
             _source(tmp_path,'Base','1',2,enabled=False),
             _source(tmp_path,'Dynamic','Expression.Evaluate("Base")',3)]
    graph=analyze_dependency_graph(sources)
    states={node.name:node.status for node in graph.nodes}
    assert states['Consumer']=='Disabled dependency: Base'
    assert states['Base']=='Disabled'
    assert states['Dynamic'].startswith('Ambiguous:')
    assert ('Consumer','Base') in graph.edges


def test_dependency_graph_marks_out_of_set_quoted_reference_without_guessing_edge(tmp_path):
    source=_source(tmp_path,'Consumer','let X = #"Missing Query" in X')
    graph=analyze_dependency_graph([source])
    assert graph.edges==()
    assert graph.nodes[0].ambiguous
    assert 'unresolved quoted identifier' in graph.nodes[0].status


def test_dependency_graph_dialog_has_visual_and_accessible_text_views(tmp_path):
    sources=[_source(tmp_path,'Final','let X = Base in X'),_source(tmp_path,'Base','1')]
    dialog=DependencyGraphDialog(sources)
    assert isinstance(dialog.graph_view,QGraphicsView)
    assert dialog.summary.isReadOnly()
    assert 'Final depends on Base' in dialog.summary.toPlainText()
    dialog.close()


def test_dependency_graph_summary_keeps_reference_after_m_unicode_line_separator_comment(tmp_path):
    sources=[_source(tmp_path,'Consumer','let X = 1 // comment\u2028 in Base'),
             _source(tmp_path,'Base','1')]
    dialog=DependencyGraphDialog(sources)
    assert dialog.analysis.edges==(('Consumer','Base'),)
    assert 'Consumer depends on Base' in dialog.summary.toPlainText()
    dialog.close()


def test_manager_exposes_dependency_graph_for_current_source_list(tmp_path,monkeypatch):
    from pq_installer.ui import main_window
    source=_source(tmp_path,'Base','1')
    window=main_window.MainWindow()
    window.query_model.append_sources([source])
    opened=[]
    monkeypatch.setattr(main_window.DependencyGraphDialog,'exec',lambda self:opened.append(self.analysis))
    button=next(button for button in window.findChildren(QPushButton) if button.text()=='Dependency graph…')
    button.click()
    assert opened and [node.name for node in opened[0].nodes]==['Base']
    window.close()
