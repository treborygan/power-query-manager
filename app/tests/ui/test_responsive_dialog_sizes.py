from pq_installer.core.models import QuerySource
from pq_installer.core.reverse_dependencies import DependencyMapReport
from pq_installer.ui.dependency_finder_dialog import DependencyFinderDialog
from pq_installer.ui.dependency_graph_dialog import DependencyGraphDialog
from pq_installer.ui.file_dependency_graph_dialog import FileDependencyGraphDialog
from pq_installer.ui.main_window import MainWindow
from pq_installer.ui.query_inspector_dialog import QueryInspectorDialog
from PySide6.QtWidgets import QApplication


def test_vector_dialogs_open_within_a_960_by_600_workspace(tmp_path):
    app = QApplication.instance() or QApplication([])
    parent = MainWindow()
    parent.resize(960, 600)
    parent.show()
    app.processEvents()

    finder = DependencyFinderDialog(parent)
    assert finder.size().width() <= parent.width()
    assert finder.size().height() <= parent.height()

    report = DependencyMapReport((), (), (), (), (), "deep")
    file_graph = FileDependencyGraphDialog(report, finder)
    query_graph = DependencyGraphDialog((), parent)
    source_path = tmp_path / "Source.m"
    source_path.write_text("let Source = #table({}, {}) in Source", encoding="utf-8")
    inspector = QueryInspectorDialog(QuerySource(source_path, "Source"), None, parent)

    for dialog in (file_graph, query_graph, inspector):
        owner = dialog.parentWidget()
        assert dialog.size().width() <= owner.width()
        assert dialog.size().height() <= owner.height()

    inspector.close()
    query_graph.close()
    file_graph.close()
    finder.close()
    parent.close()
