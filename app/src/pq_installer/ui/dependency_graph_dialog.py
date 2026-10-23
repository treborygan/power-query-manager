"""Visual and accessible advisory view of confidently detected M references."""
from __future__ import annotations

from dataclasses import dataclass

from PySide6.QtCore import Qt
from PySide6.QtGui import QBrush, QColor, QFont, QPainter, QPainterPath, QPen
from PySide6.QtWidgets import (
    QDialog, QDialogButtonBox, QGraphicsPathItem, QGraphicsScene,
    QGraphicsView, QLabel, QPlainTextEdit, QVBoxLayout,
)

from pq_installer.core.dependencies import analyze_query_references, find_dependency_cycles
from pq_installer.core.query_loader import read_m_code


@dataclass(frozen=True)
class DependencyNode:
    name: str
    status: str
    enabled: bool
    ambiguous: bool = False


@dataclass(frozen=True)
class DependencyGraph:
    nodes: tuple[DependencyNode, ...]
    edges: tuple[tuple[str, str], ...]
    cycles: tuple[tuple[str, ...], ...]


def analyze_dependency_graph(sources) -> DependencyGraph:
    """Return only confident edges; ambiguous M stays visible without guessed links."""
    sources = tuple(sources)
    names = [source.name for source in sources]
    by_name = {source.name.casefold(): source for source in sources}
    analyses = {}
    for source in sources:
        if not source.enabled:
            continue
        try:
            analyses[source.name.casefold()] = analyze_query_references(read_m_code(source), names)
        except Exception as exc:
            analyses[source.name.casefold()] = exc

    edges = []
    graph = {source.name: [] for source in sources if source.enabled}
    for source in sources:
        analysis = analyses.get(source.name.casefold())
        if source.enabled and analysis is not None and not isinstance(analysis, Exception) and not analysis.ambiguous:
            for dependency in analysis.dependencies:
                edges.append((source.name, dependency))
                if dependency in graph:
                    graph[source.name].append(dependency)

    cycles = find_dependency_cycles(graph)
    cycle_for = {name: cycle for cycle in cycles for name in cycle[:-1]}
    nodes = []
    for source in sources:
        if not source.enabled:
            status, ambiguous = 'Disabled', False
        else:
            analysis = analyses.get(source.name.casefold())
            if isinstance(analysis, Exception):
                status, ambiguous = f'Unknown: {analysis}', True
            elif analysis is None:
                status, ambiguous = 'Unknown dependency status', True
            elif analysis.ambiguous:
                status = 'Ambiguous' + (f': {analysis.reason}' if analysis.reason else ' M reference')
                ambiguous = True
            else:
                disabled_refs = [name for name in analysis.dependencies
                                 if name.casefold() in by_name and not by_name[name.casefold()].enabled]
                if disabled_refs:
                    status, ambiguous = 'Disabled dependency: ' + ', '.join(disabled_refs), False
                elif source.name in cycle_for:
                    status, ambiguous = 'Cycle: ' + ' -> '.join(cycle_for[source.name]), False
                elif analysis.dependencies:
                    status, ambiguous = 'Depends on: ' + ', '.join(analysis.dependencies), False
                else:
                    status, ambiguous = 'No direct query references detected', False
        nodes.append(DependencyNode(source.name, status, source.enabled, ambiguous))

    order = {source.name: index for index, source in enumerate(sources)}
    edges.sort(key=lambda edge: (order.get(edge[0], 0), order.get(edge[1], 0), edge))
    return DependencyGraph(tuple(nodes), tuple(edges), cycles)


class _GraphView(QGraphicsView):
    def __init__(self, scene, parent=None):
        super().__init__(scene, parent)
        self.setRenderHint(QPainter.RenderHint.Antialiasing)
        self.setDragMode(QGraphicsView.DragMode.ScrollHandDrag)
        self.setAccessibleName('Query dependency graph')
        self.setMinimumHeight(280)

    def resizeEvent(self, event):
        super().resizeEvent(event)
        if self.scene().items():
            self.fitInView(self.sceneRect(), Qt.AspectRatioMode.KeepAspectRatio)


class DependencyGraphDialog(QDialog):
    def __init__(self, sources, parent=None):
        super().__init__(parent)
        self.setWindowTitle('Query dependencies')
        self.resize(900, 620)
        self.analysis = analyze_dependency_graph(sources)
        layout = QVBoxLayout(self)
        note = QLabel('Arrows point from a query to a detected dependency. This bounded M scanner is advisory, not a full language parser. Ambiguous references are shown without guessed arrows; import order is unchanged.')
        note.setWordWrap(True)
        layout.addWidget(note)
        self.graph_view = _GraphView(self._build_scene(), self)
        self.graph_view.setAccessibleName('Visual query dependency diagram')
        layout.addWidget(self.graph_view, 3)
        self.summary = QPlainTextEdit(self)
        self.summary.setReadOnly(True)
        self.summary.setAccessibleName('Text summary of query dependencies')
        self.summary.setMaximumHeight(170)
        self.summary.setPlainText(self._summary_text())
        layout.addWidget(self.summary, 1)
        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Close)
        buttons.setAccessibleName('Query dependency dialog actions')
        close_button=buttons.button(QDialogButtonBox.StandardButton.Close)
        if close_button is not None:
            close_button.setAccessibleName('Close query dependency graph')
        buttons.rejected.connect(self.reject)
        buttons.accepted.connect(self.accept)
        layout.addWidget(buttons)

    def _summary_text(self):
        lines = []
        for node in self.analysis.nodes:
            outgoing = [target for source, target in self.analysis.edges if source == node.name]
            if outgoing:
                lines.append(f"{node.name} depends on {', '.join(outgoing)}")
            else:
                lines.append(f'{node.name}: {node.status}')
        for cycle in self.analysis.cycles:
            lines.append('Detected cycle: ' + ' → '.join(cycle))
        return '\n'.join(lines) or 'No source queries are loaded.'

    def _build_scene(self):
        scene = QGraphicsScene(self)
        if not self.analysis.nodes:
            scene.addText('Add query sources to see detected dependencies.')
            scene.setSceneRect(0, 0, 420, 120)
            return scene

        nodes = {node.name: node for node in self.analysis.nodes}
        cycle_nodes = {name for cycle in self.analysis.cycles for name in cycle[:-1]}
        outgoing = {name: [] for name in nodes}
        incoming = {name: [] for name in nodes}
        for source, dependency in self.analysis.edges:
            if source in outgoing and dependency in nodes:
                outgoing[source].append(dependency)
                incoming[dependency].append(source)

        levels = {}
        def level(name, visiting=frozenset()):
            if name in levels:
                return levels[name]
            if name in visiting or name in cycle_nodes:
                levels[name] = 0
            else:
                dependencies = outgoing[name]
                levels[name] = max((level(dep, visiting | {name}) + 1 for dep in dependencies), default=0)
            return levels[name]
        for name in nodes:
            level(name)

        order = {node.name: index for index, node in enumerate(self.analysis.nodes)}
        columns = {}
        for name, node_level in levels.items():
            columns.setdefault(node_level, []).append(name)
        for col in columns:
            columns[col].sort(key=lambda name: order[name])

        width, height, gap_x, gap_y = 220, 72, 72, 28
        positions = {}
        for col, names in sorted(columns.items()):
            for row, name in enumerate(names):
                positions[name] = (24 + col * (width + gap_x), 24 + row * (height + gap_y))

        # Draw connectors first so node cards remain on top.
        for source, dependency in self.analysis.edges:
            if source not in positions or dependency not in positions:
                continue
            sx, sy = positions[source]; dx, dy = positions[dependency]
            start = (sx, sy + height / 2)
            end = (dx + width, dy + height / 2)
            path = QPainterPath()
            path.moveTo(*start)
            middle = (start[0] + end[0]) / 2
            path.cubicTo(middle, start[1], middle, end[1], *end)
            line = QGraphicsPathItem(path)
            line.setPen(QPen(QColor('#59758f'), 2))
            line.setZValue(0)
            scene.addItem(line)
            direction = QGraphicsPathItem()
            arrow_path = QPainterPath()
            arrow_path.moveTo(end[0], end[1])
            arrow_path.lineTo(end[0] - 10, end[1] - 5)
            arrow_path.lineTo(end[0] - 10, end[1] + 5)
            arrow_path.closeSubpath()
            direction.setPath(arrow_path)
            direction.setBrush(QBrush(QColor('#59758f')))
            direction.setPen(QPen(QColor('#59758f')))
            scene.addItem(direction)

        for name, node in nodes.items():
            x, y = positions[name]
            if not node.enabled:
                fill = QColor('#e8edf2')
            elif node.ambiguous:
                fill = QColor('#fff0ca')
            elif name in cycle_nodes:
                fill = QColor('#fde2e1')
            else:
                fill = QColor('#e5f1ff')
            card = scene.addRect(x, y, width, height, QPen(QColor('#73869a'), 1.3), QBrush(fill))
            card.setToolTip(f'{name}\n{node.status}')
            card.setZValue(1)
            title = scene.addText(name, QFont('', 10, QFont.Weight.Bold))
            title.setPos(x + 9, y + 5)
            title.setToolTip(f'{name}\n{node.status}')
            status = scene.addText(node.status, QFont('', 8))
            status.setTextWidth(width - 16)
            status.setPos(x + 8, y + 31)
            status.setDefaultTextColor(QColor('#30465b'))
            status.setToolTip(node.status)
        cols = max(columns, default=0) + 1
        rows = max((len(names) for names in columns.values()), default=1)
        scene.setSceneRect(0, 0, cols * (width + gap_x) + 40, rows * (height + gap_y) + 50)
        return scene
