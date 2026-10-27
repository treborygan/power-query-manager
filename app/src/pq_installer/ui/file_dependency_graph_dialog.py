"""Advisory diagram of file-level incoming/outgoing dependencies."""
from __future__ import annotations

from math import atan2, cos, hypot, sin
from pathlib import Path

from PySide6.QtCore import QPointF, Qt
from PySide6.QtGui import QBrush, QColor, QFont, QPainter, QPainterPath, QPen, QPolygonF
from PySide6.QtWidgets import (
    QDialog, QDialogButtonBox, QGraphicsScene, QGraphicsView,
    QLabel, QPlainTextEdit, QVBoxLayout,
)


_MAX_GRAPH_NODES = 80  # full adjacency remains in the report/CSV/JSON


class FileDependencyGraphDialog(QDialog):
    def __init__(self, report, parent=None):
        super().__init__(parent)
        self.report = report
        self.setWindowTitle("File dependencies — advisory map")
        self.resize(1000, 680)
        layout = QVBoxLayout(self)
        note = QLabel("Arrows point from a file to the file it references. Dashed arrows are tentative. "
                      "Use the Dependency Finder table to filter incoming or outgoing links; the map is read-only.")
        note.setWordWrap(True)
        layout.addWidget(note)
        self.graph_view = QGraphicsView(self._scene(), self)
        self.graph_view.setAccessibleName("Visual file dependency diagram")
        self.graph_view.setRenderHint(QPainter.Antialiasing)
        self.graph_view.setDragMode(QGraphicsView.ScrollHandDrag)
        layout.addWidget(self.graph_view, 4)
        self.summary = QPlainTextEdit(self)
        self.summary.setReadOnly(True)
        self.summary.setAccessibleName("Text summary of file dependency arrows")
        self.summary.setMaximumHeight(165)
        lines = [f"{Path(edge.consumer_path).name} depends on {Path(edge.target_path).name} [{edge.confidence}]"
                 for edge in report.edges[:200]]
        if len(report.edges) > 200:
            lines.append(f"…plus {len(report.edges) - 200} more references in the main results table/export.")
        if len(report.files) > _MAX_GRAPH_NODES:
            lines.append(f"Diagram limited to {_MAX_GRAPH_NODES} of {len(report.files)} files; the report has the full index.")
        self.summary.setPlainText("\n".join(lines) or "No file-to-file dependencies were detected.")
        layout.addWidget(self.summary, 1)
        buttons = QDialogButtonBox(QDialogButtonBox.Close)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    def _scene(self):
        scene = QGraphicsScene(self)
        paths = self.report.files[:_MAX_GRAPH_NODES]
        positions = {}
        width, height, gap_x, gap_y, columns = 225, 69, 60, 44, 4
        for index, name in enumerate(paths):
            positions[name] = (20 + (index % columns) * (width + gap_x),
                               20 + (index // columns) * (height + gap_y))
        for edge in self.report.edges:
            if edge.consumer_path not in positions or edge.target_path not in positions:
                continue
            sx, sy = positions[edge.consumer_path]
            dx, dy = positions[edge.target_path]
            source_center = (sx + width / 2, sy + height / 2)
            target_center = (dx + width / 2, dy + height / 2)
            vector = (target_center[0] - source_center[0], target_center[1] - source_center[1])
            if hypot(*vector) == 0:
                continue
            # Keep both arrow endpoints outside their rectangles so the
            # target box does not obscure the arrowhead.
            fraction = min((width / 2 + 4) / abs(vector[0]) if vector[0] else 1,
                           (height / 2 + 4) / abs(vector[1]) if vector[1] else 1)
            start = (source_center[0] + vector[0] * fraction,
                     source_center[1] + vector[1] * fraction)
            end = (target_center[0] - vector[0] * fraction,
                   target_center[1] - vector[1] * fraction)
            shape = QPainterPath()
            shape.moveTo(*start)
            control = ((start[0] + end[0]) / 2 + 16, (start[1] + end[1]) / 2 - 22)
            shape.quadTo(*control, *end)
            color = QColor("#2871a3")
            pen = QPen(color, 2)
            if edge.confidence != "exact_path":
                pen.setStyle(Qt.DashLine)
            scene.addPath(shape, pen)
            angle = atan2(end[1] - control[1], end[0] - control[0])
            wings = QPolygonF([
                QPointF(*end),
                QPointF(end[0] - 11 * cos(angle - .45), end[1] - 11 * sin(angle - .45)),
                QPointF(end[0] - 11 * cos(angle + .45), end[1] - 11 * sin(angle + .45)),
            ])
            scene.addPolygon(wings, QPen(color), QBrush(color))
        for file_path, (x, y) in positions.items():
            name = Path(file_path).name
            fill = QBrush(QColor("#e4f0fc"))
            scene.addRect(x, y, width, height, QPen(QColor("#477193")), fill)
            label = scene.addText(name, QFont("", 10, QFont.Bold))
            label.setTextWidth(width - 14)
            label.setPos(x + 7, y + 5)
            label.setToolTip(file_path)
        rows = max(1, (len(paths) + columns - 1) // columns)
        scene.setSceneRect(0, 0, columns * (width + gap_x) + 20,
                           rows * (height + gap_y) + 20)
        return scene
