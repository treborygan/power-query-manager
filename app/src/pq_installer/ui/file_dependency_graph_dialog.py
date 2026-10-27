"""Vector advisory file graph: readable, zoomable, with the full text index retained."""
from __future__ import annotations

from collections import Counter
from math import atan2, cos, hypot, sin
from pathlib import Path

from PySide6.QtCore import QPointF, Qt
from PySide6.QtGui import QBrush, QColor, QFont, QPainter, QPainterPath, QPen, QPolygonF
from PySide6.QtWidgets import (
    QDialog, QDialogButtonBox, QGraphicsItem, QGraphicsScene, QGraphicsView,
    QHBoxLayout, QLabel, QPlainTextEdit, QPushButton, QVBoxLayout,
)

from pq_installer.ui import apply_vector_theme, vector_colors


_MAX_GRAPH_NODES = 80  # the complete index always remains in the report/CSV/JSON


class _VectorGraphView(QGraphicsView):
    """User-controlled zoom; hand-drag pan remains available in QGraphicsView."""

    def wheelEvent(self, event):
        step = event.angleDelta().y()
        if not step:
            return super().wheelEvent(event)
        factor = 1.12 if step > 0 else 1 / 1.12
        scale = self.transform().m11()
        if 0.30 <= scale * factor <= 3.5:
            self.scale(factor, factor)
        event.accept()


class FileDependencyGraphDialog(QDialog):
    def __init__(self, report, parent=None):
        super().__init__(parent)
        self.report = report
        dark = getattr(parent, "dark_theme_enabled",
                       getattr(getattr(parent, "parent", lambda: None)(), "dark_theme_enabled", True))
        self._colors = vector_colors(dark)
        self.setWindowTitle("File dependencies — advisory map")
        self.resize(1180, 790)
        self.setMinimumSize(680, 500)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 16, 20, 18)
        layout.setSpacing(10)

        eyebrow = QLabel("DEPENDENCY INTELLIGENCE    /    GRAPH VIEW")
        eyebrow.setObjectName("vectorEyebrow")
        layout.addWidget(eyebrow)
        title_row = QHBoxLayout()
        title = QLabel("FILE DEPENDENCY NETWORK")
        title.setObjectName("vectorHero")
        title_row.addWidget(title)
        title_row.addStretch()
        self.fit_button = QPushButton("Fit graph")
        self.fit_button.setAccessibleName("Fit dependency diagram to viewport")
        self.fit_button.clicked.connect(self._fit_graph)
        self.reset_button = QPushButton("100%")
        self.reset_button.setAccessibleName("Reset graph zoom to 100 percent")
        self.reset_button.clicked.connect(self._reset_zoom)
        title_row.addWidget(self.fit_button)
        title_row.addWidget(self.reset_button)
        layout.addLayout(title_row)

        note = QLabel(
            "Arrows point FROM the referencing file TO its dependency. "
            "Dashed amber arrows are tentative. Click a node to inspect its filename; "
            "scroll to zoom, drag the canvas to pan. Full results remain in the Finder table and exports."
        )
        note.setObjectName("secondaryText")
        note.setWordWrap(True)
        layout.addWidget(note)

        self.selection_status = QLabel(
            f"{min(len(report.files), _MAX_GRAPH_NODES)} visible files  /  "
            f"{len(report.edges)} detected relationships    •    "
            "LIME = EXACT    /    AMBER DASH = TENTATIVE"
        )
        self.selection_status.setObjectName("vectorStatusLabel")
        self.selection_status.setAccessibleName("Graph status and selected file")
        layout.addWidget(self.selection_status)

        self.graph_view = _VectorGraphView(self._scene(), self)
        self.graph_view.setAccessibleName("Visual file dependency diagram")
        self.graph_view.setRenderHint(QPainter.Antialiasing)
        self.graph_view.setRenderHint(QPainter.TextAntialiasing)
        self.graph_view.setDragMode(QGraphicsView.ScrollHandDrag)
        self.graph_view.setTransformationAnchor(QGraphicsView.AnchorUnderMouse)
        self.graph_view.setResizeAnchor(QGraphicsView.AnchorViewCenter)
        self.graph_view.setStyleSheet("QGraphicsView { border: 1px solid #30434a; border-radius: 5px; }")
        self.graph_view.scene().selectionChanged.connect(self._selection_changed)
        layout.addWidget(self.graph_view, 4)

        summary_label = QLabel("EDGE INDEX    /    ACCESSIBLE TEXT VIEW")
        summary_label.setObjectName("vectorEyebrow")
        layout.addWidget(summary_label)
        self.summary = QPlainTextEdit(self)
        self.summary.setReadOnly(True)
        self.summary.setAccessibleName("Text summary of file dependency arrows")
        self.summary.setMaximumHeight(165)
        grouped = Counter((edge.consumer_path, edge.target_path, edge.confidence)
                          for edge in report.edges)
        lines = [
            f"{Path(consumer).name} depends on {Path(target).name} [{confidence}]"
            + (f"  ×{count}" if count > 1 else "")
            for (consumer, target, confidence), count in list(grouped.items())[:200]
        ]
        if len(grouped) > 200:
            lines.append(
                f"…plus {len(grouped) - 200} distinct relationships in the main results table/export."
            )
        if len(report.edges) > len(grouped):
            lines.append(f"{len(report.edges)} total source/query records across "
                         f"{len(grouped)} distinct file relationships.")
        if len(report.files) > _MAX_GRAPH_NODES:
            lines.append(
                f"Diagram limited to {_MAX_GRAPH_NODES} of {len(report.files)} files; "
                "the report has the full index."
            )
        self.summary.setPlainText("\n".join(lines) or "No file-to-file dependencies were detected.")
        layout.addWidget(self.summary, 1)
        buttons = QDialogButtonBox(QDialogButtonBox.Close)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)
        apply_vector_theme(self, dark)

    def _fit_graph(self):
        rect = self.graph_view.scene().sceneRect()
        if not rect.isEmpty():
            self.graph_view.fitInView(rect, Qt.KeepAspectRatio)

    def _reset_zoom(self):
        self.graph_view.resetTransform()

    def _selection_changed(self):
        nodes = self.graph_view.scene().selectedItems()
        if nodes:
            file_path = nodes[0].data(0)
            if file_path:
                self.selection_status.setText("SELECTED  /  " + Path(str(file_path)).name)
                self.selection_status.setToolTip(str(file_path))
                return
        self.selection_status.setText(
            f"{min(len(self.report.files), _MAX_GRAPH_NODES)} visible files  /  "
            f"{len(self.report.edges)} detected relationships    •    "
            "LIME = EXACT    /    AMBER DASH = TENTATIVE"
        )
        self.selection_status.setToolTip("")

    def _scene(self):
        colors = self._colors
        scene = QGraphicsScene(self)
        scene.setBackgroundBrush(QBrush(QColor("#0c171a" if colors["window"] == "#080e12" else "#edf5f1")))
        paths = self.report.files[:_MAX_GRAPH_NODES]
        inbound = Counter(edge.target_path for edge in self.report.edges)
        outbound = Counter(edge.consumer_path for edge in self.report.edges)
        # Sources with the most outgoing references lead; alphabetical order breaks ties.
        ordered = sorted(paths, key=lambda name: (-outbound[name], -inbound[name], Path(name).name.casefold()))
        positions = {}
        width, height, gap_x, gap_y, columns = 246, 82, 82, 62, 4
        for index, name in enumerate(ordered):
            positions[name] = (
                28 + (index % columns) * (width + gap_x),
                26 + (index // columns) * (height + gap_y),
            )
        # A workbook can contain many formulas referencing the same file; render
        # one arrow per path/confidence instead of obscuring nodes with duplicates.
        unique_edges = []
        drawn = set()
        for edge in self.report.edges:
            identity = (edge.consumer_path, edge.target_path, edge.confidence)
            if identity not in drawn:
                drawn.add(identity)
                unique_edges.append(edge)
        for edge_index, edge in enumerate(unique_edges):
            if edge.consumer_path not in positions or edge.target_path not in positions:
                continue
            sx, sy = positions[edge.consumer_path]
            dx, dy = positions[edge.target_path]
            source_center = (sx + width / 2, sy + height / 2)
            target_center = (dx + width / 2, dy + height / 2)
            vector = (target_center[0] - source_center[0], target_center[1] - source_center[1])
            if hypot(*vector) == 0:
                continue
            fraction = min(
                (width / 2 + 6) / abs(vector[0]) if vector[0] else 1,
                (height / 2 + 6) / abs(vector[1]) if vector[1] else 1,
            )
            start = (
                source_center[0] + vector[0] * fraction,
                source_center[1] + vector[1] * fraction,
            )
            end = (
                target_center[0] - vector[0] * fraction,
                target_center[1] - vector[1] * fraction,
            )
            shape = QPainterPath()
            shape.moveTo(*start)
            bend = (16 + (edge_index % 3) * 10) * (1 if edge_index % 2 else -1)
            control = ((start[0] + end[0]) / 2 + bend, (start[1] + end[1]) / 2 - bend)
            shape.quadTo(*control, *end)
            exact = edge.confidence == "exact_path"
            dark = colors["window"] == "#080e12"
            link_color = QColor(("#c8f477" if exact else "#ffc88b") if dark
                                else ("#386c20" if exact else "#9c4e13"))
            pen = QPen(link_color, 2.2)
            if not exact:
                pen.setStyle(Qt.DashLine)
            scene.addPath(shape, pen)
            angle = atan2(end[1] - control[1], end[0] - control[0])
            wings = QPolygonF([
                QPointF(*end),
                QPointF(end[0] - 12 * cos(angle - .48), end[1] - 12 * sin(angle - .48)),
                QPointF(end[0] - 12 * cos(angle + .48), end[1] - 12 * sin(angle + .48)),
            ])
            scene.addPolygon(wings, QPen(link_color), QBrush(link_color))
        for file_path, (x, y) in positions.items():
            name = Path(file_path).name
            is_access = Path(name).suffix.casefold() in {".accdb", ".mdb"}
            accent = QColor("#81e4d3" if is_access else "#cff66d")
            fill = QBrush(QColor("#17252b"))
            path = QPainterPath()
            path.addRoundedRect(0, 0, width, height, 7, 7)
            node = scene.addPath(path, QPen(QColor("#48646b"), 1.2), fill)
            node.setPos(x, y)
            node.setFlag(QGraphicsItem.ItemIsSelectable, True)
            node.setData(0, str(file_path))
            node.setToolTip(str(file_path))
            stripe = scene.addRect(0, 0, 4, height, QPen(Qt.NoPen), QBrush(accent))
            stripe.setParentItem(node)
            stripe.setPos(0, 0)
            stripe.setAcceptedMouseButtons(Qt.NoButton)
            name_label = scene.addText(name, QFont("Segoe UI", 10, QFont.Bold))
            name_label.setDefaultTextColor(QColor("#f0f9f8"))
            name_label.setTextWidth(width - 26)
            name_label.setParentItem(node)
            name_label.setPos(12, 8)
            name_label.setAcceptedMouseButtons(Qt.NoButton)
            category = "ACCESS DATABASE" if is_access else "EXCEL WORKBOOK"
            detail = (
                f"{category}   ·   {outbound[file_path]} outgoing  /  {inbound[file_path]} incoming"
            )
            sublabel = scene.addText(detail, QFont("Consolas", 8))
            sublabel.setDefaultTextColor(QColor("#b3c8c9"))
            sublabel.setTextWidth(width - 20)
            sublabel.setParentItem(node)
            sublabel.setPos(12, 48)
            sublabel.setAcceptedMouseButtons(Qt.NoButton)
        rows = max(1, (len(paths) + columns - 1) // columns)
        scene.setSceneRect(
            0, 0, columns * (width + gap_x) + 20,
            rows * (height + gap_y) + 20,
        )
        return scene
