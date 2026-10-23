from PySide6.QtCore import Signal
from PySide6.QtWidgets import QTableView


class SourceTableView(QTableView):
    filesDropped=Signal(list)
    def __init__(self,parent=None):
        super().__init__(parent);self.setAcceptDrops(True);self.viewport().setAcceptDrops(True)
    def dragEnterEvent(self,event):
        if event.mimeData().hasUrls():
            paths=[url.toLocalFile() for url in event.mimeData().urls() if url.isLocalFile()]
            if any(path.casefold().endswith(('.m','.pq','.zip')) for path in paths):
                event.acceptProposedAction();return
        super().dragEnterEvent(event)
    def dragMoveEvent(self,event):
        if event.mimeData().hasUrls():event.acceptProposedAction();return
        super().dragMoveEvent(event)
    def dropEvent(self,event):
        if event.mimeData().hasUrls():
            paths=[url.toLocalFile() for url in event.mimeData().urls() if url.isLocalFile()]
            self.filesDropped.emit(paths);event.acceptProposedAction();return
        super().dropEvent(event)
