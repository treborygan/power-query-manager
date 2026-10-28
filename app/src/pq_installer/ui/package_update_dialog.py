from pathlib import Path

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout, QLabel,
                               QTableWidget, QTableWidgetItem, QPlainTextEdit,
                               QSplitter, QPushButton, QHeaderView)
from pq_installer.ui import apply_vector_theme


class QueryPackageUpdateDialog(QDialog):
    """Read-only package-to-package change preview; it never applies workbook actions."""

    _STATUS_LABELS={
        'ADDED':'Added', 'MODIFIED':'Modified', 'UNCHANGED':'Unchanged',
        'REMOVED':'Removed', 'CONFLICT':'Name conflict',
    }

    def __init__(self, base_package, updated_package, changes, parent=None):
        super().__init__(parent)
        self.changes=tuple(changes)
        self.setWindowTitle('Query package update preview')
        apply_vector_theme(self, getattr(parent, "dark_theme_enabled", True))
        self.resize(920,620)
        self.setMinimumSize(720,450)
        root=QVBoxLayout(self)
        sources=QLabel(f'Current: {Path(base_package).name}    →    Updated: {Path(updated_package).name}',self)
        sources.setTextInteractionFlags(Qt.TextSelectableByMouse)
        root.addWidget(sources)

        counts={status:sum(row.status==status for row in self.changes)
                for status in ('ADDED','MODIFIED','UNCHANGED','REMOVED','CONFLICT')}
        if not self.changes:
            summary='No package changes. The validated packages contain no query definitions.'
        else:
            summary=(f"{counts['ADDED']} added · {counts['MODIFIED']} modified · "
                     f"{counts['UNCHANGED']} unchanged · {counts['REMOVED']} removed")
            if counts['CONFLICT']:
                summary+=f" · {counts['CONFLICT']} name conflict(s)"
        self.summary_label=QLabel(summary,self)
        self.summary_label.setAccessibleName('Package update summary')
        root.addWidget(self.summary_label)

        panes=QSplitter(Qt.Orientation.Vertical,self)
        self.change_table=QTableWidget(len(self.changes),3,panes)
        self.change_table.setAccessibleName('Query package changes')
        self.change_table.setCornerButtonEnabled(False)
        self.change_table.setHorizontalHeaderLabels(['Query','Change','Previous name'])
        self.change_table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.change_table.setSelectionMode(QTableWidget.SelectionMode.SingleSelection)
        self.change_table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.change_table.setAlternatingRowColors(True)
        header=self.change_table.horizontalHeader()
        header.setSectionResizeMode(0,QHeaderView.ResizeMode.Stretch)
        header.setSectionResizeMode(1,QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(2,QHeaderView.ResizeMode.Stretch)
        for row,change in enumerate(self.changes):
            for column,value in enumerate((change.name,self._STATUS_LABELS.get(change.status,change.status),change.previous_name or '')):
                item=QTableWidgetItem(value)
                if column==0:item.setData(Qt.ItemDataRole.UserRole,row)
                self.change_table.setItem(row,column,item)
        self.diff_view=QPlainTextEdit(panes)
        self.diff_view.setAccessibleName('Selected query package diff')
        self.diff_view.setReadOnly(True)
        self.diff_view.setPlaceholderText('Select a package change to inspect its exact M diff.')
        panes.addWidget(self.change_table);panes.addWidget(self.diff_view)
        panes.setStretchFactor(0,1);panes.setStretchFactor(1,2)
        root.addWidget(panes,1)

        self.safety_note=QLabel('This is a read-only package comparison. Removed entries are listed for review; no workbook queries are deleted or changed.',self)
        self.safety_note.setWordWrap(True)
        self.safety_note.setAccessibleName('No workbook changes safety notice')
        root.addWidget(self.safety_note)
        buttons=QHBoxLayout();buttons.addStretch()
        close=QPushButton('Close',self);close.setAccessibleName('Close query package comparison');close.clicked.connect(self.accept);close.setDefault(True)
        buttons.addWidget(close);root.addLayout(buttons)
        self.change_table.currentCellChanged.connect(self._show_change)
        if self.changes:self.change_table.setCurrentCell(0,0)

    def _show_change(self,row,*_args):
        if not 0<=row<len(self.changes):
            self.diff_view.clear();return
        change=self.changes[row]
        header=f"{change.name} — {self._STATUS_LABELS.get(change.status,change.status)}"
        if change.previous_name and change.previous_name!=change.name:
            header+=f" (previous name: {change.previous_name})"
        self.diff_view.setPlainText(header+'\n\n'+(change.diff or 'No M changes.'))
