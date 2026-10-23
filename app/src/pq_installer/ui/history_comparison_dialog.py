from PySide6.QtCore import Qt
from PySide6.QtWidgets import (QDialog,QVBoxLayout,QHBoxLayout,QLabel,QTableWidget,
    QTableWidgetItem,QPlainTextEdit,QPushButton,QAbstractItemView)


class HistoryComparisonDialog(QDialog):
    def __init__(self,before,after,changes,parent=None):
        super().__init__(parent)
        self.before=before;self.after=after;self.changes=tuple(changes)
        self.setWindowTitle('Workbook version comparison');self.resize(980,620)
        layout=QVBoxLayout(self)
        self.summary_label=QLabel(self._summary());self.summary_label.setWordWrap(True);layout.addWidget(self.summary_label)
        self.change_table=QTableWidget(len(self.changes),3,self)
        self.change_table.setHorizontalHeaderLabels(['Query','Status','Details'])
        self.change_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.change_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.change_table.setSelectionMode(QAbstractItemView.SingleSelection)
        self.change_table.setAccessibleName('Query changes between workbook snapshots')
        self.change_table.setCornerButtonEnabled(False)
        for row,change in enumerate(self.changes):
            for col,value in enumerate((change.name,change.status,change.reason)):
                self.change_table.setItem(row,col,QTableWidgetItem(value))
        self.change_table.horizontalHeader().setStretchLastSection(True)
        layout.addWidget(self.change_table,2)
        self.diff_view=QPlainTextEdit(self);self.diff_view.setReadOnly(True)
        self.diff_view.setAccessibleName('Selected query formula diff')
        self.diff_view.setPlaceholderText('Select a query to inspect its exact M formula diff.')
        layout.addWidget(self.diff_view,3)
        close=QPushButton('Close',self);close.setAccessibleName('Close workbook version comparison');close.clicked.connect(self.accept)
        row=QHBoxLayout();row.addStretch();row.addWidget(close);layout.addLayout(row)
        self.change_table.currentCellChanged.connect(self._show_diff)
        if self.changes:self.change_table.setCurrentCell(0,0)
        else:self.diff_view.setPlainText('No query definitions were found in either snapshot.')

    def _summary(self):
        counts={}
        for change in self.changes:counts[change.status]=counts.get(change.status,0)+1
        detail=' · '.join(f'{counts[key]} {key.lower()}' for key in
            ('ADDED','MODIFIED','REMOVED','UNCHANGED','CONFLICT','UNKNOWN') if counts.get(key))
        before=self.before.get('occurred_at','unknown time')
        after=self.after.get('occurred_at','unknown time')
        before_name=self.before.get('workbook_name') or self.before.get('workbook_path','earlier workbook')
        after_name=self.after.get('workbook_name') or self.after.get('workbook_path','later workbook')
        return f"{before_name} ({before}) → {after_name} ({after}). {detail or 'No query changes.'} Comparison opens both snapshots read-only and never refreshes external data."

    def _show_diff(self,row,*_):
        if row<0 or row>=len(self.changes):self.diff_view.clear();return
        change=self.changes[row]
        self.diff_view.setPlainText(change.diff or change.reason)
