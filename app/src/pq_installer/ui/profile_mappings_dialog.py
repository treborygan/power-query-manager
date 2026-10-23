from PySide6.QtCore import Qt
from PySide6.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout, QLabel,
                               QTabWidget, QWidget, QTableWidget, QTableWidgetItem,
                               QPushButton, QHeaderView, QMessageBox)


class ProfileMappingsDialog(QDialog):
    """Edit profile-scoped source path and recognized M source-location mappings."""

    columns=('Key','From','To','Enabled')
    checked_state=Qt.CheckState.Checked
    unchecked_state=Qt.CheckState.Unchecked

    def __init__(self,path_mappings=(),environment_mappings=(),parent=None):
        super().__init__(parent)
        self.setWindowTitle('Edit source environment mappings')
        self.resize(900,560)
        self.setMinimumSize(700,420)
        root=QVBoxLayout(self)
        intro=QLabel('Edit mappings stored in the import profile. Changes stay in this dialog until you choose Save mappings.',self)
        intro.setWordWrap(True);root.addWidget(intro)
        self.tabs=QTabWidget(self);self.tabs.setAccessibleName('Profile mapping categories');root.addWidget(self.tabs,1)

        path_page=QWidget(self.tabs);path_layout=QVBoxLayout(path_page)
        self.path_help=QLabel('Source-file paths only: map profile source roots to the locations used in this environment. Loading the profile shows a path preview before applying.',path_page)
        self.path_help.setWordWrap(True);path_layout.addWidget(self.path_help)
        self.path_table=self._create_table(path_mappings,path_page,'Profile source path mappings')
        path_layout.addWidget(self.path_table,1)
        path_actions=QHBoxLayout()
        add_path=QPushButton('Add path mapping',path_page);add_path.setAccessibleName('Add source path mapping');add_path.clicked.connect(self.add_path_mapping)
        self.select_all_path_button=QPushButton('Select all path mappings',path_page);self.select_all_path_button.setObjectName('selectAllPathMappings');self.select_all_path_button.setAccessibleName('Select all source path mappings');self.select_all_path_button.clicked.connect(self.path_table.selectAll)
        remove_path=QPushButton('Remove selected mapping',path_page);remove_path.setAccessibleName('Remove selected source path mappings');remove_path.clicked.connect(lambda:self._remove_selected(self.path_table))
        path_actions.addWidget(add_path);path_actions.addWidget(self.select_all_path_button);path_actions.addWidget(remove_path);path_actions.addStretch();path_layout.addLayout(path_actions)
        self.tabs.addTab(path_page,'Source paths')

        environment_page=QWidget(self.tabs);environment_layout=QVBoxLayout(environment_page)
        self.environment_help=QLabel('M source locations only: rewrite recognized literal paths in supported M source functions. Loading the profile shows exact formula diffs and requires approval. This is not a full M parser.',environment_page)
        self.environment_help.setWordWrap(True);environment_layout.addWidget(self.environment_help)
        self.environment_table=self._create_table(environment_mappings,environment_page,'M source location mappings')
        environment_layout.addWidget(self.environment_table,1)
        environment_actions=QHBoxLayout()
        add_environment=QPushButton('Add M source mapping',environment_page);add_environment.setAccessibleName('Add M source location mapping');add_environment.clicked.connect(self.add_environment_mapping)
        self.select_all_environment_button=QPushButton('Select all M source mappings',environment_page);self.select_all_environment_button.setObjectName('selectAllEnvironmentMappings');self.select_all_environment_button.setAccessibleName('Select all M source location mappings');self.select_all_environment_button.clicked.connect(self.environment_table.selectAll)
        remove_environment=QPushButton('Remove selected mapping',environment_page);remove_environment.setAccessibleName('Remove selected M source location mappings');remove_environment.clicked.connect(lambda:self._remove_selected(self.environment_table))
        environment_actions.addWidget(add_environment);environment_actions.addWidget(self.select_all_environment_button);environment_actions.addWidget(remove_environment);environment_actions.addStretch();environment_layout.addLayout(environment_actions)
        self.tabs.addTab(environment_page,'M source locations')

        buttons=QHBoxLayout();buttons.addStretch()
        cancel=QPushButton('Cancel',self);cancel.setAccessibleName('Cancel mapping edits');cancel.clicked.connect(self.reject)
        self.save_button=QPushButton('Save mappings',self);self.save_button.setAccessibleName('Save profile mappings');self.save_button.setDefault(True);self.save_button.clicked.connect(self._accept_valid)
        buttons.addWidget(cancel);buttons.addWidget(self.save_button);root.addLayout(buttons)

    @classmethod
    def _create_table(cls,mappings,parent,accessible_name):
        table=QTableWidget(0,4,parent)
        table.setAccessibleName(accessible_name)
        table.setCornerButtonEnabled(False)
        table.setHorizontalHeaderLabels(cls.columns)
        table.setAlternatingRowColors(True)
        table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        table.setSelectionMode(QTableWidget.SelectionMode.ExtendedSelection)
        table.setEditTriggers(QTableWidget.EditTrigger.DoubleClicked|QTableWidget.EditTrigger.EditKeyPressed|QTableWidget.EditTrigger.AnyKeyPressed)
        header=table.horizontalHeader()
        for column in (0,1,2):header.setSectionResizeMode(column,QHeaderView.ResizeMode.Stretch)
        header.setSectionResizeMode(3,QHeaderView.ResizeMode.ResizeToContents)
        for mapping in mappings:
            cls._append_row(table,mapping)
        return table

    @classmethod
    def _append_row(cls,table,mapping=None):
        mapping=dict(mapping or {})
        row=table.rowCount();table.insertRow(row)
        for column,key in enumerate(('key','from','to')):
            table.setItem(row,column,QTableWidgetItem(str(mapping.get(key,''))))
        enabled=QTableWidgetItem('')
        enabled.setFlags(Qt.ItemFlag.ItemIsEnabled|Qt.ItemFlag.ItemIsUserCheckable|Qt.ItemFlag.ItemIsSelectable)
        enabled.setCheckState(cls.checked_state if mapping.get('enabled',True) else cls.unchecked_state)
        table.setItem(row,3,enabled)
        table.setCurrentCell(row,0)

    def add_path_mapping(self):
        self._append_row(self.path_table)
        self.tabs.setCurrentIndex(0)
        self.path_table.setFocus()

    def add_environment_mapping(self):
        self._append_row(self.environment_table)
        self.tabs.setCurrentIndex(1)
        self.environment_table.setFocus()

    @staticmethod
    def _remove_selected(table):
        rows=sorted({index.row() for index in table.selectionModel().selectedRows()},reverse=True)
        if not rows and table.currentRow()>=0:rows=[table.currentRow()]
        for row in rows:table.removeRow(row)

    @classmethod
    def _read_table(cls,table,label):
        mappings=[];keys=set()
        for row in range(table.rowCount()):
            values=[table.item(row,column).text() if table.item(row,column) else '' for column in range(3)]
            key=values[0].strip()
            if not key or not values[1].strip() or not values[2].strip():
                raise ValueError(f'{label} row {row+1} needs a key, from value, and to value.')
            if key.casefold() in keys:
                raise ValueError(f'{label} contains duplicate mapping key {key!r}.')
            keys.add(key.casefold())
            enabled=table.item(row,3).checkState()==cls.checked_state
            mappings.append({'key':key,'from':values[1],'to':values[2],'enabled':enabled})
        return mappings

    def mappings(self):
        """Return validated copies without mutating the caller's profile state."""
        return (self._read_table(self.path_table,'Source path mappings'),
                self._read_table(self.environment_table,'M source location mappings'))

    def _accept_valid(self):
        try:self.mappings()
        except ValueError as exc:
            QMessageBox.warning(self,'Check source mappings',str(exc));return
        self.accept()
