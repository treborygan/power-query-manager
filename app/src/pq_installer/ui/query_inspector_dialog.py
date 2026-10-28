from pathlib import Path
import shutil
from PySide6.QtCore import Qt
from PySide6.QtGui import QKeySequence, QShortcut
from PySide6.QtWidgets import (QDialog,QVBoxLayout,QHBoxLayout,QSplitter,QGroupBox,
                               QPushButton,QFileDialog,QMessageBox,QLabel,QInputDialog,QLineEdit)
from pq_installer.ui.code_editor import CodeEditor
from pq_installer.core.manager import QueryInfo,diff_m_code
from pq_installer.core.models import QuerySource
from pq_installer.core.parameters import inspect_parameter,edit_parameter
from pq_installer.ui.window_sizing import resize_dialog_to_parent


class QueryInspectorDialog(QDialog):
    def __init__(self,source:QuerySource,existing:QueryInfo|None,parent=None,on_source_saved=None):
        super().__init__(parent)
        self.source=source;self.existing=existing;self._on_source_saved=on_source_saved
        self.setWindowTitle(f'Query inspector — {source.name}')
        self._original_bytes=Path(source.path).read_bytes()
        try:self._source_text=self._original_bytes.decode('utf-8-sig')
        except UnicodeDecodeError as exc:raise ValueError('Query source must be UTF-8 before it can be inspected or edited.') from exc
        root=QVBoxLayout(self);header=QLabel(f'Source file: {source.path}');header.setTextInteractionFlags(Qt.TextSelectableByMouse);root.addWidget(header)
        search_row=QHBoxLayout();self.search_edit=QLineEdit(self);self.search_edit.setPlaceholderText('Search source, workbook, or diff…');self.search_edit.setAccessibleName('Search M expressions and diff')
        search_row.addWidget(self.search_edit,1)
        for label,target in (('Find in source','source'),('Find in workbook','workbook'),('Find in diff','diff')):
            button=QPushButton(label,self);button.setAccessibleName(label);button.clicked.connect(lambda _=False,name=target:self._find_in(name));search_row.addWidget(button)
        root.addLayout(search_row);self.search_edit.returnPressed.connect(lambda:self._find_in('source'))
        self._find_shortcut=QShortcut(QKeySequence.Find,self);self._find_shortcut.activated.connect(self.search_edit.setFocus)
        panes=QSplitter(Qt.Horizontal,self)
        source_box=QGroupBox('Source M expression');source_layout=QVBoxLayout(source_box)
        self.source_editor=CodeEditor(source_box,mode='m');self.source_editor.setAccessibleName('Editable source M expression');self.source_editor.setTabChangesFocus(True);self.source_editor.setPlainText(self._source_text);self.source_editor.document().setModified(False);source_layout.addWidget(self.source_editor)
        workbook_box=QGroupBox('Workbook M expression');workbook_layout=QVBoxLayout(workbook_box)
        self.workbook_view=CodeEditor(workbook_box,mode='m');self.workbook_view.setAccessibleName('Read-only workbook M expression');self.workbook_view.setTabChangesFocus(True);self.workbook_view.setReadOnly(True)
        self.workbook_view.setPlainText(existing.formula if existing and existing.formula is not None else '[Formula unavailable; comparison is unknown]');workbook_layout.addWidget(self.workbook_view)
        panes.addWidget(source_box);panes.addWidget(workbook_box);root.addWidget(panes,3)
        diff_box=QGroupBox('Exact line diff (workbook → source)');diff_layout=QVBoxLayout(diff_box)
        self.diff_view=CodeEditor(diff_box,mode='diff');self.diff_view.setAccessibleName('Line-by-line workbook and source M diff');self.diff_view.setTabChangesFocus(True);self.diff_view.setReadOnly(True)
        self.diff_view.setPlainText(diff_m_code(self._source_text,existing.formula if existing else None,source_name=str(source.path),workbook_name=existing.name if existing else 'Workbook'));diff_layout.addWidget(self.diff_view);root.addWidget(diff_box,2)
        buttons=QHBoxLayout();buttons.addStretch();self.parameter_button=QPushButton('Edit declared parameter…',self);self.parameter_button.setAccessibleName('Edit supported declared M parameter');self.parameter_button.clicked.connect(self._edit_parameter);buttons.addWidget(self.parameter_button)
        self.save_button=QPushButton('Save edited source copy…',self);self.save_button.setAccessibleName('Save edited M as a separate source copy');self.save_button.clicked.connect(self._choose_save_copy);buttons.addWidget(self.save_button)
        close=QPushButton('Close',self);close.setAccessibleName('Close query inspector');close.clicked.connect(self.close);buttons.addWidget(close);root.addLayout(buttons)
        self.source_editor.textChanged.connect(self._refresh_diff);self._refresh_parameter_control()
        resize_dialog_to_parent(self,1100,700)

    def _find_in(self,target):
        editors={'source':self.source_editor,'workbook':self.workbook_view,'diff':self.diff_view}
        editor=editors[target];text=self.search_edit.text()
        if not text:return
        if not editor.find_text(text):
            self.search_edit.setToolTip(f'No match in {target}.')
        else:self.search_edit.setToolTip(f'Found in {target}.')

    def _refresh_parameter_control(self):
        parameter=inspect_parameter(self.source_editor.toPlainText())
        self.parameter_button.setEnabled(parameter is not None)
        self.parameter_button.setToolTip(
            f'Declared {parameter.kind} literal parameter' if parameter else
            'Only a simple literal with IsParameterQuery=true and a supported Text, Number, or Logical type can be edited here.'
        )

    def _edit_parameter(self):
        parameter=inspect_parameter(self.source_editor.toPlainText())
        if parameter is None:
            self._refresh_parameter_control();return
        current=('true' if parameter.value else 'false') if parameter.kind=='Logical' else str(parameter.value)
        value,accepted=QInputDialog.getText(self,'Edit declared parameter',
            f'New {parameter.kind} value for {self.source.name}:',text=current)
        if not accepted:return
        try:edited=edit_parameter(self.source_editor.toPlainText(),value)
        except ValueError as exc:
            QMessageBox.warning(self,'Invalid parameter value',str(exc));return
        self.source_editor.setPlainText(edited);self.source_editor.document().setModified(True)
        self._refresh_diff()

    def _refresh_diff(self):
        old=self.existing.formula if self.existing else None
        self.diff_view.setPlainText(diff_m_code(self.source_editor.toPlainText(),old,source_name=str(self.source.path),workbook_name=self.existing.name if self.existing else 'Workbook'))
        self.setWindowTitle(f'Query inspector — {self.source.name}' + (' *' if self.source_editor.document().isModified() else ''))

    def save_copy(self,destination):
        destination=Path(destination)
        if destination.exists():raise FileExistsError(f'Refusing to overwrite an existing source file: {destination}')
        destination.parent.mkdir(parents=True,exist_ok=True)
        if self.source_editor.document().isModified():
            text=self.source_editor.toPlainText().replace('\r\n','\n').replace('\r','\n')
            newline='\r\n' if b'\r\n' in self._original_bytes else '\n'
            text=text.replace('\n',newline)
            bom=b'\xef\xbb\xbf' if self._original_bytes.startswith(b'\xef\xbb\xbf') else b''
            data=bom+text.encode('utf-8')
            with destination.open('xb') as stream:stream.write(data)
            self._original_bytes=data
        else:
            shutil.copy2(self.source.path,destination)
            self._original_bytes=destination.read_bytes()
        self.source_editor.document().setModified(False);self.source.path=destination
        if self._on_source_saved:self._on_source_saved(destination)
        return destination

    def _choose_save_copy(self):
        default=str(Path(self.source.path).with_name(Path(self.source.path).stem+'_edited'+Path(self.source.path).suffix))
        path,_=QFileDialog.getSaveFileName(self,'Save edited source copy',default,'Power Query sources (*.m *.pq)')
        if not path:return
        try:self.save_copy(path)
        except (OSError,FileExistsError) as exc:QMessageBox.warning(self,'Cannot save source copy',str(exc));return
        QMessageBox.information(self,'Source copy saved',f'Saved a separate source copy to:\n{path}\n\nThe original source file was not changed.')

    def closeEvent(self,event):
        if not self.source_editor.document().isModified():
            event.accept();return
        choice=QMessageBox.warning(self,'Unsaved M edits','Save a separate source copy before closing?',
                                   QMessageBox.Save|QMessageBox.Discard|QMessageBox.Cancel,QMessageBox.Save)
        if choice==QMessageBox.Discard:
            event.accept()
        elif choice==QMessageBox.Save:
            self._choose_save_copy()
            event.accept() if not self.source_editor.document().isModified() else event.ignore()
        else:
            event.ignore()
