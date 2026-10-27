import os
os.environ.setdefault('QT_QPA_PLATFORM', 'offscreen')

from pathlib import Path
from PySide6.QtCore import Qt, QModelIndex
from PySide6.QtWidgets import QApplication, QMessageBox
from pq_installer.core.models import QuerySource
from pq_installer.ui.query_table_model import QueryTableModel, QuerySortFilterProxyModel
from pq_installer.ui.main_window import MainWindow
from pq_installer.core.manager import QueryInfo

app = QApplication.instance() or QApplication([])


def test_manager_window_uses_query_table_model_view():
    from PySide6.QtWidgets import QTableView
    window=MainWindow()
    assert isinstance(window.table,QTableView)
    assert isinstance(window.table.model(),QuerySortFilterProxyModel)
    assert isinstance(window.table.model().sourceModel(),QueryTableModel)
    window.close()


def test_append_sources_deduplicates_duplicate_paths_within_one_batch(tmp_path):
    path=tmp_path/'same.pq';path.write_text('let x=1 in x')
    first=QuerySource(path,'First')
    duplicate=QuerySource(path,'Second')
    model=QueryTableModel()

    assert model.append_sources([first,duplicate])==1
    assert model.rowCount()==1
    assert model.source_at(0) is first


def test_manager_interactive_widgets_have_accessible_names():
    from PySide6.QtWidgets import (QAbstractButton, QComboBox, QLineEdit,
        QPlainTextEdit, QSpinBox, QTableView)
    window=MainWindow();window.show();app.processEvents()
    interactive=(QAbstractButton,QComboBox,QLineEdit,QPlainTextEdit,QSpinBox,QTableView)
    unnamed=[]
    for widget_type in interactive:
        for widget in window.findChildren(widget_type):
            if widget.isVisibleTo(window) and widget.isEnabled() and not widget.accessibleName().strip():
                parent=widget.parentWidget()
                unnamed.append((widget.metaObject().className(),widget.objectName(),
                    parent.metaObject().className() if parent else None,
                    parent.objectName() if parent else None,
                    widget.text() if hasattr(widget,'text') else ''))
    window.close()
    assert not unnamed, f'Interactive manager controls lack accessible names: {unnamed}'


def test_named_select_all_action_selects_every_query_row(tmp_path):
    from PySide6.QtTest import QTest
    from pq_installer.core.models import QuerySource
    sources=[]
    for name in ('First','Second'):
        path=tmp_path/f'{name}.pq';path.write_text('1')
        sources.append(QuerySource(path,name))
    window=MainWindow();window._merge_sources(sources);window.show();app.processEvents()
    assert window.select_all_button.accessibleName()=='Select all query rows'
    QTest.mouseClick(window.select_all_button,Qt.LeftButton);app.processEvents()
    assert len(window.table.selectionModel().selectedRows())==2
    window.close()


def test_all_skipped_import_completion_is_reported_as_no_changes(monkeypatch, tmp_path):
    from pq_installer.core.models import InstallResult, QueryStatus

    window = MainWindow()
    window.workbook_edit.setText(str(tmp_path / 'report.xlsx'))
    history_events = []
    window._record_history = lambda event, *_args, **_kwargs: history_events.append(event)
    dialogs = []
    monkeypatch.setattr(QMessageBox, 'information',
        lambda *args: dialogs.append(args[1:]))

    window._on_completed({'results': (InstallResult(
        'Q', QueryStatus.SKIPPED, 'Existing query kept unchanged.'),)})

    assert window.statusBar().currentMessage() == 'No changes: all selected queries were skipped'
    assert dialogs and dialogs[0][0] == 'No changes'
    assert 'selected queries were skipped' in dialogs[0][1].casefold()
    assert 'successfully' not in dialogs[0][1].casefold()
    assert history_events == ['import_skipped']
    window.close()


def test_busy_state_disables_inspector_and_view_menu_actions():
    window=MainWindow()
    window._set_busy(True,'Working…')
    assert not window.centralWidget().isEnabled()
    assert not window.inspector_dock.isEnabled()
    assert not window.menuBar().isEnabled()
    window._set_busy(False)
    assert window.centralWidget().isEnabled()
    assert window.inspector_dock.isEnabled()
    assert window.menuBar().isEnabled()
    window.close()


def test_recovery_control_tracks_matching_staged_workbooks(tmp_path):
    window=MainWindow()
    workbook=tmp_path/'report.xlsx';workbook.write_bytes(b'original')
    window.workbook_edit.setText(str(workbook))
    assert not window.recover_stage_btn.isEnabled()
    stage=tmp_path/'.report.pqstage-0123456789abcdef.xlsx';stage.write_bytes(b'staged')
    window._refresh_recovery_state()
    assert window.recover_stage_btn.isEnabled()
    window._set_busy(True,'Working…')
    assert not window.recover_stage_btn.isEnabled()
    window._set_busy(False)
    assert window.recover_stage_btn.isEnabled()
    window.close()


def test_docked_inspector_tracks_selected_source_and_workbook_formula(tmp_path):
    from PySide6.QtCore import QItemSelectionModel
    from PySide6.QtWidgets import QDockWidget, QPlainTextEdit
    window=MainWindow()
    source_path=tmp_path/'sales.m';source_path.write_text('let Source = 1 in Source')
    window._merge_sources([QuerySource(source_path,'Sales')])
    window.inventory=(QueryInfo('Sales','let Source = 2 in Source'),)
    index=window.query_proxy.index(0,QueryTableModel.NAME)
    window.table.selectionModel().select(index,QItemSelectionModel.Select|QItemSelectionModel.Rows)
    window.inspector_toggle.setChecked(True)
    assert isinstance(window.inspector_dock,QDockWidget)
    assert isinstance(window.inspector_source_text,QPlainTextEdit)
    assert 'let Source = 1 in Source' in window.inspector_source_text.toPlainText()
    assert 'let Source = 2 in Source' in window.inspector_workbook_text.toPlainText()
    assert window.inspector_status.text().startswith('Sales')
    window.close()


def test_docked_inspector_reports_unknown_workbook_formula(tmp_path):
    from PySide6.QtCore import QItemSelectionModel
    window=MainWindow();path=tmp_path/'unknown.pq';path.write_text('let x=1 in x')
    window._merge_sources([QuerySource(path,'Unknown')])
    window.inventory=(QueryInfo('Other',None),)
    index=window.query_proxy.index(0,QueryTableModel.NAME)
    window.table.selectionModel().select(index,QItemSelectionModel.Select|QItemSelectionModel.Rows)
    assert 'not present' in window.inspector_status.text().casefold()
    assert 'not present' in window.inspector_workbook_text.toPlainText().casefold()
    window.close()


def test_editing_path_clears_temporary_profile_source_reference(tmp_path):
    original=tmp_path/'original.pq';original.write_text('let x=1 in x')
    mapped=tmp_path/'mapped.pq';mapped.write_text('let x=2 in x')
    replacement=tmp_path/'replacement.pq'
    source=QuerySource(mapped,'Mapped',profile_source_path=original)
    model=QueryTableModel([source])
    assert model.setData(model.index(0,QueryTableModel.PATH),str(replacement),Qt.EditRole)
    assert source.path==replacement and source.profile_source_path is None


def test_manager_history_records_actions_and_exposes_history_controls(tmp_path):
    from PySide6.QtWidgets import QPushButton
    from pq_installer.core.history import HistoryStore
    window=MainWindow()
    window.history_store=HistoryStore(tmp_path/'history.sqlite3')
    assert isinstance(window.history_button,QPushButton)
    window._record_history('profile_saved',{'profile_name':'Morning','formula':'let secret=1 in secret'})
    row=window.history_store.list()[0]
    assert row['event']=='profile_saved'
    assert row['details']=={'profile_name':'Morning'}
    window.close()


def test_successful_operations_link_opted_in_snapshot_to_history(tmp_path,monkeypatch):
    from pq_installer.core.history import HistoryStore
    from pq_installer.core.excel_refresh import RefreshResult
    from pq_installer.core.models import InstallResult,QueryStatus
    monkeypatch.setattr('pq_installer.ui.main_window.QMessageBox.information',lambda *args,**kwargs:QMessageBox.Ok)
    workbook=tmp_path/'report.xlsx';workbook.write_bytes(b'contents')
    store=HistoryStore(tmp_path/'history.sqlite3');snapshot=store.create_snapshot(workbook,event='workbook_snapshot_before_import')
    window=MainWindow();window.history_store=store;window.workbook_edit.setText(str(workbook))
    window._on_completed({'results':(InstallResult('Q',QueryStatus.CREATED,'created'),),
                          'snapshot_id':snapshot['snapshot_id']})
    row=next(row for row in store.list() if row['event']=='import_completed')
    assert row['event']=='import_completed' and row['details']['snapshot_id']==snapshot['snapshot_id']
    window._on_refresh_completed({'results':(RefreshResult('Q','refreshed','done'),),
                                  'snapshot_id':snapshot['snapshot_id']})
    row=next(row for row in store.list() if row['event']=='refresh_completed')
    assert row['event']=='refresh_completed' and row['details']['snapshot_id']==snapshot['snapshot_id']
    window.close()


def test_main_window_bulk_enable_uses_proxy_selection_mapping(tmp_path):
    from PySide6.QtCore import QItemSelectionModel
    window=MainWindow();path=tmp_path/'bulk.m';path.write_text('let x=1 in x')
    window._merge_sources([QuerySource(path,'Bulk',0)])
    proxy_index=window.query_proxy.index(0,QueryTableModel.NAME)
    window.table.selectionModel().select(proxy_index,QItemSelectionModel.Select|QItemSelectionModel.Rows)
    window._set_selected_enabled(False)
    assert window.query_model.source_at(0).enabled is False
    window.close()


def test_dry_run_review_never_starts_excel_worker(tmp_path,monkeypatch):
    from PySide6.QtWidgets import QMessageBox
    monkeypatch.setattr(QMessageBox,'question',lambda *args,**kwargs:QMessageBox.Yes)
    window=MainWindow();workbook=tmp_path/'book.xlsx';workbook.write_bytes(b'disposable')
    source=tmp_path/'q.pq';source.write_text('let x=1 in x')
    window.workbook_edit.setText(str(workbook));window._merge_sources([QuerySource(source,'Q')]);window.dry_run_check.setChecked(True)
    window._install()
    assert window._thread is None and 'DRY RUN' in window.log.toPlainText()
    window.close()


def test_zip_import_sources_are_isolated_and_cleaned_when_window_closes(tmp_path,monkeypatch):
    import zipfile
    from PySide6.QtWidgets import QFileDialog
    archive=tmp_path/'bundle.zip'
    with zipfile.ZipFile(archive,'w') as package:package.writestr('queries/1_one.pq','let x=1 in x')
    monkeypatch.setattr(QFileDialog,'getOpenFileName',lambda *args,**kwargs:(str(archive),''))
    window=MainWindow();window._add_zip()
    assert window.query_model.rowCount()==1
    extracted=window.query_model.source_at(0).path
    assert extracted.is_file() and extracted.parent.parent.name.startswith('pq-import-')
    root=next(iter(window._staged_import_dirs));window.close()
    assert not root.exists()


def test_model_has_stable_identity_and_edits_without_replacing_rows(tmp_path):
    path=tmp_path/'one.m'; path.write_text('let x=1 in x')
    source=QuerySource(path,'One',order=1)
    model=QueryTableModel([source]); stable=model.stable_id_at(0)
    assert model.data(model.index(0,2),Qt.DisplayRole)=='One'
    assert model.setData(model.index(0,2),'Renamed',Qt.EditRole)
    assert source.name=='Renamed' and model.stable_id_at(0)==stable
    assert model.setData(model.index(0,0),Qt.Unchecked,Qt.CheckStateRole)
    assert source.enabled is False


def test_table_edits_support_undo_redo_without_changing_stable_identity(tmp_path):
    path=tmp_path/'query.pq';path.write_text('let x=1 in x')
    source=QuerySource(path,'Before',order=2)
    model=QueryTableModel([source]);stable=model.stable_id_at(0)
    assert model.setData(model.index(0,QueryTableModel.NAME),'After',Qt.EditRole)
    assert source.name=='After' and model.undo_stack.canUndo()
    model.undo_stack.undo()
    assert source.name=='Before' and model.stable_id_at(0)==stable
    model.undo_stack.redo()
    assert source.name=='After' and model.stable_id_at(0)==stable


def test_undo_of_path_edit_restores_profile_source_origin(tmp_path):
    original=tmp_path/'original.pq';original.write_text('original')
    mapped=tmp_path/'mapped.pq';mapped.write_text('mapped')
    replacement=tmp_path/'replacement.pq'
    source=QuerySource(mapped,'Mapped',profile_source_path=original)
    model=QueryTableModel([source])
    assert model.setData(model.index(0,QueryTableModel.PATH),str(replacement),Qt.EditRole)
    assert source.path==replacement and source.profile_source_path is None
    model.undo_stack.undo()
    assert source.path==mapped and source.profile_source_path==original
    model.undo_stack.redo()
    assert source.path==replacement and source.profile_source_path is None


def test_manager_has_undo_redo_actions_and_buttons(tmp_path):
    from PySide6.QtWidgets import QPushButton
    window=MainWindow()
    assert isinstance(window.undo_button,QPushButton)
    assert isinstance(window.redo_button,QPushButton)
    path=tmp_path/'q.pq';path.write_text('1')
    window._merge_sources([QuerySource(path,'Before')])
    window.query_model.setData(window.query_model.index(0,QueryTableModel.NAME),'After',Qt.EditRole)
    window.undo_button.click()
    assert window.query_model.source_at(0).name=='Before'
    window.redo_button.click()
    assert window.query_model.source_at(0).name=='After'
    window.close()


def test_light_dark_theme_toggle_applies_readable_theme_styles():
    window=MainWindow()
    light=window.styleSheet()
    assert 'background: #f4f7fb' in light
    window.theme_toggle.setChecked(True)
    dark=window.styleSheet()
    assert window.dark_theme_enabled and 'background: #1c232b' in dark
    assert 'QLabel#secondaryText { color: #c2ccd8; }' in dark
    window.theme_toggle.setChecked(False)
    assert not window.dark_theme_enabled and 'background: #f4f7fb' in window.styleSheet()
    window.close()


def test_manager_initial_size_and_accessible_names_fit_laptop_layout(tmp_path):
    from PySide6.QtTest import QTest
    from PySide6.QtCore import Qt
    from PySide6.QtWidgets import QScrollArea
    window=MainWindow()
    workbook=tmp_path/'target.xlsx';workbook.write_bytes(b'disposable')
    window.workbook_edit.setText(str(workbook))
    assert isinstance(window.centralWidget(),QScrollArea)
    assert window.workbook_edit.accessibleName()=='Target workbook path'
    assert window.table.accessibleName()=='Query source settings table'
    assert window.search_edit.accessibleName()=='Filter query source rows'
    for size in ((960,600),(1366,768)):
        window.resize(*size);window.show();app.processEvents()
        assert window.width()<=1366 and window.height()<=768
        assert window.minimumWidth()<=window.width() and window.minimumHeight()<=window.height()
        assert window.table.height()>=140
        assert window.centralWidget().horizontalScrollBar().maximum()==0
    assert window.centralWidget().verticalScrollBar().maximum()>0
    window.workbook_edit.setFocus();QTest.keyClick(window.workbook_edit,Qt.Key_Tab);app.processEvents()
    assert window.focusWidget().text()=='Browse…'
    expected=('Open selected workbook in Excel','Inspect selected workbook queries read-only','Probe Excel load capabilities')
    for accessible_name in expected:
        QTest.keyClick(window.focusWidget(),Qt.Key_Tab);app.processEvents()
        assert window.focusWidget().accessibleName()==accessible_name
    assert window.install_btn.accessibleName()=='Review and import selected queries'
    assert window.log.accessibleName()=='Operation log output'
    assert window.log.objectName()=='operationLogOutput'
    assert window.log.accessibleDescription()=='Live read-only progress and result details for the current operation.'
    window.close()


def test_query_table_supports_keyboard_tab_navigation(tmp_path):
    from PySide6.QtTest import QTest
    path=tmp_path/'keyboard.pq';path.write_text('1')
    window=MainWindow();window._merge_sources([QuerySource(path,'Keyboard')])
    window.resize(960,600);window.show();app.processEvents()
    window.table.setFocus()
    first=window.query_proxy.index(0,QueryTableModel.NAME)
    window.table.setCurrentIndex(first)
    QTest.keyClick(window.table,Qt.Key_Tab);app.processEvents()
    assert window.table.hasFocus()
    assert window.table.currentIndex().row()==first.row()
    assert window.table.currentIndex().column()==first.column()+1
    window.close()


def test_bulk_edit_is_one_undo_step_and_restores_every_selected_row(tmp_path):
    from PySide6.QtCore import QItemSelectionModel
    paths=[tmp_path/f'{name}.pq' for name in ('One','Two')]
    for path in paths:path.write_text('1')
    window=MainWindow();window._merge_sources([QuerySource(path,path.stem) for path in paths])
    selection=window.table.selectionModel()
    for row in range(2):
        index=window.query_proxy.index(row,QueryTableModel.NAME)
        selection.select(index,QItemSelectionModel.Select|QItemSelectionModel.Rows)
    before=window.query_model.undo_stack.count()
    window._set_selected_enabled(False)
    assert [window.query_model.source_at(i).enabled for i in range(2)]==[False,False]
    assert window.query_model.undo_stack.count()==before+1
    window.query_model.undo_stack.undo()
    assert [window.query_model.source_at(i).enabled for i in range(2)]==[True,True]
    window.close()


def test_destination_configuration_is_undoable_by_stable_row_identity(tmp_path):
    path=tmp_path/'q.pq';path.write_text('1')
    model=QueryTableModel([QuerySource(path,'Q')]);stable=model.stable_id_at(0)
    model.set_destination(stable,sheet_name='Results',cell='C4',table_name='tblQ')
    source=model.source_at(0)
    assert (source.sheet_name,source.cell,source.table_name)==('Results','C4','tblQ')
    model.undo_stack.undo()
    assert (source.sheet_name,source.cell,source.table_name)==(None,None,None)
    assert model.stable_id_at(0)==stable
    model.undo_stack.redo()
    assert (source.sheet_name,source.cell,source.table_name)==('Results','C4','tblQ')


def test_cancelled_bulk_destination_dialog_does_not_partially_edit_rows(tmp_path,monkeypatch):
    from PySide6.QtCore import QItemSelectionModel
    from pq_installer.core.models import LoadMode
    from pq_installer.ui import main_window
    paths=[tmp_path/f'{name}.pq' for name in ('One','Two')]
    for path in paths:path.write_text('1')
    sources=[QuerySource(path,path.stem,load_mode=LoadMode.NEW_SHEET) for path in paths]
    window=MainWindow();window._merge_sources(sources)
    selection=window.table.selectionModel()
    for row in range(2):
        selection.select(window.query_proxy.index(row,QueryTableModel.NAME),
                         QItemSelectionModel.Select|QItemSelectionModel.Rows)
    answers=iter([('SheetOne',True),('TableOne',True),('',False)])
    monkeypatch.setattr(main_window.QInputDialog,'getText',lambda *args,**kwargs:next(answers))
    window._configure_destinations()
    assert [(source.sheet_name,source.table_name) for source in sources]==[(None,None),(None,None)]
    assert window.query_model.undo_stack.count()==0
    window.close()


def test_visual_sort_does_not_mutate_execution_order(tmp_path):
    a=tmp_path/'a.m';b=tmp_path/'b.m';a.write_text('a');b.write_text('b')
    first,second=QuerySource(a,'Zulu',2),QuerySource(b,'Alpha',1)
    model=QueryTableModel([first,second]); proxy=QuerySortFilterProxyModel();proxy.setSourceModel(model)
    proxy.sort(2,Qt.AscendingOrder)
    assert proxy.data(proxy.index(0,2))=='Alpha'
    assert [first.order,second.order]==[2,1]
    assert [model.stable_id_at(i) for i in range(2)]==[model.stable_id_at(0),model.stable_id_at(1)]


def test_drag_reorder_preserves_stable_row_identity_and_updates_order(tmp_path):
    paths=[tmp_path/f'{n}.m' for n in 'abc']
    for p in paths:p.write_text('x')
    rows=[QuerySource(p,p.stem,i) for i,p in enumerate(paths)]
    model=QueryTableModel(rows); ids=[model.stable_id_at(i) for i in range(3)]
    assert model.moveRows(QModelIndex(),0,1,QModelIndex(),3)
    assert [model.stable_id_at(i) for i in range(3)]==[ids[1],ids[2],ids[0]]
    assert [model.source_at(i).order for i in range(3)]==[0,1,2]


def test_drag_reorder_can_be_undone_and_redone_without_changing_row_ids(tmp_path):
    paths=[tmp_path/f'{name}.pq' for name in ('A','B','C')]
    for path in paths:path.write_text('1')
    model=QueryTableModel([QuerySource(path,path.stem,index) for index,path in enumerate(paths)])
    original=[model.stable_id_at(index) for index in range(3)]
    assert model.moveRows(QModelIndex(),0,1,QModelIndex(),3)
    moved=[model.stable_id_at(index) for index in range(3)]
    assert moved==[original[1],original[2],original[0]]
    model.undo_stack.undo()
    assert [model.stable_id_at(index) for index in range(3)]==original
    model.undo_stack.redo()
    assert [model.stable_id_at(index) for index in range(3)]==moved


def test_drag_reorder_preserves_selection_for_the_same_stable_row(tmp_path):
    from PySide6.QtWidgets import QTableView
    paths=[tmp_path/f'{name}.pq' for name in ('A','B','C')]
    for path in paths:path.write_text('1')
    model=QueryTableModel([QuerySource(path,path.stem,index) for index,path in enumerate(paths)])
    view=QTableView();view.setModel(model);view.selectRow(1)
    selected_id=model.stable_id_at(view.selectionModel().selectedRows()[0].row())
    model.moveRows(QModelIndex(),0,1,QModelIndex(),3)
    selected_rows=view.selectionModel().selectedRows()
    assert len(selected_rows)==1
    assert model.stable_id_at(selected_rows[0].row())==selected_id


def test_inventory_comparison_updates_status_without_changing_source_or_order(tmp_path):
    paths=[tmp_path/f'{n}.pq' for n in ('same','changed','new','unk')]
    for p in paths:p.write_text('same' if p.stem=='same' else 'current')
    model=QueryTableModel([QuerySource(p,p.stem,i) for i,p in enumerate(paths)])
    model.apply_inventory([QueryInfo('same','same'),QueryInfo('changed','old'),QueryInfo('unk',None)])
    assert [model.data(model.index(i,QueryTableModel.STATUS)) for i in range(4)]==['UNCHANGED','MODIFIED','NEW','UNKNOWN']
    assert [model.source_at(i).order for i in range(4)]==[0,1,2,3]


def test_inventory_duplicate_names_are_conflicts_even_under_skip_policy(tmp_path):
    path=tmp_path/'q.pq';path.write_text('let x=1 in x')
    model=QueryTableModel([QuerySource(path,'Q')])
    model.apply_inventory([QueryInfo('Q','let x=1 in x'),QueryInfo('q','let x=2 in x')],skip_existing=True)
    assert model.data(model.index(0,QueryTableModel.STATUS))=='CONFLICT'


def test_duplicate_source_query_names_are_conflicts_before_import(tmp_path):
    first=tmp_path/'first.pq';second=tmp_path/'second.pq'
    first.write_text('1');second.write_text('2')
    model=QueryTableModel([QuerySource(first,'Shared'),QuerySource(second,'shared')])
    model.apply_inventory([])
    assert [model.data(model.index(i,QueryTableModel.STATUS)) for i in range(2)]==['CONFLICT','CONFLICT']


def test_bulk_row_removal_can_be_undone_and_redone_with_original_identity_and_order(tmp_path):
    paths=[tmp_path/f'{name}.pq' for name in ('A','B','C','D')]
    for path in paths:path.write_text('1')
    sources=[QuerySource(path,path.stem,index) for index,path in enumerate(paths)]
    model=QueryTableModel(sources)
    original_ids=[model.stable_id_at(index) for index in range(4)]
    before_count=model.undo_stack.count()

    model.remove_rows([1,3])

    assert [model.stable_id_at(index) for index in range(model.rowCount())]==[original_ids[0],original_ids[2]]
    assert model.undo_stack.count()==before_count+1
    model.undo_stack.undo()
    assert [model.stable_id_at(index) for index in range(model.rowCount())]==original_ids
    assert [model.source_at(index).name for index in range(model.rowCount())]==['A','B','C','D']
    model.undo_stack.redo()
    assert [model.stable_id_at(index) for index in range(model.rowCount())]==[original_ids[0],original_ids[2]]


def test_filtering_keeps_selection_mapping_to_underlying_row(tmp_path):
    paths=[tmp_path/f'{n}.m' for n in ('query-zzmatch','hide')]
    for p in paths:p.write_text('x')
    model=QueryTableModel([QuerySource(p,p.stem,i) for i,p in enumerate(paths)])
    proxy=QuerySortFilterProxyModel();proxy.setSourceModel(model);proxy.setFilterCaseSensitivity(Qt.CaseInsensitive);proxy.setFilterFixedString('zzmatch')
    assert proxy.rowCount()==1
    row=proxy.mapToSource(proxy.index(0,2)).row()
    assert model.source_at(row).name=='query-zzmatch'


def test_profile_replace_restores_stable_row_ids(tmp_path):
    p=tmp_path/'q.m';p.write_text('m')
    model=QueryTableModel();model.replace_sources([QuerySource(p,'q')],['profile-id'])
    assert model.row_for_id('profile-id')==0


def test_load_mode_requires_successful_excel_capability_probe(tmp_path):
    p=tmp_path/'q.m';p.write_text('x');model=QueryTableModel([QuerySource(p,'q')])
    assert not model.setData(model.index(0,QueryTableModel.LOAD),'connection_only',Qt.EditRole)
    assert model.setData(model.index(0,QueryTableModel.LOAD),'preserve_existing',Qt.EditRole)
    model.set_supported_load_modes(['definition_only','connection_only'])
    assert model.setData(model.index(0,QueryTableModel.LOAD),'connection_only',Qt.EditRole)
    assert model.setData(model.index(0,QueryTableModel.LOAD),'preserve_existing',Qt.EditRole)


def test_capability_probe_keeps_preserve_existing_available():
    from types import SimpleNamespace

    window=MainWindow()
    capabilities=SimpleNamespace(office_version='16.0',build='20430.0',
        supported_load_modes=('definition_only','connection_only'),reasons={})
    window._on_capabilities(capabilities)

    assert 'preserve_existing' in window.query_model.supported_load_modes
    assert 'preserve_existing' in window.load_mode_delegate.supported_modes
    window.close()


@__import__('pytest').mark.parametrize('accept',[True,False])
def test_profile_mapping_preview_requires_explicit_ui_acceptance(tmp_path,monkeypatch,accept):
    import json
    import pq_installer.ui.main_window as ui
    profile_dir=tmp_path/'profile';profile_dir.mkdir()
    source=profile_dir/'source'/'q.pq';source.parent.mkdir();source.write_text('let x=1 in x')
    mapping=profile_dir/'mapped'/'q.pq'
    profile=profile_dir/'deployment.json'
    profile.write_text(json.dumps({'schema_version':1,'profile_name':'Deploy','workbook':{'path':'book.xlsx'},
      'options':{'existing_query_policy':'replace','backup':True,'dry_run':True,'refresh':'none'},
      'path_mappings':[{'key':'target','from':'source','to':'mapped','enabled':True}],
      'queries':[{'id':'query-id','source_path':'source/q.pq','query_name':'Q','enabled':True,'order':0,'load':{'mode':'definition_only'}}]}))
    monkeypatch.setattr(ui.QFileDialog,'getOpenFileName',lambda *args:(str(profile),''))
    class Prompt:
        AcceptRole=1;RejectRole=2;Question=3
        def __init__(self,parent):self.buttons=[];self.details=''
        def setWindowTitle(self,*args):pass
        def setText(self,*args):pass
        def setDetailedText(self,text):self.details=text
        def setIcon(self,*args):pass
        def addButton(self,label,*args):button=object();self.buttons.append((label,button));return button
        def setDefaultButton(self,*args):pass
        def exec(self):pass
        def clickedButton(self):return self.buttons[0 if accept else 1][1]
    prompts=[]
    def make_prompt(parent):
        prompt=Prompt(parent);prompts.append(prompt);return prompt
    make_prompt.Question=Prompt.Question;make_prompt.AcceptRole=Prompt.AcceptRole;make_prompt.RejectRole=Prompt.RejectRole
    make_prompt.information=lambda *args:None
    monkeypatch.setattr(ui,'QMessageBox',make_prompt)
    window=MainWindow();window._load_profile()
    assert f'{source.resolve()} -> {mapping.resolve()}' in prompts[0].details
    if accept:
        assert window.query_model.source_at(0).path==mapping.resolve()
    else:
        assert window.query_model.rowCount()==0
    window.close()


def test_versioned_package_import_restores_original_names_and_order(tmp_path):
    from pq_installer.core.manager import QueryInfo,export_query_package
    package=tmp_path/'queries.pqm.zip'
    export_query_package([QueryInfo('Daily Sales','let x=1 in x'),QueryInfo('Helper (v2)','// helper\nlet x=2 in x')],package)
    window=MainWindow();window._import_zip(package)
    assert [(window.query_model.source_at(i).name,window.query_model.source_at(i).order)
            for i in range(window.query_model.rowCount())]==[('Daily Sales',0),('Helper (v2)',1)]
    assert len(window._staged_import_dirs)==1
    staged=window._staged_import_dirs[0]
    window.close()
    assert not staged.exists()


def test_refresh_requires_explicit_external_data_confirmation(tmp_path,monkeypatch):
    import pq_installer.ui.main_window as ui
    from PySide6.QtCore import QItemSelectionModel
    from PySide6.QtWidgets import QDialog, QTableView, QMessageBox
    from pq_installer.ui.inventory_model import QueryInventoryModel
    window=MainWindow()
    path=tmp_path/'book.xlsx';path.write_bytes(b'disposable')
    window.workbook_edit.setText(str(path))
    window.inventory=(QueryInfo('Sales','let x=1 in x'),)
    view=QTableView();view.setModel(QueryInventoryModel(window.inventory,view))
    view.selectionModel().select(view.model().index(0,0),QItemSelectionModel.Select|QItemSelectionModel.Rows)
    dialog=QDialog(window)
    monkeypatch.setattr(QDialog,'exec',lambda self:QDialog.Accepted)
    monkeypatch.setattr(QMessageBox,'question',lambda *args,**kwargs:QMessageBox.Cancel)
    window._refresh_inventory(dialog,view,selected_only=True)
    assert dialog.result()!=QDialog.Accepted
    assert window._refresh_thread is None
    window.close()
