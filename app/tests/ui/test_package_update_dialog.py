import os
os.environ.setdefault('QT_QPA_PLATFORM','offscreen')

from PySide6.QtCore import Qt
from PySide6.QtWidgets import QApplication, QDialog

from pq_installer.core.manager import QueryPackageChange
from pq_installer.ui.package_update_dialog import QueryPackageUpdateDialog

app=QApplication.instance() or QApplication([])


def test_package_update_dialog_shows_change_list_and_selected_diff():
    changes=(
        QueryPackageChange('Added','ADDED',None,None,'let A = 1 in A','+let A = 1 in A'),
        QueryPackageChange('Stable','UNCHANGED','Stable','1','1',''),
        QueryPackageChange('Changed','MODIFIED','Changed','let A = 1 in A','let A = 2 in A','-let A = 1 in A\n+let A = 2 in A'),
        QueryPackageChange('Removed','REMOVED','Removed','let A = 9 in A',None,'-let A = 9 in A'),
        QueryPackageChange('same','CONFLICT','Same','1','1',''),
    )
    dialog=QueryPackageUpdateDialog('base.pqm.zip','next.pqm.zip',changes)
    assert dialog.change_table.rowCount()==5
    assert dialog.summary_label.text().startswith('1 added · 1 modified · 1 unchanged · 1 removed')
    index=dialog.change_table.model().index(2,0)
    dialog.change_table.setCurrentCell(index.row(),index.column())
    assert '-let A = 1 in A' in dialog.diff_view.toPlainText()
    assert 'no workbook queries are deleted' in dialog.safety_note.text()
    dialog.close()


def test_package_update_dialog_empty_change_set_is_readable():
    dialog=QueryPackageUpdateDialog('base.pqm.zip','next.pqm.zip',())
    assert dialog.change_table.rowCount()==0
    assert 'No package changes' in dialog.summary_label.text()
    dialog.close()


def test_manager_package_preview_is_read_only_and_records_counts(tmp_path,monkeypatch):
    from pq_installer.core.manager import QueryInfo,export_query_package
    from pq_installer.ui import main_window
    base=export_query_package([QueryInfo('Keep','1'),QueryInfo('Drop','2')],tmp_path/'base.pqm.zip')
    update=export_query_package([QueryInfo('Keep','1'),QueryInfo('New','3')],tmp_path/'update.pqm.zip')
    selected=iter((str(base),str(update)))
    monkeypatch.setattr(main_window.QFileDialog,'getOpenFileName',lambda *args,**kwargs:(next(selected),''))
    shown=[]
    monkeypatch.setattr(QueryPackageUpdateDialog,'exec',lambda self:(shown.append(self) or QDialog.Accepted))
    window=main_window.MainWindow()
    before_sources=tuple(window.sources)
    history=[]
    monkeypatch.setattr(window,'_record_history',lambda event,details:history.append((event,details)))
    window._preview_query_package_update()
    assert len(shown)==1
    assert [(row.name,row.status) for row in shown[0].changes]==[
        ('Keep','UNCHANGED'),('New','ADDED'),('Drop','REMOVED')]
    assert tuple(window.sources)==before_sources
    assert window.query_model.rowCount()==0
    assert history==[('query_package_update_previewed',
        {'ADDED':1,'MODIFIED':0,'UNCHANGED':1,'REMOVED':1,'CONFLICT':0})]
    window.close()
