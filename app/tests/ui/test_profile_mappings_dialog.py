import os
os.environ.setdefault('QT_QPA_PLATFORM','offscreen')

import pytest
from PySide6.QtWidgets import QApplication, QDialog

from pq_installer.ui.profile_mappings_dialog import ProfileMappingsDialog

app=QApplication.instance() or QApplication([])


def test_profile_mappings_dialog_edits_path_and_m_source_rules():
    dialog=ProfileMappingsDialog(
        [{'key':'workspace','from':'C:/old','to':'D:/new','enabled':True}],
        [{'key':'lake','from':'C:/data','to':'D:/data','enabled':False}])
    dialog.path_table.item(0,0).setText('local')
    dialog.path_table.item(0,2).setText('E:/queries')
    dialog.environment_table.item(0,3).setCheckState(dialog.unchecked_state)
    path_rules,environment_rules=dialog.mappings()
    assert path_rules==[{'key':'local','from':'C:/old','to':'E:/queries','enabled':True}]
    assert environment_rules==[{'key':'lake','from':'C:/data','to':'D:/data','enabled':False}]
    assert 'source-file paths' in dialog.path_help.text().casefold()
    assert 'M source functions' in dialog.environment_help.text()
    dialog.close()


def test_profile_mappings_dialog_rejects_duplicate_keys_without_mutating_input():
    original=[{'key':'share','from':'C:/one','to':'D:/one','enabled':True}]
    dialog=ProfileMappingsDialog(original,[])
    dialog.add_path_mapping()
    dialog.path_table.item(1,0).setText('SHARE')
    dialog.path_table.item(1,1).setText('C:/two')
    dialog.path_table.item(1,2).setText('D:/two')
    with pytest.raises(ValueError,match='duplicate mapping key'):
        dialog.mappings()
    assert original==[{'key':'share','from':'C:/one','to':'D:/one','enabled':True}]
    dialog.close()


def test_profile_mapping_dialog_named_select_all_action_selects_every_path_row():
    from PySide6.QtCore import Qt
    from PySide6.QtTest import QTest
    dialog=ProfileMappingsDialog([
        {'key':'one','from':'C:/one','to':'D:/one'},
        {'key':'two','from':'C:/two','to':'D:/two'},
    ],[])
    dialog.show();app.processEvents()
    button=dialog.select_all_path_button
    assert button.accessibleName()=='Select all source path mappings'
    QTest.mouseClick(button,Qt.MouseButton.LeftButton);app.processEvents()
    assert len(dialog.path_table.selectionModel().selectedRows())==2
    dialog.close()


def test_manager_applies_only_accepted_mapping_edits(monkeypatch):
    from pq_installer.ui import main_window
    from pq_installer.core.models import QuerySource

    window=main_window.MainWindow()
    window._loaded_profile_path_mappings=[{'key':'old','from':'C:/old','to':'D:/old','enabled':True}]
    source=QuerySource(name='Existing',path='C:/query.m',enabled=True)
    window.sources=[source]
    window.query_model.replace_sources([source],["stable-existing"])
    original_source=window.sources[0]
    recorded=[]
    monkeypatch.setattr(window,'_record_history',lambda event,details:recorded.append((event,details)))

    class Accepted:
        def __init__(self,*args): pass
        def exec(self): return QDialog.DialogCode.Accepted
        def mappings(self):
            return ([{'key':'new','from':'C:/old','to':'E:/new','enabled':True}],[])

    monkeypatch.setattr(main_window,'ProfileMappingsDialog',Accepted)
    window._edit_source_mappings()
    assert window._loaded_profile_path_mappings==[{'key':'new','from':'C:/old','to':'E:/new','enabled':True}]
    assert window.sources[0] is original_source
    assert recorded==[('profile_source_mappings_edited',{'path_mapping_count':1,'environment_mapping_count':0})]

    class Cancelled(Accepted):
        def exec(self): return QDialog.DialogCode.Rejected

    monkeypatch.setattr(main_window,'ProfileMappingsDialog',Cancelled)
    window._edit_source_mappings()
    assert window._loaded_profile_path_mappings[0]['key']=='new'
    assert len(recorded)==1
    window.close()
