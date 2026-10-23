import os
os.environ.setdefault('QT_QPA_PLATFORM','offscreen')
from PySide6.QtWidgets import QApplication
from PySide6.QtCore import Qt
from pq_installer.core.manager import QueryInfo
from pq_installer.ui.inventory_model import QueryInventoryModel

app=QApplication.instance() or QApplication([])

def test_inventory_model_displays_formula_description_and_unknown_load():
    model=QueryInventoryModel([QueryInfo('Sales','let x=1 in x','daily','unknown')])
    assert model.rowCount()==1 and model.columnCount()==4
    assert model.data(model.index(0,0),Qt.DisplayRole)=='Sales'
    assert model.data(model.index(0,1),Qt.DisplayRole)=='let x=1 in x'
    assert model.data(model.index(0,3),Qt.DisplayRole)=='unknown'
    assert not (model.flags(model.index(0,0)) & Qt.ItemIsEditable)
