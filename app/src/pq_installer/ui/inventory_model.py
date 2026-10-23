from PySide6.QtCore import QAbstractTableModel, Qt, QModelIndex


class QueryInventoryModel(QAbstractTableModel):
    HEADERS=('Query name','M formula','Description','Load status')
    def __init__(self,queries=(),parent=None):
        super().__init__(parent);self.queries=tuple(queries)
    def rowCount(self,parent=QModelIndex()):return 0 if parent.isValid() else len(self.queries)
    def columnCount(self,parent=QModelIndex()):return 0 if parent.isValid() else len(self.HEADERS)
    def headerData(self,section,orientation,role=Qt.DisplayRole):
        if role==Qt.DisplayRole and orientation==Qt.Horizontal and 0<=section<len(self.HEADERS):return self.HEADERS[section]
        return super().headerData(section,orientation,role)
    def data(self,index,role=Qt.DisplayRole):
        if not index.isValid() or role not in (Qt.DisplayRole,Qt.ToolTipRole):return None
        query=self.queries[index.row()]
        value=(query.name,query.formula,query.description or '',query.load_state)[index.column()]
        return value if role==Qt.DisplayRole else str(value)
    def flags(self,index):return Qt.ItemIsEnabled|Qt.ItemIsSelectable if index.isValid() else Qt.NoItemFlags
