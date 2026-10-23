"""Stable-identity Qt table model; view sorting is independent of execution order."""
from dataclasses import dataclass
from uuid import uuid4
from pathlib import Path

from PySide6.QtCore import QAbstractTableModel, QSortFilterProxyModel, Qt, QModelIndex, QByteArray, QMimeData
from PySide6.QtGui import QUndoCommand, QUndoStack

from pq_installer.core.models import QuerySource, LoadMode
from pq_installer.core.dependencies import analyze_import_dependencies


@dataclass
class QueryRow:
    stable_id: str
    source: QuerySource
    status: str = 'Not inspected'
    dependency_status: str = 'Not analyzed'


class _RowEditCommand(QUndoCommand):
    def __init__(self, model, stable_id, column, before, after, label):
        super().__init__(label)
        self.model=model;self.stable_id=stable_id;self.column=column
        self.before=before;self.after=after
    def undo(self):self.model._apply_edit_state(self.stable_id,self.column,self.before)
    def redo(self):self.model._apply_edit_state(self.stable_id,self.column,self.after)


class _RowOrderCommand(QUndoCommand):
    def __init__(self,model,before,after):
        super().__init__('Reorder queries')
        self.model=model;self.before=tuple(before);self.after=tuple(after)
    def undo(self):self.model._apply_row_order(self.before)
    def redo(self):self.model._apply_row_order(self.after)


class _RowsRemoveCommand(QUndoCommand):
    def __init__(self, model, removed_rows):
        super().__init__('Remove selected queries')
        self.model=model
        self.removed_rows=tuple(removed_rows)

    def undo(self):self.model._restore_removed_rows(self.removed_rows)

    def redo(self):self.model._remove_rows_by_ids(row.stable_id for _,row in self.removed_rows)


class QueryTableModel(QAbstractTableModel):
    ENABLED, ORDER, NAME, PATH, LOAD, STATUS, DEPENDENCIES = range(7)
    HEADERS = ('Enabled', 'Order', 'Query name', 'Source path', 'Load mode', 'Workbook status', 'Dependencies')
    MIME = 'application/x-power-query-row-ids'

    def __init__(self, sources=(), parent=None, supported_load_modes=('definition_only',)):
        super().__init__(parent)
        self._rows = [QueryRow(uuid4().hex, source) for source in sources]
        self.supported_load_modes=set(supported_load_modes)
        # This is a preservation policy, not an Office load capability.
        self.supported_load_modes.add(LoadMode.PRESERVE_EXISTING.value)
        self.undo_stack=QUndoStack(self)

    def rowCount(self, parent=QModelIndex()):
        return 0 if parent.isValid() else len(self._rows)

    def columnCount(self, parent=QModelIndex()):
        return 0 if parent.isValid() else len(self.HEADERS)

    def headerData(self, section, orientation, role=Qt.DisplayRole):
        if role == Qt.DisplayRole and orientation == Qt.Horizontal and 0 <= section < len(self.HEADERS):
            return self.HEADERS[section]
        return super().headerData(section, orientation, role)

    def data(self, index, role=Qt.DisplayRole):
        if not index.isValid() or not (0 <= index.row() < len(self._rows)):
            return None
        row = self._rows[index.row()]
        source = row.source
        if index.column() == self.ENABLED and role == Qt.CheckStateRole:
            return Qt.Checked if source.enabled else Qt.Unchecked
        if role in (Qt.DisplayRole, Qt.EditRole):
            return (None, source.order, source.name, str(source.path), source.load_mode.value, row.status, row.dependency_status)[index.column()]
        if role == Qt.UserRole:
            return row.stable_id
        if role == Qt.ToolTipRole:
            if index.column() == self.PATH:return str(source.path)
            if index.column() == self.STATUS:return row.status
            if index.column() == self.LOAD:
                if source.load_mode==LoadMode.EXISTING_SHEET:return f'Worksheet {source.sheet_name or "(unset)"}, starting cell {source.cell or "(unset)"}; table {source.table_name or "automatic"}'
                if source.load_mode==LoadMode.NEW_SHEET:return f'New worksheet {source.sheet_name or "(automatic name)"}; table {source.table_name or "automatic"}'
            return None
        return None

    def flags(self, index):
        if not index.isValid():
            return Qt.ItemIsDropEnabled
        flags = Qt.ItemIsEnabled | Qt.ItemIsSelectable | Qt.ItemIsDragEnabled
        if index.column() == self.ENABLED:
            return flags | Qt.ItemIsUserCheckable
        if index.column() in (self.ORDER, self.NAME, self.PATH, self.LOAD):
            return flags | Qt.ItemIsEditable | Qt.ItemIsDropEnabled
        return flags

    def setData(self, index, value, role=Qt.EditRole):
        if not index.isValid() or not (0 <= index.row() < len(self._rows)):
            return False
        row=self._rows[index.row()]
        source=row.source
        try:
            if index.column() == self.ENABLED and role == Qt.CheckStateRole:
                after={'enabled':value==Qt.Checked}
            elif role == Qt.EditRole and index.column() == self.ORDER:
                order = int(value)
                if order < 0: return False
                after={'order':order,'order_automatic':False}
            elif role == Qt.EditRole and index.column() == self.NAME:
                after={'name':str(value)}
            elif role == Qt.EditRole and index.column() == self.PATH:
                after={'path':Path(str(value)),'profile_source_path':None}
            elif role == Qt.EditRole and index.column() == self.LOAD:
                if str(value) not in self.supported_load_modes:return False
                after={'load_mode':LoadMode(str(value))}
            else:
                return False
        except (TypeError, ValueError):
            return False
        before=self._edit_state(source,index.column())
        if before==after:return True
        self.undo_stack.push(_RowEditCommand(self,row.stable_id,index.column(),before,after,
            f'Edit {self.HEADERS[index.column()].lower()} for {source.name}'))
        return True

    @staticmethod
    def _edit_state(source,column):
        if column==QueryTableModel.ENABLED:return {'enabled':source.enabled}
        if column==QueryTableModel.ORDER:return {'order':source.order,'order_automatic':source.order_automatic}
        if column==QueryTableModel.NAME:return {'name':source.name}
        if column==QueryTableModel.PATH:return {'path':source.path,'profile_source_path':source.profile_source_path}
        if column==QueryTableModel.LOAD:return {'load_mode':source.load_mode}
        if column==QueryTableModel.DESTINATION:return {'sheet_name':source.sheet_name,'cell':source.cell,'table_name':source.table_name}
        raise ValueError('Unsupported editable query field.')

    DESTINATION=-1

    def _apply_edit_state(self,stable_id,column,state):
        row_index=self.row_for_id(stable_id)
        if row_index<0:return
        source=self._rows[row_index].source
        for attribute,value in state.items():setattr(source,attribute,value)
        display_column=self.LOAD if column==self.DESTINATION else column
        index=self.index(row_index,display_column)
        self.dataChanged.emit(index,index,[Qt.DisplayRole,Qt.EditRole,Qt.CheckStateRole,Qt.ToolTipRole])

    def set_destination(self,stable_id,*,sheet_name=None,cell=None,table_name=None):
        row_index=self.row_for_id(stable_id)
        if row_index<0:return False
        source=self._rows[row_index].source
        before=self._edit_state(source,self.DESTINATION)
        after={'sheet_name':sheet_name,'cell':cell,'table_name':table_name}
        if before==after:return True
        self.undo_stack.push(_RowEditCommand(self,stable_id,self.DESTINATION,before,after,
            f'Configure destination for {source.name}'))
        return True

    def source_at(self, row):
        return self._rows[row].source

    def stable_id_at(self, row):
        return self._rows[row].stable_id

    def row_for_id(self, stable_id):
        return next((i for i, row in enumerate(self._rows) if row.stable_id == stable_id), -1)

    def set_status(self, stable_id, status):
        row = self.row_for_id(stable_id)
        if row < 0: return False
        status = str(status)
        if self._rows[row].status == status: return True
        self._rows[row].status = status
        self.dataChanged.emit(self.index(row, self.STATUS), self.index(row, self.STATUS))
        return True

    def set_dependency_status(self, stable_id, status):
        row=self.row_for_id(stable_id)
        if row<0:return False
        status=str(status)
        if self._rows[row].dependency_status==status:return True
        self._rows[row].dependency_status=status
        self.dataChanged.emit(self.index(row,self.DEPENDENCIES),self.index(row,self.DEPENDENCIES))
        return True

    def apply_dependency_analysis(self):
        statuses=analyze_import_dependencies([entry.source for entry in self._rows])
        for entry in self._rows:
            self.set_dependency_status(entry.stable_id,statuses.get(entry.source.name,'Unknown dependency status'))

    def append_sources(self, sources):
        seen_paths = {str(row.source.path).casefold() for row in self._rows}
        additions = []
        for source in sources:
            path_key = str(source.path).casefold()
            if path_key in seen_paths:
                continue
            seen_paths.add(path_key)
            additions.append(source)
        if not additions: return 0
        first = len(self._rows); self.beginInsertRows(QModelIndex(), first, first + len(additions) - 1)
        self._rows.extend(QueryRow(uuid4().hex, source) for source in additions)
        self.endInsertRows()
        return len(additions)

    def refresh_all(self):
        if self._rows:
            self.dataChanged.emit(self.index(0, 0), self.index(len(self._rows) - 1, len(self.HEADERS) - 1))

    def clear(self):
        if not self._rows: return
        self.undo_stack.clear();self.beginResetModel(); self._rows.clear(); self.endResetModel()

    def replace_sources(self, sources, stable_ids):
        if len(sources)!=len(stable_ids) or len(set(stable_ids))!=len(stable_ids):
            raise ValueError('Profile row IDs must be unique and match the source rows.')
        self.undo_stack.clear();self.beginResetModel()
        self._rows=[QueryRow(str(stable_id),source) for source,stable_id in zip(sources,stable_ids)]
        self.endResetModel()

    def set_supported_load_modes(self,modes):
        # Capability probes gate destinations that create Excel connections or
        # output. Existing-load preservation is safe and available on every build.
        self.supported_load_modes=set(modes)
        self.supported_load_modes.add(LoadMode.PRESERVE_EXISTING.value)
        if self.rowCount():self.dataChanged.emit(self.index(0,self.LOAD),self.index(self.rowCount()-1,self.LOAD))

    def apply_inventory(self, inventory, *, skip_existing=False):
        from pq_installer.core.manager import compare_source
        existing={}
        for query in inventory:existing.setdefault(query.name.casefold(),[]).append(query)
        source_groups={}
        for entry in self._rows:source_groups.setdefault(entry.source.name.casefold(),[]).append(entry)
        duplicate_sources={key for key,entries in source_groups.items()
                           if len(entries)>1 and any(entry.source.enabled for entry in entries)}
        for row,entry in enumerate(self._rows):
            try:
                raw=entry.source.path.read_bytes()
                formula=raw.decode('utf-8-sig')
                if not formula.strip(): raise ValueError('empty source')
            except Exception:
                status='UNKNOWN'
            else:
                key=entry.source.name.casefold()
                matches=existing.get(key,[])
                if key in duplicate_sources or len(matches)>1:
                    status='CONFLICT'
                    self.set_status(entry.stable_id,status)
                    continue
                match=matches[0] if matches else None
                verdict=compare_source(entry.source.name,formula,match,skip=skip_existing)
                status=verdict.value
            self.set_status(entry.stable_id,status)

    def remove_rows(self, rows):
        valid=sorted({row for row in rows if 0 <= row < len(self._rows)})
        if not valid:return
        removed_rows=[(row,self._rows[row]) for row in valid]
        self.undo_stack.push(_RowsRemoveCommand(self,removed_rows))

    def _remove_rows_by_ids(self,stable_ids):
        rows=sorted({self.row_for_id(stable_id) for stable_id in stable_ids if self.row_for_id(stable_id)>=0},
                    reverse=True)
        for row in rows:
            self.beginRemoveRows(QModelIndex(),row,row)
            self._rows.pop(row)
            self.endRemoveRows()

    def _restore_removed_rows(self,removed_rows):
        for row,row_data in removed_rows:
            if self.row_for_id(row_data.stable_id)>=0:continue
            row=min(row,len(self._rows))
            self.beginInsertRows(QModelIndex(),row,row)
            self._rows.insert(row,row_data)
            self.endInsertRows()

    def _apply_row_order(self,stable_ids):
        if len(stable_ids)!=len(self._rows) or set(stable_ids)!={row.stable_id for row in self._rows}:return
        parent=QModelIndex()
        for target,stable_id in enumerate(stable_ids):
            current=next(index for index,row in enumerate(self._rows) if row.stable_id==stable_id)
            if current==target:continue
            destination=target if current>target else target+1
            if not self.beginMoveRows(parent,current,current,parent,destination):return
            row=self._rows.pop(current)
            self._rows.insert(destination-1 if current<destination else destination,row)
            self.endMoveRows()
        for order,row in enumerate(self._rows):
            row.source.order=order;row.source.order_automatic=False
        if self._rows:
            self.dataChanged.emit(self.index(0,self.ORDER),self.index(len(self._rows)-1,self.ORDER),
                                  [Qt.DisplayRole,Qt.EditRole])

    def moveRows(self, source_parent, source_row, count, destination_parent, destination_child):
        if source_parent.isValid() or destination_parent.isValid() or count != 1:
            return False
        if source_row < 0 or source_row >= len(self._rows) or destination_child < 0 or destination_child > len(self._rows):
            return False
        if destination_child in (source_row, source_row + 1): return False
        reordered=list(self._rows)
        row = reordered.pop(source_row)
        target = destination_child - 1 if source_row < destination_child else destination_child
        reordered.insert(target, row)
        before=[entry.stable_id for entry in self._rows]
        after=[entry.stable_id for entry in reordered]
        self.undo_stack.push(_RowOrderCommand(self,before,after))
        return True

    def mimeTypes(self):
        return [self.MIME]

    def mimeData(self, indexes):
        payload = QMimeData()
        rows = sorted({index.row() for index in indexes if index.isValid()})
        payload.setData(self.MIME, QByteArray(','.join(self.stable_id_at(i) for i in rows).encode()))
        return payload

    def dropMimeData(self, data, action, row, column, parent):
        if action == Qt.IgnoreAction: return True
        if action != Qt.MoveAction or not data.hasFormat(self.MIME): return False
        ids = bytes(data.data(self.MIME)).decode().split(',')
        if len(ids) != 1: return False
        source_row = self.row_for_id(ids[0])
        destination = row if row >= 0 else parent.row() if parent.isValid() else len(self._rows)
        return self.moveRows(QModelIndex(), source_row, 1, QModelIndex(), destination)

    def supportedDropActions(self):
        return Qt.MoveAction


class QuerySortFilterProxyModel(QSortFilterProxyModel):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setFilterKeyColumn(-1)
        self.setFilterCaseSensitivity(Qt.CaseInsensitive)

    def filterAcceptsRow(self, source_row, source_parent):
        needle = self.filterRegularExpression().pattern().casefold()
        if not needle:
            return True
        model = self.sourceModel()
        return any(needle in str(model.data(model.index(source_row, column), Qt.DisplayRole) or '').casefold()
                   for column in range(model.columnCount()))

    def lessThan(self, left, right):
        source = self.sourceModel()
        if left.column() == QueryTableModel.ORDER:
            return int(source.data(left, Qt.DisplayRole)) < int(source.data(right, Qt.DisplayRole))
        return str(source.data(left, Qt.DisplayRole) or '').casefold() < str(source.data(right, Qt.DisplayRole) or '').casefold()
