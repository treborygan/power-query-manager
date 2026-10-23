from pq_installer.core.history_comparison import compare_query_inventories
from pq_installer.core.manager import QueryInfo


def test_snapshot_inventory_comparison_reports_formula_load_add_remove_and_unknown():
    before=(
        QueryInfo('Stable','let A=1 in A',load_state='definition_only'),
        QueryInfo('Formula','let A=1 in A',load_state='connection_only'),
        QueryInfo('Load','let A=1 in A',load_state='definition_only'),
        QueryInfo('Gone','let A=2 in A',load_state='definition_only'),
        QueryInfo('Unknown',None,load_state='unknown'),
    )
    after=(
        QueryInfo('Stable','let A=1 in A',load_state='definition_only'),
        QueryInfo('Formula','let A=2 in A',load_state='connection_only'),
        QueryInfo('Load','let A=1 in A',load_state='new_sheet'),
        QueryInfo('Added','let A=3 in A',load_state='definition_only'),
        QueryInfo('Unknown',None,load_state='unknown'),
    )
    changes=compare_query_inventories(before,after)
    assert [(row.name,row.status) for row in changes]==[
        ('Stable','UNCHANGED'),('Formula','MODIFIED'),('Load','MODIFIED'),
        ('Gone','REMOVED'),('Unknown','UNKNOWN'),('Added','ADDED')]
    assert '-let A=1 in A' in changes[1].diff and '+let A=2 in A' in changes[1].diff
    assert 'load destination' in changes[2].reason.casefold()


def test_snapshot_inventory_comparison_distinguishes_observed_worksheet_and_model_loads():
    before = [QueryInfo('Q', 'let A=1 in A', load_state='worksheet: Output!$A$1:$A$3')]
    after = [QueryInfo('Q', 'let A=1 in A', load_state='data_model')]

    changes = compare_query_inventories(before, after)

    assert len(changes) == 1
    assert changes[0].status == 'MODIFIED'
    assert changes[0].before.load_state == 'worksheet: Output!$A$1:$A$3'
    assert changes[0].after.load_state == 'data_model'
    assert 'worksheet: Output!$A$1:$A$3' in changes[0].reason
    assert 'data_model' in changes[0].reason


def test_snapshot_inventory_comparison_marks_case_collisions_conflict():
    changes=compare_query_inventories(
        [QueryInfo('Sales','1'),QueryInfo('sales','2')],
        [QueryInfo('SALES','3')])
    assert len(changes)==3
    assert all(row.status=='CONFLICT' for row in changes)
    assert [row.before.formula for row in changes if row.before] == ['1','2']
    assert [row.after.formula for row in changes if row.after] == ['3']


def test_snapshot_inventory_comparison_reports_query_display_name_change():
    changes=compare_query_inventories([QueryInfo('Sales','1',load_state='definition_only')],
                                     [QueryInfo('SALES','1',load_state='definition_only')])
    assert changes[0].status=='MODIFIED'
    assert 'display name changed' in changes[0].reason.casefold()


def test_compare_worker_revalidates_snapshots_and_inspects_read_only(tmp_path):
    from pq_installer.core.history import HistoryStore
    from pq_installer.ui.history_worker import SnapshotCompareWorker
    store=HistoryStore(tmp_path/'history.sqlite3')
    old=tmp_path/'old.xlsx';old.write_bytes(b'old workbook bytes')
    new=tmp_path/'new.xlsx';new.write_bytes(b'new workbook bytes')
    before=store.create_snapshot(old)['snapshot_id']
    after=store.create_snapshot(new)['snapshot_id']
    inventories=iter(((QueryInfo('Q','1',load_state='definition_only'),),
                       (QueryInfo('Q','2',load_state='definition_only'),)))
    opened=[]
    class Inspector:
        def inspect(self,path):
            opened.append(path)
            return next(inventories)
    outcomes=[];failures=[]
    worker=SnapshotCompareWorker(store,before,after,lambda:Inspector())
    worker.completed.connect(outcomes.append);worker.failed.connect(failures.append);worker.run()
    assert not failures and len(outcomes)==1
    assert len(opened)==2
    assert outcomes[0]['changes'][0].status=='MODIFIED'
    assert '-1' in outcomes[0]['changes'][0].diff and '+2' in outcomes[0]['changes'][0].diff


def test_history_comparison_dialog_shows_diff_and_read_only_refresh_notice():
    import os
    os.environ.setdefault('QT_QPA_PLATFORM','offscreen')
    from PySide6.QtWidgets import QApplication
    from pq_installer.core.history_comparison import QueryVersionChange
    from pq_installer.ui.history_comparison_dialog import HistoryComparisonDialog
    app=QApplication.instance() or QApplication([])
    dialog=HistoryComparisonDialog({'occurred_at':'earlier'},{'occurred_at':'later'},(
        QueryVersionChange('Q','MODIFIED',QueryInfo('Q','1'),QueryInfo('Q','2'),
                           'M formula changed.','-1\n+2'),))
    assert dialog.change_table.rowCount()==1
    assert '-1' in dialog.diff_view.toPlainText()
    assert 'never refreshes external data' in dialog.summary_label.text()
    assert dialog.diff_view.isReadOnly()
    dialog.close()

def test_snapshot_comparison_keeps_every_later_duplicate_in_conflict_results():
    before = (QueryInfo('Sales', 'old formula', load_state='definition_only'),)
    after = (
        QueryInfo('SALES', 'new formula one', load_state='definition_only'),
        QueryInfo('sales', 'new formula two', load_state='connection_only'),
    )

    changes = compare_query_inventories(before, after)

    assert len(changes) == 3
    assert all(change.status == 'CONFLICT' for change in changes)
    assert [change.before.formula if change.before else None for change in changes] == [
        'old formula', None, None,
    ]
    assert [change.after.formula if change.after else None for change in changes] == [
        None, 'new formula one', 'new formula two',
    ]
