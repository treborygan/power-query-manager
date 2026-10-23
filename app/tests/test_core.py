from pathlib import Path
from datetime import datetime
from types import SimpleNamespace
import os
import pytest
from pq_installer.core.query_loader import discover_query_files, discover_m_files, read_m_code
from pq_installer.core.validation import validate_workbook, validate_workbook_for_write, validate_query_sources, ordered_sources, order_warnings
from pq_installer.core.backup import create_timestamped_backup
from pq_installer.core.models import QuerySource


def test_windows_reparse_point_detector_uses_lstat_attributes(monkeypatch, tmp_path):
    from pq_installer.core.path_safety import contains_reparse_point, is_reparse_point

    target = tmp_path / 'junction.xlsx'
    monkeypatch.setattr('pq_installer.core.path_safety.os.lstat', lambda _path: SimpleNamespace(st_file_attributes=0x400, st_mode=0))
    assert is_reparse_point(target)
    assert contains_reparse_point(target)


def test_write_preflight_rejects_windows_reparse_point(monkeypatch, tmp_path):
    workbook = tmp_path / 'junction.xlsx'
    workbook.write_bytes(b'workbook')
    monkeypatch.setattr('pq_installer.core.validation.contains_reparse_point', lambda _path: True)
    with pytest.raises(ValueError, match='reparse point'):
        validate_workbook_for_write(workbook, active_file_users=lambda _path: None)


def test_workbook_path_check_rejects_reparse_point_in_parent(monkeypatch, tmp_path):
    import pq_installer.core.path_safety as path_safety

    workbook = tmp_path / 'through-junction' / 'book.xlsx'
    workbook.parent.mkdir()
    workbook.write_bytes(b'workbook')
    real_lstat = os.lstat
    def lstat_with_junction(path):
        if Path(path) == workbook.parent:
            return SimpleNamespace(st_file_attributes=0x400, st_mode=0)
        return real_lstat(path)
    monkeypatch.setattr(path_safety.os, 'lstat', lstat_with_junction)
    assert path_safety.contains_reparse_point(workbook)


def test_mixed_extensions_and_prefix_order(tmp_path):
    names=['1_qQueryName.pq','02-Helper.m','003 Main.pq','z-last.m','a-first.pq','4digit.m','bad.txt']
    for name in names: (tmp_path/name).write_text('let x=1 in x',encoding='utf-8')
    found=discover_query_files([tmp_path])
    assert [(x.order,x.name,x.path.suffix) for x in found]==[(1,'qQueryName','.pq'),(2,'Helper','.m'),(3,'Main','.pq'),(4,'4digit','.m'),(5,'a-first','.pq'),(6,'z-last','.m')]

def test_numeric_prefix_requires_separator(tmp_path):
    for n in ['2foo.m','02 Foo.m','0_bar.pq']: (tmp_path/n).write_text('x')
    found={s.path.name:(s.order,s.name) for s in discover_query_files([tmp_path])}
    assert found['2foo.m']==(3,'2foo')
    assert found['02 Foo.m']==(2,'Foo')
    assert found['0_bar.pq']==(0,'bar')

def test_mixed_explicit_order_ties_are_visible_and_manual_order_wins(tmp_path):
    for n in ['2_b.m','2_a.pq','1_c.m']: (tmp_path/n).write_text('x')
    src=discover_query_files([tmp_path]); assert [s.name for s in ordered_sources(src)]==['c','a','b']
    assert order_warnings(src)
    src[0].order=9
    assert [s.name for s in ordered_sources(src)]==['a','b','c']

def test_adding_sources_preserves_source_file_bytes_unicode_and_order(tmp_path):
    p=tmp_path/'x.pq'; content='// café\nlet x="雪" in x'; p.write_text(content,encoding='utf-8')
    s=discover_query_files([p])[0]; s.order=17; assert read_m_code(s)==content and s.order==17

def test_duplicate_names_are_case_insensitive(tmp_path):
    a=tmp_path/'a.m'; b=tmp_path/'b.pq'; a.write_text('x'); b.write_text('y')
    src=[QuerySource(a,'Q'),QuerySource(b,'q')]
    with pytest.raises(ValueError,match='case-insensitive'): validate_query_sources(src)

def test_missing_and_unsupported_files_actionable(tmp_path):
    p=tmp_path/'bad.txt'; p.write_text('x')
    with pytest.raises(ValueError,match=r'\.m or \.pq'): validate_query_sources([QuerySource(p,'q')])
    with pytest.raises(ValueError,match='does not exist'): validate_query_sources([QuerySource(tmp_path/'gone.m','gone')])

def test_invalid_query_names_rejected_before_excel(tmp_path):
    p=tmp_path/'x.m'; p.write_text('x')
    for name in ['', ' bad','bad/name']:
        with pytest.raises(ValueError): validate_query_sources([QuerySource(p,name)])

def test_empty_and_undecodable_file_errors(tmp_path):
    p=tmp_path/'q.m'; p.write_text('   ',encoding='utf8')
    with pytest.raises(ValueError,match='empty'): read_m_code(discover_query_files([p])[0])
    p.write_bytes(b'\xff\xfe\x81')
    with pytest.raises(ValueError,match='UTF-8'): read_m_code(discover_query_files([p])[0])

def test_workbook_extension_validation(tmp_path):
    good=tmp_path/'book.xlsm'; good.write_bytes(b'x'); assert validate_workbook(good)==good
    bad=tmp_path/'book.csv'; bad.write_bytes(b'x')
    with pytest.raises(ValueError,match='xlsx or .xlsm'): validate_workbook(bad)

def test_workbook_write_preflight_rejects_excel_lock_file(tmp_path):
    from pq_installer.core.validation import validate_workbook_for_write
    book=tmp_path/'book.xlsx';book.write_bytes(b'disposable')
    (tmp_path/'~$book.xlsx').write_bytes(b'Excel owner lock')
    with pytest.raises(ValueError,match='appears to be open in Excel'):
        validate_workbook_for_write(book)

def test_workbook_write_preflight_rejects_symbolic_link_target(tmp_path):
    from pq_installer.core.validation import validate_workbook_for_write
    real=tmp_path/'real.xlsx';real.write_bytes(b'real workbook')
    link=tmp_path/'selected.xlsx'
    try:
        link.symlink_to(real)
    except OSError as exc:
        pytest.skip(f'Windows symbolic-link creation unavailable: {exc}')
    with pytest.raises(ValueError,match='symbolic link'):
        validate_workbook_for_write(link,active_file_users=lambda _path:pytest.fail('link must fail before active-use checks'))
    assert real.read_bytes()==b'real workbook' and link.is_symlink()

def test_workbook_write_preflight_rejects_active_file_use_without_owner_lock(tmp_path,monkeypatch):
    from pq_installer.core import validation
    book=tmp_path/'open-without-owner-lock.xlsx';book.write_bytes(b'disposable')
    assert not book.with_name('~$'+book.name).exists()
    monkeypatch.setattr(validation.os,'open',lambda *args,**kwargs:pytest.fail('write probe must not run for an active workbook'))
    with pytest.raises(ValueError,match='currently open or in use'):
        validation.validate_workbook_for_write(book,active_file_users=lambda path:('EXCEL.EXE',))

def test_workbook_write_preflight_fails_closed_if_active_use_check_errors(tmp_path):
    from pq_installer.core.validation import validate_workbook_for_write
    book=tmp_path/'unverifiable.xlsx';book.write_bytes(b'disposable')
    def unavailable(path):raise OSError('Restart Manager unavailable')
    with pytest.raises(ValueError,match='could not verify'):
        validate_workbook_for_write(book,active_file_users=unavailable)

def test_running_object_table_matches_only_exact_normalized_document_path(tmp_path):
    from pq_installer.core.windows_file_users import _running_object_table_users
    target=tmp_path/'Active.xlsx'
    unrelated=tmp_path/'other.xlsx'
    class Moniker:
        def __init__(self,name):self.name=name
        def GetDisplayName(self,_context,_left):return self.name
    class Enumerator:
        def __init__(self,rows):self.rows=iter(rows)
        def Next(self,count):
            try:return (next(self.rows),)
            except StopIteration:return ()
    class Table:
        def EnumRunning(self):return Enumerator((Moniker('Book1'),Moniker(str(unrelated)),Moniker(str(target))))
    class FakeCom:
        initialized=0
        uninitialized=0
        def CoInitialize(self):self.initialized+=1
        def CoUninitialize(self):self.uninitialized+=1
        def GetRunningObjectTable(self):return Table()
        def CreateBindCtx(self,_reserved):return object()
    com=FakeCom()
    assert _running_object_table_users(target,com)==('Microsoft Excel (open workbook)',)
    assert com.initialized==com.uninitialized==1

def test_running_object_table_inspection_failure_fails_closed_and_uninitializes(tmp_path):
    from pq_installer.core.windows_file_users import _running_object_table_users
    class Moniker:
        def GetDisplayName(self,_context,_left):raise OSError('unreadable moniker')
    class Enumerator:
        def __init__(self):self.first=True
        def Next(self,_count):
            if self.first:
                self.first=False
                return (Moniker(),)
            return ()
    class Table:
        def EnumRunning(self):return Enumerator()
    class FakeCom:
        initialized=0
        uninitialized=0
        def CoInitialize(self):self.initialized+=1
        def CoUninitialize(self):self.uninitialized+=1
        def GetRunningObjectTable(self):return Table()
        def CreateBindCtx(self,_reserved):return object()
    com=FakeCom()
    with pytest.raises(RuntimeError,match='inspect a running document moniker'):
        _running_object_table_users(tmp_path/'target.xlsx',com)
    assert com.initialized==com.uninitialized==1

@pytest.mark.skipif(os.name!='nt',reason='Windows Restart Manager is available only on Windows')
def test_windows_restart_manager_detects_an_open_file_without_excel_owner_lock(tmp_path):
    import subprocess,sys
    from pq_installer.core.windows_file_users import active_file_users
    book=tmp_path/'open-workbook.xlsx';book.write_bytes(b'disposable')
    child=subprocess.Popen([sys.executable,'-c',
        "import sys,time; handle=open(sys.argv[1],'rb'); print('READY',flush=True); time.sleep(20)",str(book)],
        stdout=subprocess.PIPE,stderr=subprocess.PIPE,text=True)
    try:
        assert child.stdout.readline().strip()=='READY'
        assert active_file_users(book)
        from pq_installer.core.validation import validate_workbook_for_write
        with pytest.raises(ValueError,match='currently open or in use'):
            validate_workbook_for_write(book)
    finally:
        child.terminate();child.wait(timeout=5)

def test_workbook_write_preflight_requires_existing_supported_file(tmp_path):
    from pq_installer.core.validation import validate_workbook_for_write
    with pytest.raises(ValueError,match='does not exist'):
        validate_workbook_for_write(tmp_path/'missing.xlsx')
    unsupported=tmp_path/'book.csv';unsupported.write_text('x')
    with pytest.raises(ValueError,match='xlsx or .xlsm'):
        validate_workbook_for_write(unsupported)

def test_workbook_write_preflight_rejects_read_only_file(tmp_path):
    import os
    import stat
    from pq_installer.core.validation import validate_workbook_for_write
    book=tmp_path/'book.xlsx';book.write_bytes(b'disposable')
    os.chmod(book,stat.S_IREAD)
    try:
        with pytest.raises(ValueError,match='read-only or cannot be opened for writing'):
            validate_workbook_for_write(book)
    finally:
        os.chmod(book,stat.S_IWRITE|stat.S_IREAD)

def test_backup_preserves_extension_and_bytes(tmp_path):
    wb=tmp_path/'book.xlsm'; wb.write_bytes(b'abc123')
    backup=create_timestamped_backup(wb,datetime(2026,8,16,19,30,5))
    assert backup.name=='book_backup_20260816_193005.xlsm' and backup.read_bytes()==b'abc123'

def test_folder_recursion_is_opt_in(tmp_path):
    nested=tmp_path/'nested';nested.mkdir();(nested/'1_inner.m').write_text('x')
    assert discover_query_files([tmp_path])==[]
    recursive=discover_query_files([tmp_path],recursive=True)
    assert len(recursive)==1 and recursive[0].name=='inner'


def test_worksheet_load_destinations_validate_excel_names_cells_and_table_names(tmp_path):
    from pq_installer.core.models import LoadMode
    path=tmp_path/'q.m';path.write_text('let x=1 in x')
    validate_query_sources([QuerySource(path,'Q',load_mode=LoadMode.NEW_SHEET,
                                        sheet_name='Output',table_name='PQ_Output')])
    validate_query_sources([QuerySource(path,'Q',load_mode=LoadMode.EXISTING_SHEET,
                                        sheet_name='Results',cell='$D$5',table_name='PQ_Results')])
    for source in (
        QuerySource(path,'Q',load_mode=LoadMode.EXISTING_SHEET,sheet_name='Results'),
        QuerySource(path,'Q',load_mode=LoadMode.EXISTING_SHEET,sheet_name='Bad/Name',cell='A1'),
        QuerySource(path,'Q',load_mode=LoadMode.EXISTING_SHEET,sheet_name='Results',cell='XFE1'),
        QuerySource(path,'Q',load_mode=LoadMode.NEW_SHEET,cell='A1'),
        QuerySource(path,'Q',load_mode=LoadMode.NEW_SHEET,table_name='Bad Table'),
    ):
        with pytest.raises(ValueError):
            validate_query_sources([source])

def test_workbook_marked_from_internet_is_refused_before_excel_automation(tmp_path):
    from pq_installer.core.validation import validate_workbook

    workbook = tmp_path / 'internet.xlsx'
    workbook.write_bytes(b'disposable workbook')
    with Path(str(workbook) + ':Zone.Identifier').open('w', encoding='ascii') as stream:
        stream.write('[ZoneTransfer]\r\nZoneId=3\r\n')
    with pytest.raises(ValueError, match='Internet or Restricted zone'):
        validate_workbook(workbook)


def test_local_intranet_mark_is_not_treated_as_internet_protected_view(tmp_path):
    from pq_installer.core.validation import validate_workbook

    workbook = tmp_path / 'intranet.xlsx'
    workbook.write_bytes(b'disposable workbook')
    with Path(str(workbook) + ':Zone.Identifier').open('w', encoding='ascii') as stream:
        stream.write('[ZoneTransfer]\r\nZoneId=2\r\n')
    assert validate_workbook(workbook) == workbook

def test_internet_zone_workbook_is_rejected_before_inspector_creates_excel(tmp_path):
    from pq_installer.core.excel_inspector import ExcelQueryInspector

    workbook = tmp_path / 'internet.xlsx'
    original = b'disposable workbook bytes'
    workbook.write_bytes(original)
    with Path(str(workbook) + ':Zone.Identifier').open('w', encoding='ascii') as stream:
        stream.write('[ZoneTransfer]\r\nZoneId=3\r\n')
    factories = []
    inspector = ExcelQueryInspector(lambda: factories.append('created'))
    with pytest.raises(ValueError, match='Internet or Restricted zone'):
        inspector.inspect(workbook)
    assert factories == []
    assert workbook.read_bytes() == original

def test_internet_zone_workbook_is_rejected_before_import_stage_or_excel(tmp_path):
    from pq_installer.core.excel_queries import ExcelQueryInstaller
    from pq_installer.core.models import ExistingQueryMode, QuerySource

    workbook = tmp_path / 'internet.xlsx'
    original = b'disposable workbook bytes'
    workbook.write_bytes(original)
    with Path(str(workbook) + ':Zone.Identifier').open('w', encoding='ascii') as stream:
        stream.write('[ZoneTransfer]\r\nZoneId=3\r\n')
    source = tmp_path / 'query.m'
    source.write_text('let Source = 1 in Source', encoding='utf-8')
    factories = []
    installer = ExcelQueryInstaller(lambda: factories.append('created'))
    with pytest.raises(ValueError, match='Internet or Restricted zone'):
        installer.install(workbook, [QuerySource(source, 'SafeQuery')], ExistingQueryMode.REPLACE)
    assert factories == []
    assert workbook.read_bytes() == original
    assert not list(tmp_path.glob('*.pqstage-*'))

def test_uninterpretable_windows_zone_stream_fails_closed(tmp_path):
    from pq_installer.core.validation import validate_workbook

    workbook = tmp_path / 'unknown-zone.xlsx'
    workbook.write_bytes(b'disposable workbook')
    with Path(str(workbook) + ':Zone.Identifier').open('w', encoding='ascii') as stream:
        stream.write('[ZoneTransfer]\r\nZoneId=2\r\nZoneId=3\r\n')
    with pytest.raises(ValueError, match='security-zone metadata is present'):
        validate_workbook(workbook)
