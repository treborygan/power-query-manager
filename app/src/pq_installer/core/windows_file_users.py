"""Read-only Windows Restart Manager check for processes using a workbook path."""
from __future__ import annotations

import ctypes
import os
import uuid


class _FileTime(ctypes.Structure):
    _fields_=[('low',ctypes.c_uint32),('high',ctypes.c_uint32)]


class _UniqueProcess(ctypes.Structure):
    _fields_=[('process_id',ctypes.c_uint32),('start_time',_FileTime)]


class _ProcessInfo(ctypes.Structure):
    _fields_=[('process',_UniqueProcess),('app_name',ctypes.c_wchar*256),
        ('service_short_name',ctypes.c_wchar*64),('application_type',ctypes.c_int),
        ('app_status',ctypes.c_uint32),('terminal_session_id',ctypes.c_uint32),
        ('restartable',ctypes.c_int)]


def active_file_users(path):
    """Return processes registered against the path; never shuts them down.

    Non-Windows hosts have no Restart Manager API and return an empty result. On
    Windows, Restart Manager finds ordinary file handles. Excel may release its
    OS file handle after loading a read-only workbook without creating an owner
    lock, so the COM Running Object Table is also checked for an exact workbook
    path moniker. API errors raise so callers fail closed before staging.
    """
    if os.name!='nt':return ()
    manager=ctypes.WinDLL('rstrtmgr.dll',use_last_error=True)
    start=manager.RmStartSession
    start.argtypes=[ctypes.POINTER(ctypes.c_uint32),ctypes.c_uint32,ctypes.c_wchar_p]
    start.restype=ctypes.c_uint32
    register=manager.RmRegisterResources
    register.argtypes=[ctypes.c_uint32,ctypes.c_uint32,ctypes.POINTER(ctypes.c_wchar_p),
        ctypes.c_uint32,ctypes.c_void_p,ctypes.c_uint32,ctypes.c_void_p]
    register.restype=ctypes.c_uint32
    get_list=manager.RmGetList
    get_list.argtypes=[ctypes.c_uint32,ctypes.POINTER(ctypes.c_uint32),
        ctypes.POINTER(ctypes.c_uint32),ctypes.POINTER(_ProcessInfo),ctypes.POINTER(ctypes.c_uint32)]
    get_list.restype=ctypes.c_uint32
    end=manager.RmEndSession
    end.argtypes=[ctypes.c_uint32];end.restype=ctypes.c_uint32

    session=ctypes.c_uint32()
    key=ctypes.create_unicode_buffer(uuid.uuid4().hex,33)
    result=start(ctypes.byref(session),0,key)
    if result!=0:raise RuntimeError('Windows could not start a file-use check.')
    try:
        filename=os.path.abspath(os.fspath(path))
        file_array=(ctypes.c_wchar_p*1)(filename)
        result=register(session.value,1,file_array,0,None,0,None)
        if result!=0:raise RuntimeError('Windows could not register the workbook for an active-use check.')
        needed=ctypes.c_uint32();count=ctypes.c_uint32();reasons=ctypes.c_uint32()
        result=get_list(session.value,ctypes.byref(needed),ctypes.byref(count),None,ctypes.byref(reasons))
        restart_manager_users=set()
        if result==234:
            for _ in range(3):
                count=ctypes.c_uint32(needed.value)
                entries=(_ProcessInfo*max(1,count.value))()
                result=get_list(session.value,ctypes.byref(needed),ctypes.byref(count),entries,ctypes.byref(reasons))
                if result==0:
                    restart_manager_users.update(entry.app_name or entry.service_short_name or 'Another application'
                        for entry in entries[:count.value])
                    break
                if result!=234:raise RuntimeError('Windows could not query active workbook users.')
            else:
                raise RuntimeError('The workbook user list changed repeatedly during the active-use check.')
        elif result!=0:
            raise RuntimeError('Windows could not query active workbook users.')
    finally:
        if end(session.value)!=0:
            # A failed session cleanup must not be mistaken for an unlocked workbook.
            raise RuntimeError('Windows could not close the active-use check session.')

    return tuple(sorted(restart_manager_users | set(_running_object_table_users(path))))


def _running_object_table_users(path,pythoncom_module=None):
    """Find exact open-document path monikers without binding to or calling Excel."""
    if pythoncom_module is None:
        import pythoncom as pythoncom_module
    initialized=not _com_apartment_initialized()
    table=monikers=context=rows=moniker=None
    try:
        if initialized:
            pythoncom_module.CoInitialize()
        table=pythoncom_module.GetRunningObjectTable()
        monikers=table.EnumRunning()
        context=pythoncom_module.CreateBindCtx(0)
        target=os.path.normcase(os.path.abspath(os.fspath(path)))
        matches=set()
        while True:
            rows=monikers.Next(1)
            if not rows:
                break
            moniker=rows[0]
            try:
                display_name=moniker.GetDisplayName(context,None)
            except Exception as exc:
                raise RuntimeError('Windows could not inspect a running document moniker.') from exc
            try:
                candidate=os.path.normcase(os.path.abspath(str(display_name)))
            except (OSError,TypeError,ValueError):
                continue
            if candidate==target:
                matches.add('Microsoft Excel (open workbook)')
        return tuple(sorted(matches))
    finally:
        # Release all COM proxies while this thread is still initialized.
        rows=moniker=context=monikers=table=None
        import gc
        gc.collect()
        if initialized:
            pythoncom_module.CoUninitialize()


def _com_apartment_initialized():
    """Check COM state without changing the caller thread's initialization count."""
    if os.name!='nt':
        # Portable tests supply a fake pythoncom object. There is no Windows
        # COM apartment to query on other hosts, so initialize that fake once.
        return False
    ole32=ctypes.WinDLL('ole32',use_last_error=True)
    get_token=ole32.CoGetContextToken
    get_token.argtypes=[ctypes.POINTER(ctypes.c_void_p)]
    get_token.restype=ctypes.c_long
    token=ctypes.c_void_p()
    result=int(get_token(ctypes.byref(token)))
    if result==0:
        return True
    if result==-2147221008:  # CO_E_NOTINITIALIZED
        return False
    raise RuntimeError('Windows could not determine whether this thread initialized COM.')


def ensure_no_active_file_users(path,checker=None):
    """Fail closed if Windows reports a process using ``path`` or cannot check."""
    check=checker or active_file_users
    try:users=tuple(check(path))
    except Exception as exc:
        raise RuntimeError('Windows could not verify that the workbook is unused. Close Excel and other workbook users, then retry.') from exc
    if users:
        raise RuntimeError('The workbook is currently open or in use by another application. Close it before continuing.')
