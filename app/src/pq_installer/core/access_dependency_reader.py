"""Inspect Access link metadata without opening or changing the source database.

DAO may create .laccdb/.ldb lock files when a database is opened, including a
read-only open. Always hand DAO a disposable local copy instead of the input.
No records are fetched, queries run, or links refreshed.
"""
from __future__ import annotations

from pathlib import Path
import shutil
import tempfile


class AccessEngineUnavailable(RuntimeError):
    """A compatible Access DAO engine is not installed/registered."""


def _default_engine():
    try:
        import win32com.client
    except ImportError as exc:
        raise AccessEngineUnavailable("Microsoft Access DAO requires Windows and pywin32.") from exc
    for progid in ("DAO.DBEngine.160", "DAO.DBEngine.120", "DAO.DBEngine.36"):
        try:
            return win32com.client.Dispatch(progid)
        except Exception:
            continue
    raise AccessEngineUnavailable("No compatible Microsoft Access DAO engine is registered.")


def read_access_references(path, *, engine_factory=None):
    """Return only source definitions (linked tables and saved queries).

    A local temporary copy is used even if the caller requested read-only.
    All COM work happens on the calling thread; the GUI calls this in its worker.
    """
    from pq_installer.core.reverse_dependencies import DependencyReference

    source = Path(path)
    if source.suffix.casefold() not in {".mdb", ".accdb"}:
        raise ValueError("Expected an Access .mdb or .accdb file.")
    com = None
    if engine_factory is None:
        try:
            import pythoncom
        except ImportError as exc:
            raise AccessEngineUnavailable("Microsoft Access DAO requires Windows and pywin32.") from exc
        pythoncom.CoInitialize()
        com = pythoncom
    try:
        with tempfile.TemporaryDirectory(prefix="pqm-access-scan-") as temporary:
            copy = Path(temporary) / ("scan" + source.suffix)
            shutil.copyfile(source, copy)
            engine = None
            database = None
            table = None
            query = None
            try:
                engine = (engine_factory or _default_engine)()
                database = engine.OpenDatabase(str(copy), False, True)
                references = []
                for table in database.TableDefs:
                    connect = str(getattr(table, "Connect", "") or "")
                    if connect:
                        references.append(DependencyReference(
                            "Access linked table", str(table.Name), connect, True
                        ))
                for query in database.QueryDefs:
                    name = str(query.Name)
                    connect = str(getattr(query, "Connect", "") or "")
                    sql = str(getattr(query, "SQL", "") or "")
                    if connect:
                        references.append(DependencyReference("Access query connection", name, connect, True))
                    if sql:
                        references.append(DependencyReference("Access saved query", name, sql, False))
                return tuple(references)
            finally:
                # Release linked/query COM proxies before closing DAO; and
                # release the engine before Windows deletes the local copy.
                table = None
                query = None
                try:
                    if database is not None:
                        database.Close()
                finally:
                    database = None
                    engine = None
    finally:
        if com is not None:
            com.CoUninitialize()
