"""Filesystem path checks for boundaries that must not follow links."""
from __future__ import annotations

import os
import stat
from pathlib import Path


_FILE_ATTRIBUTE_REPARSE_POINT = 0x0400


def is_reparse_point(path: os.PathLike | str) -> bool:
    """Detect POSIX symlinks and Windows reparse points without following them."""
    try:
        info = os.lstat(Path(path))
    except FileNotFoundError:
        return False
    attributes = getattr(info, 'st_file_attributes', 0)
    return stat.S_ISLNK(info.st_mode) or bool(attributes & _FILE_ATTRIBUTE_REPARSE_POINT)


def contains_reparse_point(path: os.PathLike | str) -> bool:
    """Check the entry and every lexical parent without resolving through it."""
    entry = Path(os.path.abspath(os.fspath(path)))
    return any(is_reparse_point(candidate) for candidate in (entry, *entry.parents))
