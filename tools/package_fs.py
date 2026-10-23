"""Small filesystem helpers for safe release-package publication."""
from __future__ import annotations

import os
import time
from collections.abc import Callable
from typing import TypeVar


T = TypeVar("T")
_WINDOWS_TRANSIENT_RENAME_ERRORS = {5, 32, 33}
_RETRY_DELAYS = (0.05, 0.1, 0.2, 0.4, 0.8)


def replace_directory(
    source: os.PathLike[str] | str,
    destination: os.PathLike[str] | str,
    *,
    replace: Callable[[os.PathLike[str] | str, os.PathLike[str] | str], T] = os.replace,
    sleep: Callable[[float], object] = time.sleep,
    windows: bool | None = None,
) -> T:
    """Replace a directory, retrying only known transient Windows lock errors."""
    is_windows = os.name == "nt" if windows is None else windows
    for attempt in range(len(_RETRY_DELAYS) + 1):
        try:
            return replace(source, destination)
        except OSError as exc:
            if (
                not is_windows
                or getattr(exc, "winerror", None) not in _WINDOWS_TRANSIENT_RENAME_ERRORS
                or attempt == len(_RETRY_DELAYS)
            ):
                raise
            sleep(_RETRY_DELAYS[attempt])
    raise AssertionError("unreachable")
