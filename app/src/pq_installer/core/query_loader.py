from pathlib import Path
import re
from .models import QuerySource

SUPPORTED_EXTENSIONS = {'.m', '.pq'}
_ORDER_PREFIX = re.compile(r'^(\d+)(?:[_\-\s]+)(.+)$')

def _source_for(path: Path) -> QuerySource:
    match = _ORDER_PREFIX.match(path.stem)
    if match:
        order = int(match.group(1))
        name = match.group(2).strip()
    else:
        order = 0
        name = path.stem
    return QuerySource(path.resolve(), name, order, order_from_prefix=bool(match))

def discover_query_files(paths, *, recursive=False):
    files = []
    for raw in paths:
        path = Path(raw)
        if path.is_dir():
            candidates=path.rglob('*') if recursive else path.iterdir()
            files.extend(p for p in candidates if p.is_file() and p.suffix.lower() in SUPPORTED_EXTENSIONS)
        elif path.is_file() and path.suffix.lower() in SUPPORTED_EXTENSIONS:
            files.append(path)
    unique = {str(p.resolve()).casefold(): p.resolve() for p in files}
    sources = [_source_for(p) for p in unique.values()]
    # Give unprefixed sources explicit deterministic positions after the largest numeric prefix.
    prefixed = [s for s in sources if _ORDER_PREFIX.match(s.path.stem)]
    unprefixed = sorted((s for s in sources if s not in prefixed), key=lambda s: (s.path.name.casefold(), str(s.path).casefold()))
    next_order = max((s.order for s in prefixed), default=0) + 1
    for offset, source in enumerate(unprefixed):
        source.order = next_order + offset
    return sorted(sources, key=lambda s: (s.order, s.path.name.casefold(), str(s.path).casefold()))

def discover_m_files(paths, *, recursive=False):
    """Compatibility alias retained for existing callers and integrations."""
    return discover_query_files(paths,recursive=recursive)

def read_m_code(source: QuerySource) -> str:
    try:
        text = source.path.read_text(encoding='utf-8-sig')
    except UnicodeDecodeError as exc:
        raise ValueError(f'Cannot decode {source.path} as UTF-8. Save the Power Query source as UTF-8 and retry.') from exc
    except OSError as exc:
        raise ValueError(f'Cannot read query source {source.path}: {exc}') from exc
    if not text.strip():
        raise ValueError(f'Power Query source file is empty: {source.path}')
    return text
