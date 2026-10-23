from dataclasses import dataclass
from enum import Enum
from pathlib import Path

class ExistingQueryMode(str, Enum):
    REPLACE = 'replace'
    SKIP = 'skip'

class LoadMode(str, Enum):
    DEFINITION_ONLY = 'definition_only'
    CONNECTION_ONLY = 'connection_only'
    NEW_SHEET = 'new_sheet'
    EXISTING_SHEET = 'existing_sheet'
    DATA_MODEL = 'data_model'
    PRESERVE_EXISTING = 'preserve_existing'
    QUERY_ONLY = 'definition_only'  # compatibility alias; never means a workbook connection/load

class QueryStatus(str, Enum):
    CREATED = 'Created'
    REPLACED = 'Replaced'
    SKIPPED = 'Skipped'
    FAILED = 'Failed'
    ROLLED_BACK = 'Rolled back'
    NOT_ATTEMPTED = 'Not attempted'

@dataclass
class QuerySource:
    path: Path
    name: str
    order: int = 0
    enabled: bool = True
    load_mode: LoadMode = LoadMode.DEFINITION_ONLY
    order_from_prefix: bool = False
    order_automatic: bool = True
    source_sha256: str | None = None
    sheet_name: str | None = None
    cell: str | None = None
    table_name: str | None = None
    profile_source_path: Path | None = None

@dataclass(frozen=True)
class InstallResult:
    query_name: str
    status: QueryStatus
    message: str = ''


@dataclass(frozen=True)
class SealedQuery:
    """Immutable source snapshot captured before Excel is opened."""
    name: str
    path: Path
    formula: str
    sha256: str
    order: int
