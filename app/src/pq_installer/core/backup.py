from datetime import datetime
from pathlib import Path
import hashlib
import os
import shutil
import uuid


def _sha256(path):
    digest = hashlib.sha256()
    with Path(path).open('rb') as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b''):
            digest.update(chunk)
    return digest.hexdigest()

def create_timestamped_backup(workbook, now=None):
    workbook = Path(workbook)
    source_hash = _sha256(workbook)
    stamp = (now or datetime.now()).strftime('%Y%m%d_%H%M%S')
    backup = workbook.with_name(f'{workbook.stem}_backup_{stamp}{workbook.suffix}')
    staged = workbook.with_name(f'.{workbook.name}.backup-{uuid.uuid4().hex}.tmp')
    try:
        shutil.copy2(workbook, staged)
        if _sha256(workbook) != source_hash or _sha256(staged) != source_hash:
            raise RuntimeError(f'Backup verification failed for {workbook.name}; the original workbook was not modified.')
        index = 0
        while True:
            candidate = backup if index == 0 else workbook.with_name(f'{workbook.stem}_backup_{stamp}_{index}{workbook.suffix}')
            if candidate.exists():
                index += 1
                continue
            try:
                # Hard-link publication is atomic and fails if the destination
                # appeared after the existence check; unlike POSIX rename it
                # can never replace another process's backup in that race.
                os.link(staged, candidate)
            except FileExistsError:
                index += 1
                continue
            return candidate
    finally:
        try:
            staged.unlink(missing_ok=True)
        except OSError:
            pass
