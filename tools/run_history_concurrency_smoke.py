"""Stress local history writes and snapshots from separate Windows processes."""
from __future__ import annotations

import hashlib
import multiprocessing
from pathlib import Path
import sys
import tempfile

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / 'app' / 'src'))
from pq_installer.core.history import HistoryStore


def _writer(history_path: str, workbook_path: str, worker_id: int, result_queue) -> None:
    try:
        store = HistoryStore(history_path)
        snapshot_ids = []
        for sequence in range(24):
            store.record('history_concurrency_smoke', {
                'profile_name': f'worker-{worker_id}',
                'action': 'write',
                'query_count': sequence,
                'status': 'verified',
            })
            if sequence in (0, 12):
                snapshot_ids.append(store.create_snapshot(workbook_path)['snapshot_id'])
        result_queue.put((worker_id, 'ok', snapshot_ids))
    except BaseException as exc:
        result_queue.put((worker_id, 'error', type(exc).__name__))
        raise


def _recover_writer(history_path: str, worker_id: int, barrier, result_queue) -> None:
    try:
        recover = HistoryStore._recover_corrupt_database

        def synchronized_recovery(store):
            barrier.wait(timeout=20)
            return recover(store)

        HistoryStore._recover_corrupt_database = synchronized_recovery
        store = HistoryStore(history_path)
        archive = str(store.recovery_archive) if store.recovery_archive else None
        store.record('history_recovery_concurrency_smoke', {
            'profile_name': f'recovery-worker-{worker_id}',
            'action': 'recovered',
            'status': 'verified',
        })
        result_queue.put((worker_id, 'ok', archive))
    except BaseException as exc:
        result_queue.put((worker_id, 'error', type(exc).__name__))
        raise


def main() -> int:
    context = multiprocessing.get_context('spawn')
    with tempfile.TemporaryDirectory(prefix='pqm-history-concurrency-') as temporary:
        root = Path(temporary)
        database = root / 'history.sqlite3'
        workbook = root / 'disposable.xlsx'
        workbook_bytes = b'Power Query Manager local history concurrency fixture\n'
        workbook.write_bytes(workbook_bytes)
        HistoryStore(database)

        result_queue = context.Queue()
        workers = [context.Process(target=_writer,
            args=(str(database), str(workbook), worker_id, result_queue),
            name=f'pqm-history-worker-{worker_id}') for worker_id in range(4)]
        for process in workers:
            process.start()
        try:
            for process in workers:
                process.join(45)
            stuck = [process for process in workers if process.is_alive()]
            for process in stuck:
                process.terminate()
            for process in stuck:
                process.join(5)
            if stuck:
                raise TimeoutError(f'{len(stuck)} smoke-owned history worker(s) exceeded 45 seconds.')
            failures = [(process.name, process.exitcode) for process in workers if process.exitcode != 0]
            if failures:
                raise RuntimeError(f'History writer processes failed: {failures!r}.')

            outcomes = [result_queue.get(timeout=5) for _ in workers]
            errors = [outcome for outcome in outcomes if outcome[1] != 'ok']
            if errors:
                raise RuntimeError(f'History writer reported an error: {errors!r}.')
            snapshot_ids = [snapshot_id for _worker, _status, ids in outcomes for snapshot_id in ids]
            if len(snapshot_ids) != 8 or len(set(snapshot_ids)) != 8:
                raise AssertionError(f'Expected eight distinct snapshots, got {len(snapshot_ids)}.')

            store = HistoryStore(database)
            events = store.list(limit=500)
            snapshots = store.list_snapshots(limit=500)
            expected_event_count = 4 * 24 + 8
            if len(events) != expected_event_count or len(snapshots) != 8:
                raise AssertionError(
                    f'Expected {expected_event_count} history events and 8 snapshots; '
                    f'observed {len(events)} events and {len(snapshots)} snapshots.')
            worker_events = [row for row in events if row['event'] == 'history_concurrency_smoke']
            counts = {f'worker-{worker_id}': 0 for worker_id in range(4)}
            for event in worker_events:
                profile_name = event['details'].get('profile_name')
                if profile_name not in counts:
                    raise AssertionError(f'Unexpected worker event metadata: {event!r}')
                counts[profile_name] += 1
            if counts != {f'worker-{worker_id}': 24 for worker_id in range(4)}:
                raise AssertionError(f'Per-process event counts differ: {counts!r}')

            expected_hash = hashlib.sha256(workbook_bytes).hexdigest()
            verified = 0
            for snapshot_id in snapshot_ids:
                snapshot = store.get_snapshot(snapshot_id)
                snapshot_path = Path(snapshot['snapshot_path'])
                if (snapshot['sha256'] != expected_hash or snapshot_path.read_bytes() != workbook_bytes
                        or hashlib.sha256(snapshot_path.read_bytes()).hexdigest() != expected_hash):
                    raise AssertionError(f'Snapshot failed content/hash verification: {snapshot_id}')
                verified += 1

            corrupt_database = root / 'history-recovery.sqlite3'
            corrupt_bytes = b'PQM intentionally corrupt disposable SQLite bytes\x00\xff'
            corrupt_database.write_bytes(corrupt_bytes)
            barrier = context.Barrier(2)
            recovery_workers = [context.Process(target=_recover_writer,
                args=(str(corrupt_database), worker_id, barrier, result_queue),
                name=f'pqm-history-recovery-worker-{worker_id}') for worker_id in range(2)]
            for process in recovery_workers:
                process.start()
            for process in recovery_workers:
                process.join(45)
            stuck = [process for process in recovery_workers if process.is_alive()]
            for process in stuck:
                process.terminate()
            for process in stuck:
                process.join(5)
            if stuck:
                raise TimeoutError(f'{len(stuck)} smoke-owned recovery worker(s) exceeded 45 seconds.')
            failures = [(process.name, process.exitcode) for process in recovery_workers if process.exitcode != 0]
            if failures:
                raise RuntimeError(f'Concurrent history recovery processes failed: {failures!r}.')
            recovery_outcomes = [result_queue.get(timeout=5) for _ in recovery_workers]
            if any(outcome[1] != 'ok' for outcome in recovery_outcomes):
                raise RuntimeError(f'Concurrent recovery reported an error: {recovery_outcomes!r}.')
            archives = list(root.glob('history-recovery.corrupt-*.sqlite3'))
            recovered = HistoryStore(corrupt_database)
            recovery_events = [row for row in recovered.list(limit=100)
                               if row['event'] == 'history_recovery_concurrency_smoke']
            archive_ids = [outcome[2] for outcome in recovery_outcomes if outcome[2] is not None]
            if (len(archives) != 1 or archives[0].read_bytes() != corrupt_bytes
                    or len(archive_ids) != 1 or Path(archive_ids[0]) != archives[0]
                    or len(recovery_events) != 2
                    or Path(str(corrupt_database) + '.recovery.lock').exists()):
                raise AssertionError(
                    f'Concurrent recovery did not preserve one archive and both events: '
                    f'archives={archives!r}, outcomes={recovery_outcomes!r}, events={len(recovery_events)}.')
            print(f'PASS: 4 history writers; {len(worker_events)} concurrent events; '
                  f'{verified} hash-verified snapshots; 2 synchronized corrupt-store recovery processes; '
                  f'one exact-byte archive and both recovery events; SQLite stores remained readable.')
            return 0
        finally:
            for process in workers:
                if process.is_alive():
                    process.terminate()
                    process.join(5)
            result_queue.close()
            result_queue.join_thread()


if __name__ == '__main__':
    raise SystemExit(main())
