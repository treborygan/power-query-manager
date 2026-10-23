"""Best-effort helpers for Windows UI Automation acceptance scripts."""

import time


def read_optional_operation_log(window, *, attempts=4, retry_delay=0.15, until_contains=None):
    """Read the named log widget without letting unknown UIA nodes fail a smoke."""
    if isinstance(attempts, bool) or not isinstance(attempts, int) or attempts < 1:
        raise ValueError('UI Automation retry count must be a positive integer.')
    if isinstance(retry_delay, bool) or not isinstance(retry_delay, (int, float)) or retry_delay < 0:
        raise ValueError('UI Automation retry delay must be non-negative.')
    if until_contains is not None and not isinstance(until_contains, str):
        raise ValueError('Requested UI Automation text marker must be text.')
    best = None
    for attempt in range(attempts):
        try:
            descendants = window.descendants()
        except Exception:
            descendants = ()
        for item in descendants:
            try:
                info = item.element_info
                try:
                    name = info.name
                except Exception:
                    name = ''
                try:
                    automation_id = info.automation_id
                except Exception:
                    automation_id = ''
                if name != 'Operation log output' and automation_id != 'operationLogOutput':
                    continue
                values = [item.window_text()]
                for pattern in ('value', 'text'):
                    try:
                        interface = getattr(item, f'iface_{pattern}')
                        if pattern == 'text':
                            value = interface.DocumentRange.GetText(-1)
                        else:
                            value = interface.CurrentValue
                        values.append(str(value))
                    except Exception:
                        pass
                unique = list(dict.fromkeys(value for value in values if value))
                if unique:
                    candidate = ' | '.join(unique)
                    if best is None or len(candidate) > len(best):
                        best = candidate
                    if until_contains is None or until_contains in candidate:
                        return candidate
            except Exception:
                continue
        if attempt + 1 < attempts and retry_delay:
            time.sleep(retry_delay)
    return best
