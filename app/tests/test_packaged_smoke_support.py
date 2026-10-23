from pathlib import Path
import sys
import pytest

TOOLS = Path(__file__).resolve().parents[2] / 'tools'
sys.path.insert(0, str(TOOLS))

from packaged_smoke_support import read_optional_operation_log
from run_packaged_keyboard_smoke import _tab_cycle


def test_operation_log_diagnostic_does_not_fail_on_unknown_uia_control_type():
    class Window:
        def descendants(self):
            raise KeyError(None)

    assert read_optional_operation_log(Window()) is None


def test_operation_log_diagnostic_reads_named_edit_when_available():
    class ElementInfo:
        name = 'Operation log output'

    class Element:
        element_info = ElementInfo()

        def window_text(self):
            return 'Excel automation finished.'

        @property
        def iface_value(self):
            return type('Value', (), {'CurrentValue': 'Excel automation finished.'})()

    class Window:
        def descendants(self):
            return [Element()]

    assert read_optional_operation_log(Window()) == 'Excel automation finished.'


def test_operation_log_diagnostic_falls_back_to_stable_automation_id():
    class ElementInfo:
        name = ''
        automation_id = 'operationLogOutput'

    class Element:
        element_info = ElementInfo()

        def window_text(self):
            return 'Workbook saved.'

        @property
        def iface_value(self):
            raise RuntimeError('ValuePattern is not supported')

    class Window:
        def descendants(self):
            return [Element()]

    assert read_optional_operation_log(Window()) == 'Workbook saved.'


def test_operation_log_diagnostic_retries_transiently_empty_uia_tree():
    class ElementInfo:
        name = 'Operation log output'
        automation_id = 'operationLogOutput'

    class Element:
        element_info = ElementInfo()

        def window_text(self):
            return 'Operation completed.'

    class Window:
        calls = 0

        def descendants(self):
            self.calls += 1
            return [Element()] if self.calls >= 3 else []

    window = Window()
    assert read_optional_operation_log(window, attempts=3, retry_delay=0) == 'Operation completed.'
    assert window.calls == 3


def test_operation_log_diagnostic_waits_for_requested_completion_record():
    class ElementInfo:
        name = 'Operation log output'
        automation_id = 'operationLogOutput'

    class Window:
        calls = 0

        def descendants(self):
            self.calls += 1
            call = self.calls

            class Element:
                element_info = ElementInfo()

                def window_text(self):
                    return 'Working…' if call < 3 else 'Working…\nExcel automation finished.'

            return [Element()]

    window = Window()
    result = read_optional_operation_log(
        window, attempts=3, retry_delay=0, until_contains='Excel automation finished.')
    assert result == 'Working…\nExcel automation finished.'
    assert window.calls == 3


def test_packaged_keyboard_smoke_records_a_complete_tab_cycle():
    class Focus:
        def __init__(self, runtime_id, name, process_id=42):
            self.runtime_id = runtime_id
            self.CurrentName = name
            self.CurrentProcessId = process_id

        def GetRuntimeId(self):
            return self.runtime_id

    class Automation:
        def __init__(self):
            self.elements = [Focus([1], 'Start'), Focus([2], 'Edit'),
                             Focus([3], 'Save'), Focus([1], 'Start')]
            self.index = 0

        def GetFocusedElement(self):
            return self.elements[self.index]

    class Window:
        def __init__(self, automation):
            self.automation = automation

        def type_keys(self, _keys, set_foreground):
            assert set_foreground
            self.automation.index += 1

    automation = Automation()
    visited = _tab_cycle(Window(automation), automation, 42, 'test dialog')

    assert visited == ['Start', 'Edit', 'Save']


def test_packaged_keyboard_smoke_fails_if_tab_focus_leaves_manager():
    class Focus:
        CurrentName = 'Outside'
        CurrentProcessId = 99

        def GetRuntimeId(self):
            return [2]

    class Automation:
        def __init__(self):
            self.calls = 0

        def GetFocusedElement(self):
            self.calls += 1
            if self.calls == 1:
                return type('InitialFocus', (), {
                    'CurrentName': 'Start', 'CurrentProcessId': 42,
                    'GetRuntimeId': lambda self: [1],
                })()
            return Focus()

    class Window:
        def type_keys(self, _keys, set_foreground):
            assert set_foreground

    with pytest.raises(AssertionError, match='outside the packaged manager'):
        _tab_cycle(Window(), Automation(), 42, 'test dialog')


def test_packaged_keyboard_smoke_uses_tab_to_traverse_editor_and_action():
    class Focus:
        def __init__(self, runtime_id, name):
            self.runtime_id = runtime_id
            self.CurrentName = name
            self.CurrentProcessId = 42

        def GetRuntimeId(self):
            return self.runtime_id

    class Automation:
        def __init__(self):
            self.elements = [Focus([1], 'M editor'), Focus([2], 'Close'), Focus([1], 'M editor')]
            self.index = 0

        def GetFocusedElement(self):
            return self.elements[self.index]

    class Window:
        def __init__(self, automation):
            self.automation = automation
            self.keys = []

        def type_keys(self, keys, set_foreground):
            assert set_foreground
            self.keys.append(keys)
            self.automation.index += 1

    automation = Automation()
    window = Window(automation)
    visited = _tab_cycle(window, automation, 42, 'test inspector')

    assert visited == ['M editor', 'Close']
    assert window.keys == ['{TAB}', '{TAB}']
