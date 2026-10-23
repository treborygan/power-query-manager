from pq_installer.ui.instructions import INSTRUCTIONS

def test_instructions_cover_current_workflow_and_limits():
    text=INSTRUCTIONS.lower()
    for phrase in ['.xlsx or .xlsm','.m/.pq','order','query name','source path','backup','query definition only','does not execute the m code','existing query']:
        assert phrase in text
