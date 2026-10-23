# Baseline evidence

- Date: 2026-10-21 (Europe/Paris)
- Host: Windows; Python 3.11.9 (`C:\Users\Robert\AppData\Local\Programs\Python\Python311\python.exe`); pytest 9.1.1 installed for this run.
- Excel: not yet established; no COM integration or packaged executable was run for this baseline.
- Handoff integrity: `python tools/verify_handoff.py` — PASS, 22 baseline files accounted for, hashes match.
- Initial baseline invocation `python -m pytest -q` — BLOCKED because pytest was absent (`No module named pytest`). Installed pytest 9.1.1 in the user Python environment; no app files changed.
- Baseline command from `app/`: `$env:PYTHONPATH='src'; python -m pytest -q` — PASS, 18 passed in 0.23s.

This is the measured source test baseline only; it does not establish Windows Excel/COM or executable acceptance.
