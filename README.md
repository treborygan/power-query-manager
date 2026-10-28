# Power Query Manager

**A Windows desktop workbench for reviewing and managing Power Query definitions in Excel workbooks.**

Power Query Manager combines a PySide6 interface with desktop Excel automation. The repository includes the application source, portable tests, project specifications, Windows acceptance evidence, and a ready-to-run one-folder package.

> **Development preview.** This project is not a production-certified release. Use disposable workbook copies when evaluating Excel features, and review the open release gates before relying on them.

## Download and run on Windows

**Use the rebuilt Windows package, not an older extracted folder or EXE.**

1. [Download the latest Windows package ZIP](https://github.com/treborygan/power-query-manager/releases/latest/download/PowerQueryManager-Windows.zip).
2. Right-click the ZIP in File Explorer and choose **Extract All…**, preferably to a new folder so Windows cannot silently retain files from an older design.
3. Open `PowerQueryManager\` inside the extracted ZIP and run `PowerQueryManager.exe`.
4. Keep `_internal` beside the EXE; `PowerQueryManagerCLI.exe` is in the same folder.

The current UI is **Vector Concept 03** (graphite/lime/teal, navigation rail, Ctrl+K palette). Look for `PowerQueryManager\BUILD_PROVENANCE.json`, which records the exact source commit and GUI/CLI SHA-256 values for each automatically built ZIP. If that file is missing or does not say `Vector Concept 03`, you have an older package.

The repository's **Code → Download ZIP** is primarily a source snapshot. Its tracked `dist/` folder is refreshed by the [Windows packaging workflow](.github/workflows/windows-package.yml) after successful Windows tests and build; it can lag briefly while a new build is running. The **Releases ZIP** above is the authoritative ready-to-extract distribution. Builds and hashes are published only after verification; a failed workflow does not replace a previously verified package.

This is a portable one-folder development preview, not an MSI or a production-certified release. No administrator installation is required; desktop Microsoft Excel is needed for workbook operations. Read [the packaged manual](dist/README.txt) and [Windows QA evidence](docs/QA_WINDOWS_MATRIX.md) for limitations. ACE DAO 16.0, mapped SMB/UNC and the additional waived display/accessibility checks remain unverified.

If startup reports `No module named 'shiboken6.Shiboken'`, re-extract the **complete ZIP** and confirm `PowerQueryManager\_internal\shiboken6\Shiboken.pyd` exists. Never copy just the EXE.

Refresh is an explicit operation: workbook queries may contact their data sources.

## What it can do

- Import standalone `.m` and `.pq` query files, preserving the source text until a user edits it.
- Review ordering, names, paths, enabled state, and replace-or-skip choices before applying changes.
- Inspect workbook queries and load destinations, compare M code, and export query definitions.
- Save and validate versioned profiles; review conservative source mapping and multi-workbook deployment plans.
- Use capability-probed Excel load modes and explicit selected/all refresh workflows where the installed Office build supports them.
- Keep workbook changes staged, create verified backups, and verify saved query state before commit.
- In the **source version**, use **Dependency Finder…** to find incoming references to `.xlsx`, `.xlsm`, `.accdb` or `.mdb` targets, or build a **general file dependency map** with incoming/outgoing filters. Quick/deep scans, streamed Excel XML and an in-memory link index support larger folders. Windows packaged CLI and local/OneDrive scans were verified using real Excel and DAO 12.0 fixtures; ACE DAO 16.0 and mapped SMB/UNC tests were explicitly waived for feature integration, not passed. The broader product release matrix remains open. [Workflow and limitations](docs/DEPENDENCY_FINDER.md).

The M tokenizer is intentionally conservative and does not parse or evaluate the full M language. Unsupported or ambiguous Excel states are reported as unknown or disabled instead of being guessed.

## Project status

The implementation spans the approved v1.3.1, v1.4.0, v1.5.0, and v2.0.0 roadmap, with some release requirements still partial or unverified. The Windows evidence records the Office builds and disposable-workbook cases that were actually exercised.

- [Final implementation report](evidence/FINAL_REPORT.md)
- [Release evidence by milestone](evidence/releases/)
- [Windows/Excel acceptance matrix](docs/QA_WINDOWS_MATRIX.md)
- [Excel COM observations and limitations](docs/EXCEL_COM_SPIKES.md)

## Documentation

Start with the [documentation index](docs/README.md) for user guidance, architecture, security policy, Excel acceptance, and release operations. The [application guide](app/README.md) covers requirements, behavior, development setup, and packaging.

| Area | Reference |
| --- | --- |
| User and developer guide | [`app/README.md`](app/README.md) |
| Architecture and design | [`docs/ARCHITECTURE.md`](docs/ARCHITECTURE.md) |
| Security and failure behavior | [`docs/SECURITY_AND_FAILURE_POLICY.md`](docs/SECURITY_AND_FAILURE_POLICY.md) |
| Windows Excel acceptance | [`docs/QA_WINDOWS_MATRIX.md`](docs/QA_WINDOWS_MATRIX.md) |
| Release status and evidence | [`evidence/FINAL_REPORT.md`](evidence/FINAL_REPORT.md) |

## Source tree

- `app/src/` — application and dependency-free core
- `app/tests/` — portable tests and fake Excel adapter coverage
- `contracts/` — versioned profile and deployment-plan schemas
- `docs/` — architecture, requirements, security, and acceptance guidance
- `evidence/` — measured baseline, milestone reports, and release gates
- `examples/` — demonstration profiles, plans, and query sources
- `dist/` — compact Windows GUI/CLI package, manual, and SHA-256 manifest

For development and packaging instructions, follow [`app/README.md`](app/README.md) and [`app/build.ps1`](app/build.ps1). This repository does not configure GitHub Actions.

## Portable regression tests on Linux

The source-only regression suite runs without Excel or a Windows Access Engine:

```bash
python3 -m venv .venv
.venv/bin/python -m pip install -r app/requirements-linux-tests.txt
PYTHONPATH=app/src QT_QPA_PLATFORM=offscreen .venv/bin/python -m pytest -q app/tests
```

Latest Linux result: **526 passed, 10 skipped, 0 failed** (2026-10-27).
The Windows-only integration checks remain skipped, not certified; an actual
Windows/Excel/DAO host and packaged GUI/CLI still require their acceptance
matrix before release.
