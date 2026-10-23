# Power Query Manager

Windows desktop application for managing Power Query definitions in Excel workbooks. The repository contains the Python/PySide6 source, specifications, tests, release evidence, and a compact one-folder Windows package.

## Run the packaged application

On Windows, launch [`dist/PowerQueryManager/PowerQueryManager.exe`](dist/PowerQueryManager/PowerQueryManager.exe). The adjacent `PowerQueryManagerCLI.exe` supports offline profile checks and explicitly staged operations; see the packaged manual in `dist/README.txt`.

The package is a development preview, not a production-certified release. Excel automation features require desktop Excel. Review the current limitations and release gates in [`evidence/FINAL_REPORT.md`](evidence/FINAL_REPORT.md) and [`evidence/releases/`](evidence/releases/).

## Source and development

- Application source and setup: [`app/README.md`](app/README.md)
- Architecture and design: [`docs/`](docs/)
- Profile and plan contracts: [`contracts/`](contracts/)
- Tests: `app/tests/`
- Build instructions: [`app/build.ps1`](app/build.ps1)

This repository does not configure GitHub Actions. Use disposable workbooks for Excel acceptance and consult [`docs/QA_WINDOWS_MATRIX.md`](docs/QA_WINDOWS_MATRIX.md) before evaluating Excel-specific behavior.
