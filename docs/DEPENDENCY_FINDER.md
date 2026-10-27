# Dependency Finder — read-only reverse references

Dependency Finder asks **“Which files depend on this file?”** The selected
`Warehouse.accdb` or `Replenishment v1.0.xlsx` is a *target*, not the file to
inspect for outgoing connections. The tool inspects other files for incoming
references to one or more targets. It is independent of SharePoint: local
folders, mapped drives, UNC paths and locally synced folders use the same
filesystem workflow.

## Desktop workflow (source implementation)

1. Open **Dependency Finder…** from the main window.
2. Add multiple Excel/Access files or a folder of possible *consumers*. Folder
   discovery optionally includes subfolders (on by default). Select the
   candidates to inspect, or use **Select all**. Discovery itself runs on a
   background worker and can be cancelled if a mapped folder is slow.
3. Add one or more *dependency targets* with **Add target file(s)…**, or copy
   selected inventory entries using **Use selected files as targets**.
4. Choose **Run read-only dependency scan**. The target files are excluded
   from their own consumer scans. You can cancel between files.
5. Review the referencing file, query/connection or linked-table name,
   reference type, confidence and all coverage warnings. Export JSON or CSV
   to a **new** file if needed. CSV contains target-summary and coverage-issue
   rows, including when no references are found. Formula-like text in CSV
   cells is prefixed to prevent spreadsheet execution when opened in Excel.

Supported *consumer* formats are `.xlsx`, `.xlsm`, `.accdb` and `.mdb`.
Legacy `.xls` and `.xlsb` appear in discovery but explicitly report
`unsupported`; they are not treated as successfully inspected.

## CLI example

```powershell
power-query-manager-cli dependency-scan `
  --target "Z:\Warehouse\Warehouse.accdb" `
  --target "Z:\Warehouse\Replenishment v1.0.xlsx" `
  --folder "Z:\Warehouse" --recursive `
  --output "dependency-report.json"
```

Alternatively repeat `--file PATH` to inspect individual files. `--recursive`
applies to folders in the CLI and is not on by default. When `--output` is
omitted, JSON is printed to stdout. Output files must not already exist.

Exit codes: `0` when the scan completed without reported coverage issues;
`2` for invalid arguments/paths/output; `5` when some input was unsupported,
unreadable, dynamic or otherwise not fully inspected. Code `0` is **not** proof
that every possible dependency was found.

## What is inspected

| Consumer | Read-only evidence |
| --- | --- |
| `.xlsx`, `.xlsm` | Power Query M definitions embedded in `customXml` DataMashup packages, `xl/connections.xml`, external workbook relationship targets, defined names and worksheet formulas. Excel is **not launched**. |
| `.accdb`, `.mdb` | DAO linked-table `TableDef.Connect`, saved query `QueryDef.SQL`, and `QueryDef.Connect`. Requires Windows, pywin32 and a compatible installed Microsoft Access Database Engine/DAO. |

Access files are copied to a temporary local directory *before* DAO opens
them read-only; this keeps any DAO `.laccdb`/`.ldb` lock beside the disposable
copy, not the original. The temporary copy is removed on normal completion or
failure. Scanning a cloud-placeholder or mapped file can still download/hydrate
its bytes or register a network read; scanning large databases requires local
temporary disk space.

Neither Excel nor Access query content is executed. There is no refresh,
connection test, credential entry, macro execution, workbook save, Access
table read, SharePoint API call or source-file mutation.

## Confidence and incomplete coverage

- `exact_path`: a known connection/source argument or link contains an exact
  normalized path to the selected target.
- `filename_only`: a known source/connection definition mentions the same
  basename but its stored path is different or insufficient to confirm it is
  the selected file (for example a different mapped-drive alias).
- `possible_literal`: a formula, saved SQL or other string literal mentions
  the target; the scanner cannot prove the string is actually used as a
  connection. This is a lead to check manually, not a confirmed dependency.

Name collisions, relative paths, UNC/mapped-drive aliases and dynamically
constructed paths can introduce false negatives or ambiguous results. The
static M tokenizer does not fully evaluate M bindings or arbitrary custom
connectors. External apps, VBA, Access modules/macros, scripts, scheduled
jobs, encrypted/password-protected files and references outside the selected
scan scope are not comprehensively discoverable. In particular, **“no
references found” means no references found in the definitions inspected**,
not that the selected file can safely be renamed or deleted.

An absent or incompatible DAO engine, unreadable/corrupt file, dynamic M
source expression, unsupported file or size/format limit produces a visible
coverage issue. Selecting only targets, with no other files to inspect, is
also reported as incomplete. If dependencies may be critical, verify results using the
business owners and affected applications before retiring a source file.

Excel ZIP/XML metadata is bounded by individual-part limits and a 256 MiB
total parsed-metadata budget per workbook; files exceeding this budget are
reported with incomplete coverage rather than exhausting memory.

## Verification state

Portable pytest tests cover synthetic Office ZIP/Mashup metadata, dependency
matching, Access DAO fakes and isolated-copy behavior, GUI selection/worker
and CLI report behavior. **Windows Access Engine inspection against real
disposable `.accdb` and `.mdb` files, genuine Excel Power Query package
compatibility, mapped-drive behavior and the packaged executable remain
UNVERIFIED** until a Windows smoke matrix is recorded. The repository's
existing `dist/` package has not been rebuilt for this source change.

See [acceptance criteria](ACCEPTANCE_CRITERIA.md) and
[security policy](SECURITY_AND_FAILURE_POLICY.md).
