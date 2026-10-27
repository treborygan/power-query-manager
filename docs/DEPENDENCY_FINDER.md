# Dependency Finder — read-only reverse references and general mapping

Dependency Finder asks **“Which files depend on this file?”** The selected
`Warehouse.accdb` or `Replenishment v1.0.xlsx` is a *target*, not the file to
inspect for outgoing connections. The tool inspects other files for incoming
references to one or more targets. It is independent of SharePoint: local
folders, mapped drives, UNC paths and locally synced folders use the same
filesystem workflow. The same workspace also provides **General Dependency
Map**: all detected file-to-file relationships in the selected scan scope.

## Desktop workflow (source implementation)

1. Open **Dependency Finder…** from the main window. Choose a mode:
   - **Find dependents**: one or more specific target files (incoming links).
   - **Map all selected files**: build a reusable file dependency index
     (incoming and outgoing links, no target selection required).
2. Add multiple Excel/Access files or a folder of possible *consumers*. Folder
   discovery optionally includes subfolders (on by default). Select the
   candidates to inspect, or use **Select all**. Discovery itself runs on a
   background worker and can be cancelled if a mapped folder is slow.
3. In **Find dependents** mode, add one or more *dependency targets* with
   **Add target file(s)…**, or copy selected inventory entries using
   **Use selected files as targets**. General Map needs no targets.
4. Set **Quick** (Power Query, connections and workbook links) or **Deep**
   (also worksheet formulas and defined names). General Map defaults to Quick;
   Find Dependents defaults to Deep to preserve its earlier coverage.
5. Run the read-only scan. Find Dependents excludes target files as consumers;
   **General Map scans all selected files, including interdependent targets**.
   You can cancel between files.
6. Review the referencing file, query/connection or linked-table name,
   reference type, confidence and all coverage warnings. Export JSON or CSV
   to a **new** file if needed. CSV contains target-summary and coverage-issue
   rows, including when no references are found. Formula-like text in CSV
   cells is prefixed to prevent spreadsheet execution when opened in Excel.
   In General Map, choose **All mapped files** or select one file and filter
   **incoming**, **outgoing**, or **both**. These filters use the scan's
   in-memory index — **no file is reopened or rescanned**. **Visual file graph…**
   displays an advisory diagram. The diagram is limited to 80 file nodes for
   responsiveness; the complete edge index remains in the result table and
   JSON/CSV report.

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

# Map every selected file to other scanned files; no --target argument:
power-query-manager-cli dependency-map `
  --folder "Z:\Warehouse" --recursive --depth quick `
  --output "dependency-map.json"
```

Alternatively repeat `--file PATH` to inspect individual files. `--recursive`
applies to folders in the CLI and is not on by default. When `--output` is
omitted, JSON is printed to stdout. Output files must not already exist.
Both CLI commands accept `--depth quick|deep`; `dependency-map` defaults to
`quick` and `dependency-scan` defaults to `deep`. General Map reports internal
edges plus safe filename-only descriptions for references outside the selected
scope; it never embeds raw connection strings, SQL or M formulas in reports.

Exit codes: `0` when the scan completed without reported coverage issues;
`2` for invalid arguments/paths/output; `5` when some input was unsupported,
unreadable, dynamic or otherwise not fully inspected. Code `0` is **not** proof
that every possible dependency was found.

## What is inspected

| Consumer | Read-only evidence |
| --- | --- |
| `.xlsx`, `.xlsm` | Power Query M definitions embedded in `customXml` DataMashup packages, `xl/connections.xml`, external workbook relationship targets, defined names and worksheet formulas. Excel is **not launched**. |
| `.accdb`, `.mdb` | DAO linked-table `TableDef.Connect`, saved query `QueryDef.SQL`, and `QueryDef.Connect`. Requires Windows, pywin32 and a compatible installed Microsoft Access Database Engine/DAO. |

Excel **worksheet XML, defined names, connection XML and external-link
relationships are streamed incrementally**, with completed XML nodes released
as they are processed. Formulas are matched/discarded one at a time instead
of buffering an entire worksheet's cached values/formulas. An embedded Power
Query `customXml` DataMashup part is still read/decoded in bounded memory;
Office sheet cells are never evaluated. The General Map's compact
dependency edges, safe unresolved filenames and file inventory stay in RAM
until the window is closed (or explicitly exported).
Repeated formula matches are deduplicated **as they are read**, not held in
a large intermediate list. Peak RAM still depends on the number of *distinct*
relationships and unresolved references found and on the bounded Power Query
decoding buffers.

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
reported with incomplete coverage rather than exhausting memory. The 256 MiB
budget is **cumulative processed metadata**, not a strict peak-RAM guarantee:
Power Query package decoding and Python objects can use additional memory.
Whole Access `.mdb`/`.accdb` inputs are copied to temporary disk, not loaded
wholesale into Python RAM, but require roughly their own size in disk space.

The general map is **advisory**. Filename-only matches are tentative; if
several scanned files share that name and no exact path is available, the
mapper records an `ambiguous_reference` issue and deliberately draws no edge.
Find Dependents applies the same rule when multiple selected targets have
the same basename. A single formula or saved query that mentions several
Office files contributes a separate relationship for each safely extracted
filename; raw query text is never used as an unresolved filename.
Outside-folder sources appear as unresolved filename hints (not complete
connection strings). Dynamic VBA, other connectors, path aliases and links
outside the scan scope cannot be proven absent.
General Map CSV summaries distinguish `links_found` from `no_links` based
on both incoming and outgoing edges; dependency rows contain each consumer
and referenced target path.

## Verification state and explicit Windows waivers (2026-10-27)

The owner authorized **feature-branch integration with two outstanding checks waived**, not a claim that they passed. The Windows suite reported 537 passing tests. A clean-revision Windows harness run on `a13ff507e76e4cca80fd204c69781a033ffbed39` used real disposable DAO 12.0 `.accdb`/`.mdb` fixtures and Excel Mashup/external-workbook links, verified 12 dependency records, tested local and OneDrive-synchronized Quick/Deep scans, in-process GUI filtering/graph/JSON/CSV/progress/cancellation, and packaged CLI behavior. Source hashes, sizes and timestamps stayed unchanged; Access lock/temp files did not remain; fixture-owned Excel exited naturally in 3.23 s. Both modes found the same edges on the 31,200-external-formula-cell fixture. These results do not imply exhaustive dependency detection.

**Not validated (waived for PR #1 merge only):**
- **ACE DAO 16.0:** `DAO.DBEngine.160` was not registered on the host; the ACE 16 preflight correctly stopped without fixtures. DAO 12.0 compatibility is supporting evidence **only**, not proof for DAO 16.0.
- **Mapped SMB and UNC:** there were no connected test shares/mappings. OneDrive synchronized-folder success is not network share compatibility evidence.

The packaged GUI launched and showed the new controls; the detailed GUI filter/graph/export interactions were exercised via the dialog on the Windows host, rather than as a complete packaged-GUI user-journey automation. Other Windows/Office versions, giant real corporate workbooks, and the broader product release matrix retain their own limitations. Do not present the two waived environments as verified or use a zero-match result as proof that a file can safely be retired.

See [acceptance criteria](ACCEPTANCE_CRITERIA.md), [Windows evidence and W297 waiver](QA_WINDOWS_MATRIX.md) and [security policy](SECURITY_AND_FAILURE_POLICY.md).
