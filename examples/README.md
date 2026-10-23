# Safe sample fixtures

Example sources are in `examples/sources/` and intended for tests or disposable Excel workbooks. Queries use in-memory synthetic data, so no customer datasets or connections are required. `examples/profiles/demonstration.profile.json` references a **nonexistent placeholder** workbook and defaults to dry run. Replace that path with a disposable `.xlsx` and ensure source_path references resolve relative to the profile JSON folder.

`03_Report.pq` deliberately references `QueryHelper` as a cross-query dependency. `05_CommentSafety.pq` contains a false positive inside a comment and string. All `.m/.pq` are standalone query expressions; none contain section documents.
