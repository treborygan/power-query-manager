# Agent C: tests, threat review, release

Own new test fixtures/harness, build/smoke scripts and evidence. Start with failure-injection cases in `docs/ACCEPTANCE_CRITERIA.md`; cover safe fallback deletion, backups, read-only, macro controls, staging atomicity, UI busy-lock, profiles, schema, ZIP extraction, load/refresh and COM process ownership. Preserve existing 18 tests; do not weaken assertions.

Create Windows-only integration harness using disposable workbooks, never a user's work files; collect Windows/Office build, screenshots/reopen checks and precise status for matrix W01..W20. Check against source and Microsoft Office COM documentation. Verify a **launched** onedir EXE, not merely a build exit code. Check dependency licenses, artifact contents, `SHA256SUMS`, lack of debug material/private paths, that no admin is necessary under permissive policies. No GitHub Actions.

Report: test cases/expected/observed, exact commands, environment, passes/blocks and gaps, package manifest and hashes. Do not mark real Excel as passed if no host is available.
