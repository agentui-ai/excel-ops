# Changelog

## 0.1.0

First release. The `excel-ops` skill: nine documented failure modes in spreadsheet
work, reference implementations for reading (`read-table.mjs`), generating
(`write-report.mjs`) and diagnosing (`inspect.mjs`) workbooks, and a deterministic
benchmark that scores the skill's recipes against the from-memory version.

Behaviours were verified against exceljs 4.4 rather than assumed — including that a
merged follower cell already carries its master's value, that `eachRow` skips blanks
so row numbers are not contiguous, and that a serial number with a date format is
converted to a `Date` on read.
