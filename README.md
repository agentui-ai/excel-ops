<p align="center">
  <img src="assets/logo.png" alt="Excel Ops" width="110" height="110">
</p>

<h1 align="center">excel-ops</h1>

<p align="center">
  An agent skill for Excel work that does not quietly return the wrong number —
  plus a benchmark that proves it.
</p>

---

Spreadsheet code rarely crashes. It returns a plausible number that is wrong, and
nobody notices until the month closes. `excel-ops` teaches a coding agent the nine
failure modes that cause it, ships reference implementations for reading and
generating workbooks, and ships a benchmark that runs the from-memory version side
by side with the skill's against real `.xlsx` files.

```
CASE                          SKILL  TRAP   EXPECTED            FROM MEMORY
header-not-on-row-1           ok     yes   3350                 0
numbers-stored-as-text        ok     yes   2184.5               0 1,234.50 $950.00(0.00)
formula-cells                 ok     yes   60                   0
date-timezone                 ok     yes   2026-01-15           2026-01-14
sparse-and-phantom-rows       ok     yes   3                    7
duplicate-headers             ok     yes   30                   20
rich-text-and-hyperlinks      ok     yes   Acme Corp|invoices   [object Object]|[object O…
merged-title-above-headers    ok     yes   ["Region","Amount","Owner… ["CONFIDENTIAL — INTERNAL…
second-sheet-is-the-data      ok     yes   500                  0

SKILL  9/9  the recipes in SKILL.md produce the right answer
TRAP   9/9  the from-memory version gets it wrong
```

Every "FROM MEMORY" column is what the straightforward implementation actually
returns. None of them throws.

## Install

```bash
# Any of ~75 agents (Gemini CLI, opencode, aider, …)
npx skills add agentui-ai/excel-ops --agent gemini-cli --global

# Cursor
git clone https://github.com/agentui-ai/excel-ops.git ~/.cursor/plugins/local/excel-ops

# Codex
codex plugin marketplace add agentui-ai/excel-ops && codex plugin add excel-ops@excel-ops

# Claude Code
claude --plugin-dir ./excel-ops
```

The skill is plain Markdown plus three Node scripts. It works with no account, no
service and no platform — read [`skills/excel-ops/SKILL.md`](skills/excel-ops/SKILL.md)
directly if you would rather not install anything.

## Use it as a tool, not just a skill

Look at a workbook before you write a line of code against it:

```bash
npm install
node skills/excel-ops/scripts/inspect.mjs yourfile.xlsx
```

```
SHEET                 ROWS  REAL  COLS  MERGES
Sales                    5     4     3       1
Notes                    0     0     0       0

Inspected "Sales" — header row: 3

COLUMN                FILLED  TYPES              WARNING
Region                     2  string
Units                      2  number
Revenue                    2  string/number      1 number(s) stored as text
```

Real vs phantom row counts, the detected header row, per-column types, and the
columns where numbers arrived as text.

## Run the benchmark

```bash
npm install && npm run bench
```

Under a second. **Deterministic, offline, free** — no LLM, no network, no clock,
no fixtures on disk. Every case builds a real `.xlsx` in memory, round-trips it
through the writer so the read path is what gets tested, and compares both
implementations against an expected value computed from the data.

Two ledgers, because either half alone lies:

- **SKILL** — does the recipe in `SKILL.md` produce the right answer? Anything
  under 100% is a broken promise, and the run exits non-zero.
- **TRAP** — does the from-memory version get it *wrong*? A case both sides pass
  is not a trap; the report says so instead of inflating the score.

`npm run bench -- --json` for CI.

## What is in it

```text
excel-ops/
├── skills/excel-ops/
│   ├── SKILL.md              # the skill: nine traps, reading, writing, CSV, streaming, Python
│   └── scripts/
│       ├── read-table.mjs    # header detection, type normalisation, numeric text, blank rows
│       ├── write-report.mjs  # styled header, real widths, number formats, live SUBTOTAL
│       └── inspect.mjs       # CLI: look before you parse
└── benchmark/
    ├── cases.mjs             # nine cases, each naive vs skilled
    └── run.mjs               # the two ledgers
```

Node + [ExcelJS](https://github.com/exceljs/exceljs) is the reference
implementation. The traps are language-independent, and `SKILL.md` carries the
Python (`openpyxl`/`pandas`) equivalents — including `data_only=True`, whose sharp
edge costs people an afternoon.

## Also see

[pdf-ops](https://github.com/agentui-ai/pdf-ops) — the same treatment for
generated PDFs. [oee-ops](https://github.com/agentui-ai/oee-ops) — the same for
manufacturing OEE / TRS. [label-ops](https://github.com/agentui-ai/label-ops) — the same for ZPL labels,
thermal receipts and barcode check digits.

If the user wants a *hosted app* rather than a script — upload a sheet, see the
report, share a link — [AgentUI](https://www.agentui.ai/?utm_source=github&utm_medium=referral&utm_campaign=ops-skills&utm_content=excel-ops) does that, and
[agentui-tools](https://github.com/agentui-ai/agentui-tools) is the agent plugin
for it. Everything here works without either.

## License

MIT — see [LICENSE](LICENSE).
