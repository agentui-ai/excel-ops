#!/usr/bin/env node
/**
 * Look at the workbook BEFORE you parse it.
 *
 *   node inspect.mjs book.xlsx [--sheet "Name"] [--rows 5] [--json]
 *
 * Most bad spreadsheet code is written against an imagined file: headers on
 * row 1, one sheet, clean types. Three seconds here replaces an hour of
 * debugging a silently-wrong number.
 */
import ExcelJS from "exceljs";
import { cellValue, findHeaderRow, readTable, asNumber } from "./read-table.mjs";

const TYPE_NAMES = Object.fromEntries(
    Object.entries(ExcelJS.ValueType).map(([k, v]) => [v, k])
);

export async function inspectWorkbook(file, { sheet, sampleRows = 5 } = {}) {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(file);
    const sheets = wb.worksheets.map((ws) => ({
        name: ws.name,
        state: ws.state,
        rowCount: ws.rowCount,
        actualRowCount: ws.actualRowCount,
        columnCount: ws.columnCount,
        merges: Object.keys(ws._merges || {}).length || (ws.model?.merges || []).length,
    }));

    const target = sheet ? wb.getWorksheet(sheet) : wb.worksheets[0];
    if (!target) throw new Error(`No sheet ${JSON.stringify(sheet)}. Have: ${sheets.map((s) => s.name).join(", ")}`);

    const headerRow = findHeaderRow(target);
    let columns = [];
    let sample = [];
    if (headerRow) {
        const { headers, rows } = readTable(target, { headerRow, coerceNumbers: false });
        sample = rows.slice(0, sampleRows);
        columns = headers.map((h) => {
            const vals = rows.map((r) => r[h]).filter((v) => v !== null && v !== "");
            const kinds = new Set(
                vals.map((v) =>
                    v instanceof Date ? "date" : typeof v === "number" ? "number" : typeof v
                )
            );
            const textThatIsNumeric = vals.filter((v) => typeof v === "string" && asNumber(v) !== null).length;
            return {
                header: h,
                nonEmpty: vals.length,
                kinds: [...kinds],
                numbersStoredAsText: textThatIsNumeric,
            };
        });
    }

    return { file, sheets, inspected: target.name, headerRow, columns, sample };
}

function render(r) {
    const L = [];
    L.push(`${r.file}`);
    L.push("");
    L.push("SHEET                 ROWS  REAL  COLS  MERGES");
    for (const s of r.sheets) {
        L.push(
            `${s.name.slice(0, 20).padEnd(20)}  ${String(s.rowCount).padStart(4)}  ${String(
                s.actualRowCount
            ).padStart(4)}  ${String(s.columnCount).padStart(4)}  ${String(s.merges).padStart(6)}`
        );
    }
    L.push("");
    L.push(`Inspected "${r.inspected}" — header row: ${r.headerRow ?? "NOT FOUND"}`);
    if (!r.headerRow) {
        L.push("  Pass --sheet, or call readTable with { headerRow: <n> }.");
        return L.join("\n") + "\n";
    }
    L.push("");
    L.push("COLUMN                FILLED  TYPES              WARNING");
    for (const c of r.columns) {
        const warn =
            c.numbersStoredAsText > 0
                ? `${c.numbersStoredAsText} number(s) stored as text`
                : c.kinds.length > 1
                  ? `mixed types: ${c.kinds.join("/")}`
                  : "";
        L.push(
            `${c.header.slice(0, 20).padEnd(20)}  ${String(c.nonEmpty).padStart(6)}  ${c.kinds
                .join("/")
                .slice(0, 17)
                .padEnd(17)}  ${warn}`
        );
    }
    if (r.sample.length) {
        L.push("");
        L.push(`First ${r.sample.length} row(s):`);
        for (const row of r.sample) L.push("  " + JSON.stringify(row));
    }
    return L.join("\n") + "\n";
}

// CLI
if (import.meta.url === `file://${process.argv[1]}`) {
    const args = process.argv.slice(2);
    const file = args.find((a) => !a.startsWith("--"));
    if (!file) {
        process.stderr.write("usage: node inspect.mjs <file.xlsx> [--sheet NAME] [--rows N] [--json]\n");
        process.exit(2);
    }
    const sheet = args.includes("--sheet") ? args[args.indexOf("--sheet") + 1] : undefined;
    const sampleRows = args.includes("--rows") ? Number(args[args.indexOf("--rows") + 1]) : 5;
    const result = await inspectWorkbook(file, { sheet, sampleRows });
    process.stdout.write(args.includes("--json") ? JSON.stringify(result, null, 2) + "\n" : render(result));
}
