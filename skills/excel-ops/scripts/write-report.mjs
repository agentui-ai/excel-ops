/**
 * Write a spreadsheet a person can actually read: real column widths, a frozen
 * styled header, number formats that survive a round trip, and an optional
 * totals row that is marked as one.
 *
 * The default `new Worksheet()` output — 8-character columns, no header
 * distinction, raw floats — is the reason people re-format every export by
 * hand before sending it on.
 */
import ExcelJS from "exceljs";

/** Number formats worth knowing by name instead of by incantation. */
export const FMT = {
    integer: "#,##0",
    decimal: "#,##0.00",
    money: '"$"#,##0.00',
    moneyRed: '"$"#,##0.00;[Red]-"$"#,##0.00',
    percent: "0.0%",
    date: "yyyy-mm-dd",
    dateTime: "yyyy-mm-dd hh:mm",
    duration: "[h]:mm:ss",
    text: "@",
};

/**
 * @param {object} spec
 * @param {string} [spec.sheetName]
 * @param {{key:string,header:string,width?:number,format?:string,align?:string}[]} spec.columns
 * @param {object[]} spec.rows
 * @param {object} [spec.totals]  e.g. { label: "Total", columns: ["amount"] }
 * @param {boolean} [spec.autoFilter=true]
 * @returns {ExcelJS.Workbook}
 */
export function buildReport({ sheetName = "Report", columns, rows, totals, autoFilter = true }) {
    if (!Array.isArray(columns) || !columns.length) throw new Error("buildReport needs columns");
    const wb = new ExcelJS.Workbook();
    wb.created = new Date(0); // deterministic output; set a real date in production
    const ws = wb.addWorksheet(sheetName);

    ws.columns = columns.map((c) => ({
        key: c.key,
        header: c.header ?? c.key,
        width: c.width ?? autoWidth(c, rows),
        style: {
            numFmt: c.format,
            alignment: { horizontal: c.align ?? (c.format && c.format !== FMT.text ? "right" : "left") },
        },
    }));

    for (const r of rows) ws.addRow(r);

    // Header: bold on a solid fill, frozen, with a filter. Doing this after
    // addRow keeps the header style from leaking into the data rows.
    const header = ws.getRow(1);
    header.font = { bold: true, color: { argb: "FFFFFFFF" } };
    header.alignment = { vertical: "middle", horizontal: "left" };
    header.height = 20;
    header.eachCell((cell) => {
        cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF334155" } };
        cell.border = { bottom: { style: "thin", color: { argb: "FF94A3B8" } } };
    });
    ws.views = [{ state: "frozen", ySplit: 1 }];
    if (autoFilter && rows.length) {
        ws.autoFilter = { from: { row: 1, column: 1 }, to: { row: 1, column: columns.length } };
    }

    if (totals?.columns?.length) {
        const first = ws.getColumn(1).key;
        const data = { [first]: totals.label ?? "Total" };
        const row = ws.addRow(data);
        for (const key of totals.columns) {
            const col = ws.getColumn(key);
            if (!col?.number) continue;
            const letter = col.letter;
            // A live formula, not a baked number: the reader can re-sort and
            // filter and the total still means something.
            row.getCell(col.number).value = {
                formula: `SUBTOTAL(109,${letter}2:${letter}${ws.rowCount - 1})`,
            };
        }
        row.font = { bold: true };
        row.eachCell((cell) => {
            cell.border = { top: { style: "thin", color: { argb: "FF334155" } } };
        });
    }

    return wb;
}

/** Width from the widest rendered value, clamped so one long cell cannot blow out the sheet. */
function autoWidth(col, rows) {
    let max = String(col.header ?? col.key).length;
    for (const r of rows) {
        const v = r[col.key];
        const len =
            v instanceof Date ? 10 : v === null || v === undefined ? 0 : String(v).length;
        if (len > max) max = len;
    }
    return Math.min(Math.max(max + 2, 8), 50);
}

/** Convenience: build and write in one call. */
export async function writeReport(spec, file) {
    const wb = buildReport(spec);
    await wb.xlsx.writeFile(file);
    return file;
}
