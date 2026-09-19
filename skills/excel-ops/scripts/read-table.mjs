/**
 * Read a rectangular table out of a worksheet that real people made.
 *
 * Every rule here exists because the naive version returns something that
 * looks right and is wrong. The behaviours are verified against exceljs 4.4
 * (see benchmark/), not assumed.
 */
import ExcelJS from "exceljs";

/** exceljs ValueType numbers, named. `cell.type` returns these. */
export const T = ExcelJS.ValueType;

/**
 * One cell → a plain JS value.
 *
 * `cell.text` is a trap for two of these: it renders a Date in the machine's
 * LOCAL timezone (a UTC midnight becomes the previous day west of Greenwich)
 * and it stringifies numbers. `cell.value` keeps the type; this normalises it.
 */
export function cellValue(cell) {
    if (!cell) return null;
    switch (cell.type) {
        case T.Null:
            return null;
        case T.Merge:
            // A merged follower already carries the master's value in 4.4, but
            // go through `master` so this stays true if that ever changes.
            return cellValue(cell.master);
        case T.Formula: {
            // A formula cell holds {formula, result}. `result` is what Excel
            // last calculated — it can be a Date, a number, a string, or an
            // error object, so run it back through the same normaliser.
            const r = cell.value?.result;
            if (r && typeof r === "object" && "error" in r) return null;
            return r === undefined ? null : normalisePrimitive(r);
        }
        case T.RichText:
            return (cell.value?.richText || []).map((p) => p.text).join("");
        case T.Hyperlink:
            return cell.value?.text ?? null;
        case T.Error:
            return null;
        default:
            return normalisePrimitive(cell.value);
    }
}

function normalisePrimitive(v) {
    if (v instanceof Date) return v;
    if (typeof v === "string") return v.trim();
    return v ?? null;
}

/**
 * "42", " 42 ", "1,234.5", "$1,234.50", "(12)" → numbers. Anything else → null.
 *
 * Numbers arriving as text is the most common defect in an exported sheet, and
 * it is silent: `sum += cell.value` concatenates strings and yields "0421218".
 */
export function asNumber(v) {
    if (typeof v === "number") return Number.isFinite(v) ? v : null;
    if (v instanceof Date) return null;
    if (typeof v !== "string") return null;
    let s = v.trim();
    if (!s) return null;
    let sign = 1;
    if (/^\(.*\)$/.test(s)) {
        sign = -1;
        s = s.slice(1, -1);
    }
    s = s.replace(/[\s,_]/g, "").replace(/^[$€£¥]/, "").replace(/%$/, "");
    if (!/^[+-]?(\d+\.?\d*|\.\d+)$/.test(s)) return null;
    const n = Number(s) * sign;
    return Number.isFinite(n) ? n : null;
}

/** Is this row visually empty? */
function rowIsBlank(row) {
    let blank = true;
    row.eachCell({ includeEmpty: false }, (cell) => {
        const v = cellValue(cell);
        if (v !== null && v !== "") blank = false;
    });
    return blank;
}

/**
 * Find the header row.
 *
 * Reports do not start at A1. They start with a title, a logo, a date, two
 * blank rows, and THEN the headers. Scanning for the first dense row of
 * distinct text beats hardcoding row 1, and `expect` pins it exactly when you
 * already know a column name.
 */
export function findHeaderRow(ws, { scan = 25, minCells = 2, expect = [] } = {}) {
    const want = expect.map((h) => String(h).trim().toLowerCase());
    let best = null;
    const limit = Math.min(ws.rowCount || scan, scan);
    for (let n = 1; n <= limit; n++) {
        const row = ws.getRow(n);
        const texts = [];
        row.eachCell({ includeEmpty: false }, (cell) => {
            const v = cellValue(cell);
            if (typeof v === "string" && v !== "") texts.push(v.toLowerCase());
        });
        if (texts.length < minCells) continue;
        if (want.length) {
            if (want.every((w) => texts.includes(w))) return n;
            continue;
        }
        const distinct = new Set(texts).size;
        // Prefer the densest row of distinct labels; ties go to the earlier row.
        if (!best || distinct > best.distinct) best = { n, distinct };
    }
    if (want.length) return null;
    return best ? best.n : null;
}

/**
 * Read a table as an array of objects.
 *
 * @param {ExcelJS.Worksheet} ws
 * @param {object} [opts]
 * @param {number} [opts.headerRow]     pin the header row instead of detecting it
 * @param {string[]} [opts.expect]      header names that must be present (pins detection)
 * @param {number} [opts.stopAfterBlank] consecutive blank rows that end the table (default 2)
 * @param {boolean} [opts.coerceNumbers] turn numeric-looking text into numbers (default true)
 * @returns {{headerRow:number, headers:string[], rows:object[], rowNumbers:number[]}}
 */
export function readTable(ws, opts = {}) {
    const {
        headerRow: pinned,
        expect = [],
        stopAfterBlank = 2,
        coerceNumbers = true,
    } = opts;

    const headerRow = pinned ?? findHeaderRow(ws, { expect });
    if (!headerRow) {
        throw new Error(
            `No header row found in "${ws.name}" within the first 25 rows. ` +
                `Pass { headerRow: <n> } or { expect: ["Some Column"] }.`
        );
    }

    // Header cells, by column index. `row.values` is 1-based with a hole at
    // [0] — iterating it with .map() or forEach silently shifts every column.
    const headers = [];
    ws.getRow(headerRow).eachCell({ includeEmpty: false }, (cell, col) => {
        const v = cellValue(cell);
        if (v === null || v === "") return;
        headers[col] = String(v);
    });
    const used = headers.map((h, i) => (h ? i : null)).filter((i) => i !== null);
    if (!used.length) throw new Error(`Row ${headerRow} of "${ws.name}" has no header labels.`);

    // Duplicate labels ("Amount", "Amount") would overwrite each other in an
    // object, so the second becomes "Amount (2)" rather than vanishing.
    const seen = new Map();
    for (const col of used) {
        const base = headers[col];
        const n = (seen.get(base) || 0) + 1;
        seen.set(base, n);
        if (n > 1) headers[col] = `${base} (${n})`;
    }

    const rows = [];
    const rowNumbers = [];
    let blankStreak = 0;
    // rowCount counts phantom rows that only carry formatting; actualRowCount
    // counts real ones. Walk to rowCount but let the blank streak stop us.
    for (let n = headerRow + 1; n <= ws.rowCount; n++) {
        const row = ws.getRow(n);
        if (rowIsBlank(row)) {
            if (++blankStreak >= stopAfterBlank) break;
            continue;
        }
        blankStreak = 0;
        const obj = {};
        for (const col of used) {
            let v = cellValue(row.getCell(col));
            if (coerceNumbers && typeof v === "string") {
                const num = asNumber(v);
                if (num !== null) v = num;
            }
            obj[headers[col]] = v;
        }
        rows.push(obj);
        rowNumbers.push(n);
    }

    return { headerRow, headers: used.map((c) => headers[c]), rows, rowNumbers };
}

/** Convenience: open a file and read one sheet (by name or 1-based index). */
export async function readTableFromFile(file, sheet = 1, opts = {}) {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(file);
    const ws = typeof sheet === "number" ? wb.worksheets[sheet - 1] : wb.getWorksheet(sheet);
    if (!ws) {
        throw new Error(
            `No sheet ${JSON.stringify(sheet)} in ${file}. Sheets: ${wb.worksheets.map((s) => s.name).join(", ")}`
        );
    }
    return readTable(ws, opts);
}
