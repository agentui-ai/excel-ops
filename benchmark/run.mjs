#!/usr/bin/env node
/**
 *   npm run bench            table + exit 1 if the skill misses anything
 *   npm run bench -- --json  machine-readable
 *
 * Two ledgers, because either half alone lies:
 *
 *   SKILL   does the recipe this skill teaches produce the right answer?
 *           Anything under 100% is a broken promise in SKILL.md.
 *   TRAP    does the from-memory version get it WRONG?
 *           A case both sides pass is not a trap — it is documentation of
 *           something the library already handles, and it says so rather than
 *           inflating the score.
 */
import { buildAll } from "./cases.mjs";

const eq = (a, b) => JSON.stringify(a) === JSON.stringify(b);

function attempt(fn, wb) {
    try {
        return { ok: true, value: fn(wb) };
    } catch (err) {
        return { ok: false, value: `threw: ${err.message.split("\n")[0].slice(0, 60)}` };
    }
}

export async function runBenchmark() {
    const cases = await buildAll();
    const results = cases.map((c) => {
        const naive = attempt(c.naive, c.wb);
        const skilled = attempt(c.skilled, c.wb);
        const skillPass = skilled.ok && eq(skilled.value, c.expected);
        const trapReal = !(naive.ok && eq(naive.value, c.expected));
        return { name: c.name, trap: c.trap, expected: c.expected, naive, skilled, skillPass, trapReal };
    });
    const skillScore = results.filter((r) => r.skillPass).length;
    const trapScore = results.filter((r) => r.trapReal).length;
    return { results, skillScore, trapScore, total: results.length };
}

function short(v) {
    const s = typeof v === "string" ? v : JSON.stringify(v);
    return s === undefined ? "undefined" : s.length > 26 ? s.slice(0, 25) + "…" : s;
}

export function formatReport({ results, skillScore, trapScore, total }) {
    const L = [];
    L.push("");
    L.push("excel-ops benchmark — deterministic, offline, no LLM");
    L.push("");
    L.push("CASE                          SKILL  TRAP   EXPECTED            FROM MEMORY");
    for (const r of results) {
        L.push(
            [
                r.name.slice(0, 28).padEnd(28),
                (r.skillPass ? " ok  " : " FAIL").padEnd(6),
                (r.trapReal ? " yes " : " no  ").padEnd(6),
                short(r.expected).padEnd(20),
                short(r.naive.value),
            ].join(" ")
        );
    }
    L.push("");
    L.push(`SKILL  ${skillScore}/${total}  the recipes in SKILL.md produce the right answer`);
    L.push(`TRAP   ${trapScore}/${total}  the from-memory version gets it wrong`);
    if (trapScore < total) {
        const soft = results.filter((r) => !r.trapReal).map((r) => r.name);
        L.push("");
        L.push(`Not traps on this machine (both approaches agree): ${soft.join(", ")}`);
        L.push("That is information, not a failure — exceljs already handles them.");
    }
    const failed = results.filter((r) => !r.skillPass);
    if (failed.length) {
        L.push("");
        for (const f of failed) {
            L.push(`FAIL ${f.name}`);
            L.push(`  expected: ${JSON.stringify(f.expected)}`);
            L.push(`  skilled:  ${JSON.stringify(f.skilled.value)}`);
        }
    }
    L.push("");
    return L.join("\n");
}

const result = await runBenchmark();
if (process.argv.includes("--json")) {
    process.stdout.write(JSON.stringify(result, null, 2) + "\n");
} else {
    process.stdout.write(formatReport(result));
}
// The skill half is the gate. The trap half is reporting.
process.exitCode = result.skillScore === result.total ? 0 : 1;
