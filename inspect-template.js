const XLSX = require("xlsx");
const path = require("path");

const tp = path.resolve("./templates/报关资料模板(7.9).xlsx");

console.log("===".repeat(25));
console.log("TEMPLATE INSPECTION");
console.log("===".repeat(25));

const wb = XLSX.readFile(tp, { cellStyles: true, cellFormula: true, cellDates: true, sheetStubs: true, bookProps: true, bookViews: true, cellNF: true });

console.log("
Sheet count:", wb.SheetNames.length);
console.log("Sheets:", wb.SheetNames.join(", "));

for (const sn of wb.SheetNames) {
  const ws = wb.Sheets[sn];
  const cellKeys = Object.keys(ws).filter(k => k[0] !== "!");
  console.log("
" + "---".repeat(15));
  console.log("SHEET: " + sn);
  console.log("---".repeat(15));

  // !ref
  const ref = ws["!ref"] || "(none)";
  console.log("!ref:", ref);
  if (ref !== "(none)") {
    const rng = XLSX.utils.decode_range(ref);
    console.log("  Rows:", rng.s.r+1, "to", rng.e.r+1, "Cols:", XLSX.utils.encode_col(rng.s.c), "to", XLSX.utils.encode_col(rng.e.c));
  }

  // !merges
  console.log("
MERGED CELLS (!merges):");
  if (ws["!merges"]) {
    ws["!merges"].forEach((m,i) => {
      console.log("  ["+i+"]", XLSX.utils.encode_cell(m.s) + ":" + XLSX.utils.encode_cell(m.e), "(" + XLSX.utils.encode_range(m) + ")");
    });
  } else console.log("  (none)");

  // !cols
  console.log("
COLUMNS (!cols):");
  if (ws["!cols"]) {
    ws["!cols"].forEach((c,i) => {
      if (c && c.wch != null) console.log("  Col", XLSX.utils.encode_col(i), "width:", c.wch.toFixed(1));
    });
  } else console.log("  (none)");

  // !rows
  console.log("
ROWS (!rows):");
  if (ws["!rows"]) {
    ws["!rows"].forEach((r,i) => {
      if (r && r.hpt != null) console.log("  Row", i+1, "height:", r.hpt.toFixed(1), "pt");
    });
  } else console.log("  (none)");

  // CELL DUMP - first 40 rows
  console.log("
=== CELL DUMP (rows 1-40) ===");
  const rng = XLSX.utils.decode_range(ref === "(none)" ? "A1:A1" : ref);
  const maxRow = Math.min(rng.e.r, 39);

  for (let r = rng.s.r; r <= maxRow; r++) {
    let parts = [];
    for (let c = rng.s.c; c <= rng.e.c; c++) {
      const addr = XLSX.utils.encode_cell({r,c});
      const cell = ws[addr];
      if (cell) {
        let detail = addr + ":";
        if (cell.t) detail += " type=" + cell.t;
        if (cell.v !== undefined) detail += " v=" + JSON.stringify(cell.v).substring(0,80);
        if (cell.w !== undefined && cell.w !== "") detail += " w=" + JSON.stringify(cell.w).substring(0,80);
        if (cell.f) detail += " f=" + JSON.stringify(cell.f).substring(0,80);
        if (cell.z) detail += " nf=" + JSON.stringify(cell.z).substring(0,80);
        if (cell.s != null) detail += " s=" + cell.s;
        parts.push(detail);
      }
    }
    if (parts.length > 0) {
      console.log("Row" + (r+1) + ":", parts.join(" | "));
    } else {
      console.log("Row" + (r+1) + ": (empty)");
    }
  }

  // Summary
  console.log("
SUMMARY:");
  console.log("  Data cells:", cellKeys.length);
  const withVals = cellKeys.filter(k => ws[k].v !== undefined).length;
  const withFm = cellKeys.filter(k => ws[k].f).length;
  console.log("  With values:", withVals);
  console.log("  With formulas:", withFm);
  console.log("  Merged ranges:", ws["!merges"] ? ws["!merges"].length : 0);
}

console.log("
" + "===".repeat(25));
console.log("DONE");