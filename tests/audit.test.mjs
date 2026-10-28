import { test } from "node:test";
import assert from "node:assert/strict";
import ExcelJS from "exceljs";
import {
  isAuditSheet,
  prepareWorkbook,
  saveEntry,
  readEntries,
  duplicateIds,
  readDetails,
  writeDetails,
  totals,
} from "../lib/audit.mjs";
function template() {
  const w = new ExcelJS.Workbook(),
    s = w.addWorksheet("Audit");
  for (const a of ["B4", "J4"]) s.getCell(a).value = "B/C";
  for (const a of ["C4", "K4"]) s.getCell(a).value = "Product Name";
  for (let r = 5; r <= 72; r++)
    for (const [target, formula] of [
      [`F${r}`, `D${r}-E${r}`],
      [`H${r}`, `F${r}*G${r}`],
      [`N${r}`, `L${r}-M${r}`],
      [`P${r}`, `N${r}*O${r}`],
    ])
      s.getCell(target).value = { formula };
  for (const [a, formula] of [
    ["F78", 'SUMIF(H5:H72,"<0")+SUMIF(P5:P72,"<0")'],
    ["F80", 'SUMIF(H5:H76,">0")+SUMIF(P5:P76,">0")'],
    ["F82", "SUM(F78:I80)"],
  ])
    s.getCell(a).value = { formula };
  prepareWorkbook(w);
  return { w, s };
}
const product = (id = "000123", physical = 8, system = 10) => ({
  id,
  name: "Sample product پاکستان",
  physical,
  system,
  cost: 20,
});
test("duplicate IDs are blocked across both sections and whitespace/case variants", () => {
  const { s } = template();
  saveEntry(s, product("abc"));
  assert.throws(() => saveEntry(s, product(" ABC ")), /already exists/);
  s.getCell("J5").value = "right-id";
  s.getCell("K5").value = "Right side";
  assert.throws(() => saveEntry(s, product("RIGHT-ID")), /already exists/);
  assert.equal(readEntries(s).length, 2);
});
test("editing updates one existing row, retaining leading zero ID and formulas", async () => {
  const { w, s } = template();
  const key = saveEntry(s, product());
  saveEntry(s, product("000123", 14, 10), key);
  assert.equal(readEntries(s).length, 1);
  assert.equal(s.getCell("B5").value, "000123");
  assert.equal(s.getCell("F5").formula, "D5-E5");
  assert.equal(s.getCell("F5").result, 4);
  assert.equal(s.getCell("H5").result, 80);
  const copy = new ExcelJS.Workbook();
  await copy.xlsx.load(await w.xlsx.writeBuffer());
  assert.equal(copy.worksheets[0].getCell("B5").value, "000123");
  assert.equal(copy.worksheets[0].getCell("F5").formula, "D5-E5");
});
test("shortage, excess and net totals use both panels", () => {
  const { s } = template();
  saveEntry(s, product("a", 8, 10));
  saveEntry(s, product("b", 14, 10));
  assert.deepEqual(totals(readEntries(s)), { short: -40, excess: 80, net: 40 });
  assert.equal(s.getCell("F78").result, -40);
  assert.equal(s.getCell("F80").result, 80);
  assert.equal(s.getCell("F82").result, 40);
});
test("invalid quantities and missing formulas do not modify data", () => {
  const { s } = template();
  for (const cost of ["", -1, "bad", Infinity])
    assert.throws(() => saveEntry(s, { ...product(), cost }));
  s.getCell("H5").value = null;
  assert.throws(() => saveEntry(s, product()), /formula/);
  assert.equal(s.getCell("B5").value, null);
});
test("full template stops before totals and never overwrites entries", () => {
  const { s } = template();
  for (let i = 0; i < 136; i++) saveEntry(s, product(String(i)));
  assert.throws(() => saveEntry(s, product("extra")), /full/);
  assert.equal(readEntries(s).length, 136);
  assert.ok(s.getCell("F78").formula);
});
test("existing duplicate rows are reported without silently deleting them", () => {
  const { s } = template();
  s.getCell("B5").value = "one";
  s.getCell("J5").value = "ONE";
  assert.deepEqual(duplicateIds(readEntries(s)), ["ONE"]);
  assert.equal(readEntries(s).length, 2);
});
test("report metadata remains editable and unrelated cells are preserved", () => {
  const { s } = template();
  assert.ok(isAuditSheet(s));
  s.getCell("L78").value = "Signature";
  writeDetails(s, {
    title: "DAIRY AUDIT",
    date: "2026-10-29",
    store: "Department A",
    lastAudit: "Previous audit",
    preparedBy: "Auditor",
  });
  assert.equal(readDetails(s).date, "2026-10-29");
  assert.equal(s.getCell("A1").value, "DAIRY AUDIT");
  assert.equal(s.getCell("L78").value, "Signature");
});

test("export retains duplicate-highlighting rules unsupported by the XLSX writer", async () => {
  const { default: JSZip } = await import("jszip");
  const { retainFormatting, exportAudit } = await import(
    "../lib/audit-export.mjs"
  );
  const { w } = template();
  const original = await JSZip.loadAsync(await w.xlsx.writeBuffer());
  const rule =
    '<conditionalFormatting sqref="B5:B72 J5:J72"><cfRule type="duplicateValues" dxfId="0" priority="1"/></conditionalFormatting>';
  const styles =
    '<dxfs count="1"><dxf><font><color rgb="FFFF0000"/></font></dxf></dxfs>';
  original.file(
    "xl/worksheets/sheet1.xml",
    (await original.file("xl/worksheets/sheet1.xml").async("string")).replace(
      "<pageMargins",
      rule + "<pageMargins",
    ),
  );
  original.file(
    "xl/styles.xml",
    (await original.file("xl/styles.xml").async("string")).replace(
      /<dxfs\b[^>]*\/>|<dxfs\b[\s\S]*?<\/dxfs>/,
      styles,
    ),
  );
  const bytes = await original.generateAsync({ type: "uint8array" });
  const loaded = new ExcelJS.Workbook();
  await loaded.xlsx.load(bytes);
  await retainFormatting(loaded, bytes);
  prepareWorkbook(loaded);
  saveEntry(loaded.worksheets[0], product());
  const exported = await JSZip.loadAsync(await exportAudit(loaded));
  assert.ok(
    (await exported.file("xl/worksheets/sheet1.xml").async("string")).includes(
      rule,
    ),
  );
  assert.ok(
    (await exported.file("xl/styles.xml").async("string")).includes(styles),
  );
});
