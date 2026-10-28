import { test } from "node:test";
import assert from "node:assert/strict";
import ExcelJS from "exceljs";
import {
  productRows,
  duplicateProducts,
  loadProductCheck,
  scanProducts,
  fixProduct,
  clearProduct,
  exportCheckedProducts,
} from "../lib/product-check.mjs";
function make() {
  const workbook = new ExcelJS.Workbook(),
    sheet = workbook.addWorksheet("Products");
  sheet.addRow(["Product ID", "Product Name", "Quantity", "Value"]);
  return { workbook, sheet, dirty: false };
}
test("checks names and IDs independently, with case and whitespace normalization", () => {
  const s = make();
  s.sheet.addRow([" 001 ", "Milk", 2]);
  s.sheet.addRow(["001", "Bread", 3]);
  s.sheet.addRow(["002", "  MILK  ", 4]);
  const scan = scanProducts(s);
  assert.equal(scan.groups.length, 2);
  assert.deepEqual(
    scan.groups.map((g) => g.field),
    ["id", "name"],
  );
  assert.equal(scan.rows.length, 3);
});
test("leading-zero IDs stay distinct, blanks do not create false duplicate groups", () => {
  assert.equal(
    duplicateProducts([
      { id: "001", name: "" },
      { id: "1", name: "" },
      { id: "", name: "Other" },
    ]).length,
    0,
  );
});
test("duplicate names match Unicode width and repeated whitespace", () => {
  assert.equal(
    duplicateProducts([
      { id: "A", name: "Ｍｉｌｋ  Powder" },
      { id: "B", name: "milk powder" },
    ]).length,
    1,
  );
});
test("finds duplicates across worksheets and repeated product panels", () => {
  const s = make();
  s.sheet.addRow(["A", "Milk"]);
  const other = s.workbook.addWorksheet("Second");
  other.addRow(["B/C", "Product Name", null, "B/C", "Product Name"]);
  other.addRow(["B", "Bread", null, "A", "Eggs"]);
  const scan = scanProducts(s);
  assert.equal(scan.rows.length, 3);
  assert.equal(scan.groups[0].field, "id");
  assert.deepEqual(
    scan.groups[0].rows.map((r) => r.sheetName),
    ["Products", "Second"],
  );
});
test("corrections must clear both duplicate types before export and retain quantity, formulas and names", async () => {
  const s = make();
  s.sheet.addRow(["001", "Milk", 5, { formula: "C2*10", result: 50 }]);
  s.sheet.addRow(["001", "Milk", 3]);
  const bytes = await s.workbook.xlsx.writeBuffer();
  const loaded = await loadProductCheck(bytes);
  const row = scanProducts(loaded).rows[1];
  fixProduct(loaded, row.key, "002", "Milk");
  await assert.rejects(() => exportCheckedProducts(loaded), /Resolve all/);
  fixProduct(loaded, row.key, "002", "Bread");
  const out = new ExcelJS.Workbook();
  await out.xlsx.load(await exportCheckedProducts(loaded));
  assert.equal(out.worksheets[0].getCell("A3").value, "002");
  assert.equal(out.worksheets[0].getCell("B3").value, "Bread");
  assert.equal(out.worksheets[0].getCell("C3").value, 3);
  assert.equal(out.worksheets[0].getCell("D2").formula, "C2*10");
  assert.equal(scanProducts(loaded).groups.length, 0);
});
test("explicit clearing removes quantities but preserves formula cells and unrelated entries", () => {
  const s = make();
  s.sheet.addRow(["A", "Milk", 2, { formula: "C2*5", result: 10 }]);
  s.sheet.addRow(["A", "Milk", 8, { formula: "C3*5", result: 40 }]);
  clearProduct(s, scanProducts(s).rows[1].key);
  assert.equal(s.sheet.getCell("A3").value, null);
  assert.equal(s.sheet.getCell("C3").value, null);
  assert.equal(s.sheet.getCell("D3").formula, "C3*5");
  assert.equal(s.sheet.getCell("C2").value, 2);
  assert.equal(scanProducts(s).groups.length, 0);
});
test("unrecognized sheets are reported as unchecked and unsupported files fail clearly", async () => {
  const s = make();
  s.sheet.addRow(["A", "Milk"]);
  s.workbook.addWorksheet("Notes").getCell("A1").value = "Not a product table";
  assert.deepEqual(scanProducts(s).skipped, ["Notes"]);
  const w = new ExcelJS.Workbook();
  w.addWorksheet("Unknown").addRow(["Unrecognized"]);
  const bytes = await w.xlsx.writeBuffer();
  await assert.rejects(() => loadProductCheck(bytes), /No product table/);
});
