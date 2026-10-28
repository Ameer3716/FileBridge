// Workbook adapter for the supplied two-panel audit template. No ERP or cloud mutation.
export const FIRST_ROW = 5,
  LAST_ROW = 72;
export const normalizeId = (value) =>
  String(value ?? "")
    .trim()
    .toUpperCase();
const text = (cell) => cell.text.trim();
const number = (cell) =>
  typeof cell.value === "number"
    ? cell.value
    : Number(cell.result ?? cell.text) || 0;
export function isAuditSheet(sheet) {
  return (
    ["B4", "J4"].every((a) => /B\s*\/\s*C/i.test(text(sheet.getCell(a)))) &&
    ["C4", "K4"].every((a) => /product name/i.test(text(sheet.getCell(a)))) &&
    ["F5", "H5", "N5", "P5", "F78", "F80", "F82"].every(
      (a) => !!sheet.getCell(a).formula,
    )
  );
}
export function readEntries(sheet) {
  const rows = [];
  for (const offset of [0, 8])
    for (let row = FIRST_ROW; row <= LAST_ROW; row++) {
      const c = (n) => sheet.getCell(row, n + offset);
      if (
        text(c(2)) ||
        text(c(3)) ||
        [4, 5, 7].some((n) => text(c(n)) !== "" && Number(text(c(n))) !== 0)
      )
        rows.push({
          key: `${offset}:${row}`,
          row,
          offset,
          id: text(c(2)),
          name: text(c(3)),
          physical: number(c(4)),
          system: number(c(5)),
          cost: number(c(7)),
          balance: number(c(4)) - number(c(5)),
          value: (number(c(4)) - number(c(5))) * number(c(7)) || 0,
        });
    }
  return rows;
}
export function duplicateIds(entries) {
  const seen = new Set(),
    duplicates = new Set();
  for (const e of entries) {
    const id = normalizeId(e.id);
    if (id && seen.has(id)) duplicates.add(id);
    seen.add(id);
  }
  return [...duplicates];
}
export function totals(entries) {
  return entries.reduce(
    (t, e) => {
      t.short += Math.min(0, e.value);
      t.excess += Math.max(0, e.value);
      t.net += e.value;
      return t;
    },
    { short: 0, excess: 0, net: 0 },
  );
}
export function readDetails(sheet) {
  const raw = text(sheet.getCell("A2"));
  const match = raw.match(/(\d{2})-(\d{2})-(\d{4})/);
  return {
    title: text(sheet.getCell("A1")),
    date: match ? `${match[3]}-${match[2]}-${match[1]}` : "",
    store: text(sheet.getCell("D2")),
    lastAudit: text(sheet.getCell("L2")),
    preparedBy: text(sheet.getCell("K78")),
  };
}
export function writeDetails(sheet, details) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(details.date) || !details.title.trim())
    throw Error("Enter a report title and audit date.");
  sheet.getCell("A1").value = details.title.trim();
  sheet.getCell("A2").value =
    "DATE:-  " + details.date.split("-").reverse().join("-");
  for (const [key, address] of [
    ["store", "D2"],
    ["lastAudit", "L2"],
    ["preparedBy", "K78"],
  ])
    sheet.getCell(address).value = details[key];
}
export function prepareWorkbook(workbook) {
  // Resolve shared formulas before editing, so each original formula remains independent.
  for (const sheet of workbook.worksheets)
    sheet.eachRow((row) =>
      row.eachCell((cell) => {
        if (cell.formula)
          cell.value = {
            formula: cell.formula,
            ...(cell.result !== undefined ? { result: cell.result } : {}),
          };
      }),
    );
  workbook.calcProperties.fullCalcOnLoad = true;
}
/** @param {any} sheet @param {any} input @param {string|null} editingKey */
export function saveEntry(sheet, input, editingKey = null) {
  const id = String(input.id).trim(),
    name = String(input.name).trim();
  if (!id || !name) throw Error("Product ID and name are required.");
  const physical =
    String(input.physical ?? "").trim() === "" ? 0 : Number(input.physical);
  const cost = String(input.cost ?? "").trim() === "" ? 0 : Number(input.cost);
  const values = [physical, input.system, cost];
  if (
    values.some(
      (v) =>
        String(v).trim() === "" || !Number.isFinite(Number(v)) || Number(v) < 0,
    )
  )
    throw Error(
      "ERP quantity is required. Quantities and unit cost must be valid non-negative numbers.",
    );
  const entries = readEntries(sheet);
  if (
    entries.some(
      (e) => normalizeId(e.id) === normalizeId(id) && e.key !== editingKey,
    )
  )
    throw Error("This Product ID already exists. Open its entry to update it.");
  let slot = editingKey ? entries.find((e) => e.key === editingKey) : null;
  if (editingKey && !slot)
    throw Error("The selected entry no longer exists. Search again.");
  if (!slot) {
    for (const offset of [0, 8]) {
      for (let row = FIRST_ROW; row <= LAST_ROW; row++) {
        if (!entries.some((e) => e.row === row && e.offset === offset)) {
          slot = { row, offset };
          break;
        }
      }
      if (slot) break;
    }
  }
  if (!slot)
    throw Error(
      "This report is full (136 products). Download it and import a new audit template.",
    );
  const { row, offset } = slot,
    c = (n) => sheet.getCell(row, n + offset);
  // Never write into a template whose calculation cells have been replaced.
  if (!c(6).formula || !c(8).formula)
    throw Error(
      "This row is missing its balance or price formula. Restore the template formulas before editing.",
    );
  c(2).value = id;
  c(2).numFmt = "@";
  c(3).value = name;
  c(4).value = physical;
  c(5).value = Number(input.system);
  c(7).value = cost;
  refreshResults(sheet);
  return `${offset}:${row}`;
}
export function refreshResults(sheet) {
  for (const offset of [0, 8])
    for (let row = FIRST_ROW; row <= LAST_ROW; row++) {
      const c = (n) => sheet.getCell(row, n + offset),
        balance = number(c(4)) - number(c(5));
      for (const [n, result] of [
        [6, balance],
        [8, balance * number(c(7)) || 0],
      ])
        if (c(n).formula) c(n).value = { formula: c(n).formula, result };
    }
  const t = totals(readEntries(sheet));
  for (const [a, result] of [
    ["F78", t.short],
    ["F80", t.excess],
    ["F82", t.net],
  ]) {
    const c = sheet.getCell(a);
    if (c.formula) c.value = { formula: c.formula, result };
  }
}
