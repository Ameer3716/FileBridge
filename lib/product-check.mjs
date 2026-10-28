import { isAuditSheet, prepareWorkbook, refreshResults } from "./audit.mjs";
import { retainFormatting, exportAudit } from "./audit-export.mjs";
export const normalizeProductName = (v) =>
  String(v ?? "")
    .normalize("NFKC")
    .trim()
    .replace(/\s+/g, " ")
    .toLowerCase();
const header = (v) =>
  String(v ?? "")
    .toLowerCase()
    .replace(/[^a-z0-9]/g, "");
const idHeaders = new Set([
  "bc",
  "barcode",
  "id",
  "productid",
  "productcode",
  "itemid",
  "itemcode",
  "sku",
]);
const nameHeaders = new Set([
  "name",
  "productname",
  "itemname",
  "productdescription",
  "description",
]);
/** @typedef {{key:string,sheetId:number,sheetName:string,row:number,idCol:number,nameCol:number,endCol:number,id:string,name:string}} ProductRow */
/** @returns {{rows:ProductRow[],supported:string[],skipped:string[]}} */
export function productRows(workbook) {
  const rows = [],
    supported = [];
  for (const sheet of workbook.worksheets) {
    let groups = [];
    if (isAuditSheet(sheet))
      groups = [
        { headerRow: 4, idCol: 2, nameCol: 3, endCol: 8, lastRow: 72 },
        { headerRow: 4, idCol: 10, nameCol: 11, endCol: 16, lastRow: 72 },
      ];
    else
      for (let r = 1; r <= Math.min(30, sheet.rowCount); r++) {
        const ids = [],
          names = [];
        for (let c = 1; c <= Math.min(sheet.columnCount, 256); c++) {
          const h = header(sheet.getCell(r, c).text);
          if (idHeaders.has(h)) ids.push(c);
          if (nameHeaders.has(h)) names.push(c);
        }
        for (let i = 0; i < ids.length; i++) {
          const idCol = ids[i],
            next = ids[i + 1] ?? sheet.columnCount + 1,
            nameCol =
              names.find((c) => c > idCol && c < next) ??
              (ids.length === 1 ? names[0] : undefined);
          if (nameCol)
            groups.push({
              headerRow: r,
              idCol,
              nameCol,
              endCol: next - 1,
              lastRow: sheet.rowCount,
            });
        }
        if (groups.length) break;
      }
    if (!groups.length) continue;
    supported.push(sheet.name);
    for (const g of groups)
      for (let row = g.headerRow + 1; row <= g.lastRow; row++) {
        const id = sheet.getCell(row, g.idCol).text.trim(),
          name = sheet.getCell(row, g.nameCol).text.trim();
        if (!id && !name) continue;
        if (idHeaders.has(header(id)) && nameHeaders.has(header(name)))
          continue;
        rows.push({
          key: `${sheet.id}:${row}:${g.idCol}`,
          sheetId: sheet.id,
          sheetName: sheet.name,
          row,
          idCol: g.idCol,
          nameCol: g.nameCol,
          endCol: g.endCol,
          id,
          name,
        });
      }
  }
  return {
    rows,
    supported,
    skipped: workbook.worksheets
      .filter((s) => !supported.includes(s.name))
      .map((s) => s.name),
  };
}
/** @param {ProductRow[]} rows */
export function duplicateProducts(rows) {
  /** @type {{field:string,value:string,rows:ProductRow[]}[]} */
  const groups = [];
  for (const field of ["id", "name"]) {
    const map = new Map();
    for (const row of rows) {
      const value =
        field === "name"
          ? normalizeProductName(row[field])
          : String(row[field]).trim().toUpperCase();
      if (!value) continue;
      const list = map.get(value) ?? [];
      list.push(row);
      map.set(value, list);
    }
    for (const [value, matches] of map)
      if (matches.length > 1) groups.push({ field, value, rows: matches });
  }
  return groups;
}
export function scanProducts(session) {
  const scan = productRows(session.workbook);
  return { ...scan, groups: duplicateProducts(scan.rows) };
}
export async function loadProductCheck(bytes) {
  if (bytes.byteLength > 10 * 1024 * 1024)
    throw Error(
      "Product checking supports XLSX files up to 10 MiB. Split this workbook or check it in Excel.",
    );
  const { default: ExcelJS } = await import("exceljs");
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(bytes);
  await retainFormatting(workbook, bytes);
  prepareWorkbook(workbook);
  const session = { workbook, dirty: false };
  const scan = scanProducts(session);
  if (!scan.supported.length)
    throw Error(
      "No product table found. Use Product ID / Barcode / B/C and Product Name column headings in the first 30 rows.",
    );
  return session;
}
export function fixProduct(session, key, id, name) {
  const row = productRows(session.workbook).rows.find((r) => r.key === key);
  if (!row) throw Error("Product row no longer exists.");
  if (!String(id).trim() || !String(name).trim())
    throw Error(
      "Product ID and name cannot be blank. Use Clear entry to remove an extra entry.",
    );
  const sheet = session.workbook.getWorksheet(row.sheetId);
  if (
    [row.idCol, row.nameCol].some(
      (c) =>
        sheet.getCell(row.row, c).formula || sheet.getCell(row.row, c).isMerged,
    )
  )
    throw Error(
      "This identity uses a formula or merged cell. Edit it in Excel to keep its structure intact.",
    );
  sheet.getCell(row.row, row.idCol).value = String(id).trim();
  sheet.getCell(row.row, row.idCol).numFmt = "@";
  sheet.getCell(row.row, row.nameCol).value = String(name).trim();
  session.dirty = true;
}
export function clearProduct(session, key) {
  const row = productRows(session.workbook).rows.find((r) => r.key === key);
  if (!row) throw Error("Product row no longer exists.");
  const sheet = session.workbook.getWorksheet(row.sheetId);
  if (
    [row.idCol, row.nameCol].some(
      (c) =>
        sheet.getCell(row.row, c).formula || sheet.getCell(row.row, c).isMerged,
    )
  )
    throw Error(
      "This identity uses a formula or merged cell. Clear it in Excel to preserve its structure.",
    );
  for (let c = Math.min(row.idCol, row.nameCol); c <= row.endCol; c++) {
    const cell = sheet.getCell(row.row, c);
    if (cell.formula) cell.value = { formula: cell.formula };
    else if (!cell.isMerged) cell.value = null;
  }
  if (isAuditSheet(sheet)) refreshResults(sheet);
  session.dirty = true;
}
export async function exportCheckedProducts(session) {
  const scan = scanProducts(session);
  if (scan.groups.length)
    throw Error(
      "Resolve all duplicate IDs and names before downloading the checked copy.",
    );
  return exportAudit(session.workbook);
}
