import JSZip from "jszip";
// ExcelJS does not write the template's duplicateValues conditional formatting.
// Retain all original conditional rules and their differential styles verbatim.
const formatting = new WeakMap();
const conditional =
  /<conditionalFormatting\b[^>]*\/>|<conditionalFormatting\b[\s\S]*?<\/conditionalFormatting>/g;
const differential = /<dxfs\b[^>]*\/>|<dxfs\b[\s\S]*?<\/dxfs>/;
export async function retainFormatting(workbook, bytes) {
  const zip = await JSZip.loadAsync(bytes),
    sheets = new Map();
  for (const name of Object.keys(zip.files))
    if (/^xl\/worksheets\/sheet\d+\.xml$/.test(name)) {
      const xml = await zip.file(name).async("string");
      sheets.set(name, xml.match(conditional)?.join("") || "");
    }
  const styles = await zip.file("xl/styles.xml")?.async("string");
  formatting.set(workbook, { sheets, dxfs: styles?.match(differential)?.[0] });
}
export async function exportAudit(workbook) {
  const buffer = await workbook.xlsx.writeBuffer(),
    original = formatting.get(workbook);
  if (!original) return buffer;
  const zip = await JSZip.loadAsync(buffer);
  for (const [name, rules] of original.sheets) {
    const part = zip.file(name);
    if (!part)
      throw Error(
        "Workbook sheet structure changed. Reimport the original audit template.",
      );
    let xml = (await part.async("string")).replace(conditional, "");
    // conditionalFormatting belongs after merges/phoneticPr and before validations.
    const following =
      /<(?:dataValidations|hyperlinks|printOptions|pageMargins|pageSetup|headerFooter|rowBreaks|colBreaks|drawing|legacyDrawing|extLst)\b/;
    const index = xml.search(following),
      at = index < 0 ? xml.lastIndexOf("</worksheet>") : index;
    xml = xml.slice(0, at) + rules + xml.slice(at);
    zip.file(name, xml);
  }
  if (original.dxfs) {
    const part = zip.file("xl/styles.xml");
    let xml = await part.async("string");
    xml = differential.test(xml)
      ? xml.replace(differential, original.dxfs)
      : xml.replace("</styleSheet>", original.dxfs + "</styleSheet>");
    zip.file("xl/styles.xml", xml);
  }
  return zip.generateAsync({ type: "uint8array", compression: "DEFLATE" });
}
