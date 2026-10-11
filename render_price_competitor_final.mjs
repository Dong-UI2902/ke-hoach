import fs from "node:fs/promises";
import path from "node:path";
import { SpreadsheetFile } from "@oai/artifact-tool";

const input = "D:/kehoach/InviBrowser_Business_Plan_Giu_Nguyen_Format_v2.xlsx";
const output = "D:/kehoach/qa_price_competitor_final";
await fs.mkdir(output, { recursive: true });
const workbook = await SpreadsheetFile.importXlsx(await fs.readFile(input));

const inspections = await Promise.all([
  workbook.inspect({ kind: "table", range: "'Giá InviBrowser'!A1:J12", include: "values,formulas", tableMaxRows: 15, tableMaxCols: 10, maxChars: 16000 }),
  workbook.inspect({ kind: "table", range: "'Đối thủ'!A34:L44", include: "values,formulas", tableMaxRows: 15, tableMaxCols: 12, maxChars: 16000 }),
  workbook.inspect({ kind: "match", searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A|#NUM!|#NULL!|#SPILL!|#CALC!", options: { useRegex: true, maxResults: 200 }, summary: "formula error scan" }),
]);
for (const item of inspections) console.log(item.ndjson);

const specs = [
  ["price", "Giá InviBrowser", "A1:J12"],
  ["competitor_keywords", "Đối thủ", "A33:L44"],
];
for (const [name, sheetName, range] of specs) {
  const image = await workbook.render({ sheetName, range, scale: 1, format: "png" });
  await fs.writeFile(path.join(output, `${name}.png`), new Uint8Array(await image.arrayBuffer()));
}
