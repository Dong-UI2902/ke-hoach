import fs from "node:fs/promises";
import path from "node:path";
import { SpreadsheetFile } from "@oai/artifact-tool";

const input = "D:/kehoach/InviBrowser_Business_Plan_Giu_Nguyen_Format_v2.xlsx";
const output = "D:/kehoach/qa_finance_update";
await fs.mkdir(output, { recursive: true });
const workbook = await SpreadsheetFile.importXlsx(await fs.readFile(input));

const inspections = await Promise.all([
  workbook.inspect({ kind: "table", range: "'Kế hoạch tài chính Needly'!A1:D60", include: "values,formulas", tableMaxRows: 65, tableMaxCols: 4, maxChars: 18000 }),
  workbook.inspect({ kind: "table", range: "'Kế hoạch tài chính Needly（Copy）'!A1:W60", include: "values,formulas", tableMaxRows: 65, tableMaxCols: 23, maxChars: 18000 }),
  workbook.inspect({ kind: "match", searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A|#NUM!|#NULL!|#SPILL!|#CALC!", options: { useRegex: true, maxResults: 200 }, summary: "formula error scan" }),
]);
for (const item of inspections) console.log(item.ndjson);

const specs = [
  ["plan_left", "Kế hoạch tài chính Needly", "A1:D60"],
  ["plan_top", "Kế hoạch tài chính Needly", "A1:AN8"],
  ["reference_left", "Kế hoạch tài chính Needly（Copy）", "A1:D55"],
  ["reference_top", "Kế hoạch tài chính Needly（Copy）", "A1:W8"],
];
for (const [name, sheetName, range] of specs) {
  const image = await workbook.render({ sheetName, range, scale: 1, format: "png" });
  await fs.writeFile(path.join(output, `${name}.png`), new Uint8Array(await image.arrayBuffer()));
}
