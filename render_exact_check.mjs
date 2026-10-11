import fs from "node:fs/promises";
import path from "node:path";
import { SpreadsheetFile } from "@oai/artifact-tool";

const input = "D:/kehoach/InviBrowser_Business_Plan_Giu_Nguyen_Format_v2.xlsx";
const output = "D:/kehoach/qa_exact_template";
await fs.mkdir(output, { recursive: true });
const workbook = await SpreadsheetFile.importXlsx(await fs.readFile(input));

const checks = await Promise.all([
  workbook.inspect({ kind: "table", range: "'Kế hoạch kinh doanh'!B1:H90", include: "values,formulas", tableMaxRows: 95, tableMaxCols: 7, maxChars: 18000 }),
  workbook.inspect({ kind: "table", range: "'Timeline'!A1:H30", include: "values,formulas", tableMaxRows: 32, tableMaxCols: 8, maxChars: 10000 }),
  workbook.inspect({ kind: "match", searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A|#NUM!|#NULL!|#SPILL!|#CALC!", options: { useRegex: true, maxResults: 200 }, summary: "formula error scan" }),
]);
for (const check of checks) console.log(check.ndjson);

const specs = [
  ["Kế hoạch kinh doanh", "B1:H90"],
  ["Kế hoạch tài chính Needly", "A1:AN12"],
  ["Giá Needly đơn đầu", "A1:J10"],
  ["Nhân sự dự án", "A1:S11"],
  ["Đối thủ", "A1:L32"],
  ["Timeline", "A1:BE30"],
];
for (const [sheetName, range] of specs) {
  const rendered = await workbook.render({ sheetName, range, scale: 1, format: "png" });
  const safe = sheetName.replace(/[\\/:*?"<>|]/g, "_");
  await fs.writeFile(path.join(output, `${safe}.png`), new Uint8Array(await rendered.arrayBuffer()));
}
