import fs from "node:fs/promises";
import path from "node:path";
import { SpreadsheetFile } from "@oai/artifact-tool";

const input = "D:/kehoach/InviBrowser_Business_Plan_Giu_Nguyen_Format_v2.xlsx";
const output = "D:/kehoach/qa_full_audit";
await fs.mkdir(output, { recursive: true });
const workbook = await SpreadsheetFile.importXlsx(await fs.readFile(input));

const tableSpecs = [
  ["Kế hoạch kinh doanh", "B1:H89", 95, 7],
  ["Tài chính InviBrowser", "A1:D60", 65, 4],
  ["Tài chính InviBrowser", "E1:AN12", 15, 36],
  ["Khung tài chính mẫu", "A1:D55", 60, 4],
  ["Giá InviBrowser", "A1:J12", 15, 10],
  ["Nhân sự dự án", "A1:S11", 15, 19],
  ["Đối thủ & nội dung", "A1:L44", 50, 12],
  ["Timeline GTM", "A1:H36", 40, 8],
  ["Timeline GTM", "I1:BE36", 40, 49],
];
for (const [sheet, range, rows, cols] of tableSpecs) {
  const result = await workbook.inspect({
    kind: "table",
    range: `'${sheet}'!${range}`,
    include: "values,formulas",
    tableMaxRows: rows,
    tableMaxCols: cols,
    maxChars: 30000,
  });
  console.log(result.ndjson);
}

const scan = await workbook.inspect({
  kind: "match",
  searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A|#NUM!|#NULL!|#SPILL!|#CALC!",
  options: { useRegex: true, maxResults: 200 },
  summary: "formula error scan",
});
console.log(scan.ndjson);

const renderSpecs = [
  ["01_business_top", "Kế hoạch kinh doanh", "A1:J40"],
  ["01_business_bottom", "Kế hoạch kinh doanh", "A41:J90"],
  ["02_finance_left", "Tài chính InviBrowser", "A1:D60"],
  ["02_finance_months", "Tài chính InviBrowser", "A1:AN12"],
  ["03_finance_template", "Khung tài chính mẫu", "A1:D55"],
  ["04_price", "Giá InviBrowser", "A1:J12"],
  ["05_people", "Nhân sự dự án", "A1:S11"],
  ["06_competitor_top", "Đối thủ & nội dung", "A1:L32"],
  ["06_competitor_bottom", "Đối thủ & nội dung", "A33:L44"],
  ["07_timeline_labels", "Timeline GTM", "A1:H36"],
  ["07_timeline_bars", "Timeline GTM", "I1:BE36"],
];
for (const [name, sheetName, range] of renderSpecs) {
  const image = await workbook.render({ sheetName, range, scale: 1, format: "png" });
  await fs.writeFile(path.join(output, `${name}.png`), new Uint8Array(await image.arrayBuffer()));
}
