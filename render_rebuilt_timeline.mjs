import fs from "node:fs/promises";
import path from "node:path";
import { SpreadsheetFile } from "@oai/artifact-tool";

const input = "D:/kehoach/InviBrowser_Business_Plan_Giu_Nguyen_Format_v2.xlsx";
const output = "D:/kehoach/qa_timeline_rebuilt";
await fs.mkdir(output, { recursive: true });
const workbook = await SpreadsheetFile.importXlsx(await fs.readFile(input));

const inspection = await workbook.inspect({
  kind: "table",
  range: "'Timeline GTM'!A1:AQ9",
  include: "values,formulas",
  tableMaxRows: 12,
  tableMaxCols: 43,
  maxChars: 30000,
});
console.log(inspection.ndjson);

for (const [name, range] of [["left", "A1:K9"], ["gantt", "J1:AQ9"]]) {
  const image = await workbook.render({ sheetName: "Timeline GTM", range, scale: 1, format: "png" });
  await fs.writeFile(path.join(output, `${name}.png`), new Uint8Array(await image.arrayBuffer()));
}
