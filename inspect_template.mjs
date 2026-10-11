import fs from "node:fs/promises";
import path from "node:path";
import { SpreadsheetFile } from "@oai/artifact-tool";

const input = "C:/Users/Admin/Desktop/Bussiness Plan mẫu.xlsx";
const outDir = "D:/kehoach/template_preview";
await fs.mkdir(outDir, { recursive: true });
const bytes = await fs.readFile(input);
const wb = await SpreadsheetFile.importXlsx(bytes);
const overview = await wb.inspect({
  kind: "workbook,sheet,table,drawing,definedName",
  maxChars: 16000,
  tableMaxRows: 12,
  tableMaxCols: 12,
  tableMaxCellChars: 120,
});
console.log(overview.ndjson);
for (const ws of wb.worksheets.items) {
  const used = ws.getUsedRange();
  const address = used?.address ?? "A1";
  const detail = await wb.inspect({
    kind: "region,formula,computedStyle",
    sheetId: ws.name,
    range: address,
    maxChars: 10000,
    tableMaxRows: 60,
    tableMaxCols: 24,
  });
  console.log(`\n### ${ws.name} ${address}\n${detail.ndjson}`);
  const image = await wb.render({ sheetName: ws.name, autoCrop: "all", scale: 1, format: "png" });
  await fs.writeFile(path.join(outDir, `${ws.name.replace(/[\\/:*?"<>|]/g, "_")}.png`), new Uint8Array(await image.arrayBuffer()));
}
