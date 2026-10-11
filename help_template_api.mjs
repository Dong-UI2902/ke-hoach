import fs from "node:fs/promises";
import { SpreadsheetFile } from "@oai/artifact-tool";
const wb = await SpreadsheetFile.importXlsx(await fs.readFile("C:/Users/Admin/Desktop/Bussiness Plan mẫu.xlsx"));
for (const q of ["worksheet rename name visibility hidden", "worksheet.delete", "worksheets.delete"]) {
  const h = await wb.help(q, { include: "index,examples,notes", maxChars: 4000 });
  console.log(`QUERY ${q}\n${h.ndjson}`);
}
