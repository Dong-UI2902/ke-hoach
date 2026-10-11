import fs from "node:fs/promises";
import { SpreadsheetFile } from "@oai/artifact-tool";
const wb = await SpreadsheetFile.importXlsx(await fs.readFile("C:/Users/Admin/Desktop/Bussiness Plan mẫu.xlsx"));
const ws = wb.worksheets.getItemAt(1);
ws.name = "Kế hoạch tài chính";
const out = await SpreadsheetFile.exportXlsx(wb);
await out.save("D:/kehoach/test_rename.xlsx");
console.log(wb.worksheets.items.map(s => s.name));
