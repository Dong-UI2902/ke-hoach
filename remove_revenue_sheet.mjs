import fs from "node:fs/promises";
import { SpreadsheetFile } from "@oai/artifact-tool";

const inputPath = "D:/kehoach/InviBrowser_Business_Plan_Giu_Nguyen_Format_v2.xlsx";
const outputPath = "D:/kehoach/InviBrowser_Business_Plan_Giu_Nguyen_Format_v2.no_revenue_sheet.xlsx";

const workbook = await SpreadsheetFile.importXlsx(await fs.readFile(inputPath));
const sheet = workbook.worksheets.getItem("KH doanh thu và P&L");
sheet.delete();

const exported = await SpreadsheetFile.exportXlsx(workbook);
await exported.save(outputPath);
console.log(`SAVED ${outputPath}`);
