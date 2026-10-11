import fs from "node:fs/promises";
import { SpreadsheetFile } from "@oai/artifact-tool";

const inputPath = "D:/kehoach/InviBrowser_Business_Plan_Giu_Nguyen_Format_v2.xlsx";
const outputPath = "D:/kehoach/InviBrowser_Business_Plan_Giu_Nguyen_Format_v2.timeline_labels.xlsx";

const workbook = await SpreadsheetFile.importXlsx(await fs.readFile(inputPath));
const sheet = workbook.worksheets.getItem("Timeline GTM");
sheet.getRange("L1").values = [["T9"]];
sheet.getRange("M1").values = [["THÁNG 10/2026"]];

const exported = await SpreadsheetFile.exportXlsx(workbook);
await exported.save(outputPath);
console.log(`SAVED ${outputPath}`);
