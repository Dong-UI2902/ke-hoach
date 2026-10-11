import fs from "node:fs/promises";
import { SpreadsheetFile } from "@oai/artifact-tool";

const inputPath = "D:/kehoach/InviBrowser_Business_Plan_Giu_Nguyen_Format_v2.xlsx";
const outputPath = "D:/kehoach/InviBrowser_Business_Plan_Giu_Nguyen_Format_v2.timeline_view_editing.xlsx";

const workbook = await SpreadsheetFile.importXlsx(await fs.readFile(inputPath));
const timeline = workbook.worksheets.getItem("Timeline");
timeline.freezePanes.unfreeze();
timeline.freezePanes.freezeRows(2);
timeline.freezePanes.freezeColumns(3);

const exported = await SpreadsheetFile.exportXlsx(workbook);
await exported.save(outputPath);
console.log(`SAVED ${outputPath}`);
