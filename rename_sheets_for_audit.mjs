import fs from "node:fs/promises";
import { SpreadsheetFile } from "@oai/artifact-tool";

const inputPath = "D:/kehoach/InviBrowser_Business_Plan_Giu_Nguyen_Format_v2.xlsx";
const outputPath = "D:/kehoach/InviBrowser_Business_Plan_Giu_Nguyen_Format_v2.renamed.xlsx";

const workbook = await SpreadsheetFile.importXlsx(await fs.readFile(inputPath));

const renames = [
  ["Kế hoạch tài chính Needly", "Tài chính InviBrowser"],
  ["Kế hoạch tài chính Needly（Copy）", "Khung tài chính mẫu"],
  ["Đối thủ", "Đối thủ & nội dung"],
  ["Timeline", "Timeline GTM"],
];

for (const [oldName, newName] of renames) {
  workbook.worksheets.getItem(oldName).name = newName;
}

const exported = await SpreadsheetFile.exportXlsx(workbook);
await exported.save(outputPath);
console.log(`SAVED ${outputPath}`);
