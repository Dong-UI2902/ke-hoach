import fs from "node:fs/promises";
import { SpreadsheetFile } from "@oai/artifact-tool";

const input = "D:/kehoach/InviBrowser_Business_Plan_Giu_Nguyen_Format_v2.xlsx";
const output = "D:/kehoach/qa_timeline_final.png";
const workbook = await SpreadsheetFile.importXlsx(await fs.readFile(input));
const image = await workbook.render({ sheetName: "Timeline", range: "A1:Q30", scale: 1, format: "png" });
await fs.writeFile(output, new Uint8Array(await image.arrayBuffer()));
console.log(`SAVED ${output}`);
