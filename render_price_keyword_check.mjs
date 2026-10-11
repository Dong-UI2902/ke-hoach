import fs from "node:fs/promises";
import path from "node:path";
import { SpreadsheetFile } from "@oai/artifact-tool";

const input = "D:/kehoach/InviBrowser_Business_Plan_Giu_Nguyen_Format_v2.xlsx";
const output = "D:/kehoach/qa_price_keyword_before";
await fs.mkdir(output, { recursive: true });
const workbook = await SpreadsheetFile.importXlsx(await fs.readFile(input));

const specs = [
  ["price", "Giá Needly đơn đầu", "A1:J15"],
  ["keyword", "Phân tích từ khóa_Đối thủ cạnh tranh", "A1:M12"],
  ["competitor", "Đối thủ", "A1:L32"],
];

for (const [name, sheetName, range] of specs) {
  const image = await workbook.render({ sheetName, range, scale: 1, format: "png" });
  await fs.writeFile(path.join(output, `${name}.png`), new Uint8Array(await image.arrayBuffer()));
}
