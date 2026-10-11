import fs from "node:fs/promises";
import path from "node:path";
import { SpreadsheetFile } from "@oai/artifact-tool";

const input = "D:/kehoach/InviBrowser_Business_Plan_Theo_Mau.xlsx";
const output = "D:/kehoach/qa_final";
await fs.mkdir(output, { recursive: true });
const workbook = await SpreadsheetFile.importXlsx(await fs.readFile(input));
const specs = [
  ["Kế hoạch kinh doanh", "B1:H90"],
  ["Kế hoạch tài chính", "A1:AN60"],
  ["Bảng giá & gói dịch vụ", "A1:J10"],
  ["Nhân sự dự án", "A1:S11"],
  ["Đối thủ", "A1:L15"],
  ["Timeline", "A1:BE36"],
  ["Tài chính chi tiết", "B1:F10"],
  ["Doanh thu và P&L", "B1:N10"],
  ["Từ khóa & đối thủ", "A1:M10"],
];
for (const [sheetName, range] of specs) {
  const rendered = await workbook.render({
    sheetName,
    range,
    scale: 1,
    format: "png",
  });
  const safeName = sheetName.replace(/[\\/:*?"<>|]/g, "_");
  await fs.writeFile(
    path.join(output, safeName + ".png"),
    new Uint8Array(await rendered.arrayBuffer()),
  );
}
