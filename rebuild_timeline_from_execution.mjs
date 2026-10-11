import fs from "node:fs/promises";
import { SpreadsheetFile } from "@oai/artifact-tool";

const inputPath = "D:/kehoach/InviBrowser_Business_Plan_Giu_Nguyen_Format_v2.xlsx";
const outputPath = "D:/kehoach/InviBrowser_Business_Plan_Giu_Nguyen_Format_v2.timeline_rebuilt.xlsx";

const workbook = await SpreadsheetFile.importXlsx(await fs.readFile(inputPath));
const sheet = workbook.worksheets.getItem("Timeline GTM");

const oldMerges = [
  "A31:A36", "I1:AA1", "C1:C2", "AB1:BE1", "E1:E2", "AI35:AK35",
  "A7:A12", "A25:A30", "A3:A6", "B1:B2", "H1:H2", "AI23:AK23",
  "AI17:AK17", "G1:G2", "AI29:AK29", "A13:A18", "D1:D2", "F1:F2",
  "A19:A24", "A1:A2",
];
for (const range of oldMerges) sheet.unmergeCells(range);
sheet.getRange("A1:BE206").dataValidation.clear();
sheet.getRange("A1:BE206").clear({ applyTo: "all" });
sheet.freezePanes.unfreeze();
sheet.freezePanes.freezeRows(2);
sheet.freezePanes.freezeColumns(3);

sheet.mergeCells("A1:K1");
sheet.getRange("A1").values = [["TIMELINE GTM — VÒNG THỰC THI 1 (30/09–27/10/2026)"]];
sheet.getRange("L1").values = [["Sep 2026"]];
sheet.mergeCells("M1:AQ1");
sheet.getRange("M1").values = [["Oct 2026"]];

sheet.getRange("A2:K2").values = [[
  "Tuần", "Mục tiêu tuần", "ICP ưu tiên", "Thực thi chính", "Phụ trách",
  "Bắt đầu", "Kết thúc", "Số ngày", "KPI chính", "Trạng thái", "Tiến độ",
]];

const rows = [
  ["W1", "Khởi động test và tạo baseline", "Team Ads + Seller + chung", "3 nội dung chính + 1 demo nhỏ; hỏi 5 user; theo dõi nguồn và bước tạo profile đầu tiên.", "Marketing + Intern", new Date("2026-09-30T00:00:00Z"), new Date("2026-10-02T00:00:00Z"), 3, "Signup → First Value", "Bị chậm", 0],
  ["W2", "Test đủ 4 P1 ICP", "Affiliate + Seeding + Team Ads + Seller", "5 nội dung chính; mỗi ICP có ít nhất một test; follow-up user đã đăng ký; luân phiên các nhóm seeding.", "Marketing + Intern", new Date("2026-10-05T00:00:00Z"), new Date("2026-10-09T00:00:00Z"), 5, "Tạo profile + D7", "Đang làm", 0],
  ["W3", "Đào sâu organization/team", "4 P1 + bài so sánh chung", "5 nội dung về organization, handoff và permission; 5–10 trao đổi thật; thử hai cách thông điệp.", "Marketing + Intern", new Date("2026-10-12T00:00:00Z"), new Date("2026-10-16T00:00:00Z"), 5, "Signup + Feedback + D7", "Chưa bắt đầu", 0],
  ["W4", "Tăng proof/case demo", "4 P1", "4 case/demo + 1 bài giáo dục; so sánh nội dung có proof với nội dung chỉ nói pain; follow-up hội thoại cũ.", "Marketing + Intern", new Date("2026-10-19T00:00:00Z"), new Date("2026-10-23T00:00:00Z"), 5, "Tạo profile + Feature Use", "Chưa bắt đầu", 0],
  ["W5", "Chốt learning vòng 1", "Affiliate + Seeding", "2 nội dung chính; tổng hợp bốn tuần; chốt KEEP / IMPROVE / PAUSE / SCALE và bước tiếp theo.", "Marketing", new Date("2026-10-26T00:00:00Z"), new Date("2026-10-27T00:00:00Z"), 2, "D7 + Feedback + Paid signal", "Chưa bắt đầu", 0],
];
sheet.getRange("A3:K7").values = rows;
sheet.getRange("F3:G7").format.numberFormat = "dd/mm/yyyy";
sheet.getRange("K3:K7").format.numberFormat = "0%";

const start = new Date("2026-09-30T00:00:00Z");
const dates = [];
for (let index = 0; index < 32; index += 1) {
  dates.push(new Date(start.getTime() + index * 86400000));
}
sheet.getRange("L2:AQ2").values = [dates];
sheet.getRange("L2:AQ2").format.numberFormat = "dd";

const columnName = (number) => {
  let result = "";
  let value = number;
  while (value > 0) {
    value -= 1;
    result = String.fromCharCode(65 + (value % 26)) + result;
    value = Math.floor(value / 26);
  }
  return result;
};

// Grid and weekend shading.
sheet.getRange("L3:AQ7").format = {
  fill: "#FFFFFF",
  borders: {
    bottom: { style: "thin", color: "#D9E2F3" },
    right: { style: "thin", color: "#E5E7EB" },
  },
};
for (let index = 0; index < dates.length; index += 1) {
  const day = dates[index].getUTCDay();
  if (day === 0 || day === 6) {
    const column = columnName(12 + index);
    sheet.getRange(`${column}3:${column}7`).format.fill = "#F3F4F6";
  }
}

const barColors = ["#F4CCCC", "#FFE599", "#D9EAD3", "#CFE2F3", "#D9D2E9"];
for (let index = 0; index < rows.length; index += 1) {
  const startOffset = Math.round((rows[index][5].getTime() - start.getTime()) / 86400000);
  const endOffset = Math.round((rows[index][6].getTime() - start.getTime()) / 86400000);
  const startColumn = columnName(12 + startOffset);
  const endColumn = columnName(12 + endOffset);
  sheet.getRange(`${startColumn}${3 + index}:${endColumn}${3 + index}`).format.fill = barColors[index];
}

sheet.getRange("A1:K1").format = {
  fill: "#203864",
  font: { name: "Calibri", size: 12, bold: true, color: "#FFFFFF" },
  horizontalAlignment: "center",
  verticalAlignment: "middle",
  rowHeightPx: 36,
};
sheet.getRange("L1:AQ1").format = {
  fill: "#4472C4",
  font: { name: "Calibri", size: 11, bold: true, color: "#FFFFFF" },
  horizontalAlignment: "center",
  verticalAlignment: "middle",
};
sheet.getRange("A2:AQ2").format = {
  fill: "#5B9BD5",
  font: { name: "Calibri", size: 9, bold: true, color: "#FFFFFF" },
  horizontalAlignment: "center",
  verticalAlignment: "middle",
  wrapText: true,
  rowHeightPx: 42,
  borders: { bottom: { style: "thin", color: "#FFFFFF" }, right: { style: "thin", color: "#D9E2F3" } },
};
sheet.getRange("A3:K7").format = {
  fill: "#FFFFFF",
  font: { name: "Calibri", size: 10, color: "#1F2937" },
  verticalAlignment: "middle",
  wrapText: true,
  rowHeightPx: 78,
  borders: { bottom: { style: "thin", color: "#D9E2F3" }, right: { style: "thin", color: "#E5E7EB" } },
};
sheet.getRange("A4:K4").format.fill = "#F7FAFC";
sheet.getRange("A6:K6").format.fill = "#F7FAFC";
sheet.getRange("A3:A7").format.font = { name: "Calibri", size: 10, bold: true, color: "#1F4E78" };
sheet.getRange("B3:B7").format.font = { name: "Calibri", size: 10, bold: true, color: "#1F4E78" };
sheet.getRange("J3").format.fill = "#F4CCCC";
sheet.getRange("J4").format.fill = "#FFE599";
sheet.getRange("J5:J7").format.fill = "#DDEBF7";
sheet.getRange("J3:J7").dataValidation = {
  list: { source: ["Chưa bắt đầu", "Đang làm", "Bị chậm", "Hoàn thành"], inCellDropDown: true },
  allowBlank: false,
};

sheet.getRange("A1:A10").format.columnWidthPx = 55;
sheet.getRange("B1:B10").format.columnWidthPx = 205;
sheet.getRange("C1:C10").format.columnWidthPx = 175;
sheet.getRange("D1:D10").format.columnWidthPx = 325;
sheet.getRange("E1:E10").format.columnWidthPx = 125;
sheet.getRange("F1:G10").format.columnWidthPx = 92;
sheet.getRange("H1:H10").format.columnWidthPx = 58;
sheet.getRange("I1:I10").format.columnWidthPx = 150;
sheet.getRange("J1:J10").format.columnWidthPx = 100;
sheet.getRange("K1:K10").format.columnWidthPx = 70;
sheet.getRange("L1:AQ7").format.columnWidthPx = 31;

sheet.mergeCells("A9:K9");
sheet.getRange("A9").values = [["Nguồn: InviBrowser_GTM_Execution_v1.9_Clean.xlsx — sheet 01 Kế hoạch tuần. Thanh màu bám đúng ngày bắt đầu và kết thúc."]];
sheet.getRange("A9:K9").format = {
  fill: "#FFF2CC",
  font: { name: "Calibri", size: 9, italic: true, color: "#7F6000" },
  verticalAlignment: "middle",
  wrapText: true,
  rowHeightPx: 34,
};

const exported = await SpreadsheetFile.exportXlsx(workbook);
await exported.save(outputPath);
console.log(`SAVED ${outputPath}`);
