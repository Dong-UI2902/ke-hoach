import fs from "node:fs/promises";
import { SpreadsheetFile } from "@oai/artifact-tool";

const inputPath = "D:/kehoach/InviBrowser_Business_Plan_Giu_Nguyen_Format_v2.xlsx";
const outputPath = "D:/kehoach/InviBrowser_Business_Plan_Giu_Nguyen_Format_v2.price_competitor_editing.xlsx";

const workbook = await SpreadsheetFile.importXlsx(await fs.readFile(inputPath));

// Update the pricing sheet with the public pricing currently shown on invibrowser.com.
const price = workbook.worksheets.getItem("Giá Needly đơn đầu");
price.name = "Giá InviBrowser";
price.getRange("A2:J48").clear({ applyTo: "contents" });

for (let row = 3; row <= 10; row += 1) {
  price.getRange(`A${row}:J${row}`).copyFrom(price.getRange("A2:J2"), "all");
}

const priceRows = [
  [1, "Free", "Free", "Người mới / nhu cầu thử", "10 hồ sơ cá nhân; Team: 5 hồ sơ/seat × 2 seat", "0đ vĩnh viễn", "Đang áp dụng", "Gói miễn phí", "0đ", "Xuất tối đa 2 hồ sơ/lần; 5 workflow, 10 lượt chạy/ngày/workflow; đầy đủ anti-detect"],
  [2, "Cá nhân", "Solo Pro", "Cá nhân chạy nhiều tài khoản", "100 hồ sơ", "182.000đ/tháng", "Đang áp dụng", "Nâng cấp từ Free", "1.950.000đ/năm", "Xuất hồ sơ, workflow và lượt chạy không giới hạn; thao tác hàng loạt; Smart Trash"],
  [3, "Đội nhóm", "Team Pro", "Team cần phân quyền", "50 hồ sơ/seat × 10 seat", "520.000đ/tháng", "Đang áp dụng", "Nâng cấp từ Free", "5.564.000đ/năm", "Phân quyền nâng cao; chặn website theo rule; export và workflow không giới hạn"],
  [4, "Trọn đời", "Lifetime", "Cá nhân/team mua một lần", "Không giới hạn hồ sơ; 50 seat", "5.200.000đ một lần", "Đang áp dụng", "Không nêu", "Không áp dụng", "Mọi tính năng cao cấp; cập nhật miễn phí trọn đời"],
  [5, "Doanh nghiệp", "Business", "Tổ chức cần triển khai riêng", "Tùy biến", "Liên hệ", "Đang áp dụng", "Theo tư vấn", "Theo báo giá", "Thiết kế riêng hạn mức, thanh toán, triển khai và hỗ trợ"],
  [6, "Mua thêm", "Seat bổ sung", "Khách hàng cần thêm thành viên", "Cộng vào gói hiện tại", "125.000đ/seat", "Đang áp dụng", "Không áp dụng", "Không áp dụng", "Cộng trực tiếp vào hạn mức, không cần nâng cấp gói"],
  [7, "Mua thêm", "1–50 hồ sơ", "Khách hàng cần thêm profile", "1–50 hồ sơ", "25.000đ/hồ sơ", "Đang áp dụng", "Không áp dụng", "Không áp dụng", "Cộng trực tiếp vào hạn mức, không cần nâng cấp gói"],
  [8, "Mua thêm", "51–100 hồ sơ", "Khách hàng cần thêm profile", "51–100 hồ sơ", "18.000đ/hồ sơ", "Đang áp dụng", "Không áp dụng", "Không áp dụng", "Cộng trực tiếp vào hạn mức, không cần nâng cấp gói"],
  [9, "Mua thêm", "Trên 100 hồ sơ", "Khách hàng cần thêm profile", "Trên 100 hồ sơ", "12.000đ/hồ sơ", "Đang áp dụng", "Không áp dụng", "Không áp dụng", "Cộng trực tiếp vào hạn mức, không cần nâng cấp gói"],
];
price.getRange("A2:J10").values = priceRows;

price.getRange("A2:J10").format = {
  fill: "#FFFFFF",
  font: { name: "Calibri", size: 10, color: "#1F2937" },
  verticalAlignment: "middle",
  horizontalAlignment: "center",
  wrapText: true,
  rowHeightPx: 72,
  borders: { bottom: { style: "thin", color: "#D9E2F3" } },
};
price.getRange("C2:D10").format.horizontalAlignment = "left";
price.getRange("J2:J10").format.horizontalAlignment = "left";
price.getRange("A3:J3").format.fill = "#F7FAFC";
price.getRange("A5:J5").format.fill = "#F7FAFC";
price.getRange("A7:J7").format.fill = "#F7FAFC";
price.getRange("A9:J9").format.fill = "#F7FAFC";
price.getRange("A2:A10").format.font = { name: "Calibri", size: 10, bold: true, color: "#203864" };
price.getRange("C2:C10").format.font = { name: "Calibri", size: 10, bold: true, color: "#203864" };
price.getRange("F2:F10").format.font = { name: "Calibri", size: 10, bold: true, color: "#C65911" };
price.getRange("A1:A12").format.columnWidthPx = 56;
price.getRange("B1:B12").format.columnWidthPx = 82;
price.getRange("C1:C12").format.columnWidthPx = 150;
price.getRange("D1:D12").format.columnWidthPx = 215;
price.getRange("E1:E12").format.columnWidthPx = 190;
price.getRange("F1:F12").format.columnWidthPx = 145;
price.getRange("G1:G12").format.columnWidthPx = 105;
price.getRange("H1:H12").format.columnWidthPx = 125;
price.getRange("I1:I12").format.columnWidthPx = 145;
price.getRange("J1:J12").format.columnWidthPx = 285;
price.getRange("A1:J1").format.rowHeightPx = 48;
price.mergeCells("A12:J12");
price.getRange("A12").values = [["Nguồn: https://invibrowser.com — truy cập 11/10/2026. Giá trên website đã gồm VAT và thanh toán bằng VNĐ."]];
price.getRange("A12:J12").format = {
  fill: "#FFF2CC",
  font: { name: "Calibri", size: 10, italic: true, color: "#7F6000" },
  verticalAlignment: "middle",
  wrapText: true,
  rowHeightPx: 36,
};

// Fold the useful keyword themes into the competitor sheet, where they support positioning decisions.
const competitor = workbook.worksheets.getItem("Đối thủ");
competitor.mergeCells("A34:L34");
competitor.getRange("A34").values = [["4. CHỦ ĐỀ TÌM KIẾM & NỘI DUNG CẠNH TRANH"]];

const groups = ["A35:B35", "C35:E35", "F35:H35", "I35:L35"];
for (const group of groups) competitor.mergeCells(group);
competitor.getRange("A35").values = [["Mức ưu tiên"]];
competitor.getRange("C35").values = [["Chủ đề / từ khóa"]];
competitor.getRange("F35").values = [["Đối tượng & nhu cầu"]];
competitor.getRange("I35").values = [["Nội dung / kênh / CTA"]];

const themes = [
  ["Ưu tiên 1", "quản lý nhiều tài khoản", "Team Ads, Seller — cần tổ chức nhiều account rõ ràng", "Bài problem → solution; SEO/Social; CTA: thử 3 tài khoản"],
  ["Ưu tiên 1", "Chrome profile vs antidetect", "Người đang vận hành thủ công, muốn hiểu khi nào cần đổi công cụ", "Bài so sánh; SEO/Social; CTA: kiểm tra workflow hiện tại"],
  ["Ưu tiên 1", "phân quyền tài khoản team ads", "Team Ads — cần giao việc, bàn giao và kiểm soát quyền", "Guide + demo; Facebook/Outbound; CTA: thử 3–5 tài khoản"],
  ["Ưu tiên 1", "quản lý nhiều shop", "Seller — cần tách môi trường và quản lý nhiều cửa hàng", "Checklist + demo; TikTok/Group; CTA: đưa 3 shop vào test"],
  ["Ưu tiên 1", "sync nhiều profile", "Social Ops — cần giảm thao tác lặp lại trên nhiều profile", "Product proof; TikTok/Facebook; CTA: test một batch nhỏ"],
  ["Ưu tiên 2", "antidetect browser / multi-account browser", "Người đang tìm hiểu loại công cụ hoặc giải pháp tổng quát", "Comparison + education; SEO/Search; cần nghiên cứu SERP trước"],
  ["Sau benchmark", "chuyển profile antidetect", "Người đang dùng công cụ khác, quan tâm migration và tính liên tục", "Migration guide + proof; SEO/Outbound; CTA: test song song"],
];

for (let index = 0; index < themes.length; index += 1) {
  const row = 36 + index;
  for (const range of [`A${row}:B${row}`, `C${row}:E${row}`, `F${row}:H${row}`, `I${row}:L${row}`]) {
    competitor.mergeCells(range);
  }
  competitor.getRange(`A${row}`).values = [[themes[index][0]]];
  competitor.getRange(`C${row}`).values = [[themes[index][1]]];
  competitor.getRange(`F${row}`).values = [[themes[index][2]]];
  competitor.getRange(`I${row}`).values = [[themes[index][3]]];
}

competitor.mergeCells("A44:L44");
competitor.getRange("A44").values = [["Ghi chú: Đây là định hướng nội dung phục vụ GTM, chưa phải bộ keyword SEO hoàn chỉnh. Chỉ tách thành sheet SEO riêng khi có search volume, keyword difficulty, thứ hạng hiện tại và SERP benchmark."]];

competitor.getRange("A34:L34").format = {
  fill: "#5B9BD5",
  font: { name: "Calibri", size: 11, bold: true, color: "#FFFFFF" },
  verticalAlignment: "middle",
  wrapText: true,
  rowHeightPx: 32,
};
competitor.getRange("A35:L35").format = {
  fill: "#203864",
  font: { name: "Calibri", size: 10, bold: true, color: "#FFFFFF" },
  horizontalAlignment: "center",
  verticalAlignment: "middle",
  wrapText: true,
  rowHeightPx: 44,
  borders: { bottom: { style: "thin", color: "#D9E2F3" } },
};
for (let row = 36; row <= 42; row += 1) {
  competitor.getRange(`A${row}:L${row}`).format = {
    fill: row % 2 === 0 ? "#FFFFFF" : "#EAF2F8",
    font: { name: "Calibri", size: 10, color: "#1F2937" },
    verticalAlignment: "middle",
    wrapText: true,
    rowHeightPx: 62,
    borders: { bottom: { style: "thin", color: "#D9E2F3" } },
  };
  competitor.getRange(`A${row}:B${row}`).format.font = { name: "Calibri", size: 10, bold: true, color: "#1F4E78" };
  competitor.getRange(`C${row}:E${row}`).format.font = { name: "Calibri", size: 10, bold: true, color: "#1F4E78" };
}
competitor.getRange("A44:L44").format = {
  fill: "#FFF2CC",
  font: { name: "Calibri", size: 10, italic: true, color: "#7F6000" },
  verticalAlignment: "middle",
  wrapText: true,
  rowHeightPx: 48,
};

workbook.worksheets.getItem("Phân tích từ khóa_Đối thủ cạnh tranh").delete();

const exported = await SpreadsheetFile.exportXlsx(workbook);
await exported.save(outputPath);
console.log(`SAVED ${outputPath}`);
