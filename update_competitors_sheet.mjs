import fs from "node:fs/promises";
import { SpreadsheetFile } from "@oai/artifact-tool";

const input = "D:/kehoach/InviBrowser_Business_Plan_Giu_Nguyen_Format_v2.xlsx";
const output = "D:/kehoach/InviBrowser_Business_Plan_Giu_Nguyen_Format_v2.editing.xlsx";

const workbook = await SpreadsheetFile.importXlsx(await fs.readFile(input));
const sheet = workbook.worksheets.getItemAt(7);

sheet.unmergeCells("A1:L30");
sheet.getRange("A1:L40").clear({ applyTo: "all" });
sheet.deleteAllDrawings();
sheet.showGridLines = false;

const mergeGroups = (row, groups) => {
  for (const [start, end] of groups) sheet.mergeCells(`${start}${row}:${end}${row}`);
};
sheet.mergeCells("A1:L1");
sheet.mergeCells("A2:L2");
sheet.mergeCells("A4:L4");
for (let row = 5; row <= 9; row++) mergeGroups(row, [["A","B"],["C","E"],["F","I"],["J","L"]]);
sheet.mergeCells("A11:L11");
for (let row = 12; row <= 21; row++) mergeGroups(row, [["A","B"],["C","E"],["F","H"],["I","L"]]);
sheet.mergeCells("A23:L23");
for (let row = 24; row <= 30; row++) mergeGroups(row, [["A","E"],["F","L"]]);
sheet.mergeCells("A32:L32");

sheet.getRange("A1").values = [["BẢN ĐỒ ĐỐI THỦ INViBROWSER"]];
sheet.getRange("A2").values = [["Phân tầng mối đe dọa, lợi thế cạnh tranh và hàm ý cho chiến lược thâm nhập thị trường"]];
sheet.getRange("A4").values = [["1. PHÂN TẦNG ĐỐI THỦ"]];
sheet.getRange("A5:J9").values = [
  ["Nhóm", "", "Đối thủ", "", "", "Vì sao đưa vào", "", "", "", "Vai trò khi phân tích"],
  ["Tier 1 - threat trực tiếp", "", "GPM Login, AdsPower, GenLogin, Hidemyacc", "", "", "Overlap cao với ICP/use case; có lợi thế rõ về distribution, team ops, automation hoặc portability.", "", "", "", "Hiểu trận cạnh tranh chính và rào cản chuyển đổi."],
  ["Tier 2 - threat theo năng lực", "", "GemLogin, Hidemium, GoLogin, OmniLogin", "", "", "Mỗi bên mạnh ở một trục: ecosystem, API/technical, cloud/cross-platform hoặc integrated stack.", "", "", "", "Soi capability gap và xu hướng sản phẩm."],
  ["Tier 3 - product benchmark", "", "KikiLogin", "", "", "Feature breadth khá rộng; cần giữ trong benchmark dù GTM tại Việt Nam chưa cho thấy threat bằng Tier 1.", "", "", "", "Tránh bỏ sót đối thủ có năng lực sản phẩm."],
  ["Substitute", "", "Chrome Profiles, browser thường, nhiều máy, sheet/manual, script tự làm", "", "", "Cùng giải bài toán bằng cách khác và là trạng thái hiện tại của nhiều user.", "", "", "", "Đối thủ thật của beachhead manual → structured system."],
];

sheet.getRange("A11").values = [["2. ĐỐI THỦ CHÍNH"]];
sheet.getRange("A12:I21").values = [
  ["Đối thủ", "", "Họ mạnh ở đâu", "", "", "Vì sao đây là lợi thế", "", "", "InviBrowser cần lưu ý"],
  ["GPM Login", "", "Volume profile lớn, lifetime, local distribution", "", "", "User volume lớn dễ kiểm soát chi phí; thương hiệu quen trong cộng đồng Việt Nam; seeding mạnh.", "", "", "Không nên đánh trực diện bằng 'nhiều profile giá rẻ'."],
  ["AdsPower", "", "Permission, sharing, RPA, API, migration", "", "", "Team đã đưa workflow vào AdsPower sẽ có rào cản chuyển tool cao.", "", "", "Không lấy agency lớn đang dùng sâu AdsPower làm beachhead."],
  ["GenLogin", "", "Automation/no-code, Sync và mini-app", "", "", "Hấp dẫn nhóm seeding/affiliate không biết code nhưng muốn tự động hóa.", "", "", "Bulk + Sync + Automation không thể là khác biệt duy nhất."],
  ["Hidemyacc", "", "Share, transfer, backup, session sync và automation", "", "", "Giảm rủi ro bàn giao, đổi máy và chuyển hệ thống.", "", "", "Nếu nói portability/migration phải benchmark thật."],
  ["GemLogin", "", "GemStore và hệ sinh thái automation", "", "", "User có thể dùng app/script sẵn; hệ sinh thái giúp tăng lock-in.", "", "", "InviBrowser hiện chưa có marketplace tương tự."],
  ["Hidemium", "", "Cấu hình sâu, API, automation cho user nâng cao", "", "", "Phù hợp operator kỹ thuật và workflow phức tạp.", "", "", "Không nên lao vào trận advanced technical ngay."],
  ["GoLogin", "", "Cloud, web và cross-platform", "", "", "User không phụ thuộc một máy Windows; phù hợp remote/team.", "", "", "Windows-only hiện là gap cần ghi nhận."],
  ["OmniLogin", "", "Browser + proxy + automation trong một hệ", "", "", "Giảm nhu cầu ghép nhiều công cụ rời rạc.", "", "", "Theo dõi vì overlap sản phẩm khá cao."],
  ["KikiLogin", "", "Feature breadth khá rộng", "", "", "Có khả năng cạnh tranh về product capability.", "", "", "GTM tại Việt Nam hiện chưa cho thấy threat bằng Tier 1."],
];

sheet.getRange("A23").values = [["3. HÀM Ý CHIẾN LƯỢC"]];
sheet.getRange("A24:F30").values = [
  ["Key finding", "", "", "", "", "Implication cho InviBrowser"],
  ["Feature parity ở core operations khá cao", "", "", "", "", "Không định vị bằng một feature phổ biến. Bán workflow/use case hoàn chỉnh cho một nhóm cụ thể."],
  ["Lợi thế cạnh tranh không chỉ nằm ở product feature", "", "", "", "", "Xem distribution, ecosystem và platform coverage là một phần của competitive strategy."],
  ["Switching cost tăng khi workflow đã ăn sâu vào tool", "", "", "", "", "Ưu tiên user manual/Chrome; với user mature, bắt đầu bằng batch nhỏ hoặc use case phụ."],
  ["Price/profile-volume là trận khó để dẫn đầu", "", "", "", "", "Dùng Free 10 profiles làm cơ chế trial/activation, không dùng làm trục định vị."],
  ["Ecosystem và platform breadth tạo lock-in", "", "", "", "", "Theo dõi gap về marketplace, cloud/cross-platform và integrated stack trong roadmap dài hạn."],
  ["Một số claim quan trọng chưa được chứng minh", "", "", "", "", "Chỉ claim performance, stability, fingerprint và ease of use sau khi có benchmark."],
];
sheet.getRange("A32").values = [["Kết luận: Giữ định vị manual/Chrome → multi-account có hệ thống; không đấu bằng số profile hoặc claim kỹ thuật chưa chứng minh; dùng workflow proof để thắng nhóm có switching cost thấp trước."]];

const NAVY = "#1F4E78";
const BLUE = "#5B9BD5";
const HEADER = "#203864";
const PALE_BLUE = "#EAF2F8";
const PALE_YELLOW = "#FFF2CC";
const BORDER = "#D9E2F3";
const TEXT = "#1F2937";

sheet.getRange("A1:L32").format.font = { name: "Calibri", size: 10, color: TEXT };
sheet.getRange("A1:L32").format.verticalAlignment = "center";
sheet.getRange("A1:L32").format.wrapText = true;
sheet.getRange("A1:L32").format.borders = { preset: "all", style: "thin", color: BORDER };

sheet.getRange("A1:L1").format.fill = NAVY;
sheet.getRange("A1:L1").format.font = { name: "Calibri", size: 18, bold: true, color: "#FFFFFF" };
sheet.getRange("A1:L1").format.horizontalAlignment = "center";
sheet.getRange("A2:L2").format.fill = "#D9EAF7";
sheet.getRange("A2:L2").format.font = { name: "Calibri", size: 10, italic: true, color: "#35546F" };
sheet.getRange("A2:L2").format.horizontalAlignment = "center";

for (const row of [4, 11, 23]) {
  sheet.getRange(`A${row}:L${row}`).format.fill = BLUE;
  sheet.getRange(`A${row}:L${row}`).format.font = { name: "Calibri", size: 11, bold: true, color: "#FFFFFF" };
}
for (const row of [5, 12, 24]) {
  sheet.getRange(`A${row}:L${row}`).format.fill = HEADER;
  sheet.getRange(`A${row}:L${row}`).format.font = { name: "Calibri", size: 10, bold: true, color: "#FFFFFF" };
  sheet.getRange(`A${row}:L${row}`).format.horizontalAlignment = "center";
}
for (const row of [7, 9, 14, 16, 18, 20, 26, 28, 30]) {
  sheet.getRange(`A${row}:L${row}`).format.fill = PALE_BLUE;
}
sheet.getRange("A32:L32").format.fill = PALE_YELLOW;
sheet.getRange("A32:L32").format.font = { name: "Calibri", size: 10, bold: true, color: "#7F6000" };

for (const row of [6, 7, 8, 9, 13, 14, 15, 16, 17, 18, 19, 20, 21]) {
  sheet.getRange(`A${row}:B${row}`).format.font = { name: "Calibri", size: 10, bold: true, color: NAVY };
}

const widths = { A:10, B:10, C:11, D:11, E:11, F:11, G:11, H:11, I:11, J:11, K:11, L:11 };
for (const [column, width] of Object.entries(widths)) sheet.getRange(`${column}:${column}`).format.columnWidth = width;
const heights = {1:32,2:30,3:8,4:24,5:34,6:58,7:58,8:64,9:64,10:8,11:24,12:34,13:62,14:62,15:62,16:62,17:62,18:62,19:62,20:62,21:62,22:8,23:24,24:34,25:54,26:54,27:58,28:54,29:58,30:58,31:8,32:56};
for (const [row, height] of Object.entries(heights)) sheet.getRange(`${row}:${row}`).format.rowHeight = height;

const exported = await SpreadsheetFile.exportXlsx(workbook);
await exported.save(output);
console.log(`SAVED ${output}`);
