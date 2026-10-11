import fs from "node:fs/promises";
import { SpreadsheetFile } from "@oai/artifact-tool";

const currentPath = "D:/kehoach/InviBrowser_Business_Plan_Giu_Nguyen_Format_v2.xlsx";
const templatePath = "C:/Users/Admin/Desktop/Bussiness Plan mẫu.xlsx";
const outputPath = "D:/kehoach/InviBrowser_Business_Plan_Giu_Nguyen_Format_v2.finance_editing.xlsx";

const [workbook, template] = await Promise.all([
  SpreadsheetFile.importXlsx(await fs.readFile(currentPath)),
  SpreadsheetFile.importXlsx(await fs.readFile(templatePath)),
]);

const plan = workbook.worksheets.getItemAt(1);
const reference = workbook.worksheets.getItemAt(2);
const templateReference = template.worksheets.getItemAt(2);

for (const sheet of [plan, reference]) {
  const used = sheet.getUsedRange();
  if (used) used.clear({ applyTo: "contents" });
}

function put(sheet, values) {
  for (const [address, value] of Object.entries(values)) {
    sheet.getRange(address).values = [[value]];
  }
}

put(plan, {
  A1: "KẾ HOẠCH TÀI CHÍNH INViBROWSER 12 THÁNG",
  E1: "0–30 ngày: Acquire + Activate + Learn",
  N1: "31–60 ngày: Focus + Improve",
  Z1: "61–90 ngày: Scale What Works",
  E2: "Tạo baseline; test 4 P1; onboarding tạo profile đầu tiên",
  N2: "Giữ 2–3 hướng mạnh; tối ưu message, onboarding và blocker",
  Z2: "Scale kênh/message đã chứng minh; case study; trigger nâng cấp",
  B3: "Hạng mục", C3: "Nội dung", D3: "Chú thích",
  E3: "Tháng 1", H3: "Tháng 2", K3: "Tháng 3", N3: "Tháng 4",
  Q3: "Tháng 5", T3: "Tháng 6", W3: "Tháng 7", Z3: "Tháng 8",
  AC3: "Tháng 9", AF3: "Tháng 10", AI3: "Tháng 11", AL3: "Tháng 12",
  B4: "Căn cứ lập kế hoạch: 45 signup, 31 activated (68,9%), 7 user có tín hiệu D7. Tuần 05–11/10 có 7 signup và 4 activated (57,1%).",
  B5: "Mục tiêu: cải thiện chất lượng acquisition, activation và product learning; sau đó tập trung và scale hướng đã chứng minh hiệu quả.",
  B6: "Sản phẩm", C6: "InviBrowser", D6: "SaaS freemium; Free 10 profiles là cơ chế trial/activation.",

  A8: 1, B8: "Funnel và doanh thu",
  A9: 1.1, B9: "Website traffic", D9: "Chưa có dữ liệu nguồn; cần theo dõi theo channel/campaign.",
  A10: 1.2, B10: "Signup", D10: "Baseline hiện có: 45 signup; cần bổ sung source và ICP.",
  A11: 1.3, B11: "Activated users", D11: "31 user, tương ứng 68,9%; activation = tạo profile đầu tiên.",
  A12: 1.4, B12: "D7 retained users", D12: "Hiện có 7 user có tín hiệu D7; cần chuẩn hóa cohort.",
  B13: "Weekly Activated Users", D13: "North Star v0.1: user tạo profile đầu tiên trong tuần.",
  B14: "Feature usage có ý nghĩa", D14: "Theo dõi permission/handoff, organization, bulk, sync và workflow.",
  B15: "Pricing page views", D15: "Chưa có dữ liệu trong nguồn.",
  B16: "Price inquiries", D16: "Chưa có dữ liệu trong nguồn.",
  B17: "Free-limit hits", D17: "Theo dõi user chạm giới hạn 10 profiles.",
  B18: "Team requests", D18: "Theo dõi nhu cầu permission, sharing và handoff.",
  B19: "Paid users và ARPU", D19: "Chưa có pricing, paid users hoặc ARPU trong nguồn.",
  B20: "Doanh thu", D20: "Để trống đến khi có dữ liệu billing/paid thực tế.",

  A21: 2, B21: "Chi phí GTM và vận hành trực tiếp",
  A22: 2.1, B22: "Paid media", D22: "Facebook/TikTok/Search; ngân sách chưa có trong nguồn.",
  B23: "Ngân sách và CAC", D23: "Cần bổ sung spend, paid users và attribution.",
  A24: 2.2, B24: "Content production", D24: "20 nội dung: 5 đã đăng, 4 đang sản xuất, 11 đã lên kế hoạch.",
  B25: "Chi phí trên mỗi nội dung", D25: "Chưa có dữ liệu chi phí.",
  A26: 2.3, B26: "Community seeding", D26: "85 lượt: 27 đã đăng, 35 chờ duyệt, 12 bị xóa, 11 bị từ chối.",
  B27: "Chi phí và hiệu quả seeding", D27: "Đo hội thoại, click, signup biết nguồn và activation; chưa có cost.",
  A28: 2.4, B28: "Cold outreach", D28: "Lập source list đúng ICP; CTA thử 3–5 account hoặc một batch nhỏ.",
  B29: "Email/outreach tools", D29: "Chưa có dữ liệu chi phí.",
  A30: 2.5, B30: "SEO và Google Search", D30: "Bắt nhu cầu category, review, comparison và problem-aware.",
  B31: "Website và landing page", D31: "Chuyển traffic thành download/signup; cần tracking theo source.",
  A32: 2.6, B32: "Onboarding và retention", D32: "In-app, email, tutorial và docs để tăng first value và D7.",
  B33: "Chi phí lifecycle", D33: "Chưa có dữ liệu chi phí.",
  A34: 2.7, B34: "Analytics và tracking", D34: "Theo dõi source, ICP, signup, activation, D7, feature use và paid signal.",
  B35: "Công cụ đo lường", D35: "Chưa có danh sách công cụ và chi phí.",
  A36: 2.8, B36: "Customer interview", D36: "Hỏi user về pain, blocker, use case, churn reason và feature request.",
  B37: "Incentive nghiên cứu", D37: "Chưa có dữ liệu chi phí.",
  A38: 2.9, B38: "Competitive benchmark", D38: "Đo performance, stability, fingerprint và ease of use trước khi claim.",
  B39: "Chi phí benchmark", D39: "Chưa có dữ liệu chi phí.",
  A40: 2.1, B40: "Case study và retarget", D40: "Triển khai khi pain/proof và activation đã có tín hiệu rõ.",
  B41: "Chi phí retarget", D41: "Chưa có dữ liệu chi phí.",
  A42: 2.11, B42: "Product fixes theo feedback", D42: "Ưu tiên blocker lặp lại; thông báo lại và đo user quay lại.",
  B43: "Nguồn lực sản phẩm", D43: "Chưa có chi phí nhân sự/engineering.",
  A44: 2.12, B44: "Thử nghiệm khác", D44: "Chỉ mở rộng khi có quyết định KEEP/SCALE/IMPROVE/PAUSE.",
  B45: "Ngân sách dự phòng", D45: "Chưa có dữ liệu.",

  A46: 3, B46: "Đóng góp sau chi phí trực tiếp", D46: "Chưa tính được do thiếu revenue và cost.",
  A47: 4, B47: "Chi phí bán hàng và quản lý doanh nghiệp",
  A48: 4.1, B48: "Marketing và GTM", D48: "Marketing/Quản lý và Marketing + Intern; payroll chưa có.",
  B49: "Ngân sách marketing", D49: "Chưa có dữ liệu.",
  A50: 4.2, B50: "Product và infrastructure", D50: "Hạ tầng sản phẩm, tooling, analytics và vận hành kỹ thuật.",
  B51: "Chi phí product/infra", D51: "Chưa có dữ liệu.",
  A52: 4.3, B52: "Sales và Customer Success", D52: "Follow-up signup, qualify paid signal và đóng vòng feedback.",
  B53: "Chi phí Sales/CS", D53: "Chưa có dữ liệu.",
  A54: 4.4, B54: "Software và vận hành", D54: "CRM, email, content, research và công cụ cộng tác.",
  B55: "Chi phí software", D55: "Chưa có dữ liệu.",
  A56: 4.5, B56: "Chi phí dự phòng", D56: "Chưa có tỷ lệ hoặc ngân sách được phê duyệt.",
  B57: "Nguyên tắc", D57: "Không tự đặt giả định khi chưa có owner và nguồn dữ liệu.",
  A58: 5, B58: "Lợi nhuận trước thuế", D58: "Chưa tính được.",
  B59: "Thuế và nghĩa vụ tài chính", D59: "Chưa đủ dữ liệu và chính sách áp dụng.",
  A60: 6, B60: "Lợi nhuận sau thuế", D60: "Chưa tính được.",
});

const triplets = [
  ["E","F","G"],["H","I","J"],["K","L","M"],["N","O","P"],
  ["Q","R","S"],["T","U","V"],["W","X","Y"],["Z","AA","AB"],
  ["AC","AD","AE"],["AF","AG","AH"],["AI","AJ","AK"],["AL","AM","AN"],
];
for (const [a,b,c] of triplets) {
  plan.getRange(`${a}4`).values = [["Acquisition"]];
  plan.getRange(`${b}4`).values = [["Activation"]];
  plan.getRange(`${c}4`).values = [["Retention / Learning"]];
}

// Restore only the first four columns from the original reference sheet.
const sourceRange = templateReference.getRange("A1:D60");
const sourceValues = sourceRange.values;
const referenceColumns = sourceValues.map((row) => row.map((value) =>
  typeof value === "string" && value.startsWith("=") ? "" : value
));
reference.getRange("A1:D60").values = referenceColumns;

const exported = await SpreadsheetFile.exportXlsx(workbook);
await exported.save(outputPath);
console.log(`SAVED ${outputPath}`);
