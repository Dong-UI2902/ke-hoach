import fs from "node:fs/promises";
import path from "node:path";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const outputPath = "D:/kehoach/InviBrowser_Business_Plan_2026_2027.xlsx";
const previewDir = "D:/kehoach/qa_invibrowser_plan";
await fs.mkdir(previewDir, { recursive: true });

const wb = Workbook.create();
const NAVY = "#1F2A44";
const NAVY2 = "#314563";
const YELLOW = "#FFF2CC";
const YELLOW_STRONG = "#FFF258";
const PALE_BLUE = "#D9EAF7";
const PALE_GREEN = "#E2F0D9";
const PALE_RED = "#FCE4D6";
const GRAY = "#F2F2F2";
const BORDER = "#D9D9D9";
const TEXT = "#1F2329";
const FONT = "Arial";

function baseSheet(name) {
  const ws = wb.worksheets.add(name);
  ws.showGridLines = false;
  return ws;
}

function setWidth(ws, col, width) {
  ws.getRange(`${col}:${col}`).format.columnWidth = width;
}

function title(ws, range, text) {
  ws.mergeCells(range);
  const r = ws.getRange(range);
  r.values = [[text]];
  r.format.fill = YELLOW;
  r.format.font = { name: FONT, size: 15, bold: true, color: "#7F6000" };
  r.format.horizontalAlignment = "center";
  r.format.verticalAlignment = "center";
  r.format.borders = { preset: "outside", style: "thin", color: BORDER };
}

function section(ws, row, text, endCol = "H") {
  const range = `B${row}:${endCol}${row}`;
  ws.mergeCells(range);
  const r = ws.getRange(range);
  r.values = [[text]];
  r.format.fill = NAVY;
  r.format.font = { name: FONT, size: 11, bold: true, color: "#FFFFFF" };
  r.format.borders = { preset: "outside", style: "thin", color: NAVY };
  r.format.rowHeight = 24;
}

function narrative(ws, row, label, text, endCol = "H", height = 48) {
  ws.getRange(`B${row}`).values = [[label]];
  ws.getRange(`B${row}`).format = {
    fill: GRAY,
    font: { name: FONT, size: 10, bold: true, color: TEXT },
    wrapText: true,
    verticalAlignment: "top",
    borders: { preset: "all", style: "thin", color: BORDER },
  };
  ws.mergeCells(`C${row}:${endCol}${row}`);
  const c = ws.getRange(`C${row}:${endCol}${row}`);
  c.values = [[text]];
  c.format = {
    font: { name: FONT, size: 10, color: TEXT },
    wrapText: true,
    verticalAlignment: "top",
    borders: { preset: "all", style: "thin", color: BORDER },
  };
  ws.getRange(`B${row}:${endCol}${row}`).format.rowHeight = height;
}

function tableHeader(ws, range) {
  const r = ws.getRange(range);
  r.format.fill = NAVY;
  r.format.font = { name: FONT, size: 10, bold: true, color: "#FFFFFF" };
  r.format.horizontalAlignment = "center";
  r.format.verticalAlignment = "center";
  r.format.wrapText = true;
  r.format.borders = { preset: "all", style: "thin", color: "#FFFFFF" };
  r.format.rowHeight = 30;
}

function bodyTable(ws, range) {
  const r = ws.getRange(range);
  r.format.font = { name: FONT, size: 10, color: TEXT };
  r.format.borders = { preset: "all", style: "thin", color: BORDER };
  r.format.wrapText = true;
  r.format.verticalAlignment = "top";
}

// 1. Business plan following the reference workbook framework.
const plan = baseSheet("Kế hoạch kinh doanh");
setWidth(plan, "A", 3); setWidth(plan, "B", 28);
for (const c of ["C","D","E","F","G","H"]) setWidth(plan, c, 15);
for (const c of ["I","J","K","L","M","N"]) setWidth(plan, c, 11);
title(plan, "B1:H1", "KẾ HOẠCH KINH DOANH VÀ THÂM NHẬP THỊ TRƯỜNG INViBROWSER TRONG 12 THÁNG");
plan.getRange("B2:C4").values = [
  ["Thời gian thực hiện", "12 tháng từ 10/2026 đến 09/2027"],
  ["Mô hình kinh doanh", "SaaS freemium cho cá nhân và đội nhóm vận hành nhiều tài khoản"],
  ["Chiến lược kinh doanh", "Acquire → Activate → Learn → Focus → Scale"],
];
plan.getRange("B2:B4").format.fill = GRAY;
plan.getRange("B2:B4").format.font = { name: FONT, size: 10, bold: true, color: TEXT };
plan.getRange("B2:C4").format.borders = { preset: "all", style: "thin", color: BORDER };
plan.getRange("C2:C4").format.wrapText = true;

section(plan, 6, "I. EXECUTIVE SUMMARY");
narrative(plan, 7, "1. Sản phẩm", "InviBrowser là trình duyệt multi-account/antidetect hỗ trợ tách môi trường, tổ chức profile, phân quyền đội nhóm và mở rộng vận hành bằng bulk, đồng bộ và workflow.", "H", 42);
narrative(plan, 8, "2. Định vị", "Giúp cá nhân và team đang chuyển từ Chrome/profile thủ công sang một hệ thống multi-account có tổ chức, tách biệt và có thể mở rộng. Điểm thâm nhập là người đã có pain nhiều tài khoản nhưng chưa bị khóa sâu vào công cụ đối thủ.", "H", 58);
narrative(plan, 9, "3. Mục tiêu", "Giai đoạn đầu ưu tiên Acquisition, Activation và Product Learning. Paid là tín hiệu thương mại quan trọng nhưng chưa phải KPI chính. North Star v0.1 là Weekly Activated Users: người dùng trong tuần đã tạo profile đầu tiên và nhận được giá trị ban đầu.", "H", 58);
narrative(plan, 10, "4. Nhóm ưu tiên", "Team Ads nhỏ-vừa; Seller nhiều shop/account; Seeding/Social Ops; Affiliate đang scale. Đây là bốn nhóm có pain rõ, phù hợp với sản phẩm hiện tại và rào cản chuyển đổi chưa quá cao.", "H", 48);

section(plan, 12, "II. MARKET OPPORTUNITY");
narrative(plan, 13, "Phạm vi thị trường", "Category chính: Antidetect / Multi-account Browser. Bài toán rộng hơn là vận hành nhiều tài khoản, profile, quyền truy cập và workflow.", "H", 42);
narrative(plan, 14, "Ba loại nhu cầu", "1) Đã biết loại công cụ: tìm pricing, review, comparison. 2) Đã biết vấn đề: nhiều account/shop/team nhưng chưa biết antidetect. 3) Muốn chuyển tool: cần migration, test song song và lý do đổi rõ ràng.", "H", 52);
narrative(plan, 15, "Cửa thâm nhập", "Tập trung nhóm đang dùng Chrome profiles, nhiều browser, nhiều máy hoặc quy trình sheet/manual. Đây là đối thủ thực tế gần nhất và là nhóm có switching cost thấp hơn các power user đã gắn chặt với hệ sinh thái đối thủ.", "H", 52);
narrative(plan, 16, "Cụm pain ưu tiên", "Tổ chức nhiều account và tách môi trường; vận hành số lượng lớn; phân quyền và bàn giao; rủi ro liên đới; thiết lập hàng loạt. Chuyển dữ liệu, độ ổn định và hiệu năng cần tiếp tục benchmark trước khi dùng làm claim.", "H", 56);
narrative(plan, 17, "Cạnh tranh", "Tier 1: GPM Login, AdsPower, GenLogin, Hidemyacc. Tier 2: GemLogin, Hidemium, GoLogin, OmniLogin. KikiLogin dùng làm product benchmark. Tránh đấu bằng số profile, giá rẻ hoặc claim kỹ thuật chưa có bằng chứng.", "H", 56);

section(plan, 19, "III. CHIẾN LƯỢC GTM ĐỀ XUẤT");
narrative(plan, 20, "Chiến lược tổng thể", "Pain first, proof close to conversion. Dùng Free 10 profile để giảm rào cản thử và tạo nguồn học từ người dùng. Chỉ mở rộng nguồn lực khi có tín hiệu activation, D7, feedback hoặc commercial signal.", "H", 52);
narrative(plan, 21, "0–30 ngày", "Acquire + Activate + Learn: hoàn thiện tracking, kiểm tra bốn nhóm P1, chạy content/community/outbound, phỏng vấn người dùng và sửa blocker ảnh hưởng tạo profile.", "H", 48);
narrative(plan, 22, "31–60 ngày", "Focus + Improve: giữ 2–3 hướng mạnh, tối ưu thông điệp và onboarding theo use case, cải thiện sản phẩm dựa trên pain lặp lại và thiết lập benchmark activation.", "H", 48);
narrative(plan, 23, "61–90 ngày", "Scale What Works: tăng nguồn lực cho kênh và thông điệp đã chứng minh hiệu quả, tạo case study, retarget và kiểm tra trigger nâng cấp.", "H", 48);
narrative(plan, 24, "Nguyên tắc quyết định", "KEEP khi có tín hiệu nhưng cần thêm dữ liệu; SCALE khi activation/retention/feedback tốt; IMPROVE khi có tiềm năng nhưng funnel có bottleneck; PAUSE khi đã kiểm tra đủ mà tín hiệu vẫn yếu.", "H", 52);

section(plan, 26, "IV. CHIẾN LƯỢC KHÁCH HÀNG");
narrative(plan, 27, "Team Ads nhỏ-vừa", "Pain: nhiều account, nhiều người và share thủ công. Lời hứa: kiểm soát account và quyền truy cập gọn hơn. Proof: team, permission, isolation, Sync. CTA: thử 3–5 account.", "H", 52);
narrative(plan, 28, "Seller nhiều shop/account", "Pain: nhiều shop/account, dễ lẫn môi trường và session. Lời hứa: mỗi shop/account một môi trường, quản lý rõ ràng. Proof: isolation, group, session, team. CTA: đưa 3 shop vào test.", "H", 52);
narrative(plan, 29, "Seeding / Social Ops", "Pain: volume lớn và thao tác lặp. Lời hứa: scale dàn account mà không tăng thao tác tay tương ứng. Proof: Bulk, Sync, Workflow. CTA: thử một batch nhỏ.", "H", 52);
narrative(plan, 30, "Affiliate đang scale", "Pain: từ vài account lên nhiều campaign/account. Lời hứa: chuyển từ manual sang workflow multi-account có hệ thống. Proof: Group, Bulk, Automation. CTA: thử batch đầu tiên.", "H", 52);

section(plan, 32, "V. CHIẾN LƯỢC SẢN PHẨM VÀ GIÁ TRỊ");
narrative(plan, 33, "Giá trị cốt lõi", "Tách môi trường; tổ chức profile; vận hành nhiều account; phân quyền/bàn giao; scale bằng bulk, đồng bộ và workflow.", "H", 42);
narrative(plan, 34, "Activation", "Mốc activation dùng chung: người dùng tạo profile đầu tiên. Theo dõi tiếp D7, feature use lặp lại và blocker thay vì chỉ đếm đăng ký.", "H", 42);
narrative(plan, 35, "Free plan", "Free 10 profile là cơ chế trial và activation. Không dùng số profile miễn phí làm trục định vị chính hoặc cam kết lợi thế giá dài hạn.", "H", 42);
narrative(plan, 36, "Product loop", "Kéo user → kích hoạt → quan sát → hỏi → phân loại feedback → sửa/build → thông báo lại → đo quay lại. Bug khiến profile không mở và pain lặp ở nhóm P1 có độ ưu tiên cao.", "H", 52);
narrative(plan, 37, "Claim cần kiểm chứng", "Hiệu năng, độ ổn định, chất lượng fingerprint, mức dễ dùng, migration/portability và độ sâu automation. Chỉ dùng làm thông điệp cạnh tranh sau khi có benchmark hoặc proof thực tế.", "H", 48);

section(plan, 39, "VI. CHIẾN LƯỢC KÊNH");
narrative(plan, 40, "Tạo nhu cầu", "TikTok, Facebook và cộng đồng: cho người dùng thấy pain tăng theo số account, team và thao tác lặp.", "H", 42);
narrative(plan, 41, "Đón nhu cầu", "SEO, Google Search, group/forum: guide, comparison và nội dung category cho người đã tìm giải pháp.", "H", 42);
narrative(plan, 42, "Chủ động tiếp cận", "Cold email và direct outreach theo từng segment. Bắt đầu bằng use case nhỏ, đúng pain, có CTA thử batch/account cụ thể.", "H", 42);
narrative(plan, 43, "Chuyển đổi", "Website, landing, pricing, use-case page và demo. Càng gần conversion càng cần proof rõ hơn.", "H", 42);
narrative(plan, 44, "Kích hoạt và giữ user", "In-app, email, tutorial và docs để đưa người dùng từ signup đến tạo profile, dùng feature có ý nghĩa và quay lại.", "H", 42);

section(plan, 46, "VII. CHIẾN LƯỢC CONTENT");
narrative(plan, 47, "Problem", "Cho thấy khi số account tăng, Chrome/profile thủ công bắt đầu vỡ: lẫn session, khó tìm, khó bàn giao, nhiều thao tác và thiếu kiểm soát.", "H", 48);
narrative(plan, 48, "Education", "Giải thích khi nào browser profile thông thường còn đủ và khi nào cần hệ thống multi-account; hướng dẫn tổ chức theo shop, campaign, team và batch.", "H", 48);
narrative(plan, 49, "Product proof", "Demo bulk creation, Sync Control, team permission, group/profile structure và end-to-end workflow trên batch nhỏ.", "H", 44);
narrative(plan, 50, "Comparison", "So sánh workflow before/after và hướng chuyển từ manual. Không claim hơn đối thủ khi chưa có benchmark.", "H", 42);
narrative(plan, 51, "Retention content", "Quick start, tutorial tạo profile đầu tiên, follow-up D7 và nội dung dùng sâu hơn theo use case.", "H", 42);

section(plan, 53, "VIII. CHIẾN LƯỢC COMMUNITY VÀ OUTBOUND");
narrative(plan, 54, "Community", "Xây kho nguồn theo Facebook Ads, TikTok Ads, Seller, Affiliate, MMO và Social Ops. Ghi từng lượt seeding theo nguồn ID, góc tiếp cận, trạng thái, trả lời, hội thoại chất lượng và đăng ký biết nguồn.", "H", 56);
narrative(plan, 55, "Góc tiếp cận", "Ưu tiên nêu vấn đề và chia sẻ cách làm trước khi giới thiệu sản phẩm trực tiếp. Theo dõi tỷ lệ bài được đăng, phản hồi chất lượng và đăng ký thay vì chỉ đếm lượt đăng.", "H", 48);
narrative(plan, 56, "Outbound", "Lập danh sách team nhỏ và operator có pain rõ; cá nhân hóa theo số account, cách bàn giao và workflow hiện tại; mời thử 3–5 account hoặc một batch nhỏ.", "H", 48);

section(plan, 58, "IX. CHIẾN LƯỢC SALES VÀ CUSTOMER LEARNING");
narrative(plan, 59, "Định nghĩa Lead", "Contact chỉ được tính là Lead khi có use case hoặc vấn đề thật liên quan InviBrowser. Signup dùng thử không tự động trở thành Lead.", "H", 44);
narrative(plan, 60, "Qualification", "Qualified khi đúng nhóm khách hàng, pain rõ và có khả năng sử dụng thực tế. Opportunity/Paid chỉ ghi khi có tín hiệu thương mại thật.", "H", 44);
narrative(plan, 61, "Theo dõi", "Đăng ký → liên hệ → tạo profile → D7 → tín hiệu thương mại → Paid. Paid và doanh thu tính theo ngày trả phí; Marketing attribution tách khỏi tổng user thực tế.", "H", 48);
narrative(plan, 62, "Feedback", "Ghi nguyên văn lý do chưa dùng, bug, feature request, objection, lost reason và purchase intent. Ưu tiên pain lặp ở nhóm P1 và vấn đề ảnh hưởng activation/retention.", "H", 48);

section(plan, 64, "X. MASTER TIMELINE 12 THÁNG", "N");
plan.getRange("B65:N65").values = [["Hoạt động","T1","T2","T3","T4","T5","T6","T7","T8","T9","T10","T11","T12"]];
tableHeader(plan, "B65:N65");
const timeline = [
  ["Tracking, baseline và chuẩn hóa dữ liệu",1,1,1,0,0,0,0,0,0,0,0,0],
  ["Kiểm tra 4 nhóm khách hàng P1",1,1,1,1,0,0,0,0,0,0,0,0],
  ["Content pain và education",1,1,1,1,1,1,1,1,1,1,1,1],
  ["Demo và product proof",1,1,1,1,1,1,1,1,1,1,1,1],
  ["Community seeding",1,1,1,1,1,1,1,1,1,1,1,1],
  ["Cold outreach theo ICP",1,1,1,1,1,1,1,1,1,1,1,1],
  ["Onboarding theo use case",1,1,1,1,1,1,1,1,1,1,1,1],
  ["Customer interview và product feedback",1,1,1,1,1,1,1,1,1,1,1,1],
  ["Benchmark đối thủ và claim",1,1,1,1,1,1,0,0,0,0,0,0],
  ["Chọn 1–2 ICP nổi bật",0,0,1,1,1,1,0,0,0,0,0,0],
  ["Case study và retarget",0,0,1,1,1,1,1,1,1,1,1,1],
  ["Kiểm tra trigger nâng cấp",0,0,1,1,1,1,1,1,1,1,1,1],
  ["Scale kênh và message hiệu quả",0,0,0,1,1,1,1,1,1,1,1,1],
];
plan.getRange(`B66:N${65+timeline.length}`).values = timeline;
bodyTable(plan, `B66:N${65+timeline.length}`);
plan.getRange(`C66:N${65+timeline.length}`).format.horizontalAlignment = "center";
plan.getRange(`C66:N${65+timeline.length}`).conditionalFormats.addCustom("=C66=1", { fill: PALE_GREEN, font: { color: "#375623", bold: true } });
plan.getRange(`C66:N${65+timeline.length}`).conditionalFormats.addCustom("=C66=0", { font: { color: "#B7B7B7" } });

section(plan, 81, "XI. KẾ HOẠCH NHÂN SỰ", "H");
narrative(plan, 82, "GTM / Marketing Lead", "Chốt ưu tiên ICP, kế hoạch test, message, kênh, review tuần và quyết định KEEP/SCALE/IMPROVE/PAUSE.", "H", 42);
narrative(plan, 83, "Content / Intern", "Sản xuất và đăng nội dung theo lịch, ghi kết quả theo mã nội dung, tái sử dụng nội dung giữa Facebook, TikTok và website.", "H", 42);
narrative(plan, 84, "Community team", "Seeding theo nguồn ID, tuân thủ quy định nhóm, ghi trạng thái và tín hiệu chất lượng. Hiện dữ liệu có ba người phụ trách: Quỳnh, Thảo và Đông.", "H", 44);
narrative(plan, 85, "Sales / CS", "Liên hệ đúng hạn, qualify lead, ghi lý do chưa dùng, tín hiệu thương mại, ngày trả phí và doanh thu.", "H", 42);
narrative(plan, 86, "Product", "Theo dõi activation, D7, blocker và feedback lặp; ưu tiên sửa lỗi ảnh hưởng tạo/mở profile và các pain được nhiều user P1 xác nhận.", "H", 42);
plan.freezePanes.freezeRows(1);

// 2. 90-day plan.
const exec = baseSheet("Kế hoạch 90 ngày");
setWidth(exec,"A",9); setWidth(exec,"B",15); setWidth(exec,"C",24); setWidth(exec,"D",22);
setWidth(exec,"E",30); setWidth(exec,"F",38); setWidth(exec,"G",34); setWidth(exec,"H",28);
setWidth(exec,"I",18); setWidth(exec,"J",15);
title(exec,"A1:J1","KẾ HOẠCH THỰC THI GTM 90 NGÀY");
exec.getRange("A3:J3").values = [["Giai đoạn","Khoảng ngày","Mục tiêu","Nhóm ưu tiên","Giả thuyết cần kiểm tra","Bản thử / cách kiểm tra","Điều kiện đạt","Đầu ra","Phụ trách","Trạng thái"]];
tableHeader(exec,"A3:J3");
const executionRows = [
  ["W1","30/09–02/10","Khởi động test và tạo baseline","Team Ads + Seller + chung","Pain manual/Chrome và demo bulk tạo signup chất lượng","3 nội dung chính + 1 demo nhỏ; hỏi 5 user; theo dõi nguồn và tạo profile đầu tiên","Có tín hiệu từ ít nhất 2 nguồn; có user tạo profile; có feedback use case rõ","Baseline và danh sách blocker","Marketing + Intern","Bị chậm"],
  ["W2","05/10–09/10","Test đủ 4 P1 ICP","4 nhóm P1","Mỗi ICP phản ứng khác với pain/proof","Mỗi nhóm có ít nhất 1 nội dung hoặc lượt seeding có chủ đích; follow-up user","Nhìn ra 1–2 nhóm nổi bật về activation, feedback, D7 hoặc tín hiệu mua","So sánh tín hiệu theo ICP","Marketing + Intern","Đang làm"],
  ["W3","12/10–16/10","Đào sâu organization/team","4 P1 + bài so sánh","Organization/handoff tạo nhu cầu tốt ở user manual","2–3 góc nội dung + 5–10 trao đổi thật; test pain team/handoff","Pain lặp ở nhiều user đúng nhóm; có dùng thật hoặc hỏi sâu về team/permission","Learning về team và handoff","Marketing + Intern","Chưa bắt đầu"],
  ["W4","19/10–23/10","Tăng proof/case demo","4 P1","Proof workflow tăng tỷ lệ tạo profile hơn content pain","So sánh content có demo với content chỉ nói pain","Demo tạo click, signup, activation hoặc feedback tốt hơn baseline","Case/demo tốt nhất","Marketing + Intern","Chưa bắt đầu"],
  ["W5","26/10–27/10","Chốt learning vòng 1","Affiliate + Seeding","End-to-end workflow tạo tín hiệu sử dụng sâu","Tổng hợp 4 tuần; không mở hướng mới","Mỗi hướng có quyết định, lý do, bài học và bước tiếp theo","Quyết định vòng 2","Marketing","Chưa bắt đầu"],
  ["Ngày 31–60","11–12/2026","Focus + Improve","2–3 hướng mạnh nhất","Onboarding và message theo use case cải thiện activation","Tối ưu landing/onboarding; sửa blocker; tiếp tục VOC và D7 follow-up","Có 1–2 ICP nổi bật; có benchmark activation và pain lặp","ICP ưu tiên v1 và onboarding use case","GTM + Product","Chưa bắt đầu"],
  ["Ngày 61–90","01/2027","Scale What Works","ICP nổi bật","Kênh/message đã chứng minh có thể tăng quy mô","Scale content/community/outbound; case study; retarget; test upgrade trigger","Kênh hiệu quả, D7 cải thiện, trigger paid rõ hơn","Primary ICP v1 và kế hoạch scale","GTM + Sales + Product","Chưa bắt đầu"],
];
exec.getRange(`A4:J${3+executionRows.length}`).values = executionRows;
bodyTable(exec,`A4:J${3+executionRows.length}`);
exec.getRange(`A4:J${3+executionRows.length}`).format.rowHeight = 76;
exec.getRange(`J4:J${3+executionRows.length}`).dataValidation = { rule: { type: "list", values: ["Chưa bắt đầu","Đang làm","Bị chậm","Hoàn thành"] } };
exec.getRange(`J4:J${3+executionRows.length}`).conditionalFormats.add("containsText", { text: "Hoàn thành", format: { fill: PALE_GREEN, font: { color: "#375623", bold: true } } });
exec.getRange(`J4:J${3+executionRows.length}`).conditionalFormats.add("containsText", { text: "Bị chậm", format: { fill: PALE_RED, font: { color: "#9C0006", bold: true } } });
exec.getRange(`J4:J${3+executionRows.length}`).conditionalFormats.add("containsText", { text: "Đang làm", format: { fill: YELLOW, font: { color: "#7F6000", bold: true } } });
exec.freezePanes.freezeRows(3);

// 3. Content and experiments.
const content = baseSheet("Nội dung & thử nghiệm");
const widths = {A:10,B:13,C:8,D:22,E:18,F:18,G:25,H:25,I:18,J:38,K:28,L:16};
for (const [c,w] of Object.entries(widths)) setWidth(content,c,w);
title(content,"A1:L1","LỊCH NỘI DUNG VÀ THỬ NGHIỆM GTM");
content.getRange("A3:L3").values = [["Mã","Ngày","Tuần","Nhóm khách hàng","Mục tiêu","Giai đoạn KH","Cụm vấn đề","Chủ đề","Định dạng","Hook / nội dung mở","CTA","Trạng thái"]];
tableHeader(content,"A3:L3");
const contentRows = [
  ["CT-001",new Date("2026-09-30"),"W1","Team Ads nhỏ-vừa","Tạo nhu cầu","Biết vấn đề","Phân quyền & bàn giao","Permission / handoff","Video","5 người chạy Ads nhưng vẫn share account bằng Chrome profile?","Xem cách setup team Ads nhỏ","Đã đăng"],
  ["CT-002",new Date("2026-10-01"),"W1","Seller nhiều shop/account","Tạo nhu cầu","Biết vấn đề","Tổ chức & tách môi trường","Mỗi shop một môi trường","Video ngắn","10 shop nhưng tất cả vẫn vận hành bằng browser/profile rời rạc?","Thử đưa 3 shop vào InviBrowser","Đã đăng"],
  ["CT-003",new Date("2026-10-02"),"W1","Chung / P1","Chứng minh sản phẩm","Biết giải pháp","Thiết lập hàng loạt","Bulk creation","Video demo","100 account không nên đồng nghĩa với 100 lần setup từ đầu.","Thử tạo một batch nhỏ","Đã đăng"],
  ["CT-004",new Date("2026-10-05"),"W2","Affiliate đang scale","Tạo nhu cầu","Biết vấn đề","Vận hành quy mô lớn","Scale từ manual","Video","Affiliate nhỏ thì vài account không vấn đề. Scale lên mới bắt đầu đau đầu.","Xem cách tổ chức theo campaign","Đã đăng"],
  ["CT-005",new Date("2026-10-06"),"W2","Seeding / Social Ops","Chứng minh","Biết giải pháp","Vận hành quy mô lớn","Sync / thao tác lặp","Video demo","20 profile cùng một tác vụ: làm tay hay đồng bộ?","Test một batch nhỏ","Đã đăng"],
  ["CT-006",new Date("2026-10-07"),"W2","Team Ads nhỏ-vừa","Chứng minh","Cân nhắc","Phân quyền & bàn giao","Phân quyền team","Video demo","Nhân viên có cần được thấy toàn bộ profile của team không?","Thử setup 2–3 thành viên","Đang sản xuất"],
  ["CT-007",new Date("2026-10-08"),"W2","Seller nhiều shop/account","Giáo dục","Biết vấn đề","Tổ chức & tách môi trường","Tổ chức nhiều shop","Bài ảnh / viết","Nhiều shop chưa đáng sợ. Đáng sợ là không còn biết shop nào đang nằm ở đâu.","Lưu checklist quản lý shop","Đang sản xuất"],
  ["CT-008",new Date("2026-10-09"),"W2","Chung / P1","Activation","Đã thử SP","Tổ chức & tách môi trường","Tạo profile đầu tiên","Tutorial video","Tạo profile đầu tiên trên InviBrowser mất bao lâu?","Tạo profile đầu tiên","Đang sản xuất"],
  ["CT-009",new Date("2026-10-12"),"W3","Affiliate đang scale","Giáo dục","Biết vấn đề","Vận hành quy mô lớn","Tổ chức theo campaign","Bài ảnh / viết","10 campaign, 20 account nhưng vẫn quản lý bằng Chrome profile?","Xem workflow mẫu","Đang sản xuất"],
  ["CT-010",new Date("2026-10-13"),"W3","Seeding / Social Ops","Chứng minh","Biết giải pháp","Thiết lập hàng loạt","Bulk setup","Video demo","Dựng một batch 50 profile mất bao lâu nếu vẫn làm từng cái?","Thử tạo batch riêng","Lên kế hoạch"],
  ["CT-011",new Date("2026-10-14"),"W3","Team Ads nhỏ-vừa","Tạo nhu cầu","Biết vấn đề","Phân quyền & bàn giao","Bàn giao nhân sự","Bài ảnh / viết","Nhân viên nghỉ việc: bạn thu hồi quyền account kiểu gì?","Xem cách bàn giao profile","Lên kế hoạch"],
  ["CT-012",new Date("2026-10-15"),"W3","Seller nhiều shop/account","Tạo nhu cầu","Biết vấn đề","Phân quyền & bàn giao","Team vận hành shop","Bài ảnh / video","Khi 3 người cùng vận hành 10 shop, vấn đề không còn chỉ là đăng nhập.","Thử setup team nhỏ","Lên kế hoạch"],
  ["CT-013",new Date("2026-10-16"),"W3","Chung / P1","Giáo dục","Biết giải pháp","Tổ chức & tách môi trường","Chrome Profile vs Multi-account","Bài so sánh","Chrome Profile dùng tốt đến khi nào?","Tự kiểm tra workflow hiện tại","Lên kế hoạch"],
  ["CT-014",new Date("2026-10-19"),"W4","Affiliate đang scale","Chứng minh","Biết giải pháp","Vận hành quy mô lớn","Workflow lặp","Video demo","Scale traffic nhưng mỗi account vẫn phải thao tác tay từng cái?","Thử một workflow mẫu","Lên kế hoạch"],
  ["CT-015",new Date("2026-10-20"),"W4","Seeding / Social Ops","Giáo dục","Biết vấn đề","Vận hành quy mô lớn","Scale account vs manpower","Bài ảnh / viết","Scale account mà số người và số thao tác cũng tăng theo thì chưa phải scale.","Xem demo Sync/Workflow","Lên kế hoạch"],
  ["CT-016",new Date("2026-10-21"),"W4","Team Ads nhỏ-vừa","Chứng minh","Cân nhắc","Phân quyền & bàn giao","Team 5 người","Case demo","Một team Ads 5 người nên tổ chức profile và quyền như thế nào?","Test 3–5 account trước","Lên kế hoạch"],
  ["CT-017",new Date("2026-10-22"),"W4","Seller nhiều shop/account","Chứng minh","Cân nhắc","Vận hành quy mô lớn","Multi-store workflow","Case demo","10 shop, một hệ thống quản lý profile: workflow trông như thế nào?","Đưa 3 shop vào test","Lên kế hoạch"],
  ["CT-018",new Date("2026-10-23"),"W4","Chung / P1","Giáo dục","Biết vấn đề","Rủi ro liên đới","Rủi ro liên đới","Bài ảnh / viết","Không bị khóa không có nghĩa mọi account đang hoàn toàn bình thường.","Kiểm tra cách tách môi trường","Lên kế hoạch"],
  ["CT-019",new Date("2026-10-26"),"W5","Affiliate đang scale","Chứng minh","Cân nhắc","Phân quyền & bàn giao","Từ solo sang team","Bài ảnh / video","Affiliate bắt đầu thuê thêm người: lúc nào Chrome profile không còn đủ?","Thử workspace nhỏ","Lên kế hoạch"],
  ["CT-020",new Date("2026-10-27"),"W5","Seeding / Social Ops","Chứng minh","Cân nhắc","Vận hành quy mô lớn","End-to-end batch workflow","Video demo","Từ tạo batch đến chạy workflow: một quy trình multi-account đầy đủ trông ra sao?","Test quy trình trên batch riêng","Lên kế hoạch"],
];
content.getRange(`A4:L${3+contentRows.length}`).values = contentRows;
bodyTable(content,`A4:L${3+contentRows.length}`);
content.getRange(`B4:B${3+contentRows.length}`).format.numberFormat = "dd/mm/yyyy";
content.getRange(`A4:L${3+contentRows.length}`).format.rowHeight = 58;
content.getRange(`L4:L${3+contentRows.length}`).dataValidation = { rule: { type: "list", values: ["Lên kế hoạch","Đang sản xuất","Đã đăng","Bị hoãn"] } };
content.getRange(`L4:L${3+contentRows.length}`).conditionalFormats.add("containsText", { text: "Đã đăng", format: { fill: PALE_GREEN, font: { color: "#375623", bold: true } } });
content.getRange(`L4:L${3+contentRows.length}`).conditionalFormats.add("containsText", { text: "Đang sản xuất", format: { fill: YELLOW, font: { color: "#7F6000", bold: true } } });
content.getRange(`L4:L${3+contentRows.length}`).conditionalFormats.add("containsText", { text: "Bị hoãn", format: { fill: PALE_RED, font: { color: "#9C0006", bold: true } } });
section(content, 26, "MA TRẬN THỬ NGHIỆM ƯU TIÊN", "L");
content.getRange("A27:G27").values = [["Mã kiểm tra","Nhóm khách hàng","Pain","Kênh","KPI chính","Điều kiện đạt","Quyết định"]];
tableHeader(content,"A27:G27");
const tests = [
  ["KH01-VD03","Team Ads nhỏ-vừa","Phân quyền & bàn giao","Facebook Group / Cold Email","Signup → Tạo profile","Có phản hồi đúng pain và user tạo profile cho account/shop", ""],
  ["KH02-VD01","Seller nhiều shop/account","Tổ chức & tách môi trường","TikTok / Group Seller","Signup → Tạo profile","Có signup đúng nhóm và đưa shop/account vào dùng thử", ""],
  ["KH03-VD02","Seeding / Social Ops","Vận hành quy mô lớn","Facebook / Community","Feature Use → D7","Có user thử batch/sync và quay lại hoặc phản hồi rõ về giảm thao tác", ""],
  ["KH04-VD02","Affiliate đang scale","Vận hành quy mô lớn","TikTok / Group Affiliate","Tạo profile → D7","Có activation, D7 hoặc tín hiệu commercial từ workflow multi-account", ""],
];
content.getRange("A28:G31").values = tests;
bodyTable(content,"A28:G31");
content.getRange("A28:G31").format.rowHeight = 52;
content.getRange("G28:G31").dataValidation = { rule: { type: "list", values: ["KEEP","SCALE","IMPROVE","PAUSE"] } };
content.freezePanes.freezeRows(3);

// 4. KPI and funnel snapshot.
const kpi = baseSheet("KPI & Funnel");
for (const [c,w] of Object.entries({A:26,B:15,C:15,D:15,E:35,F:4,G:18,H:18,I:18,J:18,K:18,L:18,M:18,N:18})) setWidth(kpi,c,w);
title(kpi,"A1:N1","KPI VÀ FUNNEL GO-TO-MARKET");
kpi.getRange("A2:E2").merge();
kpi.getRange("A2:E2").values = [["Ảnh chụp dữ liệu đến 10/10/2026. Các số retention là trạng thái đã đánh dấu trong tracker, không phải cohort đã trưởng thành đầy đủ."]];
kpi.getRange("A2:E2").format.font = { name: FONT, size: 9, italic: true, color: "#666666" };
section(kpi,4,"CHỈ SỐ HIỆN TẠI","E");
kpi.getRange("A5:E5").values = [["Chỉ số","Hiện tại","Mục tiêu","Khoảng cách","Cách đọc"]];
tableHeader(kpi,"A5:E5");
const metrics = [
  ["Signup tổng",45,"","","Tổng user có ngày đăng ký trong tracker"],
  ["Đã tạo profile",31,"","","Mốc activation dùng chung"],
  ["Tỷ lệ tạo profile",null,"","","Đã tạo profile / Signup tổng"],
  ["D7 đã đánh dấu",7,"","","User được đánh dấu còn dùng sau D7"],
  ["Signup tuần 05–11/10",7,"","","Cohort đăng ký trong tuần đang xem"],
  ["Activated tuần 05–11/10",4,"","","User trong cohort tuần đã tạo profile"],
  ["Tỷ lệ activation tuần",null,"","","Activated tuần / Signup tuần"],
  ["Nội dung đã đăng",5,"","","Trong 20 nội dung được lập lịch"],
  ["Lượt seeding ghi nhận",85,"","","Nhật ký từ 05–09/10"],
  ["Seeding đã đăng",27,"","","Trạng thái Đã đăng trong nhật ký"],
];
kpi.getRange("A6:E15").values = metrics;
kpi.getRange("B8").formulas = [["=IFERROR(B7/B6,0)"]];
kpi.getRange("B12").formulas = [["=IFERROR(B11/B10,0)"]];
kpi.getRange("D6:D15").formulasR1C1 = Array.from({length:10},()=>["=IF(RC[-1]=\"\",\"\",RC[-1]-RC[-2])"]);
bodyTable(kpi,"A6:E15");
kpi.getRange("C6:C15").format.fill = YELLOW;
kpi.getRange("C6:C15").format.font = { name: FONT, size: 10, color: "#0070C0" };
kpi.getRange("B8:B8").format.numberFormat = "0.0%";
kpi.getRange("B12:B12").format.numberFormat = "0.0%";
kpi.getRange("B6:D15").format.horizontalAlignment = "right";

section(kpi,17,"FUNNEL HIỆN TẠI","E");
kpi.getRange("A18:E18").values = [["Bước","Số lượng","Tỷ lệ từ bước trước","Định nghĩa","Lưu ý"]];
tableHeader(kpi,"A18:E18");
kpi.getRange("A19:E22").values = [
  ["Signup",45,1,"Đã đăng ký tài khoản","Tổng lịch sử trong tracker"],
  ["Tạo profile",31,null,"Đã tạo profile đầu tiên","Activation"],
  ["D7 đã đánh dấu",7,null,"Còn dùng sau D7","Cần đọc theo cohort đủ ngày"],
  ["Paid","","","Đã thanh toán","Chưa có dữ liệu được nhập"],
];
kpi.getRange("C20").formulas = [["=IFERROR(B20/B19,0)"]];
kpi.getRange("C21").formulas = [["=IFERROR(B21/B20,0)"]];
kpi.getRange("C19:C22").format.numberFormat = "0.0%";
bodyTable(kpi,"A19:E22");
const funnelChart = kpi.charts.add("bar", kpi.getRange("A18:B21"));
funnelChart.title = "Funnel người dùng hiện tại";
funnelChart.titleTextStyle.typeface = FONT;
funnelChart.titleTextStyle.fontSize = 13;
funnelChart.hasLegend = false;
funnelChart.xAxis = { axisType: "textAxis", textStyle: { typeface: FONT, fontSize: 10 } };
funnelChart.yAxis = { numberFormatCode: "0", numberFormatSourceLinked: false, textStyle: { typeface: FONT, fontSize: 10 } };
funnelChart.setPosition("G4","N18");

section(kpi,25,"TRẠNG THÁI THỰC THI","N");
kpi.getRange("A26:D26").values = [["Nội dung","Số lượng","Seeding","Số lượng"]];
tableHeader(kpi,"A26:D26");
kpi.getRange("A27:D29").values = [
  ["Lên kế hoạch",11,"Chờ duyệt",35],
  ["Đang sản xuất",4,"Đã đăng",27],
  ["Đã đăng",5,"Bị xóa / từ chối",23],
];
bodyTable(kpi,"A27:D29");
kpi.getRange("F26:N26").merge();
kpi.getRange("F26:N26").values = [["Điểm cần xử lý"]];
kpi.getRange("F26:N26").format.fill = NAVY;
kpi.getRange("F26:N26").format.font = { name: FONT, size: 10, bold: true, color: "#FFFFFF" };
kpi.getRange("F27:N30").merge(true);
kpi.getRange("F27:F30").values = [["1. Chưa có nguồn/ICP được điền cho 45 user trong Lead Pipeline."],["2. Chưa có dữ liệu Sales Stage, Paid và Revenue."],["3. W1 bị chậm; W2 đang làm nhưng tiến độ % chưa cập nhật."],["4. 23/85 lượt seeding bị xóa hoặc từ chối; cần tối ưu format và quy định nguồn."]];
kpi.getRange("F27:N30").format.wrapText = true;
kpi.getRange("F27:N30").format.borders = { preset: "all", style: "thin", color: BORDER };
kpi.getRange("F27:N30").format.rowHeight = 28;
kpi.freezePanes.freezeRows(5);

// 5. Driver-based financial and resource plan. Inputs remain blank where sources provide no real assumptions.
const fin = baseSheet("Tài chính & nguồn lực");
setWidth(fin,"A",4); setWidth(fin,"B",34);
for (const c of ["C","D","E","F","G","H","I","J","K","L","M","N"]) setWidth(fin,c,14);
title(fin,"B1:N1","KHUNG TÀI CHÍNH VÀ NGUỒN LỰC 12 THÁNG");
fin.getRange("B2:N2").merge();
fin.getRange("B2:N2").values = [["Nguồn hiện tại chưa có pricing, ngân sách, CAC, payroll hoặc forecast doanh thu. Điền các ô vàng để mô hình tự tính; không dùng các ô trống như cam kết kinh doanh."]];
fin.getRange("B2:N2").format.font = { name: FONT, size: 9, italic: true, color: "#666666" };
fin.getRange("B4:N4").values = [["Hạng mục",...Array.from({length:12},(_,i)=>new Date(2026,9+i,1))]];
tableHeader(fin,"B4:N4");
fin.getRange("C4:N4").format.numberFormat = "mm/yyyy";
section(fin,6,"GIẢ ĐỊNH ĐẦU VÀO","N");
const inputLabels = ["Website visits","Tỷ lệ signup","Tỷ lệ activation","Tỷ lệ D7","Tỷ lệ Paid từ D7","ARPU bình quân (VND)","Marketing spend (VND)","Payroll (VND)","Infrastructure / tools (VND)","Chi phí khác (VND)"];
fin.getRange("B7:B16").values = inputLabels.map(x=>[x]);
fin.getRange("B7:N16").format.borders = { preset: "all", style: "thin", color: BORDER };
fin.getRange("C7:N16").format.fill = YELLOW;
fin.getRange("C7:N16").format.font = { name: FONT, size: 10, color: "#0070C0" };
fin.getRange("C8:N11").format.numberFormat = "0.0%";
fin.getRange("C12:N16").format.numberFormat = "#,##0";
section(fin,18,"KẾT QUẢ TÍNH TOÁN","N");
const outputLabels = ["Signup dự kiến","Activated users","D7 retained users","Paid users","Doanh thu","Tổng chi phí","Lợi nhuận vận hành","CAC","Biên lợi nhuận vận hành"];
fin.getRange("B19:B27").values = outputLabels.map(x=>[x]);
for (let col=3; col<=14; col++) {
  const letter = String.fromCharCode(64+col);
  fin.getRange(`${letter}19`).formulas = [[`=IF(OR(${letter}7="",${letter}8=""),"",${letter}7*${letter}8)`]];
  fin.getRange(`${letter}20`).formulas = [[`=IF(OR(${letter}19="",${letter}9=""),"",${letter}19*${letter}9)`]];
  fin.getRange(`${letter}21`).formulas = [[`=IF(OR(${letter}20="",${letter}10=""),"",${letter}20*${letter}10)`]];
  fin.getRange(`${letter}22`).formulas = [[`=IF(OR(${letter}21="",${letter}11=""),"",${letter}21*${letter}11)`]];
  fin.getRange(`${letter}23`).formulas = [[`=IF(OR(${letter}22="",${letter}12=""),"",${letter}22*${letter}12)`]];
  fin.getRange(`${letter}24`).formulas = [[`=IF(COUNT(${letter}13:${letter}16)=0,"",SUM(${letter}13:${letter}16))`]];
  fin.getRange(`${letter}25`).formulas = [[`=IF(OR(${letter}23="",${letter}24=""),"",${letter}23-${letter}24)`]];
  fin.getRange(`${letter}26`).formulas = [[`=IFERROR(${letter}13/${letter}22,"")`]];
  fin.getRange(`${letter}27`).formulas = [[`=IFERROR(${letter}25/${letter}23,"")`]];
}
bodyTable(fin,"B19:N27");
fin.getRange("C19:N22").format.numberFormat = "#,##0";
fin.getRange("C23:N26").format.numberFormat = "#,##0";
fin.getRange("C27:N27").format.numberFormat = "0.0%";
section(fin,29,"BASELINE THAM CHIẾU","N");
fin.getRange("B30:F30").values = [["Ngày chốt","Signup tổng","Activated","D7 đã đánh dấu","Ghi chú"]];
tableHeader(fin,"B30:F30");
fin.getRange("B31:F31").values = [[new Date("2026-10-10"),45,31,7,"Snapshot từ user tracker; không coi là forecast tháng"]];
bodyTable(fin,"B31:F31");
fin.getRange("B31").format.numberFormat = "dd/mm/yyyy";
fin.freezePanes.freezeRows(4);

// 6. Sources, definitions, assumptions and research gaps.
const src = baseSheet("Nguồn & giả định");
for (const [c,w] of Object.entries({A:4,B:30,C:42,D:54,E:20,F:20})) setWidth(src,c,w);
title(src,"B1:F1","NGUỒN DỮ LIỆU VÀ GIẢ ĐỊNH");
section(src,3,"NGUỒN ĐÃ SỬ DỤNG","F");
src.getRange("B4:E4").values = [["Nguồn","Vai trò","Nội dung được dùng","Tình trạng"]];
tableHeader(src,"B4:E4");
src.getRange("B5:E8").values = [
  ["GTM_Strategy.docx","Chiến lược","Thị trường, ICP, cạnh tranh, định vị, thông điệp, kênh, KPI, roadmap","Đã chuyển"],
  ["InviBrowser_GTM_Execution_v1.9_Clean.xlsx","Thực thi","Kế hoạch tuần, test matrix, content calendar, seeding và quyết định","Đã chuyển"],
  ["InviBrowser_Sales_Marketing_Dashboard_v2.0_Clean.xlsx","Đo lường","Định nghĩa funnel, user tracker, KPI Sales/Marketing và baseline hiện tại","Đã chuyển"],
  ["Bussiness Plan mẫu.xlsx","Khung tham chiếu","Cấu trúc business plan, timeline, kế hoạch tài chính và cách trình bày","Đã áp dụng"],
];
bodyTable(src,"B5:E8");
src.getRange("B5:E8").format.rowHeight = 46;
section(src,10,"ĐỊNH NGHĨA DỮ LIỆU","F");
src.getRange("B11:E11").values = [["Thuật ngữ","Định nghĩa","Nguồn sở hữu","Lưu ý"]];
tableHeader(src,"B11:E11");
src.getRange("B12:E18").values = [
  ["Signup","Đã đăng ký tài khoản InviBrowser","Lead Pipeline / backend","Không tự động là Lead"],
  ["Activated","Đã tạo profile đầu tiên","Lead Pipeline / Product","Mốc first value dùng chung"],
  ["D7 retained","Còn dùng sau 7 ngày kể từ mốc activation","Product / follow-up","Chỉ đọc cohort đã đủ ngày"],
  ["Lead","Contact có use case hoặc pain thật","Sales","Không ép mọi signup thành Lead"],
  ["Qualified","Đúng ICP, pain rõ, có khả năng dùng thực tế","Sales","Cần ghi bằng chứng qualify"],
  ["Paid","Đã thanh toán","Billing / Sales","Tính theo ngày trả phí"],
  ["Marketing attribution","Signup/activation/D7 biết được nguồn hoạt động","Execution","Không phải tổng user"],
];
bodyTable(src,"B12:E18");
section(src,20,"GIẢ ĐỊNH VÀ PHẦN CẦN BỔ SUNG","F");
src.getRange("B21:E21").values = [["Hạng mục","Vì sao cần","Ảnh hưởng","Chủ sở hữu đề xuất"]];
tableHeader(src,"B21:E21");
src.getRange("B22:E30").values = [
  ["Pricing và ARPU","Chưa có giá hoặc blended revenue/user","Forecast doanh thu","Product + Finance"],
  ["Ngân sách marketing","Chưa có budget theo kênh","CAC và kế hoạch scale","GTM + Finance"],
  ["Chi phí nhân sự và hạ tầng","Chưa có payroll/tooling","P&L và runway","Finance + Ops"],
  ["Activation/D7 theo cohort","Tracker mới chỉ có trạng thái tổng","Ưu tiên ICP và onboarding","Product Analytics"],
  ["Nguồn và ICP của user hiện tại","45 user chưa được phân nguồn/segment","Attribution và chất lượng kênh","Marketing + Sales"],
  ["Benchmark hiệu năng/reliability","Chưa có bằng chứng so sánh","Cho phép hoặc loại bỏ claim kỹ thuật","Product"],
  ["Hands-on ease of use đối thủ","Chưa kiểm tra đủ","Thông điệp dễ dùng","Product Marketing"],
  ["Competitive creative research","Chưa audit hook/format/channel","Tối ưu content","Content"],
  ["Localization và market-share proxy","Working finding chưa xác nhận","Tiering và local positioning","GTM"],
];
bodyTable(src,"B22:E30");
src.getRange("B22:E30").format.rowHeight = 42;
src.freezePanes.freezeRows(4);

// Tab colors echo the reference workbook's restrained navy/yellow palette.
plan.tabColor = NAVY;
exec.tabColor = NAVY2;
content.tabColor = NAVY2;
kpi.tabColor = NAVY;
fin.tabColor = "#BF9000";
src.tabColor = "#A6A6A6";

wb.recalculate();

const checks = await Promise.all([
  wb.inspect({ kind: "table", range: "'Kế hoạch kinh doanh'!B1:N86", include: "values,formulas", tableMaxRows: 90, tableMaxCols: 14, maxChars: 14000 }),
  wb.inspect({ kind: "table", range: "'KPI & Funnel'!A1:N30", include: "values,formulas", tableMaxRows: 35, tableMaxCols: 14, maxChars: 10000 }),
  wb.inspect({ kind: "table", range: "'Tài chính & nguồn lực'!B4:N31", include: "values,formulas", tableMaxRows: 35, tableMaxCols: 14, maxChars: 12000 }),
  wb.inspect({ kind: "match", searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A|#NUM!|#NULL!|#SPILL!|#CALC!", options: { useRegex: true, maxResults: 300 }, summary: "final formula error scan" }),
]);
for (const c of checks) console.log(c.ndjson);

for (const ws of wb.worksheets.items) {
  const image = await wb.render({ sheetName: ws.name, autoCrop: "all", scale: 1, format: "png" });
  const safe = ws.name.replace(/[\\/:*?"<>|]/g, "_");
  await fs.writeFile(path.join(previewDir, `${safe}.png`), new Uint8Array(await image.arrayBuffer()));
}

const out = await SpreadsheetFile.exportXlsx(wb);
await out.save(outputPath);
console.log(`SAVED ${outputPath}`);
