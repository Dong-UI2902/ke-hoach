import fs from "node:fs/promises";
import path from "node:path";
import { SpreadsheetFile } from "@oai/artifact-tool";

const templatePath = "C:/Users/Admin/Desktop/Bussiness Plan mẫu.xlsx";
const outputPath = "D:/kehoach/InviBrowser_Business_Plan_Theo_Mau.xlsx";
const previewDir = "D:/kehoach/qa_invibrowser_template_based";
await fs.mkdir(previewDir, { recursive: true });

const wb = await SpreadsheetFile.importXlsx(await fs.readFile(templatePath));
const NAVY = "#1F2A44";
const YELLOW = "#FFF2CC";
const INPUT = "#FFF2CC";
const BORDER = "#D9D9D9";
const GREEN = "#E2F0D9";
const FONT = "Arial";

// Keep the template's workbook architecture, widths, row styles, merges and hidden sheets.
const business = wb.worksheets.getItemAt(0);
const finance = wb.worksheets.getItemAt(1);
const financeDetail = wb.worksheets.getItemAt(2);
const pnl = wb.worksheets.getItemAt(3);
const pricing = wb.worksheets.getItemAt(4);
const keyword = wb.worksheets.getItemAt(5);
const people = wb.worksheets.getItemAt(6);
const competitors = wb.worksheets.getItemAt(7);
const timeline = wb.worksheets.getItemAt(8);

finance.name = "Kế hoạch tài chính";
financeDetail.name = "Tài chính chi tiết";
pnl.name = "Doanh thu và P&L";
pricing.name = "Bảng giá & gói dịch vụ";
keyword.name = "Từ khóa & đối thủ";

for (const ws of wb.worksheets.items) {
  const used = ws.getUsedRange();
  if (used) used.clear({ applyTo: "contents" });
  ws.showGridLines = false;
}

function setRows(ws, map) {
  for (const [cell, value] of Object.entries(map)) ws.getRange(cell).values = [[value]];
}
function rowHeight(ws, row, height) { ws.getRange(`${row}:${row}`).format.rowHeight = height; }
function header(ws, range) {
  const r = ws.getRange(range);
  r.format.fill = NAVY;
  r.format.font = { name: FONT, size: 10, bold: true, color: "#FFFFFF" };
  r.format.horizontalAlignment = "center";
  r.format.verticalAlignment = "center";
  r.format.wrapText = true;
  r.format.borders = { preset: "all", style: "thin", color: BORDER };
}
function body(ws, range) {
  const r = ws.getRange(range);
  r.format.font = { name: FONT, size: 10, color: "#1F2329" };
  r.format.wrapText = true;
  r.format.verticalAlignment = "top";
  r.format.borders = { preset: "all", style: "thin", color: BORDER };
}

// Main business-plan sheet: same section order and same coordinates as the sample.
setRows(business, {
  B1: "KẾ HOẠCH KINH DOANH VÀ THÂM NHẬP THỊ TRƯỜNG INViBROWSER TRONG 12 THÁNG",
  B2: "Thời gian thực hiện", C2: "12 tháng từ 10/2026 đến 09/2027",
  B3: "Mô hình kinh doanh", C3: "SaaS freemium cho cá nhân và đội nhóm vận hành nhiều tài khoản",
  B4: "Chiến lược kinh doanh", C4: "Acquire → Activate → Learn → Focus → Scale",
  B6: "I. EXECUTIVE SUMMARY",
  B7: "1. Sản phẩm", C7: "InviBrowser là trình duyệt multi-account/antidetect hỗ trợ tách môi trường, tổ chức profile, phân quyền đội nhóm và mở rộng vận hành bằng bulk, đồng bộ và workflow.",
  B8: "2. Định vị", C8: "Giúp cá nhân và team đang chuyển từ Chrome/profile thủ công sang một hệ thống multi-account có tổ chức, tách biệt và có thể scale. Điểm thâm nhập là người đã có pain nhiều tài khoản nhưng chưa bị khóa sâu vào công cụ đối thủ.",
  B9: "3. Mục tiêu", C9: "Ưu tiên Acquisition, Activation và Product Learning; xây công thức GTM có thể lặp lại theo ICP × Pain × Message × Channel × CTA. Paid là tín hiệu thương mại quan trọng nhưng chưa phải KPI chính.",
  C10: "Baseline 10/10/2026: 45 signup, 31 activated (68,9%), 7 user có tín hiệu D7. Tuần 05–11/10: 7 signup, 4 activated (57,1%).",
  C11: "Execution baseline: 20 nội dung — 5 đã đăng, 4 đang sản xuất, 11 đã lên kế hoạch. 85 lượt seeding — 27 đã đăng, 35 chờ duyệt, 12 bị xóa, 11 bị từ chối. Data gap: source, ICP, sales stage, paid và revenue chưa được ghi đầy đủ.",
  B12: "II. MARKET OPPORTUNITY",
  B13: "Phạm vi thị trường", C13: "Category chính: Antidetect / Multi-account Browser. Bài toán rộng hơn là tổ chức nhiều tài khoản, profile, quyền truy cập và workflow.",
  B14: "Loại nhu cầu", C14: "1. Đã biết loại công cụ: tìm pricing, review, comparison.\n2. Đã biết vấn đề: nhiều account/shop/team nhưng chưa biết antidetect.\n3. Muốn chuyển tool: cần migration, test song song và lý do đổi rõ ràng.",
  B15: "Cụm vấn đề ưu tiên", C15: "1. Tổ chức nhiều account và tách môi trường.\n2. Vận hành số lượng lớn.\n3. Phân quyền và bàn giao.\n4. Rủi ro liên đới giữa account.\n5. Thiết lập hàng loạt.\n6. Chuyển dữ liệu/profile.\n7. Độ ổn định và hiệu năng cần benchmark.",
  D15: "Thị trường có cả đối thủ trực tiếp và substitute. GPM Login mạnh về local distribution và value; AdsPower mạnh về team/permission/RPA; GenLogin mạnh về automation/no-code; Hidemyacc mạnh về share/transfer/backup; GemLogin có ecosystem; GoLogin có cloud/cross-platform. Chrome Profiles và quy trình manual là trạng thái hiện tại của nhiều user P1.",
  C16: "Cửa phù hợp nhất là manual/Chrome → structured multi-account. Không lấy power user đã bị lock-in sâu vào hệ sinh thái đối thủ làm beachhead.",
  B17: "Key findings", C17: "Core feature parity khá cao; khác biệt cạnh tranh còn nằm ở distribution, ecosystem và platform coverage.",
  C18: "Switching cost tăng mạnh khi team permission, automation và migration đã ăn sâu vào workflow.",
  C19: "Price/profile-volume là trận khó để dẫn đầu. Free 10 profile nên là cơ chế trial/activation.",
  C20: "Không claim nhẹ hơn, ổn định hơn, dễ dùng hơn, fingerprint tốt hơn hoặc automation mạnh hơn khi chưa có benchmark.",
  C21: "Tiếng Việt và local support có khả năng là hygiene factor; cần audit trước khi dùng làm USP.",
  C22: "Pain trước feature. Càng gần conversion càng cần proof cụ thể hơn.",
  C23: "North Star v0.1: Weekly Activated Users — user trong tuần đã tạo profile đầu tiên và nhận được first value.",
  C24: "Quyết định chiến lược: thắng nhóm có pain thật nhưng switching cost thấp trước, sau đó mới mở rộng sang nhóm mature.",
  B26: "III. CHIẾN LƯỢC GTM ĐỀ XUẤT",
  B27: "CHIẾN LƯỢC TỔNG THỂ",
  B28: "Giả thuyết định vị", C28: "InviBrowser giúp cá nhân và team chuyển từ quản lý account thủ công sang vận hành multi-account có tổ chức, tách biệt và có thể scale bằng bulk, đồng bộ và workflow.",
  B29: "Hướng triển khai", C29: "ACQUIRE → ACTIVATE → LEARN → FOCUS → SCALE",
  B30: "- Tháng 1 – tháng 2", C30: "BUILD & TEST: Hoàn thiện tracking, kiểm tra bốn nhóm P1, kéo user đúng nhóm và sửa blocker activation.",
  B31: "- Tháng 3 – tháng 4", C31: "FOCUS & IMPROVE: Giữ 2–3 hướng mạnh, tối ưu message, onboarding và product theo use case.",
  B32: "- Tháng 5 – tháng 6", C32: "SCALE WHAT WORKS: Tăng nguồn lực cho kênh/message đã chứng minh hiệu quả, tạo case study và test trigger nâng cấp.",
  B33: "- Tháng 7 – tháng 12", C33: "Mở rộng ICP/kênh có kiểm soát, tiếp tục benchmark sản phẩm và xây cơ chế giữ chân/nâng cấp.",
  B34: "IV. CHIẾN LƯỢC KHÁCH HÀNG",
  B35: "Phân khúc 1", C35: "Team Ads nhỏ-vừa\nPain: nhiều account, nhiều người, share thủ công.\nLời hứa: kiểm soát account và quyền truy cập gọn hơn.\nProof: Team, permission, isolation, Sync.\nCTA: thử 3–5 account.",
  B36: "Phân khúc 2", C36: "Seller nhiều shop/account\nPain: dễ lẫn môi trường/session và khó tổ chức.\nLời hứa: mỗi shop/account một môi trường rõ ràng.\nProof: isolation, group, session, team.\nCTA: đưa 3 shop vào test.",
  B37: "Phân khúc 3", C37: "Seeding / Social Ops\nPain: volume lớn và thao tác lặp.\nLời hứa: scale dàn account mà không tăng thao tác tay tương ứng.\nProof: Bulk, Sync, Workflow.\nCTA: test một batch nhỏ.",
  B38: "Phân khúc 4", C38: "Affiliate đang scale\nPain: từ vài account lên nhiều campaign/account.\nLời hứa: chuyển từ manual sang workflow multi-account có hệ thống.\nProof: Group, Bulk, Automation.\nCTA: test batch đầu tiên.",
  B39: "Nhóm sau", C39: "YouTube/content network, Freelancer/Solo, MMO advanced, Agency lớn và Enterprise tiếp tục theo dõi. Chưa phân bổ nguồn lực ngang bốn nhóm P1.",
  B40: "V. CHIẾN LƯỢC SẢN PHẨM",
  B41: "Nhóm 1: First Value", C41: "Tạo và mở profile đầu tiên. Đây là mốc activation dùng chung. Mục tiêu: giúp user đi từ signup đến first value nhanh, rõ và đo được.",
  B42: "Nhóm 2: Organization", C42: "Group, đặt tên, tách môi trường, quản lý session và cấu trúc profile theo shop/campaign/team. Mục tiêu: giải pain rối khi số account tăng.",
  B43: "Nhóm 3: Team Operations", C43: "Permission, sharing và handoff. Mục tiêu: kiểm soát quyền, bàn giao và thu hồi quyền khi nhiều người cùng vận hành.",
  B44: "Nhóm 4: Scale", C44: "Bulk, Sync và Workflow. Mục tiêu: tăng số account mà không tăng tương ứng số thao tác tay. Automation không được dùng riêng làm USP khi chưa benchmark độ sâu.",
  B45: "VI. CHIẾN LƯỢC KÊNH PHÂN PHỐI",
  B46: "Phase 1", C46: "Digital acquisition và product-led trial:\n1. TikTok/Facebook: tạo nhu cầu bằng pain và workflow.\n2. Community/Group: chạm đúng operator và tạo hội thoại.\n3. SEO/Google: đón nhu cầu category, review, comparison.\n4. Cold email/direct outreach: tiếp cận đúng ICP với CTA thử batch nhỏ.\n5. Website/pricing/use-case page: chuyển quan tâm thành signup.\n6. In-app/email/tutorial/docs: kích hoạt và giữ user.",
  B47: "Phase 2", C47: "Mở rộng sau khi đã có tín hiệu: retarget, case study, partner/referral, migration campaign và account-based outreach cho team/agency phù hợp. Chỉ tăng ngân sách khi activation, D7 và feedback chứng minh chất lượng.",
  B48: "VII. CHIẾN LƯỢC CONTENT", C48: "5 chủ đề",
  B49: "Chủ đề 1: Problem", C49: "Lẫn session; không biết account nằm ở đâu; khó bàn giao; quá nhiều thao tác; rủi ro dùng chung môi trường; Chrome profiles bắt đầu vỡ khi scale.",
  B50: "Chủ đề 2: Product", C50: "Demo tạo profile, group, bulk, Sync Control, team permission và end-to-end workflow.",
  B51: "Chủ đề 3: Proof", C51: "Workflow before/after, demo thời gian thao tác, feedback user, case theo Team Ads/Seller/Seeding/Affiliate. Không dùng claim kỹ thuật chưa chứng minh.",
  B52: "Chủ đề 4: Education", C52: "Khi nào Chrome Profile còn đủ; khi nào cần multi-account; cách tổ chức theo shop/campaign/team; checklist bàn giao và phân quyền.",
  B53: "Chủ đề 5: Retention", C53: "Quick start, tutorial tạo profile đầu tiên, follow-up D7, hướng dẫn dùng sâu hơn theo use case và thông báo cải tiến từ feedback.",
  B54: "VIII. CHIẾN LƯỢC COMMUNITY / OUTBOUND",
  B55: "Community seeding", C55: "Xây kho nguồn theo Facebook Ads, TikTok Ads, Seller, Affiliate, MMO và Social Ops. Mỗi lượt dùng nguồn ID và ghi góc tiếp cận, trạng thái, reply, hội thoại tốt, inbox/click, signup biết nguồn.",
  B56: "Góc tiếp cận", C56: "Ưu tiên nêu vấn đề và chia sẻ cách làm trước khi giới thiệu sản phẩm. So sánh hiệu quả theo mảng/ngành, nguồn và người seeding.",
  B57: "Direct outreach", C57: "Lập danh sách operator/team có pain rõ, cá nhân hóa theo số account và cách bàn giao hiện tại, mời thử 3–5 account hoặc một batch nhỏ.",
  B58: "Nguyên tắc\nQuy mô hoạt động phụ thuộc chất lượng nguồn và năng lực follow-up. Không xem số lượt đăng là kết quả nếu không có hội thoại, activation hoặc learning.",
  B59: "IX. CHIẾN LƯỢC SALES",
  B60: "STT", C60: "Hoạt động", D60: "Mục tiêu", E60: "Kết quả đầu ra",
  B61: "Tháng 1", C61: "Chuẩn hóa tracking; hoàn thiện bốn ICP/Pain test; xuất bản 3 nội dung; lập kho nguồn; hỏi 5 user; theo dõi tạo profile đầu tiên.", D61: "Tạo baseline và dữ liệu đầu tiên", E61: "Baseline acquisition/activation; danh sách pain và blocker; nguồn có tín hiệu.",
  B62: "Tháng 2", C62: "Test đủ bốn P1; luân phiên content/community; follow-up user đã signup; so sánh pain/proof; hoàn thiện tutorial first profile.", D62: "Xác định nhóm phản ứng tốt", E62: "1–2 ICP nổi bật; benchmark activation; pain lặp và message có tín hiệu.",
  B63: "Tháng 3", C63: "Đào sâu organization, handoff, permission, batch và Sync; chạy 5–10 trao đổi thật; so sánh content có demo với content chỉ nói pain.", D63: "Tìm công thức GTM có thể lặp", E63: "Ma trận ICP × Pain × Message × Channel × CTA với KEEP/SCALE/IMPROVE/PAUSE.",
  B64: "Tháng 4", C64: "Tập trung hướng thắng; sửa onboarding và blocker; tạo case demo; xây landing/use-case page; retarget user có tín hiệu.", D64: "Tăng activation và D7", E64: "Onboarding use case, case study đầu tiên và backlog product theo pain lặp.",
  B65: "Tháng 5", C65: "Scale content/community/outbound tốt; thử trigger nâng cấp; tăng proof và follow-up commercial signal.", D65: "Tăng quy mô hướng đã chứng minh", E65: "Kênh/message chủ lực và trigger paid rõ hơn.",
  B66: "Tháng 6", C66: "Review nửa năm; chốt Primary ICP v1; cập nhật roadmap, pricing test và kế hoạch 6 tháng tiếp theo.", D66: "Chốt hướng tăng trưởng", E66: "Primary ICP v1, funnel benchmark và kế hoạch scale.",
  B68: "X. MASTER TIMELINE",
  B69: "Hoạt động", C69: "Tháng 1", D69: "Tháng 2", E69: "Tháng 3", F69: "Tháng 4", G69: "Tháng 5", H69: "Tháng 6",
  B70: "Tracking và baseline", B71: "Kiểm tra 4 ICP P1", B72: "Content pain/education", B73: "Demo và product proof", B74: "Community seeding", B75: "Cold outreach", B76: "Onboarding theo use case", B77: "Customer interview & feedback", B78: "Benchmark đối thủ/claim", B79: "Case study & retarget", B80: "Kiểm tra trigger nâng cấp", B81: "Scale kênh/message hiệu quả", B82: "Pricing test", B83: "Review ICP và roadmap", B84: "Chốt kế hoạch 6 tháng tiếp",
  B86: "XI. KẾ HOẠCH NHÂN SỰ", C86: "Nhóm chức năng tối thiểu 5–6 người hoặc năng lực tương đương",
  B87: "GTM / Marketing Lead (1)", C87: "Chốt ICP, giả thuyết, message, kênh, kế hoạch test và quyết định KEEP/SCALE/IMPROVE/PAUSE.",
  B88: "Content / Video (1)", C88: "Sản xuất nội dung, demo, tutorial, tái sử dụng giữa Facebook, TikTok và website; gắn mã content/test.",
  B89: "Community / Intern (2)", C89: "Vận hành seeding, kho nguồn, log hoạt động và tín hiệu; tuân thủ quy định cộng đồng.",
  B90: "Sales/CS (1) + Product/Analytics (1)\nSales/CS qualify và follow-up; Product/Analytics theo dõi activation, D7, blocker và feedback lặp để ưu tiên roadmap.",
});
const businessHeights = {7:44,8:58,9:52,10:42,11:64,13:44,14:68,15:110,16:46,17:36,18:36,19:36,20:42,21:36,22:36,23:36,24:40,28:52,30:42,31:42,32:42,33:42,35:76,36:76,37:76,38:76,39:48,41:48,42:48,43:48,44:54,46:112,47:68,49:48,50:42,51:48,52:54,53:48,55:62,56:46,57:48,58:44,61:70,62:76,63:76,64:70,65:62,66:58,87:48,88:48,89:48,90:54};
for (const [r,h] of Object.entries(businessHeights)) rowHeight(business,Number(r),h);
const timelineMatrix = [
  [1,1,1,0,0,0],[1,1,1,1,0,0],[1,1,1,1,1,1],[1,1,1,1,1,1],[1,1,1,1,1,1],[1,1,1,1,1,1],[1,1,1,1,1,1],[1,1,1,1,1,1],[1,1,1,1,0,0],[0,0,1,1,1,1],[0,0,1,1,1,1],[0,0,0,1,1,1],[0,0,1,1,1,1],[0,0,1,1,1,1],[0,0,0,0,1,1],
];
business.getRange("C70:H84").values = timelineMatrix;

// Financial plan: retain the template's wide monthly/plan layout.
setRows(finance, {
  A1: "  KẾ HOẠCH TÀI CHÍNH INViBROWSER 12 THÁNG (10/2026–09/2027)",
  E1: "Phase 1: Acquire + Activate + Learn", N1: "Phase 2: Focus + Improve", W1: "Phase 3: Scale What Works", AF1: "Phase 4: Expand Carefully",
  E2: "Hoàn thiện tracking, kiểm tra bốn ICP, content/community/outbound và activation.",
  N2: "Giữ 2–3 hướng mạnh, tối ưu onboarding và product theo use case.",
  W2: "Scale kênh/message tốt, case study, retarget và trigger nâng cấp.",
  AF2: "Mở rộng ICP/kênh có kiểm soát; cập nhật pricing và roadmap.",
  B3: "Hạng mục", C3: "Nội dung", D3: "Chú thích",
  B4: "Kế hoạch theo gói khách hàng", B5: "Ô màu vàng là giả định cần nhập. Chưa có pricing, CAC, budget hay payroll nên không tự đặt số.",
  B6: "Sản phẩm", C6: "InviBrowser", D6: "Ba cột mỗi tháng: Solo Pro, Team Pro, Business",
  A8: 1, B8: "Funnel & doanh thu", A9: 1.1, B9: "Website visits", D9: "Input theo plan/tháng",
  A10: 1.2, B10: "Tỷ lệ signup", D10: "Signup / website visits",
  A11: 1.3, B11: "Signup", D11: "Website visits × tỷ lệ signup",
  A12: 1.4, B12: "Tỷ lệ activation", D12: "Activated / signup",
  B13: "Activated users", D13: "Signup × tỷ lệ activation",
  B14: "Tỷ lệ D7", D14: "D7 / activated",
  B15: "D7 retained users", D15: "Activated × tỷ lệ D7",
  B16: "Tỷ lệ Paid từ D7", D16: "Paid / D7",
  B17: "Paid users", D17: "D7 × tỷ lệ paid",
  B18: "ARPU bình quân", D18: "Input theo plan/tháng",
  B19: "Doanh thu", D19: "Paid users × ARPU",
  A21: 2, B21: "Tổng chi phí", A22: 2.1, B22: "Marketing spend", D22: "Paid media, content, community và outbound",
  A24: 2.2, B24: "Payroll", D24: "Lương, BHXH và thưởng",
  A26: 2.3, B26: "Infrastructure / tools", D26: "Hosting, tooling, analytics và support",
  A28: 2.4, B28: "Sales & customer support", D28: "CRM, support và vận hành khách hàng",
  A30: 2.5, B30: "Chi phí khác", D30: "Legal, office, contractor và dự phòng",
  B32: "Tổng chi phí vận hành", D32: "Tổng các khoản chi phí",
  A46: 3, B46: "Lợi nhuận vận hành", D46: "Doanh thu – tổng chi phí",
  A58: 5, B58: "Lợi nhuận trước thuế", B59: "Thuế TNDN / nghĩa vụ thuế", A60: 6, B60: "Lợi nhuận sau thuế",
});
const monthCells = ["E","H","K","N","Q","T","W","Z","AC","AF","AI","AL"];
const months = ["Tháng 10/2026","Tháng 11/2026","Tháng 12/2026","Tháng 01/2027","Tháng 02/2027","Tháng 03/2027","Tháng 04/2027","Tháng 05/2027","Tháng 06/2027","Tháng 07/2027","Tháng 08/2027","Tháng 09/2027"];
for (let i=0;i<monthCells.length;i++) {
  const start = monthCells[i];
  finance.getRange(`${start}3`).values = [[months[i]]];
  const c = String.fromCharCode(start.charCodeAt(0)+1);
  const d = String.fromCharCode(start.charCodeAt(0)+2);
  // Handle double-letter columns explicitly below.
}
const triples = [["E","F","G"],["H","I","J"],["K","L","M"],["N","O","P"],["Q","R","S"],["T","U","V"],["W","X","Y"],["Z","AA","AB"],["AC","AD","AE"],["AF","AG","AH"],["AI","AJ","AK"],["AL","AM","AN"]];
for (let i=0;i<triples.length;i++) {
  const [a,b,c] = triples[i];
  finance.getRange(`${a}3`).values = [[months[i]]];
  finance.getRange(`${a}4:${c}4`).values = [["Solo Pro","Team Pro","Business"]];
  finance.getRange(`${a}9:${c}10`).format.fill = INPUT;
  finance.getRange(`${a}12:${c}12`).format.fill = INPUT;
  finance.getRange(`${a}14:${c}14`).format.fill = INPUT;
  finance.getRange(`${a}16:${c}16`).format.fill = INPUT;
  finance.getRange(`${a}18:${c}18`).format.fill = INPUT;
  finance.getRange(`${a}22:${c}30`).format.fill = INPUT;
  for (const col of [a,b,c]) {
    finance.getRange(`${col}11`).formulas = [[`=IF(OR(${col}9="",${col}10=""),"",${col}9*${col}10)`]];
    finance.getRange(`${col}13`).formulas = [[`=IF(OR(${col}11="",${col}12=""),"",${col}11*${col}12)`]];
    finance.getRange(`${col}15`).formulas = [[`=IF(OR(${col}13="",${col}14=""),"",${col}13*${col}14)`]];
    finance.getRange(`${col}17`).formulas = [[`=IF(OR(${col}15="",${col}16=""),"",${col}15*${col}16)`]];
    finance.getRange(`${col}19`).formulas = [[`=IF(OR(${col}17="",${col}18=""),"",${col}17*${col}18)`]];
    finance.getRange(`${col}32`).formulas = [[`=IF(COUNT(${col}22,${col}24,${col}26,${col}28,${col}30)=0,"",SUM(${col}22,${col}24,${col}26,${col}28,${col}30))`]];
    finance.getRange(`${col}46`).formulas = [[`=IF(OR(${col}19="",${col}32=""),"",${col}19-${col}32)`]];
    finance.getRange(`${col}58`).formulas = [[`=${col}46`]];
    finance.getRange(`${col}60`).formulas = [[`=IF(${col}58="","",${col}58-${col}59)`]];
  }
}
finance.getRange("E10:AN10").format.numberFormat = "0.0%";
finance.getRange("E12:AN12").format.numberFormat = "0.0%";
finance.getRange("E14:AN14").format.numberFormat = "0.0%";
finance.getRange("E16:AN16").format.numberFormat = "0.0%";
finance.getRange("E9:AN60").format.wrapText = true;
finance.getRange("D7:D8").clear({ applyTo: "all" });
finance.getRange("D7").format.fill = YELLOW;
finance.getRange("D8").format.fill = "#34C724";
finance.getRange("D8").format.font = { name: FONT, size: 10, color: "#34C724" };

// Pricing sheet reuses the sample's product-price table.
pricing.getRange("I2:J46").clear({ applyTo: "all" });
pricing.getRange("A1:J1").values = [["STT","Gói","Đối tượng","Giới hạn profile","Số seat","Giá tháng (VND)","Trạng thái","Trial","Giá năm (VND)","Ghi chú"]];
pricing.getRange("A2:J6").values = [
  [1,"Free","User mới / nhu cầu nhỏ",10,1,"","Đang áp dụng","Có","","Cơ chế trial và activation"],
  [2,"Solo Pro","Freelancer / Affiliate solo","","","","Cần xác nhận","","","Chưa có pricing trong nguồn"],
  [3,"Team Pro","Team Ads / Seller / Social Ops","","","","Cần xác nhận","","","Ưu tiên permission và handoff"],
  [4,"Lifetime","User volume / nhạy giá","","","","Cần đánh giá","","","Không dùng làm trục định vị"],
  [5,"Business","Agency / team lớn","","","","Tương lai","","","Cần reliability, audit, SLA và support"],
];
pricing.getRange("I2:I6").formulas = [
  ["=IF(F2=\"\",\"\",F2*12)"],
  ["=IF(F3=\"\",\"\",F3*12)"],
  ["=IF(F4=\"\",\"\",F4*12)"],
  ["=IF(F5=\"\",\"\",F5*12)"],
  ["=IF(F6=\"\",\"\",F6*12)"],
];
pricing.getRange("A8:J8").values = [["Lưu ý","Không tự đặt giá khi chưa có quyết định pricing. Các ô trống là dữ liệu cần bổ sung.","","","","","","","",""]];
header(pricing,"A1:J1"); body(pricing,"A2:J8");
pricing.getRange("F2:F6").format.fill = INPUT; pricing.getRange("I2:I6").format.fill = "#F2F2F2";

// People sheet retains the original payroll framework.
people.getRange("A1:S1").values = [[" ","Vị trí","JD","Số lượng","Ngày OB","Lương cứng gross (điền)","Lương đóng BHXH","Phụ cấp khác","BHXH công ty đóng","BHXH nhân viên đóng","Thuế TNCN","Lương cứng net","Thưởng/KPI","Tổng thu nhập","Tổng chi phí công ty","Khung lương tham khảo","Thu nhập trung bình","Quy định thưởng","Ghi chú"]];
people.getRange("A2:S7").values = [
  [1,"GTM / Marketing Lead","ICP, test plan, message, channel, weekly review",1,new Date("2026-10-01"),"","","","","","","","","","","","","",""],
  [2,"Content / Video","Content, demo, tutorial và tái sử dụng đa kênh",1,new Date("2026-10-01"),"","","","","","","","","","","","","",""],
  [3,"Community / Intern","Kho nguồn, seeding, log tín hiệu và quy định nhóm",2,new Date("2026-10-01"),"","","","","","","","","","","","","",""],
  [4,"Sales / CS","Follow-up, qualify, objection, paid và revenue",1,new Date("2026-10-15"),"","","","","","","","","","","","","",""],
  [5,"Product / Analytics","Activation, D7, blocker, feedback và prioritization",1,new Date("2026-10-15"),"","","","","","","","","","","","","",""],
  [6,"Product / Engineering","Sửa blocker và phát triển capability theo pain lặp","Theo nguồn lực",new Date("2026-10-01"),"","","","","","","","","","","","","",""],
];
people.getRange("A9:C9").values = [["Tổng cộng"]];
people.getRange("D9").formulas = [["=SUM(D2:D7)"]];
for (const c of ["F","I","J","K","L","M","O"]) people.getRange(`${c}9`).formulas = [[`=SUM(${c}2:${c}7)`]];
people.getRange("N9").formulas = [["=L9+M9"]];
people.getRange("R11").values = [["Nguyên tắc đề xuất: KPI thưởng gắn với dữ liệu có thể kiểm chứng theo vai trò. GTM/Content/Community ưu tiên activation, D7, hội thoại chất lượng và learning; Sales ưu tiên qualified opportunity, paid và revenue; Product ưu tiên giảm blocker và cải thiện activation/retention."]];
people.getRange("E2:E7").format.numberFormat = "dd/mm/yyyy";
people.getRange("F2:O7").format.fill = INPUT;

// Competitor sheet replaces the original embedded competitor area.
competitors.unmergeCells("A1:L30");
competitors.mergeCells("A1:L1");
competitors.getRange("A1").values = [["BẢN ĐỒ ĐỐI THỦ INViBROWSER"]];
competitors.getRange("A1:L1").format.fill = YELLOW;
competitors.getRange("A1:L1").format.font = { name: FONT, size: 14, bold: true, color: "#7F6000" };
competitors.getRange("A1:L1").format.horizontalAlignment = "center";
competitors.getRange("A3:L3").values = [["Tier","Đối thủ","Điểm mạnh","Vì sao đáng chú ý","Hàm ý cho InviBrowser","Profile/Value","Team","Automation","Migration","Ecosystem","Cloud","Bằng chứng"]];
header(competitors,"A3:L3");
competitors.getRange("A4:L13").values = [
  ["Tier 1","GPM Login","Volume profile, lifetime, local distribution","Thương hiệu quen tại VN và value perception mạnh","Không đấu bằng số profile giá rẻ","Mạnh","TB","TB","TB","TB","Thấp","Có tín hiệu rõ"],
  ["Tier 1","AdsPower","Team, permission, RPA, API, migration","Workflow team tạo switching cost cao","Không lấy agency dùng sâu làm beachhead","Mạnh","Rất mạnh","Rất mạnh","Mạnh","Mạnh","Mạnh","Có bằng chứng"],
  ["Tier 1","GenLogin","Automation/no-code, Sync, mini-app","Hấp dẫn Seeding/Affiliate muốn tự động hóa","Bulk/Sync/Automation không đủ làm USP","Mạnh","Mạnh","Rất mạnh","TB","Mạnh","TB","Có bằng chứng"],
  ["Tier 1","Hidemyacc","Share, transfer, backup, session sync","Giảm rủi ro bàn giao và đổi máy/tool","Phải benchmark portability trước khi claim","Mạnh","Mạnh","Mạnh","Rất mạnh","TB","TB","Có bằng chứng"],
  ["Tier 2","GemLogin","GemStore / ecosystem automation","Marketplace tạo value và lock-in","Theo dõi ecosystem gap","Mạnh","TB","Mạnh","TB","Rất mạnh","TB","Có bằng chứng"],
  ["Tier 2","Hidemium","Cấu hình sâu, API, automation","Phù hợp operator kỹ thuật","Không lao vào advanced technical sớm","Mạnh","Mạnh","Rất mạnh","TB","TB","TB","Có tín hiệu"],
  ["Tier 2","GoLogin","Cloud, web, cross-platform","Phù hợp remote/team, giảm phụ thuộc Windows","Windows-only là gap cần theo dõi","Mạnh","Mạnh","Mạnh","Mạnh","TB","Rất mạnh","Có bằng chứng"],
  ["Tier 2","OmniLogin","Browser + proxy + automation","Integrated stack giảm nhu cầu ghép tool","Verify depth trước khi so sánh","Mạnh","TB","Mạnh","TB","TB","TB","Cần verify"],
  ["Tier 3","KikiLogin","Feature breadth","Product benchmark dù GTM VN chưa mạnh","Giữ trong capability benchmark","Mạnh","Mạnh","Mạnh","TB","TB","TB","Cần verify"],
  ["Substitute","Chrome / manual","Miễn phí, quen thuộc, không cần đổi thói quen","Trạng thái hiện tại của nhiều user P1","Wedge chính: manual → structured system","Thấp","Thấp","Thấp","Thấp","Thấp","TB","Quan sát thị trường"],
];
body(competitors,"A4:L13");
for (const [c,w] of Object.entries({A:12,B:18,C:30,D:35,E:38,F:12,G:12,H:12,I:12,J:12,K:12,L:18})) competitors.getRange(`${c}:${c}`).format.columnWidth = w;
competitors.getRange("A4:L13").format.rowHeight = 58;

// Timeline retains the template's Gantt coordinates and formatting.
setRows(timeline, {A1:"Nhóm công việc",B1:"Hạng mục",C1:"Task",D1:"Người phụ trách",E1:"Ngày bắt đầu",F1:"Ngày hoàn thành",G1:"Days",H1:"Trạng thái",I1:"Oct",AB1:"Nov"});
timeline.getRange("I2").values = [[new Date("2026-10-12")]];
timeline.getRange("J2").formulas = [["=I2+1"]];
timeline.getRange("J2:BE2").fillRight();
timeline.getRange("I2:BE2").format.numberFormat = "dd";
const groups = [
  [3,"Strategy & Tracking",["Chuẩn hóa tracking và định nghĩa funnel","Review baseline signup/activation/D7","Hoàn thiện test matrix P1","Chốt KPI và nhịp review tuần"]],
  [7,"Content",["Sản xuất content W2","Demo permission và handoff","Tutorial tạo profile đầu tiên","Lên lịch content W3","Case demo workflow","Review kết quả content"]],
  [13,"Community & Outbound",["Chuẩn hóa kho nguồn seeding","Review bài bị xóa/từ chối","Luân phiên Seller/Affiliate/Ads","Test nêu vấn đề vs chia sẻ cách làm","Cold outreach Team Ads","Review hội thoại chất lượng"]],
  [19,"Product & Onboarding",["Audit blocker tạo profile","Quick-start theo use case","Follow-up D7","Tổng hợp bug/feature request","Ưu tiên backlog theo pain lặp","Thông báo lại cho user"]],
  [25,"Sales & Learning",["Bổ sung nguồn và ICP user","Follow-up user mới trong 24h","Ghi lý do chưa dùng","Qualify signal thương mại","Theo dõi paid/revenue","Review objection/lost reason"]],
  [31,"Review & Scale",["Review KEEP/SCALE/IMPROVE/PAUSE","Chọn 1–2 ICP nổi bật","Chốt case study đầu tiên","Test retarget","Test trigger nâng cấp","Lập kế hoạch vòng tiếp theo"]],
];
let taskNo=1;
for (const [start,group,tasks] of groups) {
  timeline.getRange(`A${start}`).values = [[group]];
  for (let i=0;i<tasks.length;i++) {
    const r=start+i;
    timeline.getRange(`B${r}`).values = [[tasks[i]]];
    timeline.getRange(`C${r}`).values = [[`Task ${taskNo++}`]];
    timeline.getRange(`D${r}`).values = [[group.includes("Product")?"Product":group.includes("Sales")?"Sales/CS":"GTM / Marketing"]];
    timeline.getRange(`E${r}`).values = [[new Date(2026,9,12 + (r-3)*2)]];
    timeline.getRange(`F${r}`).values = [[new Date(2026,9,15 + (r-3)*2)]];
    timeline.getRange(`G${r}`).formulas = [[`=F${r}-E${r}+1`]];
    timeline.getRange(`H${r}`).values = [[r<7?"Đang làm":"Chưa bắt đầu"]];
  }
}
timeline.getRange("E3:F36").format.numberFormat = "dd/mm/yyyy";

// Hidden supporting sheets keep their role in the sample while using clean InviBrowser tables.
const financeDetailMerges = ["D27:D28","E1:M1","N2:P2","L3:L4","B1:D1","N3:N4","D21:D22","F3:F4","N1:V1","H2:J2","E6:G6","H7:J7","Q6:S6","T2:V2","E2:G2","K2:M2","G3:G4","M3:M4","I3:I4","U3:U4","T7:V7","E7:G7","K7:M7","N6:P6","Q2:S2","T6:V6","D33:D34","K6:M6","P3:P4","H3:H4","R3:R4","T3:T4","J3:J4","D37:D38","D25:D26","Q7:S7","D29:D30","D39:D40","N7:P7","B3:D3","K3:K4","O3:O4","E3:E4","V3:V4","Q3:Q4","D31:D32","D23:D24","S3:S4","D35:D36","B4:C4","H6:J6"];
const pnlMerges = ["AE3:AP3","AD1:AD2","Q5:AB5","AG1:AH1","AI1:AJ1","AA1:AB1","M1:N1","Q4:AB4","C1:D1","E1:F1","K1:L1","W1:X1","Y1:Z1","B1:B2","AM1:AN1","Q1:R1","AK1:AL1","AE5:AP5","AE1:AF1","AO1:AP1","C3:N3","AE4:AP4","I1:J1","Q3:AB3","G1:H1","C5:N5","S1:T1","U1:V1","P1:P2","C4:N4","A1:A2"];
for (const range of financeDetailMerges) financeDetail.unmergeCells(range);
for (const range of pnlMerges) pnl.unmergeCells(range);
keyword.unmergeCells("M2:R2");
for (const ws of [financeDetail,pnl,keyword]) ws.deleteAllDrawings();
financeDetail.getRange("A1:W50").clear({ applyTo: "formats" });
pnl.getRange("A1:AP50").clear({ applyTo: "formats" });
keyword.getRange("A1:AB30").clear({ applyTo: "formats" });

financeDetail.getRange("B1:F1").merge();
financeDetail.getRange("B1").values = [["TÀI CHÍNH CHI TIẾT INViBROWSER"]];
financeDetail.getRange("B3:F3").values = [["Hạng mục","Định nghĩa","Nguồn","Tình trạng","Ghi chú"]];
financeDetail.getRange("B4:F10").values = [
  ["Pricing / ARPU","Giá và doanh thu bình quân theo plan","Product/Finance","Cần bổ sung","Không tự đặt số"],
  ["Marketing budget","Ngân sách theo kênh/tháng","GTM/Finance","Cần bổ sung","Dùng để tính CAC"],
  ["Payroll","Chi phí nhân sự","HR/Finance","Cần bổ sung","Link sheet Nhân sự dự án"],
  ["Infrastructure","Hosting/tooling/support","Product/Finance","Cần bổ sung","Phân bổ theo tháng"],
  ["Paid conversion","D7 → Paid","Sales/Billing","Cần đo","Đọc theo cohort"],
  ["Revenue","Paid users × ARPU","Billing","Chưa có dữ liệu","Tính theo ngày trả phí"],
  ["CAC","Marketing spend / paid users","Finance","Chưa đủ dữ liệu","Không dùng zero giả"],
];
header(financeDetail,"B3:F3"); body(financeDetail,"B4:F10");
financeDetail.getRange("B1:F1").format.fill = YELLOW;
financeDetail.getRange("B1:F1").format.font = { name: FONT, size: 14, bold: true, color: "#7F6000" };
financeDetail.getRange("B1:F1").format.horizontalAlignment = "center";
for (const [c,w] of Object.entries({B:24,C:28,D:20,E:18,F:28})) financeDetail.getRange(c + ":" + c).format.columnWidth = w;

pnl.getRange("B1:N1").merge();
pnl.getRange("B1").values = [["DOANH THU VÀ P&L INViBROWSER"]];
pnl.getRange("B3:N3").values = [["Hạng mục","T10/2026","T11/2026","T12/2026","T01/2027","T02/2027","T03/2027","T04/2027","T05/2027","T06/2027","T07/2027","T08/2027","T09/2027"]];
pnl.getRange("B4:B10").values = [["Doanh thu"],["Marketing"],["Payroll"],["Infrastructure / tools"],["Chi phí khác"],["Lợi nhuận vận hành"],["Biên lợi nhuận"]];
header(pnl,"B3:N3"); body(pnl,"B4:N10");
pnl.getRange("B1:N1").format.fill = NAVY;
pnl.getRange("B1:N1").format.font = { name: FONT, size: 14, bold: true, color: "#FFFFFF" };
pnl.getRange("B1:N1").format.horizontalAlignment = "center";
pnl.getRange("B:B").format.columnWidth = 26;
pnl.getRange("C:N").format.columnWidth = 13;

keyword.getRange("A1:M1").merge();
keyword.getRange("A1").values = [["KHUNG NGHIÊN CỨU TỪ KHÓA VÀ ĐỐI THỦ"]];
keyword.getRange("A2:M2").values = [["STT","Từ khóa / chủ đề","Mức nhận thức","ICP","Pain","Search intent","Đối thủ xuất hiện","Loại nội dung","CTA","Kênh","Trạng thái","Bằng chứng","Ghi chú"]];
header(keyword,"A2:M2");
keyword.getRange("A3:M10").values = [
  [1,"antidetect browser","Biết loại công cụ","Chung","Category","Review / comparison","Cần research","Comparison","Xem demo","SEO/Search","Cần nghiên cứu","",""],
  [2,"multi-account browser","Biết giải pháp","4 P1","Nhiều account","Guide / tool","Cần research","Education","Tạo profile","SEO/Search","Cần nghiên cứu","",""],
  [3,"quản lý nhiều tài khoản","Biết vấn đề","Seller/Ads","Organization","How-to","Chrome/manual","Problem → solution","Thử 3 account","Social/SEO","Ưu tiên","",""],
  [4,"Chrome profile vs antidetect","Biết giải pháp","Chung","Manual workflow","Comparison","Chrome","Comparison","Kiểm tra workflow","SEO/Social","Ưu tiên","",""],
  [5,"phân quyền tài khoản team ads","Biết vấn đề","Team Ads","Permission/handoff","How-to","AdsPower/Hidemyacc","Education/demo","Setup 2–3 thành viên","Facebook/Outbound","Ưu tiên","",""],
  [6,"quản lý nhiều shop","Biết vấn đề","Seller","Organization/isolation","How-to","Chrome/manual","Checklist/demo","Đưa 3 shop vào test","TikTok/Group","Ưu tiên","",""],
  [7,"sync nhiều profile","Biết giải pháp","Social Ops","Repeated tasks","Feature/demo","GenLogin/AdsPower","Product proof","Test batch nhỏ","TikTok/Facebook","Ưu tiên","",""],
  [8,"chuyển profile antidetect","Muốn chuyển tool","Power user","Migration","Migration guide","Hidemyacc/AdsPower","Migration proof","Test song song","SEO/Outbound","Sau benchmark","",""],
];
body(keyword,"A3:M10");
keyword.getRange("A1:M1").format.fill = YELLOW;
keyword.getRange("A1:M1").format.font = { name: FONT, size: 14, bold: true, color: "#7F6000" };
keyword.getRange("A1:M1").format.horizontalAlignment = "center";
for (const [c,w] of Object.entries({A:7,B:24,C:18,D:18,E:22,F:18,G:22,H:22,I:18,J:16,K:16,L:20,M:24})) keyword.getRange(c + ":" + c).format.columnWidth = w;
keyword.getRange("A3:M10").format.rowHeight = 48;

wb.recalculate();
const inspections = await Promise.all([
  wb.inspect({ kind: "table", range: "'Kế hoạch kinh doanh'!B1:H90", include: "values,formulas", tableMaxRows: 95, tableMaxCols: 7, maxChars: 16000 }),
  wb.inspect({ kind: "table", range: "'Kế hoạch tài chính'!A1:AN60", include: "values,formulas", tableMaxRows: 65, tableMaxCols: 40, maxChars: 14000 }),
  wb.inspect({ kind: "table", range: "'Timeline'!A1:BE36", include: "values,formulas", tableMaxRows: 40, tableMaxCols: 57, maxChars: 12000 }),
  wb.inspect({ kind: "match", searchTerm: "NEEDLY|Needly|#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A|#NUM!|#NULL!|#SPILL!|#CALC!", options: { useRegex: true, maxResults: 300 }, summary: "stale template and formula error scan" }),
]);
for (const x of inspections) console.log(x.ndjson);

const renderSpecs = [
  [business,"B1:H90"],[finance,"A1:AN60"],[pricing,"A1:J10"],[people,"A1:S11"],[competitors,"A1:L15"],[timeline,"A1:BE36"],
  [financeDetail,"B1:F10"],[pnl,"B1:N10"],[keyword,"A1:M10"],
];
for (const [ws,range] of renderSpecs) {
  const image = await wb.render({ sheetName: ws.name, range, scale: 1, format: "png" });
  const safe = ws.name.replace(/[\\/:*?"<>|]/g,"_");
  await fs.writeFile(path.join(previewDir, `${safe}.png`), new Uint8Array(await image.arrayBuffer()));
}

const out = await SpreadsheetFile.exportXlsx(wb);
await out.save(outputPath);
console.log(`SAVED ${outputPath}`);
