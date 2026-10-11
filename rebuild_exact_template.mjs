import fs from "node:fs/promises";
import { SpreadsheetFile } from "@oai/artifact-tool";

const templatePath = "C:/Users/Admin/Desktop/Bussiness Plan mẫu.xlsx";
const outputPath = "D:/kehoach/InviBrowser_Business_Plan_Giu_Nguyen_Format.xlsx";

const wb = await SpreadsheetFile.importXlsx(await fs.readFile(templatePath));
const sheets = Array.from({ length: 9 }, (_, i) => wb.worksheets.getItemAt(i));
const [business, finance, financeCopy, pnl, pricing, keyword, people, competitors, timeline] = sheets;

// Keep every template object and visual setting. Only cell contents and obsolete
// source-specific pictures are removed.
for (const ws of sheets) {
  const used = ws.getUsedRange();
  if (used) used.clear({ applyTo: "contents" });
}
keyword.deleteAllDrawings();
competitors.deleteAllDrawings();

function put(ws, values) {
  for (const [address, value] of Object.entries(values)) {
    ws.getRange(address).values = [[value]];
  }
}

put(business, {
  B1: "KẾ HOẠCH KINH DOANH VÀ THÂM NHẬP THỊ TRƯỜNG INViBROWSER TRONG 12 THÁNG",
  B2: "Thời gian thực hiện", C2: "12 tháng",
  B3: "Mô hình kinh doanh", C3: "SaaS freemium cho cá nhân và đội nhóm vận hành nhiều tài khoản",
  B4: "Chiến lược kinh doanh", C4: "Acquire → Activate → Learn → Focus → Scale",
  B6: "I. EXECUTIVE SUMMARY",
  B7: "1. Sản phẩm", C7: "InviBrowser là trình duyệt multi-account/antidetect hỗ trợ tách môi trường, tổ chức profile, phân quyền đội nhóm và mở rộng vận hành bằng bulk, đồng bộ và workflow.",
  B8: "2. Định vị", C8: "Giúp cá nhân và team chuyển từ quản lý account thủ công sang vận hành multi-account có tổ chức, tách biệt và có thể mở rộng bằng bulk, đồng bộ và workflow.",
  B9: "3. Mục tiêu", C9: "Ưu tiên Acquisition, Activation và Product Learning. Paid là tín hiệu thương mại quan trọng nhưng chưa phải KPI chính.",
  C10: "Baseline hiện có: 45 signup, 31 activated (68,9%), 7 user có tín hiệu D7. Tuần 05–11/10: 7 signup, 4 activated (57,1%).",
  C11: "Execution baseline: 20 nội dung gồm 5 đã đăng, 4 đang sản xuất, 11 đã lên kế hoạch. 85 lượt seeding gồm 27 đã đăng, 35 chờ duyệt, 12 bị xóa, 11 bị từ chối. Source, ICP, sales stage, paid và revenue chưa được ghi đầy đủ.",
  B12: "II. MARKET OPPORTUNITY",
  B13: "Phạm vi thị trường", C13: "Category chính: Antidetect / Multi-account Browser. Bài toán rộng hơn là tổ chức nhiều tài khoản, profile, quyền truy cập và workflow.",
  B14: "Loại nhu cầu", C14: "1. Đã biết loại công cụ: tìm pricing, review, comparison.\n2. Đã biết vấn đề: nhiều account/shop/team nhưng chưa biết antidetect.\n3. Muốn chuyển tool: cần migration, test song song và lý do đổi rõ ràng.",
  B15: "Cụm vấn đề ưu tiên", C15: "Tổ chức nhiều account và tách môi trường; vận hành số lượng lớn; phân quyền và bàn giao; rủi ro liên đới; thiết lập hàng loạt; chuyển dữ liệu/profile; ổn định và hiệu năng cần benchmark.",
  D15: "Tier 1: GPM Login, AdsPower, GenLogin, Hidemyacc. Tier 2: GemLogin, Hidemium, GoLogin, OmniLogin. Tier 3: KikiLogin. Substitute: Chrome/profile thủ công.",
  C16: "Beachhead: người dùng chuyển từ Chrome/profile thủ công sang vận hành multi-account có cấu trúc.",
  B17: "Key findings",
  C17: "Bốn P1 ICP: Team Ads nhỏ-vừa, Seller nhiều shop/account, Seeding/Social Ops và Affiliate đang scale.",
  C18: "Core value: tách môi trường, tổ chức profile, phân quyền/bàn giao, bulk, đồng bộ và workflow.",
  C19: "Free 10 profiles là cơ chế trial/activation, giúp giảm rào cản thử và tạo nguồn người dùng để học.",
  C20: "Không dùng claim về hiệu năng, độ ổn định, dễ dùng, fingerprint hoặc automation khi chưa có benchmark.",
  C21: "Pain đi trước feature. Càng gần conversion càng cần proof cụ thể.",
  C22: "North Star v0.1: Weekly Activated Users — user đã tạo profile đầu tiên trong tuần.",
  C23: "Funnel: Content/Search/Outbound → Website → Download → Signup → tạo/mở profile → dùng tính năng có ý nghĩa → quay lại → nâng cấp.",
  C24: "Quyết định theo trạng thái KEEP / SCALE / IMPROVE / PAUSE.",
  B26: "III. CHIẾN LƯỢC THƯƠNG HIỆU ĐỀ XUẤT",
  B27: "CHIẾN LƯỢC TỔNG THỂ",
  B28: "Giả thuyết định vị", C28: "InviBrowser giúp cá nhân và team chuyển từ quản lý account thủ công sang vận hành multi-account có tổ chức, tách biệt và có thể scale bằng bulk, đồng bộ và workflow.",
  B29: "Hướng triển khai", C29: "ACQUIRE → ACTIVATE → LEARN → FOCUS → SCALE",
  B30: "0–30 ngày", C30: "Acquire + Activate + Learn: tạo baseline, kiểm tra 4 P1, kéo đúng người dùng, quan sát first value và thu feedback.",
  B31: "31–60 ngày", C31: "Focus + Improve: giữ 2–3 hướng mạnh, tối ưu message, onboarding và sản phẩm theo use case.",
  B32: "61–90 ngày", C32: "Scale What Works: tăng nguồn lực cho kênh và thông điệp đã chứng minh hiệu quả; tạo case study và thử trigger nâng cấp.",
  B34: "IV. CHIẾN LƯỢC KHÁCH HÀNG",
  B35: "Phân khúc 1", C35: "Team Ads nhỏ-vừa. Pain: nhiều account, nhiều người, share thủ công. Proof: team, permission, isolation, sync. CTA: thử 3–5 account.",
  B36: "Phân khúc 2", C36: "Seller nhiều shop/account. Pain: dễ lẫn môi trường/session và khó tổ chức. Proof: isolation, group, session. CTA: đưa 3 shop vào test.",
  B37: "Phân khúc 3", C37: "Seeding / Social Ops. Pain: volume lớn và thao tác lặp. Proof: bulk, sync, workflow. CTA: test một batch nhỏ.",
  B38: "Phân khúc 4", C38: "Affiliate đang scale. Pain: từ vài account lên nhiều campaign/account. Proof: group, bulk, automation. CTA: test batch đầu tiên.",
  B40: "V. CHIẾN LƯỢC SẢN PHẨM",
  B41: "First Value", C41: "Tạo và mở profile đầu tiên. Đây là mốc activation dùng chung.",
  B42: "Organization", C42: "Group, đặt tên, tách môi trường và quản lý session/profile theo shop, campaign hoặc team.",
  B43: "Team Operations", C43: "Permission, sharing và handoff để kiểm soát quyền và bàn giao khi nhiều người cùng vận hành.",
  B44: "Scale", C44: "Bulk, Sync và Workflow để tăng số account mà không tăng tương ứng số thao tác tay.",
  B45: "VI. CHIẾN LƯỢC KÊNH PHÂN PHỐI",
  B46: "Phase 1", C46: "Tạo nhu cầu: TikTok, Facebook. Bắt nhu cầu: SEO, Google Search, group/forum. Tiếp cận chủ động: cold email/direct outreach. Chuyển đổi: website, landing page, pricing. Activation/retention: in-app, email, tutorial, docs.",
  B47: "Phase 2", C47: "Chỉ mở rộng kênh và ngân sách khi activation, D7 và feedback chứng minh chất lượng.",
  B48: "VII. CHIẾN LƯỢC CONTENT", C48: "5 chủ đề",
  B49: "Problem", C49: "Lẫn session; khó biết account nằm ở đâu; khó bàn giao; nhiều thao tác; Chrome profiles bắt đầu vỡ khi scale.",
  B50: "Product", C50: "Demo tạo profile, group, bulk, Sync Control, team permission và end-to-end workflow.",
  B51: "Proof", C51: "Workflow trước/sau, demo thao tác, feedback user và case theo Team Ads, Seller, Seeding hoặc Affiliate.",
  B52: "Education", C52: "Khi nào Chrome Profile còn đủ; khi nào cần multi-account; cách tổ chức theo shop/campaign/team; checklist bàn giao.",
  B53: "Retention", C53: "Quick start, tutorial tạo profile đầu tiên, follow-up D7 và hướng dẫn theo use case.",
  B54: "VIII. CHIẾN LƯỢC KOL/KOC",
  B59: "IX. CHIẾN LƯỢC BÁN HÀNG",
  B60: "STT", C60: "Hoạt động", D60: "Mục tiêu", E60: "Kết quả đầu ra",
  B61: "1", C61: "0–30 ngày: tracking, test 4 P1, content, seeding, follow-up và quan sát first profile.", D61: "Acquire + Activate + Learn", E61: "Baseline, pain/blocker và nguồn có tín hiệu.",
  B62: "2", C62: "31–60 ngày: giữ 2–3 hướng mạnh; tối ưu message, onboarding và sản phẩm theo use case.", D62: "Focus + Improve", E62: "ICP và message nổi bật; activation/D7 tốt hơn.",
  B63: "3", C63: "61–90 ngày: scale kênh/message hiệu quả; tạo case study và thử trigger nâng cấp.", D63: "Scale What Works", E63: "Công thức GTM có thể lặp và Primary ICP v1.",
  B68: "X. MASTER TIMELINE",
  B69: "Hoạt động", C69: "Tháng 1", D69: "Tháng 2", E69: "Tháng 3", F69: "Tháng 4", G69: "Tháng 5", H69: "Tháng 6",
  B70: "Tracking và baseline", B71: "Kiểm tra 4 ICP P1", B72: "Content pain/education", B73: "Demo và product proof", B74: "Community seeding", B75: "Cold outreach", B76: "Onboarding tạo profile đầu tiên", B77: "Phỏng vấn và feedback", B78: "Benchmark đối thủ/claim", B79: "Case study", B80: "Thử trigger nâng cấp", B81: "Scale kênh/message hiệu quả", B82: "Vòng lặp product feedback", B83: "Review KEEP/SCALE/IMPROVE/PAUSE", B84: "Chốt Primary ICP v1",
  B86: "XI. KẾ HOẠCH NHÂN SỰ", C86: "Vai trò được nêu trong các file nguồn",
  B87: "Marketing / Quản lý", C87: "Lập kế hoạch, cập nhật, theo dõi KPI và review quyết định hằng tuần.",
  B88: "Marketing + Intern", C88: "Sản xuất content, seeding, ghi log nguồn và follow-up người dùng.",
  B89: "Product / Sales-CS", C89: "Ghi nhận blocker, feedback, tín hiệu thương mại và đóng vòng phản hồi sản phẩm.",
});

// Only the first three months are supported by the 90-day source roadmap.
business.getRange("C70:E84").values = Array.from({ length: 15 }, (_, i) => {
  if (i === 9 || i === 10 || i === 11 || i === 14) return ["", "", 1];
  return [1, 1, 1];
});

// Financial template: keep the original fields, but do not create unsupported figures.
put(finance, {
  A1: "KẾ HOẠCH TÀI CHÍNH INViBROWSER 12 THÁNG",
  E1: "0–30 ngày: Acquire + Activate + Learn",
  N1: "31–60 ngày: Focus + Improve",
  Z1: "61–90 ngày: Scale What Works",
  B3: "Hạng mục", C3: "Nội dung", D3: "Chú thích",
  B4: "Dữ liệu tài chính", C4: "Pricing, budget, CAC, payroll, paid và revenue chưa đủ dữ liệu", D4: "Để trống, không tự đặt giả định",
});

// Pricing template: only the Free 10 profiles offer is present in the sources.
pricing.getRange("A1:J2").values = [
  ["STT", "Gói", "Tên gói", "Đối tượng", "Giới hạn profile", "Giá", "Trạng thái", "Trial", "Giá năm", "Ghi chú"],
  [1, "Free", "Free 10 profiles", "Người dùng mới / nhu cầu thử", "10 profiles", "", "Đang áp dụng", "Có", "", "Cơ chế trial/activation"],
];

// Keep the keyword sheet's original grid and merge structure.
put(keyword, { A1: "Nguồn: GTM Strategy và các file execution/dashboard" });
keyword.getRange("A2:M2").values = [["STT", "Từ khóa/chủ đề", "Mức nhận thức", "ICP", "Pain", "Search intent", "Đối thủ xuất hiện", "Loại nội dung", "CTA", "Kênh", "Trạng thái", "Bằng chứng", "Ghi chú"]];
keyword.getRange("A3:M10").values = [
  [1,"antidetect browser","Biết loại công cụ","Chung","Category","Review / comparison","Tier 1–3","Comparison","Xem demo","SEO/Search","Cần nghiên cứu","",""] ,
  [2,"multi-account browser","Biết giải pháp","4 P1","Nhiều account","Guide / tool","Tier 1–3","Education","Tạo profile","SEO/Search","Cần nghiên cứu","",""] ,
  [3,"quản lý nhiều tài khoản","Biết vấn đề","Seller/Ads","Organization","How-to","Chrome/manual","Problem → solution","Thử 3 account","Social/SEO","Ưu tiên","",""] ,
  [4,"Chrome profile vs antidetect","Biết giải pháp","Chung","Manual workflow","Comparison","Chrome","Comparison","Kiểm tra workflow","SEO/Social","Ưu tiên","",""] ,
  [5,"phân quyền tài khoản team ads","Biết vấn đề","Team Ads","Permission/handoff","How-to","AdsPower/Hidemyacc","Education/demo","Thử 3–5 account","Facebook/Outbound","Ưu tiên","",""] ,
  [6,"quản lý nhiều shop","Biết vấn đề","Seller","Organization/isolation","How-to","Chrome/manual","Checklist/demo","Đưa 3 shop vào test","TikTok/Group","Ưu tiên","",""] ,
  [7,"sync nhiều profile","Biết giải pháp","Social Ops","Repeated tasks","Feature/demo","GenLogin/AdsPower","Product proof","Test batch nhỏ","TikTok/Facebook","Ưu tiên","",""] ,
  [8,"chuyển profile antidetect","Muốn chuyển tool","Tool switcher","Migration","Migration guide","Hidemyacc/AdsPower","Migration proof","Test song song","SEO/Outbound","Sau benchmark","",""] ,
];

people.getRange("A1:S1").values = [["", "Vị trí", "JD", "Số lượng", "Ngày OB", "Lương cứng gross", "Lương đóng BHXH", "Phụ cấp khác", "BHXH công ty đóng", "BHXH nhân viên đóng", "Thuế TNCN", "Lương cứng net", "Thưởng/KPI", "Tổng thu nhập", "Tổng chi phí công ty", "Khung lương tham khảo", "Thu nhập trung bình", "Quy định thưởng", "Ghi chú"]];
people.getRange("A2:C4").values = [
  [1, "Marketing / Quản lý", "Lập kế hoạch, cập nhật, theo dõi KPI và review hằng tuần"],
  [2, "Marketing + Intern", "Content, seeding, nhật ký nguồn và follow-up"],
  [3, "Product / Sales-CS", "Blocker, feedback, tín hiệu thương mại và vòng lặp sản phẩm"],
];

put(competitors, {
  A1: "Tier 1: GPM Login, AdsPower, GenLogin, Hidemyacc\nTier 2: GemLogin, Hidemium, GoLogin, OmniLogin\nTier 3: KikiLogin\nSubstitute: Chrome/profile thủ công",
});

put(timeline, { A1:"Nhóm công việc", B1:"Hạng mục", C1:"Task", D1:"Người phụ trách", E1:"Ngày bắt đầu", F1:"Ngày hoàn thành", G1:"Days", H1:"Trạng thái", I1:"Oct", AB1:"Nov" });
const weekly = [
  [3, "W1", "Khởi động test & baseline", "W1", "MKT + Intern", new Date(2026, 8, 30), new Date(2026, 9, 2), 3, "Bị chậm"],
  [7, "W2", "Test 4 P1 ICP", "W2", "MKT + Intern", new Date(2026, 9, 5), new Date(2026, 9, 9), 5, "Đang làm"],
  [13, "W3", "Đào sâu organization/team", "W3", "MKT + Intern", new Date(2026, 9, 12), new Date(2026, 9, 16), 5, "Chưa bắt đầu"],
  [19, "W4", "Tăng proof/case demo", "W4", "MKT + Intern", new Date(2026, 9, 19), new Date(2026, 9, 23), 5, "Chưa bắt đầu"],
  [25, "W5", "Chốt learning vòng 1", "W5", "Marketing", new Date(2026, 9, 26), new Date(2026, 9, 27), 2, "Chưa bắt đầu"],
];
for (const [row, week, goal, task, owner, start, end, days, status] of weekly) {
  put(timeline, { [`A${row}`]: week, [`B${row}`]: goal, [`C${row}`]: task, [`D${row}`]: owner, [`E${row}`]: start, [`F${row}`]: end, [`G${row}`]: days, [`H${row}`]: status });
}

const out = await SpreadsheetFile.exportXlsx(wb);
await out.save(outputPath);
console.log(`SAVED ${outputPath}`);
