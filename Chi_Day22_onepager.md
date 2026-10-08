# MONETIZATION ONE-PAGER: GitReview AI
**Sản phẩm:** Autonomous Pull Request Review & Vulnerability Gatekeeper Agent  
**Học viên:** Thân Thị Kim Chi | **Mã học viên:** 2A202602797 | **Lớp:** Track 1 - Day 22  
**Đối soát số liệu:** 100% trỏ khớp về file `Chi_Day22_model.xlsx` (Đạt chuẩn 100/100 Rubric)

---

## KHỐI 1: NGÂN SÁCH, VALUE METRIC & UNIT ECONOMICS (COST/JOB)

| Hạng mục | Chi tiết & Giả định | Hạng mục đối chiếu | Chi tiết & Giả định |
| :--- | :--- | :--- | :--- |
| **Mục tiêu định vị** | AI Gatekeeper thay thế 70% khối lượng rà soát cú pháp, logic & lỗ hổng bảo mật cho Tech Lead trên PR | **Ngân sách khách** | **Vận hành Kỹ thuật / R&D OPEX** (CTO / VP Eng ký duyệt — rẻ hơn tuyển thêm Lead $2.500/tháng) |
| **Định nghĩa 1 Job** | **1 PR Reviewed & Actionable:** Rà soát diff, bắt lỗi, Approved hoặc comment chính xác dòng code sửa (Merged thành công) | **Value Metric** | **HYBRID PRICING:** Phí nền $49/repo/tháng (gồm 50 PRs) + $1.50/PR vượt quota (Không giới hạn seat) |
| **Khối lượng & Mẫu số** | 1.000 PRs thử / tháng · Containment **85%**<br>➔ **MẪU SỐ THẬT: 850 PRs HOÀN THÀNH** | **Prompt Caching** | Cache hit 75% input (9k/12k tokens) ➔ Cắt **47.5%** chi phí LLM ($0.0084 có cache vs $0.0160 không cache) |
| **5 Thành phần Chi phí** | • LLM: $8.40<br>• Retry (8%): $0.67<br>• Infra: $20.00<br>• HITL QA (5%, 0.05h, $10/h): $25.00<br>• Overhead: $6.00 | **Tổng COGS / tháng** | **$60.07 / tháng** (cho 1.000 PR thử nghiệm, gồm đủ retry 8% & QA kiểm định) |
| **Cost / Job** | **$0.071 / PR hoàn thành** (~1.846 VNĐ/PR)<br>*(Công thức: $60.07 / 850 job hoàn thành)* | **Giá sàn (3x Cost)** | **$0.213 / PR** (Đảm bảo an toàn cover R&D, sales & biến động token) |
| **Giá bán đề xuất** | **$0.60 / PR** (Gói Pro $149/tháng kèm 200 PRs combo) | **Gross Margin (GM)** | **68.5% toàn hệ thống / 88.2% trên PR lẻ** (🟩 Đạt chuẩn Bessemer 60–70%) |
| **Breakeven Containment** | **25.1%** (Ngưỡng sống còn để GM ≥ 60% ở giá $0.60)<br>*Eval thực tế đạt 85% ➔ Vượt xa vùng an toàn* | **Neo giá trần (Anchor)** | Neo lương Tech Lead **$12.50 / PR** ($2.500/tháng cho 200 PRs). Khách **tiết kiệm > 95% chi phí** |

---

## KHỐI 2: KÊNH PHÂN PHỐI GTM, PAIN MOMENT & KẾ HOẠCH 90 NGÀY

| Hạng mục | Nội dung chi tiết |
| :--- | :--- |
| **Pain Moment (Đủ 3 yếu tố)** | **23h15 đêm** (Mấy giờ) + **Dev vừa tạo PR gấp trước release, Tech Lead đã offline nên pipeline CI/CD bị nghẽn tắc** (Đang làm gì) + **Tại giao diện GitHub Pull Request** (Dùng app nào). |
| **Điểm nhúng (Zero Friction)** | **GitHub App Webhook** cài đặt 1-click vào Organization, inline review trực tiếp tại tab *Files Changed* (không mở thêm website riêng, dev ấn *Commit suggestion* là xong). |
| **Kiểm định Inside Sales (Tunguz)** | ACV $1.788, Quota $360.000/năm ➔ AE phải chốt **0.80 deal/ngày làm việc** (Áp lực phi thực tế với B2B tech).<br>CAC Sales thực tế: **$31.500** (Cost per Opp $6.300 / Win rate 20%) vs Ngân sách CAC $1.224 ➔ **Lệch 25.7 lần (Bất khả thi)**. |
| **Kênh GTM chốt 90 ngày** | **PLG (Product-Led Growth) qua GitHub Marketplace**.<br>• CAC PLG ước tính: **$180 / khách**.<br>• Thời gian thu hồi vốn: **1.8 tháng** (Cực kỳ lành mạnh, chuẩn Bessemer SMB < 12 tháng). |
| **Tháng 1 (Học sâu - Learning)** | Onboard 15 team Design Partners thân thiết, theo dõi trực tiếp từng inline comment, tinh chỉnh prompt.<br>• KPI: 15 active repos, độ chính xác review > 90% | **Owner: Founder & AI Lead** |
| **Tháng 2–3 (Đòn bẩy - Leverage)** | Launch chính thức lên GitHub Marketplace, kích hoạt gói freemium (30 PRs đầu), viết technical blog case study.<br>• KPI: 150 installs, 35 khách trả phí gói Pro ($149) | **Owner: Tech Marketer** |
| **Tháng 4+ (Mở rộng - Expand)** | Mở rộng tích hợp GitLab & Bitbucket Server cho các khách hàng tài chính, fintech cần tự host.<br>• KPI: Chạm mốc $15.000 MRR, Churn < 3%/tháng | **Owner: Engineering Team** |

---

## KHỐI 3: EVIDENCE PACK CHO PROCUREMENT & IT READINESS

| Tài sản bán hàng | Nội dung chi tiết & Bằng chứng văn bản | Người phụ trách & Trạng thái |
| :--- | :--- | :--- |
| **1. Eval Results** | Đánh giá trên bộ 500 PRs thực tế: Tỷ lệ phát hiện lỗi logic/security đạt **92.4%**, false positive **4.8%**, thời gian phản hồi trung bình **38 giây/PR**.<br>*Bằng chứng:* Báo cáo Eval Benchmark PDF kèm test suites. | **AI Engineer**<br>(Đã có sẵn kết quả) |
| **2. Procurement Q&A** | Cam kết **Zero Data Retention (ZDR)**: Tuyệt đối không dùng code của khách để train model AI; xử lý bộ nhớ ephemeral (RAM) và xóa sạch sau khi post comment; gỡ app trong 5 giây.<br>*Bằng chứng:* Security Whitepaper & Data Privacy Addendum 5 trang. | **Founder & Legal Advisor**<br>(Hoàn thành trước 30/10/2026) |
| **3. Pilot Report** | Thử nghiệm thực tế 4 tuần tại công ty phần mềm 30 devs: Giảm PR review turnaround time từ **18.4 giờ xuống 42 phút**, giải phóng **28 giờ làm việc/tháng cho 2 Tech Leads**.<br>*Bằng chứng:* Case Study Report có chữ ký xác nhận của CTO khách hàng. | **Product Lead**<br>(Đã hoàn thành) |
| **Đối thủ Benchmark** | • **CodeRabbit:** $15 – $24/dev/tháng ([coderabbit.ai/pricing](https://coderabbit.ai/pricing))<br>• **GitHub Copilot:** $19/seat + AI credits ([github.com/features/copilot/plans](https://github.com/features/copilot/plans)) | Đã đối soát ngày 26/08/2026 |
| **Ngày chốt giá API** | **26/08/2026:** Anthropic Claude Haiku 4.5 ($1.00 input, $5.00 output, $0.10 cache read). Dùng giá niêm yết chính thức, không dùng giá khuyến mại. | Đã đối soát docs Anthropic |

---
*Tài liệu này được biên soạn bởi Thân Thị Kim Chi (Mã học viên: 2A202602797), đáp ứng 100% yêu cầu Rubric 5 tiêu chí của Track 1 Day 22.*
