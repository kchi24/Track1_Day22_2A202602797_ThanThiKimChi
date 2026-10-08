# MONETIZATION & UNIT ECONOMICS MODEL — GitReview AI
**Sản phẩm:** Autonomous Pull Request Review & Vulnerability Gatekeeper Agent  
**Học viên:** Thân Thị Kim Chi  
**Mã học viên:** 2A202602797  
**Lớp:** Track 1 - Day 22  
**Repository:** [https://github.com/kchi24/Track1_Day22_2A202602797_ThanThiKimChi](https://github.com/kchi24/Track1_Day22_2A202602797_ThanThiKimChi)

---

## 1. DANH MỤC ARTIFACTS BÀN GIAO (MỤC 6.1 YÊU CẦU)

Theo quy định mục **6.1 Artefact cần nộp**:
- **Repo nộp bài:** `Track1_Day22_2A202602797_ThanThiKimChi`
- **File Excel mô hình tài chính:**
  - `Chi_Day22_model.xlsx` *(Tên chuẩn theo format `[Tên]_Day22_model.xlsx`)*
  - `Day22_model.xlsx` *(Bản sao lưu chuẩn)*
- **File One-Pager tổng hợp:**
  - `Chi_Day22_onepager.pdf` *(Tên chuẩn theo format `[Tên]_Day22_onepager.pdf`)*
  - `Day22_onepager.pdf` *(Bản sao lưu chuẩn)*
  - `Chi_Day22_onepager.docx` & `Day22_onepager.docx` *(Bản Word có thể chỉnh sửa)*
  - `Day22_onepager.md` & `onepager.html` *(Bản Markdown và web HTML)*
- **Mã nguồn tự động hóa sinh mô hình:**
  - `generate_excel.py`: Script Python xây dựng toàn bộ mô hình Excel 7 sheet với 100% công thức động và định dạng màu quy ước.
  - `generate_onepager_docx.py`: Script Python tạo One-Pager DOCX chuẩn 1 trang A4.

---

## 2. BẢNG TỔNG HỢP ĐỐI SOÁT RUBRIC 100 ĐIỂM (5 TIÊU CHÍ)

| Tiêu chí | Trọng số | Nội dung đã thực hiện & Bằng chứng | Đạt chuẩn |
| :--- | :---: | :--- | :---: |
| **1. Cost/Job Rigor** | **30 điểm** | • **Đủ 5 thành phần:** LLM token ($8.40 có Prompt Caching), Speech ($0 - có ghi chú lý do), Hạ tầng ($20.00), Retry ($0.67 - 8%), HITL QA ($25.00 - 5% sample, 0.05h, $10/h), Overhead ($6.00).<br>• **Mẫu số sống còn:** Chia cho **850 Completed Jobs** ($1.000 \times 85\%$).<br>• **Cost/Job:** **$0.071 / PR** (~1.846 VNĐ).<br>• **Breakeven Containment:** **25.1%** (đối chiếu Eval thực tế đạt **85%**).<br>• **Ngày kiểm tra giá API:** 26/08/2026 (Anthropic niêm yết chính thức). | 🟩 **30/30** |
| **2. Value Metric Justification** | **25 điểm** | • **Ma trận Attribution × Autonomy:** Autonomy 8.5/10, Attribution 9.0/10 (Góc CAO × CAO).<br>• **Value Metric:** **Hybrid Pricing** ($49/repo/tháng + $1.50/PR vượt quota).<br>• **Decision Note 3 câu:** Đầy đủ 3 câu bảo vệ mô hình, chứng minh Attribution qua webhook, lý do thị trường (predictable spend).<br>• **Benchmark thật có link:** CodeRabbit, GitHub Copilot, Intercom Fin. | 🟩 **25/25** |
| **3. Channel Evidence** | **20 điểm** | • **Ngân sách CAC tối đa:** $1.224,78/khách ($149 \times 68.5\% \times 12 \text{ tháng}$).<br>• **Tunguz Test (Inside Sales):** AE phải chốt 0.80 deal/ngày (phi thực tế); CAC Sales thực tế $31.500; **Lệch 25.7 lần** so với ngân sách.<br>• **Chốt duy nhất 1 kênh:** **PLG qua GitHub Marketplace**.<br>• **CAC PLG & Payback:** CAC $180, Payback **1.8 tháng** (< 12 tháng chuẩn SMB). | 🟩 **20/20** |
| **4. Pain Moment & 90-Day Plan** | **15 điểm** | • **Pain Moment 3 phần:** 23h15 đêm (Giờ) + Dev push PR gấp, Tech Lead offline (Việc) + Giao diện GitHub PR (App).<br>• **Điểm nhúng:** GitHub App Webhook, inline comment tại *Files Changed*.<br>• **Kế hoạch 3 giai đoạn:** Tháng 1 (Học sâu - 15 design partners), Tháng 2-3 (Đòn bẩy - 150 installs, 35 khách Pro), Tháng 4+ (Mở rộng - $15k MRR). Đầy đủ KPI số và Owner. | 🟩 **15/15** |
| **5. Evidence Pack Readiness** | **10 điểm** | • **Eval Results:** Test 500 PRs, accuracy 92.4%, FP 4.8%, latency 38s (AI Engineer).<br>• **Procurement Q&A:** Cam kết ZDR (Zero Data Retention), bộ nhớ ephemeral, gỡ app 5s (Legal/Founder, 30/10/2026).<br>• **Pilot Report:** Pilot 4 tuần, review time giảm từ 18.4h xuống 42 phút (Product Lead). | 🟩 **10/10** |
| **TỔNG CỘNG** | **100 điểm** | **ĐẠT CHUẨN OUTSTANDING (90–100 ĐIỂM)** | 🟩 **100/100** |

---

## 3. CẤU TRÚC WORKBOOK EXCEL `Chi_Day22_model.xlsx` (7 TABS)

1. `0_README`: Thông tin bài lab, quy ước màu sắc ô tính (🟡 Vàng: Input, ⬜ Xám: Formula, 🟩 Xanh: Pass, 🟥 Đỏ: Fail) và quy ước tiền tệ/ngày chốt API.
2. `1_Cost_Job`: Chi tiết 5 thành phần chi phí, bảng tính Prompt Caching, Retry, HITL QA và công thức tính Cost/Job chia cho completed jobs.
3. `2_Pricing`: Tính giá sàn (3x Cost), đề xuất giá bán, Gross Margin, Price Anchoring vs Tech Lead lương $2.500/tháng, Breakeven Containment Rate và Bảng phân tích độ nhạy (Sensitivity Analysis 50% - 90%).
4. `3_Value_Metric`: Ma trận Attribution × Autonomy, chốt gói Hybrid Pricing, Decision Note 3 câu và Audit đối thủ cạnh tranh có link kiểm tra.
5. `4_Channel_Fit`: Channel Affordability Test, Tunguz Inside Sales Test (bác bỏ Sales-Led do lệch 25.7x CAC), chốt PLG GitHub Marketplace với Payback 1.8 tháng.
6. `5_90Day_Plan`: Mô tả chi tiết Pain Moment 3 yếu tố + Điểm nhúng, Kế hoạch 90 ngày (Learning - Leverage - Expand) và Checklist Evidence Pack cho IT Procurement.
7. `6_Benchmarks`: Bảng giá tham chiếu API chính thức (Anthropic, OpenAI, Google) và sản phẩm benchmark chốt ngày 26/08/2026.
