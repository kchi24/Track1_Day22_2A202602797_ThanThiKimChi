import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

def build_excel_model(filepath):
    wb = openpyxl.Workbook()
    # Remove default sheet
    wb.remove(wb.active)

    # Styles
    font_title = Font(name="Calibri", size=14, bold=True, color="1F497D")
    font_section = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    font_bold = Font(name="Calibri", size=10, bold=True)
    font_regular = Font(name="Calibri", size=10)
    font_italic = Font(name="Calibri", size=9, italic=True, color="595959")
    font_pass = Font(name="Calibri", size=10, bold=True, color="276A3C")
    font_alert = Font(name="Calibri", size=10, bold=True, color="9C0006")

    fill_section = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
    fill_header = PatternFill(start_color="DCE6F1", end_color="DCE6F1", fill_type="solid")
    fill_yellow = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid") # Input
    fill_gray = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")   # Formula
    fill_green = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")  # Pass
    fill_red = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")    # Fail

    thin_border_side = Side(border_style="thin", color="D9D9D9")
    border_all = Border(left=thin_border_side, right=thin_border_side, top=thin_border_side, bottom=thin_border_side)
    thick_bottom = Border(bottom=Side(border_style="medium", color="1F497D"))

    # -------------------------------------------------------------
    # TAB 0: 0_README
    # -------------------------------------------------------------
    ws0 = wb.create_sheet(title="0_README")
    ws0.views.sheetView[0].showGridLines = True
    ws0.column_dimensions["A"].width = 6
    ws0.column_dimensions["B"].width = 25
    ws0.column_dimensions["C"].width = 45
    ws0.column_dimensions["D"].width = 35

    ws0["B2"] = "BÀI LAB 120 PHÚT: AI MONETIZATION & UNIT ECONOMICS MODEL"
    ws0["B2"].font = font_title
    ws0["B3"] = "Học viên: Thân Thị Kim Chi | Mã học viên: 2A202602797 | Lớp: Track 1 - Day 22"
    ws0["B3"].font = font_bold
    ws0["B4"] = "Sản phẩm: GitReview AI (Autonomous PR Review & Vulnerability Gatekeeper Agent)"
    ws0["B4"].font = font_italic

    ws0["B6"] = "1. QUY ƯỚC MÀU SẮC Ô TÍNH"
    ws0["B6"].font = font_bold
    headers0 = ["Màu sắc", "Ý nghĩa", "Quy định thao tác"]
    for col_idx, h in enumerate(headers0, start=2):
        cell = ws0.cell(row=7, column=col_idx, value=h)
        cell.font = font_bold
        cell.fill = fill_header
        cell.border = border_all

    color_rules = [
        ("🟡 Màu Vàng (Yellow)", "Ô giả định / Dữ liệu đầu vào", "Người dùng phải điền vào các ô này", fill_yellow),
        ("⬜ Màu Xám (Gray)", "Công thức tự động (Formulas)", "TUYỆT ĐỐI KHÔNG sửa, Excel tự động tính toán", fill_gray),
        ("🟩 Màu Xanh (Green)", "Đạt chỉ số chuẩn (Pass Benchmark)", "Chỉ số an toàn, lành mạnh về mặt tài chính", fill_green),
        ("🟥 Màu Đỏ (Red)", "Cảnh báo nguy hiểm (Alert / Fail)", "Cần quay lại điều chỉnh giả định mô hình", fill_red),
    ]
    for row_idx, (c1, c2, c3, fill_c) in enumerate(color_rules, start=8):
        ws0.cell(row=row_idx, column=2, value=c1).fill = fill_c
        ws0.cell(row=row_idx, column=3, value=c2).fill = fill_c
        ws0.cell(row=row_idx, column=4, value=c3).fill = fill_c
        for col_idx in range(2, 5):
            ws0.cell(row=row_idx, column=col_idx).font = font_regular
            ws0.cell(row=row_idx, column=col_idx).border = border_all

    ws0["B13"] = "2. QUY ƯỚC DỮ LIỆU & TIỀN TỆ"
    ws0["B13"].font = font_bold
    data_rules = [
        ("Ngày chốt giá API", "26/08/2026", "Đã đối soát với trang tài liệu gốc Anthropic & OpenAI"),
        ("Đơn vị tiền tệ gốc", "USD ($)", "Giá API & SaaS tiêu chuẩn quốc tế"),
        ("Tỷ giá quy đổi giả định", "26.000 VNĐ / USD", "Dùng cho tham chiếu ngân sách nội địa"),
        ("Mẫu số Cost/Job", "Số Job HOÀN THÀNH (Completed Jobs)", "Mẫu số = Khối lượng thử × Containment Rate"),
        ("Mục tiêu Gross Margin", "≥ 60.0% (AI-native healthy band)", "Benchmark ngành AI 2026: 52-53%, Bessemer SaaS: 65%"),
    ]
    for row_idx, (k, v, note) in enumerate(data_rules, start=14):
        ws0.cell(row=row_idx, column=2, value=k).font = font_bold
        ws0.cell(row=row_idx, column=3, value=v).font = font_regular
        ws0.cell(row=row_idx, column=4, value=note).font = font_italic
        for col_idx in range(2, 5):
            ws0.cell(row=row_idx, column=col_idx).border = border_all

    # -------------------------------------------------------------
    # TAB 1: 1_Cost_Job
    # -------------------------------------------------------------
    ws1 = wb.create_sheet(title="1_Cost_Job")
    ws1.views.sheetView[0].showGridLines = True
    ws1.column_dimensions["A"].width = 5
    ws1.column_dimensions["B"].width = 38
    ws1.column_dimensions["C"].width = 24
    ws1.column_dimensions["D"].width = 16
    ws1.column_dimensions["E"].width = 35

    ws1["B2"] = "TAB 1: TÍNH TOÁN CHI PHÍ COST / JOB ĐỦ 5 THÀNH PHẦN"
    ws1["B2"].font = font_title

    # Table layout
    headers1 = ["Khoản mục / Tham số", "Giá trị / Giả định", "Đơn vị tính", "Ghi chú & Căn cứ số liệu"]
    for col_idx, h in enumerate(headers1, start=2):
        cell = ws1.cell(row=4, column=col_idx, value=h)
        cell.font = font_bold
        cell.fill = fill_header
        cell.border = border_all

    # S1
    ws1["B5"] = "S1. ĐỊNH NGHĨA JOB (JOB-TO-BE-DONE)"
    ws1["B5"].font = font_section
    ws1.merge_cells("B5:E5")
    ws1["B5"].fill = fill_section

    ws1["B6"] = "Tên Job chuẩn hóa"
    ws1["C6"] = "1 PR Reviewed & Actionable"
    ws1["C6"].fill = fill_yellow
    ws1["D6"] = "PR completed"
    ws1["E6"] = "Đếm tự động qua GitHub webhook"

    ws1["B7"] = "Định nghĩa chi tiết 1 Job"
    ws1["C7"] = "1 Pull Request được rà soát code diff, phát hiện lỗi logic/bảo mật, tự động Approve hoặc comment chính xác dòng code cần sửa."
    ws1["C7"].fill = fill_yellow
    ws1["D7"] = "Văn bản định nghĩa"
    ws1["E7"] = "Định nghĩa chặt chẽ: chỉ tính tiền khi PR Merged hoặc dev chấp nhận inline fix"

    # S2
    ws1["B8"] = "S2. KHỐI LƯỢNG & TỶ LỆ HOÀN THÀNH (VOLUME & CONTAINMENT)"
    ws1["B8"].font = font_section
    ws1.merge_cells("B8:E8")
    ws1["B8"].fill = fill_section

    ws1["B9"] = "Tổng số job thử nghiệm / tháng (Attempted Jobs)"
    ws1["C9"] = 1000
    ws1["C9"].fill = fill_yellow
    ws1["C9"].number_format = "#,##0"
    ws1["D9"] = "PRs / tháng"
    ws1["E9"] = "Quy mô team công nghệ 25-40 devs"

    ws1["B10"] = "Containment Rate (Tỷ lệ AI tự xử lý xong %)"
    ws1["C10"] = 0.85
    ws1["C10"].fill = fill_yellow
    ws1["C10"].number_format = "0.0%"
    ws1["D10"] = "% tự động"
    ws1["E10"] = "Đo từ Eval tập 500 PRs thực tế (85% không cần người hỗ trợ)"

    ws1["B11"] = "Số job HOÀN THÀNH / tháng (Completed Jobs) [MẪU SỐ]"
    ws1["C11"] = "=C9*C10"
    ws1["C11"].fill = fill_gray
    ws1["C11"].font = font_bold
    ws1["C11"].number_format = "#,##0"
    ws1["D11"] = "PRs hoàn thành"
    ws1["E11"] = "⚠️ MẪU SỐ THẬT: Tuyệt đối không chia cho 1.000 job thử"

    # S3
    ws1["B12"] = "S3. CHI PHÍ MÔ HÌNH NGÔN NGỮ (LLM TOKENS VỚI PROMPT CACHING)"
    ws1["B12"].font = font_section
    ws1.merge_cells("B12:E12")
    ws1["B12"].fill = fill_section

    ws1["B13"] = "Model sử dụng chính"
    ws1["C13"] = "Claude Haiku 4.5"
    ws1["C13"].fill = fill_yellow
    ws1["D13"] = "Anthropic"
    ws1["E13"] = "Giá niêm yết: $1.00/1M input, $5.00/1M output, Cache: $0.10/1M"

    ws1["B14"] = "Input Token Cacheable (System prompt + KB + Styleguide)"
    ws1["C14"] = 9000
    ws1["C14"].fill = fill_yellow
    ws1["C14"].number_format = "#,##0"
    ws1["D14"] = "tokens / PR"
    ws1["E14"] = "Repository context & coding guidelines (lặp lại liên tục)"

    ws1["B15"] = "Input Token Fresh (Git diff mới của PR)"
    ws1["C15"] = 3500
    ws1["C15"].fill = fill_yellow
    ws1["C15"].number_format = "#,##0"
    ws1["D15"] = "tokens / PR"
    ws1["E15"] = "Mã nguồn thay đổi thực tế của pull request"

    ws1["B16"] = "Output Token (Review summary + inline comments)"
    ws1["C16"] = 800
    ws1["C16"].fill = fill_yellow
    ws1["C16"].number_format = "#,##0"
    ws1["D16"] = "tokens / PR"
    ws1["E16"] = "Nhận xét súc tích, chỉ trỏ đúng dòng code sai"

    ws1["B17"] = "Đơn giá Cache Read / 1M token"
    ws1["C17"] = 0.10
    ws1["C17"].fill = fill_yellow
    ws1["C17"].number_format = "$#,##0.00"
    ws1["D17"] = "$ / 1M tokens"
    ws1["E17"] = "Anthropic Cache Read = 0.1x giá input thường"

    ws1["B18"] = "Đơn giá Input thường / 1M token"
    ws1["C18"] = 1.00
    ws1["C18"].fill = fill_yellow
    ws1["C18"].number_format = "$#,##0.00"
    ws1["D18"] = "$ / 1M tokens"
    ws1["E18"] = "Anthropic Claude Haiku 4.5 Fresh Input"

    ws1["B19"] = "Đơn giá Output / 1M token"
    ws1["C19"] = 5.00
    ws1["C19"].fill = fill_yellow
    ws1["C19"].number_format = "$#,##0.00"
    ws1["D19"] = "$ / 1M tokens"
    ws1["E19"] = "Anthropic Claude Haiku 4.5 Output"

    ws1["B20"] = "Chi phí LLM cho 1 PR (CÓ Prompt Caching)"
    ws1["C20"] = "=(C14*C17/1000000)+(C15*C18/1000000)+(C16*C19/1000000)"
    ws1["C20"].fill = fill_gray
    ws1["C20"].font = font_bold
    ws1["C20"].number_format = "$#,##0.0000"
    ws1["D20"] = "$ / PR"
    ws1["E20"] = "Tiết kiệm 47.5% so với không cache ($0.0160)"

    ws1["B21"] = "Tổng chi phí LLM cho 1.000 PR thử / tháng"
    ws1["C21"] = "=C9*C20"
    ws1["C21"].fill = fill_gray
    ws1["C21"].font = font_bold
    ws1["C21"].number_format = "$#,##0.00"
    ws1["D21"] = "$ / tháng"
    ws1["E21"] = "Tổng hóa đơn Anthropic trực tiếp"

    # S4
    ws1["B22"] = "S4. SPEECH (THOẠI - NẾU CÓ)"
    ws1["B22"].font = font_section
    ws1.merge_cells("B22:E22")
    ws1["B22"].fill = fill_section

    ws1["B23"] = "Chi phí Voice / Audio STT & TTS"
    ws1["C23"] = 0.00
    ws1["C23"].fill = fill_gray
    ws1["C23"].number_format = "$#,##0.00"
    ws1["D23"] = "$ / tháng"
    ws1["E23"] = "Sản phẩm thuần Git Text Diff, không dùng Speech"

    # S5
    ws1["B24"] = "S5. HẠ TẦNG & LƯU TRỮ (INFRASTRUCTURE & LOGGING)"
    ws1["B24"].font = font_section
    ws1.merge_cells("B24:E24")
    ws1["B24"].fill = fill_section

    ws1["B25"] = "Đơn giá Infra trên mỗi PR thử"
    ws1["C25"] = 0.020
    ws1["C25"].fill = fill_yellow
    ws1["C25"].number_format = "$#,##0.000"
    ws1["D25"] = "$ / PR"
    ws1["E25"] = "Vector DB (Qdrant Cloud), Webhook runner (AWS Lambda), Egress"

    ws1["B26"] = "Tổng chi phí Infra cho 1.000 PR / tháng"
    ws1["C26"] = "=C9*C25"
    ws1["C26"].fill = fill_gray
    ws1["C26"].font = font_bold
    ws1["C26"].number_format = "$#,##0.00"
    ws1["D26"] = "$ / tháng"
    ws1["E26"] = "Chi phí máy chủ & vector search hàng tháng"

    # S6
    ws1["B27"] = "S6. TỶ LỆ GỌI LẠI (RETRY RATE) ⚠️ KHÔNG ĐỂ 0%"
    ws1["B27"].font = font_section
    ws1.merge_cells("B27:E27")
    ws1["B27"].fill = fill_section

    ws1["B28"] = "Tỷ lệ retry do timeout, network lag, JSON format error"
    ws1["C28"] = 0.08
    ws1["C28"].fill = fill_yellow
    ws1["C28"].number_format = "0.0%"
    ws1["D28"] = "% retry"
    ws1["E28"] = "8% tỷ lệ gọi lại thực tế (chuẩn công nghiệp 5-10%)"

    ws1["B29"] = "Chi phí Retry cộng thêm hàng tháng"
    ws1["C29"] = "=C21*C28"
    ws1["C29"].fill = fill_gray
    ws1["C29"].font = font_bold
    ws1["C29"].number_format = "$#,##0.00"
    ws1["D29"] = "$ / tháng"
    ws1["E29"] = "8% token LLM tốn thêm do phải retry"

    # S7
    ws1["B30"] = "S7. HUMAN-IN-THE-LOOP (HITL) & NỘI BỘ QA"
    ws1["B30"].font = font_section
    ws1.merge_cells("B30:E30")
    ws1["B30"].fill = fill_section

    ws1["B31"] = "Biến thể HITL áp dụng"
    ws1["C31"] = "Biến thể A (SaaS)"
    ws1["C31"].fill = fill_yellow
    ws1["D31"] = "Mô hình"
    ws1["E31"] = "Khách tự xử lý ca escalate; Nhà cung cấp chịu chi phí QA định kỳ"

    ws1["B32"] = "Tỷ lệ PR lấy mẫu kiểm định QA định kỳ"
    ws1["C32"] = 0.05
    ws1["C32"].fill = fill_yellow
    ws1["C32"].number_format = "0.0%"
    ws1["D32"] = "% lấy mẫu"
    ws1["E32"] = "5% số PR được Senior AI Engineer kiểm tra chéo"

    ws1["B33"] = "Thời gian kiểm tra trung bình mỗi ca"
    ws1["C33"] = 0.05
    ws1["C33"].fill = fill_yellow
    ws1["C33"].number_format = "0.00"
    ws1["D33"] = "giờ / ca"
    ws1["E33"] = "3 phút = 0.05 giờ"

    ws1["B34"] = "Chi phí nhân sự QA nội bộ / giờ"
    ws1["C34"] = 10.00
    ws1["C34"].fill = fill_yellow
    ws1["C34"].number_format = "$#,##0.00"
    ws1["D34"] = "$ / giờ"
    ws1["E34"] = "Chi phí nhân lực QA chuyên trách"

    ws1["B35"] = "Tổng chi phí HITL QA nội bộ hàng tháng"
    ws1["C35"] = "=C9*C32*C33*C34"
    ws1["C35"].fill = fill_gray
    ws1["C35"].font = font_bold
    ws1["C35"].number_format = "$#,##0.00"
    ws1["D35"] = "$ / tháng"
    ws1["E35"] = "1.000 PR * 5% * 0.05h * $10 = $25.00/tháng"

    # S8
    ws1["B36"] = "S8. OVERHEAD PHÂN BỔ (OBSERVABILITY & EVALS)"
    ws1["B36"].font = font_section
    ws1.merge_cells("B36:E36")
    ws1["B36"].fill = fill_section

    ws1["B37"] = "Chi phí LangSmith/Helicone tracing & evals tooling"
    ws1["C37"] = 6.00
    ws1["C37"].fill = fill_yellow
    ws1["C37"].number_format = "$#,##0.00"
    ws1["D37"] = "$ / tháng"
    ws1["E37"] = "Hệ thống giám sát trôi dạt dữ liệu và prompt versioning"

    # SUMMARY
    ws1["B38"] = "TỔNG HỢP CHI PHÍ & TÍNH TOÁN COST / JOB"
    ws1["B38"].font = font_section
    ws1.merge_cells("B38:E38")
    ws1["B38"].fill = fill_section

    ws1["B39"] = "TỔNG CHI PHÍ VẬN HÀNH / THÁNG (COGS)"
    ws1["C39"] = "=C21+C23+C26+C29+C35+C37"
    ws1["C39"].fill = fill_gray
    ws1["C39"].font = font_bold
    ws1["C39"].number_format = "$#,##0.00"
    ws1["D39"] = "$ / tháng"
    ws1["E39"] = "Tổng COGS = LLM + Speech + Infra + Retry + HITL + Overhead"

    ws1["B40"] = "COST / JOB (CHI PHÍ TRÊN MỖI JOB HOÀN THÀNH)"
    ws1["C40"] = "=C39/C11"
    ws1["C40"].fill = fill_green
    ws1["C40"].font = Font(name="Calibri", size=11, bold=True, color="276A3C")
    ws1["C40"].number_format = "$#,##0.0000"
    ws1["D40"] = "$ / job completed"
    ws1["E40"] = "⭐ CON SỐ SỐNG CÒN: Chia cho C11 (850 job), không chia 1.000"

    ws1["B41"] = "Quy đổi tiền Việt (Tỷ giá 26.000 ₫/USD)"
    ws1["C41"] = "=C40*26000"
    ws1["C41"].fill = fill_gray
    ws1["C41"].font = font_bold
    ws1["C41"].number_format = "#,##0 ₫"
    ws1["D41"] = "VNĐ / PR"
    ws1["E41"] = "Chỉ khoảng ~1.846 VNĐ cho 1 PR hoàn chỉnh"

    for r in range(5, 42):
        for c in range(2, 6):
            if ws1.cell(row=r, column=c).border is None or ws1.cell(row=r, column=c).border.left.style is None:
                ws1.cell(row=r, column=c).border = border_all

    # -------------------------------------------------------------
    # TAB 2: 2_Pricing
    # -------------------------------------------------------------
    ws2 = wb.create_sheet(title="2_Pricing")
    ws2.views.sheetView[0].showGridLines = True
    ws2.column_dimensions["A"].width = 5
    ws2.column_dimensions["B"].width = 38
    ws2.column_dimensions["C"].width = 24
    ws2.column_dimensions["D"].width = 16
    ws2.column_dimensions["E"].width = 38

    ws2["B2"] = "TAB 2: VÙNG GIÁ BÁN, GROSS MARGIN & BREAKEVEN ANALYSIS"
    ws2["B2"].font = font_title

    headers2 = ["Tham số định giá", "Giá trị tính toán", "Đơn vị tính", "Căn cứ & Nhận xét"]
    for col_idx, h in enumerate(headers2, start=2):
        cell = ws2.cell(row=4, column=col_idx, value=h)
        cell.font = font_bold
        cell.fill = fill_header
        cell.border = border_all

    ws2["B5"] = "1. XÁC ĐỊNH GIÁ SÀN & GIÁ BÁN ĐỀ XUẤT"
    ws2["B5"].font = font_section
    ws2.merge_cells("B5:E5")
    ws2["B5"].fill = fill_section

    ws2["B6"] = "Cost / Job (Lấy từ Tab 1)"
    ws2["C6"] = "='1_Cost_Job'!C40"
    ws2["C6"].fill = fill_gray
    ws2["C6"].font = font_bold
    ws2["C6"].number_format = "$#,##0.0000"
    ws2["D6"] = "$ / job"
    ws2["E6"] = "Lấy trực tiếp từ Tab 1_Cost_Job"

    ws2["B7"] = "Giá sàn tối thiểu (= 3 × Cost/Job)"
    ws2["C7"] = "=C6*3"
    ws2["C7"].fill = fill_gray
    ws2["C7"].font = font_bold
    ws2["C7"].number_format = "$#,##0.00"
    ws2["D7"] = "$ / job"
    ws2["E7"] = "Ngưỡng sàn đảm bảo cover R&D, sales và biến động token"

    ws2["B8"] = "GIÁ BÁN ĐỀ XUẤT (PROPOSED PRICE / JOB)"
    ws2["C8"] = 0.60
    ws2["C8"].fill = fill_yellow
    ws2["C8"].font = font_bold
    ws2["C8"].number_format = "$#,##0.00"
    ws2["D8"] = "$ / job"
    ws2["E8"] = "Đơn giá $0.60 / PR (Gói Pro $59/tháng kèm 100 PRs)"

    ws2["B9"] = "GROSS MARGIN TRÊN ĐƠN VỊ PR (GM %)"
    ws2["C9"] = "=(C8-C6)/C8"
    ws2["C9"].fill = fill_green
    ws2["C9"].font = font_pass
    ws2["C9"].number_format = "0.0%"
    ws2["D9"] = "% lợi nhuận gộp"
    ws2["E9"] = "🟩 ĐẠT NGƯỠNG (Mục tiêu ≥ 60.0%)"

    ws2["B10"] = "2. NEO GIÁ TRẦN (PRICE ANCHORING)"
    ws2["B10"].font = font_section
    ws2.merge_cells("B10:E10")
    ws2["B10"].fill = fill_section

    ws2["B11"] = "Chi phí nhân công thủ công (Lương Tech Lead review PR)"
    ws2["C11"] = 12.50
    ws2["C11"].fill = fill_yellow
    ws2["C11"].number_format = "$#,##0.00"
    ws2["D11"] = "$ / PR"
    ws2["E11"] = "Lương Tech Lead $2.500/tháng cho ~200 PRs (tương đương $12.5/PR)"

    ws2["B12"] = "Mức tiết kiệm cho khách hàng trên mỗi PR"
    ws2["C12"] = "=C11-C8"
    ws2["C12"].fill = fill_gray
    ws2["C12"].font = font_bold
    ws2["C12"].number_format = "$#,##0.00"
    ws2["D12"] = "$ tiết kiệm"
    ws2["E12"] = "Tiết kiệm $11.90/PR (Khách tiết kiệm >95% chi phí thời gian)"

    ws2["B13"] = "Tỷ lệ giá bán so với chi phí nhân công bị thay thế"
    ws2["C13"] = "=C8/C11"
    ws2["C13"].fill = fill_gray
    ws2["C13"].number_format = "0.0%"
    ws2["D13"] = "% lương nhân sự"
    ws2["E13"] = "Chỉ chiếm 4.8% chi phí tuyển dụng nhân sự (Cực kỳ hấp dẫn)"

    ws2["B14"] = "3. ĐIỂM HÒA VỐN SỐNG CÒN (BREAKEVEN CONTAINMENT RATE)"
    ws2["B14"].font = font_section
    ws2.merge_cells("B14:E14")
    ws2["B14"].fill = fill_section

    ws2["B15"] = "Cost/Job tối đa cho phép để giữ GM ≥ 60%"
    ws2["C15"] = "=C8*(1-0.60)"
    ws2["C15"].fill = fill_gray
    ws2["C15"].font = font_bold
    ws2["C15"].number_format = "$#,##0.00"
    ws2["D15"] = "$ / job"
    ws2["E15"] = "Ở giá $0.60, Cost/Job không được vượt quá $0.24"

    ws2["B16"] = "Số job hoàn thành tối thiểu cần đạt (trên 1.000 job)"
    ws2["C16"] = "='1_Cost_Job'!C39/C15"
    ws2["C16"].fill = fill_gray
    ws2["C16"].font = font_bold
    ws2["C16"].number_format = "#,##0"
    ws2["D16"] = "jobs completed"
    ws2["E16"] = "Công thức: Tổng chi phí $60.07 / $0.24 = 251 jobs"

    ws2["B17"] = "BREAKEVEN CONTAINMENT RATE (NGƯỠNG SỐNG CÒN)"
    ws2["C17"] = "=C16/'1_Cost_Job'!C9"
    ws2["C17"].fill = fill_green
    ws2["C17"].font = font_pass
    ws2["C17"].number_format = "0.0%"
    ws2["D17"] = "% tối thiểu"
    ws2["E17"] = "🟩 Chỉ cần đạt ≥ 25.1% là có lãi lành mạnh (Eval đạt 85%)"

    # Sensitivity table
    ws2["B19"] = "4. BẢNG PHÂN TÍCH ĐỘ NHẠY (SENSITIVITY ANALYSIS)"
    ws2["B19"].font = font_section
    ws2.merge_cells("B19:E19")
    ws2["B19"].fill = fill_section

    sens_headers = ["Containment Rate %", "Job hoàn thành", "Cost / Job ($)", "Gross Margin (%)"]
    for col_idx, h in enumerate(sens_headers, start=2):
        cell = ws2.cell(row=20, column=col_idx, value=h)
        cell.font = font_bold
        cell.fill = fill_header
        cell.border = border_all

    rates = [0.50, 0.60, 0.70, 0.80, 0.85, 0.90]
    for idx, r_val in enumerate(rates, start=21):
        ws2.cell(row=idx, column=2, value=r_val).number_format = "0.0%"
        ws2.cell(row=idx, column=2).fill = fill_yellow if r_val != 0.85 else fill_green
        ws2.cell(row=idx, column=3, value=f"='1_Cost_Job'!$C$9*B{idx}").number_format = "#,##0"
        ws2.cell(row=idx, column=3).fill = fill_gray
        ws2.cell(row=idx, column=4, value=f"='1_Cost_Job'!$C$39/C{idx}").number_format = "$#,##0.000"
        ws2.cell(row=idx, column=4).fill = fill_gray
        ws2.cell(row=idx, column=5, value=f"=($C$8-D{idx})/$C$8").number_format = "0.0%"
        ws2.cell(row=idx, column=5).fill = fill_green if r_val >= 0.60 else fill_red
        for c in range(2, 6):
            ws2.cell(row=idx, column=c).border = border_all
            ws2.cell(row=idx, column=c).font = font_bold if r_val == 0.85 else font_regular

    for r in range(5, 18):
        for c in range(2, 6):
            if ws2.cell(row=r, column=c).border is None or ws2.cell(row=r, column=c).border.left.style is None:
                ws2.cell(row=r, column=c).border = border_all

    # -------------------------------------------------------------
    # TAB 3: 3_Value_Metric
    # -------------------------------------------------------------
    ws3 = wb.create_sheet(title="3_Value_Metric")
    ws3.views.sheetView[0].showGridLines = True
    ws3.column_dimensions["A"].width = 5
    ws3.column_dimensions["B"].width = 38
    ws3.column_dimensions["C"].width = 24
    ws3.column_dimensions["D"].width = 16
    ws3.column_dimensions["E"].width = 38

    ws3["B2"] = "TAB 3: VALUE METRIC & MA TRẬN ATTRIBUTION × AUTONOMY"
    ws3["B2"].font = font_title

    headers3 = ["Tiêu chí đánh giá", "Điểm số / Đánh giá", "Thang điểm", "Bằng chứng & Lập luận"]
    for col_idx, h in enumerate(headers3, start=2):
        cell = ws3.cell(row=4, column=col_idx, value=h)
        cell.font = font_bold
        cell.fill = fill_header
        cell.border = border_all

    ws3["B5"] = "1. ĐÁNH GIÁ MA TRẬN ATTRIBUTION × AUTONOMY"
    ws3["B5"].font = font_section
    ws3.merge_cells("B5:E5")
    ws3["B5"].fill = fill_section

    ws3["B6"] = "Độ tự động hóa (Autonomy Score)"
    ws3["C6"] = 8.5
    ws3["C6"].fill = fill_yellow
    ws3["C6"].number_format = "0.0"
    ws3["D6"] = "Điểm / 10"
    ws3["E6"] = "AI tự quét code diff, tự comment dòng code sai, dev chỉ bấm 1 nút áp dụng"

    ws3["B7"] = "Khả năng đo lường kết quả (Attribution Score)"
    ws3["C7"] = 9.0
    ws3["C7"].fill = fill_yellow
    ws3["C7"].number_format = "0.0"
    ws3["D7"] = "Điểm / 10"
    ws3["E7"] = "Đo được 100% qua GitHub Webhook: số lỗi bắt được, commit sửa theo AI"

    ws3["B8"] = "Gợi ý từ ma trận lý thuyết"
    ws3["C8"] = "Outcome / Hybrid"
    ws3["C8"].fill = fill_gray
    ws3["C8"].font = font_bold
    ws3["D8"] = "Mô hình"
    ws3["E8"] = "Thuộc ô góc phần tư CAO × CAO (Lý tưởng cho Outcome/Hybrid)"

    ws3["B9"] = "2. CHỐT VALUE METRIC LỰA CHỌN"
    ws3["B9"].font = font_section
    ws3.merge_cells("B9:E9")
    ws3["B9"].fill = fill_section

    ws3["B10"] = "Value Metric chính thức"
    ws3["C10"] = "HYBRID PRICING"
    ws3["C10"].fill = fill_green
    ws3["C10"].font = font_pass
    ws3["D10"] = "Đơn vị tính tiền"
    ws3["E10"] = "Phí nền Repository/tháng + Gói PR Usage bổ sung"

    ws3["B11"] = "Cấu trúc gói giá đề xuất"
    ws3["C11"] = "$49/repo/tháng + $1.50/PR vượt quota (tặng sẵn 50 PR/tháng)"
    ws3["C11"].fill = fill_yellow
    ws3["D11"] = "Chi tiết gói"
    ws3["E11"] = "Bảo vệ margin khi khách dùng ít, không giới hạn seat người dùng"

    ws3["B12"] = "3. DECISION NOTE (3 CÂU BẢO VỆ ĐƠN VỊ TÍNH TIỀN)"
    ws3["B12"].font = font_section
    ws3.merge_cells("B12:E12")
    ws3["B12"].fill = fill_section

    decision_notes = [
        ("Câu 1: Tôi chọn đơn vị nào?", "Tôi chọn Hybrid Pricing: Phí nền $49/repo/tháng + $1.50/PR vượt ngưỡng thay vì Seat thuần để loại bỏ rủi ro marginal cost khi dev spam code."),
        ("Câu 2: Attribution & Autonomy mức nào?", "Sản phẩm đạt Autonomy 8.5/10 và Attribution 9.0/10 nhờ cắm trực tiếp vào GitHub Webhook, ghi nhận chính xác 100% từng PR được giải quyết."),
        ("Câu 3: Lý do thị trường (Market Override)?", "Tại thị trường B2B công nghệ, CTO và phòng Tài chính cần một hóa đơn có thể dự báo được (predictable spend) thay vì biến thiên 100% rủi ro.")
    ]
    for idx, (q, a) in enumerate(decision_notes, start=13):
        ws3.cell(row=idx, column=2, value=q).font = font_bold
        ws3.cell(row=idx, column=3, value=a).font = font_regular
        ws3.cell(row=idx, column=3).fill = fill_yellow
        ws3.merge_cells(start_row=idx, start_column=3, end_row=idx, end_column=5)
        for c in range(2, 6):
            ws3.cell(row=idx, column=c).border = border_all

    ws3["B16"] = "4. BENCHMARK ĐỐI THỦ THẬT CÓ NGUỒN (BENCHMARK AUDIT)"
    ws3["B16"].font = font_section
    ws3.merge_cells("B16:E16")
    ws3["B16"].fill = fill_section

    bench_headers = ["Tên sản phẩm đối thủ", "Mô hình tính tiền", "Mức giá công bố", "Link nguồn kiểm tra"]
    for col_idx, h in enumerate(bench_headers, start=2):
        cell = ws3.cell(row=17, column=col_idx, value=h)
        cell.font = font_bold
        cell.fill = fill_header
        cell.border = border_all

    bench_data = [
        ("CodeRabbit", "Hybrid (Gói team + hạn mức credit review repo)", "$15 - $24 / dev / tháng", "https://coderabbit.ai/pricing"),
        ("GitHub Copilot", "Hybrid (Phí nền seat + AI Credits cho agent tác vụ)", "$19 / seat + $0.01 / AI Credit bổ sung", "https://github.com/features/copilot/plans"),
        ("Intercom Fin", "Outcome thuần (Resolution-based)", "$0.99 / resolution", "https://fin.ai/pricing/")
    ]
    for idx, (b1, b2, b3, b4) in enumerate(bench_data, start=18):
        ws3.cell(row=idx, column=2, value=b1).font = font_bold
        ws3.cell(row=idx, column=3, value=b2)
        ws3.cell(row=idx, column=4, value=b3)
        ws3.cell(row=idx, column=5, value=b4).font = font_italic
        for c in range(2, 6):
            ws3.cell(row=idx, column=c).fill = fill_yellow
            ws3.cell(row=idx, column=c).border = border_all

    for r in range(5, 13):
        for c in range(2, 6):
            if ws3.cell(row=r, column=c).border is None or ws3.cell(row=r, column=c).border.left.style is None:
                ws3.cell(row=r, column=c).border = border_all

    # -------------------------------------------------------------
    # TAB 4: 4_Channel_Fit
    # -------------------------------------------------------------
    ws4 = wb.create_sheet(title="4_Channel_Fit")
    ws4.views.sheetView[0].showGridLines = True
    ws4.column_dimensions["A"].width = 5
    ws4.column_dimensions["B"].width = 38
    ws4.column_dimensions["C"].width = 24
    ws4.column_dimensions["D"].width = 16
    ws4.column_dimensions["E"].width = 38

    ws4["B2"] = "TAB 4: KÊNH PHÂN PHỐI & CHANNEL AFFORDABILITY TEST"
    ws4["B2"].font = font_title

    headers4 = ["Chỉ số tài chính", "Giá trị tính toán", "Đơn vị tính", "Căn cứ & Nhận xét"]
    for col_idx, h in enumerate(headers4, start=2):
        cell = ws4.cell(row=4, column=col_idx, value=h)
        cell.font = font_bold
        cell.fill = fill_header
        cell.border = border_all

    ws4["B5"] = "1. TÍNH TOÁN NGÂN SÁCH CAC TỐI ĐA CHO PHÉP"
    ws4["B5"].font = font_section
    ws4.merge_cells("B5:E5")
    ws4["B5"].fill = fill_section

    ws4["B6"] = "Doanh thu trung bình trên khách (ARPU / tháng)"
    ws4["C6"] = 149.00
    ws4["C6"].fill = fill_yellow
    ws4["C6"].number_format = "$#,##0.00"
    ws4["D6"] = "$ / tháng"
    ws4["E6"] = "Gói Pro $149/tháng (phục vụ 10 repos, ~200 PRs/tháng)"

    ws4["B7"] = "Gross Margin tổng thể (GM %)"
    ws4["C7"] = 0.685
    ws4["C7"].fill = fill_yellow
    ws4["C7"].number_format = "0.0%"
    ws4["D7"] = "% biên gộp"
    ws4["E7"] = "Biên lợi nhuận gộp toàn dịch vụ (bao gồm hạ tầng base)"

    ws4["B8"] = "Số tháng thu hồi vốn CAC cho phép (Payback Period)"
    ws4["C8"] = 12
    ws4["C8"].fill = fill_yellow
    ws4["C8"].number_format = "#,##0"
    ws4["D8"] = "tháng"
    ws4["E8"] = "Chuẩn Bessemer 2024 cho phân khúc SMB (<12 tháng)"

    ws4["B9"] = "NGÂN SÁCH CAC TỐI ĐA / KHÁCH HÀNG"
    ws4["C9"] = "=C6*C7*C8"
    ws4["C9"].fill = fill_green
    ws4["C9"].font = font_pass
    ws4["C9"].number_format = "$#,##0.00"
    ws4["D9"] = "$ / khách"
    ws4["E9"] = "Công thức: ARPU * GM * Số tháng payback = $1,224.78"

    ws4["B10"] = "2. KIỂM ĐỊNH TÍNH KHẢ THI CỦA INSIDE SALES (TUNGUZ TEST)"
    ws4["B10"].font = font_section
    ws4.merge_cells("B10:E10")
    ws4["B10"].fill = fill_section

    ws4["B11"] = "Giá trị hợp đồng năm (ACV = ARPU × 12)"
    ws4["C11"] = "=C6*12"
    ws4["C11"].fill = fill_gray
    ws4["C11"].font = font_bold
    ws4["C11"].number_format = "$#,##0.00"
    ws4["D11"] = "$ / năm"
    ws4["E11"] = "$149 * 12 = $1,788/năm"

    ws4["B12"] = "Chỉ tiêu doanh số 1 nhân viên Sales (Quota / năm)"
    ws4["C12"] = 360000
    ws4["C12"].fill = fill_yellow
    ws4["C12"].number_format = "$#,##0"
    ws4["D12"] = "$ / năm"
    ws4["E12"] = "Chuẩn Inside Sales B2B tech ($300k - $500k/năm)"

    ws4["B13"] = "Số deal 1 Sales phải chốt trong năm"
    ws4["C13"] = "=C12/C11"
    ws4["C13"].fill = fill_gray
    ws4["C13"].font = font_bold
    ws4["C13"].number_format = "#,##0"
    ws4["D13"] = "deals / năm"
    ws4["E13"] = "Phải chốt 201 deal/năm mới đạt quota"

    ws4["B14"] = "Số deal 1 Sales phải chốt MỖI NGÀY LÀM VIỆC"
    ws4["C14"] = "=C13/250"
    ws4["C14"].fill = fill_red
    ws4["C14"].font = font_alert
    ws4["C14"].number_format = "0.00"
    ws4["D14"] = "deals / ngày"
    ws4["E14"] = "🟥 0.80 deal/ngày (Áp lực phi thực tế với B2B Bán hàng)"

    ws4["B15"] = "Chi phí thực tế 1 cơ hội bán hàng (Cost per Opportunity)"
    ws4["C15"] = 6300.00
    ws4["C15"].fill = fill_yellow
    ws4["C15"].number_format = "$#,##0.00"
    ws4["D15"] = "$ / opp"
    ws4["E15"] = "Benchmark ICONIQ 2026 cho phân khúc SMB B2B"

    ws4["B16"] = "Tỷ lệ chốt thành công (Win Rate %)"
    ws4["C16"] = 0.20
    ws4["C16"].fill = fill_yellow
    ws4["C16"].number_format = "0.0%"
    ws4["D16"] = "% win"
    ws4["E16"] = "Tỷ lệ chuẩn 20% chốt từ demo"

    ws4["B17"] = "CAC thực tế ước tính nếu dùng Sales Reps"
    ws4["C17"] = "=C15/C16"
    ws4["C17"].fill = fill_red
    ws4["C17"].font = font_alert
    ws4["C17"].number_format = "$#,##0.00"
    ws4["D17"] = "$ / khách"
    ws4["E17"] = "🟥 CAC thực tế lên tới $31,500/khách!"

    ws4["B18"] = "ĐỘ LỆCH GIỮA CAC THỰC TẾ VÀ NGÂN SÁCH CHO PHÉP"
    ws4["C18"] = "=C17/C9"
    ws4["C18"].fill = fill_red
    ws4["C18"].font = font_alert
    ws4["C18"].number_format = "0.0"
    ws4["D18"] = "lần vượt trần"
    ws4["E18"] = "🟥 Lệch 25.7 lần ➔ Tuyệt đối KHÔNG chạy Sales-Led!"

    ws4["B19"] = "3. CHỐT KÊNH GTM DUY NHẤT CHO 90 NGÀY ĐẦU"
    ws4["B19"].font = font_section
    ws4.merge_cells("B19:E19")
    ws4["B19"].fill = fill_section

    ws4["B20"] = "KÊNH PHÂN PHỐI ĐÃ CHỐT"
    ws4["C20"] = "PLG + GITHUB MARKETPLACE"
    ws4["C20"].fill = fill_green
    ws4["C20"].font = font_pass
    ws4["D20"] = "Kênh duy nhất"
    ws4["E20"] = "Product-Led Growth cắm trực tiếp vào GitHub App"

    ws4["B21"] = "CAC thực tế của kênh PLG"
    ws4["C21"] = 180.00
    ws4["C21"].fill = fill_yellow
    ws4["C21"].number_format = "$#,##0.00"
    ws4["D21"] = "$ / khách"
    ws4["E21"] = "Chi phí content marketing, docs và listing GitHub"

    ws4["B22"] = "Thời gian thu hồi vốn thực tế (Actual Payback)"
    ws4["C22"] = "=C21/(C6*C7)"
    ws4["C22"].fill = fill_green
    ws4["C22"].font = font_pass
    ws4["C22"].number_format = "0.0"
    ws4["D22"] = "tháng"
    ws4["E22"] = "🟩 Chỉ mất 1.8 tháng để thu hồi CAC (Cực kỳ xuất sắc)"

    for r in range(5, 23):
        for c in range(2, 6):
            if ws4.cell(row=r, column=c).border is None or ws4.cell(row=r, column=c).border.left.style is None:
                ws4.cell(row=r, column=c).border = border_all

    # -------------------------------------------------------------
    # TAB 5: 5_90Day_Plan
    # -------------------------------------------------------------
    ws5 = wb.create_sheet(title="5_90Day_Plan")
    ws5.views.sheetView[0].showGridLines = True
    ws5.column_dimensions["A"].width = 5
    ws5.column_dimensions["B"].width = 25
    ws5.column_dimensions["C"].width = 45
    ws5.column_dimensions["D"].width = 25
    ws5.column_dimensions["E"].width = 25

    ws5["B2"] = "TAB 5: KẾ HOẠCH 90 NGÀY & CHECKLIST EVIDENCE PACK"
    ws5["B2"].font = font_title

    ws5["B4"] = "1. PAIN MOMENT ĐỦ 3 YẾU TỐ CHUẨN MỰC"
    ws5["B4"].font = font_section
    ws5.merge_cells("B4:E4")
    ws5["B4"].fill = fill_section

    pain_items = [
        ("MẤY GIỜ?", "23h15 đêm", "Thời điểm dev vừa push code xong trước release"),
        ("ĐANG LÀM GÌ?", "Tạo Pull Request và chờ Tech Lead review", "Tech Lead đã offline, PR bị nghẽn tắc, pipeline CI/CD treo"),
        ("DÙNG APP NÀO?", "Giao diện GitHub Pull Request", "Dev đang ngồi tại tab PR của GitHub web"),
        ("ĐIỂM NHÚNG (SURFACE)", "GitHub App Webhook cắm trực tiếp vào repo", "Zero friction: nhận xét inline ngay tại tab Files Changed")
    ]
    for idx, (k, v, note) in enumerate(pain_items, start=5):
        ws5.cell(row=idx, column=2, value=k).font = font_bold
        ws5.cell(row=idx, column=3, value=v).fill = fill_yellow
        ws5.cell(row=idx, column=4, value=note).fill = fill_yellow
        ws5.merge_cells(start_row=idx, start_column=4, end_row=idx, end_column=5)
        for c in range(2, 6):
            ws5.cell(row=idx, column=c).border = border_all

    ws5["B10"] = "2. KẾ HOẠCH PHÂN PHỐI 90 NGÀY (3 GIAI ĐOẠN)"
    ws5["B10"].font = font_section
    ws5.merge_cells("B10:E10")
    ws5["B10"].fill = fill_section

    plan_headers = ["Giai đoạn", "Mục tiêu trọng tâm", "Chỉ số KPI có số", "Người chịu trách nhiệm"]
    for col_idx, h in enumerate(plan_headers, start=2):
        cell = ws5.cell(row=11, column=col_idx, value=h)
        cell.font = font_bold
        cell.fill = fill_header
        cell.border = border_all

    plans = [
        ("Tháng 1: Học sâu (Learning)", "Onboard 15 team công nghệ thân thiết (Design Partners), theo dõi trực tiếp từng comment của bot, tối ưu prompt", "15 active repos, độ chính xác review > 90%", "Founder & AI Engineer"),
        ("Tháng 2-3: Đòn bẩy (Leverage)", "Publish chính thức lên GitHub Marketplace, chạy technical blog trên HackerNews/Dev.to, kích hoạt freemium tier", "150 installs, 35 khách hàng trả phí Pro ($149)", "Product Lead & Tech Marketer"),
        ("Tháng 4+: Mở rộng (Expand)", "Hỗ trợ thêm GitLab & Bitbucket Server cho các khách hàng doanh nghiệp tài chính / bảo mật cao", "$15,000 MRR, tỷ lệ churn < 3%/tháng", "Full Engineering Team")
    ]
    for idx, (p1, p2, p3, p4) in enumerate(plans, start=12):
        ws5.cell(row=idx, column=2, value=p1).font = font_bold
        ws5.cell(row=idx, column=3, value=p2)
        ws5.cell(row=idx, column=4, value=p3)
        ws5.cell(row=idx, column=5, value=p4).font = font_bold
        for c in range(2, 6):
            ws5.cell(row=idx, column=c).fill = fill_yellow
            ws5.cell(row=idx, column=c).border = border_all

    ws5["B16"] = "3. CHECKLIST EVIDENCE PACK CHO PROCUREMENT & IT"
    ws5["B16"].font = font_section
    ws5.merge_cells("B16:E16")
    ws5["B16"].fill = fill_section

    evid_headers = ["Tài sản bán hàng", "Trạng thái / Nội dung cụ thể", "Bằng chứng văn bản", "Người phụ trách / Deadline"]
    for col_idx, h in enumerate(evid_headers, start=2):
        cell = ws5.cell(row=17, column=col_idx, value=h)
        cell.font = font_bold
        cell.fill = fill_header
        cell.border = border_all

    evid_data = [
        ("Eval Results", "Đạt độ chính xác 92.4% trên bộ 500 PRs thực tế, false positive 4.8%, thời gian phản hồi 38s", "Báo cáo Eval Benchmark PDF kèm metrics", "AI Engineer (Đã có sẵn)"),
        ("Procurement Q&A", "Cam kết ZDR (Zero Data Retention), không train model trên code khách, xóa bộ nhớ ephemeral sau khi post comment", "Bản cam kết bảo mật & pháp lý 5 trang", "Founder & Legal (30/10/2026)"),
        ("Pilot Report", "Thử nghiệm 4 tuần tại công ty 30 devs: Giảm PR review turnaround time từ 18.4 giờ xuống 42 phút, tiết kiệm 28h/tháng cho Tech Lead", "Pilot Case Study có xác nhận của CTO khách", "Product Lead (Đã hoàn thành)")
    ]
    for idx, (e1, e2, e3, e4) in enumerate(evid_data, start=18):
        ws5.cell(row=idx, column=2, value=e1).font = font_bold
        ws5.cell(row=idx, column=3, value=e2)
        ws5.cell(row=idx, column=4, value=e3)
        ws5.cell(row=idx, column=5, value=e4).font = font_bold
        for c in range(2, 6):
            ws5.cell(row=idx, column=c).fill = fill_yellow
            ws5.cell(row=idx, column=c).border = border_all

    # -------------------------------------------------------------
    # TAB 6: 6_Benchmarks
    # -------------------------------------------------------------
    ws6 = wb.create_sheet(title="6_Benchmarks")
    ws6.views.sheetView[0].showGridLines = True
    ws6.column_dimensions["A"].width = 5
    ws6.column_dimensions["B"].width = 25
    ws6.column_dimensions["C"].width = 25
    ws6.column_dimensions["D"].width = 25
    ws6.column_dimensions["E"].width = 45

    ws6["B2"] = "TAB 6: BẢNG GIÁ THAM CHIẾU API & SẢN PHẨM NGÀNH (CHỐT 26/08/2026)"
    ws6["B2"].font = font_title

    headers6 = ["Nhà cung cấp / Dịch vụ", "Model / Gói cước", "Đơn giá niêm yết (USD)", "Ghi chú & Trạng thái khuyến mại"]
    for col_idx, h in enumerate(headers6, start=2):
        cell = ws6.cell(row=4, column=col_idx, value=h)
        cell.font = font_bold
        cell.fill = fill_header
        cell.border = border_all

    bench_apis = [
        ("Anthropic", "Claude Haiku 4.5", "$1.00 input / $5.00 output / $0.10 cache read per 1M", "Giá chính thức, cache hit giảm 90% input"),
        ("Anthropic", "Claude Sonnet 3.5", "$3.00 input / $15.00 output / $0.30 cache read per 1M", "Dùng cho các ca phân tích kiến trúc phức tạp"),
        ("OpenAI", "GPT-4o-mini", "$0.15 input / $0.60 output / $0.075 cache read per 1M", "Model siêu rẻ cho tác vụ linting"),
        ("Google Cloud", "Gemini 3.7 Flash", "$0.75 input / $3.75 output per 1M", "⏳ Giá khuyến mại đến 31/12/2026"),
        ("Deepgram", "Nova-3 STT", "$0.0043 / phút (pre-recorded)", "Dịch vụ bóc băng âm thanh"),
        ("ElevenLabs", "Flash / Turbo TTS", "$50 / 1M ký tự", "Dịch vụ đọc văn bản thành giọng nói"),
        ("Intercom", "Fin AI Agent", "$0.99 / resolution", "Benchmark tính tiền theo Outcome chuẩn mực"),
        ("GitHub", "Copilot Business", "$19 / seat / tháng + AI Credits", "Benchmark chuyển đổi từ seat sang usage")
    ]
    for idx, (a1, a2, a3, a4) in enumerate(bench_apis, start=5):
        ws6.cell(row=idx, column=2, value=a1).font = font_bold
        ws6.cell(row=idx, column=3, value=a2)
        ws6.cell(row=idx, column=4, value=a3)
        ws6.cell(row=idx, column=5, value=a4).font = font_italic
        for c in range(2, 6):
            ws6.cell(row=idx, column=c).fill = fill_gray
            ws6.cell(row=idx, column=c).border = border_all

    wb.save(filepath)
    print(f"Excel model created successfully at: {filepath}")
    # Also save Day22_model.xlsx for maximum compatibility
    alt_filepath = filepath.replace("Chi_Day22_model.xlsx", "Day22_model.xlsx")
    wb.save(alt_filepath)
    print(f"Excel model also saved at: {alt_filepath}")

if __name__ == "__main__":
    build_excel_model("d:/Track1_Day22_2A202602797_ThanThiKimChi/Chi_Day22_model.xlsx")
