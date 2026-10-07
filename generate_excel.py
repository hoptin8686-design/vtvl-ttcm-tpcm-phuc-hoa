# -*- coding: utf-8 -*-
"""
Script sinh file Excel: Đề án Vị trí việc làm Viên chức – THPT Phục Hòa
Căn cứ Nghị định số 232/2026/NĐ-CP ngày 26/6/2026 của Chính phủ
Áp dụng ĐÚNG GIỐNG MẪU BẢNG MÔ TẢ VỊ TRÍ VIỆC LÀM TRÊN MÀN HÌNH MÁY TÍNH
ĐẦY ĐỦ TRỌN BỘ 35 BIÊN CHẾ (03 NHÓM DANH MỤC):
1. Vị trí việc làm Quản lý (07 người): Hiệu trưởng, Phó Hiệu trưởng, TTCM, TPCM.
2. Vị trí việc làm Chuyên môn nghiệp vụ theo số các môn học trong phụ lục (22 người - 13 môn học):
   Ngữ văn (3), Toán (3), Ngoại ngữ (2), Lịch sử (2), GDTC (2), GDQP-AN (1), Địa lí (2),
   GD KT&PL (1), Vật lí (1), Hóa học (1), Sinh học (1), Công nghệ (2), Tin học (1).
3. Vị trí việc làm Hỗ trợ 06 người: Kế toán, Văn thư, Thiết bị TN, Giáo vụ, Y tế học đường, Thủ quỹ.
Kèm Khung năng lực chuẩn 5 cấp độ và KPI đo lường định lượng.
"""

import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
import os
import sys

if sys.stdout:
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass

def create_full_vtvl_excel(output_paths):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)  # Remove default sheet

    # Colors
    NAVY_FILL = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
    BLUE_HEADER = PatternFill(start_color="2563EB", end_color="2563EB", fill_type="solid")
    TEAL_HEADER = PatternFill(start_color="0D9488", end_color="0D9488", fill_type="solid")
    PURPLE_HEADER = PatternFill(start_color="6D28D9", end_color="6D28D9", fill_type="solid")
    AMBER_HEADER = PatternFill(start_color="D97706", end_color="D97706", fill_type="solid")
    EMERALD_HEADER = PatternFill(start_color="059669", end_color="059669", fill_type="solid")

    SECTION_FILL = PatternFill(start_color="DBEAFE", end_color="DBEAFE", fill_type="solid")
    SECTION_FILL_GREEN = PatternFill(start_color="CCFBF1", end_color="CCFBF1", fill_type="solid")
    SECTION_FILL_AMBER = PatternFill(start_color="FEF3C7", end_color="FEF3C7", fill_type="solid")
    ZEBRA_FILL = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")

    FONT_FAMILY = "Times New Roman"

    FONT_TITLE_MAIN = Font(name=FONT_FAMILY, size=13, bold=True, color="1E3A8A")
    FONT_HEADER_WHITE = Font(name=FONT_FAMILY, size=11, bold=True, color="FFFFFF")
    FONT_SECTION = Font(name=FONT_FAMILY, size=11, bold=True, color="1E3A8A")
    FONT_BOLD = Font(name=FONT_FAMILY, size=11, bold=True, color="000000")
    FONT_REGULAR = Font(name=FONT_FAMILY, size=11, bold=False, color="000000")
    FONT_ITALIC = Font(name=FONT_FAMILY, size=10, italic=True, color="4B5563")

    THIN_BORDER_SIDE = Side(border_style="thin", color="CBD5E1")
    MEDIUM_BORDER_SIDE = Side(border_style="medium", color="64748B")
    BORDER_CELL = Border(left=THIN_BORDER_SIDE, right=THIN_BORDER_SIDE, top=THIN_BORDER_SIDE, bottom=THIN_BORDER_SIDE)
    BORDER_HEADER = Border(left=THIN_BORDER_SIDE, right=THIN_BORDER_SIDE, top=MEDIUM_BORDER_SIDE, bottom=MEDIUM_BORDER_SIDE)

    ALIGN_CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ALIGN_LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)

    def apply_row_styles(ws, row_idx, font=FONT_REGULAR, alignment=ALIGN_LEFT, fill=None, border=BORDER_CELL):
        for col_idx in range(1, ws.max_column + 1):
            cell = ws.cell(row=row_idx, column=col_idx)
            if font: cell.font = font
            if alignment: cell.alignment = alignment
            if fill: cell.fill = fill
            if border: cell.border = border

    def auto_fit_columns(ws, max_cols=None, max_width_limit=68):
        ws.views.sheetView[0].showGridLines = True
        cols = max_cols if max_cols else ws.max_column
        for col in range(1, cols + 1):
            col_letter = get_column_letter(col)
            max_len = 0
            for row in range(1, ws.max_row + 1):
                val = ws.cell(row=row, column=col).value
                if val is not None:
                    lines = str(val).split("\n")
                    line_max = max(len(l) for l in lines) if lines else 0
                    if line_max > max_len and line_max < 120:
                        max_len = line_max
            adjusted_width = min(max(max_len + 3, 11), max_width_limit)
            ws.column_dimensions[col_letter].width = adjusted_width

    # ==============================================================================
    # SHEET 1: ĐỀ ÁN & CĂN CỨ NĐ 232
    # ==============================================================================
    ws1 = wb.create_sheet(title="1. Đề án & Căn cứ NĐ 232")
    ws1.column_dimensions['A'].width = 8
    ws1.column_dimensions['B'].width = 28
    ws1.column_dimensions['C'].width = 30
    ws1.column_dimensions['D'].width = 18
    ws1.column_dimensions['E'].width = 18
    ws1.column_dimensions['F'].width = 40

    ws1.merge_cells("A1:C1")
    ws1["A1"] = "SỞ GIÁO DỤC VÀ ĐÀO TẠO CAO BẰNG"
    ws1["A1"].font = Font(name=FONT_FAMILY, size=11, bold=True)
    ws1["A1"].alignment = ALIGN_CENTER

    ws1.merge_cells("A2:C2")
    ws1["A2"] = "TRƯỜNG THPT PHỤC HÒA"
    ws1["A2"].font = Font(name=FONT_FAMILY, size=11, bold=True, color="1E3A8A")
    ws1["A2"].alignment = ALIGN_CENTER

    ws1.merge_cells("D1:F1")
    ws1["D1"] = "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM"
    ws1["D1"].font = Font(name=FONT_FAMILY, size=11, bold=True)
    ws1["D1"].alignment = ALIGN_CENTER

    ws1.merge_cells("D2:F2")
    ws1["D2"] = "Độc lập - Tự do - Hạnh phúc"
    ws1["D2"].font = Font(name=FONT_FAMILY, size=11, bold=True, italic=True)
    ws1["D2"].alignment = ALIGN_CENTER

    ws1.merge_cells("A4:F4")
    ws1["A4"] = "ĐỀ ÁN VỊ TRÍ VIỆC LÀM VIÊN CHỨC NĂM HỌC 2026 - 2027"
    ws1["A4"].font = FONT_TITLE_MAIN
    ws1["A4"].alignment = ALIGN_CENTER

    ws1.merge_cells("A5:F5")
    ws1["A5"] = "(Xây dựng theo Nghị định 232/2026/NĐ-CP & Phụ lục I C3 Phục Hòa - Tổng 35 biên chế)"
    ws1["A5"].font = FONT_ITALIC
    ws1["A5"].alignment = ALIGN_CENTER

    legal_rows = [
        ("I", "CĂN CỨ PHÁP LÝ XÂY DỰNG ĐỀ ÁN", "", "", "", ""),
        ("1", "Nghị định số 232/2026/NĐ-CP", "Chính phủ ban hành 26/6/2026 quy định về vị trí việc làm viên chức", "Chính phủ", "26/06/2026", "Hiệu lực từ 01/7/2026"),
        ("2", "Luật Viên chức số 129/2025/QH15", "Quy định quyền, nghĩa vụ của viên chức, tuyển dụng, sử dụng và quản lý viên chức", "Quốc hội", "2025", "Văn bản nền tảng"),
        ("3", "Thông tư số 20/2023/TT-BGDĐT", "Hướng dẫn về vị trí việc làm, cơ cấu viên chức theo chức danh nghề nghiệp và định mức người làm việc trong các cơ sở GDPT", "Bộ GD&ĐT", "30/10/2023", "Quy định định mức GV/lớp"),
        ("4", "Phụ lục I VTVL Trường THPT Phục Hòa", "Danh mục vị trí việc làm và bậc nghề nghiệp được sử dụng trong từng vị trí của Trường THPT Phục Hòa: 35 biên chế", "THPT Phục Hòa", "10/2026", "Phê duyệt của Sở GD&ĐT"),
        ("II", "NGUYÊN TẮC XÂY DỰNG & QUY MÔ TRƯỜNG LỚP", "", "", "", ""),
        ("1", "Quy mô trường lớp", "Tổng số lớp: 18 lớp THPT (Khối 10: 6 lớp, Khối 11: 6 lớp, Khối 12: 6 lớp). Tổng số học sinh: ~720 học sinh.", "", "", ""),
        ("2", "Tổng số biên chế theo Đề án", "Tổng số biên chế người làm việc: 35 người (đúng theo Phụ lục I Danh mục VTVL của nhà trường).", "", "", ""),
        ("3", "Cơ cấu 3 Nhóm vị trí", "- Nhóm I (Viên chức quản lý): 07 người (Hiệu trưởng 1, Phó HT 2, TTCM 2, TPCM 2)\n- Nhóm II (Viên chức chuyên môn, nghiệp vụ): 22 người (13 môn học trong phụ lục)\n- Nhóm III (Viên chức hỗ trợ): 06 người (Kế toán 1, Văn thư 1, TB-TN 1, Giáo vụ 1, Y tế 1, Thủ quỹ 1)", "", "", ""),
        ("4", "Chuẩn hóa theo Mẫu NĐ 232", "Toàn bộ bản mô tả VTVL được lập chuẩn theo 6 cột nghiệp vụ và Phần B Phụ lục IV/VI của Nghị định số 232/2026/NĐ-CP.", "", "", "")
    ]

    cur_row = 7
    for item in legal_rows:
        if item[0] in ["I", "II"]:
            ws1.merge_cells(start_row=cur_row, start_column=1, end_row=cur_row, end_column=6)
            ws1.cell(cur_row, 1, f"{item[0]}. {item[1]}")
            apply_row_styles(ws1, cur_row, font=FONT_SECTION, alignment=ALIGN_LEFT, fill=SECTION_FILL)
        else:
            for c_idx, val in enumerate(item, start=1):
                ws1.cell(cur_row, c_idx, val)
            apply_row_styles(ws1, cur_row, font=FONT_REGULAR, alignment=ALIGN_LEFT)
            ws1.cell(cur_row, 1).alignment = ALIGN_CENTER
            ws1.cell(cur_row, 4).alignment = ALIGN_CENTER
            ws1.cell(cur_row, 5).alignment = ALIGN_CENTER
        cur_row += 1

    # ==============================================================================
    # SHEET 2: PHỤ LỤC I - DANH MỤC 35 VTVL TỔNG HỢP
    # ==============================================================================
    ws2 = wb.create_sheet(title="2. Phụ lục I - DM 35 VTVL")
    ws2.column_dimensions['A'].width = 8
    ws2.column_dimensions['B'].width = 38
    ws2.column_dimensions['C'].width = 20
    ws2.column_dimensions['D'].width = 24
    ws2.column_dimensions['E'].width = 14
    ws2.column_dimensions['F'].width = 22
    ws2.column_dimensions['G'].width = 40

    ws2.merge_cells("A1:G1")
    ws2["A1"] = "DANH MỤC VỊ TRÍ VIỆC LÀM VÀ BẬC NGHỀ NGHIỆP TRƯỜNG THPT PHỤC HÒA"
    ws2["A1"].font = FONT_TITLE_MAIN
    ws2["A1"].alignment = ALIGN_CENTER

    ws2.merge_cells("A2:G2")
    ws2["A2"] = "(Phụ lục I Danh mục VTVL của Trường THPT Phục Hòa - Căn cứ Nghị định số 232/2026/NĐ-CP)"
    ws2["A2"].font = FONT_ITALIC
    ws2["A2"].alignment = ALIGN_CENTER

    headers_ws2 = ["STT", "Tên vị trí việc làm", "Mã vị trí việc làm", "Bậc nghề nghiệp sử dụng", "Số biên chế", "Phụ cấp / Định mức", "Ghi chú vai trò cốt lõi"]
    ws2.append([])
    ws2.append(headers_ws2)
    for col_idx in range(1, len(headers_ws2) + 1):
        cell = ws2.cell(row=4, column=col_idx)
        cell.font = FONT_HEADER_WHITE
        cell.fill = NAVY_FILL
        cell.alignment = ALIGN_CENTER
        cell.border = BORDER_HEADER

    pl1_data = [
        ("I", "Vị trí việc làm viên chức quản lý: 07 người", "", "", "7", "", "Lãnh đạo, quản lý và điều hành trường"),
        ("1", "Hiệu trưởng", "HT-THPT-01", "Bậc 3 đến Bậc 5", "1", "PC: 0.70 | Dạy 2t/t", "Người đứng đầu, quản lý điều hành toàn diện"),
        ("2", "Phó Hiệu trưởng", "PHT-THPT-01", "Bậc 3 đến Bậc 5", "2", "PC: 0.50 | Dạy 4t/t", "Giúp Hiệu trưởng phụ trách chuyên môn, CSVC"),
        ("3", "Tổ trưởng chuyên môn", "TTCM-THPT-01", "Bậc 3 đến Bậc 4", "2", "PC: 0.25 | Dạy 14t/t", "Quản lý sinh hoạt CM tổ, kế hoạch dạy học"),
        ("4", "Tổ phó chuyên môn", "TPCM-THPT-01", "Bậc 3 đến Bậc 4", "2", "PC: 0.15 | Dạy 16t/t", "Giúp TTCM kiểm tra hồ sơ, theo dõi tiến độ"),

        ("II", "Vị trí việc làm viên chức chuyên môn, nghiệp vụ: 22 người (13 môn)", "", "", "22", "", "Giảng dạy 13 môn học theo CT GDPT 2018"),
        ("1", "Giáo viên THPT môn Ngữ văn", "GV-THPT-01", "Bậc 3 đến Bậc 4", "3", "17 tiết/tuần", "Dạy Ngữ văn 10, 11, 12, bồi dưỡng HSG Văn"),
        ("2", "Giáo viên THPT môn Toán", "GV-THPT-01", "Bậc 3 đến Bậc 4", "3", "17 tiết/tuần", "Dạy Toán 10, 11, 12, bồi dưỡng HSG Toán"),
        ("3", "Giáo viên THPT môn Ngoại ngữ (Tiếng Anh)", "GV-THPT-01", "Bậc 3 đến Bậc 4", "2", "17 tiết/tuần", "Dạy Ngoại ngữ, phòng lab, giao tiếp"),
        ("4", "Giáo viên THPT môn Lịch sử", "GV-THPT-01", "Bậc 3 đến Bậc 4", "2", "17 tiết/tuần", "Dạy Lịch sử bắt buộc và lựa chọn"),
        ("5", "Giáo viên THPT môn GDTC", "GV-THPT-01", "Bậc 3 đến Bậc 4", "2", "17 tiết/tuần", "Dạy GD Thể chất, Hội khỏe Phù Đổng"),
        ("6", "Giáo viên THPT môn GDQP-AN", "GV-THPT-01", "Bậc 3 đến Bậc 4", "1", "17 tiết/tuần", "Dạy GDQP-AN, hội thao, chủ quyền biên giới"),
        ("7", "Giáo viên THPT môn Địa lí", "GV-THPT-01", "Bậc 3 đến Bậc 4", "2", "17 tiết/tuần", "Dạy Địa lí, Atlat, GIS, tài nguyên môi trường"),
        ("8", "Giáo viên THPT môn GD KT&PL", "GV-THPT-01", "Bậc 3 đến Bậc 4", "1", "17 tiết/tuần", "Dạy GD Kinh tế & Pháp luật, xử lý tình huống"),
        ("9", "Giáo viên THPT môn Vật lí", "GV-THPT-01", "Bậc 3 đến Bậc 4", "1", "17 tiết/tuần", "Dạy Vật lí, phòng bộ môn Vật lí, STEM"),
        ("10", "Giáo viên THPT môn Hóa học", "GV-THPT-01", "Bậc 3 đến Bậc 4", "1", "17 tiết/tuần", "Dạy Hóa học, thực hành phòng TN, PCCC"),
        ("11", "Giáo viên THPT môn Sinh học", "GV-THPT-01", "Bậc 3 đến Bậc 4", "1", "17 tiết/tuần", "Dạy Sinh học, kính hiển vi, vườn thực nghiệm"),
        ("12", "Giáo viên THPT môn Công nghệ", "GV-THPT-01", "Bậc 3 đến Bậc 4", "2", "17 tiết/tuần", "Dạy Công nghệ công nghiệp & Nông nghiệp"),
        ("13", "Giáo viên THPT môn Tin học", "GV-THPT-01", "Bậc 3 đến Bậc 4", "1", "17 tiết/tuần", "Dạy Tin học, lập trình Python, phòng máy, AI"),

        ("III", "Vị trí việc làm viên chức hỗ trợ: 06 người", "", "", "6", "", "Bảo đảm điều kiện hoạt động của trường"),
        ("1", "Kế toán", "KT-THPT-01", "Bậc 1 đến Bậc 4", "1", "40 giờ/tuần", "Quản lý ngân sách, tài chính, thanh toán lương"),
        ("2", "Văn thư", "VT-THPT-01", "Bậc 1 đến Bậc 4", "1", "40 giờ/tuần", "Tiếp nhận, phát hành công văn, quản lý con dấu"),
        ("3", "Thiết bị, thí nghiệm", "TBTN-THPT-01", "Bậc 2 đến Bậc 3", "1", "40 giờ/tuần", "Quản lý thiết bị dạy học, phòng thí nghiệm TN"),
        ("4", "Giáo vụ", "GVU-THPT-01", "Bậc 2 đến Bậc 3", "1", "40 giờ/tuần", "Quản lý hồ sơ học sinh, sổ đăng bộ, tuyển sinh"),
        ("5", "Y tế trường học", "YT-THPT-01", "Bậc 1 đến Bậc 3", "1", "40 giờ/tuần", "Chăm sóc sức khỏe, sơ cấp cứu, phòng chống dịch"),
        ("6", "Thủ quỹ", "TQ-THPT-01", "Bậc 1 đến Bậc 3", "1", "40 giờ/tuần", "Quản lý quỹ tiền mặt, thu chi đúng phiếu chi/thu"),

        ("TỔNG CỘNG", "Toàn bộ 03 nhóm VTVL trường THPT Phục Hòa", "35 VTVL", "Bậc 1 đến Bậc 5", "35", "Đủ 35 biên chế", "Đúng định mức được giao")
    ]

    cur_row = 5
    for item in pl1_data:
        for c_idx, val in enumerate(item, start=1):
            ws2.cell(cur_row, c_idx, val)

        if item[0] in ["I", "II", "III"]:
            apply_row_styles(ws2, cur_row, font=FONT_SECTION, alignment=ALIGN_LEFT, fill=SECTION_FILL)
            ws2.cell(cur_row, 1).alignment = ALIGN_CENTER
            ws2.cell(cur_row, 5).alignment = ALIGN_CENTER
        elif item[0] == "TỔNG CỘNG":
            apply_row_styles(ws2, cur_row, font=FONT_TITLE_MAIN, alignment=ALIGN_LEFT, fill=SECTION_FILL_AMBER)
            ws2.cell(cur_row, 1).alignment = ALIGN_CENTER
            ws2.cell(cur_row, 5).alignment = ALIGN_CENTER
        else:
            apply_row_styles(ws2, cur_row, font=FONT_REGULAR, alignment=ALIGN_LEFT)
            ws2.cell(cur_row, 1).alignment = ALIGN_CENTER
            ws2.cell(cur_row, 3).alignment = ALIGN_CENTER
            ws2.cell(cur_row, 4).alignment = ALIGN_CENTER
            ws2.cell(cur_row, 5).alignment = ALIGN_CENTER
        cur_row += 1

    # ==============================================================================
    # SHEET 3: DANH MỤC 1: BẢNG MÔ TẢ VTVL QUẢN LÝ (07 BIÊN CHẾ)
    # ==============================================================================
    ws3 = wb.create_sheet(title="3. DM1 - VTVL Quản lý")
    ws3.column_dimensions['A'].width = 8
    ws3.column_dimensions['B'].width = 28
    ws3.column_dimensions['C'].width = 30
    ws3.column_dimensions['D'].width = 46
    ws3.column_dimensions['E'].width = 12
    ws3.column_dimensions['F'].width = 40

    ws3.merge_cells("A1:F1")
    ws3["A1"] = "DANH MỤC 1: BẢN MÔ TẢ VỊ TRÍ VIỆC LÀM VIÊN CHỨC QUẢN LÝ (07 NGƯỜI)"
    ws3["A1"].font = FONT_TITLE_MAIN
    ws3["A1"].alignment = ALIGN_CENTER

    headers_6col = ["STT", "Vị trí việc làm & Mã số", "Nhiệm vụ quản lý chính", "Nội dung công việc cụ thể & Tiêu chí đo lường", "Tỷ trọng", "Sản phẩm / Kết quả đầu ra"]
    ws3.append([])
    ws3.append(headers_6col)
    for c_i in range(1, len(headers_6col) + 1):
        cell = ws3.cell(row=3, column=c_i)
        cell.font = FONT_HEADER_WHITE
        cell.fill = BLUE_HEADER
        cell.alignment = ALIGN_CENTER
        cell.border = BORDER_HEADER

    dm1_rows = [
        ("1", "Hiệu trưởng\n(HT-THPT-01)\nPhụ cấp: 0.70 | 01 người\nĐịnh mức: 02 tiết/tuần",
         "1. Chiến lược & Kế hoạch phát triển",
         "- Xây dựng Chiến lược 5 năm, Kế hoạch năm học trình cấp có thẩm quyền phê duyệt.\n- Ban hành Quy chế dân chủ, Quy chế chi tiêu nội bộ, Quy chế làm việc của trường.",
         "15%",
         "- Chiến lược phát triển trường 5 năm.\n- Kế hoạch năm học & Quy chế làm việc.\n- Báo cáo định kỳ Sở GD&ĐT."),
        ("2", "Hiệu trưởng\n(HT-THPT-01)",
         "2. Quản lý nhân sự & Đánh giá thi đua",
         "- Phân công nhiệm vụ, bổ nhiệm tổ trưởng, tổ phó; đánh giá xếp loại VC hàng năm theo NĐ 232/2026/NĐ-CP.\n- Thực hiện công tác quy hoạch, đào tạo, bồi dưỡng và thi đua khen thưởng.",
         "20%",
         "- QĐ phân công nhiệm vụ cán bộ giáo viên.\n- QĐ xếp loại thi đua, khen thưởng viên chức.\n- Hồ sơ quy hoạch cán bộ quản lý."),
        ("3", "Hiệu trưởng\n(HT-THPT-01)",
         "3. Quản lý tài chính & Cơ sở vật chất",
         "- Chủ tài khoản, duyệt thu - chi ngân sách, kiểm tra công tác kế toán, quản lý tài sản công đúng luật.\n- Chỉ đạo sửa chữa, nâng cấp CSVC trường học bảo đảm dạy và học an toàn.",
         "15%",
         "- Dự toán & Báo cáo quyết toán ngân sách.\n- Quy chế chi tiêu nội bộ chuẩn hóa.\n- Báo cáo kiểm kê tài sản hàng năm."),
        ("4", "Hiệu trưởng\n(HT-THPT-01)",
         "4. Chỉ đạo chuyên môn & Giảng dạy",
         "- Chỉ đạo kỳ thi tốt nghiệp THPT, kỳ thi chọn HSG; trực tiếp giảng dạy 02 tiết/tuần theo quy định.\n- Thanh tra nội bộ, dự giờ thăm lớp, đổi mới PPDH và ứng dụng AI toàn trường.",
         "20%",
         "- Báo cáo kết quả kỳ thi tốt nghiệp THPT.\n- Sổ giáo án cá nhân giảng dạy 02 tiết/tuần.\n- Biên bản kiểm tra chuyên môn nội bộ."),
        ("5", "Phó Hiệu trưởng\n(PHT-THPT-01)\nPhụ cấp: 0.50 | 02 người\nĐịnh mức: 04 tiết/tuần",
         "1. Chỉ đạo & điều hành chuyên môn",
         "- Thẩm định và duyệt Kế hoạch giáo dục nhà trường, kế hoạch các tổ CM theo CT GDPT 2018.\n- Chỉ đạo đổi mới phương pháp dạy học, giáo dục STEM/STEAM, chuyển đổi số dạy học.",
         "25%",
         "- Kế hoạch chuyên môn năm học được phê duyệt.\n- Lịch thi, thời khóa biểu toàn trường.\n- Báo cáo sơ kết, tổng kết chuyên môn."),
        ("6", "Phó Hiệu trưởng\n(PHT-THPT-01)",
         "2. Khảo thí, ôn thi TN & Ký số học bạ",
         "- Xây dựng ma trận đặc tả, ngân hàng đề kiểm tra định kỳ; quản lý ôn thi TN THPT khối 12.\n- Kiểm tra, ký duyệt học bạ điện tử, kiểm tra hồ sơ giáo án giáo viên; dạy 04 tiết/tuần.",
         "20%",
         "- Ngân hàng đề kiểm tra định kỳ chuẩn quy chế.\n- Kế hoạch ôn thi TN THPT chi tiết theo môn.\n- 100% học bạ số được duyệt đúng hạn."),
        ("7", "Tổ trưởng CM\n(TTCM-THPT-01)\nPhụ cấp: 0.25 | 02 người\nĐịnh mức: 14 tiết/tuần",
         "1. Quản lý toàn diện chuyên môn tổ",
         "- Xây dựng kế hoạch dạy học môn học (PL1, PL2, PL3) theo CT GDPT 2018; phân công GV dạy đúng chuyên ngành.\n- Tổ chức sinh hoạt chuyên môn theo NCBL (≥ 2 lần/tháng); kiểm tra giáo án, dự giờ GV.",
         "40%",
         "- Kế hoạch giáo dục tổ chuyên môn.\n- Biên bản họp sinh hoạt chuyên môn theo NCBL.\n- Phiếu đánh giá giáo viên trong tổ."),
        ("8", "Tổ phó CM\n(TPCM-THPT-01)\nPhụ cấp: 0.15 | 02 người\nĐịnh mức: 16 tiết/tuần",
         "1. Quản lý nề nếp, sổ sách & tiến độ",
         "- Theo dõi tiến độ chương trình hàng tuần, lịch báo giảng điện tử, kế hoạch dạy thay/bù của tổ viên.\n- Kiểm tra nề nếp hồ sơ sổ sách, đôn đốc ký số học bạ; giảng dạy 16 tiết/tuần.",
         "40%",
         "- Sổ theo dõi tiến độ chương trình hàng tuần.\n- Biên bản kiểm tra nề nếp hồ sơ chuyên môn.\n- Báo cáo tổng hợp nề nếp tổ viên.")
    ]

    r_idx = 4
    for it in dm1_rows:
        for c_i, v in enumerate(it, start=1):
            ws3.cell(r_idx, c_i, v)
        apply_row_styles(ws3, r_idx, font=FONT_REGULAR)
        ws3.cell(r_idx, 1).alignment = ALIGN_CENTER
        ws3.cell(r_idx, 5).alignment = ALIGN_CENTER
        if "Hiệu trưởng\n(" in it[1] or "Phó Hiệu trưởng\n(" in it[1] or "Tổ trưởng" in it[1] or "Tổ phó" in it[1]:
            ws3.cell(r_idx, 2).font = FONT_BOLD
        r_idx += 1

    # ==============================================================================
    # SHEET 4: DANH MỤC 2: BẢNG MÔ TẢ VTVL 13 MÔN HỌC (22 BIÊN CHẾ)
    # ==============================================================================
    ws4 = wb.create_sheet(title="4. DM2 - VTVL 13 Môn học")
    ws4.column_dimensions['A'].width = 8
    ws4.column_dimensions['B'].width = 30
    ws4.column_dimensions['C'].width = 30
    ws4.column_dimensions['D'].width = 46
    ws4.column_dimensions['E'].width = 12
    ws4.column_dimensions['F'].width = 42

    ws4.merge_cells("A1:F1")
    ws4["A1"] = "DANH MỤC 2: BẢNG MÔ TẢ VỊ TRÍ VIỆC LÀM VIÊN CHỨC CHUYÊN MÔN THEO 13 MÔN HỌC (22 BIÊN CHẾ)"
    ws4["A1"].font = FONT_TITLE_MAIN
    ws4["A1"].alignment = ALIGN_CENTER

    ws4.append([])
    ws4.append(headers_6col)
    for c_i in range(1, len(headers_6col) + 1):
        cell = ws4.cell(row=3, column=c_i)
        cell.font = FONT_HEADER_WHITE
        cell.fill = TEAL_HEADER
        cell.alignment = ALIGN_CENTER
        cell.border = BORDER_HEADER

    # Data for 13 subjects, each with full 4 detailed duties and specific weights & products
    subjects_specs = [
        {
            "name": "Giáo viên môn Ngữ văn", "code": "GV-THPT-01", "grade": "Bậc 3 đến Bậc 4", "quota": "03 biên chế",
            "d1_name": "1. Giảng dạy môn Ngữ văn & Đổi mới KTĐG",
            "d1_desc": "- Giảng dạy môn Ngữ văn 10, 11, 12 theo CT GDPT 2018 (định mức 17 tiết/tuần).\n- Rèn luyện 4 kỹ năng Đọc - Viết - Nói - Nghe; đổi mới PPDH tác phẩm văn học, tích hợp ngữ liệu địa phương Cao Bằng.\n- Đổi mới ra đề kiểm tra đánh giá theo hướng phát triển năng lực, tránh sao chép văn mẫu.",
            "d1_w": "40%",
            "d1_out": "- Kế hoạch bài dạy (giáo án) môn Văn chuẩn bị chu đáo.\n- Đề kiểm tra định kỳ có ma trận đặc tả.\n- Học bạ và sổ điểm điện tử cập nhật đúng hạn.\n- Tỷ lệ học sinh đạt chuẩn môn Văn ≥ 98%.",
            "d2_name": "2. Công tác chủ nhiệm & Giáo dục đạo đức",
            "d2_desc": "- Quản lý toàn diện học sinh lớp chủ nhiệm, xây dựng tập thể lớp đoàn kết, kỷ cương.\n- Phối hợp chặt chẽ với cha mẹ học sinh qua sổ liên lạc điện tử; ký duyệt học bạ số.",
            "d2_w": "25%",
            "d2_out": "- Kế hoạch chủ nhiệm năm học.\n- Biên bản họp CMHS đầu năm, cuối kỳ.\n- Sổ theo dõi nề nếp học sinh.\n- Học bạ số ký duyệt đầy đủ.",
            "d3_name": "3. Bồi dưỡng chuyên môn, NCBL & Trợ lý AI",
            "d3_desc": "- Tham gia sinh hoạt chuyên môn theo NCBL đầy đủ (≥ 2 lần/tháng); dự giờ đồng nghiệp.\n- Ứng dụng công nghệ thông tin và AI (trợ lý AI gợi ý ngữ liệu văn học, tóm tắt tác phẩm, đề đọc hiểu).",
            "d3_w": "20%",
            "d3_out": "- Sổ tích lũy chuyên môn cá nhân.\n- Phiếu dự giờ đồng nghiệp (≥ 1 tiết/kỳ).\n- Kho học liệu số môn Văn đóng góp cho tổ.",
            "d4_name": "4. Bồi dưỡng HSG môn Văn & Ôn thi tốt nghiệp",
            "d4_desc": "- Bồi dưỡng đội tuyển HSG môn Ngữ văn dự thi cấp tỉnh theo kế hoạch.\n- Phụ đạo học sinh yếu kém, phụ đạo ôn thi tốt nghiệp THPT khối 12 đạt phổ điểm cao.",
            "d4_w": "15%",
            "d4_out": "- Giáo án bồi dưỡng HSG và phụ đạo.\n- Học sinh đạt giải HSG môn Văn cấp tỉnh.\n- Tỷ lệ tốt nghiệp THPT môn Văn đạt ≥ 98.5%."
        },
        {
            "name": "Giáo viên môn Toán", "code": "GV-THPT-01", "grade": "Bậc 3 đến Bậc 4", "quota": "03 biên chế",
            "d1_name": "1. Giảng dạy môn Toán & Mô hình hóa thực tiễn",
            "d1_desc": "- Giảng dạy môn Toán 10, 11, 12 (Đại số, Giải tích, Hình học, Xác suất - Thống kê) theo định mức 17 tiết/tuần.\n- Phát triển tư duy logic, mô hình hóa toán học; ứng dụng phần mềm GeoGebra vẽ hình động, máy tính bỏ túi; dạy học STEM tích hợp Toán học.",
            "d1_w": "40%",
            "d1_out": "- Kế hoạch bài dạy môn Toán chất lượng cao.\n- Ngân hàng câu hỏi trắc nghiệm đúng/sai, trả lời ngắn theo cấu trúc mới.\n- Tỷ lệ HS đạt chuẩn môn Toán ≥ 96%.",
            "d2_name": "2. Công tác chủ nhiệm & Tư vấn hướng nghiệp KHTN",
            "d2_desc": "- Chủ nhiệm lớp, nắm bắt tâm lý học sinh, duy trì nề nếp kỷ luật lớp.\n- Tư vấn hướng nghiệp và định hướng lựa chọn tổ hợp môn KHTN / thi đại học.",
            "d2_w": "25%",
            "d2_out": "- Kế hoạch chủ nhiệm năm học.\n- Sổ theo dõi học sinh lớp chủ nhiệm.\n- Học bạ số được nhận xét và ký duyệt đúng hạn.",
            "d3_name": "3. Sinh hoạt NCBL & Ứng dụng AI tạo đề thi",
            "d3_desc": "- Sinh hoạt chuyên môn theo NCBL; tự làm thiết bị dạy học số.\n- Ứng dụng AI sinh đề trắc nghiệm theo 4 mức độ nhận thức của Bộ GD&ĐT.",
            "d3_w": "20%",
            "d3_out": "- Phiếu dự giờ đồng nghiệp.\n- Bộ bài giảng điện tử tương tác môn Toán.\n- Bài học STEM môn Toán được nghiệm thu.",
            "d4_name": "4. Bồi dưỡng HSG Toán & Ôn thi TN THPT",
            "d4_desc": "- Bồi dưỡng đội tuyển HSG Toán dự thi cấp tỉnh; phụ đạo học sinh có nguy cơ hổng kiến thức.\n- Trực tiếp giảng dạy ôn thi tốt nghiệp THPT môn Toán đạt phổ điểm khá giỏi.",
            "d4_w": "15%",
            "d4_out": "- Chuyên đề bồi dưỡng HSG Toán.\n- Học sinh đạt giải HSG Toán cấp tỉnh.\n- Phổ điểm tốt nghiệp môn Toán đạt bình quân ≥ 6.8."
        },
        {
            "name": "Giáo viên môn Ngoại ngữ (Tiếng Anh)", "code": "GV-THPT-01", "grade": "Bậc 3 đến Bậc 4", "quota": "02 biên chế",
            "d1_name": "1. Giảng dạy môn Ngoại ngữ & Phòng Lab tiếng",
            "d1_desc": "- Giảng dạy Tiếng Anh 10, 11, 12 theo khung năng lực 6 bậc Việt Nam (định mức 17 tiết/tuần).\n- Rèn luyện kỹ năng Nghe - Nói - Đọc - Viết; sử dụng phòng lab học tiếng, học liệu âm thanh đa phương tiện; tạo môi trường giao tiếp tiếng Anh tự nhiên.",
            "d1_w": "40%",
            "d1_out": "- Kế hoạch bài dạy Tiếng Anh tương tác.\n- File âm thanh bài kiểm tra kỹ năng nghe/nói.\n- Đề kiểm tra trắc nghiệm theo chuẩn cấu trúc mới.\n- Tỷ lệ HS đạt chuẩn môn Ngoại ngữ ≥ 95%.",
            "d2_name": "2. Chủ nhiệm lớp & Thúc đẩy CLB Tiếng Anh",
            "d2_desc": "- Công tác chủ nhiệm, quản lý nề nếp lớp; thúc đẩy học sinh tham gia Câu lạc bộ Tiếng Anh và hoạt động ngoại khóa giao lưu quốc tế.",
            "d2_w": "25%",
            "d2_out": "- Kế hoạch chủ nhiệm năm học.\n- Kế hoạch và sản phẩm sinh hoạt CLB Tiếng Anh.\n- Học bạ số ký duyệt đúng hạn.",
            "d3_name": "3. Tự bồi dưỡng chuẩn C1 & Ứng dụng AI luyện nói",
            "d3_desc": "- Nâng cao năng lực sư phạm tiếng Anh theo chuẩn Châu Âu (C1/B2); tham gia sinh hoạt chuyên môn cụm.\n- Ứng dụng AI luyện phát âm tiếng Anh và tạo bài tập giao tiếp.",
            "d3_w": "20%",
            "d3_out": "- Chứng chỉ bồi dưỡng chuyên môn.\n- Kho học liệu nghe nhìn tiếng Anh số hóa.\n- Phiếu dự giờ đồng nghiệp.",
            "d4_name": "4. Bồi dưỡng HSG Tiếng Anh & Ôn thi tốt nghiệp",
            "d4_desc": "- Bồi dưỡng HSG Tiếng Anh thi cấp tỉnh; phụ đạo học sinh có nguy cơ trượt tốt nghiệp.\n- Hướng dẫn học sinh tiếp cận chứng chỉ quốc tế (IELTS/VSTEP).",
            "d4_w": "15%",
            "d4_out": "- Đội tuyển HSG Tiếng Anh có giải cấp tỉnh.\n- Học sinh vùng cao tự tin vượt qua kỳ thi TN THPT môn Tiếng Anh."
        },
        {
            "name": "Giáo viên môn Lịch sử", "code": "GV-THPT-01", "grade": "Bậc 3 đến Bậc 4", "quota": "02 biên chế",
            "d1_name": "1. Giảng dạy môn Lịch sử & Giáo dục truyền thống",
            "d1_desc": "- Giảng dạy Lịch sử bắt buộc và các chuyên đề lựa chọn (17 tiết/tuần).\n- Giáo dục lòng yêu nước, tự hào dân tộc, gắn liền di tích lịch sử chiến thắng Biên giới và truyền thống cách mạng quê hương Cao Bằng.\n- Đổi mới PPDH, phát triển năng lực tìm hiểu và nhận thức lịch sử.",
            "d1_w": "40%",
            "d1_out": "- Kế hoạch bài dạy Lịch sử đổi mới phương pháp.\n- Sơ đồ tư duy dòng thời gian lịch sử số hóa.\n- Đề kiểm tra định kỳ có ma trận chuẩn.\n- Tỷ lệ HS đạt chuẩn môn Sử ≥ 98%.",
            "d2_name": "2. Công tác chủ nhiệm & Hoạt động Đoàn thanh niên",
            "d2_desc": "- Công tác chủ nhiệm lớp; giáo dục truyền thống cách mạng, đạo đức lối sống cho học sinh; phối hợp chặt chẽ với Đoàn thanh niên.",
            "d2_w": "25%",
            "d2_out": "- Kế hoạch chủ nhiệm năm học.\n- Báo cáo hoạt động giáo dục truyền thống lớp.\n- Học bạ số được phê duyệt đúng hạn.",
            "d3_name": "3. Sinh hoạt NCBL & Phim tài liệu lịch sử số",
            "d3_desc": "- Sinh hoạt chuyên môn theo NCBL; số hóa các thước phim tài liệu lịch sử.\n- Ứng dụng CNTT và AI xây dựng bài giảng lịch sử trực quan, sinh động.",
            "d3_w": "20%",
            "d3_out": "- Kho tư liệu phim/hình ảnh lịch sử số hóa.\n- Phiếu đánh giá dự giờ đồng nghiệp.\n- Bài giảng trình chiếu đa phương tiện.",
            "d4_name": "4. Bồi dưỡng HSG Lịch sử & Ôn thi tốt nghiệp KHXH",
            "d4_desc": "- Bồi dưỡng đội tuyển HSG Lịch sử dự thi cấp tỉnh; phụ đạo học sinh yếu.\n- Tổ chức ôn thi tốt nghiệp THPT môn Lịch sử khối KHXH đạt kết quả xuất sắc.",
            "d4_w": "15%",
            "d4_out": "- Chuyên đề bồi dưỡng HSG Lịch sử.\n- Học sinh đạt giải HSG môn Lịch sử cấp tỉnh.\n- 100% học sinh đỗ tốt nghiệp môn Lịch sử."
        },
        {
            "name": "Giáo viên môn Giáo dục thể chất (GDTC)", "code": "GV-THPT-01", "grade": "Bậc 3 đến Bậc 4", "quota": "02 biên chế",
            "d1_name": "1. Giảng dạy môn GDTC & Rèn luyện thể lực",
            "d1_desc": "- Giảng dạy GDTC theo CT 2018 (Điền kinh, Bóng chuyền, Bóng đá, Cầu lông, Thể dục nhịp điệu) đủ 17 tiết/tuần.\n- Rèn luyện tố chất thể lực nhanh, mạnh, bền, khéo léo; bảo đảm an toàn tuyệt đối trong giờ học ngoài trời.",
            "d1_w": "40%",
            "d1_out": "- Kế hoạch bài dạy GDTC chuẩn bị kỹ lưỡng.\n- Bảng kiểm tra đánh giá tiêu chuẩn rèn luyện thể lực học sinh.\n- 100% tiết dạy an toàn thể lực.",
            "d2_name": "2. Chủ nhiệm lớp & Thể dục giữa giờ toàn trường",
            "d2_desc": "- Công tác chủ nhiệm lớp; rèn nề nếp kỷ luật, tác phong nhanh nhẹn cho học sinh.\n- Phụ trách duy trì tập thể dục giữa giờ và phong trào thể thao toàn trường.",
            "d2_w": "25%",
            "d2_out": "- Kế hoạch chủ nhiệm năm học.\n- Sổ theo dõi nề nếp tập luyện thể dục.\n- Học bạ số ký duyệt đúng hạn.",
            "d3_name": "3. Sinh hoạt NCBL & Huấn luyện thể thao số",
            "d3_desc": "- Sinh hoạt chuyên môn tổ; tự học nâng cao phương pháp huấn luyện thể thao trường học.\n- Ứng dụng công nghệ phân tích kỹ thuật động tác qua video quay chậm.",
            "d3_w": "20%",
            "d3_out": "- Phiếu dự giờ đồng nghiệp.\n- Kế hoạch huấn luyện thể thao trường học.\n- Sáng kiến cải tiến dụng cụ tập luyện.",
            "d4_name": "4. Huấn luyện HKPĐ & Bảo quản sân bãi TDTT",
            "d4_desc": "- Huấn luyện đội tuyển học sinh tham gia Hội khỏe Phù Đổng cấp huyện, cấp tỉnh.\n- Quản lý sân bóng đá, sân bóng chuyền, đường chạy và dụng cụ thể thao nhà trường.",
            "d4_w": "15%",
            "d4_out": "- Đội tuyển đạt giải/huy chương HKPĐ cấp tỉnh.\n- Biên bản kiểm kê dụng cụ TDTT trường học."
        },
        {
            "name": "Giáo viên môn GDQP-AN", "code": "GV-THPT-01", "grade": "Bậc 3 đến Bậc 4", "quota": "01 biên chế",
            "d1_name": "1. Giảng dạy môn GDQP-AN & Kỹ thuật quân sự",
            "d1_desc": "- Giảng dạy GDQP-AN 10, 11, 12 (Đường lối quân sự của Đảng, Điều lệnh đội ngũ, Bắn súng tiểu liên AK, Băng bó cứu thương, Chiến thuật bộ binh) định mức 17 tiết/tuần.\n- Giáo dục ý thức bảo vệ chủ quyền biên giới quốc gia tại huyện biên giới Phục Hòa.",
            "d1_w": "40%",
            "d1_out": "- Kế hoạch bài dạy GDQP-AN đầy đủ.\n- Bảng điểm đánh giá năng lực QP-AN học sinh.\n- 100% tiết thực hành bắn súng laser an toàn tuyệt đối.",
            "d2_name": "2. Chủ nhiệm lớp & Nếp sống nội vụ kỷ luật",
            "d2_desc": "- Công tác chủ nhiệm lớp; rèn luyện nếp sống nội vụ quân sự, tính kỷ luật, tự giác, tinh thần đồng đội cho học sinh.",
            "d2_w": "25%",
            "d2_out": "- Kế hoạch chủ nhiệm năm học.\n- Biên bản đánh giá rèn luyện học sinh.\n- Học bạ số hoàn thành đúng hạn.",
            "d3_name": "3. Tập huấn quân sự & Bảo quản vũ khí mô hình",
            "d3_desc": "- Tham gia tập huấn quân sự của BCH Quân sự tỉnh và Sở GD&ĐT.\n- Bảo quản tủ súng mô hình, trang thiết bị quân trang học đường; ứng dụng mô phỏng 3D bài học quân sự.",
            "d3_w": "20%",
            "d3_out": "- Nhật ký bảo quản vũ khí, trang bị mô hình.\n- Phiếu dự giờ đồng nghiệp.\n- Hồ sơ tập huấn QP-AN hàng năm.",
            "d4_name": "4. Huấn luyện Hội thao GDQP-AN cấp tỉnh",
            "d4_desc": "- Huấn luyện đội tuyển học sinh tham gia Hội thao GDQP-AN cấp tỉnh.\n- Tổ chức các buổi học tập trải nghiệm tại đồn biên phòng, bảo vệ đường biên mốc giới.",
            "d4_w": "15%",
            "d4_out": "- Đội tuyển đạt giải Hội thao GDQP-AN cấp tỉnh.\n- Chương trình trải nghiệm biên cương thực tế."
        },
        {
            "name": "Giáo viên môn Địa lí", "code": "GV-THPT-01", "grade": "Bậc 3 đến Bậc 4", "quota": "02 biên chế",
            "d1_name": "1. Giảng dạy môn Địa lí & Kỹ năng Atlat, GIS",
            "d1_desc": "- Giảng dạy Địa lí 10, 11, 12 (Địa lí tự nhiên, Địa lí kinh tế - xã hội thế giới và Việt Nam) định mức 17 tiết/tuần.\n- Rèn kỹ năng đọc bản đồ, biểu đồ, Atlat Địa lí Việt Nam; khai thác kinh tế cửa khẩu Phục Hòa; giáo dục bảo vệ môi trường, biến đổi khí hậu.",
            "d1_w": "40%",
            "d1_out": "- Kế hoạch bài dạy Địa lí chuẩn quy chế.\n- Hệ thống câu hỏi khai thác Atlat thực hành.\n- Đề kiểm tra định kỳ có bản đặc tả.\n- Tỷ lệ HS đạt chuẩn môn Địa lí ≥ 98%.",
            "d2_name": "2. Chủ nhiệm lớp & Phong trào trường học xanh",
            "d2_desc": "- Công tác chủ nhiệm; giáo dục tinh thần bảo vệ môi trường, phân loại rác thải và giữ gìn trường lớp xanh - sạch - đẹp.",
            "d2_w": "25%",
            "d2_out": "- Kế hoạch chủ nhiệm năm học.\n- Sổ theo dõi lớp chủ nhiệm.\n- Học bạ số hoàn thành đúng tiến độ.",
            "d3_name": "3. Sinh hoạt NCBL & Google Earth, Bản đồ số",
            "d3_desc": "- Sinh hoạt chuyên môn theo NCBL; khai thác phần mềm Google Earth, hệ thống thông tin địa lý GIS.\n- Ứng dụng AI cập nhật số liệu kinh tế xã hội mới nhất.",
            "d3_w": "20%",
            "d3_out": "- Kho học liệu số bản đồ tương tác.\n- Phiếu dự giờ đồng nghiệp.\n- Chuyên đề đổi mới PPDH Địa lí.",
            "d4_name": "4. Bồi dưỡng HSG Địa lí & Ôn thi tốt nghiệp",
            "d4_desc": "- Bồi dưỡng đội tuyển HSG Địa lí cấp tỉnh; ôn tập thi tốt nghiệp THPT môn Địa lí đạt kết quả cao.",
            "d4_w": "15%",
            "d4_out": "- Giáo án bồi dưỡng HSG Địa lí.\n- Học sinh đạt giải HSG môn Địa lí cấp tỉnh.\n- Tỷ lệ đỗ tốt nghiệp môn Địa lí đạt 100%."
        },
        {
            "name": "Giáo viên môn GD KT&PL", "code": "GV-THPT-01", "grade": "Bậc 3 đến Bậc 4", "quota": "01 biên chế",
            "d1_name": "1. Giảng dạy GD KT&PL & Tình huống thực tiễn",
            "d1_desc": "- Giảng dạy GD KT&PL 10, 11, 12 (Kinh tế, Thị trường, Ngân sách nhà nước, Hệ thống pháp luật Việt Nam) định mức 17 tiết/tuần.\n- Rèn kỹ năng xử lý tình huống pháp luật đời sống, quản lý tài chính cá nhân; giáo dục công dân sống và làm việc theo Hiến pháp và pháp luật.",
            "d1_w": "40%",
            "d1_out": "- Kế hoạch bài dạy tình huống thực tiễn.\n- Đề kiểm tra đánh giá năng lực pháp luật.\n- Sổ điểm điện tử cập nhật đúng hạn.\n- Tỷ lệ HS đạt chuẩn môn học ≥ 98%.",
            "d2_name": "2. Chủ nhiệm & Tuyên truyền an toàn pháp luật",
            "d2_desc": "- Công tác chủ nhiệm lớp; tuyên truyền phổ biến giáo dục pháp luật ATGT, phòng chống bạo lực học đường, tệ nạn xã hội trong nhà trường.",
            "d2_w": "25%",
            "d2_out": "- Kế hoạch chủ nhiệm năm học.\n- Hồ sơ tuyên truyền phổ biến pháp luật.\n- Học bạ số được phê duyệt đúng kỳ.",
            "d3_name": "3. Sinh hoạt NCBL & Trợ lý AI pháp luật học đường",
            "d3_desc": "- Sinh hoạt chuyên môn theo NCBL; cập nhật các luật mới ban hành của Quốc hội.\n- Ứng dụng AI phân tích tình huống pháp luật thực tế cho học sinh.",
            "d3_w": "20%",
            "d3_out": "- Bộ tư liệu tình huống pháp luật số hóa.\n- Phiếu dự giờ đồng nghiệp.\n- Chuyên đề giáo dục pháp luật học đường.",
            "d4_name": "4. Bồi dưỡng HSG KT&PL & Ôn thi tốt nghiệp",
            "d4_desc": "- Bồi dưỡng đội tuyển HSG môn GD KT&PL thi cấp tỉnh; phụ đạo học sinh ôn thi tốt nghiệp THPT tổ hợp KHXH.",
            "d4_w": "15%",
            "d4_out": "- Chuyên đề bồi dưỡng HSG KT&PL.\n- Học sinh đạt giải HSG cấp tỉnh môn KT&PL.\n- Tỷ lệ học sinh đỗ tốt nghiệp môn học đạt 100%."
        },
        {
            "name": "Giáo viên môn Vật lí", "code": "GV-THPT-01", "grade": "Bậc 3 đến Bậc 4", "quota": "01 biên chế",
            "d1_name": "1. Giảng dạy môn Vật lí, Phòng thí nghiệm & STEM",
            "d1_desc": "- Giảng dạy môn Vật lí 10, 11, 12 (Cơ học, Nhiệt học, Điện từ học, Quang học, Hạt nhân) định mức 17 tiết/tuần.\n- Thực hiện đầy đủ các tiết thí nghiệm thực hành tại phòng bộ môn; thiết kế các bài học STEM Vật lí ứng dụng thực tiễn; ôn thi TN THPT KHTN.",
            "d1_w": "40%",
            "d1_out": "- Kế hoạch bài dạy môn Vật lí tích hợp thí nghiệm.\n- Kế hoạch bài dạy STEM được phê duyệt.\n- Đề thi trắc nghiệm theo ma trận mới.\n- 100% tiết thực hành bảo đảm an toàn.",
            "d2_name": "2. Chủ nhiệm lớp & Hướng nghiệp khối Kỹ thuật",
            "d2_desc": "- Công tác chủ nhiệm lớp; quản lý nề nếp học tập, rèn luyện kỹ năng sống và tư vấn hướng nghiệp khối ngành kỹ thuật công nghệ cho học sinh.",
            "d2_w": "25%",
            "d2_out": "- Kế hoạch chủ nhiệm năm học.\n- Biên bản sinh hoạt lớp hàng tuần.\n- Học bạ điện tử nhận xét và ký số đúng hạn.",
            "d3_name": "3. Sinh hoạt NCBL, Thí nghiệm ảo PhET & AI",
            "d3_desc": "- Sinh hoạt tổ chuyên môn; bảo quản thiết bị bộ môn Vật lí.\n- Sử dụng phần mềm mô phỏng thí nghiệm ảo PhET; ứng dụng AI tạo câu hỏi trắc nghiệm vật lí định tính và định lượng.",
            "d3_w": "20%",
            "d3_out": "- Kho bài giảng thí nghiệm ảo Vật lí.\n- Phiếu dự giờ đồng nghiệp.\n- Sáng kiến kinh nghiệm cấp cơ sở.",
            "d4_name": "4. Bồi dưỡng HSG Vật lí, NCKH & Ôn thi TN",
            "d4_desc": "- Bồi dưỡng đội tuyển HSG môn Vật lí cấp tỉnh; hướng dẫn học sinh nghiên cứu KHKT dành cho học sinh trung học; ôn thi TN THPT môn Vật lí.",
            "d4_w": "15%",
            "d4_out": "- Học sinh đạt giải HSG môn Vật lí cấp tỉnh.\n- Dự án KHKT/STEM tham gia cấp tỉnh.\n- Phổ điểm tốt nghiệp môn Vật lí đạt cao."
        },
        {
            "name": "Giáo viên môn Hóa học", "code": "GV-THPT-01", "grade": "Bậc 3 đến Bậc 4", "quota": "01 biên chế",
            "d1_name": "1. Giảng dạy môn Hóa học & Thực hành phòng Lab",
            "d1_desc": "- Giảng dạy Hóa học 10, 11, 12 (Cấu tạo nguyên tử, Bảng tuần hoàn, Liên kết hóa học, Nhiệt hóa học, Vô cơ và Hữu cơ) định mức 17 tiết/tuần.\n- Hướng dẫn thực hành thí nghiệm an toàn tuyệt đối; sử dụng phần mềm ChemDraw vẽ cấu tạo phân tử; giáo dục bảo vệ môi trường chống ô nhiễm hóa chất.",
            "d1_w": "40%",
            "d1_out": "- Kế hoạch bài dạy Hóa học gắn thực nghiệm.\n- Báo cáo kết quả thực hành thí nghiệm học sinh.\n- Đề kiểm tra định kỳ có ma trận chuẩn.\n- Tỷ lệ HS đạt chuẩn môn Hóa ≥ 96%.",
            "d2_name": "2. Chủ nhiệm lớp & Hướng nghiệp khối Y Dược",
            "d2_desc": "- Công tác chủ nhiệm lớp; định hướng nghề nghiệp khối ngành Y - Dược, Hóa thực phẩm, Nông nghiệp sạch; quản lý học sinh toàn diện.",
            "d2_w": "25%",
            "d2_out": "- Kế hoạch chủ nhiệm năm học.\n- Hồ sơ theo dõi học sinh lớp chủ nhiệm.\n- Học bạ số ký duyệt đúng kỳ hạn.",
            "d3_name": "3. Sinh hoạt NCBL & An toàn hóa chất phòng Lab",
            "d3_desc": "- Sinh hoạt chuyên môn theo NCBL; quản lý hóa chất, tủ hút phòng bộ môn; phối hợp nhân viên thiết bị thí nghiệm; ứng dụng AI tra cứu phản ứng hóa học.",
            "d3_w": "20%",
            "d3_out": "- Phiếu dự giờ đồng nghiệp.\n- Sổ an toàn phòng thí nghiệm Hóa học.\n- Bài giảng tương tác mô hình 3D phân tử.",
            "d4_name": "4. Bồi dưỡng HSG Hóa học & Ôn thi tốt nghiệp",
            "d4_desc": "- Bồi dưỡng đội tuyển HSG môn Hóa học cấp tỉnh; ôn thi tốt nghiệp THPT môn Hóa đạt phổ điểm cao; phụ đạo học sinh yếu môn Hóa.",
            "d4_w": "15%",
            "d4_out": "- Học sinh đạt giải HSG môn Hóa cấp tỉnh.\n- Kết quả tốt nghiệp môn Hóa đạt chỉ tiêu giao.\n- Chuyên đề bồi dưỡng HSG Hóa học."
        },
        {
            "name": "Giáo viên môn Sinh học", "code": "GV-THPT-01", "grade": "Bậc 3 đến Bậc 4", "quota": "01 biên chế",
            "d1_name": "1. Giảng dạy môn Sinh học & Tiêu bản hiển vi",
            "d1_desc": "- Giảng dạy Sinh học 10, 11, 12 (Tế bào, Vi sinh vật, Sinh học cơ thể, Di truyền, Sinh thái) định mức 17 tiết/tuần.\n- Hướng dẫn học sinh thao tác sử dụng kính hiển vi quang học, làm tiêu bản hiển vi; giáo dục sức khỏe sinh sản, bảo tồn đa dạng sinh học rừng Cao Bằng.",
            "d1_w": "40%",
            "d1_out": "- Kế hoạch bài dạy Sinh học thực nghiệm.\n- Báo cáo thực hành kính hiển vi của học sinh.\n- Đề kiểm tra định kỳ có ma trận chuẩn.\n- Tỷ lệ HS đạt chuẩn môn Sinh ≥ 97%.",
            "d2_name": "2. Chủ nhiệm lớp & Chăm sóc vườn sinh thái",
            "d2_desc": "- Công tác chủ nhiệm; xây dựng phong trào chăm sóc vườn hoa cây cảnh, bảo vệ môi trường sinh thái trường học xanh sạch đẹp; quản lý nề nếp học sinh.",
            "d2_w": "25%",
            "d2_out": "- Kế hoạch chủ nhiệm năm học.\n- Sổ theo dõi nề nếp học sinh.\n- Học bạ số được ký số chuẩn quy chế.",
            "d3_name": "3. Sinh hoạt NCBL & Mô hình 3D Sinh học số",
            "d3_desc": "- Sinh hoạt chuyên môn theo NCBL; xây dựng mô hình sinh thái thực nghiệm.\n- Ứng dụng công cụ AI phân tích quy luật di truyền và tra cứu cơ sở dữ liệu sinh học.",
            "d3_w": "20%",
            "d3_out": "- Kho học liệu hình ảnh, video sinh học tế bào.\n- Phiếu dự giờ đồng nghiệp.\n- Chuyên đề đổi mới PPDH môn Sinh.",
            "d4_name": "4. Bồi dưỡng HSG Sinh học & Ôn thi tốt nghiệp",
            "d4_desc": "- Bồi dưỡng đội tuyển HSG môn Sinh học thi cấp tỉnh; hướng dẫn học sinh đề tài KHKT thuộc lĩnh vực Y - Sinh học; ôn thi TN THPT khối KHTN.",
            "d4_w": "15%",
            "d4_out": "- Học sinh đạt giải HSG Sinh học cấp tỉnh.\n- Dự án NCKH học sinh tham gia cấp tỉnh.\n- Tỷ lệ học sinh đỗ tốt nghiệp môn Sinh đạt 100%."
        },
        {
            "name": "Giáo viên môn Công nghệ", "code": "GV-THPT-01", "grade": "Bậc 3 đến Bậc 4", "quota": "02 biên chế",
            "d1_name": "1. Giảng dạy môn Công nghệ (Công nghiệp & Nông nghiệp)",
            "d1_desc": "- Giảng dạy Công nghệ định hướng Công nghiệp (Bản vẽ, Cơ khí, Điện) và Nông nghiệp công nghệ cao (Trồng trọt, Chăn nuôi, Thủy sản sạch) định mức 17 tiết/tuần.\n- Hướng dẫn thực hành an toàn điện và mô hình nông nghiệp thích ứng khí hậu vùng núi Phục Hòa.",
            "d1_w": "40%",
            "d1_out": "- Kế hoạch bài dạy Công nghệ gắn thực tiễn.\n- Sản phẩm thực hành mạch điện của học sinh.\n- Báo cáo dự án nông nghiệp công nghệ cao.\n- Tỷ lệ HS đạt chuẩn môn Công nghệ ≥ 98%.",
            "d2_name": "2. Chủ nhiệm lớp & Phân luồng đào tạo nghề",
            "d2_desc": "- Công tác chủ nhiệm lớp; giáo dục hướng nghiệp, phân luồng học sinh sau THPT tiếp cận học nghề kỹ thuật công nghiệp hoặc đại học công nghệ.",
            "d2_w": "25%",
            "d2_out": "- Kế hoạch chủ nhiệm năm học.\n- Hồ sơ tư vấn hướng nghiệp học sinh.\n- Học bạ số hoàn thành đúng hạn.",
            "d3_name": "3. Sinh hoạt NCBL & Mô phỏng mạch điện AutoCAD",
            "d3_desc": "- Sinh hoạt tổ chuyên môn; khai thác phần mềm AutoCAD, mô phỏng mạch điện.\n- Ứng dụng AI tìm kiếm ý tưởng thiết kế sản phẩm STEM công nghệ tiện ích.",
            "d3_w": "20%",
            "d3_out": "- Kho học liệu số mô hình kỹ thuật 3D.\n- Phiếu dự giờ đồng nghiệp.\n- Sản phẩm sáng tạo kỹ thuật của giáo viên.",
            "d4_name": "4. Sáng tạo KHKT & Ngày hội STEM trường học",
            "d4_desc": "- Hướng dẫn học sinh tham gia cuộc thi Sáng tạo thanh thiếu niên nhi đồng và KHKT cấp tỉnh.\n- Nòng cốt tổ chức Ngày hội STEM trường học; phụ đạo học sinh.",
            "d4_w": "15%",
            "d4_out": "- Sản phẩm/mô hình đạt giải Sáng tạo KHKT cấp tỉnh.\n- Kế hoạch và sản phẩm Ngày hội STEM trường học."
        },
        {
            "name": "Giáo viên môn Tin học", "code": "GV-THPT-01", "grade": "Bậc 3 đến Bậc 4", "quota": "01 biên chế",
            "d1_name": "1. Giảng dạy Tin học, Lập trình Python & Quản trị phòng máy",
            "d1_desc": "- Giảng dạy Tin học 10, 11, 12 theo CT 2018 (Khoa học máy tính & Tin học ứng dụng: Python, CSDL số, Mạng, An ninh mạng, AI) định mức 17 tiết/tuần.\n- Quản lý bảo đảm 02 phòng máy vi tính vận hành ổn định phục vụ thực hành cho toàn trường.",
            "d1_w": "40%",
            "d1_out": "- Kế hoạch bài dạy Tin học lập trình Python.\n- Bài tập dự án phần mềm của học sinh.\n- Đề thi thực hành trên máy tính chấm tự động.\n- Hệ thống phòng máy vi tính hoạt động 99.9%.",
            "d2_name": "2. Chủ nhiệm lớp & Văn hóa an toàn không gian mạng",
            "d2_desc": "- Công tác chủ nhiệm lớp; giáo dục văn hóa ứng xử trên không gian mạng, phòng chống lừa đảo trực tuyến và bảo vệ dữ liệu cá nhân cho học sinh.",
            "d2_w": "25%",
            "d2_out": "- Kế hoạch chủ nhiệm năm học.\n- Biên bản sinh hoạt chuyên đề an toàn mạng.\n- Học bạ số ký duyệt đúng thời hạn.",
            "d3_name": "3. Nòng cốt CĐS & Tập huấn AI toàn trường",
            "d3_desc": "- Quản trị kỹ thuật mạng LAN, máy chủ nhà trường.\n- Tập huấn hướng dẫn cán bộ giáo viên sử dụng phần mềm dạy học số và các công cụ AI (Gemini, ChatGPT) trong soạn giảng.",
            "d3_w": "20%",
            "d3_out": "- Nhật ký bảo trì hệ thống mạng & phòng máy.\n- Tài liệu tập huấn ứng dụng AI cho giáo viên.\n- Phiếu dự giờ đồng nghiệp.",
            "d4_name": "4. Bồi dưỡng HSG Tin học & Hỗ trợ thi trực tuyến",
            "d4_desc": "- Bồi dưỡng đội tuyển HSG Tin học lập trình dự thi cấp tỉnh.\n- Trực kỹ thuật các kỳ thi khảo sát trực tuyến, thi nghề số hóa của nhà trường.",
            "d4_w": "15%",
            "d4_out": "- Đội tuyển HSG Tin học đạt giải cấp tỉnh.\n- Hệ thống thi trực tuyến thông suốt 100%."
        }
    ]

    r_idx = 4
    stt_counter = 1
    for s in subjects_specs:
        vtvl_info = f"{s['name']}\n({s['code']})\n{s['grade']} | {s['quota']}\nĐịnh mức: 17 tiết/tuần"
        
        # Duty 1
        ws4.cell(r_idx, 1, stt_counter)
        ws4.cell(r_idx, 2, vtvl_info)
        ws4.cell(r_idx, 3, s["d1_name"])
        ws4.cell(r_idx, 4, s["d1_desc"])
        ws4.cell(r_idx, 5, s["d1_w"])
        ws4.cell(r_idx, 6, s["d1_out"])
        apply_row_styles(ws4, r_idx, font=FONT_REGULAR)
        ws4.cell(r_idx, 1).alignment = ALIGN_CENTER
        ws4.cell(r_idx, 2).font = FONT_BOLD
        ws4.cell(r_idx, 5).alignment = ALIGN_CENTER
        r_idx += 1
        stt_counter += 1

        # Duty 2
        ws4.cell(r_idx, 1, stt_counter)
        ws4.cell(r_idx, 2, f"{s['name']} (tiếp)")
        ws4.cell(r_idx, 3, s["d2_name"])
        ws4.cell(r_idx, 4, s["d2_desc"])
        ws4.cell(r_idx, 5, s["d2_w"])
        ws4.cell(r_idx, 6, s["d2_out"])
        apply_row_styles(ws4, r_idx, font=FONT_REGULAR)
        ws4.cell(r_idx, 1).alignment = ALIGN_CENTER
        ws4.cell(r_idx, 5).alignment = ALIGN_CENTER
        r_idx += 1
        stt_counter += 1

        # Duty 3
        ws4.cell(r_idx, 1, stt_counter)
        ws4.cell(r_idx, 2, f"{s['name']} (tiếp)")
        ws4.cell(r_idx, 3, s["d3_name"])
        ws4.cell(r_idx, 4, s["d3_desc"])
        ws4.cell(r_idx, 5, s["d3_w"])
        ws4.cell(r_idx, 6, s["d3_out"])
        apply_row_styles(ws4, r_idx, font=FONT_REGULAR)
        ws4.cell(r_idx, 1).alignment = ALIGN_CENTER
        ws4.cell(r_idx, 5).alignment = ALIGN_CENTER
        r_idx += 1
        stt_counter += 1

        # Duty 4
        ws4.cell(r_idx, 1, stt_counter)
        ws4.cell(r_idx, 2, f"{s['name']} (tiếp)")
        ws4.cell(r_idx, 3, s["d4_name"])
        ws4.cell(r_idx, 4, s["d4_desc"])
        ws4.cell(r_idx, 5, s["d4_w"])
        ws4.cell(r_idx, 6, s["d4_out"])
        apply_row_styles(ws4, r_idx, font=FONT_REGULAR)
        ws4.cell(r_idx, 1).alignment = ALIGN_CENTER
        ws4.cell(r_idx, 5).alignment = ALIGN_CENTER
        r_idx += 1
        stt_counter += 1

    # ==============================================================================
    # SHEET 5: DANH MỤC 3: BẢNG MÔ TẢ VTVL 06 HỖ TRỢ (06 BIÊN CHẾ)
    # ==============================================================================
    ws5 = wb.create_sheet(title="5. DM3 - VTVL 06 Hỗ trợ")
    ws5.column_dimensions['A'].width = 8
    ws5.column_dimensions['B'].width = 30
    ws5.column_dimensions['C'].width = 32
    ws5.column_dimensions['D'].width = 46
    ws5.column_dimensions['E'].width = 12
    ws5.column_dimensions['F'].width = 42

    ws5.merge_cells("A1:F1")
    ws5["A1"] = "DANH MỤC 3: BẢN MÔ TẢ VỊ TRÍ VIỆC LÀM VIÊN CHỨC HỖ TRỢ (06 NGƯỜI)"
    ws5["A1"].font = FONT_TITLE_MAIN
    ws5["A1"].alignment = ALIGN_CENTER

    ws5.append([])
    ws5.append(headers_6col)
    for c_i in range(1, len(headers_6col) + 1):
        cell = ws5.cell(row=3, column=c_i)
        cell.font = FONT_HEADER_WHITE
        cell.fill = AMBER_HEADER
        cell.alignment = ALIGN_CENTER
        cell.border = BORDER_HEADER

    support_specs = [
        {
            "name": "Kế toán viên", "code": "KT-THPT-01", "grade": "Bậc 1 đến Bậc 4", "quota": "01 biên chế",
            "d1_name": "1. Quản lý dự toán, ngân sách & Báo cáo tài chính",
            "d1_desc": "- Quản lý tài chính, ngân sách nhà nước và các nguồn thu hợp pháp; lập dự toán ngân sách hàng năm.\n- Thực hiện thanh quyết toán qua Kho bạc Nhà nước; lập báo cáo tài chính quý/năm theo Luật Kế toán; lưu trữ chứng từ kế toán an toàn, bảo mật.",
            "d1_w": "60%",
            "d1_out": "- Dự toán ngân sách năm học được duyệt.\n- Báo cáo quyết toán tài chính quý/năm được Sở GD&ĐT phê chuẩn.\n- Sổ cái, sổ chi tiết các tài khoản kế toán.\n- 100% hồ sơ thanh toán Kho bạc không bị từ chối.",
            "d2_name": "2. Tiền lương, phụ cấp, BHXH & Chế độ học sinh",
            "d2_desc": "- Tính toán chi trả tiền lương, phụ cấp ưu đãi nghề, phụ cấp thâm niên, BHXH, BHYT, chế độ thừa giờ cho cán bộ giáo viên đúng hạn.\n- Chi trả chế độ chính sách miễn giảm học phí, hỗ trợ chi phí học tập cho học sinh vùng đặc biệt khó khăn kịp thời, chính xác.",
            "d2_w": "40%",
            "d2_out": "- Bảng thanh toán lương chuyển khoản trước ngày 05 hàng tháng.\n- Hồ sơ giải quyết chế độ chính sách học sinh đúng đối tượng, đủ định mức.\n- Chứng từ chi trả chế độ lưu trữ đúng quy định."
        },
        {
            "name": "Văn thư trường học", "code": "VT-THPT-01", "grade": "Bậc 1 đến Bậc 4", "quota": "01 biên chế",
            "d1_name": "1. Quản lý văn bản điện tử, công văn đi - đến",
            "d1_desc": "- Quản lý hệ thống văn bản đi - đến trên phần mềm Quản lý văn bản điện tử và điều hành (iOffice).\n- Tiếp nhận, vào sổ, chuyển giao công văn trong ngày làm việc.\n- Rà soát thể thức văn bản hành chính theo Nghị định 30/2020/NĐ-CP trước khi trình ký; phát hành văn bản điện tử ký số đúng quy trình.",
            "d1_w": "60%",
            "d1_out": "- Sổ đăng ký văn bản đi / đến điện tử cập nhật 100%.\n- 100% công văn đến được chuyển giao kịp thời trong ngày.\n- Văn bản phát hành bảo đảm đúng thể thức và kỹ thuật trình bày.",
            "d2_name": "2. Quản lý con dấu, chứng thư số & Lưu trữ hồ sơ",
            "d2_desc": "- Quản lý, sử dụng và bảo quản con dấu của trường THPT Phục Hòa và thiết bị lưu khóa bí mật (chứng thư số) an toàn tuyệt đối.\n- Lập danh mục hồ sơ cơ quan, thu thập và lập hồ sơ tài liệu lưu trữ hiện hành; phục vụ tra cứu tài liệu lưu trữ nhanh chóng, chính xác.",
            "d2_w": "40%",
            "d2_out": "- Sổ theo dõi đóng dấu và sử dụng chứng thư số.\n- Danh mục hồ sơ lưu trữ năm học chuẩn quy chế.\n- Không để xảy ra thất lạc tài liệu hoặc sử dụng con dấu sai quy định."
        },
        {
            "name": "Thiết bị, thí nghiệm", "code": "TBTN-THPT-01", "grade": "Bậc 2 đến Bậc 3", "quota": "01 biên chế",
            "d1_name": "1. Quản lý thiết bị dạy học & Chuẩn bị phòng thực hành",
            "d1_desc": "- Quản lý, bảo quản toàn bộ trang thiết bị dạy học, đồ dùng dạy học, mô hình, hóa chất tại các phòng thực hành Vật lí, Hóa học, Sinh học, Công nghệ.\n- Chuẩn bị đầy đủ thiết bị, dụng cụ, hóa chất thí nghiệm trước các tiết dạy thực hành của giáo viên theo lịch báo giảng; hướng dẫn quy định an toàn phòng thí nghiệm.",
            "d1_w": "60%",
            "d1_out": "- Sổ theo dõi sử dụng phòng học bộ môn & mượn trả thiết bị dạy học.\n- 100% tiết thực hành có thiết bị chuẩn bị sẵn sàng trước giờ học.\n- Tuyệt đối an toàn cháy nổ, không để xảy ra sự cố phòng thí nghiệm.",
            "d2_name": "2. Kiểm kê, bảo dưỡng & Đề xuất mua sắm bổ sung",
            "d2_desc": "- Kiểm kê thiết bị dạy học định kỳ cuối học kỳ và năm học; lập hồ sơ đề xuất thanh lý thiết bị hỏng không thể khắc phục.\n- Lập dự trù đề xuất mua sắm bổ sung hóa chất, dụng cụ thực hành tiêu hao đầu năm học mới; vệ sinh, bảo dưỡng trang thiết bị máy móc thường xuyên.",
            "d2_w": "40%",
            "d2_out": "- Biên bản kiểm kê tài sản thiết bị dạy học cuối năm.\n- Danh mục đề xuất mua sắm bổ sung TBDH được phê duyệt.\n- Sổ nhật ký bảo dưỡng thiết bị dạy học ghi chép định kỳ."
        },
        {
            "name": "Giáo vụ trường học", "code": "GVU-THPT-01", "grade": "Bậc 2 đến Bậc 3", "quota": "01 biên chế",
            "d1_name": "1. Quản lý sổ đăng bộ, hồ sơ học sinh & CSDL ngành",
            "d1_desc": "- Quản lý sổ đăng bộ của nhà trường; quản lý và lưu trữ hồ sơ tuyển sinh vào lớp 10, hồ sơ học sinh chuyển trường đi - chuyển trường đến, hồ sơ học sinh thôi học.\n- Quản lý cơ sở dữ liệu học sinh trên hệ thống CSDL ngành; kiểm tra tính hợp lệ và đồng bộ dữ liệu học bạ số của các khối lớp.",
            "d1_w": "50%",
            "d1_out": "- Sổ đăng bộ trường THPT số hóa chính xác 100%.\n- Hồ sơ tuyển sinh lớp 10 được phê duyệt đầy đủ.\n- Cơ sở dữ liệu ngành cập nhật đầy đủ, không sai sót thông tin học sinh.",
            "d2_name": "2. Phục vụ kỳ thi TN THPT & Quản lý cấp phát văn bằng",
            "d2_desc": "- Phối hợp tổ chức kỳ thi tốt nghiệp THPT, kỳ thi chọn học sinh giỏi cấp tỉnh và các kỳ kiểm tra chung của trường (lập danh sách thí sinh, đánh số báo danh, chuẩn bị thẻ dự thi, hồ sơ phòng thi).\n- Quản lý sổ cấp phát bằng tốt nghiệp THPT, học bạ và giấy chứng nhận tốt nghiệp tạm thời cho học sinh.",
            "d2_w": "50%",
            "d2_out": "- Hồ sơ đăng ký dự thi tốt nghiệp THPT nộp Sở GD&ĐT đúng thời hạn.\n- Sổ cấp phát văn bằng, chứng chỉ có đầy đủ chữ ký người nhận.\n- Cấp phát bằng tốt nghiệp kịp thời, chính xác 100%."
        },
        {
            "name": "Y tế trường học", "code": "YT-THPT-01", "grade": "Bậc 1 đến Bậc 3", "quota": "01 biên chế",
            "d1_name": "1. Chăm sóc sức khỏe ban đầu, sơ cấp cứu & Tủ thuốc",
            "d1_desc": "- Chăm sóc sức khỏe ban đầu, sơ cứu và cấp cứu kịp thời các tai nạn thương tích hoặc ốm đau đột xuất của học sinh và CBGVNV.\n- Quản lý tủ thuốc y tế học đường, dự trù cơ số thuốc thiết yếu, bông băng, dụng cụ y tế; lập và quản lý hồ sơ theo dõi sức khỏe từng học sinh qua các năm học.",
            "d1_w": "50%",
            "d1_out": "- Sổ theo dõi khám chữa bệnh và cấp phát thuốc.\n- Sổ theo dõi sức khỏe định kỳ học sinh.\n- Tủ thuốc y tế có đầy đủ thuốc sơ cứu còn hạn sử dụng.\n- 100% ca sơ cứu học đường được xử lý kịp thời, an toàn.",
            "d2_name": "2. Khám sức khỏe định kỳ, phòng dịch & BHYT học sinh",
            "d2_desc": "- Phối hợp với Trung tâm Y tế huyện Phục Hòa tổ chức khám sức khỏe định kỳ cho học sinh đầu năm học.\n- Triển khai công tác phòng chống dịch bệnh truyền nhiễm trong trường học (sốt xuất huyết, cúm, Covid...).\n- Kiểm tra điều kiện vệ sinh học đường, nguồn nước uống và an toàn thực phẩm; tuyên truyền bảo hiểm y tế học sinh đạt 100%.",
            "d2_w": "50%",
            "d2_out": "- Báo cáo kết quả khám sức khỏe học sinh toàn trường.\n- Kế hoạch và biên bản kiểm tra vệ sinh an toàn thực phẩm, nước uống.\n- 100% học sinh tham gia BHYT học sinh; không xảy ra dịch bệnh học đường."
        },
        {
            "name": "Thủ quỹ trường học", "code": "TQ-THPT-01", "grade": "Bậc 1 đến Bậc 3", "quota": "01 biên chế",
            "d1_name": "1. Quản lý két sắt, quỹ tiền mặt & Thu - chi đúng phiếu",
            "d1_desc": "- Trực tiếp quản lý quỹ tiền mặt tại két sắt của nhà trường; bảo đảm an toàn tuyệt đối tiền mặt trong mọi tình huống.\n- Thực hiện nghiêm ngặt nguyên tắc tài chính: chỉ thu tiền hoặc xuất tiền khi có Phiếu thu, Phiếu chi hợp pháp có đầy đủ chữ ký của Hiệu trưởng (Chủ tài khoản) và Kế toán.\n- Ký nhận và yêu cầu người nhận tiền ký nhận rõ ràng, ghi rõ họ tên trên phiếu chi.",
            "d1_w": "60%",
            "d1_out": "- 100% phiếu thu, phiếu chi có đầy đủ chữ ký hợp lệ trước khi xuất/nhập quỹ.\n- Quỹ tiền mặt an toàn tuyệt đối, không thất thoát.\n- Tiền mặt được kiểm đếm chính xác, phân loại ngăn nắp trong két sắt chống cháy.",
            "d2_name": "2. Ghi chép sổ quỹ hàng ngày & Đối chiếu khớp số dư",
            "d2_desc": "- Mở sổ quỹ tiền mặt và ghi chép sổ quỹ hàng ngày theo trình tự thời gian phát sinh nghiệp vụ thu, chi.\n- Định kỳ hàng ngày, tuần, tháng đối chiếu số tồn quỹ thực tế với Kế toán; lập biên bản kiểm kê quỹ tiền mặt đột xuất và định kỳ cuối tháng/quý/năm.\n- Nộp tiền mặt vào tài khoản Kho bạc/Ngân hàng theo đúng quy định quản lý tiền mặt.",
            "d2_w": "40%",
            "d2_out": "- Sổ quỹ tiền mặt ghi chép rõ ràng, sạch sẽ, cập nhật hàng ngày.\n- Biên bản đối chiếu khớp 100% số dư tiền mặt giữa Thủ quỹ và Kế toán.\n- Biên bản kiểm kê quỹ tiền mặt định kỳ cuối tháng/quý/năm."
        }
    ]

    r_idx = 4
    stt_counter = 1
    for su in support_specs:
        vtvl_info = f"{su['name']}\n({su['code']})\n{su['grade']} | {su['quota']}\nĐịnh mức: 40 giờ/tuần"

        # Duty 1
        ws5.cell(r_idx, 1, stt_counter)
        ws5.cell(r_idx, 2, vtvl_info)
        ws5.cell(r_idx, 3, su["d1_name"])
        ws5.cell(r_idx, 4, su["d1_desc"])
        ws5.cell(r_idx, 5, su["d1_w"])
        ws5.cell(r_idx, 6, su["d1_out"])
        apply_row_styles(ws5, r_idx, font=FONT_REGULAR)
        ws5.cell(r_idx, 1).alignment = ALIGN_CENTER
        ws5.cell(r_idx, 2).font = FONT_BOLD
        ws5.cell(r_idx, 5).alignment = ALIGN_CENTER
        r_idx += 1
        stt_counter += 1

        # Duty 2
        ws5.cell(r_idx, 1, stt_counter)
        ws5.cell(r_idx, 2, f"{su['name']} (tiếp)")
        ws5.cell(r_idx, 3, su["d2_name"])
        ws5.cell(r_idx, 4, su["d2_desc"])
        ws5.cell(r_idx, 5, su["d2_w"])
        ws5.cell(r_idx, 6, su["d2_out"])
        apply_row_styles(ws5, r_idx, font=FONT_REGULAR)
        ws5.cell(r_idx, 1).alignment = ALIGN_CENTER
        ws5.cell(r_idx, 5).alignment = ALIGN_CENTER
        r_idx += 1
        stt_counter += 1

    # ==============================================================================
    # SHEET 6: BẢN MÔ TẢ CHUẨN NGHỊ ĐỊNH 232/2026/NĐ-CP (PHẦN B PHỤ LỤC IV/VI)
    # ==============================================================================
    ws6 = wb.create_sheet(title="6. Mau Mo Ta Chuan ND 232")
    ws6.column_dimensions['A'].width = 10
    ws6.column_dimensions['B'].width = 25
    ws6.column_dimensions['C'].width = 38
    ws6.column_dimensions['D'].width = 32
    ws6.column_dimensions['E'].width = 26

    ws6.merge_cells("A1:B1")
    ws6["A1"] = "SỞ GIÁO DỤC VÀ ĐÀO TẠO CAO BẰNG\nTRƯỜNG THPT PHỤC HÒA"
    ws6["A1"].font = Font(name=FONT_FAMILY, size=10, bold=True)
    ws6["A1"].alignment = ALIGN_CENTER

    ws6.merge_cells("C1:E1")
    ws6["C1"] = "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM\nĐộc lập – Tự do – Hạnh phúc\n_______________________"
    ws6["C1"].font = Font(name=FONT_FAMILY, size=10, bold=True)
    ws6["C1"].alignment = ALIGN_CENTER

    ws6.merge_cells("A3:E3")
    ws6["A3"] = "BẢN MÔ TẢ CÔNG VIỆC VÀ KHUNG NĂNG LỰC VỊ TRÍ VIỆC LÀM VIÊN CHỨC"
    ws6["A3"].font = Font(name=FONT_FAMILY, size=13, bold=True, color="1E3A8A")
    ws6["A3"].alignment = ALIGN_CENTER

    ws6.merge_cells("A4:E4")
    ws6["A4"] = "(Mẫu chuẩn theo Phần B Phụ lục IV/VI của Nghị định số 232/2026/NĐ-CP ngày 26/6/2026)"
    ws6["A4"].font = FONT_ITALIC
    ws6["A4"].alignment = ALIGN_CENTER

    doc_content = [
        ("I", "THÔNG TIN CHUNG VỀ VỊ TRÍ VIỆC LÀM", ""),
        ("1", "Tên vị trí việc làm", "Giáo viên trung học phổ thông (Giáo viên 13 môn học trong Phụ lục I) / Viên chức Hỗ trợ"),
        ("2", "Mã số vị trí việc làm", "GV-THPT-01 (Chuyên môn) | KT, VT, TBTN, GVU, YT, TQ (Hỗ trợ)"),
        ("3", "Bậc chức danh nghề nghiệp", "Chuyên môn: Bậc 3 đến Bậc 4 | Hỗ trợ: Bậc 1 đến Bậc 4"),
        ("4", "Cơ quan, đơn vị sử dụng", "Trường THPT Phục Hòa, huyện Phục Hòa, tỉnh Cao Bằng"),
        ("5", "Vị trí việc làm cấp trên", "Tổ trưởng chuyên môn, Phó Hiệu trưởng, Hiệu trưởng Trường THPT Phục Hòa"),
        ("6", "Định mức làm việc", "Giáo viên: 17 tiết/tuần | Viên chức hỗ trợ: 40 giờ/tuần"),
        ("II", "MỤC TIÊU VỊ TRÍ VIỆC LÀM", 
         "Đảm nhận nhiệm vụ giảng dạy và giáo dục học sinh theo Chương trình GDPT 2018 (đối với GV 13 môn học) và bảo đảm các điều kiện về tài chính, văn thư, thiết bị thí nghiệm, giáo vụ, y tế học đường, ngân quỹ phục vụ vận hành thông suốt của nhà trường (đối với 06 viên chức hỗ trợ)."),
        ("III", "CÁC CÔNG VIỆC VÀ KẾT QUẢ ĐẦU RA CỤ THỂ", 
         "Chi tiết cụ thể theo từng môn học và từng vị trí hỗ trợ được quy định tại Sheet 3 (Quản lý), Sheet 4 (13 Môn học Chuyên môn) và Sheet 5 (06 Vị trí Hỗ trợ)."),
        ("IV", "KHUNG NĂNG LỰC VỊ TRÍ VIỆC LÀM",
         "1. Năng lực chung:\n - Đạo đức nghề nghiệp, liêm chính sư phạm, phục vụ nhân dân (Cấp độ 4 - 5)\n - Kỷ luật, trách nhiệm và tinh thần hợp tác đồng nghiệp (Cấp độ 4 - 5)\n - Đổi mới sáng tạo, chuyển đổi số và ứng dụng AI (Cấp độ 3 - 4)\n2. Năng lực chuyên môn nghiệp vụ:\n - Nắm vững kiến thức bộ môn theo CT GDPT 2018 (Cấp độ 4 - 5)\n - Kỹ năng sư phạm, tổ chức hoạt động học tập, dạy học STEM (Cấp độ 4 - 5)\n - Đánh giá năng lực học sinh theo thông tư Bộ GD&ĐT (Cấp độ 4)"),
        ("V", "MỐI QUAN HỆ CÔNG TÁC",
         "1. Quan hệ bên trong: Báo cáo trực tiếp Tổ trưởng CM, Phó Hiệu trưởng, Hiệu trưởng; phối hợp với GV chủ nhiệm, GV bộ môn, nhân viên Thiết bị, Thư viện, Y tế, Giáo vụ, Thủ quỹ.\n2. Quan hệ bên ngoài: Phối hợp thường xuyên với Cha mẹ học sinh; Hội Khuyến học địa phương; Phòng chuyên môn Sở GD&ĐT Cao Bằng."),
        ("VI", "PHẠM VI QUYỀN HẠN",
         "1. Về chuyên môn: Chủ động lựa chọn PPDH, kiểm tra đánh giá theo kế hoạch giáo dục môn học; đánh giá xếp loại học sinh theo quy chế; tham gia biên soạn tài liệu giáo dục địa phương.\n2. Về quản lý: Quản lý học sinh trong giờ học và các hoạt động giáo dục phân công.\n3. Về bảo đảm điều kiện: Được trang bị đầy đủ phương tiện dạy học, phòng thực hành bộ môn, thiết bị CNTT, chế độ phụ cấp và bảo hộ lao động theo quy định."),
        ("VII", "YÊU CẦU VỀ TRÌNH ĐỘ, KINH NGHIỆM, PHẨM CHẤT",
         "- Về trình độ đào tạo: Bằng Cử nhân (Đại học) sư phạm trở lên đúng chuyên ngành đối với Giáo viên; Bằng Đại học/Cao đẳng/Trung cấp chuyên ngành phù hợp đối với Viên chức hỗ trợ (Kế toán, Văn thư, Thiết bị, Y tế...).\n- Về chứng chỉ: Có chứng chỉ bồi dưỡng tiêu chuẩn chức danh nghề nghiệp viên chức theo quy định.\n- Về kỹ năng số: Sử dụng thành thạo máy vi tính, phần mềm quản lý giáo dục điện tử, ký số học bạ, khai thác công cụ AI hỗ trợ công việc.\n- Phẩm chất đạo đức: Mẫu mực, tâm huyết với sự nghiệp giáo dục vùng cao Phục Hòa.")
    ]

    r_idx = 6
    for it in doc_content:
        if it[0] in ["I", "II", "III", "IV", "V", "VI", "VII"]:
            ws6.merge_cells(start_row=r_idx, start_column=1, end_row=r_idx, end_column=5)
            ws6.cell(r_idx, 1, f"{it[0]}. {it[1]}")
            apply_row_styles(ws6, r_idx, font=FONT_SECTION, fill=SECTION_FILL)
            r_idx += 1
            if it[2]:
                ws6.merge_cells(start_row=r_idx, start_column=1, end_row=r_idx, end_column=5)
                ws6.cell(r_idx, 1, it[2])
                apply_row_styles(ws6, r_idx, font=FONT_REGULAR)
                r_idx += 1
        else:
            ws6.cell(r_idx, 1, it[0])
            ws6.cell(r_idx, 2, it[1])
            ws6.merge_cells(start_row=r_idx, start_column=3, end_row=r_idx, end_column=5)
            ws6.cell(r_idx, 3, it[2])
            apply_row_styles(ws6, r_idx, font=FONT_REGULAR)
            ws6.cell(r_idx, 1).alignment = ALIGN_CENTER
            ws6.cell(r_idx, 2).font = FONT_BOLD
            r_idx += 1

    # Khung ký duyệt
    r_idx += 1
    ws6.merge_cells(start_row=r_idx, start_column=1, end_row=r_idx, end_column=2)
    ws6.cell(r_idx, 1, "NGƯỜI LẬP BIỂU\n(Ký và ghi rõ họ tên)")
    ws6.cell(r_idx, 1).font = FONT_BOLD
    ws6.cell(r_idx, 1).alignment = ALIGN_CENTER

    ws6.merge_cells(start_row=r_idx, start_column=3, end_row=r_idx, end_column=5)
    ws6.cell(r_idx, 3, "Phục Hòa, ngày .... tháng .... năm 2026\nHIỆU TRƯỞNG\n(Ký tên, đóng dấu)")
    ws6.cell(r_idx, 3).font = FONT_BOLD
    ws6.cell(r_idx, 3).alignment = ALIGN_CENTER

    # ==============================================================================
    # SHEET 7: KHUNG NĂNG LỰC TOÀN DIỆN 3 NHÓM (5 CẤP ĐỘ)
    # ==============================================================================
    ws7 = wb.create_sheet(title="7. Khung năng lực 3 Nhóm")
    ws7.column_dimensions['A'].width = 8
    ws7.column_dimensions['B'].width = 24
    ws7.column_dimensions['C'].width = 30
    ws7.column_dimensions['D'].width = 16
    ws7.column_dimensions['E'].width = 18
    ws7.column_dimensions['F'].width = 16
    ws7.column_dimensions['G'].width = 46

    ws7.merge_cells("A1:G1")
    ws7["A1"] = "KHUNG NĂNG LỰC TOÀN DIỆN CHO 03 NHÓM VỊ TRÍ VIỆC LÀM (5 CẤP ĐỘ)"
    ws7["A1"].font = FONT_TITLE_MAIN
    ws7["A1"].alignment = ALIGN_CENTER

    headers_ws7 = ["STT", "Nhóm năng lực", "Tên năng lực cụ thể", "DM1: Quản lý", "DM2: Chuyên môn (13 môn)", "DM3: Hỗ trợ (6 VC)", "Mô tả chuẩn hành vi yêu cầu"]
    ws7.append([])
    ws7.append(headers_ws7)
    for col_idx in range(1, len(headers_ws7) + 1):
        cell = ws7.cell(row=3, column=col_idx)
        cell.font = FONT_HEADER_WHITE
        cell.fill = PURPLE_HEADER
        cell.alignment = ALIGN_CENTER
        cell.border = BORDER_HEADER

    comp_rows = [
        ("1", "Năng lực chung", "Phẩm chất chính trị & Đạo đức nghề nghiệp", "Cấp độ 4 - 5", "Cấp độ 3 - 4", "Cấp độ 2 - 3", "Tuyệt đối trung thành, gương mẫu, giữ gìn phẩm chất nhà giáo, liêm chính."),
        ("2", "Năng lực chung", "Giao tiếp, ứng xử & Tinh thần hợp tác", "Cấp độ 4 - 5", "Cấp độ 3 - 4", "Cấp độ 2 - 3", "Giao tiếp sư phạm chuẩn mực, tôn trọng học sinh, phụ huynh và đồng nghiệp."),
        ("3", "Năng lực chung", "Đổi mới, sáng tạo & Chuyển đổi số", "Cấp độ 4 - 5", "Cấp độ 3 - 4", "Cấp độ 2 - 3", "Chủ động đề xuất cải tiến quy trình làm việc, ứng dụng CNTT và AI hiệu quả."),
        ("4", "Năng lực chung", "Kỹ năng số, CNTT & Công cụ AI", "Cấp độ 4 - 5", "Cấp độ 3 - 4", "Cấp độ 2 - 3", "Khai thác thành thạo phần mềm quản lý, ký số học bạ điện tử, công cụ AI."),
        ("5", "Năng lực quản lý", "Tư duy chiến lược & Lập kế hoạch", "Cấp độ 4 - 5", "Cấp độ 2 - 3", "Cấp độ 1 - 2", "Hoạch định chiến lược phát triển trường, kế hoạch giáo dục tổ chuyên môn."),
        ("6", "Năng lực quản lý", "Tổ chức điều hành & Kiểm tra giám sát", "Cấp độ 4 - 5", "Cấp độ 2 - 3", "Cấp độ 1 - 2", "Phân công nhiệm vụ hợp lý, kiểm tra đánh giá tiến độ và kết quả công việc."),
        ("7", "Năng lực quản lý", "Quản trị nhân lực & Phát triển đội ngũ", "Cấp độ 4 - 5", "Cấp độ 2 - 3", "Cấp độ 1", "Xây dựng khối đoàn kết nội bộ, động viên khích lệ cán bộ giáo viên phát triển."),
        ("8", "Năng lực chuyên môn", "Nghiệp vụ sư phạm / Chuyên ngành", "Cấp độ 4 - 5", "Cấp độ 3 - 5", "Cấp độ 2 - 3", "Nắm vững chuyên môn kiến thức môn học, phương pháp dạy học tích cực, STEM."),
        ("9", "Năng lực chuyên môn", "Xây dựng kế hoạch dạy học & Bài dạy", "Cấp độ 4 - 5", "Cấp độ 3 - 5", "Cấp độ 1 - 2", "Thiết kế giáo án chuẩn CT GDPT 2018, ma trận và đề kiểm tra định kỳ."),
        ("10", "Năng lực chuyên môn", "Nghiên cứu khoa học & Bồi dưỡng HSG", "Cấp độ 3 - 4", "Cấp độ 3 - 5", "Cấp độ 1 - 2", "Có sáng kiến kinh nghiệm cấp cơ sở, bồi dưỡng học sinh đạt giải cấp tỉnh.")
    ]

    r_idx = 4
    for it in comp_rows:
        for c_i, v in enumerate(it, start=1):
            ws7.cell(r_idx, c_i, v)
        apply_row_styles(ws7, r_idx, font=FONT_REGULAR)
        ws7.cell(r_idx, 1).alignment = ALIGN_CENTER
        ws7.cell(r_idx, 4).alignment = ALIGN_CENTER
        ws7.cell(r_idx, 5).alignment = ALIGN_CENTER
        ws7.cell(r_idx, 6).alignment = ALIGN_CENTER
        r_idx += 1

    # ==============================================================================
    # SHEET 8: BỘ CHỈ SỐ KPI ĐO LƯỜNG HIỆU QUẢ 3 NHÓM
    # ==============================================================================
    ws8 = wb.create_sheet(title="8. KPI đo lường 3 Nhóm")
    ws8.column_dimensions['A'].width = 8
    ws8.column_dimensions['B'].width = 20
    ws8.column_dimensions['C'].width = 28
    ws8.column_dimensions['D'].width = 34
    ws8.column_dimensions['E'].width = 26
    ws8.column_dimensions['F'].width = 16
    ws8.column_dimensions['G'].width = 14

    ws8.merge_cells("A1:G1")
    ws8["A1"] = "BỘ TIÊU CHÍ KPI ĐO LƯỜNG HIỆU QUẢ CÔNG VIỆC CHO 03 NHÓM VTVL"
    ws8["A1"].font = FONT_TITLE_MAIN
    ws8["A1"].alignment = ALIGN_CENTER

    headers_ws8 = ["STT", "Danh mục VTVL", "Vị trí áp dụng", "Chỉ số KPI chính", "Mục tiêu định lượng", "Tần suất", "Trọng số"]
    ws8.append([])
    ws8.append(headers_ws8)
    for col_idx in range(1, len(headers_ws8) + 1):
        cell = ws8.cell(row=3, column=col_idx)
        cell.font = FONT_HEADER_WHITE
        cell.fill = EMERALD_HEADER
        cell.alignment = ALIGN_CENTER
        cell.border = BORDER_HEADER

    kpi_rows = [
        ("1", "DM1: Quản lý", "Hiệu trưởng", "Tỷ lệ tốt nghiệp THPT toàn trường", "Đạt ≥ 98.5%", "Hàng năm", "25 điểm"),
        ("2", "DM1: Quản lý", "Hiệu trưởng", "Quyết toán ngân sách & Tài chính công", "100% đúng hạn, minh bạch", "Quý / Năm", "20 điểm"),
        ("3", "DM1: Quản lý", "Phó Hiệu trưởng", "Tiến độ chương trình & Ký số học bạ", "100% đúng tiến độ quy định", "Tháng / Kỳ", "25 điểm"),
        ("4", "DM1: Quản lý", "Tổ trưởng CM", "Sinh hoạt NCBL & Đổi mới PPDH tổ", "≥ 02 chuyên đề / học kỳ", "Hàng tháng", "25 điểm"),
        ("5", "DM1: Quản lý", "Tổ phó CM", "Nề nếp sổ sách & Ký số tổ CM", "100% thành viên đúng hạn", "Hàng tháng", "25 điểm"),
        ("6", "DM2: Chuyên môn", "Giáo viên 13 môn", "Tiến độ giảng dạy & Kế hoạch bài dạy", "100% tiết dạy có giáo án chuẩn", "Hàng tuần", "30 điểm"),
        ("7", "DM2: Chuyên môn", "Giáo viên 13 môn", "Chất lượng học sinh môn học phụ trách", "Học sinh đạt chuẩn ≥ 95%", "Học kỳ / Năm", "30 điểm"),
        ("8", "DM2: Chuyên môn", "Giáo viên 13 môn", "Ký số học bạ & Sổ điểm điện tử", "100% đúng hạn, không sai sót", "Cuối kỳ", "20 điểm"),
        ("9", "DM2: Chuyên môn", "Giáo viên 13 môn", "Bồi dưỡng HSG cấp tỉnh / Sáng kiến", "Có giải HSG hoặc SKKN nghiệm thu", "Hàng năm", "20 điểm"),
        ("10", "DM3: Hỗ trợ", "Kế toán viên", "Thanh toán lương & Báo cáo tài chính", "Lương trước ngày 05; 0 sai sót", "Hàng tháng", "30 điểm"),
        ("11", "DM3: Hỗ trợ", "Văn thư", "Xử lý công văn đi/đến & Quản lý con dấu", "100% công văn chuyển trong ngày", "Hàng ngày", "30 điểm"),
        ("12", "DM3: Hỗ trợ", "Thiết bị, thí nghiệm", "Chuẩn bị thiết bị thực hành & An toàn PCCC", "100% tiết thực hành sẵn sàng", "Hàng tuần", "30 điểm"),
        ("13", "DM3: Hỗ trợ", "Giáo vụ", "Sổ đăng bộ & CSDL học sinh số hóa", "Chính xác 100%, đúng tiến độ", "Học kỳ / Năm", "30 điểm"),
        ("14", "DM3: Hỗ trợ", "Y tế trường học", "Sơ cấp cứu, tủ thuốc & Vệ sinh ATTP", "100% an toàn học đường, 0 ngộ độc", "Thường xuyên", "30 điểm"),
        ("15", "DM3: Hỗ trợ", "Thủ quỹ", "Quản lý quỹ tiền mặt & Khớp sổ quỹ", "Khớp 100% số dư với kế toán", "Hàng ngày", "30 điểm")
    ]

    r_idx = 4
    for it in kpi_rows:
        for c_i, v in enumerate(it, start=1):
            ws8.cell(r_idx, c_i, v)
        apply_row_styles(ws8, r_idx, font=FONT_REGULAR)
        ws8.cell(r_idx, 1).alignment = ALIGN_CENTER
        ws8.cell(r_idx, 6).alignment = ALIGN_CENTER
        ws8.cell(r_idx, 7).alignment = ALIGN_CENTER
        r_idx += 1

    # Auto-fit columns for all sheets
    for ws in wb.worksheets:
        auto_fit_columns(ws)

    # Save to all requested paths
    for p in output_paths:
        try:
            d = os.path.dirname(p)
            if d: os.makedirs(d, exist_ok=True)
            wb.save(p)
            print(f"[OK] Đã lưu file thành công: {p}")
        except Exception as e:
            print(f"[ERROR] Không thể lưu file {p}: {e}")

if __name__ == "__main__":
    local_excel_name = "Bang-mo-ta-De-an-VTVL-THPT-Phuc-Hoa-ND232-2026.xlsx"
    local_excel_alt = "Bang-mo-ta-VTVL-TTCM-TPCM-THPT-Phuc-Hoa.xlsx"
    desktop_excel_1 = r"D:\Desktop\Bang-mo-ta-De-an-VTVL-THPT-Phuc-Hoa-ND232-2026.xlsx"
    desktop_excel_2 = r"D:\Desktop\Bang-mo-ta-VTVL-TTCM-TPCM-THPT-Phuc-Hoa.xlsx"
    
    paths = [local_excel_name, local_excel_alt, desktop_excel_1, desktop_excel_2]
    create_full_vtvl_excel(paths)
