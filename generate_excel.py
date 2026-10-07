# -*- coding: utf-8 -*-
"""
Script sinh file Excel: Đề án Vị trí việc làm Viên chức – THPT Phục Hòa
Căn cứ Nghị định số 232/2026/NĐ-CP ngày 26/6/2026 của Chính phủ
Theo Phụ lục I Danh mục VTVL của Trường THPT Phục Hòa (35 biên chế):
1. Vị trí việc làm viên chức quản lý: 07 người (Hiệu trưởng, 2 Phó HT, 2 TTCM, 2 TPCM)
2. Vị trí việc làm viên chức chuyên môn, nghiệp vụ: 22 người (13 môn học trong phụ lục)
3. Vị trí việc làm viên chức hỗ trợ: 06 người (Kế toán, Văn thư, Thiết bị thí nghiệm, Giáo vụ, Y tế học đường, Thủ quỹ)
"""

import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
import os
import shutil
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
    NAVY_FILL = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")       # #1E3A8A
    BLUE_HEADER = PatternFill(start_color="2563EB", end_color="2563EB", fill_type="solid")     # #2563EB
    TEAL_HEADER = PatternFill(start_color="0D9488", end_color="0D9488", fill_type="solid")     # #0D9488
    PURPLE_HEADER = PatternFill(start_color="6D28D9", end_color="6D28D9", fill_type="solid")   # #6D28D9
    AMBER_HEADER = PatternFill(start_color="D97706", end_color="D97706", fill_type="solid")    # #D97706
    SLATE_HEADER = PatternFill(start_color="334155", end_color="334155", fill_type="solid")    # #334155
    EMERALD_HEADER = PatternFill(start_color="059669", end_color="059669", fill_type="solid")  # #059669

    SECTION_FILL = PatternFill(start_color="DBEAFE", end_color="DBEAFE", fill_type="solid")    # #DBEAFE
    SECTION_FILL_GREEN = PatternFill(start_color="CCFBF1", end_color="CCFBF1", fill_type="solid")
    SECTION_FILL_PURPLE = PatternFill(start_color="EDE9FE", end_color="EDE9FE", fill_type="solid")
    SECTION_FILL_AMBER = PatternFill(start_color="FEF3C7", end_color="FEF3C7", fill_type="solid")

    FONT_FAMILY = "Times New Roman"

    FONT_TITLE_MAIN = Font(name=FONT_FAMILY, size=14, bold=True, color="1E3A8A")
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

    def auto_fit_columns(ws, max_cols=None, max_width_limit=65):
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
    # SHEET 1: ĐỀ ÁN & CĂN CỨ PHÁP LÝ
    # ==============================================================================
    ws1 = wb.create_sheet(title="1. Đề án & Căn cứ")
    ws1.column_dimensions['A'].width = 8
    ws1.column_dimensions['B'].width = 28
    ws1.column_dimensions['C'].width = 28
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
    ws1["A5"] = "(Xây dựng theo Nghị định số 232/2026/NĐ-CP ngày 26/6/2026 của Chính phủ & Danh mục VTVL Phụ lục I của Trường)"
    ws1["A5"].font = FONT_ITALIC
    ws1["A5"].alignment = ALIGN_CENTER

    legal_rows = [
        ("I", "CĂN CỨ PHÁP LÝ XÂY DỰNG ĐỀ ÁN", "", "", "", ""),
        ("1", "Nghị định số 232/2026/NĐ-CP", "Chính phủ ban hành 26/6/2026 quy định về vị trí việc làm viên chức", "Chính phủ", "26/06/2026", "Hiệu lực từ 01/7/2026"),
        ("2", "Luật Viên chức số 129/2025/QH15", "Quy định quyền, nghĩa vụ của viên chức, tuyển dụng, sử dụng và quản lý viên chức", "Quốc hội", "2025", "Văn bản nền tảng"),
        ("3", "Thông tư số 20/2023/TT-BGDĐT", "Hướng dẫn về vị trí việc làm, cơ cấu viên chức theo chức danh nghề nghiệp và định mức người làm việc trong các cơ sở GDPT", "Bộ GD&ĐT", "30/10/2023", "Quy định định mức GV/lớp"),
        ("4", "Thông tư số 15/2026/TT-BGDĐT", "Ban hành Điều lệ trường trung học cơ sở, trường trung học phổ thông và trường phổ thông có nhiều cấp học", "Bộ GD&ĐT", "2026", "Quy định quyền hạn HT, TTCM"),
        ("5", "Phụ lục I VTVL Trường THPT Phục Hòa", "Danh mục vị trí việc làm và bậc nghề nghiệp được sử dụng trong từng vị trí của Trường THPT Phục Hòa: 35 biên chế", "THPT Phục Hòa", "10/2026", "Phê duyệt của Sở GD&ĐT"),
        ("II", "NGUYÊN TẮC XÂY DỰNG & QUY MÔ TRƯỜNG LỚP", "", "", "", ""),
        ("1", "Quy mô trường lớp", "Tổng số lớp: 18 lớp THPT (Khối 10: 6 lớp, Khối 11: 6 lớp, Khối 12: 6 lớp). Tổng số học sinh: ~720 học sinh.", "", "", ""),
        ("2", "Tổng số biên chế theo Đề án", "Tổng số biên chế người làm việc: 35 người (đúng theo Phụ lục I Danh mục VTVL của nhà trường).", "", "", ""),
        ("3", "Cơ cấu 3 Nhóm vị trí", "- Nhóm I (Viên chức quản lý): 07 người\n- Nhóm II (Viên chức chuyên môn, nghiệp vụ): 22 người (13 môn học)\n- Nhóm III (Viên chức hỗ trợ): 06 người", "", "", ""),
        ("4", "Nguyên tắc bố trí", "Bảo đảm mỗi vị trí việc làm gắn với sản phẩm đầu ra, thẩm quyền rõ ràng, ứng dụng CNTT & AI nâng cao chất lượng giáo dục toàn diện.", "", "", "")
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
    # SHEET 2: TỔNG HỢP DANH MỤC VTVL CHUẨN PHỤ LỤC I (35 BIÊN CHẾ)
    # ==============================================================================
    ws2 = wb.create_sheet(title="2. Tổng hợp Danh mục Phụ lục I")
    ws2.column_dimensions['A'].width = 8
    ws2.column_dimensions['B'].width = 38
    ws2.column_dimensions['C'].width = 18
    ws2.column_dimensions['D'].width = 22
    ws2.column_dimensions['E'].width = 16
    ws2.column_dimensions['F'].width = 18
    ws2.column_dimensions['G'].width = 35

    ws2.merge_cells("A1:G1")
    ws2["A1"] = "DANH MỤC VỊ TRÍ VIỆC LÀM VÀ BẬC NGHỀ NGHIỆP SỬ DỤNG TRƯỜNG THPT PHỤC HÒA"
    ws2["A1"].font = FONT_TITLE_MAIN
    ws2["A1"].alignment = ALIGN_CENTER

    ws2.merge_cells("A2:G2")
    ws2["A2"] = "(Kèm theo Phụ lục I Danh mục VTVL của Trường THPT Phục Hòa - Căn cứ Nghị định số 232/2026/NĐ-CP)"
    ws2["A2"].font = FONT_ITALIC
    ws2["A2"].alignment = ALIGN_CENTER

    headers_ws2 = ["STT", "Tên vị trí việc làm", "Mã vị trí việc làm", "Bậc nghề nghiệp sử dụng", "Số biên chế", "Phụ cấp / Định mức", "Ghi chú vai trò cốt lõi"]
    ws2.append([]) # Row 3 empty
    ws2.append(headers_ws2) # Row 4
    for col_idx in range(1, len(headers_ws2) + 1):
        cell = ws2.cell(row=4, column=col_idx)
        cell.font = FONT_HEADER_WHITE
        cell.fill = NAVY_FILL
        cell.alignment = ALIGN_CENTER
        cell.border = BORDER_HEADER

    pl1_data = [
        # Nhóm I
        ("I", "Vị trí việc làm viên chức quản lý: 07 người", "", "", "7", "", "Lãnh đạo, quản lý và điều hành trường"),
        ("1", "Hiệu trưởng", "HT-THPT-01", "Bậc 3 đến Bậc 5", "1", "PC: 0.70 | Dạy 2t/t", "Người đứng đầu, quản lý điều hành toàn diện"),
        ("2", "Phó Hiệu trưởng", "PHT-THPT-01", "Bậc 3 đến Bậc 5", "2", "PC: 0.50 | Dạy 4t/t", "Giúp Hiệu trưởng phụ trách chuyên môn, CSVC"),
        ("3", "Tổ trưởng chuyên môn", "TTCM-THPT-01", "Bậc 3 đến Bậc 4", "2", "PC: 0.25 | Dạy 14t/t", "Quản lý sinh hoạt CM tổ, kế hoạch dạy học"),
        ("4", "Tổ phó chuyên môn", "TPCM-THPT-01", "Bậc 3 đến Bậc 4", "2", "PC: 0.15 | Dạy 16t/t", "Giúp TTCM kiểm tra hồ sơ, theo dõi tiến độ"),

        # Nhóm II
        ("II", "Vị trí việc làm viên chức chuyên môn, nghiệp vụ: 22 người", "", "", "22", "", "Giảng dạy các môn học theo CT GDPT 2018"),
        ("1", "Giáo viên THPT môn Ngữ văn", "GV-THPT-01", "Bậc 3 đến Bậc 4", "3", "17 tiết/tuần", "Dạy Ngữ văn 10, 11, 12, ôn thi TN THPT"),
        ("2", "Giáo viên THPT môn Toán", "GV-THPT-01", "Bậc 3 đến Bậc 4", "3", "17 tiết/tuần", "Dạy Toán 10, 11, 12, bồi dưỡng HSG Toán"),
        ("3", "Giáo viên THPT môn Ngoại ngữ (Tiếng Anh)", "GV-THPT-01", "Bậc 3 đến Bậc 4", "2", "17 tiết/tuần", "Dạy Ngoại ngữ, phát triển năng lực giao tiếp"),
        ("4", "Giáo viên THPT môn Lịch sử", "GV-THPT-01", "Bậc 3 đến Bậc 4", "2", "17 tiết/tuần", "Dạy Lịch sử bắt buộc và lựa chọn"),
        ("5", "Giáo viên THPT môn GDTC", "GV-THPT-01", "Bậc 3 đến Bậc 4", "2", "17 tiết/tuần", "Dạy Giáo dục thể chất, HKPĐ, rèn luyện thể lực"),
        ("6", "Giáo viên THPT môn GDQP-AN", "GV-THPT-01", "Bậc 3 đến Bậc 4", "1", "17 tiết/tuần", "Dạy GDQP-AN, hội thao QP-AN, an ninh quốc gia"),
        ("7", "Giáo viên THPT môn Địa lí", "GV-THPT-01", "Bậc 3 đến Bậc 4", "2", "17 tiết/tuần", "Dạy Địa lí, giáo dục bảo vệ tài nguyên môi trường"),
        ("8", "Giáo viên THPT môn GD KT&PL", "GV-THPT-01", "Bậc 3 đến Bậc 4", "1", "17 tiết/tuần", "Dạy Giáo dục kinh tế và pháp luật"),
        ("9", "Giáo viên THPT môn Vật lí", "GV-THPT-01", "Bậc 3 đến Bậc 4", "1", "17 tiết/tuần", "Dạy Vật lí, phòng bộ môn Vật lí, STEM"),
        ("10", "Giáo viên THPT môn Hóa học", "GV-THPT-01", "Bậc 3 đến Bậc 4", "1", "17 tiết/tuần", "Dạy Hóa học, thực hành hóa nghiệm, an toàn PCCC"),
        ("11", "Giáo viên THPT môn Sinh học", "GV-THPT-01", "Bậc 3 đến Bậc 4", "1", "17 tiết/tuần", "Dạy Sinh học, thực hành phòng thí nghiệm Sinh"),
        ("12", "Giáo viên THPT môn Công nghệ", "GV-THPT-01", "Bậc 3 đến Bậc 4", "2", "17 tiết/tuần", "Dạy Công nghệ công nghiệp & Nông nghiệp"),
        ("13", "Giáo viên THPT môn Tin học", "GV-THPT-01", "Bậc 3 đến Bậc 4", "1", "17 tiết/tuần", "Dạy Tin học, phòng máy tính, CĐS & AI"),

        # Nhóm III
        ("III", "Vị trí việc làm viên chức hỗ trợ: 06 người", "", "", "6", "", "Bảo đảm điều kiện hoạt động của trường"),
        ("1", "Kế toán", "KT-THPT-01", "Bậc 1 đến Bậc 4", "1", "40 giờ/tuần", "Quản lý ngân sách, tài chính, thanh toán lương"),
        ("2", "Văn thư", "VT-THPT-01", "Bậc 1 đến Bậc 4", "1", "40 giờ/tuần", "Tiếp nhận, phát hành công văn, quản lý con dấu"),
        ("3", "Thiết bị, thí nghiệm", "TBTN-THPT-01", "Bậc 2 đến Bậc 3", "1", "40 giờ/tuần", "Quản lý thiết bị dạy học, phòng thí nghiệm TN"),
        ("4", "Giáo vụ", "GVU-THPT-01", "Bậc 2 đến Bậc 3", "1", "40 giờ/tuần", "Quản lý hồ sơ học sinh, sổ đăng bộ, tuyển sinh"),
        ("5", "Y tế trường học", "YT-THPT-01", "Bậc 1 đến Bậc 3", "1", "40 giờ/tuần", "Chăm sóc sức khỏe, sơ cấp cứu, phòng chống dịch"),
        ("6", "Thủ quỹ", "TQ-THPT-01", "Bậc 1 đến Bậc 3", "1", "40 giờ/tuần", "Quản lý quỹ tiền mặt, thu chi đúng phiếu chi/thu"),

        # Tổng cộng
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

    note_row = cur_row + 1
    ws2.merge_cells(start_row=note_row, start_column=1, end_row=note_row, end_column=7)
    ws2.cell(note_row, 1, "* Ghi chú: Số lượng VTVL xây dựng đảm bảo tương ứng với số biên chế được cơ quan có thẩm quyền giao theo đúng Phụ lục I của Trường THPT Phục Hòa.")
    ws2.cell(note_row, 1).font = FONT_ITALIC

    # ==============================================================================
    # SHEET 3: DM1 - VTVL QUẢN LÝ (04 VỊ TRÍ, 7 BIÊN CHẾ)
    # ==============================================================================
    ws3 = wb.create_sheet(title="3. DM1 - VTVL Quản lý")
    ws3.column_dimensions['A'].width = 8
    ws3.column_dimensions['B'].width = 24
    ws3.column_dimensions['C'].width = 38
    ws3.column_dimensions['D'].width = 30
    ws3.column_dimensions['E'].width = 12
    ws3.column_dimensions['F'].width = 24

    ws3.merge_cells("A1:F1")
    ws3["A1"] = "BẢNG MÔ TẢ VỊ TRÍ VIỆC LÀM VIÊN CHỨC QUẢN LÝ (07 BIÊN CHẾ)"
    ws3["A1"].font = FONT_TITLE_MAIN
    ws3["A1"].alignment = ALIGN_CENTER

    ws3.merge_cells("A2:F2")
    ws3["A2"] = "(Căn cứ Phụ lục I Nghị định 232/2026/NĐ-CP & Điều lệ trường trung học)"
    ws3["A2"].font = FONT_ITALIC
    ws3["A2"].alignment = ALIGN_CENTER

    headers_ws3 = ["STT", "Vị trí việc làm", "Nhiệm vụ trọng tâm cốt lõi", "Sản phẩm / Kết quả đầu ra", "Tỷ trọng", "Tiêu chuẩn / Phụ cấp"]
    ws3.append([])
    ws3.append(headers_ws3)
    for col_idx in range(1, len(headers_ws3) + 1):
        cell = ws3.cell(row=4, column=col_idx)
        cell.font = FONT_HEADER_WHITE
        cell.fill = BLUE_HEADER
        cell.alignment = ALIGN_CENTER
        cell.border = BORDER_HEADER

    dm1_desc_data = [
        # 1. Hiệu trưởng
        ("1", "Hiệu trưởng (1 người)\nMã: HT-THPT-01\nPC chức vụ: 0.70\nĐịnh mức: 2 tiết/tuần\nBậc: 3 - 5", 
         "1. Xây dựng Chiến lược phát triển trường 2026-2030, kế hoạch giáo dục hàng năm theo CT GDPT 2018.", "Kế hoạch giáo dục được Sở GD&ĐT phê duyệt; 08 quy chế nội bộ ban hành trước 05/9.", "15%", "Đại học SP trở lên, QLGD, CC LLCT"),
        ("", "", "2. Quản lý tổ chức bộ máy, nhân sự; phân công nhiệm vụ, đánh giá xếp loại viên chức.", "Quyết định phân công nhiệm vụ, đánh giá VC theo NĐ 90/2020 & Chuẩn TT 20/2018.", "20%", ""),
        ("", "", "3. Quản lý tài chính, tài sản công, chủ tài khoản trường học.", "Dự toán, quyết toán ngân sách minh bạch, kiểm toán đúng quy định, tiết kiệm chi.", "15%", ""),
        ("", "", "4. Chỉ đạo chuyên môn dạy học, thi TN THPT, thi tuyển sinh 10, kiểm định CLGD.", "Tỷ lệ tốt nghiệp THPT ≥ 98.5%, đạt chuẩn quốc gia mức độ 2, KĐCLGD cấp độ 3.", "20%", ""),
        ("", "", "5. Chuyển đổi số, an toàn trường học, đối ngoại và trực tiếp giảng dạy 2 tiết/tuần.", "100% học bạ số ký đúng hạn; trường học an toàn; thực hiện dạy đủ 2 tiết/tuần.", "30%", ""),

        # 2. Phó Hiệu trưởng
        ("2", "Phó Hiệu trưởng (2 người)\nMã: PHT-THPT-01\nPC chức vụ: 0.50\nĐịnh mức: 4 tiết/tuần\nBậc: 3 - 5", 
         "1. Chỉ đạo trực tiếp công tác chuyên môn dạy học, duyệt kế hoạch bài dạy và phân phối chương trình.", "Thời khóa biểu khoa học, 100% giáo viên dạy đúng tiến độ chương trình 2018.", "25%", "Đại học SP trở lên, QLGD, Trung cấp LLCT"),
        ("", "", "2. Chỉ đạo khảo thí, kiểm tra đánh giá, bồi dưỡng HSG và ôn thi tốt nghiệp THPT.", "Ngân hàng đề kiểm tra định kỳ; tổ chức ôn tập khối 12; vượt chỉ tiêu giải HSG tỉnh.", "20%", ""),
        ("", "", "3. Kiểm tra nội bộ chuyên môn, dự giờ thăm lớp, bồi dưỡng giáo viên dạy giỏi.", "Kiểm tra 100% giáo viên theo kế hoạch năm; dự giờ 1-2 tiết/tuần; tổ chức thi GVDG.", "15%", ""),
        ("", "", "4. Phụ trách cơ sở vật chất, thiết bị dạy học, phòng thí nghiệm, công tác giáo vụ.", "CSVC bảo đảm, thiết bị phục vụ 100% tiết thực hành, quản lý sổ đăng bộ.", "15%", ""),
        ("", "", "5. Chuyển đổi số, ký duyệt học bạ số và trực tiếp giảng dạy 4 tiết/tuần.", "Ký số học bạ đúng hạn 100%; hoàn thành đủ 4 tiết dạy/tuần theo phân công.", "25%", ""),

        # 3. Tổ trưởng chuyên môn
        ("3", "Tổ trưởng chuyên môn (2 người)\nMã: TTCM-THPT-01\nPC chức vụ: 0.25\nĐịnh mức: Dạy 14 tiết/tuần\nBậc: 3 - 4", 
         "1. Xây dựng và tổ chức thực hiện kế hoạch dạy học môn học của tổ theo CT GDPT 2018.", "Kế hoạch giáo dục tổ chuyên môn (PL1, PL2, PL3) phê duyệt trước 30/8.", "25%", "ĐH Sư phạm bộ môn, GVDG cấp trường trở lên"),
        ("", "", "2. Tổ chức sinh hoạt chuyên môn theo nghiên cứu bài học (NCBL) định kỳ ≥ 2 lần/tháng.", "Biên bản sinh hoạt tổ chuyên môn; tối thiểu 02 chuyên đề đổi mới PPDH/học kỳ.", "20%", ""),
        ("", "", "3. Kiểm tra hồ sơ giáo án, dự giờ thành viên trong tổ, bồi dưỡng giáo viên trẻ.", "Dự giờ tối thiểu 2 tiết/giáo viên/học kỳ; 100% thành viên có giáo án đạt chuẩn.", "20%", ""),
        ("", "", "4. Tổ chức bồi dưỡng HSG môn học, phụ đạo học sinh yếu, đổi mới kiểm tra đánh giá.", "Đội tuyển HSG có thành tích giải tỉnh; học sinh yếu được bồi dưỡng kịp thời.", "15%", ""),
        ("", "", "5. Thực hiện giảng dạy 14 tiết/tuần và ứng dụng AI, công nghệ số trong soạn giảng.", "Hoàn thành 14 tiết dạy/tuần; số hóa học liệu, đề thi trắc nghiệm trên hệ thống số.", "20%", ""),

        # 4. Tổ phó chuyên môn
        ("4", "Tổ phó chuyên môn (2 người)\nMã: TPCM-THPT-01\nPC chức vụ: 0.15\nĐịnh mức: Dạy 16 tiết/tuần\nBậc: 3 - 4", 
         "1. Giúp Tổ trưởng theo dõi tiến độ thực hiện chương trình môn học hàng tuần của các thành viên.", "Báo cáo tiến độ chương trình hàng tháng; không có trường hợp dạy chậm/cháy giáo án.", "25%", "ĐH Sư phạm bộ môn, có kinh nghiệm chuyên môn"),
        ("", "", "2. Đôn đốc việc lập lịch báo giảng điện tử, kế hoạch dạy bù, dạy thay của tổ.", "100% giáo viên cập nhật lịch báo giảng trước thứ Hai hàng tuần; bố trí dạy bù kịp thời.", "20%", ""),
        ("", "", "3. Kiểm tra việc sử dụng thiết bị dạy học, đồ dùng trực quan, phòng bộ môn.", "Sổ theo dõi mượn thiết bị và sử dụng phòng thực hành đầy đủ, đúng tiết dạy.", "15%", ""),
        ("", "", "4. Đôn đốc cập nhật điểm số, ký số học bạ điện tử của các thành viên trong tổ.", "100% thành viên vào điểm đúng hạn, ký số học bạ không để sai sót.", "15%", ""),
        ("", "", "5. Thực hiện giảng dạy 16 tiết/tuần và điều hành tổ khi Tổ trưởng vắng mặt.", "Hoàn thành 16 tiết dạy/tuần; xử lý kịp thời công việc đột xuất của tổ theo ủy quyền.", "25%", "")
    ]

    cur_row = 5
    for item in dm1_desc_data:
        for c_idx, val in enumerate(item, start=1):
            ws3.cell(cur_row, c_idx, val)
        apply_row_styles(ws3, cur_row, font=FONT_REGULAR, alignment=ALIGN_LEFT)
        ws3.cell(cur_row, 1).alignment = ALIGN_CENTER
        ws3.cell(cur_row, 5).alignment = ALIGN_CENTER
        cur_row += 1

    # ==============================================================================
    # SHEET 4: DM2 - VTVL CHUYÊN MÔN THEO 13 MÔN HỌC (22 BIÊN CHẾ)
    # ==============================================================================
    ws4 = wb.create_sheet(title="4. DM2 - Chuyên môn (13 Môn)")
    ws4.column_dimensions['A'].width = 6
    ws4.column_dimensions['B'].width = 24
    ws4.column_dimensions['C'].width = 12
    ws4.column_dimensions['D'].width = 14
    ws4.column_dimensions['E'].width = 36
    ws4.column_dimensions['F'].width = 32
    ws4.column_dimensions['G'].width = 28

    ws4.merge_cells("A1:G1")
    ws4["A1"] = "BẢNG MÔ TẢ VỊ TRÍ VIỆC LÀM GIÁO VIÊN THPT THEO 13 MÔN HỌC (22 BIÊN CHẾ)"
    ws4["A1"].font = FONT_TITLE_MAIN
    ws4["A1"].alignment = ALIGN_CENTER

    ws4.merge_cells("A2:G2")
    ws4["A2"] = "(Căn cứ Phụ lục I Danh mục VTVL của Trường THPT Phục Hòa - Chương trình GDPT 2018)"
    ws4["A2"].font = FONT_ITALIC
    ws4["A2"].alignment = ALIGN_CENTER

    headers_ws4 = ["STT", "Tên môn học giảng dạy", "Số biên chế", "Bậc nghề nghiệp", "Đặc thù chuyên môn & Nhiệm vụ giảng dạy (17 tiết/tuần)", "Sản phẩm đầu ra & Chỉ số chất lượng", "Yêu cầu phòng bộ môn / Thiết bị / AI"]
    ws4.append([])
    ws4.append(headers_ws4)
    for col_idx in range(1, len(headers_ws4) + 1):
        cell = ws4.cell(row=4, column=col_idx)
        cell.font = FONT_HEADER_WHITE
        cell.fill = TEAL_HEADER
        cell.alignment = ALIGN_CENTER
        cell.border = BORDER_HEADER

    dm2_subjects = [
        ("1", "Giáo viên Ngữ văn", "3 biên chế", "Bậc 3 đến Bậc 4", 
         "Giảng dạy Ngữ văn 10, 11, 12 theo CT GDPT 2018. Rèn luyện 4 kỹ năng Đọc - Viết - Nói - Nghe; bồi dưỡng tư duy thẩm mỹ, nhân văn; ôn luyện thi tốt nghiệp THPT môn Văn.",
         "100% tiết dạy có giáo án chuẩn; HS đạt chuẩn môn Văn ≥ 98%; phổ điểm thi TN môn Văn ≥ 7.0 điểm; có HSG cấp tỉnh.",
         "Ứng dụng ngữ liệu ngoài SGK, máy chiếu, AI gợi ý dàn ý nghị luận văn học & xã hội."),
        
        ("2", "Giáo viên Toán", "3 biên chế", "Bậc 3 đến Bậc 4",
         "Giảng dạy môn Toán 10, 11, 12 (Đại số, Giải tích, Hình học không gian, Thống kê & Xác suất); phát triển tư duy logic và giải quyết vấn đề toán học; ôn thi TN THPT môn Toán.",
         "Kế hoạch bài dạy có tích hợp liên môn STEM; HS đạt chuẩn môn Toán ≥ 96%; điểm thi tốt nghiệp môn Toán trong top tỉnh; đạt giải HSG tỉnh môn Toán.",
         "Sử dụng phần mềm GeoGebra, máy tính Casio điện tử, AI sinh đề trắc nghiệm toán theo ma trận."),

        ("3", "Giáo viên Ngoại ngữ (Tiếng Anh)", "2 biên chế", "Bậc 3 đến Bậc 4",
         "Giảng dạy Tiếng Anh hệ 10 năm theo CT 2018; phát triển năng lực giao tiếp ngôn ngữ, phản xạ nghe - nói, ngữ pháp và từ vựng; ôn thi tốt nghiệp môn Tiếng Anh.",
         "Học sinh tự tin giao tiếp cơ bản; tỷ lệ học sinh đạt điểm trung bình trở lên thi TN môn Anh tăng ≥ 5%; có học sinh đạt chứng chỉ chuẩn quốc tế/tỉnh.",
         "Phòng học Ngoại ngữ tương tác đa phương tiện, loa nghe chất lượng cao, ứng dụng app phát âm AI."),

        ("4", "Giáo viên Lịch sử", "2 biên chế", "Bậc 3 đến Bậc 4",
         "Giảng dạy Lịch sử bắt buộc và lựa chọn; giáo dục lòng yêu nước, truyền thống lịch sử vẻ vang của dân tộc và lịch sử thế giới văn minh; ôn thi tốt nghiệp THPT môn Lịch sử.",
         "100% bài dạy có tranh ảnh, lược đồ tư liệu lịch sử; điểm thi tốt nghiệp môn Sử đạt ≥ 98% trên trung bình; có giải HSG môn Lịch sử cấp tỉnh.",
         "Bản đồ số lịch sử, video tư liệu, bảo tàng ảo 3D, ứng dụng AI tóm tắt dòng thời gian sự kiện."),

        ("5", "Giáo viên Giáo dục thể chất (GDTC)", "2 biên chế", "Bậc 3 đến Bậc 4",
         "Giảng dạy GDTC khối 10, 11, 12 (Điền kinh, Bóng chuyền, Bóng đá, Cầu lông...); rèn luyện sức khỏe, thói quen vận động thể thao lành mạnh; huấn luyện Hội khỏe Phù Đổng.",
         "100% học sinh đạt tiêu chuẩn rèn luyện thân thể; đoàn VĐV nhà trường đạt thành tích cao tại HKPĐ cấp huyện và tỉnh; không để xảy ra chấn thương.",
         "Sân vận động, nhà tập đa năng, dụng cụ đo thể lực, trang thiết bị sơ cứu chấn thương thể thao."),

        ("6", "Giáo viên GDQP-AN", "1 biên chế", "Bậc 3 đến Bậc 4",
         "Giảng dạy môn Giáo dục quốc phòng và an ninh; huấn luyện điều lệnh đội ngũ, bắn súng tiểu liên AK, chiến thuật bộ binh cơ bản; giáo dục an ninh chủ quyền biên giới quốc gia.",
         "100% học sinh nắm vững kiến thức quốc phòng, thực hiện chuẩn điều lệnh; tham gia Hội thao GDQP-AN các cấp đạt kết quả cao; bảo đảm tuyệt đối an toàn vũ khí huấn luyện.",
         "Bãi tập điều lệnh, súng mô hình tiểu liên AK, lựu đạn tập, thiết bị bắn súng laser điện tử."),

        ("7", "Giáo viên Địa lí", "2 biên chế", "Bậc 3 đến Bậc 4",
         "Giảng dạy Địa lí tự nhiên, kinh tế - xã hội Việt Nam và thế giới; khai thác bản đồ, biểu đồ số liệu, Atlat Địa lí Việt Nam; giáo dục bảo vệ chủ quyền biển đảo và biến đổi khí hậu.",
         "Học sinh sử dụng thành thạo Atlat và đọc bản đồ; bài kiểm tra định kỳ phân tích biểu đồ tốt; điểm thi TN môn Địa lí đạt kết quả cao; bồi dưỡng HSG tỉnh.",
         "Hệ thống Atlat số, Google Earth, mô hình địa bàn, sa bàn địa hình Cao Bằng và Việt Nam."),

        ("8", "Giáo viên GD KT&PL", "1 biên chế", "Bậc 3 đến Bậc 4",
         "Giảng dạy Giáo dục Kinh tế và Pháp luật; trang bị kiến thức tài chính, doanh nghiệp cơ bản, pháp luật dân sự, hình sự, hôn nhân gia đình và ý thức công dân tuân thủ pháp luật.",
         "Học sinh hiểu và tôn trọng pháp luật, không có học sinh vi phạm kỷ luật/pháp luật; tổ chức các phiên tòa giả định; học sinh đạt chuẩn môn học 100%.",
         "Văn bản quy phạm pháp luật số hóa, tình huống pháp lý thực tiễn, video phiên tòa xét xử thực tế."),

        ("9", "Giáo viên Vật lí", "1 biên chế", "Bậc 3 đến Bậc 4",
         "Giảng dạy Vật lí 10, 11, 12 (Cơ, Nhiệt, Điện, Quang, Vật lí hạt nhân); tổ chức thí nghiệm thực hành, giải thích các hiện tượng tự nhiên; giáo dục STEM/STEAM mô hình hóa.",
         "100% tiết thực hành Vật lí được thực hiện tại phòng bộ môn; sản phẩm STEM học sinh chế tạo; kết quả thi TN tổ hợp KHTN đạt chuẩn; bồi dưỡng HSG tỉnh môn Lý.",
         "Phòng thí nghiệm Vật lí đạt chuẩn, cảm biến thông minh, phần mềm mô phỏng vật lý PhET, thiết bị STEM."),

        ("10", "Giáo viên Hóa học", "1 biên chế", "Bậc 3 đến Bậc 4",
         "Giảng dạy Hóa học vô cơ và hữu cơ; trực tiếp tổ chức thí nghiệm nghiên cứu phản ứng hóa học; giáo dục an toàn hóa chất, bảo vệ môi trường và ứng dụng hóa học đời sống.",
         "100% bài thực hành an toàn tuyệt đối; học sinh nắm vững phương trình hóa học và kỹ năng thực hành ống nghiệm; kết quả thi KHTN môn Hóa đạt tỷ lệ cao.",
         "Phòng thí nghiệm Hóa học có tủ hút khí độc, hệ thống cấp thoát nước, bình xịt chữa cháy, hóa chất chuẩn."),

        ("11", "Giáo viên Sinh học", "1 biên chế", "Bậc 3 đến Bậc 4",
         "Giảng dạy Sinh học tế bào, di truyền, tiến hóa và sinh thái học; thực hành quan sát kính hiển vi, tiêu bản tế bào; giáo dục bảo tồn đa dạng sinh học và sức khỏe sinh sản.",
         "Tổ chức thành công các tiết thực hành làm tiêu bản hiển vi; học sinh hiểu sâu quy luật di truyền; đạt giải thi HSG Sinh học cấp tỉnh; tham gia thi KHKT học sinh.",
         "Kính hiển vi quang học, tiêu bản mẫu động - thực vật, vườn thực nghiệm sinh học trường học."),

        ("12", "Giáo viên Công nghệ", "2 biên chế", "Bậc 3 đến Bậc 4",
         "Giảng dạy Công nghệ định hướng Công nghiệp (Bản vẽ kỹ thuật, Cơ khí, Điện) & Nông nghiệp (Trồng trọt, Chăn nuôi, Thủy sản công nghệ cao); hướng nghiệp việc làm cho học sinh.",
         "Học sinh đọc được bản vẽ kỹ thuật cơ bản, đấu nối mạch điện an toàn, nắm kỹ thuật trồng trọt nông nghiệp sạch; có dự án sáng tạo kỹ thuật của học sinh.",
         "Xưởng thực hành công nghệ, mô hình động cơ, thiết bị điện an toàn, mô hình nông nghiệp thông minh."),

        ("13", "Giáo viên Tin học", "1 biên chế", "Bậc 3 đến Bậc 4",
         "Giảng dạy Tin học 10, 11, 12 (Lập trình Python, Khoa học dữ liệu, Mạng máy tính & An ninh số); quản lý phòng máy tính; hỗ trợ chuyển đổi số và khai thác AI toàn trường.",
         "Học sinh lập trình cơ bản thành thạo; tỷ lệ học sinh giỏi tin học tăng; hỗ trợ 100% giáo viên ứng dụng CNTT và bảo đảm mạng máy tính hoạt động thông suốt.",
         "02 Phòng máy vi tính kết nối Internet cáp quang, máy chiếu, bảng tương tác, phần mềm lập trình Python/IDE.")
    ]

    cur_row = 5
    for item in dm2_subjects:
        for c_idx, val in enumerate(item, start=1):
            ws4.cell(cur_row, c_idx, val)
        apply_row_styles(ws4, cur_row, font=FONT_REGULAR, alignment=ALIGN_LEFT)
        ws4.cell(cur_row, 1).alignment = ALIGN_CENTER
        ws4.cell(cur_row, 3).alignment = ALIGN_CENTER
        ws4.cell(cur_row, 4).alignment = ALIGN_CENTER
        cur_row += 1

    # ==============================================================================
    # SHEET 5: DM3 - VTVL HỖ TRỢ (06 VỊ TRÍ, 6 BIÊN CHẾ)
    # ==============================================================================
    ws5 = wb.create_sheet(title="5. DM3 - VTVL Hỗ trợ (6 Vị trí)")
    ws5.column_dimensions['A'].width = 8
    ws5.column_dimensions['B'].width = 24
    ws5.column_dimensions['C'].width = 16
    ws5.column_dimensions['D'].width = 16
    ws5.column_dimensions['E'].width = 38
    ws5.column_dimensions['F'].width = 30
    ws5.column_dimensions['G'].width = 25

    ws5.merge_cells("A1:G1")
    ws5["A1"] = "BẢNG MÔ TẢ VỊ TRÍ VIỆC LÀM VIÊN CHỨC HỖ TRỢ (06 NGƯỜI)"
    ws5["A1"].font = FONT_TITLE_MAIN
    ws5["A1"].alignment = ALIGN_CENTER

    ws5.merge_cells("A2:G2")
    ws5["A2"] = "(Căn cứ Phụ lục I Danh mục VTVL của Trường THPT Phục Hòa - Nghị định 232/2026/NĐ-CP)"
    ws5["A2"].font = FONT_ITALIC
    ws5["A2"].alignment = ALIGN_CENTER

    headers_ws5 = ["STT", "Vị trí việc làm", "Mã vị trí", "Bậc nghề nghiệp", "Mô tả công việc & Nhiệm vụ cốt lõi (40 giờ/tuần)", "Sản phẩm / Kết quả đầu ra", "Tiêu chuẩn chuyên môn"]
    ws5.append([])
    ws5.append(headers_ws5)
    for col_idx in range(1, len(headers_ws5) + 1):
        cell = ws5.cell(row=4, column=col_idx)
        cell.font = FONT_HEADER_WHITE
        cell.fill = AMBER_HEADER
        cell.alignment = ALIGN_CENTER
        cell.border = BORDER_HEADER

    dm3_support_data = [
        ("1", "Kế toán\n(01 biên chế)", "KT-THPT-01", "Bậc 1 đến Bậc 4",
         "1. Lập dự toán ngân sách nhà nước đầu năm, dự toán thu - chi học phí và các nguồn thu hợp pháp.\n2. Thực hiện tính toán chi trả lương, phụ cấp, BHXH, chế độ thừa giờ, nâng bậc lương đúng hạn.\n3. Quản lý hệ thống sổ sách kế toán, chứng từ thanh quyết toán qua Kho bạc Nhà nước.\n4. Quyết toán tài chính quý, năm với Sở GD&ĐT và cơ quan tài chính; lưu trữ chứng từ.",
         "Dự toán được phê duyệt; bảng lương hàng tháng trước ngày 05; báo cáo tài chính năm đúng hạn 100%, không sai sót kiểm toán.",
         "Đại học chuyên ngành Tài chính - Kế toán trở lên; chứng chỉ kế toán viên; làm chủ phần mềm MISA Mimosa."),

        ("2", "Văn thư\n(01 biên chế)", "VT-THPT-01", "Bậc 1 đến Bậc 4",
         "1. Tiếp nhận, đăng ký, chuyển giao công văn đi - đến trên phần mềm Quản lý văn bản điện tử và điều hành (iOffice).\n2. Soạn thảo, định dạng văn bản theo Nghị định 30/2020/NĐ-CP; quản lý và đóng dấu cơ quan, chứng thư chữ ký số đúng luật.\n3. Lập hồ sơ lưu trữ hiện hành, bảo quản tài liệu lưu trữ vĩnh viễn và có thời hạn.\n4. Cấp giấy giới thiệu, in ấn tài liệu phục vụ các cuộc họp, hội nghị của nhà trường.",
         "Công văn đến được xử lý trong ngày 100%; không thất lạc văn bản; quản lý con dấu an toàn tuyệt đối; hồ sơ lưu trữ khoa học.",
         "Trung cấp trở lên ngành Văn thư - Lưu trữ hoặc tương đương; thành thạo iOffice và chữ ký số."),

        ("3", "Thiết bị, thí nghiệm\n(01 biên chế)", "TBTN-THPT-01", "Bậc 2 đến Bậc 3",
         "1. Quản lý, bảo quản toàn bộ thiết bị dạy học, máy móc thí nghiệm, hóa chất phòng thực hành Lý, Hóa, Sinh, Công nghệ.\n2. Lập kế hoạch mua sắm, sửa chữa, bổ sung thiết bị dạy học đầu năm học.\n3. Chuẩn bị đầy đủ dụng cụ, mẫu vật, hóa chất cho các tiết thực hành của giáo viên bộ môn.\n4. Lập sổ theo dõi mượn - trả TBDH; kiểm kê tài sản định kỳ; bảo đảm an toàn cháy nổ PCCC.",
         "100% tiết thực hành có thiết bị chuẩn bị sẵn sàng; không xảy ra sự cố cháy nổ/tai nạn; sổ sách TBDH cập nhật hàng tuần.",
         "Cao đẳng trở lên chuyên ngành Thiết bị dạy học hoặc các ngành Lý, Hóa, Sinh; chứng chỉ an toàn hóa chất."),

        ("4", "Giáo vụ\n(01 biên chế)", "GVU-THPT-01", "Bậc 2 đến Bậc 3",
         "1. Quản lý sổ đăng bộ của nhà trường, hồ sơ trúng tuyển học sinh lớp 10, hồ sơ học sinh chuyển đi, chuyển đến.\n2. Quản lý và theo dõi học bạ số, hồ sơ đề nghị cấp phát bằng tốt nghiệp THPT.\n3. Phối hợp lập danh sách thí sinh dự thi tốt nghiệp THPT, kỳ thi học sinh giỏi các cấp.\n4. Chuẩn bị văn phòng phẩm, ấn chỉ thi, phiếu trả lời trắc nghiệm phục vụ các kỳ kiểm tra định kỳ.",
         "Sổ đăng bộ cập nhật chính xác 100%; hồ sơ tuyển sinh và hồ sơ thi tốt nghiệp nộp đúng hạn về Sở; cấp phát bằng tốt nghiệp không sai sót.",
         "Trung cấp trở lên ngành Quản lý giáo dục, Tin học hoặc tương đương; thành thạo cơ sở dữ liệu ngành GDĐT."),

        ("5", "Y tế trường học\n(01 biên chế)", "YT-THPT-01", "Bậc 1 đến Bậc 3",
         "1. Chăm sóc sức khỏe ban đầu, sơ cứu kịp thời tai nạn thương tích học sinh và cán bộ giáo viên.\n2. Lập hồ sơ theo dõi sức khỏe học sinh; phối hợp với TTYT tổ chức khám sức khỏe định kỳ cho học sinh toàn trường.\n3. Tuyên truyền phòng chống dịch bệnh học đường (sốt xuất huyết, cúm, đau mắt đỏ...), nha học đường, tật khúc xạ.\n4. Kiểm tra vệ sinh môi trường học đường, an toàn nguồn nước uống, an toàn thực phẩm căng tin trường.",
         "Tủ thuốc y tế đầy đủ thuốc thiết yếu; 100% học sinh được theo dõi sức khỏe; không xảy ra ngộ độc thực phẩm hay dịch bệnh bùng phát.",
         "Trung cấp Y sĩ hoặc Điều dưỡng trở lên; có chứng chỉ sơ cấp cứu ban đầu; nắm vững quy chuẩn vệ sinh học đường."),

        ("6", "Thủ quỹ\n(01 biên chế)", "TQ-THPT-01", "Bậc 1 đến Bậc 3",
         "1. Trực tiếp quản lý két sắt và quỹ tiền mặt của nhà trường; bảo đảm an toàn tuyệt đối quỹ tiền mặt.\n2. Thực hiện thu, chi tiền mặt đúng quy định, chỉ thu - chi khi có phiếu thu, phiếu chi hợp lệ có đầy đủ chữ ký của Hiệu trưởng và Kế toán.\n3. Hàng ngày ghi chép sổ quỹ tiền mặt đầy đủ, rõ ràng; đối chiếu khớp số dư sổ quỹ với sổ kế toán tiền mặt.\n4. Tham gia kiểm kê quỹ tiền mặt đột xuất và định kỳ cuối tháng, quý, năm.",
         "Số dư quỹ tiền mặt thực tế khớp 100% với sổ quỹ và sổ kế toán; không thiếu hụt, không thất thoát; thanh toán tiền mặt kịp thời, minh bạch.",
         "Trung cấp trở lên ngành Tài chính - Kế toán, Ngân hàng hoặc Kinh tế; phẩm chất đạo đức trung thực, cẩn trọng tuyệt đối.")
    ]

    cur_row = 5
    for item in dm3_support_data:
        for c_idx, val in enumerate(item, start=1):
            ws5.cell(cur_row, c_idx, val)
        apply_row_styles(ws5, cur_row, font=FONT_REGULAR, alignment=ALIGN_LEFT)
        ws5.cell(cur_row, 1).alignment = ALIGN_CENTER
        ws5.cell(cur_row, 3).alignment = ALIGN_CENTER
        ws5.cell(cur_row, 4).alignment = ALIGN_CENTER
        cur_row += 1

    # ==============================================================================
    # SHEET 6: KHUNG NĂNG LỰC 5 CẤP ĐỘ THEO NĐ 232/2026
    # ==============================================================================
    ws6 = wb.create_sheet(title="6. Khung năng lực 5 Cấp độ")
    ws6.column_dimensions['A'].width = 6
    ws6.column_dimensions['B'].width = 22
    ws6.column_dimensions['C'].width = 30
    ws6.column_dimensions['D'].width = 18
    ws6.column_dimensions['E'].width = 18
    ws6.column_dimensions['F'].width = 18
    ws6.column_dimensions['G'].width = 38

    ws6.merge_cells("A1:G1")
    ws6["A1"] = "KHUNG NĂNG LỰC TOÀN DIỆN CỦA 03 DANH MỤC VTVL (5 CẤP ĐỘ)"
    ws6["A1"].font = FONT_TITLE_MAIN
    ws6["A1"].alignment = ALIGN_CENTER

    ws6.merge_cells("A2:G2")
    ws6["A2"] = "(Căn cứ Điều 8 & Phụ lục IV Nghị định số 232/2026/NĐ-CP của Chính phủ)"
    ws6["A2"].font = FONT_ITALIC
    ws6["A2"].alignment = ALIGN_CENTER

    headers_ws6 = ["STT", "Nhóm năng lực", "Tên năng lực thành phần", "Yêu cầu Nhóm 1\n(Viên chức quản lý)", "Yêu cầu Nhóm 2\n(13 Môn học)", "Yêu cầu Nhóm 3\n(6 VC Hỗ trợ)", "Mô tả chuẩn hành vi & Tiêu chí đạt"]
    ws6.append([])
    ws6.append(headers_ws6)
    for col_idx in range(1, len(headers_ws6) + 1):
        cell = ws6.cell(row=4, column=col_idx)
        cell.font = FONT_HEADER_WHITE
        cell.fill = PURPLE_HEADER
        cell.alignment = ALIGN_CENTER
        cell.border = BORDER_HEADER

    comp_rows = [
        ("1", "Năng lực chung", "Đạo đức nghề nghiệp & Liêm chính sư phạm", "Cấp độ 5", "Cấp độ 4 - 5", "Cấp độ 3 - 4", "Tuyệt đối gương mẫu, không vụ lợi, chấp hành nghiêm quy tắc ứng xử trường học."),
        ("2", "Năng lực chung", "Kỷ luật, trách nhiệm & Tinh thần phụng sự", "Cấp độ 5", "Cấp độ 4", "Cấp độ 3 - 4", "Tuân thủ kỷ luật lao động, hoàn thành công việc được giao đúng hạn và chất lượng."),
        ("3", "Năng lực chung", "Kỹ năng giao tiếp, ứng xử & Phối hợp công tác", "Cấp độ 4 - 5", "Cấp độ 3 - 4", "Cấp độ 3", "Ứng xử chuẩn mực với phụ huynh, học sinh, đồng nghiệp; hợp tác hiệu quả."),
        ("4", "Năng lực chung", "Ứng dụng CNTT, CĐS & Sử dụng AI", "Cấp độ 4", "Cấp độ 3 - 5", "Cấp độ 2 - 3", "Khai thác phần mềm ngành, ký số học bạ, sử dụng AI hỗ trợ công việc chuyên môn."),
        ("5", "Năng lực quản lý", "Tư duy chiến lược & Lập kế hoạch", "Cấp độ 4 - 5", "Cấp độ 2 - 3", "Cấp độ 1 - 2", "Xây dựng tầm nhìn, phân bổ nguồn lực khoa học, dự báo tình huống quản trị."),
        ("6", "Năng lực quản lý", "Tổ chức điều hành & Kiểm tra giám sát", "Cấp độ 4 - 5", "Cấp độ 2 - 3", "Cấp độ 1 - 2", "Phân công nhiệm vụ hợp lý, kiểm tra đôn đốc tiến độ, xử lý công bằng khách quan."),
        ("7", "Năng lực chuyên môn", "Nghiệp vụ sư phạm môn học / Nghiệp vụ chuyên ngành", "Cấp độ 4 - 5", "Cấp độ 3 - 5", "Cấp độ 3 - 4", "Nắm vững kiến thức 13 môn học GDPT 2018 hoặc chuyên ngành tài chính/văn thư/y tế."),
        ("8", "Năng lực chuyên môn", "Phương pháp dạy học tích cực / Kỹ thuật nghiệp vụ", "Cấp độ 4 - 5", "Cấp độ 4 - 5", "Cấp độ 2 - 3", "Vận dụng PPDH tích cực, giáo dục STEM, kỹ thuật ghi chép và quy trình kiểm soát."),
        ("9", "Năng lực chuyên môn", "Kiểm tra đánh giá & Phân tích chất lượng", "Cấp độ 4 - 5", "Cấp độ 3 - 5", "Cấp độ 2 - 3", "Đánh giá đúng năng lực học sinh, xây dựng ma trận đề chuẩn, phân tích dữ liệu."),
        ("10", "Năng lực bổ trợ", "Xử lý tình huống sư phạm & Giải quyết tranh chấp", "Cấp độ 4 - 5", "Cấp độ 3 - 4", "Cấp độ 2 - 3", "Điềm tĩnh, thấu tình đạt lý, giải quyết xung đột trong môi trường học đường nhân văn.")
    ]

    cur_row = 5
    for item in comp_rows:
        for c_idx, val in enumerate(item, start=1):
            ws6.cell(cur_row, c_idx, val)
        apply_row_styles(ws6, cur_row, font=FONT_REGULAR, alignment=ALIGN_LEFT)
        ws6.cell(cur_row, 1).alignment = ALIGN_CENTER
        ws6.cell(cur_row, 4).alignment = ALIGN_CENTER
        ws6.cell(cur_row, 5).alignment = ALIGN_CENTER
        ws6.cell(cur_row, 6).alignment = ALIGN_CENTER
        cur_row += 1

    # ==============================================================================
    # SHEET 7: KPI ĐO LƯỜNG HIỆU QUẢ ĐẦU RA CHO 35 BIÊN CHẾ
    # ==============================================================================
    ws7 = wb.create_sheet(title="7. KPI đo lường đầu ra")
    ws7.column_dimensions['A'].width = 6
    ws7.column_dimensions['B'].width = 20
    ws7.column_dimensions['C'].width = 24
    ws7.column_dimensions['D'].width = 30
    ws7.column_dimensions['E'].width = 25
    ws7.column_dimensions['F'].width = 14
    ws7.column_dimensions['G'].width = 12
    ws7.column_dimensions['H'].width = 32

    ws7.merge_cells("A1:H1")
    ws7["A1"] = "BỘ TIÊU CHÍ KPI ĐO LƯỜNG KẾT QUẢ ĐẦU RA 03 NHÓM VTVL (35 BIÊN CHẾ)"
    ws7["A1"].font = FONT_TITLE_MAIN
    ws7["A1"].alignment = ALIGN_CENTER

    ws7.merge_cells("A2:H2")
    ws7["A2"] = "(Căn cứ Điều 9 Nghị định số 232/2026/NĐ-CP ngày 26/6/2026 của Chính phủ)"
    ws7["A2"].font = FONT_ITALIC
    ws7["A2"].alignment = ALIGN_CENTER

    headers_ws7 = ["STT", "Danh mục VTVL", "Vị trí việc làm áp dụng", "Chỉ số KPI chính", "Mục tiêu định lượng", "Tần suất", "Trọng số", "Mức hoàn thành xuất sắc"]
    ws7.append([])
    ws7.append(headers_ws7)
    for col_idx in range(1, len(headers_ws7) + 1):
        cell = ws7.cell(row=4, column=col_idx)
        cell.font = FONT_HEADER_WHITE
        cell.fill = EMERALD_HEADER
        cell.alignment = ALIGN_CENTER
        cell.border = BORDER_HEADER

    kpi_rows = [
        ("1", "I. Quản lý (7 người)", "Hiệu trưởng", "Tỷ lệ tốt nghiệp THPT toàn trường", "Đạt ≥ 98.5%", "Hàng năm", "25%", "Đạt ≥ 99.5%, phổ điểm trong top đầu tỉnh"),
        ("2", "I. Quản lý (7 người)", "Hiệu trưởng", "Quyết toán ngân sách & Tài chính", "100% đúng luật, không vi phạm", "Quý / Năm", "20%", "Tiết kiệm chi ngân sách, kiểm toán minh bạch"),
        ("3", "I. Quản lý (7 người)", "Phó Hiệu trưởng", "Tiến độ chương trình & Ký số học bạ", "100% giáo viên đúng tiến độ", "Tháng / Kỳ", "25%", "Học bạ số ký trước thời hạn quy định 2 ngày"),
        ("4", "I. Quản lý (7 người)", "Tổ trưởng CM", "Sinh hoạt NCBL & Đổi mới PPDH", "≥ 02 chuyên đề / học kỳ", "Hàng tháng", "25%", "Có chuyên đề STEM đột phá cấp trường"),
        ("5", "I. Quản lý (7 người)", "Tổ phó CM", "Lịch báo giảng & Sổ sách điện tử", "100% thành viên đúng hạn", "Hàng tuần", "25%", "Không có giáo viên vi phạm nộp trễ"),
        
        ("6", "II. Chuyên môn (22 GV)", "Giáo viên 13 môn học", "Kế hoạch bài dạy & Giảng dạy 17t/t", "100% tiết có giáo án chuẩn", "Hàng tuần", "30%", "Soạn giảng ứng dụng AI, có tiết dạy STEM"),
        ("7", "II. Chuyên môn (22 GV)", "Giáo viên 13 môn học", "Chất lượng môn học phụ trách", "Học sinh đạt chuẩn ≥ 96%", "Học kỳ / Năm", "30%", "Tăng tỷ lệ điểm Khá, Giỏi ≥ 10%"),
        ("8", "II. Chuyên môn (22 GV)", "Giáo viên 13 môn học", "Bồi dưỡng HSG & Phụ đạo yếu kém", "Có học sinh đạt giải cấp tỉnh", "Hàng năm", "20%", "Đạt từ 02 giải HSG tỉnh trở lên"),
        ("9", "II. Chuyên môn (22 GV)", "Giáo viên Tin học/Lý/Hóa/Sinh", "Phòng bộ môn, phòng máy & CĐS", "Hoạt động an toàn, thông suốt", "Hàng tháng", "20%", "100% tiết thực hành bảo đảm an toàn"),

        ("10", "III. Hỗ trợ (6 người)", "Kế toán viên", "Chi trả lương & Quyết toán chứng từ", "100% đúng hạn, chính xác tuyệt đối", "Hàng tháng", "35%", "Quyết toán không phát sinh tồn đọng, minh bạch"),
        ("11", "III. Hỗ trợ (6 người)", "Văn thư", "Xử lý công văn đi đến & Con dấu số", "Xử lý trong ngày 100%", "Hàng ngày", "35%", "Không thất lạc văn bản, số hóa tài liệu 100%"),
        ("12", "III. Hỗ trợ (6 người)", "Thiết bị, TN", "Chuẩn bị thiết bị cho tiết thực hành", "100% tiết thực hành đủ đồ dùng", "Hàng tuần", "35%", "Không xảy ra tai nạn thí nghiệm, kiểm kê tốt"),
        ("13", "III. Hỗ trợ (6 người)", "Giáo vụ", "Quản lý hồ sơ học sinh & Tuyển sinh", "100% hồ sơ chính xác, đúng hạn", "Hàng kỳ", "35%", "Cấp phát văn bằng, học bạ chính xác tuyệt đối"),
        ("14", "III. Hỗ trợ (6 người)", "Y tế học đường", "Chăm sóc sức khỏe & Vệ sinh an toàn", "100% học sinh được theo dõi SK", "Hàng kỳ", "35%", "Không bùng phát dịch bệnh, sơ cứu kịp thời 100%"),
        ("15", "III. Hỗ trợ (6 người)", "Thủ quỹ", "Quản lý quỹ tiền mặt & Thu chi hợp lệ", "Khớp số dư 100% với kế toán", "Hàng ngày", "35%", "Không chênh lệch quỹ, an toàn tuyệt đối")
    ]

    cur_row = 5
    for item in kpi_rows:
        for c_idx, val in enumerate(item, start=1):
            ws7.cell(cur_row, c_idx, val)
        apply_row_styles(ws7, cur_row, font=FONT_REGULAR, alignment=ALIGN_LEFT)
        ws7.cell(cur_row, 1).alignment = ALIGN_CENTER
        ws7.cell(cur_row, 6).alignment = ALIGN_CENTER
        ws7.cell(cur_row, 7).alignment = ALIGN_CENTER
        cur_row += 1

    # Auto-fit columns for all sheets
    for ws in wb.worksheets:
        auto_fit_columns(ws)

    # Save to all target paths
    for p in output_paths:
        target_dir = os.path.dirname(p)
        if target_dir and not os.path.exists(target_dir):
            os.makedirs(target_dir, exist_ok=True)
        wb.save(p)
        print(f"✓ Đã lưu bảng Excel Đề án thành công tại: {p}")

if __name__ == "__main__":
    paths = [
        r"d:\Du-an-web\web-vtvl-ttcm-tpcm-phuc-hoa\Bang-mo-ta-De-an-VTVL-THPT-Phuc-Hoa-ND232-2026.xlsx",
        r"D:\Desktop\Bang-mo-ta-De-an-VTVL-THPT-Phuc-Hoa-ND232-2026.xlsx"
    ]
    create_full_vtvl_excel(paths)
