# -*- coding: utf-8 -*-
"""
Script sinh file Excel: Đề án Vị trí việc làm Viên chức – THPT Phục Hòa
Căn cứ Nghị định số 232/2026/NĐ-CP ngày 26/6/2026 của Chính phủ
Áp dụng ĐÚNG GIỐNG MẪU BẢN MÔ TẢ CÔNG VIỆC VÀ KHUNG NĂNG LỰC (Phần B Phụ lục IV/VI NĐ 232/2026/NĐ-CP)
Trọn bộ 35 Biên chế theo Phụ lục I C3 Phục Hòa:
1. Vị trí việc làm Quản lý: 07 người
2. Vị trí việc làm Chuyên môn, nghiệp vụ: 22 người (13 môn học trong phụ lục)
3. Vị trí việc làm Hỗ trợ: 06 người (Kế toán, Văn thư, Thiết bị TN, Giáo vụ, Y tế học đường, Thủ quỹ)
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
    # SHEET 1: ĐỀ ÁN & CĂN CỨ
    # ==============================================================================
    ws1 = wb.create_sheet(title="1. Đề án & Căn cứ")
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
    ws1["A5"] = "(Xây dựng theo Nghị định 232/2026/NĐ-CP & Mẫu bản mô tả VTVL chuẩn - Tổng 35 biên chế)"
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
        ("3", "Cơ cấu 3 Nhóm vị trí", "- Nhóm I (Viên chức quản lý): 07 người\n- Nhóm II (Viên chức chuyên môn, nghiệp vụ): 22 người (13 môn học)\n- Nhóm III (Viên chức hỗ trợ): 06 người", "", "", ""),
        ("4", "Chuẩn hóa theo Mẫu NĐ 232", "Toàn bộ bản mô tả VTVL được lập đúng theo Mẫu tại Phần B Phụ lục IV/VI của Nghị định số 232/2026/NĐ-CP.", "", "", "")
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
    # SHEET 2: TỔNG HỢP DANH MỤC PHỤ LỤC I (35 BIÊN CHẾ)
    # ==============================================================================
    ws2 = wb.create_sheet(title="2. Tổng hợp Phụ lục I (35 BC)")
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

        ("II", "Vị trí việc làm viên chức chuyên môn, nghiệp vụ: 22 người", "", "", "22", "", "Giảng dạy 13 môn học theo CT GDPT 2018"),
        ("1", "Giáo viên THPT môn Ngữ văn", "GV-THPT-01", "Bậc 3 đến Bậc 4", "3", "17 tiết/tuần", "Dạy Ngữ văn 10, 11, 12, ôn thi TN THPT"),
        ("2", "Giáo viên THPT môn Toán", "GV-THPT-01", "Bậc 3 đến Bậc 4", "3", "17 tiết/tuần", "Dạy Toán 10, 11, 12, bồi dưỡng HSG Toán"),
        ("3", "Giáo viên THPT môn Ngoại ngữ (Tiếng Anh)", "GV-THPT-01", "Bậc 3 đến Bậc 4", "2", "17 tiết/tuần", "Dạy Ngoại ngữ, phát triển năng lực giao tiếp"),
        ("4", "Giáo viên THPT môn Lịch sử", "GV-THPT-01", "Bậc 3 đến Bậc 4", "2", "17 tiết/tuần", "Dạy Lịch sử bắt buộc và lựa chọn"),
        ("5", "Giáo viên THPT môn GDTC", "GV-THPT-01", "Bậc 3 đến Bậc 4", "2", "17 tiết/tuần", "Dạy Giáo dục thể chất, HKPĐ, rèn thể lực"),
        ("6", "Giáo viên THPT môn GDQP-AN", "GV-THPT-01", "Bậc 3 đến Bậc 4", "1", "17 tiết/tuần", "Dạy GDQP-AN, hội thao QP-AN, an ninh quốc gia"),
        ("7", "Giáo viên THPT môn Địa lí", "GV-THPT-01", "Bậc 3 đến Bậc 4", "2", "17 tiết/tuần", "Dạy Địa lí, bảo vệ tài nguyên môi trường"),
        ("8", "Giáo viên THPT môn GD KT&PL", "GV-THPT-01", "Bậc 3 đến Bậc 4", "1", "17 tiết/tuần", "Dạy Giáo dục kinh tế và pháp luật"),
        ("9", "Giáo viên THPT môn Vật lí", "GV-THPT-01", "Bậc 3 đến Bậc 4", "1", "17 tiết/tuần", "Dạy Vật lí, phòng bộ môn Vật lí, STEM"),
        ("10", "Giáo viên THPT môn Hóa học", "GV-THPT-01", "Bậc 3 đến Bậc 4", "1", "17 tiết/tuần", "Dạy Hóa học, thực hành hóa nghiệm, PCCC"),
        ("11", "Giáo viên THPT môn Sinh học", "GV-THPT-01", "Bậc 3 đến Bậc 4", "1", "17 tiết/tuần", "Dạy Sinh học, thực hành phòng thí nghiệm Sinh"),
        ("12", "Giáo viên THPT môn Công nghệ", "GV-THPT-01", "Bậc 3 đến Bậc 4", "2", "17 tiết/tuần", "Dạy Công nghệ công nghiệp & Nông nghiệp"),
        ("13", "Giáo viên THPT môn Tin học", "GV-THPT-01", "Bậc 3 đến Bậc 4", "1", "17 tiết/tuần", "Dạy Tin học, phòng máy tính, CĐS & AI"),

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
    # SHEET 3: BẢN MÔ TẢ MẪU NĐ 232 - CHUYÊN MÔN (13 MÔN HỌC)
    # ==============================================================================
    ws3 = wb.create_sheet(title="3. Mau Ban Mo Ta - Chuyen Mon")
    ws3.column_dimensions['A'].width = 10
    ws3.column_dimensions['B'].width = 25
    ws3.column_dimensions['C'].width = 38
    ws3.column_dimensions['D'].width = 32
    ws3.column_dimensions['E'].width = 26

    # Header công vụ
    ws3.merge_cells("A1:B1")
    ws3["A1"] = "SỞ GIÁO DỤC VÀ ĐÀO TẠO CAO BẰNG\nTRƯỜNG THPT PHỤC HÒA"
    ws3["A1"].font = Font(name=FONT_FAMILY, size=10, bold=True)
    ws3["A1"].alignment = ALIGN_CENTER

    ws3.merge_cells("C1:E1")
    ws3["C1"] = "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM\nĐộc lập – Tự do – Hạnh phúc\n_______________________"
    ws3["C1"].font = Font(name=FONT_FAMILY, size=10, bold=True)
    ws3["C1"].alignment = ALIGN_CENTER

    ws3.merge_cells("A3:E3")
    ws3["A3"] = "BẢN MÔ TẢ CÔNG VIỆC VÀ KHUNG NĂNG LỰC VỊ TRÍ VIỆC LÀM"
    ws3["A3"].font = Font(name=FONT_FAMILY, size=13, bold=True, color="1E3A8A")
    ws3["A3"].alignment = ALIGN_CENTER

    ws3.merge_cells("A4:E4")
    ws3["A4"] = "(Mẫu chuẩn theo Phần B Phụ lục IV/VI Nghị định số 232/2026/NĐ-CP - Áp dụng cho VTVL Giáo viên THPT 13 Môn học)"
    ws3["A4"].font = FONT_ITALIC
    ws3["A4"].alignment = ALIGN_CENTER

    sec_desc_cm = [
        ("I", "THÔNG TIN CHUNG", "", "", ""),
        ("1", "Tên vị trí việc làm:", "Giáo viên trung học phổ thông (Môn học: 13 Môn trong Phụ lục I)", "", ""),
        ("2", "Nhóm vị trí việc làm:", "Chuyên môn, nghiệp vụ (Mã số: GV-THPT-01)", "", ""),
        ("3", "Bậc nghề nghiệp sử dụng:", "Bậc 3 đến Bậc 4 (theo Phụ lục I Trường THPT Phục Hòa)", "", ""),
        ("4", "Lĩnh vực hoạt động:", "Giáo dục trung học phổ thông - Giảng dạy và giáo dục học sinh", "", ""),
        ("5", "Đơn vị:", "Trường THPT Phục Hòa, huyện Phục Hòa, tỉnh Cao Bằng", "", ""),
        ("6", "Cấp quản lý:", "Hiệu trưởng, Phó Hiệu trưởng phụ trách chuyên môn, Tổ trưởng chuyên môn", "", ""),
        ("7", "Vị trí việc làm liên quan:", "Tổ trưởng CM, Tổ phó CM, Giáo viên chủ nhiệm, Thư viện, Thiết bị TN, Giáo vụ", "", ""),
        ("8", "Số lượng người đảm nhiệm:", "22 người (gồm: Ngữ văn: 3, Toán: 3, Anh: 2, Sử: 2, GDTC: 2, GDQP: 1, Địa: 2, KTPL: 1, Lý: 1, Hóa: 1, Sinh: 1, CN: 2, Tin: 1)", "", ""),
        
        ("II", "MỤC TIÊU VỊ TRÍ VIỆC LÀM", "", "", ""),
        ("•", "Công thức mục tiêu:", "Trực tiếp giảng dạy, giáo dục học sinh theo Chương trình GDPT 2018 nhằm phát triển toàn diện phẩm chất, năng lực của học sinh THPT; bảo đảm hoàn thành chỉ tiêu chất lượng bộ môn và mục tiêu giáo dục của trường THPT Phục Hòa.", "", ""),

        ("III", "CÔNG VIỆC, KẾT QUẢ, SẢN PHẨM CHUYÊN MÔN (Bảng theo Bậc nghề nghiệp)", "", "", "")
    ]

    r_idx = 6
    for it in sec_desc_cm:
        if it[0] in ["I", "II", "III"]:
            ws3.merge_cells(start_row=r_idx, start_column=1, end_row=r_idx, end_column=5)
            ws3.cell(r_idx, 1, f"{it[0]}. {it[1]}")
            apply_row_styles(ws3, r_idx, font=FONT_SECTION, fill=SECTION_FILL)
        else:
            ws3.cell(r_idx, 1, it[0])
            ws3.cell(r_idx, 2, it[1])
            ws3.merge_cells(start_row=r_idx, start_column=3, end_row=r_idx, end_column=5)
            ws3.cell(r_idx, 3, it[2])
            apply_row_styles(ws3, r_idx, font=FONT_REGULAR)
            ws3.cell(r_idx, 1).alignment = ALIGN_CENTER
            ws3.cell(r_idx, 2).font = FONT_BOLD
        r_idx += 1

    # Bảng III chuyên môn theo Bậc
    tbl_headers_cm = ["Bậc", "Nội dung công việc chuyên môn", "Sản phẩm, Kết quả đầu ra", "Tiêu chí đánh giá", "Môn học áp dụng"]
    for c_i, h in enumerate(tbl_headers_cm, start=1):
        cell = ws3.cell(row=r_idx, column=c_i, value=h)
        cell.font = FONT_HEADER_WHITE
        cell.fill = TEAL_HEADER
        cell.alignment = ALIGN_CENTER
        cell.border = BORDER_HEADER
    r_idx += 1

    tbl_rows_cm = [
        ("Bậc 3", 
         "1. Giảng dạy môn học theo CT GDPT 2018 (định mức 17 tiết/tuần).\n2. Xây dựng Kế hoạch bài dạy (giáo án), đổi mới PPDH tích cực, tích hợp STEM.\n3. Kiểm tra đánh giá học sinh thường xuyên và định kỳ, vào điểm số điện tử đúng hạn.\n4. Thực hiện công tác giáo viên chủ nhiệm, quản lý nề nếp lớp được phân công.\n5. Bồi dưỡng học sinh yếu, phụ đạo học sinh có nguy cơ chưa đạt chuẩn.",
         "- 100% tiết dạy có giáo án chuẩn bị kỹ lưỡng trước khi lên lớp.\n- Sổ theo dõi đánh giá học sinh cập nhật chính xác.\n- Học sinh đạt chuẩn môn học ≥ 95%.\n- Học bạ điện tử được nhận xét và ký số đúng thời hạn.",
         "Chất lượng giảng dạy; tính độc lập sư phạm; khả năng giải quyết các tình huống sư phạm thông thường; tuân thủ đúng tiến độ chương trình GDPT 2018.",
         "Toàn bộ 13 môn học"),
        ("Bậc 4",
         "1. Thực hiện giảng dạy các chuyên đề nâng cao, dạy ôn thi tốt nghiệp THPT khối 12.\n2. Chủ trì hoặc nòng cốt tổ chức sinh hoạt chuyên môn theo NCBL cấp tổ/trường.\n3. Xây dựng đề kiểm tra định kỳ, ma trận đề, ngân hàng câu hỏi chuẩn quy chế.\n4. Trực tiếp bồi dưỡng đội tuyển học sinh giỏi môn học dự thi cấp tỉnh.\n5. Đổi mới sáng tạo: Có sáng kiến kinh nghiệm, bài học STEM hoặc ứng dụng AI xuất sắc được công nhận.",
         "- Ngân hàng câu hỏi ma trận đề thi chuẩn mực.\n- Có học sinh đạt giải HSG môn học cấp tỉnh.\n- Tỷ lệ tốt nghiệp THPT môn học đạt ≥ 98%.\n- Báo cáo chuyên đề đổi mới PPDH hoặc sáng kiến cấp cơ sở được nghiệm thu.",
         "Hiệu quả tổ chức chuyên môn; năng lực hướng dẫn đồng nghiệp; khả năng chuẩn hóa học liệu số; kết quả thi HSG và tốt nghiệp THPT vượt chỉ tiêu giao.",
         "Giáo viên cốt cán 13 môn")
    ]

    for row_data in tbl_rows_cm:
        for c_i, val in enumerate(row_data, start=1):
            ws3.cell(row=r_idx, column=c_i, value=val)
        apply_row_styles(ws3, r_idx, font=FONT_REGULAR)
        ws3.cell(r_idx, 1).alignment = ALIGN_CENTER
        ws3.cell(r_idx, 1).font = FONT_BOLD
        ws3.cell(r_idx, 5).alignment = ALIGN_CENTER
        r_idx += 1

    sec_next_cm = [
        ("IV", "CÔNG VIỆC, KẾT QUẢ, SẢN PHẨM QUẢN LÝ", "Không áp dụng đối với vị trí Giáo viên chuyên môn"),
        ("V", "CÔNG VIỆC, KẾT QUẢ, SẢN PHẨM HỖ TRỢ", "Không áp dụng đối với vị trí Giáo viên chuyên môn"),
        ("VI", "KHUNG NĂNG LỰC VỊ TRÍ VIỆC LÀM", 
         "1. Năng lực chung:\n - Phẩm chất đạo đức nhà giáo, chấp hành pháp luật, tinh thần phụng sự học sinh (Cấp độ 4 - 5)\n - Giao tiếp ứng xử sư phạm thân thiện với phụ huynh và học sinh (Cấp độ 4)\n - Đổi mới sáng tạo và áp dụng công nghệ số, AI vào soạn giảng (Cấp độ 4)\n - Kỹ năng sử dụng công nghệ thông tin, ngoại ngữ chuyên ngành (Cấp độ 3 - 4)\n2. Năng lực chuyên môn:\n - Nắm vững kiến thức 13 môn học GDPT 2018 (Cấp độ 4 - 5)\n - Phương pháp dạy học tích cực, giáo dục STEM (Cấp độ 4 - 5)\n - Kiểm tra đánh giá năng lực học sinh (Cấp độ 4)"),
        ("VII", "MỐI QUAN HỆ CÔNG TÁC",
         "1. Bên trong: Báo cáo trực tiếp Tổ trưởng CM, Phó Hiệu trưởng, Hiệu trưởng; phối hợp với GV chủ nhiệm, GV bộ môn khác, nhân viên Thiết bị, Thư viện, Y tế, Giáo vụ.\n2. Bên ngoài: Phối hợp thường xuyên với Cha mẹ học sinh (CMHS); Hội Khuyến học địa phương; Sở GD&ĐT (khi tham gia hội đồng thi, tập huấn chuyên môn)."),
        ("VIII", "PHẠM VI QUYỀN HẠN",
         "1. Về chuyên môn: Được chủ động lựa chọn phương pháp dạy học, kiểm tra đánh giá theo kế hoạch giáo dục môn học; đánh giá xếp loại học sinh theo quy chế; tham gia biên soạn tài liệu giáo dục địa phương.\n2. Về quản lý: Quản lý học sinh trong giờ học và các hoạt động giáo dục phân công.\n3. Về tài chính: Được bảo đảm trang thiết bị dạy học, phòng bộ môn, chế độ thừa giờ, bồi dưỡng theo quy định."),
        ("IX", "YÊU CẦU VỀ TRÌNH ĐỘ, KINH NGHIỆM, PHẨM CHẤT",
         "- Trình độ đào tạo: Bằng Cử nhân (Đại học) sư phạm trở lên đúng chuyên ngành môn học giảng dạy (Ngữ văn, Toán, Tiếng Anh, Lịch sử, Thể dục, Địa lý, Vật lý, Hóa học, Sinh học, Kỹ thuật công nghiệp/nông nghiệp, Tin học).\n- Bồi dưỡng: Có chứng chỉ bồi dưỡng theo tiêu chuẩn chức danh nghề nghiệp giáo viên THPT.\n- Ngoại ngữ & Tin học: Năng lực ngoại ngữ và tin học đạt chuẩn theo quy định hiện hành; làm chủ ký số học bạ.\n- Phẩm chất: Yêu nghề, tận tụy, mẫu mực, tâm huyết với sự nghiệp giáo dục vùng cao Phục Hòa.")
    ]

    for sec in sec_next_cm:
        ws3.merge_cells(start_row=r_idx, start_column=1, end_row=r_idx, end_column=5)
        ws3.cell(r_idx, 1, f"{sec[0]}. {sec[1]}")
        apply_row_styles(ws3, r_idx, font=FONT_SECTION, fill=SECTION_FILL)
        r_idx += 1

        ws3.merge_cells(start_row=r_idx, start_column=1, end_row=r_idx, end_column=5)
        ws3.cell(r_idx, 1, sec[2])
        apply_row_styles(ws3, r_idx, font=FONT_REGULAR)
        r_idx += 1

    # Khung ký duyệt
    r_idx += 1
    ws3.merge_cells(start_row=r_idx, start_column=1, end_row=r_idx, end_column=2)
    ws3.cell(r_idx, 1, "NGƯỜI LẬP BIỂU\n(Ký và ghi rõ họ tên)")
    ws3.cell(r_idx, 1).font = FONT_BOLD
    ws3.cell(r_idx, 1).alignment = ALIGN_CENTER

    ws3.merge_cells(start_row=r_idx, start_column=3, end_row=r_idx, end_column=5)
    ws3.cell(r_idx, 3, "Phục Hòa, ngày .... tháng .... năm 2026\nHIỆU TRƯỞNG\n(Ký tên, đóng dấu)")
    ws3.cell(r_idx, 3).font = FONT_BOLD
    ws3.cell(r_idx, 3).alignment = ALIGN_CENTER

    # ==============================================================================
    # SHEET 4: BẢN MÔ TẢ MẪU NĐ 232 - HỖ TRỢ (06 VỊ TRÍ)
    # ==============================================================================
    ws4 = wb.create_sheet(title="4. Mau Ban Mo Ta - Ho Tro")
    ws4.column_dimensions['A'].width = 10
    ws4.column_dimensions['B'].width = 25
    ws4.column_dimensions['C'].width = 38
    ws4.column_dimensions['D'].width = 32
    ws4.column_dimensions['E'].width = 26

    # Header công vụ
    ws4.merge_cells("A1:B1")
    ws4["A1"] = "SỞ GIÁO DỤC VÀ ĐÀO TẠO CAO BẰNG\nTRƯỜNG THPT PHỤC HÒA"
    ws4["A1"].font = Font(name=FONT_FAMILY, size=10, bold=True)
    ws4["A1"].alignment = ALIGN_CENTER

    ws4.merge_cells("C1:E1")
    ws4["C1"] = "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM\nĐộc lập – Tự do – Hạnh phúc\n_______________________"
    ws4["C1"].font = Font(name=FONT_FAMILY, size=10, bold=True)
    ws4["C1"].alignment = ALIGN_CENTER

    ws4.merge_cells("A3:E3")
    ws4["A3"] = "BẢN MÔ TẢ CÔNG VIỆC VÀ KHUNG NĂNG LỰC VỊ TRÍ VIỆC LÀM HỖ TRỢ"
    ws4["A3"].font = Font(name=FONT_FAMILY, size=13, bold=True, color="1E3A8A")
    ws4["A3"].alignment = ALIGN_CENTER

    ws4.merge_cells("A4:E4")
    ws4["A4"] = "(Mẫu chuẩn theo Phần B Phụ lục IV/VI Nghị định số 232/2026/NĐ-CP - Áp dụng cho 06 Viên chức hỗ trợ)"
    ws4["A4"].font = FONT_ITALIC
    ws4["A4"].alignment = ALIGN_CENTER

    sec_desc_ht = [
        ("I", "THÔNG TIN CHUNG", "", "", ""),
        ("1", "Tên nhóm vị trí:", "Viên chức hỗ trợ (Gồm: Kế toán, Văn thư, Thiết bị TN, Giáo vụ, Y tế học đường, Thủ quỹ)", "", ""),
        ("2", "Nhóm vị trí việc làm:", "Hỗ trợ, phục vụ (Định mức làm việc: 40 giờ/tuần)", "", ""),
        ("3", "Bậc nghề nghiệp sử dụng:", "Bậc 1 đến Bậc 4 (theo Phụ lục I Trường THPT Phục Hòa)", "", ""),
        ("4", "Lĩnh vực hoạt động:", "Tài chính, hành chính, thư viện số, thiết bị dạy học, giáo vụ, y tế học đường, ngân quỹ", "", ""),
        ("5", "Đơn vị:", "Trường THPT Phục Hòa, huyện Phục Hòa, tỉnh Cao Bằng", "", ""),
        ("6", "Cấp quản lý:", "Hiệu trưởng, Phó Hiệu trưởng trực tiếp phụ trách cơ sở vật chất và văn phòng", "", ""),
        ("7", "Số lượng người đảm nhiệm:", "06 người (Kế toán: 1, Văn thư: 1, Thiết bị TN: 1, Giáo vụ: 1, Y tế: 1, Thủ quỹ: 1)", "", ""),

        ("II", "MỤC TIÊU VỊ TRÍ VIỆC LÀM", "", "", ""),
        ("•", "Công thức mục tiêu:", "Thực hiện các hoạt động hỗ trợ về tài chính, văn thư lưu trữ, quản lý thiết bị dạy học, hồ sơ học sinh, chăm sóc y tế ban đầu và quản lý quỹ tiền mặt nhằm bảo đảm các điều kiện vận hành thông suốt, an toàn và đúng pháp luật của trường THPT Phục Hòa.", "", ""),

        ("V", "CÔNG VIỆC, KẾT QUẢ, SẢN PHẨM HỖ TRỢ (Bảng theo Bậc nghề nghiệp cho 06 Vị trí)", "", "", "")
    ]

    r_idx = 6
    for it in sec_desc_ht:
        if it[0] in ["I", "II", "V"]:
            ws4.merge_cells(start_row=r_idx, start_column=1, end_row=r_idx, end_column=5)
            ws4.cell(r_idx, 1, f"{it[0]}. {it[1]}")
            apply_row_styles(ws4, r_idx, font=FONT_SECTION, fill=SECTION_FILL_AMBER)
        else:
            ws4.cell(r_idx, 1, it[0])
            ws4.cell(r_idx, 2, it[1])
            ws4.merge_cells(start_row=r_idx, start_column=3, end_row=r_idx, end_column=5)
            ws4.cell(r_idx, 3, it[2])
            apply_row_styles(ws4, r_idx, font=FONT_REGULAR)
            ws4.cell(r_idx, 1).alignment = ALIGN_CENTER
            ws4.cell(r_idx, 2).font = FONT_BOLD
        r_idx += 1

    tbl_headers_ht = ["Vị trí & Bậc", "Nội dung công việc hỗ trợ cốt lõi", "Sản phẩm, Kết quả đầu ra", "Tiêu chí đánh giá", "Định mức"]
    for c_i, h in enumerate(tbl_headers_ht, start=1):
        cell = ws4.cell(row=r_idx, column=c_i, value=h)
        cell.font = FONT_HEADER_WHITE
        cell.fill = AMBER_HEADER
        cell.alignment = ALIGN_CENTER
        cell.border = BORDER_HEADER
    r_idx += 1

    tbl_rows_ht = [
        ("1. Kế toán\n(Bậc 1 - 4)",
         "Lập dự toán ngân sách, chi trả lương, phụ cấp, BHXH, đối chiếu Kho bạc Nhà nước, quyết toán tài chính quý/năm.",
         "Dự toán được duyệt; bảng lương hàng tháng trước ngày 05; báo cáo tài chính năm đúng luật, không sai sót kiểm toán.",
         "Chính xác tuyệt đối; đúng Luật Ngân sách và Luật Kế toán; 100% đúng thời hạn.",
         "40 giờ/tuần\n(01 người)"),
        ("2. Văn thư\n(Bậc 1 - 4)",
         "Xử lý công văn đi đến trên iOffice, quản lý con dấu nhà trường, lập hồ sơ lưu trữ hiện hành và lịch sử, cấp giấy giới thiệu.",
         "100% công văn xử lý trong ngày; bảo quản con dấu an toàn tuyệt đối; số hóa hồ sơ văn bản khoa học.",
         "Bảo mật, kịp thời, đúng thể thức văn bản theo Nghị định 30/2020/NĐ-CP.",
         "40 giờ/tuần\n(01 người)"),
        ("3. Thiết bị TN\n(Bậc 2 - 3)",
         "Quản lý thiết bị dạy học, hóa chất phòng thực hành Lý-Hóa-Sinh; chuẩn bị đồ dùng tiết thực hành; an toàn PCCC phòng thí nghiệm.",
         "100% tiết thực hành có thiết bị chuẩn bị sẵn sàng; sổ theo dõi mượn trả cập nhật hàng tuần; không có tai nạn cháy nổ.",
         "Bảo quản tốt 100% tài sản; an toàn tuyệt đối phòng thí nghiệm.",
         "40 giờ/tuần\n(01 người)"),
        ("4. Giáo vụ\n(Bậc 2 - 3)",
         "Quản lý sổ đăng bộ, hồ sơ tuyển sinh lớp 10, hồ sơ thi tốt nghiệp THPT, cơ sở dữ liệu học sinh, cấp phát văn bằng chứng chỉ.",
         "Sổ đăng bộ chuẩn xác 100%; hồ sơ tuyển sinh và thi tốt nghiệp nộp về Sở GD&ĐT đúng thời hạn; cấp bằng không sai sót.",
         "Chính xác, bảo mật dữ liệu học sinh, tuân thủ quy chế quản lý văn bằng.",
         "40 giờ/tuần\n(01 người)"),
        ("5. Y tế trường học\n(Bậc 1 - 3)",
         "Sơ cấp cứu tai nạn thương tích; quản lý tủ thuốc; khám sức khỏe định kỳ; phòng chống dịch bệnh; kiểm tra vệ sinh căng tin.",
         "Tủ thuốc đủ cơ số cấp cứu; 100% học sinh được theo dõi sức khỏe; không để xảy ra dịch bệnh bùng phát hay ngộ độc.",
         "Kịp thời, chu đáo, tuân thủ quy chuẩn y tế học đường.",
         "40 giờ/tuần\n(01 người)"),
        ("6. Thủ quỹ\n(Bậc 1 - 3)",
         "Quản lý két quỹ tiền mặt; thực hiện thu và chi tiền mặt đúng phiếu thu/chi có chữ ký Hiệu trưởng & Kế toán; ghi sổ quỹ hàng ngày.",
         "Số dư quỹ tiền mặt thực tế khớp 100% với sổ quỹ và sổ kế toán; thanh toán tiền mặt kịp thời, an toàn tuyệt đối.",
         "Tuyệt đối trung thực, cẩn trọng; khớp số dư 100%; không để xảy ra thiếu hụt.",
         "40 giờ/tuần\n(01 người)")
    ]

    for row_data in tbl_rows_ht:
        for c_i, val in enumerate(row_data, start=1):
            ws4.cell(row=r_idx, column=c_i, value=val)
        apply_row_styles(ws4, r_idx, font=FONT_REGULAR)
        ws4.cell(r_idx, 1).alignment = ALIGN_CENTER
        ws4.cell(r_idx, 1).font = FONT_BOLD
        ws4.cell(r_idx, 5).alignment = ALIGN_CENTER
        r_idx += 1

    sec_next_ht = [
        ("VI", "KHUNG NĂNG LỰC CHO 06 VỊ TRÍ HỖ TRỢ",
         "- Năng lực chung: Đạo đức công vụ liêm chính (Cấp độ 3-4); Kỷ luật trách nhiệm (Cấp độ 3-4); Giao tiếp lịch thiệp, phục vụ chu đáo (Cấp độ 3); Ứng dụng công nghệ thông tin và phần mềm chuyên môn (Cấp độ 2-3).\n- Năng lực chuyên môn: Nắm vững quy trình nghiệp vụ tài chính kế toán, văn thư lưu trữ, an toàn thiết bị thí nghiệm, quản lý văn bằng giáo vụ, kỹ năng sơ cứu y tế và nghiệp vụ quản lý quỹ tiền mặt."),
        ("VII", "MỐI QUAN HỆ CÔNG TÁC",
         "1. Bên trong: Chịu sự quản lý điều hành trực tiếp của Hiệu trưởng và Phó Hiệu trưởng phụ trách; phối hợp chặt chẽ với Tổ chuyên môn, Giáo viên chủ nhiệm, Giáo viên bộ môn.\n2. Bên ngoài: Quan hệ công tác với Kho bạc Nhà nước, Cơ quan Thuế, Trung tâm Y tế huyện, Bảo hiểm xã hội huyện Phục Hòa."),
        ("VIII", "PHẠM VI QUYỀN HẠN",
         "1. Về chuyên môn: Được từ chối các khoản thu - chi hoặc xuất kho thiết bị không đúng quy định, không đủ chứng từ hợp lệ; kiến nghị các biện pháp bảo đảm an toàn trường học.\n2. Về quản lý & Tài chính: Trực tiếp quản lý hệ thống sổ sách, con dấu, cơ sở dữ liệu và tài sản được phân công quản lý."),
        ("IX", "YÊU CẦU VỀ TRÌNH ĐỘ, KINH NGHIỆM CHO 06 VỊ TRÍ",
         "- Kế toán: Đại học chuyên ngành Tài chính - Kế toán; chứng chỉ kế toán viên; làm chủ phần mềm MISA Mimosa.\n- Văn thư: Trung cấp trở lên ngành Văn thư - Lưu trữ; thành thạo hệ thống iOffice và chứng thư số.\n- Thiết bị TN: Cao đẳng trở lên ngành Thiết bị trường học hoặc Lý, Hóa, Sinh; chứng chỉ an toàn hóa chất.\n- Giáo vụ: Trung cấp trở lên ngành Quản lý giáo dục, Tin học hoặc tương đương; thành thạo CSDL ngành GDĐT.\n- Y tế học đường: Trung cấp Y sĩ hoặc Điều dưỡng trở lên; chứng chỉ sơ cấp cứu ban đầu.\n- Thủ quỹ: Trung cấp trở lên ngành Tài chính - Kế toán, Ngân hàng hoặc Kinh tế; phẩm chất đạo đức trung thực tuyệt đối.")
    ]

    for sec in sec_next_ht:
        ws4.merge_cells(start_row=r_idx, start_column=1, end_row=r_idx, end_column=5)
        ws4.cell(r_idx, 1, f"{sec[0]}. {sec[1]}")
        apply_row_styles(ws4, r_idx, font=FONT_SECTION, fill=SECTION_FILL_AMBER)
        r_idx += 1

        ws4.merge_cells(start_row=r_idx, start_column=1, end_row=r_idx, end_column=5)
        ws4.cell(r_idx, 1, sec[2])
        apply_row_styles(ws4, r_idx, font=FONT_REGULAR)
        r_idx += 1

    # Khung ký duyệt
    r_idx += 1
    ws4.merge_cells(start_row=r_idx, start_column=1, end_row=r_idx, end_column=2)
    ws4.cell(r_idx, 1, "NGƯỜI LẬP BIỂU\n(Ký và ghi rõ họ tên)")
    ws4.cell(r_idx, 1).font = FONT_BOLD
    ws4.cell(r_idx, 1).alignment = ALIGN_CENTER

    ws4.merge_cells(start_row=r_idx, start_column=3, end_row=r_idx, end_column=5)
    ws4.cell(r_idx, 3, "Phục Hòa, ngày .... tháng .... năm 2026\nHIỆU TRƯỞNG\n(Ký tên, đóng dấu)")
    ws4.cell(r_idx, 3).font = FONT_BOLD
    ws4.cell(r_idx, 3).alignment = ALIGN_CENTER

    # ==============================================================================
    # SHEET 5: KHUNG NĂNG LỰC & KPI NĐ 232 (CHO 35 BIÊN CHẾ)
    # ==============================================================================
    ws5 = wb.create_sheet(title="5. Khung Nang Luc & KPI")
    ws5.column_dimensions['A'].width = 6
    ws5.column_dimensions['B'].width = 20
    ws5.column_dimensions['C'].width = 28
    ws5.column_dimensions['D'].width = 18
    ws5.column_dimensions['E'].width = 18
    ws5.column_dimensions['F'].width = 18
    ws5.column_dimensions['G'].width = 38

    ws5.merge_cells("A1:G1")
    ws5["A1"] = "KHUNG NĂNG LỰC TOÀN DIỆN CỦA 03 DANH MỤC VTVL (5 CẤP ĐỘ)"
    ws5["A1"].font = FONT_TITLE_MAIN
    ws5["A1"].alignment = ALIGN_CENTER

    ws5.merge_cells("A2:G2")
    ws5["A2"] = "(Căn cứ Điều 8, Điều 9 & Phụ lục IV/VI Nghị định số 232/2026/NĐ-CP - Chuẩn hóa 35 Biên chế)"
    ws5["A2"].font = FONT_ITALIC
    ws5["A2"].alignment = ALIGN_CENTER

    headers_ws5 = ["STT", "Nhóm năng lực", "Tên năng lực thành phần", "Yêu cầu Nhóm 1\n(07 Quản lý)", "Yêu cầu Nhóm 2\n(22 GV 13 Môn)", "Yêu cầu Nhóm 3\n(06 VC Hỗ trợ)", "Mô tả chuẩn hành vi & Tiêu chí đạt"]
    ws5.append([])
    ws5.append(headers_ws5)
    for col_idx in range(1, len(headers_ws5) + 1):
        cell = ws5.cell(row=4, column=col_idx)
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
        ("7", "Năng lực chuyên môn", "Nghiệp vụ sư phạm môn học / Chuyên môn hỗ trợ", "Cấp độ 4 - 5", "Cấp độ 3 - 5", "Cấp độ 3 - 4", "Nắm vững kiến thức 13 môn học GDPT 2018 hoặc chuyên ngành tài chính/văn thư/y tế."),
        ("8", "Năng lực chuyên môn", "Phương pháp dạy học tích cực / Kỹ thuật nghiệp vụ", "Cấp độ 4 - 5", "Cấp độ 4 - 5", "Cấp độ 2 - 3", "Vận dụng PPDH tích cực, giáo dục STEM, kỹ thuật ghi chép và quy trình kiểm soát."),
        ("9", "Năng lực chuyên môn", "Kiểm tra đánh giá & Phân tích chất lượng", "Cấp độ 4 - 5", "Cấp độ 3 - 5", "Cấp độ 2 - 3", "Đánh giá đúng năng lực học sinh, xây dựng ma trận đề chuẩn, phân tích dữ liệu."),
        ("10", "Năng lực bổ trợ", "Xử lý tình huống sư phạm & Giải quyết tranh chấp", "Cấp độ 4 - 5", "Cấp độ 3 - 4", "Cấp độ 2 - 3", "Điềm tĩnh, thấu tình đạt lý, giải quyết xung đột trong môi trường học đường nhân văn.")
    ]

    cur_row = 5
    for item in comp_rows:
        for c_idx, val in enumerate(item, start=1):
            ws5.cell(cur_row, c_idx, val)
        apply_row_styles(ws5, cur_row, font=FONT_REGULAR, alignment=ALIGN_LEFT)
        ws5.cell(cur_row, 1).alignment = ALIGN_CENTER
        ws5.cell(cur_row, 4).alignment = ALIGN_CENTER
        ws5.cell(cur_row, 5).alignment = ALIGN_CENTER
        ws5.cell(cur_row, 6).alignment = ALIGN_CENTER
        cur_row += 1

    # KPI Section trong Sheet 5
    cur_row += 2
    ws5.merge_cells(start_row=cur_row, start_column=1, end_row=cur_row, end_column=7)
    ws5.cell(cur_row, 1, "BỘ CHỈ SỐ KPI ĐO LƯỜNG HIỆU QUẢ ĐẦU RA CHO 35 BIÊN CHẾ")
    ws5.cell(cur_row, 1).font = FONT_TITLE_MAIN
    ws5.cell(cur_row, 1).alignment = ALIGN_CENTER
    cur_row += 1

    kpi_headers = ["STT", "Nhóm VTVL", "Vị trí áp dụng", "Chỉ số KPI chính", "Mục tiêu định lượng", "Tần suất", "Trọng số %"]
    for c_i, h in enumerate(kpi_headers, start=1):
        cell = ws5.cell(row=cur_row, column=c_i, value=h)
        cell.font = FONT_HEADER_WHITE
        cell.fill = EMERALD_HEADER
        cell.alignment = ALIGN_CENTER
        cell.border = BORDER_HEADER
    cur_row += 1

    kpi_items = [
        ("1", "I. Quản lý (7 người)", "Hiệu trưởng", "Tỷ lệ tốt nghiệp THPT toàn trường", "Đạt ≥ 98.5%", "Hàng năm", "25%"),
        ("2", "I. Quản lý (7 người)", "Phó Hiệu trưởng", "Tiến độ chương trình & Ký số học bạ", "100% đúng tiến độ & ký số", "Tháng / Kỳ", "25%"),
        ("3", "I. Quản lý (7 người)", "Tổ trưởng CM", "Sinh hoạt NCBL & Chuyên đề STEM", "≥ 02 chuyên đề / học kỳ", "Hàng tháng", "25%"),
        ("4", "I. Quản lý (7 người)", "Tổ phó CM", "Lịch báo giảng & Sổ sách điện tử", "100% thành viên đúng hạn", "Hàng tuần", "25%"),
        ("5", "II. Chuyên môn (22 GV)", "22 GV 13 Môn học", "Kế hoạch bài dạy & Dạy 17 tiết/tuần", "100% tiết có giáo án chuẩn", "Hàng tuần", "30%"),
        ("6", "II. Chuyên môn (22 GV)", "22 GV 13 Môn học", "Chất lượng môn học phụ trách", "Học sinh đạt chuẩn ≥ 96%", "Học kỳ / Năm", "30%"),
        ("7", "II. Chuyên môn (22 GV)", "22 GV 13 Môn học", "Bồi dưỡng HSG & Phụ đạo yếu kém", "Có giải HSG cấp tỉnh", "Hàng năm", "20%"),
        ("8", "III. Hỗ trợ (6 người)", "Kế toán viên", "Chi trả lương & Quyết toán ngân sách", "100% đúng hạn, kiểm toán tốt", "Hàng tháng", "35%"),
        ("9", "III. Hỗ trợ (6 người)", "Văn thư", "Xử lý công văn đi đến trên iOffice", "100% xử lý trong ngày", "Hàng ngày", "35%"),
        ("10", "III. Hỗ trợ (6 người)", "Thiết bị, TN", "Chuẩn bị thiết bị cho tiết thực hành", "100% tiết thực hành đủ đồ dùng", "Hàng tuần", "35%"),
        ("11", "III. Hỗ trợ (6 người)", "Giáo vụ", "Sổ đăng bộ, Tuyển sinh 10 & Thi TN", "100% hồ sơ chính xác, đúng hạn", "Hàng kỳ", "35%"),
        ("12", "III. Hỗ trợ (6 người)", "Y tế học đường", "Chăm sóc sức khỏe & Phòng chống dịch", "100% HS theo dõi sức khỏe", "Hàng kỳ", "35%"),
        ("13", "III. Hỗ trợ (6 người)", "Thủ quỹ", "Quản lý két tiền mặt & Khớp số dư", "Khớp số dư 100%, an toàn", "Hàng ngày", "35%")
    ]

    for it in kpi_items:
        for c_i, val in enumerate(it, start=1):
            ws5.cell(row=cur_row, column=c_i, value=val)
        apply_row_styles(ws5, cur_row, font=FONT_REGULAR)
        ws5.cell(cur_row, 1).alignment = ALIGN_CENTER
        ws5.cell(cur_row, 6).alignment = ALIGN_CENTER
        ws5.cell(cur_row, 7).alignment = ALIGN_CENTER
        cur_row += 1

    # Auto-fit columns
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
