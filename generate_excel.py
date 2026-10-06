# -*- coding: utf-8 -*-
"""
Script sinh file Excel: Đề án Vị trí việc làm Viên chức – THPT Phục Hòa
Căn cứ Nghị định số 232/2026/NĐ-CP ngày 26/6/2026 của Chính phủ & Thông tư 15/2026/TT-BGDĐT
Bao gồm trọn bộ 03 Danh mục (Nhóm) Vị trí việc làm:
1. DANH MỤC 1: Vị trí việc làm Quản lý (Hiệu trưởng, Phó Hiệu trưởng, TTCM, TPCM)
2. DANH MỤC 2: Vị trí việc làm Chuyên môn, nghiệp vụ (Giáo viên THPT, Thư viện, CNTT & Chuyển đổi số)
3. DANH MỤC 3: Vị trí việc làm Hỗ trợ (Kế toán, Văn thư, Thiết bị thí nghiệm, Giáo vụ, Y tế)
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

    FONT_TITLE_MAIN = Font(name=FONT_FAMILY, size=15, bold=True, color="1E3A8A")
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
            adjusted_width = min(max(max_len + 3, 10), max_width_limit)
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
    ws1.column_dimensions['F'].width = 38

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
    ws1["A4"] = "ĐỀ ÁN XÂY DỰNG VỊ TRÍ VIỆC LÀM VIÊN CHỨC TRƯỜNG THPT"
    ws1["A4"].font = FONT_TITLE_MAIN
    ws1["A4"].alignment = ALIGN_CENTER

    ws1.merge_cells("A5:F5")
    ws1["A5"] = "Căn cứ Nghị định số 232/2026/NĐ-CP ngày 26/6/2026 của Chính phủ quy định về vị trí việc làm viên chức"
    ws1["A5"].font = Font(name=FONT_FAMILY, size=11, italic=True)
    ws1["A5"].alignment = ALIGN_CENTER

    ws1.merge_cells("A6:F6")
    ws1["A6"] = "Bao gồm trọn bộ 03 Danh mục: Quản lý (Phụ lục I), Chuyên môn nghiệp vụ (Phụ lục II), Hỗ trợ (Phụ lục III)"
    ws1["A6"].font = Font(name=FONT_FAMILY, size=11, bold=True, color="1E3A8A")
    ws1["A6"].alignment = ALIGN_CENTER

    ws1.merge_cells("A7:F7")
    ws1["A7"] = "Năm học: 2026 - 2027 | Đơn vị: Trường THPT Phục Hòa, huyện Phục Hòa, tỉnh Cao Bằng"
    ws1["A7"].font = FONT_ITALIC
    ws1["A7"].alignment = ALIGN_CENTER

    # I. CĂN CỨ PHÁP LÝ
    ws1.merge_cells("A9:F9")
    ws1["A9"] = "I. CƠ SỞ PHÁP LÝ XÂY DỰNG ĐỀ ÁN VỊ TRÍ VIỆC LÀM"
    apply_row_styles(ws1, 9, font=FONT_SECTION, fill=SECTION_FILL)

    legal_rows = [
        ("1", "Nghị định số 232/2026/NĐ-CP", "Chính phủ", "26/06/2026", "01/07/2026", "Quy định vị trí việc làm viên chức trong đơn vị sự nghiệp công lập (Phụ lục I, II, III, IV)"),
        ("2", "Thông tư số 15/2026/TT-BGDĐT", "Bộ GD&ĐT", "24/03/2026", "10/05/2026", "Ban hành Điều lệ trường THCS, trường THPT và trường phổ thông có nhiều cấp học"),
        ("3", "Luật Viên chức số 129/2025/QH15", "Quốc hội", "2025", "2026", "Quy định nguyên tắc, quyền, nghĩa vụ, tuyển dụng, sử dụng và quản lý viên chức"),
        ("4", "Thông tư số 20/2023/TT-BGDĐT", "Bộ GD&ĐT", "30/10/2023", "16/12/2023", "Hướng dẫn về vị trí việc làm, cơ cấu viên chức theo chức danh nghề nghiệp và định mức số lượng người làm việc trong các cơ sở GDPT"),
        ("5", "Thông tư số 04/2021/TT-BGDĐT & TT 08/2023/TT-BGDĐT", "Bộ GD&ĐT", "02/02/2021", "20/03/2021", "Quy định mã số, tiêu chuẩn CDNN và bổ nhiệm, xếp lương giáo viên THPT"),
        ("6", "Thông tư số 20/2018/TT-BGDĐT & TT 14/2018/TT-BGDĐT", "Bộ GD&ĐT", "2018", "2018", "Ban hành Chuẩn nghề nghiệp giáo viên THPT và Chuẩn Hiệu trưởng cơ sở GDPT"),
    ]

    headers_legal = ["STT", "Văn bản căn cứ", "Cơ quan ban hành", "Ngày ký", "Hiệu lực", "Nội dung điều chỉnh chính"]
    for col_i, h in enumerate(headers_legal, 1):
        cell = ws1.cell(row=10, column=col_i, value=h)
        cell.font = FONT_HEADER_WHITE
        cell.fill = BLUE_HEADER
        cell.alignment = ALIGN_CENTER
        cell.border = BORDER_HEADER

    for r_i, ldata in enumerate(legal_rows, 11):
        for c_i, val in enumerate(ldata, 1):
            cell = ws1.cell(row=r_i, column=c_i, value=val)
            cell.font = FONT_REGULAR
            cell.border = BORDER_CELL
            cell.alignment = ALIGN_CENTER if c_i in [1, 4, 5] else ALIGN_LEFT

    # II. NGUYÊN TẮC PHÂN LOẠI 03 DANH MỤC VTVL (ĐIỀU 5 & ĐIỀU 6 NGHỊ ĐỊNH 232/2026)
    r_start = 18
    ws1.merge_cells(f"A{r_start}:F{r_start}")
    ws1[f"A{r_start}"] = "II. NGUYÊN TẮC PHÂN LOẠI VÀ MÔ TẢ 03 DANH MỤC VTVL (ĐIỀU 5, 6, 8 NĐ 232/2026/NĐ-CP)"
    apply_row_styles(ws1, r_start, font=FONT_SECTION, fill=SECTION_FILL)

    principles = [
        ("1", "Danh mục 1: Vị trí Quản lý (Phụ lục I)", "Bao gồm người đứng đầu, cấp phó người đứng đầu và viên chức quản lý cấp tổ/bộ phận. Xếp lương theo bậc chuyên môn đang giữ và hưởng phụ cấp chức vụ lãnh đạo (HT: 0.70; PHT: 0.50; TTCM: 0.25; TPCM: 0.15)."),
        ("2", "Danh mục 2: Vị trí Chuyên môn, nghiệp vụ (Phụ lục II)", "Thực hiện trực tiếp chức năng, nhiệm vụ chuyên môn cốt lõi của nhà trường (Giảng dạy, giáo dục học sinh, thư viện, công nghệ thông tin và chuyển đổi số). Áp dụng Bậc nghề nghiệp từ Bậc 1 đến Bậc 5."),
        ("3", "Danh mục 3: Vị trí Hỗ trợ (Phụ lục III)", "Thực hiện hoạt động phục vụ, bảo đảm điều kiện vận hành hoạt động của nhà trường (Kế toán, Văn thư, Thiết bị thí nghiệm, Giáo vụ, Y tế học đường). Sản phẩm đầu ra là dịch vụ nội bộ hoặc điều kiện cơ sở vật chất."),
        ("4", "Nguyên tắc xây dựng Bản mô tả & Khung năng lực", "Mỗi vị trí xây dựng 01 bản mô tả chuẩn theo Phụ lục IV (Phần B) NĐ 232; xác định rõ mục tiêu theo công thức (Động từ + Đối tượng + Kết quả cuối cùng); gắn tỷ trọng % và tiêu chí KPI đo lường được."),
    ]
    for r_i, pdata in enumerate(principles, r_start+1):
        ws1.cell(row=r_i, column=1, value=pdata[0]).alignment = ALIGN_CENTER
        ws1.cell(row=r_i, column=1).border = BORDER_CELL
        ws1.cell(row=r_i, column=2, value=pdata[1]).font = FONT_BOLD
        ws1.cell(row=r_i, column=2).alignment = ALIGN_LEFT
        ws1.cell(row=r_i, column=2).border = BORDER_CELL
        ws1.merge_cells(f"C{r_i}:F{r_i}")
        ws1.cell(row=r_i, column=3, value=pdata[2]).font = FONT_REGULAR
        ws1.cell(row=r_i, column=3).alignment = ALIGN_LEFT
        for c in range(3, 7): ws1.cell(row=r_i, column=c).border = BORDER_CELL

    # Signatures
    r_sig = r_start + len(principles) + 2
    ws1.merge_cells(f"A{r_sig}:C{r_sig}")
    ws1[f"A{r_sig}"] = "NGƯỜI LẬP ĐỀ ÁN"
    ws1[f"A{r_sig}"].font = FONT_BOLD
    ws1[f"A{r_sig}"].alignment = ALIGN_CENTER

    ws1.merge_cells(f"D{r_sig}:F{r_sig}")
    ws1[f"D{r_sig}"] = "HIỆU TRƯỞNG TRƯỜNG THPT PHỤC HÒA"
    ws1[f"D{r_sig}"].font = FONT_BOLD
    ws1[f"D{r_sig}"].alignment = ALIGN_CENTER

    # ==============================================================================
    # SHEET 2: TỔNG HỢP 3 DANH MỤC VỊ TRÍ VIỆC LÀM (MA TRẬN TOÀN TRƯỜNG)
    # ==============================================================================
    ws2 = wb.create_sheet(title="2. Tổng hợp 3 Danh mục VTVL")
    ws2.column_dimensions['A'].width = 8
    ws2.column_dimensions['B'].width = 16
    ws2.column_dimensions['C'].width = 24
    ws2.column_dimensions['D'].width = 26
    ws2.column_dimensions['E'].width = 12
    ws2.column_dimensions['F'].width = 22
    ws2.column_dimensions['G'].width = 16
    ws2.column_dimensions['H'].width = 16
    ws2.column_dimensions['I'].width = 28

    ws2.merge_cells("A1:I1")
    ws2["A1"] = "BẢNG TỔNG HỢP 03 DANH MỤC VỊ TRÍ VIỆC LÀM TRƯỜNG THPT PHỤC HÒA"
    ws2["A1"].font = FONT_TITLE_MAIN
    ws2["A1"].alignment = ALIGN_CENTER

    ws2.merge_cells("A2:I2")
    ws2["A2"] = "Thực hiện theo Nghị định số 232/2026/NĐ-CP (Phụ lục I, II, III) và Thông tư 20/2023/TT-BGDĐT"
    ws2["A2"].font = Font(name=FONT_FAMILY, size=11, italic=True)
    ws2["A2"].alignment = ALIGN_CENTER

    headers_summary = [
        "STT", "Mã số VTVL", "Tên Vị trí việc làm", "Phân nhóm danh mục (NĐ 232)", "Số lượng",
        "Bậc nghề nghiệp áp dụng", "Phụ cấp chức vụ", "Định mức công việc", "Cơ quan phê duyệt / Bổ nhiệm"
    ]
    for col_i, h in enumerate(headers_summary, 1):
        cell = ws2.cell(row=4, column=col_i, value=h)
        cell.font = FONT_HEADER_WHITE
        cell.fill = NAVY_FILL
        cell.alignment = ALIGN_CENTER
        cell.border = BORDER_HEADER

    # Data for 3 categories
    full_vtvl_summary = [
        # DANH MỤC 1: QUẢN LÝ
        ("1", "HT-THPT-01", "Hiệu trưởng", "DM1: Quản lý (Người đứng đầu ĐVSN)", "01", "Bậc 4 - Bậc 5 (GV Hạng II, I)", "0.70", "02 tiết/tuần", "Giám đốc Sở GD&ĐT Cao Bằng"),
        ("2", "PHT-THPT-01", "Phó Hiệu trưởng", "DM1: Quản lý (Cấp phó người đứng đầu)", "02", "Bậc 3 - Bậc 5 (GV Hạng III, II, I)", "0.50", "04 tiết/tuần", "Giám đốc Sở GD&ĐT Cao Bằng"),
        ("3", "TTCM-THPT-01", "Tổ trưởng chuyên môn", "DM1: Quản lý (Tổ trưởng tổ CM)", "04 - 06", "Bậc 3 - Bậc 4 (GV Hạng III, II)", "0.25", "Giảm 3 tiết (dạy 14 tiết/tuần)", "Hiệu trưởng THPT Phục Hòa"),
        ("4", "TPCM-THPT-01", "Tổ phó chuyên môn", "DM1: Quản lý (Phó Tổ trưởng tổ CM)", "04 - 06", "Bậc 3 - Bậc 4 (GV Hạng III, II)", "0.15", "Giảm 1 tiết (dạy 16 tiết/tuần)", "Hiệu trưởng THPT Phục Hòa"),

        # DANH MỤC 2: CHUYÊN MÔN, NGHIỆP VỤ
        ("5", "GV-THPT-01", "Giáo viên THPT (Môn học)", "DM2: Chuyên môn, nghiệp vụ (Cốt lõi)", "35 - 45", "Bậc 3 - Bậc 4 (hoặc Bậc 5)", "Không", "17 tiết/tuần", "Hiệu trưởng THPT Phục Hòa"),
        ("6", "TV-THPT-01", "Thư viện viên trường học", "DM2: Chuyên môn, nghiệp vụ (Chuyên ngành)", "01", "Bậc 1 - Bậc 4 (TVV Hạng IV - II)", "Không", "40 giờ/tuần (quản lý thư viện)", "Hiệu trưởng THPT Phục Hòa"),
        ("7", "CNTT-THPT-01", "Ứng dụng CNTT và CĐS", "DM2: Chuyên môn, nghiệp vụ (Kỹ thuật số)", "01", "Bậc 1 - Bậc 4 (CNTT Hạng IV - II)", "Không", "40 giờ/tuần (hệ thống số hóa)", "Hiệu trưởng THPT Phục Hòa"),

        # DANH MỤC 3: HỖ TRỢ, PHỤC VỤ
        ("8", "KT-THPT-01", "Kế toán / Phụ trách KT", "DM3: Hỗ trợ (Tài chính - Ngân sách)", "01", "Bậc 1 - Bậc 4 (KTV Hạng IV - II)", "Phụ cấp KT", "40 giờ/tuần (kế toán viên)", "Sở GD&ĐT / Hiệu trưởng duyệt"),
        ("9", "VT-THPT-01", "Văn thư trường học", "DM3: Hỗ trợ (Hành chính văn phòng)", "01", "Bậc 1 - Bậc 4 (VTV Hạng IV - II)", "Không", "40 giờ/tuần (lưu trữ, công văn)", "Hiệu trưởng THPT Phục Hòa"),
        ("10", "TBTN-THPT-01", "Thiết bị, thí nghiệm", "DM3: Hỗ trợ (Phục vụ thực hành)", "01 - 02", "Bậc 2 - Bậc 3 (ĐTV Hạng III - II)", "Không", "40 giờ/tuần (phòng bộ môn)", "Hiệu trưởng THPT Phục Hòa"),
        ("11", "GVU-THPT-01", "Giáo vụ trường học", "DM3: Hỗ trợ (Hồ sơ học sinh, thi)", "01", "Bậc 2 - Bậc 3 (GVU Hạng III - II)", "Không", "40 giờ/tuần (giáo vụ, nề nếp)", "Hiệu trưởng THPT Phục Hòa"),
        ("12", "YT-THPT-01", "Y tế trường học", "DM3: Hỗ trợ (Chăm sóc sức khỏe)", "01", "Bậc 1 - Bậc 3 (Y sĩ / Điều dưỡng)", "Không", "40 giờ/tuần (y tế học đường)", "Hiệu trưởng THPT Phục Hòa"),
    ]

    for r_i, sdata in enumerate(full_vtvl_summary, 5):
        # Color group
        is_dm1 = int(sdata[0]) <= 4
        is_dm2 = 5 <= int(sdata[0]) <= 7
        row_fill = None
        for c_i, val in enumerate(sdata, 1):
            cell = ws2.cell(row=r_i, column=c_i, value=val)
            cell.font = FONT_BOLD if c_i in [2, 3] else FONT_REGULAR
            cell.border = BORDER_CELL
            cell.alignment = ALIGN_CENTER if c_i in [1, 2, 5, 7, 8] else ALIGN_LEFT

    # ==============================================================================
    # SHEET 3: DANH MỤC 1 - BẢN MÔ TẢ VTVL QUẢN LÝ
    # ==============================================================================
    ws3 = wb.create_sheet(title="3. DM1 - VTVL Quản lý")
    ws3.column_dimensions['A'].width = 8
    ws3.column_dimensions['B'].width = 24
    ws3.column_dimensions['C'].width = 26
    ws3.column_dimensions['D'].width = 38
    ws3.column_dimensions['E'].width = 12
    ws3.column_dimensions['F'].width = 28

    ws3.merge_cells("A1:F1")
    ws3["A1"] = "DANH MỤC 1: BẢN MÔ TẢ VỊ TRÍ VIỆC LÀM VIÊN CHỨC QUẢN LÝ (PHỤ LỤC I)"
    ws3["A1"].font = FONT_TITLE_MAIN
    ws3["A1"].alignment = ALIGN_CENTER

    headers_dm = ["STT", "Vị trí việc làm & Mã số", "Nhiệm vụ quản lý chính", "Nội dung công việc cụ thể & Tiêu chí đo lường", "Tỷ trọng", "Sản phẩm / Kết quả đầu ra"]
    for col_i, h in enumerate(headers_dm, 1):
        cell = ws3.cell(row=3, column=col_i, value=h)
        cell.font = FONT_HEADER_WHITE
        cell.fill = BLUE_HEADER
        cell.alignment = ALIGN_CENTER
        cell.border = BORDER_HEADER

    dm1_data = [
        ("1", "Hiệu trưởng\n(HT-THPT-01)\nPhụ cấp: 0.70\nDạy: 02 tiết/tuần",
         "1. Chiến lược & Kế hoạch phát triển trường",
         "- Xây dựng Chiến lược 5 năm, Kế hoạch giáo dục nhà trường hàng năm theo CT GDPT 2018.\n- Ban hành Bộ 08 Quy chế quản trị nội bộ trước 05/9.\n- 100% cán bộ, giáo viên thực hiện nghiêm túc.",
         "15%",
         "- Chiến lược phát triển trường 2026-2030.\n- Kế hoạch giáo dục năm học.\n- Bộ 08 Quy chế quản trị."),
        ("2", "Hiệu trưởng\n(HT-THPT-01)",
         "2. Quản lý nhân sự & Đánh giá viên chức",
         "- Phân công nhiệm vụ, bổ nhiệm TTCM/TPCM.\n- Chủ trì Hội đồng đánh giá xếp loại viên chức theo NĐ 90/2020 & NĐ 48/2023.\n- Đánh giá Chuẩn nghề nghiệp GV theo TT 20/2018; công bằng, không khiếu nại.",
         "20%",
         "- QĐ phân công nhiệm vụ.\n- QĐ bổ nhiệm TTCM, TPCM.\n- Biên bản họp Hội đồng xếp loại VC.\n- QĐ khen thưởng, kỷ luật."),
        ("3", "Hiệu trưởng\n(HT-THPT-01)",
         "3. Quản lý tài chính & Tài sản công",
         "- Chủ tài khoản, duyệt thu - chi ngân sách theo Luật Ngân sách và Kế toán.\n- Quản lý cơ sở vật chất, tu sửa phòng học, mua sắm trang thiết bị.\n- Quyết toán tài chính minh bạch, đúng luật.",
         "15%",
         "- Dự toán & Báo cáo quyết toán năm.\n- Biên bản kiểm kê tài sản định kỳ.\n- Hồ sơ nghiệm thu CSVC hè."),
        ("4", "Hiệu trưởng\n(HT-THPT-01)",
         "4. Chỉ đạo chuyên môn, kỳ thi & KĐCLGD",
         "- Chỉ đạo kỳ thi tốt nghiệp THPT (mục tiêu đỗ ≥ 98.5%), thi tuyển sinh 10, thi HSG tỉnh.\n- Chỉ đạo kiểm định chất lượng giáo dục và duy trì trường chuẩn quốc gia.",
         "20%",
         "- Báo cáo kết quả kỳ thi tốt nghiệp.\n- Báo cáo kết quả thi HSG tỉnh.\n- Báo cáo tự đánh giá KĐCLGD."),
        ("5", "Phó Hiệu trưởng\n(PHT-THPT-01)\nPhụ cấp: 0.50\nDạy: 04 tiết/tuần",
         "1. Chỉ đạo & điều hành chuyên môn dạy học",
         "- Thẩm định và duyệt Kế hoạch dạy học của các tổ chuyên môn trước 28/8.\n- Chỉ đạo xếp thời khóa biểu khoa học, theo dõi lịch báo giảng, dạy bù, dạy thay.\n- Bảo đảm 100% giáo viên thực hiện đúng tiến độ chương trình.",
         "25%",
         "- Kế hoạch chuyên môn năm học.\n- Thời khóa biểu các đợt đã duyệt.\n- Biên bản thẩm định kế hoạch tổ.\n- Báo cáo sơ kết/tổng kết chuyên môn."),
        ("6", "Phó Hiệu trưởng\n(PHT-THPT-01)",
         "2. Chỉ đạo khảo thí, ôn thi TN & thi HSG",
         "- Xây dựng ma trận đặc tả, ngân hàng đề kiểm tra định kỳ (giữa kỳ, cuối kỳ).\n- Chỉ đạo bồi dưỡng HSG cấp tỉnh, ôn thi tốt nghiệp THPT khối 12; thi thử.\n- Phân tích phổ điểm thi thử tốt nghiệp, đề ra giải pháp hỗ trợ HS yếu.",
         "20%",
         "- Ngân hàng đề kiểm tra định kỳ.\n- Kế hoạch ôn thi TN THPT & HSG.\n- Báo cáo phân tích phổ điểm thi thử.\n- Danh sách học sinh đạt giải HSG tỉnh."),
        ("7", "Tổ trưởng CM\n(TTCM-THPT-01)\nPhụ cấp: 0.25\nGiảm 3 tiết dạy",
         "1. Quản lý toàn diện chuyên môn tổ",
         "- Xây dựng kế hoạch dạy học môn học, phân phối chương trình môn của tổ.\n- Sinh hoạt chuyên môn nghiên cứu bài học ≥ 02 lần/tháng.\n- Dự giờ 2-3 tiết/GV/kỳ; kiểm tra kế hoạch bài dạy và sổ điểm của GV.",
         "40%",
         "- Kế hoạch giáo dục tổ chuyên môn.\n- Biên bản sinh hoạt chuyên môn số hóa.\n- Báo cáo chuyên đề nghiên cứu bài học.\n- Phiếu đánh giá dự giờ GV."),
        ("8", "Tổ phó CM\n(TPCM-THPT-01)\nPhụ cấp: 0.15\nGiảm 1 tiết dạy",
         "1. Quản lý nề nếp, sổ sách & thiết bị",
         "- Theo dõi tiến độ chương trình hàng tuần, đôn đốc lịch báo giảng, dạy bù.\n- Kiểm tra nề nếp sổ sách điện tử, ký số học bạ của GV trong tổ.\n- Theo dõi khai thác thiết bị dạy học, phòng bộ môn, hoạt động STEM.",
         "40%",
         "- Sổ theo dõi tiến độ chương trình.\n- Tập biên bản họp tổ số hóa.\n- Bảng theo dõi ký số học bạ.\n- Sổ mượn trả thiết bị & phòng bộ môn."),
    ]

    for idx, rdata in enumerate(dm1_data, 4):
        for col_c in range(1, 7):
            cell = ws3.cell(row=idx, column=col_c, value=rdata[col_c-1])
            cell.font = FONT_BOLD if col_c in [1, 2, 5] else FONT_REGULAR
            cell.border = BORDER_CELL
            cell.alignment = ALIGN_CENTER if col_c in [1, 5] else ALIGN_LEFT

    # ==============================================================================
    # SHEET 4: DANH MỤC 2 - BẢN MÔ TẢ VTVL CHUYÊN MÔN, NGHIỆP VỤ
    # ==============================================================================
    ws4 = wb.create_sheet(title="4. DM2 - VTVL Chuyên môn")
    ws4.column_dimensions['A'].width = 8
    ws4.column_dimensions['B'].width = 24
    ws4.column_dimensions['C'].width = 26
    ws4.column_dimensions['D'].width = 38
    ws4.column_dimensions['E'].width = 12
    ws4.column_dimensions['F'].width = 28

    ws4.merge_cells("A1:F1")
    ws4["A1"] = "DANH MỤC 2: BẢN MÔ TẢ VỊ TRÍ VIỆC LÀM CHUYÊN MÔN, NGHIỆP VỤ (PHỤ LỤC II)"
    ws4["A1"].font = FONT_TITLE_MAIN
    ws4["A1"].alignment = ALIGN_CENTER

    for col_i, h in enumerate(headers_dm, 1):
        cell = ws4.cell(row=3, column=col_i, value=h)
        cell.font = FONT_HEADER_WHITE
        cell.fill = EMERALD_HEADER
        cell.alignment = ALIGN_CENTER
        cell.border = BORDER_HEADER

    dm2_data = [
        ("1", "Giáo viên THPT (Môn học)\n(GV-THPT-01)\nBậc: 3 - 4 (hoặc 5)\nĐịnh mức: 17 tiết/tuần",
         "1. Giảng dạy & Đánh giá môn học theo CT GDPT 2018",
         "- Thực hiện giảng dạy đủ 17 tiết/tuần theo phân công chuyên môn.\n- Soạn kế hoạch bài dạy (giáo án) chất lượng, áp dụng PPDH tích cực, giáo dục STEM/STEAM.\n- Kiểm tra, chấm bài, đánh giá học sinh đúng thông tư Bộ GD&ĐT; cập nhật điểm số điện tử đúng hạn.",
         "40%",
         "- Kế hoạch bài dạy (giáo án) cá nhân.\n- Sổ điểm, sổ đánh giá học sinh số hóa.\n- Đề kiểm tra định kỳ kèm ma trận, bản đặc tả."),
        ("2", "Giáo viên THPT (Môn học)\n(GV-THPT-01)",
         "2. Công tác chủ nhiệm & Giáo dục đạo đức học sinh",
         "- Quản lý học sinh lớp chủ nhiệm, xây dựng tập thể lớp đoàn kết, kỷ cương.\n- Phối hợp chặt chẽ với cha mẹ học sinh và các giáo viên bộ môn.\n- Ký duyệt học bạ số, nhận xét đánh giá rèn luyện học sinh cuối kỳ/năm học.",
         "25%",
         "- Kế hoạch chủ nhiệm năm học.\n- Sổ theo dõi học sinh lớp chủ nhiệm.\n- Biên bản họp cha mẹ học sinh đầu năm, cuối kỳ.\n- Học bạ số đã ký duyệt."),
        ("3", "Giáo viên THPT (Môn học)\n(GV-THPT-01)",
         "3. Tự học, bồi dưỡng chuyên môn & Ứng dụng AI",
         "- Tham gia sinh hoạt chuyên môn theo NCBL đầy đủ (≥ 2 lần/tháng).\n- Tham gia đầy đủ các đợt tập huấn CT GDPT 2018, bồi dưỡng thường xuyên.\n- Ứng dụng CNTT và công cụ AI hỗ trợ soạn bài, tạo đề trắc nghiệm, bài giảng số.",
         "20%",
         "- Phiếu dự giờ đồng nghiệp (≥ 1 tiết/kỳ).\n- Chứng nhận bồi dưỡng thường xuyên.\n- Kho học liệu số đóng góp cho tổ chuyên môn."),
        ("4", "Giáo viên THPT (Môn học)\n(GV-THPT-01)",
         "4. Bồi dưỡng HSG & Phụ đạo học sinh",
         "- Bồi dưỡng đội tuyển HSG môn học theo phân công của trường/tổ.\n- Phụ đạo học sinh yếu kém, phụ đạo ôn thi tốt nghiệp THPT khối 12.",
         "15%",
         "- Giáo án bồi dưỡng HSG / phụ đạo.\n- Danh sách học sinh dự thi và kết quả đạt giải."),
        ("5", "Thư viện viên trường học\n(TV-THPT-01)\nBậc: 1 - 4\nĐịnh mức: 40 giờ/tuần",
         "1. Quản lý, khai thác thư viện & Phát triển văn hóa đọc",
         "- Quản lý sách, báo, tạp chí, sách giáo khoa, tài liệu tham khảo trong thư viện.\n- Lập kế hoạch bổ sung sách hàng năm; phân loại, biên mục sách số hóa.\n- Tổ chức phục vụ bạn đọc (giáo viên, học sinh); tổ chức Ngày hội đọc sách.\n- Quản trị phần mềm thư viện số, kiểm kê tài sản thư viện định kỳ.",
         "60%",
         "- Sổ đăng ký cá biệt, sổ đăng ký tổng quát.\n- Báo cáo số lượt bạn đọc mượn trả sách.\n- Kế hoạch Ngày hội đọc sách trường học.\n- Biên bản kiểm kê sách thư viện cuối năm."),
        ("6", "Ứng dụng CNTT & CĐS\n(CNTT-THPT-01)\nBậc: 1 - 4\nĐịnh mức: 40 giờ/tuần",
         "1. Quản trị hệ thống công nghệ, mạng & Phần mềm số hóa",
         "- Quản trị hệ thống mạng LAN, Internet, máy chủ, website và cổng thông tin của trường.\n- Quản trị cơ sở dữ liệu ngành, phần mềm quản lý điểm, học bạ điện tử, ký số.\n- Hỗ trợ CBGVNV sử dụng thiết bị CNTT, phòng máy tính, hệ thống thi trắc nghiệm online.\n- Bảo đảm an toàn thông tin, an ninh mạng cho hệ thống nhà trường.",
         "60%",
         "- Báo cáo vận hành cổng thông tin trường.\n- Sổ nhật ký bảo trì hệ thống máy tính, mạng.\n- Kế hoạch chuyển đổi số và bảo đảm an toàn dữ liệu."),
    ]

    for idx, rdata in enumerate(dm2_data, 4):
        for col_c in range(1, 7):
            cell = ws4.cell(row=idx, column=col_c, value=rdata[col_c-1])
            cell.font = FONT_BOLD if col_c in [1, 2, 5] else FONT_REGULAR
            cell.border = BORDER_CELL
            cell.alignment = ALIGN_CENTER if col_c in [1, 5] else ALIGN_LEFT

    # ==============================================================================
    # SHEET 5: DANH MỤC 3 - BẢN MÔ TẢ VTVL HỖ TRỢ
    # ==============================================================================
    ws5 = wb.create_sheet(title="5. DM3 - VTVL Hỗ trợ")
    ws5.column_dimensions['A'].width = 8
    ws5.column_dimensions['B'].width = 24
    ws5.column_dimensions['C'].width = 26
    ws5.column_dimensions['D'].width = 38
    ws5.column_dimensions['E'].width = 12
    ws5.column_dimensions['F'].width = 28

    ws5.merge_cells("A1:F1")
    ws5["A1"] = "DANH MỤC 3: BẢN MÔ TẢ VỊ TRÍ VIỆC LÀM HỖ TRỢ, PHỤC VỤ (PHỤ LỤC III)"
    ws5["A1"].font = FONT_TITLE_MAIN
    ws5["A1"].alignment = ALIGN_CENTER

    for col_i, h in enumerate(headers_dm, 1):
        cell = ws5.cell(row=3, column=col_i, value=h)
        cell.font = FONT_HEADER_WHITE
        cell.fill = AMBER_HEADER
        cell.alignment = ALIGN_CENTER
        cell.border = BORDER_HEADER

    dm3_data = [
        ("1", "Kế toán / Phụ trách KT\n(KT-THPT-01)\nBậc: 1 - 4\nĐịnh mức: 40 giờ/tuần",
         "1. Quản lý tài chính, ngân sách & chế độ tiền lương",
         "- Lập dự toán thu - chi ngân sách nhà nước và các nguồn thu dịch vụ giáo dục hàng năm.\n- Thực hiện thanh toán lương, phụ cấp, bảo hiểm, chế độ chính sách cho CBGVNV và học sinh.\n- Đối chiếu kho bạc nhà nước, lập báo cáo quyết toán tài chính quý/năm.\n- Quản lý sổ sách kế toán, chứng từ thu chi theo Luật Kế toán.",
         "60%",
         "- Bảng thanh toán tiền lương & phụ cấp hàng tháng.\n- Báo cáo quyết toán tài chính ngân sách quý/năm.\n- Sổ cái, sổ chi tiết các tài khoản kế toán.\n- Hồ sơ chứng từ kế toán lưu trữ đúng luật."),
        ("2", "Văn thư trường học\n(VT-THPT-01)\nBậc: 1 - 4\nĐịnh mức: 40 giờ/tuần",
         "1. Quản lý văn bản, con dấu & lưu trữ hồ sơ hành chính",
         "- Tiếp nhận, đăng ký, chuyển giao công văn đi, công văn đến trên hệ thống quản lý văn bản số.\n- Quản lý và sử dụng con dấu của nhà trường đúng quy định pháp luật.\n- Lập hồ sơ lưu trữ hiện hành, giao nộp hồ sơ lưu trữ lịch sử theo quy định.\n- Đánh máy, in sao, phát hành các văn bản hành chính của Hiệu trưởng.",
         "60%",
         "- Sổ đăng ký văn bản đi / đến điện tử.\n- Hồ sơ lưu trữ văn bản lưu trữ năm học.\n- Sổ theo dõi đóng dấu và sử dụng chứng thư số."),
        ("3", "Thiết bị, thí nghiệm\n(TBTN-THPT-01)\nBậc: 2 - 3\nĐịnh mức: 40 giờ/tuần",
         "1. Quản lý phòng thực hành & thiết bị dạy học",
         "- Quản lý, bảo quản thiết bị dạy học, hóa chất, phòng thực hành Vật lý, Hóa học, Sinh học.\n- Chuẩn bị thiết bị, dụng cụ thí nghiệm cho các tiết dạy thực hành theo kế hoạch của giáo viên.\n- Lập sổ theo dõi mượn trả thiết bị; đề xuất thanh lý thiết bị hỏng, mua sắm bổ sung thiết bị mới.\n- Bảo đảm an toàn lao động, phòng chống cháy nổ tại các phòng thí nghiệm.",
         "60%",
         "- Sổ theo dõi sử dụng phòng học bộ môn & mượn trả TBDH.\n- Danh mục thiết bị dạy học theo môn/khối lớp.\n- Biên bản kiểm kê thiết bị dạy học cuối kỳ/năm.\n- Phiếu đề xuất mua sắm hóa chất, thiết bị."),
        ("4", "Giáo vụ trường học\n(GVU-THPT-01)\nBậc: 2 - 3\nĐịnh mức: 40 giờ/tuần",
         "1. Quản lý hồ sơ học sinh, tuyển sinh & kỳ thi",
         "- Quản lý hồ sơ học sinh, sổ đăng bộ, hồ sơ tuyển sinh lớp 10, hồ sơ thi tốt nghiệp THPT.\n- Theo dõi biến động học sinh (chuyển đi, chuyển đến, thôi học), lập danh sách cấp bằng tốt nghiệp.\n- Hỗ trợ công tác tổ chức thi (chuẩn bị phòng thi, số báo danh, danh sách thí sinh).",
         "50%",
         "- Sổ đăng bộ trường THPT số hóa.\n- Hồ sơ tuyển sinh vào lớp 10.\n- Hồ sơ đăng ký dự thi tốt nghiệp THPT của học sinh.\n- Sổ cấp phát văn bằng, chứng chỉ."),
        ("5", "Y tế trường học\n(YT-THPT-01)\nBậc: 1 - 3\nĐịnh mức: 40 giờ/tuần",
         "1. Chăm sóc sức khỏe ban đầu & vệ sinh an toàn trường học",
         "- Khám sức khỏe định kỳ cho học sinh và CBGVNV; sơ cấp cứu tai nạn thương tích học đường.\n- Quản lý tủ thuốc y tế, thuốc thiết yếu và trang thiết bị sơ cứu.\n- Tuyên truyền phòng chống dịch bệnh, kiểm tra vệ sinh an toàn thực phẩm, nước uống học sinh.",
         "50%",
         "- Sổ theo dõi sức khỏe học sinh.\n- Sổ khám chữa bệnh và cấp phát thuốc.\n- Báo cáo công tác y tế trường học hàng kỳ/năm.\n- Kế hoạch phòng chống dịch bệnh học đường."),
    ]

    for idx, rdata in enumerate(dm3_data, 4):
        for col_c in range(1, 7):
            cell = ws5.cell(row=idx, column=col_c, value=rdata[col_c-1])
            cell.font = FONT_BOLD if col_c in [1, 2, 5] else FONT_REGULAR
            cell.border = BORDER_CELL
            cell.alignment = ALIGN_CENTER if col_c in [1, 5] else ALIGN_LEFT

    # ==============================================================================
    # SHEET 6: KHUNG NĂNG LỰC CẢ 3 DANH MỤC (CẤP ĐỘ 1 ĐẾN 5 THEO ĐIỀU 8 NĐ 232)
    # ==============================================================================
    ws6 = wb.create_sheet(title="6. Khung năng lực 3 Nhóm")
    ws6.column_dimensions['A'].width = 8
    ws6.column_dimensions['B'].width = 16
    ws6.column_dimensions['C'].width = 28
    ws6.column_dimensions['D'].width = 15
    ws6.column_dimensions['E'].width = 15
    ws6.column_dimensions['F'].width = 15
    ws6.column_dimensions['G'].width = 32

    ws6.merge_cells("A1:G1")
    ws6["A1"] = "KHUNG NĂNG LỰC TOÀN DIỆN CHO 03 DANH MỤC VỊ TRÍ VIỆC LÀM (CẤP ĐỘ 1 - 5)"
    ws6["A1"].font = FONT_TITLE_MAIN
    ws6["A1"].alignment = ALIGN_CENTER

    headers_knl = ["STT", "Nhóm năng lực", "Tên năng lực cụ thể", "DM1: Quản lý", "DM2: Chuyên môn", "DM3: Hỗ trợ", "Mô tả chuẩn hành vi yêu cầu"]
    for col_i, h in enumerate(headers_knl, 1):
        cell = ws6.cell(row=3, column=col_i, value=h)
        cell.font = FONT_HEADER_WHITE
        cell.fill = PURPLE_HEADER
        cell.alignment = ALIGN_CENTER
        cell.border = BORDER_HEADER

    knl_matrix = [
        ("1", "Năng lực chung", "Phẩm chất chính trị & Đạo đức công vụ", "Cấp độ 4 - 5", "Cấp độ 3 - 4", "Cấp độ 2 - 3", "Tuyệt đối trung thành, gương mẫu, tận tụy phục vụ, tuân thủ pháp luật và đạo đức nghề nghiệp."),
        ("2", "Năng lực chung", "Giao tiếp, ứng xử & Tinh thần phục vụ", "Cấp độ 4 - 5", "Cấp độ 3 - 4", "Cấp độ 2 - 3", "Giao tiếp sư phạm chuẩn mực, tôn trọng đồng nghiệp, phục vụ tận tình học sinh và phụ huynh."),
        ("3", "Năng lực chung", "Đổi mới, sáng tạo & Chuyển đổi số", "Cấp độ 4 - 5", "Cấp độ 3 - 4", "Cấp độ 2 - 3", "Chủ động đề xuất cải tiến quy trình, tích cực ứng dụng công nghệ và kỹ năng số trong công việc."),
        ("4", "Năng lực chung", "Kỹ năng số, CNTT & Công cụ AI", "Cấp độ 4 - 5", "Cấp độ 3 - 4", "Cấp độ 2 - 3", "Khai thác thành thạo phần mềm chuyên ngành, sử dụng công cụ AI tăng năng suất công việc."),
        ("5", "Năng lực quản lý", "Tư duy chiến lược & Lập kế hoạch", "Cấp độ 4 - 5", "Cấp độ 2 - 3", "Cấp độ 1 - 2", "Hoạch định chiến lược phát triển, kế hoạch công tác khoa học, đo lường rõ ràng theo kết quả."),
        ("6", "Năng lực quản lý", "Tổ chức điều hành & Kiểm tra giám sát", "Cấp độ 4 - 5", "Cấp độ 2 - 3", "Cấp độ 1 - 2", "Phân công nhiệm vụ hợp lý, kiểm tra đôn đốc tiến độ, xử lý tình huống phát sinh công bằng."),
        ("7", "Năng lực quản lý", "Quản trị nhân lực & Phát triển đội ngũ", "Cấp độ 4 - 5", "Cấp độ 2 - 3", "Cấp độ 1", "Xây dựng khối đoàn kết nội bộ, động viên khen thưởng, đào tạo bồi dưỡng thế hệ kế cận."),
        ("8", "Năng lực chuyên môn", "Nghiệp vụ sư phạm / Chuyên ngành", "Cấp độ 4 - 5", "Cấp độ 3 - 5", "Cấp độ 2 - 3", "Nắm vững chuyên môn, kiến thức chuyên sâu môn học/nghiệp vụ, làm chủ chương trình đào tạo."),
        ("9", "Năng lực chuyên môn", "Phương pháp dạy học / Kỹ thuật nghiệp vụ", "Cấp độ 4 - 5", "Cấp độ 3 - 5", "Cấp độ 2 - 3", "Vận dụng thành thạo các phương pháp tích cực, kỹ thuật nghiệp vụ hiện đại, tiêu chuẩn quốc gia."),
        ("10", "Năng lực chuyên môn", "Kiểm tra đánh giá & Phân tích dữ liệu", "Cấp độ 4 - 5", "Cấp độ 3 - 4", "Cấp độ 2 - 3", "Đánh giá chính xác phẩm chất năng lực học sinh/kết quả công việc, xử lý dữ liệu tin cậy."),
    ]

    for idx, rdata in enumerate(knl_matrix, 4):
        for col_c in range(1, 8):
            cell = ws6.cell(row=idx, column=col_c, value=rdata[col_c-1])
            cell.font = FONT_BOLD if col_c in [1, 2, 3] else FONT_REGULAR
            cell.border = BORDER_CELL
            cell.alignment = ALIGN_CENTER if col_c in [1, 4, 5, 6] else ALIGN_LEFT

    # ==============================================================================
    # SHEET 7: BỘ TIÊU CHÍ KPI ĐO LƯỜNG ĐẦU RA CHO CẢ 3 DANH MỤC
    # ==============================================================================
    ws7 = wb.create_sheet(title="7. KPI đo lường 3 Nhóm")
    ws7.column_dimensions['A'].width = 8
    ws7.column_dimensions['B'].width = 18
    ws7.column_dimensions['C'].width = 24
    ws7.column_dimensions['D'].width = 28
    ws7.column_dimensions['E'].width = 22
    ws7.column_dimensions['F'].width = 14
    ws7.column_dimensions['G'].width = 14
    ws7.column_dimensions['H'].width = 24

    ws7.merge_cells("A1:H1")
    ws7["A1"] = "BỘ TIÊU CHÍ KPI ĐO LƯỜNG HIỆU QUẢ ĐẦU RA TOÀN BỘ 03 DANH MỤC VTVL"
    ws7["A1"].font = FONT_TITLE_MAIN
    ws7["A1"].alignment = ALIGN_CENTER

    headers_kpi = ["STT", "Danh mục VTVL", "Vị trí áp dụng", "Chỉ số KPI chính", "Mục tiêu định lượng", "Tần suất", "Trọng số", "Mức hoàn thành xuất sắc"]
    for col_i, h in enumerate(headers_kpi, 1):
        cell = ws7.cell(row=3, column=col_i, value=h)
        cell.font = FONT_HEADER_WHITE
        cell.fill = SLATE_HEADER
        cell.alignment = ALIGN_CENTER
        cell.border = BORDER_HEADER

    kpi_all = [
        # DM1
        ("1", "DM1: Quản lý", "Hiệu trưởng", "Tỷ lệ tốt nghiệp THPT toàn trường", "Đạt ≥ 98.5%", "Hàng năm", "25 điểm", "Đạt ≥ 99.5% và điểm TB môn trong top đầu"),
        ("2", "DM1: Quản lý", "Hiệu trưởng", "Quyết toán ngân sách & Tài chính", "100% đúng hạn, kiểm toán tốt", "Quý / Năm", "20 điểm", "Tiết kiệm chi, minh bạch 100%, không sai sót"),
        ("3", "DM1: Quản lý", "Phó Hiệu trưởng", "Tiến độ chương trình & Ký số học bạ", "100% đúng tiến độ quy định", "Tháng / Kỳ", "25 điểm", "Sớm hơn thời hạn quy định 2 ngày"),
        ("4", "DM1: Quản lý", "Tổ trưởng CM", "Sinh hoạt NCBL & Đổi mới PPDH", "≥ 02 chuyên đề / học kỳ", "Hàng tháng", "25 điểm", "Có chuyên đề STEM đột phá, chia sẻ cấp trường"),
        ("5", "DM1: Quản lý", "Tổ phó CM", "Nề nếp sổ sách & Ký số tổ CM", "100% thành viên đúng hạn", "Hàng tháng", "25 điểm", "Không có bất kỳ trường hợp nào nộp trễ"),
        # DM2
        ("6", "DM2: Chuyên môn", "Giáo viên THPT", "Tiến độ giảng dạy & Kế hoạch bài dạy", "100% tiết dạy có giáo án chuẩn", "Hàng tuần", "30 điểm", "Áp dụng phương pháp tích cực và bài học STEM"),
        ("7", "DM2: Chuyên môn", "Giáo viên THPT", "Chất lượng học sinh môn học", "Học sinh đạt chuẩn ≥ 95%", "Học kỳ / Năm", "30 điểm", "Học sinh đạt điểm Giỏi tăng ≥ 10%"),
        ("8", "DM2: Chuyên môn", "Thư viện viên", "Phục vụ bạn đọc & Bổ sung sách", "≥ 1.500 lượt bạn đọc / năm", "Hàng tháng", "30 điểm", "Vượt chỉ tiêu bạn đọc ≥ 20%, tổ chức ngày hội sách"),
        ("9", "DM2: Chuyên môn", "Ứng dụng CNTT", "Vận hành hệ thống mạng & Cổng TT", "Hệ thống thông suốt 99.9%", "Hàng tháng", "30 điểm", "Không xảy ra sự cố an toàn dữ liệu, cập nhật tin bài nhanh"),
        # DM3
        ("10", "DM3: Hỗ trợ", "Kế toán viên", "Chi trả lương & Quyết toán chứng từ", "100% đúng hạn, chính xác tuyệt đối", "Hàng tháng", "35 điểm", "Quyết toán ngân sách không có khoản chi sai phạm"),
        ("11", "DM3: Hỗ trợ", "Văn thư", "Xử lý công văn đi đến & Lưu trữ", "Xử lý trong ngày 100%", "Hàng ngày", "35 điểm", "Không thất lạc văn bản, số hóa văn bản 100%"),
        ("12", "DM3: Hỗ trợ", "Thiết bị, TN", "Chuẩn bị thiết bị cho tiết thực hành", "100% tiết thực hành có thiết bị", "Hàng tuần", "35 điểm", "Không xảy ra tai nạn thí nghiệm, bảo quản tốt 100%"),
        ("13", "DM3: Hỗ trợ", "Giáo vụ", "Quản lý hồ sơ học sinh & Tuyển sinh", "100% hồ sơ chính xác, đúng hạn", "Hàng kỳ", "35 điểm", "Cấp phát văn bằng, học bạ chính xác 100%"),
        ("14", "DM3: Hỗ trợ", "Y tế trường học", "Khám sức khỏe & Sơ cấp cứu", "100% học sinh được theo dõi SK", "Hàng kỳ", "35 điểm", "Không xảy ra dịch bệnh bùng phát trong trường"),
    ]

    for idx, rdata in enumerate(kpi_all, 4):
        for col_c in range(1, 9):
            cell = ws7.cell(row=idx, column=col_c, value=rdata[col_c-1])
            cell.font = FONT_BOLD if col_c in [1, 2, 3, 7] else FONT_REGULAR
            cell.border = BORDER_CELL
            cell.alignment = ALIGN_CENTER if col_c in [1, 2, 6, 7] else ALIGN_LEFT

    # Auto-fit all sheets
    for ws_item in [ws1, ws2, ws3, ws4, ws5, ws6, ws7]:
        auto_fit_columns(ws_item)

    # Save to all target paths
    for p in output_paths:
        os.makedirs(os.path.dirname(os.path.abspath(p)), exist_ok=True)
        wb.save(p)
        print(f"✓ Đã tạo thành công file Excel Đề án 3 Danh mục: {p}")

if __name__ == "__main__":
    paths = [
        r"d:\Du-an-web\web-vtvl-ttcm-tpcm-phuc-hoa\Bang-mo-ta-De-an-VTVL-THPT-Phuc-Hoa-ND232-2026.xlsx",
        r"d:\Du-an-web\web-vtvl-ttcm-tpcm-phuc-hoa\Bang-mo-ta-VTVL-TTCM-TPCM-THPT-Phuc-Hoa.xlsx",
        r"D:\Desktop\Bang-mo-ta-De-an-VTVL-THPT-Phuc-Hoa-ND232-2026.xlsx",
        r"D:\Desktop\Bang-mo-ta-VTVL-TTCM-TPCM-THPT-Phuc-Hoa.xlsx",
    ]
    create_full_vtvl_excel(paths)
