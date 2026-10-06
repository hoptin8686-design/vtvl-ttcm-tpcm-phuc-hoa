# -*- coding: utf-8 -*-
"""
Script sinh file Excel: Đề án Vị trí việc làm Viên chức quản lý - THPT Phục Hòa
Căn cứ Nghị định số 232/2026/NĐ-CP ngày 26/6/2026 của Chính phủ & Thông tư 15/2026/TT-BGDĐT
Bao gồm trọn bộ 04 vị trí:
1. Hiệu trưởng (HT-THPT-01)
2. Phó Hiệu trưởng (PHT-THPT-01)
3. Tổ trưởng chuyên môn (TTCM-THPT-01)
4. Tổ phó chuyên môn (TPCM-THPT-01)
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

def create_vtvl_excel(output_paths):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)  # Remove default sheet

    # Colors
    NAVY_FILL = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")       # #1E3A8A
    BLUE_HEADER = PatternFill(start_color="2563EB", end_color="2563EB", fill_type="solid")     # #2563EB
    TEAL_HEADER = PatternFill(start_color="0D9488", end_color="0D9488", fill_type="solid")     # #0D9488
    PURPLE_HEADER = PatternFill(start_color="6D28D9", end_color="6D28D9", fill_type="solid")   # #6D28D9
    AMBER_HEADER = PatternFill(start_color="D97706", end_color="D97706", fill_type="solid")    # #D97706
    LIGHT_BLUE_FILL = PatternFill(start_color="EFF6FF", end_color="EFF6FF", fill_type="solid") # #EFF6FF
    LIGHT_GRAY_FILL = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid") # #F8FAFC
    SECTION_FILL = PatternFill(start_color="DBEAFE", end_color="DBEAFE", fill_type="solid")    # #DBEAFE
    SECTION_FILL_GREEN = PatternFill(start_color="CCFBF1", end_color="CCFBF1", fill_type="solid")
    SECTION_FILL_PURPLE = PatternFill(start_color="EDE9FE", end_color="EDE9FE", fill_type="solid")

    FONT_FAMILY = "Times New Roman"

    FONT_TITLE_MAIN = Font(name=FONT_FAMILY, size=15, bold=True, color="1E3A8A")
    FONT_TITLE_SUB = Font(name=FONT_FAMILY, size=12, bold=True, color="1F2937")
    FONT_HEADER_WHITE = Font(name=FONT_FAMILY, size=11, bold=True, color="FFFFFF")
    FONT_SECTION = Font(name=FONT_FAMILY, size=11, bold=True, color="1E3A8A")
    FONT_SECTION_PURPLE = Font(name=FONT_FAMILY, size=11, bold=True, color="5B21B6")
    FONT_SECTION_GREEN = Font(name=FONT_FAMILY, size=11, bold=True, color="0F766E")
    FONT_BOLD = Font(name=FONT_FAMILY, size=11, bold=True, color="000000")
    FONT_REGULAR = Font(name=FONT_FAMILY, size=11, bold=False, color="000000")
    FONT_ITALIC = Font(name=FONT_FAMILY, size=10, italic=True, color="4B5563")

    THIN_BORDER_SIDE = Side(border_style="thin", color="CBD5E1")
    MEDIUM_BORDER_SIDE = Side(border_style="medium", color="64748B")
    BORDER_CELL = Border(left=THIN_BORDER_SIDE, right=THIN_BORDER_SIDE, top=THIN_BORDER_SIDE, bottom=THIN_BORDER_SIDE)
    BORDER_HEADER = Border(left=THIN_BORDER_SIDE, right=THIN_BORDER_SIDE, top=MEDIUM_BORDER_SIDE, bottom=MEDIUM_BORDER_SIDE)

    ALIGN_CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ALIGN_LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)
    ALIGN_RIGHT = Alignment(horizontal="right", vertical="center", wrap_text=True)

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
    ws1["A4"] = "ĐỀ ÁN XÂY DỰNG VỊ TRÍ VIỆC LÀM VIÊN CHỨC QUẢN LÝ"
    ws1["A4"].font = FONT_TITLE_MAIN
    ws1["A4"].alignment = ALIGN_CENTER

    ws1.merge_cells("A5:F5")
    ws1["A5"] = "Căn cứ Nghị định số 232/2026/NĐ-CP ngày 26/6/2026 của Chính phủ quy định về vị trí việc làm viên chức"
    ws1["A5"].font = Font(name=FONT_FAMILY, size=11, italic=True)
    ws1["A5"].alignment = ALIGN_CENTER

    ws1.merge_cells("A6:F6")
    ws1["A6"] = "Áp dụng đối với: Hiệu trưởng, Phó Hiệu trưởng, Tổ trưởng chuyên môn, Tổ phó chuyên môn"
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
        ("1", "Nghị định số 232/2026/NĐ-CP", "Chính phủ", "26/06/2026", "01/07/2026", "Quy định về vị trí việc làm viên chức trong đơn vị sự nghiệp công lập (Phụ lục I, II, IV)"),
        ("2", "Thông tư số 15/2026/TT-BGDĐT", "Bộ GD&ĐT", "24/03/2026", "10/05/2026", "Ban hành Điều lệ trường THCS, trường THPT và trường phổ thông có nhiều cấp học"),
        ("3", "Luật Viên chức số 129/2025/QH15", "Quốc hội", "2025", "2026", "Quy định nguyên tắc, quyền, nghĩa vụ, tuyển dụng, sử dụng và quản lý viên chức"),
        ("4", "Thông tư số 04/2021/TT-BGDĐT & TT 08/2023/TT-BGDĐT", "Bộ GD&ĐT", "02/02/2021", "20/03/2021", "Quy định mã số, tiêu chuẩn chức danh nghề nghiệp và bổ nhiệm, xếp lương giáo viên THPT"),
        ("5", "Thông tư số 20/2018/TT-BGDĐT", "Bộ GD&ĐT", "22/08/2018", "10/10/2018", "Ban hành Chuẩn nghề nghiệp giáo viên cơ sở giáo dục phổ thông"),
        ("6", "Thông tư số 14/2018/TT-BGDĐT", "Bộ GD&ĐT", "20/07/2018", "08/09/2018", "Ban hành Chuẩn hiệu trưởng cơ sở giáo dục phổ thông"),
        ("7", "Kế hoạch phát triển GD năm học 2026-2027", "THPT Phục Hòa", "15/08/2026", "01/09/2026", "Kế hoạch phân công nhiệm vụ, biên chế và chỉ tiêu năm học 2026-2027"),
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

    # II. NGUYÊN TẮC VÀ MỤC TIÊU ĐỀ ÁN
    r_start = 19
    ws1.merge_cells(f"A{r_start}:F{r_start}")
    ws1[f"A{r_start}"] = "II. MỤC TIÊU VÀ NGUYÊN TẮC XÂY DỰNG ĐỀ ÁN"
    apply_row_styles(ws1, r_start, font=FONT_SECTION, fill=SECTION_FILL)

    principles = [
        ("1", "Mục tiêu tổng quát", "Xác định rõ ràng, chuẩn hóa và khoa học hệ thống vị trí việc làm viên chức quản lý trường THPT Phục Hòa (Hiệu trưởng, Phó Hiệu trưởng, Tổ trưởng CM, Tổ phó CM); gắn bản mô tả công việc, khung năng lực và chỉ số đo lường hiệu quả (KPI); làm căn cứ phân công nhiệm vụ, đánh giá xếp loại, trả lương và quy hoạch đào tạo bồi dưỡng."),
        ("2", "Tuân thủ quy định NĐ 232/2026", "Áp dụng thống nhất danh mục vị trí việc làm quản lý theo Phụ lục I; xây dựng Bản mô tả và Khung năng lực theo đúng mẫu Phụ lục IV Phần B Nghị định số 232/2026/NĐ-CP của Chính phủ."),
        ("3", "Gắn với CT GDPT 2018", "Tích hợp đầy đủ yêu cầu quản trị trường học hiện đại, phát triển phẩm chất năng lực học sinh, sinh hoạt chuyên môn theo nghiên cứu bài học, đẩy mạnh ứng dụng CNTT, chuyển đổi số và công nghệ AI."),
        ("4", "Phân công rõ việc - Đo lường rõ KPI", "Mỗi vị trí có chức trách rõ ràng, không chồng chéo thẩm quyền; kết quả đầu ra đo lường được bằng sản phẩm, văn bản, hồ sơ quản lý và tỷ lệ hoàn thành nhiệm vụ theo từng tháng, học kỳ và cả năm học."),
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
        for c in range(3, 7):
            ws1.cell(row=r_i, column=c).border = BORDER_CELL

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

    ws1.merge_cells(f"A{r_sig+1}:C{r_sig+1}")
    ws1[f"A{r_sig+1}"] = "(Ký và ghi rõ họ tên)"
    ws1[f"A{r_sig+1}"].font = FONT_ITALIC
    ws1[f"A{r_sig+1}"].alignment = ALIGN_CENTER

    ws1.merge_cells(f"D{r_sig+1}:F{r_sig+1}")
    ws1[f"D{r_sig+1}"] = "(Ký tên, đóng dấu và ghi rõ họ tên)"
    ws1[f"D{r_sig+1}"].font = FONT_ITALIC
    ws1[f"D{r_sig+1}"].alignment = ALIGN_CENTER

    # ==============================================================================
    # SHEET 2: TỔNG HỢP DANH MỤC 04 VỊ TRÍ VIỆC LÀM QUẢN LÝ
    # ==============================================================================
    ws2 = wb.create_sheet(title="2. Tổng hợp 04 VTVL")
    ws2.column_dimensions['A'].width = 8
    ws2.column_dimensions['B'].width = 16
    ws2.column_dimensions['C'].width = 24
    ws2.column_dimensions['D'].width = 18
    ws2.column_dimensions['E'].width = 12
    ws2.column_dimensions['F'].width = 22
    ws2.column_dimensions['G'].width = 16
    ws2.column_dimensions['H'].width = 14
    ws2.column_dimensions['I'].width = 30

    ws2.merge_cells("A1:I1")
    ws2["A1"] = "DANH MỤC VỊ TRÍ VIỆC LÀM VIÊN CHỨC QUẢN LÝ TRƯỜNG THPT PHỤC HÒA"
    ws2["A1"].font = FONT_TITLE_MAIN
    ws2["A1"].alignment = ALIGN_CENTER

    ws2.merge_cells("A2:I2")
    ws2["A2"] = "Thực hiện theo Nghị định số 232/2026/NĐ-CP (Phụ lục I) và Thông tư 15/2026/TT-BGDĐT"
    ws2["A2"].font = Font(name=FONT_FAMILY, size=11, italic=True)
    ws2["A2"].alignment = ALIGN_CENTER

    headers_summary = [
        "STT", "Mã số VTVL", "Tên Vị trí việc làm", "Nhóm vị trí (NĐ 232)", "Số lượng",
        "Bậc nghề nghiệp sử dụng", "Phụ cấp chức vụ", "Định mức dạy", "Thẩm quyền bổ nhiệm / Phê duyệt"
    ]
    for col_i, h in enumerate(headers_summary, 1):
        cell = ws2.cell(row=4, column=col_i, value=h)
        cell.font = FONT_HEADER_WHITE
        cell.fill = NAVY_FILL
        cell.alignment = ALIGN_CENTER
        cell.border = BORDER_HEADER

    summary_vtvl_data = [
        ("1", "HT-THPT-01", "Hiệu trưởng", "Viên chức quản lý (Người đứng đầu ĐVSNCL - Mục 1 Phụ lục I)", "01",
         "Bậc 4 - Bậc 5 (GV Hạng II, Hạng I)", "0.70", "02 tiết/tuần", "Giám đốc Sở Giáo dục và Đào tạo tỉnh Cao Bằng"),
        ("2", "PHT-THPT-01", "Phó Hiệu trưởng", "Viên chức quản lý (Cấp phó người đứng đầu - Mục 2 Phụ lục I)", "02",
         "Bậc 3 - Bậc 5 (GV Hạng III, II, I)", "0.50", "04 tiết/tuần", "Giám đốc Sở Giáo dục và Đào tạo tỉnh Cao Bằng"),
        ("3", "TTCM-THPT-01", "Tổ trưởng chuyên môn", "Viên chức quản lý (Tổ trưởng tổ CM - Mục 7 Phụ lục I)", "04 - 06",
         "Bậc 3 - Bậc 4 (GV Hạng III, II)", "0.25", "Giảm 3 tiết/tuần", "Hiệu trưởng trường THPT Phục Hòa bổ nhiệm"),
        ("4", "TPCM-THPT-01", "Tổ phó chuyên môn", "Viên chức quản lý (Phó Tổ trưởng tổ CM - Mục 8 Phụ lục I)", "04 - 06",
         "Bậc 3 - Bậc 4 (GV Hạng III, II)", "0.15", "Giảm 1 tiết/tuần", "Hiệu trưởng trường THPT Phục Hòa bổ nhiệm"),
    ]

    for r_i, sdata in enumerate(summary_vtvl_data, 5):
        for c_i, val in enumerate(sdata, 1):
            cell = ws2.cell(row=r_i, column=c_i, value=val)
            cell.font = FONT_BOLD if c_i in [2, 3] else FONT_REGULAR
            cell.border = BORDER_CELL
            cell.alignment = ALIGN_CENTER if c_i in [1, 2, 4, 5, 7, 8] else ALIGN_LEFT

    # Thống kê cơ cấu và tỷ lệ theo Điều 10 Nghị định 232/2026/NĐ-CP
    r_stat = 11
    ws2.merge_cells(f"A{r_stat}:I{r_stat}")
    ws2[f"A{r_stat}"] = "III. CƠ CẤU VÀ TỶ LỆ VIÊN CHỨC BỐ TRÍ THEO BẬC NGHỀ NGHIỆP (ĐIỀU 10 NGHỊ ĐỊNH 232/2026/NĐ-CP)"
    apply_row_styles(ws2, r_stat, font=FONT_SECTION, fill=SECTION_FILL)

    ratio_rows = [
        ("1", "Vị trí có bậc sử dụng cao nhất là Bậc 5 (Hiệu trưởng, Phó Hiệu trưởng)",
         "- Bậc 5: Không vượt quá 30% tổng số biên chế vị trí.\n- Bậc 4: Không vượt quá 40% tổng số biên chế vị trí.\n- Bậc 3: Tỷ lệ còn lại theo nhu cầu thực tế và phân bổ của nhà trường."),
        ("2", "Vị trí có bậc sử dụng cao nhất là Bậc 4 (Tổ trưởng CM, Tổ phó CM)",
         "- Bậc 4: Không vượt quá 40% tổng số biên chế chức danh.\n- Bậc 3: Tỷ lệ còn lại (tối thiểu 60%) bố trí viên chức đạt chuẩn chuyên môn."),
        ("3", "Nguyên tắc xếp lương và phụ cấp chức vụ (Khoản 4 Điều 7)",
         "Viên chức đảm nhiệm vị trí việc làm quản lý được xếp lương theo bậc nghề nghiệp của chức danh chuyên môn (Giáo viên THPT) đang giữ và hưởng phụ cấp chức vụ lãnh đạo tương ứng: Hiệu trưởng (0.70), Phó Hiệu trưởng (0.50), Tổ trưởng CM (0.25), Tổ phó CM (0.15)."),
    ]

    for r_i, rdata in enumerate(ratio_rows, r_stat+1):
        ws2.cell(row=r_i, column=1, value=rdata[0]).alignment = ALIGN_CENTER
        ws2.cell(row=r_i, column=1).border = BORDER_CELL
        ws2.merge_cells(f"B{r_i}:C{r_i}")
        ws2.cell(row=r_i, column=2, value=rdata[1]).font = FONT_BOLD
        ws2.cell(row=r_i, column=2).alignment = ALIGN_LEFT
        for c in range(2, 4): ws2.cell(row=r_i, column=c).border = BORDER_CELL
        ws2.merge_cells(f"D{r_i}:I{r_i}")
        ws2.cell(row=r_i, column=4, value=rdata[2]).font = FONT_REGULAR
        ws2.cell(row=r_i, column=4).alignment = ALIGN_LEFT
        for c in range(4, 10): ws2.cell(row=r_i, column=c).border = BORDER_CELL

    # Signatures
    r_sig2 = r_stat + len(ratio_rows) + 2
    ws2.merge_cells(f"A{r_sig2}:D{r_sig2}")
    ws2[f"A{r_sig2}"] = "NGƯỜI LẬP BIỂU"
    ws2[f"A{r_sig2}"].font = FONT_BOLD
    ws2[f"A{r_sig2}"].alignment = ALIGN_CENTER

    ws2.merge_cells(f"E{r_sig2}:I{r_sig2}")
    ws2[f"E{r_sig2}"] = "HIỆU TRƯỞNG DUYỆT"
    ws2[f"E{r_sig2}"].font = FONT_BOLD
    ws2[f"E{r_sig2}"].alignment = ALIGN_CENTER

    # ==============================================================================
    # SHEET 3: BẢN MÔ TẢ VTVL HIỆU TRƯỞNG (HT-THPT-01)
    # ==============================================================================
    ws3 = wb.create_sheet(title="3. VTVL Hiệu trưởng")
    ws3.column_dimensions['A'].width = 8
    ws3.column_dimensions['B'].width = 26
    ws3.column_dimensions['C'].width = 38
    ws3.column_dimensions['D'].width = 24
    ws3.column_dimensions['E'].width = 14
    ws3.column_dimensions['F'].width = 28

    ws3.merge_cells("A1:F1")
    ws3["A1"] = "BẢN MÔ TẢ CÔNG VIỆC VÀ KHUNG NĂNG LỰC VỊ TRÍ VIỆC LÀM QUẢN LÝ"
    ws3["A1"].font = FONT_TITLE_MAIN
    ws3["A1"].alignment = ALIGN_CENTER

    ws3.merge_cells("A2:F2")
    ws3["A2"] = "VỊ TRÍ: HIỆU TRƯỞNG (MÃ SỐ: HT-THPT-01)"
    ws3["A2"].font = Font(name=FONT_FAMILY, size=13, bold=True, color="1E3A8A")
    ws3["A2"].alignment = ALIGN_CENTER

    ws3.merge_cells("A3:F3")
    ws3["A3"] = "Theo Mẫu Phụ lục IV (Phần B) - Nghị định số 232/2026/NĐ-CP ngày 26/6/2026 của Chính phủ"
    ws3["A3"].font = FONT_ITALIC
    ws3["A3"].alignment = ALIGN_CENTER

    # I. THÔNG TIN CHUNG
    ws3.merge_cells("A5:F5")
    ws3["A5"] = "I. THÔNG TIN CHUNG VỀ VỊ TRÍ VIỆC LÀM"
    apply_row_styles(ws3, 5, font=FONT_SECTION, fill=SECTION_FILL)

    ht_info = [
        ("1.1", "Tên vị trí việc làm:", "Hiệu trưởng", "1.2", "Mã số vị trí:", "HT-THPT-01"),
        ("1.3", "Nhóm vị trí việc làm:", "Viên chức quản lý (Người đứng đầu ĐVSNCL)", "1.4", "Cấp quản lý:", "Lãnh đạo đơn vị trường THPT"),
        ("1.5", "Bậc nghề nghiệp áp dụng:", "Bậc 4 - Bậc 5 (GV THPT Hạng II, Hạng I)", "1.6", "Cơ quan cấp trên trực tiếp:", "Sở Giáo dục và Đào tạo tỉnh Cao Bằng"),
        ("1.7", "Lĩnh vực hoạt động:", "Giáo dục trung học phổ thông, quản trị trường học", "1.8", "Số lượng người đảm nhiệm:", "01 người"),
        ("1.9", "Phụ cấp chức vụ lãnh đạo:", "Hệ số 0.70", "1.10", "Định mức tiết giảng dạy:", "02 tiết/tuần (theo quy định Bộ GD&ĐT)"),
    ]
    for idx, rdata in enumerate(ht_info, 6):
        ws3.cell(row=idx, column=1, value=rdata[0]).alignment = ALIGN_CENTER
        ws3.cell(row=idx, column=2, value=rdata[1]).font = FONT_BOLD
        ws3.cell(row=idx, column=3, value=rdata[2]).font = FONT_REGULAR
        ws3.cell(row=idx, column=4, value=rdata[3]).alignment = ALIGN_CENTER
        ws3.cell(row=idx, column=5, value=rdata[4]).font = FONT_BOLD
        ws3.cell(row=idx, column=6, value=rdata[5]).font = FONT_REGULAR
        for col_c in range(1, 7): ws3.cell(row=idx, column=col_c).border = BORDER_CELL

    # II. MỤC TIÊU VỊ TRÍ
    r_ht_obj = 12
    ws3.merge_cells(f"A{r_ht_obj}:F{r_ht_obj}")
    ws3[f"A{r_ht_obj}"] = "II. MỤC TIÊU CỦA VỊ TRÍ VIỆC LÀM"
    apply_row_styles(ws3, r_ht_obj, font=FONT_SECTION, fill=SECTION_FILL)

    ws3.merge_cells(f"A{r_ht_obj+1}:F{r_ht_obj+1}")
    ws3[f"A{r_ht_obj+1}"] = "Lãnh đạo, quản lý và điều hành toàn diện mọi hoạt động của trường THPT Phục Hòa theo đúng chủ trương đường lối của Đảng, chính sách pháp luật của Nhà nước và Điều lệ trường THPT; chịu trách nhiệm trước Giám đốc Sở GD&ĐT, UBND tỉnh và trước pháp luật về việc hoàn thành thắng lợi các nhiệm vụ chính trị, bảo đảm chất lượng giáo dục toàn diện, quản trị tài chính, nhân sự, cơ sở vật chất và định hướng chiến lược phát triển nhà trường trong thời kỳ kỷ nguyên số và công nghệ AI."
    ws3[f"A{r_ht_obj+1}"].font = FONT_REGULAR
    ws3[f"A{r_ht_obj+1}"].alignment = ALIGN_LEFT
    apply_row_styles(ws3, r_ht_obj+1, border=BORDER_CELL)

    # III. CÔNG VIỆC, KẾT QUẢ, SẢN PHẨM QUẢN LÝ (THEO PHỤ LỤC IV PHẦN B)
    r_ht_duties = 15
    ws3.merge_cells(f"A{r_ht_duties}:F{r_ht_duties}")
    ws3[f"A{r_ht_duties}"] = "III. CÔNG VIỆC, KẾT QUẢ, SẢN PHẨM QUẢN LÝ VÀ CHUYÊN MÔN (ĐÍNH KÈM TỶ TRỌNG & KPI)"
    apply_row_styles(ws3, r_ht_duties, font=FONT_SECTION, fill=SECTION_FILL)

    headers_duties = ["STT", "Nhiệm vụ quản lý cốt lõi", "Hoạt động công việc cụ thể", "Tiêu chí đánh giá đo lường (KPI)", "Tỷ trọng (%)", "Sản phẩm / Kết quả đầu ra"]
    for col_i, h in enumerate(headers_duties, 1):
        cell = ws3.cell(row=r_ht_duties+1, column=col_i, value=h)
        cell.font = FONT_HEADER_WHITE
        cell.fill = BLUE_HEADER
        cell.alignment = ALIGN_CENTER
        cell.border = BORDER_HEADER

    ht_duties = [
        ("1", "Xây dựng chiến lược & Kế hoạch phát triển trường học",
         "- Xây dựng Chiến lược phát triển trường giai đoạn 5 năm.\n- Ban hành Kế hoạch giáo dục nhà trường hàng năm theo CT GDPT 2018.\n- Xây dựng và ban hành Quy chế làm việc, Quy chế chi tiêu nội bộ, Quy chế dân chủ, Quy chế thi đua khen thưởng.",
         "- Kế hoạch năm học ban hành trước khai giảng (trước 05/9).\n- 100% cán bộ, giáo viên, nhân viên thực hiện nghiêm túc.\n- Được Sở GD&ĐT phê duyệt.",
         "15%",
         "- Chiến lược phát triển trường 2026-2030.\n- Kế hoạch giáo dục nhà trường 2026-2027.\n- Bộ 08 Quy chế quản trị nội bộ nhà trường."),
        ("2", "Quản lý tổ chức bộ máy, nhân sự & đánh giá viên chức",
         "- Quyết định phân công nhiệm vụ cho Phó Hiệu trưởng, TTCM, TPCM, GV, NV.\n- Chủ trì tuyển dụng (theo phân cấp), tiếp nhận, bổ nhiệm TTCM, TPCM.\n- Chủ trì đánh giá, xếp loại viên chức theo NĐ 90/2020 & NĐ 48/2023; đánh giá Chuẩn nghề nghiệp GV theo TT 20/2018; quy hoạch đào tạo bồi dưỡng.",
         "- Đúng quy trình, dân chủ, khách quan, minh bạch.\n- 100% CBGVNV được đánh giá đúng hạn, không có khiếu nại.\n- Khen thưởng, kỷ luật kịp thời, chính xác.",
         "20%",
         "- Quyết định phân công nhiệm vụ năm học.\n- Quyết định bổ nhiệm TTCM, TPCM.\n- Biên bản họp Hội đồng đánh giá xếp loại viên chức.\n- Quyết định công nhận kết quả xếp loại."),
        ("3", "Quản lý tài chính, ngân sách & tài sản công",
         "- Chủ tài khoản, chịu trách nhiệm quản lý thu - chi ngân sách nhà nước và các khoản thu dịch vụ giáo dục đúng Luật Ngân sách, Luật Kế toán.\n- Quản lý, sử dụng hiệu quả tài sản công, đất đai, cơ sở vật chất, phòng học bộ môn, thư viện.\n- Lập kế hoạch tu sửa hè, mua sắm bổ sung thiết bị dạy học.",
         "- Quyết toán tài chính minh bạch, không sai phạm, kiểm toán đạt chuẩn.\n- Cơ sở vật chất khang trang, an toàn tuyệt đối, phục vụ tốt dạy học.",
         "15%",
         "- Dự toán và Báo cáo quyết toán tài chính quý/năm.\n- Biên bản kiểm kê tài sản công định kỳ.\n- Hồ sơ tu sửa, nghiệm thu CSVC và mua sắm thiết bị."),
        ("4", "Chỉ đạo chuyên môn dạy học, kỳ thi tốt nghiệp & KĐCLGD",
         "- Chỉ đạo triển khai toàn diện Chương trình GDPT 2018; duyệt kế hoạch các tổ.\n- Chỉ đạo kỳ thi tốt nghiệp THPT, kỳ thi tuyển sinh vào 10, thi HSG cấp tỉnh.\n- Chỉ đạo công tác kiểm định chất lượng giáo dục và duy trì trường đạt chuẩn quốc gia.",
         "- Tỷ lệ đỗ tốt nghiệp THPT đạt từ 98% trở lên.\n- Đạt chỉ tiêu học sinh giỏi cấp tỉnh.\n- Hồ sơ kiểm định chất lượng giáo dục đạt Mức độ 2 trở lên.",
         "20%",
         "- Kế hoạch chỉ đạo chuyên môn năm học.\n- Quyết định thành lập các hội đồng thi.\n- Báo cáo kết quả kỳ thi tốt nghiệp THPT.\n- Báo cáo tự đánh giá KĐCLGD nhà trường."),
        ("5", "Chuyển đổi số, an ninh trường học & đối ngoại liên ngành",
         "- Chỉ đạo chuyển đổi số toàn diện: học bạ số, sổ điểm điện tử, ứng dụng AI trong giảng dạy và quản trị.\n- Bảo đảm an ninh trật tự trường học, an toàn giao thông, PCCC, phòng chống bạo lực học đường.\n- Phối hợp với Đảng ủy, chính quyền địa phương và Ban đại diện cha mẹ học sinh.",
         "- 100% hồ sơ học bạ, sổ sách điện tử được ký số đúng tiến độ.\n- Trường học an toàn, không có tai nạn thương tích hoặc vi phạm pháp luật.",
         "15%",
         "- Kế hoạch chuyển đổi số và ứng dụng CNTT.\n- Biên bản phối hợp an ninh trật tự với Công an.\n- Nghị quyết Hội nghị phối hợp với Ban đại diện CMHS."),
        ("6", "Trực tiếp tham gia giảng dạy bộ môn",
         "- Thực hiện giảng dạy 02 tiết/tuần theo quy định của Bộ GD&ĐT.\n- Nắm bắt thực tiễn việc học tập của học sinh và phương pháp của giáo viên.",
         "- Giảng dạy đủ số tiết quy định.\n- Hồ sơ giáo án đạt chuẩn, chất lượng giờ dạy tốt.",
         "15%",
         "- Giáo án điện tử cá nhân.\n- Sổ theo dõi đánh giá học sinh.\n- Kết quả kiểm tra môn giảng dạy."),
    ]

    for idx, rdata in enumerate(ht_duties, r_ht_duties+2):
        ws3.cell(row=idx, column=1, value=rdata[0]).alignment = ALIGN_CENTER
        ws3.cell(row=idx, column=2, value=rdata[1]).font = FONT_BOLD
        ws3.cell(row=idx, column=2).alignment = ALIGN_LEFT
        ws3.cell(row=idx, column=3, value=rdata[2]).font = FONT_REGULAR
        ws3.cell(row=idx, column=3).alignment = ALIGN_LEFT
        ws3.cell(row=idx, column=4, value=rdata[3]).font = FONT_REGULAR
        ws3.cell(row=idx, column=4).alignment = ALIGN_LEFT
        ws3.cell(row=idx, column=5, value=rdata[4]).font = FONT_BOLD
        ws3.cell(row=idx, column=5).alignment = ALIGN_CENTER
        ws3.cell(row=idx, column=6, value=rdata[5]).font = FONT_REGULAR
        ws3.cell(row=idx, column=6).alignment = ALIGN_LEFT
        for col_c in range(1, 7): ws3.cell(row=idx, column=col_c).border = BORDER_CELL

    # IV. KHUNG NĂNG LỰC & YÊU CẦU TIÊU CHUẨN HIỆU TRƯỞNG
    r_ht_req = r_ht_duties + len(ht_duties) + 3
    ws3.merge_cells(f"A{r_ht_req}:F{r_ht_req}")
    ws3[f"A{r_ht_req}"] = "IV. YÊU CẦU TIÊU CHUẨN, NĂNG LỰC VỊ TRÍ HIỆU TRƯỞNG (CHUẨN BẬC 4 - BẬC 5)"
    apply_row_styles(ws3, r_ht_req, font=FONT_SECTION, fill=SECTION_FILL)

    ht_reqs = [
        ("1", "Trình độ đào tạo chuyên môn", "Có bằng Cử nhân trở lên thuộc ngành đào tạo giáo viên (hoặc có bằng Thạc sĩ Quản lý giáo dục / chuyên ngành phù hợp).", "Bằng cử nhân/Thạc sĩ; Bảng điểm"),
        ("2", "Tiêu chuẩn CDNN & Bồi dưỡng", "Bổ nhiệm chức danh nghề nghiệp Giáo viên THPT Hạng II hoặc Hạng I; có chứng chỉ bồi dưỡng CBQL giáo dục; có chứng chỉ bồi dưỡng CDNN.", "Quyết định bổ nhiệm CDNN; Chứng chỉ CBQL"),
        ("3", "Năng lực quản trị trường học (Cấp độ 4-5)", "Năng lực tư duy chiến lược, quản trị nhân sự, điều hành tổ chức, quản trị tài chính, giải quyết khủng hoảng và quan hệ cộng đồng.", "Phiếu đánh giá Chuẩn Hiệu trưởng (Mức Tốt); Quyết định bổ nhiệm"),
        ("4", "Năng lực chuyển đổi số & AI", "Tiên phong chuyển đổi số trong giáo dục, ứng dụng công nghệ quản trị nhà trường thông minh, chỉ đạo ứng dụng AI trong giảng dạy và học tập.", "Sản phẩm chuyển đổi số thực tế của nhà trường"),
        ("5", "Lý luận chính trị & Đạo đức", "Có trình độ Trung cấp lý luận chính trị trở lên; đảng viên Đảng Cộng sản Việt Nam; phẩm chất chính trị kiên định, đạo đức liêm chính mẫu mực.", "Bằng Trung cấp/Cao cấp LLCT; Bản kiểm điểm đảng viên"),
    ]
    for idx, rdata in enumerate(ht_reqs, r_ht_req+1):
        ws3.cell(row=idx, column=1, value=rdata[0]).alignment = ALIGN_CENTER
        ws3.cell(row=idx, column=2, value=rdata[1]).font = FONT_BOLD
        ws3.cell(row=idx, column=2).alignment = ALIGN_LEFT
        ws3.merge_cells(f"C{idx}:D{idx}")
        ws3.cell(row=idx, column=3, value=rdata[2]).font = FONT_REGULAR
        ws3.cell(row=idx, column=3).alignment = ALIGN_LEFT
        for c in range(3, 5): ws3.cell(row=idx, column=c).border = BORDER_CELL
        ws3.merge_cells(f"E{idx}:F{idx}")
        ws3.cell(row=idx, column=5, value=rdata[3]).font = FONT_REGULAR
        ws3.cell(row=idx, column=5).alignment = ALIGN_LEFT
        for c in range(5, 7): ws3.cell(row=idx, column=c).border = BORDER_CELL
        for c in range(1, 3): ws3.cell(row=idx, column=c).border = BORDER_CELL

    # Signatures
    r_sig3 = r_ht_req + len(ht_reqs) + 2
    ws3.merge_cells(f"A{r_sig3}:C{r_sig3}")
    ws3[f"A{r_sig3}"] = "HIỆU TRƯỞNG"
    ws3[f"A{r_sig3}"].font = FONT_BOLD
    ws3[f"A{r_sig3}"].alignment = ALIGN_CENTER

    ws3.merge_cells(f"D{r_sig3}:F{r_sig3}")
    ws3[f"D{r_sig3}"] = "GIÁM ĐỐC SỞ GD&ĐT CAO BẰNG PHÊ DUYỆT"
    ws3[f"D{r_sig3}"].font = FONT_BOLD
    ws3[f"D{r_sig3}"].alignment = ALIGN_CENTER

    # ==============================================================================
    # SHEET 4: BẢN MÔ TẢ VTVL PHÓ HIỆU TRƯỞNG (PHT-THPT-01)
    # ==============================================================================
    ws4 = wb.create_sheet(title="4. VTVL Phó Hiệu trưởng")
    ws4.column_dimensions['A'].width = 8
    ws4.column_dimensions['B'].width = 26
    ws4.column_dimensions['C'].width = 38
    ws4.column_dimensions['D'].width = 24
    ws4.column_dimensions['E'].width = 14
    ws4.column_dimensions['F'].width = 28

    ws4.merge_cells("A1:F1")
    ws4["A1"] = "BẢN MÔ TẢ CÔNG VIỆC VÀ KHUNG NĂNG LỰC VỊ TRÍ VIỆC LÀM QUẢN LÝ"
    ws4["A1"].font = FONT_TITLE_MAIN
    ws4["A1"].alignment = ALIGN_CENTER

    ws4.merge_cells("A2:F2")
    ws4["A2"] = "VỊ TRÍ: PHÓ HIỆU TRƯỞNG (MÃ SỐ: PHT-THPT-01)"
    ws4["A2"].font = Font(name=FONT_FAMILY, size=13, bold=True, color="1E3A8A")
    ws4["A2"].alignment = ALIGN_CENTER

    ws4.merge_cells("A3:F3")
    ws4["A3"] = "Theo Mẫu Phụ lục IV (Phần B) - Nghị định số 232/2026/NĐ-CP ngày 26/6/2026 của Chính phủ"
    ws4["A3"].font = FONT_ITALIC
    ws4["A3"].alignment = ALIGN_CENTER

    # I. THÔNG TIN CHUNG
    ws4.merge_cells("A5:F5")
    ws4["A5"] = "I. THÔNG TIN CHUNG VỀ VỊ TRÍ VIỆC LÀM"
    apply_row_styles(ws4, 5, font=FONT_SECTION, fill=SECTION_FILL)

    pht_info = [
        ("1.1", "Tên vị trí việc làm:", "Phó Hiệu trưởng", "1.2", "Mã số vị trí:", "PHT-THPT-01"),
        ("1.3", "Nhóm vị trí việc làm:", "Viên chức quản lý (Cấp phó người đứng đầu ĐVSNCL)", "1.4", "Cấp quản lý:", "Lãnh đạo đơn vị trường THPT"),
        ("1.5", "Bậc nghề nghiệp áp dụng:", "Bậc 3 - Bậc 5 (GV THPT Hạng III, II, I)", "1.6", "Người quản lý trực tiếp:", "Hiệu trưởng trường THPT Phục Hòa"),
        ("1.7", "Lĩnh vực hoạt động:", "Chuyên môn giáo dục, kiểm định chất lượng, khảo thí, nề nếp học sinh", "1.8", "Số lượng người đảm nhiệm:", "02 người (theo quy định trường hạng II)"),
        ("1.9", "Phụ cấp chức vụ lãnh đạo:", "Hệ số 0.50", "1.10", "Định mức tiết giảng dạy:", "04 tiết/tuần (theo quy định Bộ GD&ĐT)"),
    ]
    for idx, rdata in enumerate(pht_info, 6):
        ws4.cell(row=idx, column=1, value=rdata[0]).alignment = ALIGN_CENTER
        ws4.cell(row=idx, column=2, value=rdata[1]).font = FONT_BOLD
        ws4.cell(row=idx, column=3, value=rdata[2]).font = FONT_REGULAR
        ws4.cell(row=idx, column=4, value=rdata[3]).alignment = ALIGN_CENTER
        ws4.cell(row=idx, column=5, value=rdata[4]).font = FONT_BOLD
        ws4.cell(row=idx, column=6, value=rdata[5]).font = FONT_REGULAR
        for col_c in range(1, 7): ws4.cell(row=idx, column=col_c).border = BORDER_CELL

    # II. MỤC TIÊU VỊ TRÍ
    r_pht_obj = 12
    ws4.merge_cells(f"A{r_pht_obj}:F{r_pht_obj}")
    ws4[f"A{r_pht_obj}"] = "II. MỤC TIÊU CỦA VỊ TRÍ VIỆC LÀM"
    apply_row_styles(ws4, r_pht_obj, font=FONT_SECTION, fill=SECTION_FILL)

    ws4.merge_cells(f"A{r_pht_obj+1}:F{r_pht_obj+1}")
    ws4[f"A{r_pht_obj+1}"] = "Giúp Hiệu trưởng quản lý, chỉ đạo và điều hành các mảng công tác được phân công phụ trách (chuyên môn dạy học, khảo thí kiểm định, chuyển đổi số, nề nếp học sinh, cơ sở vật chất); trực tiếp chỉ đạo nâng cao chất lượng dạy và học; điều hành toàn bộ công việc của nhà trường khi được Hiệu trưởng ủy quyền và chịu trách nhiệm trước Hiệu trưởng và pháp luật về kết quả thực hiện nhiệm vụ."
    ws4[f"A{r_pht_obj+1}"].font = FONT_REGULAR
    ws4[f"A{r_pht_obj+1}"].alignment = ALIGN_LEFT
    apply_row_styles(ws4, r_pht_obj+1, border=BORDER_CELL)

    # III. CÔNG VIỆC, KẾT QUẢ, SẢN PHẨM QUẢN LÝ
    r_pht_duties = 15
    ws4.merge_cells(f"A{r_pht_duties}:F{r_pht_duties}")
    ws4[f"A{r_pht_duties}"] = "III. CÔNG VIỆC, KẾT QUẢ, SẢN PHẨM QUẢN LÝ VÀ CHUYÊN MÔN (ĐÍNH KÈM TỶ TRỌNG & KPI)"
    apply_row_styles(ws4, r_pht_duties, font=FONT_SECTION, fill=SECTION_FILL)

    for col_i, h in enumerate(headers_duties, 1):
        cell = ws4.cell(row=r_pht_duties+1, column=col_i, value=h)
        cell.font = FONT_HEADER_WHITE
        cell.fill = BLUE_HEADER
        cell.alignment = ALIGN_CENTER
        cell.border = BORDER_HEADER

    pht_duties = [
        ("1", "Chỉ đạo & điều hành công tác chuyên môn dạy học",
         "- Thẩm định và duyệt Kế hoạch dạy học của các tổ chuyên môn, kế hoạch giáo dục môn học theo CT GDPT 2018.\n- Chỉ đạo xây dựng thời khóa biểu, lịch báo giảng, kế hoạch dạy bù, dạy thay, kiểm tra tiến độ chương trình.\n- Chỉ đạo đổi mới phương pháp dạy học, giáo dục STEM/STEAM, sinh hoạt chuyên môn theo nghiên cứu bài học.",
         "- Thời khóa biểu khoa học, duyệt kế hoạch tổ trước 28/8.\n- 100% giáo viên thực hiện đúng tiến độ chương trình, không cắt xén.\n- Tổ chức tối thiểu 02 chuyên đề cấp trường/học kỳ.",
         "25%",
         "- Kế hoạch giáo dục chuyên môn năm học.\n- Thời khóa biểu các đợt đã duyệt.\n- Biên bản thẩm định kế hoạch tổ chuyên môn.\n- Báo cáo sơ kết/tổng kết chuyên môn học kỳ."),
        ("2", "Chỉ đạo công tác khảo thí, ôn thi TN THPT & bồi dưỡng HSG",
         "- Chỉ đạo xây dựng ma trận, bản đặc tả và ngân hàng đề kiểm tra định kỳ (giữa kỳ, cuối kỳ) toàn trường.\n- Xây dựng kế hoạch bồi dưỡng HSG, kế hoạch ôn thi tốt nghiệp THPT cho khối 12.\n- Tổ chức các đợt thi thử tốt nghiệp THPT, phân tích phổ điểm để chỉ đạo giải pháp khắc phục kịp thời.",
         "- Đề kiểm tra bảo mật tuyệt đối, đúng ma trận chuẩn.\n- Tỷ lệ tốt nghiệp THPT đạt mục tiêu nhà trường đề ra.\n- Có học sinh đạt giải HSG cấp tỉnh.",
         "20%",
         "- Ma trận và ngân hàng đề kiểm tra định kỳ.\n- Kế hoạch bồi dưỡng HSG & ôn thi TN THPT.\n- Báo cáo phân tích phổ điểm thi thử tốt nghiệp THPT.\n- Bảng tổng hợp kết quả thi HSG cấp tỉnh."),
        ("3", "Kiểm tra chuyên môn nội bộ, dự giờ & phát triển giáo viên",
         "- Lập kế hoạch kiểm tra chuyên môn nội bộ, kiểm tra hồ sơ giáo án, sổ điểm, ký số của giáo viên.\n- Trực tiếp dự giờ thăm lớp (tối thiểu 1-2 tiết/tuần); chỉ đạo rút kinh nghiệm giờ dạy.\n- Tổ chức Hội thi giáo viên dạy giỏi cấp trường; bồi dưỡng giáo viên dự thi cấp tỉnh.",
         "- Kiểm tra 100% giáo viên trong năm học theo đúng kế hoạch.\n- Nhận xét đánh giá chuyên môn sâu sắc, thúc đẩy giáo viên tiến bộ.",
         "15%",
         "- Kế hoạch kiểm tra nội bộ chuyên môn.\n- Biên bản kiểm tra hồ sơ giáo viên.\n- Phiếu dự giờ, đánh giá tiết dạy.\n- Quyết định công nhận GVDG cấp trường."),
        ("4", "Chỉ đạo chuyển đổi số, thiết bị dạy học & phòng bộ môn",
         "- Giám sát việc quản lý sổ sách điện tử, học bạ điện tử, ký số đúng quy chế.\n- Chỉ đạo khai thác hiệu quả phòng thực hành, thí nghiệm, phòng máy tính, thiết bị dạy học số.\n- Đôn đốc giáo viên xây dựng kho học liệu số và ứng dụng AI trong giảng dạy.",
         "- 100% học bạ và sổ điểm được ký số đúng hạn.\n- Tỷ lệ tiết dạy có sử dụng TBDH đạt trên 90%.",
         "15%",
         "- Báo cáo kết quả chuyển đổi số chuyên môn.\n- Sổ theo dõi sử dụng phòng học bộ môn và TBDH.\n- Kho học liệu số của trường."),
        ("5", "Trực tiếp tham gia giảng dạy bộ môn",
         "- Thực hiện giảng dạy 04 tiết/tuần theo quy định của Bộ GD&ĐT đối với Phó Hiệu trưởng.\n- Thực hiện soạn giảng, kiểm tra đánh giá đúng quy chế chuyên môn.",
         "- Dạy đủ 100% số tiết quy định.\n- Chất lượng giảng dạy môn đạt kết quả tốt.",
         "15%",
         "- Kế hoạch bài dạy (giáo án) cá nhân.\n- Sổ điểm, sổ đánh giá học sinh.\n- Bài kiểm tra đã chấm trả học sinh."),
        ("6", "Phối hợp quản lý nề nếp, HĐTN & điều hành ủy quyền",
         "- Phối hợp quản lý nề nếp học sinh, phong trào Đoàn, hoạt động trải nghiệm - hướng nghiệp.\n- Thay mặt Hiệu trưởng giải quyết công việc của nhà trường khi được ủy quyền.",
         "- Xử lý công việc kịp thời, đúng thẩm quyền và báo cáo đầy đủ cho Hiệu trưởng.",
         "10%",
         "- Kế hoạch hoạt động trải nghiệm - hướng nghiệp.\n- Báo cáo công việc giải quyết theo ủy quyền."),
    ]

    for idx, rdata in enumerate(pht_duties, r_pht_duties+2):
        ws4.cell(row=idx, column=1, value=rdata[0]).alignment = ALIGN_CENTER
        ws4.cell(row=idx, column=2, value=rdata[1]).font = FONT_BOLD
        ws4.cell(row=idx, column=2).alignment = ALIGN_LEFT
        ws4.cell(row=idx, column=3, value=rdata[2]).font = FONT_REGULAR
        ws4.cell(row=idx, column=3).alignment = ALIGN_LEFT
        ws4.cell(row=idx, column=4, value=rdata[3]).font = FONT_REGULAR
        ws4.cell(row=idx, column=4).alignment = ALIGN_LEFT
        ws4.cell(row=idx, column=5, value=rdata[4]).font = FONT_BOLD
        ws4.cell(row=idx, column=5).alignment = ALIGN_CENTER
        ws4.cell(row=idx, column=6, value=rdata[5]).font = FONT_REGULAR
        ws4.cell(row=idx, column=6).alignment = ALIGN_LEFT
        for col_c in range(1, 7): ws4.cell(row=idx, column=col_c).border = BORDER_CELL

    # IV. YÊU CẦU TIÊU CHUẨN PHÓ HIỆU TRƯỞNG
    r_pht_req = r_pht_duties + len(pht_duties) + 3
    ws4.merge_cells(f"A{r_pht_req}:F{r_pht_req}")
    ws4[f"A{r_pht_req}"] = "IV. YÊU CẦU TIÊU CHUẨN, NĂNG LỰC VỊ TRÍ PHÓ HIỆU TRƯỞNG (CHUẨN BẬC 3 - BẬC 5)"
    apply_row_styles(ws4, r_pht_req, font=FONT_SECTION, fill=SECTION_FILL)

    pht_reqs = [
        ("1", "Trình độ chuyên môn đào tạo", "Có bằng Cử nhân trở lên thuộc ngành đào tạo giáo viên hoặc chuyên ngành phù hợp.", "Bằng tốt nghiệp đại học / Thạc sĩ; Bảng điểm"),
        ("2", "Tiêu chuẩn CDNN & Bồi dưỡng", "Được bổ nhiệm chức danh nghề nghiệp Giáo viên THPT Hạng III trở lên; có Chứng chỉ bồi dưỡng CBQL hoặc chứng chỉ CDNN GV THPT.", "Quyết định bổ nhiệm CDNN; Chứng chỉ CBQL"),
        ("3", "Năng lực quản lý chuyên môn (Cấp độ 3-4)", "Năng lực chỉ đạo chuyên môn sâu, tổ chức bồi dưỡng GV, điều hành khảo thí, giải quyết xung đột nghiệp vụ sư phạm.", "Đánh giá Chuẩn nghề nghiệp (Mức Tốt); Quyết định bổ nhiệm PHT"),
        ("4", "Năng lực ứng dụng công nghệ & AI", "Thành thạo công nghệ thông tin, quản lý phần mềm trường học số, thúc đẩy GV ứng dụng công nghệ số và trí tuệ nhân tạo.", "Học bạ điện tử, kho tài liệu số chỉ đạo"),
        ("5", "Phẩm chất chính trị & Đạo đức", "Đảng viên Đảng Cộng sản Việt Nam; phẩm chất chính trị vững vàng, phong cách làm việc gương mẫu, đoàn kết nội bộ.", "Bản kiểm điểm đảng viên; Hồ sơ thi đua khen thưởng"),
    ]
    for idx, rdata in enumerate(pht_reqs, r_pht_req+1):
        ws4.cell(row=idx, column=1, value=rdata[0]).alignment = ALIGN_CENTER
        ws4.cell(row=idx, column=2, value=rdata[1]).font = FONT_BOLD
        ws4.cell(row=idx, column=2).alignment = ALIGN_LEFT
        ws4.merge_cells(f"C{idx}:D{idx}")
        ws4.cell(row=idx, column=3, value=rdata[2]).font = FONT_REGULAR
        ws4.cell(row=idx, column=3).alignment = ALIGN_LEFT
        for c in range(3, 5): ws4.cell(row=idx, column=c).border = BORDER_CELL
        ws4.merge_cells(f"E{idx}:F{idx}")
        ws4.cell(row=idx, column=5, value=rdata[3]).font = FONT_REGULAR
        ws4.cell(row=idx, column=5).alignment = ALIGN_LEFT
        for c in range(5, 7): ws4.cell(row=idx, column=c).border = BORDER_CELL
        for c in range(1, 3): ws4.cell(row=idx, column=c).border = BORDER_CELL

    # Signatures
    r_sig4 = r_pht_req + len(pht_reqs) + 2
    ws4.merge_cells(f"A{r_sig4}:C{r_sig4}")
    ws4[f"A{r_sig4}"] = "PHÓ HIỆU TRƯỞNG"
    ws4[f"A{r_sig4}"].font = FONT_BOLD
    ws4[f"A{r_sig4}"].alignment = ALIGN_CENTER

    ws4.merge_cells(f"D{r_sig4}:F{r_sig4}")
    ws4[f"D{r_sig4}"] = "HIỆU TRƯỞNG PHÊ DUYỆT"
    ws4[f"D{r_sig4}"].font = FONT_BOLD
    ws4[f"D{r_sig4}"].alignment = ALIGN_CENTER

    # ==============================================================================
    # SHEET 5: BẢN MÔ TẢ VTVL TỔ TRƯỞNG CHUYÊN MÔN (TTCM-THPT-01)
    # ==============================================================================
    ws5 = wb.create_sheet(title="5. VTVL Tổ trưởng CM")
    ws5.column_dimensions['A'].width = 8
    ws5.column_dimensions['B'].width = 26
    ws5.column_dimensions['C'].width = 38
    ws5.column_dimensions['D'].width = 24
    ws5.column_dimensions['E'].width = 14
    ws5.column_dimensions['F'].width = 28

    ws5.merge_cells("A1:F1")
    ws5["A1"] = "BẢN MÔ TẢ CÔNG VIỆC VÀ KHUNG NĂNG LỰC VỊ TRÍ VIỆC LÀM QUẢN LÝ"
    ws5["A1"].font = FONT_TITLE_MAIN
    ws5["A1"].alignment = ALIGN_CENTER

    ws5.merge_cells("A2:F2")
    ws5["A2"] = "VỊ TRÍ: TỔ TRƯỞNG CHUYÊN MÔN (MÃ SỐ: TTCM-THPT-01)"
    ws5["A2"].font = Font(name=FONT_FAMILY, size=13, bold=True, color="1E3A8A")
    ws5["A2"].alignment = ALIGN_CENTER

    ws5.merge_cells("A3:F3")
    ws5["A3"] = "Theo Mẫu Phụ lục IV (Phần B) - Nghị định số 232/2026/NĐ-CP ngày 26/6/2026 của Chính phủ"
    ws5["A3"].font = FONT_ITALIC
    ws5["A3"].alignment = ALIGN_CENTER

    # I. THÔNG TIN CHUNG
    ws5.merge_cells("A5:F5")
    ws5["A5"] = "I. THÔNG TIN CHUNG VỀ VỊ TRÍ VIỆC LÀM"
    apply_row_styles(ws5, 5, font=FONT_SECTION, fill=SECTION_FILL)

    ttcm_info = [
        ("1.1", "Tên vị trí việc làm:", "Tổ trưởng chuyên môn", "1.2", "Mã số vị trí:", "TTCM-THPT-01"),
        ("1.3", "Nhóm vị trí việc làm:", "Viên chức quản lý (Mục 7 Phụ lục I NĐ 232/2026)", "1.4", "Cấp quản lý:", "Cấp tổ chuyên môn"),
        ("1.5", "Bậc nghề nghiệp áp dụng:", "Bậc 3 - Bậc 4 (GV THPT Hạng III, Hạng II)", "1.6", "Người quản lý trực tiếp:", "Hiệu trưởng / Phó Hiệu trưởng phụ trách CM"),
        ("1.7", "Đối tượng quản lý trực tiếp:", "Tổ phó chuyên môn, giáo viên và nhân viên trong tổ", "1.8", "Số lượng người đảm nhiệm:", "01 người/tổ (Tổng 04 - 06 người toàn trường)"),
        ("1.9", "Phụ cấp chức vụ lãnh đạo:", "Hệ số 0.25", "1.10", "Định mức giảm tiết giảng dạy:", "Giảm 03 tiết dạy/tuần (dạy 14 tiết/tuần)"),
    ]
    for idx, rdata in enumerate(ttcm_info, 6):
        ws5.cell(row=idx, column=1, value=rdata[0]).alignment = ALIGN_CENTER
        ws5.cell(row=idx, column=2, value=rdata[1]).font = FONT_BOLD
        ws5.cell(row=idx, column=3, value=rdata[2]).font = FONT_REGULAR
        ws5.cell(row=idx, column=4, value=rdata[3]).alignment = ALIGN_CENTER
        ws5.cell(row=idx, column=5, value=rdata[4]).font = FONT_BOLD
        ws5.cell(row=idx, column=6, value=rdata[5]).font = FONT_REGULAR
        for col_c in range(1, 7): ws5.cell(row=idx, column=col_c).border = BORDER_CELL

    # II. MỤC TIÊU VỊ TRÍ
    r_ttcm_obj = 12
    ws5.merge_cells(f"A{r_ttcm_obj}:F{r_ttcm_obj}")
    ws5[f"A{r_ttcm_obj}"] = "II. MỤC TIÊU CỦA VỊ TRÍ VIỆC LÀM"
    apply_row_styles(ws5, r_ttcm_obj, font=FONT_SECTION, fill=SECTION_FILL)

    ws5.merge_cells(f"A{r_ttcm_obj+1}:F{r_ttcm_obj+1}")
    ws5[f"A{r_ttcm_obj+1}"] = "Lãnh đạo, quản lý và điều hành toàn diện công tác chuyên môn của tổ; tổ chức triển khai hiệu quả Chương trình Giáo dục phổ thông 2018; nâng cao chất lượng dạy và học môn học; bồi dưỡng và phát triển năng lực đội ngũ giáo viên; đẩy mạnh đổi mới phương pháp, ứng dụng chuyển đổi số và công nghệ AI trong giảng dạy; chịu trách nhiệm trước Hiệu trưởng về toàn bộ hoạt động của tổ chuyên môn."
    ws5[f"A{r_ttcm_obj+1}"].font = FONT_REGULAR
    ws5[f"A{r_ttcm_obj+1}"].alignment = ALIGN_LEFT
    apply_row_styles(ws5, r_ttcm_obj+1, border=BORDER_CELL)

    # III. CÔNG VIỆC, KẾT QUẢ, SẢN PHẨM QUẢN LÝ
    r_ttcm_duties = 15
    ws5.merge_cells(f"A{r_ttcm_duties}:F{r_ttcm_duties}")
    ws5[f"A{r_ttcm_duties}"] = "III. CÔNG VIỆC, KẾT QUẢ, SẢN PHẨM QUẢN LÝ VÀ CHUYÊN MÔN (ĐÍNH KÈM TỶ TRỌNG & KPI)"
    apply_row_styles(ws5, r_ttcm_duties, font=FONT_SECTION, fill=SECTION_FILL)

    for col_i, h in enumerate(headers_duties, 1):
        cell = ws5.cell(row=r_ttcm_duties+1, column=col_i, value=h)
        cell.font = FONT_HEADER_WHITE
        cell.fill = BLUE_HEADER
        cell.alignment = ALIGN_CENTER
        cell.border = BORDER_HEADER

    ttcm_duties = [
        ("1", "Xây dựng và triển khai kế hoạch giáo dục của tổ",
         "- Xây dựng Kế hoạch dạy học môn học của tổ theo CT GDPT 2018.\n- Phân phối chương trình, xây dựng ma trận đặc tả đề kiểm tra định kỳ.\n- Trình Ban Giám hiệu phê duyệt và tổ chức giám sát thực hiện đúng tiến độ.",
         "- Kế hoạch hoàn thành trước 25/8 hàng năm.\n- 100% giáo viên trong tổ thực hiện đúng kế hoạch.\n- Được Hiệu trưởng phê duyệt.",
         "20%",
         "- Kế hoạch giáo dục tổ chuyên môn năm học.\n- Kế hoạch dạy học từng môn lớp 10, 11, 12.\n- Phân phối chương trình chi tiết."),
        ("2", "Tổ chức sinh hoạt chuyên môn & đổi mới phương pháp",
         "- Sinh hoạt chuyên môn định kỳ ít nhất 2 lần/tháng.\n- Chỉ đạo sinh hoạt chuyên môn theo nghiên cứu bài học (tối thiểu 02 chuyên đề/học kỳ).\n- Triển khai phương pháp dạy học tích cực, giáo dục STEM/STEAM, ứng dụng AI.\n- Đổi mới kiểm tra đánh giá theo định hướng phẩm chất năng lực học sinh.",
         "- Đủ số buổi sinh hoạt (>= 2 lần/tháng).\n- 100% hồ sơ biên bản sinh hoạt được số hóa lưu trữ.\n- 100% GV tham gia đầy đủ, đúng giờ.",
         "20%",
         "- Biên bản sinh hoạt chuyên môn số hóa.\n- Báo cáo chuyên đề nghiên cứu bài học.\n- Kế hoạch bài dạy (giáo án) mẫu của tổ."),
        ("3", "Quản lý, kiểm tra nội bộ và bồi dưỡng giáo viên trong tổ",
         "- Dự giờ, thăm lớp (tối thiểu 2-3 tiết/GV/học kỳ); rút kinh nghiệm giờ dạy.\n- Kiểm tra hồ sơ giáo án, sổ điểm điện tử, đề kiểm tra của giáo viên theo kế hoạch.\n- Tham gia đánh giá chuẩn nghề nghiệp GV và xếp loại thi đua cuối kỳ/năm.\n- Bồi dưỡng GV trẻ, hướng dẫn GV thi GVDG các cấp.",
         "- 100% GV trong tổ được dự giờ và kiểm tra hồ sơ đúng định kỳ.\n- Nhận xét đánh giá công tâm, đúng quy trình, khách quan.",
         "15%",
         "- Phiếu đánh giá dự giờ giáo viên.\n- Biên bản kiểm tra chuyên môn nội bộ tổ.\n- Bảng tổng hợp đánh giá chuẩn nghề nghiệp GV."),
        ("4", "Chỉ đạo bồi dưỡng HSG, ôn thi tốt nghiệp & phụ đạo học sinh",
         "- Xây dựng kế hoạch bồi dưỡng HSG các môn thuộc tổ, phân công GV phụ trách.\n- Xây dựng kế hoạch ôn thi tốt nghiệp THPT khối 12; ra đề thi thử môn học.\n- Tổ chức phụ đạo học sinh có học lực chưa đạt chuẩn, giảm tỷ lệ yếu kém.\n- Hướng dẫn học sinh tham gia nghiên cứu KHKT, STEM cấp trường, cấp tỉnh.",
         "- Đạt và vượt chỉ tiêu giải HSG cấp tỉnh trường giao.\n- Điểm trung bình thi TN THPT môn học đạt mặt bằng chung toàn tỉnh.",
         "15%",
         "- Kế hoạch bồi dưỡng HSG & phụ đạo yếu kém.\n- Danh sách đội tuyển và kết quả thi HSG.\n- Báo cáo kết quả dự thi KHKT cấp tỉnh."),
        ("5", "Trực tiếp tham gia giảng dạy trên lớp",
         "- Giảng dạy môn học theo đúng phân công chuyên môn của nhà trường.\n- Định mức tiết dạy: 14 tiết/tuần (đã giảm 3 tiết nhiệm vụ TTCM theo quy định).\n- Đảm bảo chất lượng bài dạy, soạn bài và chấm trả bài đúng quy chế.",
         "- Hoàn thành 100% số tiết dạy theo định mức.\n- Hồ sơ giáo án đạt chuẩn, đúng quy định.\n- Tỷ lệ HS đạt yêu cầu bộ môn >= 95%.",
         "20%",
         "- Kế hoạch bài dạy (giáo án) cá nhân.\n- Sổ theo dõi đánh giá học sinh.\n- Kết quả kiểm tra, đánh giá học sinh trên lớp."),
        ("6", "Chuyển đổi số & ứng dụng AI trong hoạt động của tổ",
         "- Triển khai sử dụng sổ điểm điện tử, học bạ số, ký số trên phần mềm quản lý.\n- Hướng dẫn GV tổ ứng dụng CNTT, phần mềm mô phỏng và các công cụ AI hỗ trợ soạn bài.\n- Xây dựng kho học liệu số, ngân hàng câu hỏi trắc nghiệm của tổ.",
         "- 100% GV thực hiện sổ sách điện tử, ký số đúng hạn.\n- Đóng góp tối thiểu 20 học liệu số/học kỳ vào kho dùng chung của trường.",
         "10%",
         "- Kho học liệu số của tổ chuyên môn.\n- Ngân hàng đề thi/kiểm tra số hóa của tổ.\n- Báo cáo ứng dụng CNTT và AI của tổ."),
    ]

    for idx, rdata in enumerate(ttcm_duties, r_ttcm_duties+2):
        ws5.cell(row=idx, column=1, value=rdata[0]).alignment = ALIGN_CENTER
        ws5.cell(row=idx, column=2, value=rdata[1]).font = FONT_BOLD
        ws5.cell(row=idx, column=2).alignment = ALIGN_LEFT
        ws5.cell(row=idx, column=3, value=rdata[2]).font = FONT_REGULAR
        ws5.cell(row=idx, column=3).alignment = ALIGN_LEFT
        ws5.cell(row=idx, column=4, value=rdata[3]).font = FONT_REGULAR
        ws5.cell(row=idx, column=4).alignment = ALIGN_LEFT
        ws5.cell(row=idx, column=5, value=rdata[4]).font = FONT_BOLD
        ws5.cell(row=idx, column=5).alignment = ALIGN_CENTER
        ws5.cell(row=idx, column=6, value=rdata[5]).font = FONT_REGULAR
        ws5.cell(row=idx, column=6).alignment = ALIGN_LEFT
        for col_c in range(1, 7): ws5.cell(row=idx, column=col_c).border = BORDER_CELL

    # IV. TIÊU CHUẨN NĂNG LỰC TỔ TRƯỞNG CM
    r_ttcm_req = r_ttcm_duties + len(ttcm_duties) + 3
    ws5.merge_cells(f"A{r_ttcm_req}:F{r_ttcm_req}")
    ws5[f"A{r_ttcm_req}"] = "IV. YÊU CẦU TIÊU CHUẨN, NĂNG LỰC VỊ TRÍ TỔ TRƯỞNG CHUYÊN MÔN (CHUẨN BẬC 3 - BẬC 4)"
    apply_row_styles(ws5, r_ttcm_req, font=FONT_SECTION, fill=SECTION_FILL)

    ttcm_reqs = [
        ("1", "Trình độ chuyên môn đào tạo", "Có bằng Cử nhân trở lên ngành sư phạm (hoặc cử nhân chuyên ngành kèm chứng chỉ NVSP).", "Bằng tốt nghiệp đại học; Bảng điểm"),
        ("2", "Tiêu chuẩn CDNN & Ngạch bậc", "Được bổ nhiệm CDNN Giáo viên THPT Hạng II (hoặc Hạng III có thành tích xuất sắc); có chứng chỉ bồi dưỡng CDNN.", "Quyết định bổ nhiệm CDNN Hạng II/III; Chứng chỉ CDNN"),
        ("3", "Năng lực quản lý chuyên môn tổ", "Có uy tín chuyên môn cao trong trường; năng lực điều hành sinh hoạt tổ, hướng dẫn đồng nghiệp, xử lý tình huống sư phạm.", "Đánh giá Chuẩn nghề nghiệp GV (Mức Tốt); Quyết định bổ nhiệm TTCM"),
        ("4", "Năng lực tin học & Chuyển đổi số", "Thành thạo phần mềm quản trị trường học, học bạ điện tử, ký số; ứng dụng thành thạo AI và CNTT trong bài dạy.", "Hồ sơ bài giảng số hóa, chứng nhận tập huấn CNTT/AI"),
        ("5", "Phẩm chất đạo đức & Lối sống", "Có phẩm chất chính trị vững vàng, lối sống trung thực, tận tụy, mẫu mực, đoàn kết nội bộ.", "Bản kiểm điểm Đảng viên/viên chức; Phiếu đánh giá thi đua"),
    ]
    for idx, rdata in enumerate(ttcm_reqs, r_ttcm_req+1):
        ws5.cell(row=idx, column=1, value=rdata[0]).alignment = ALIGN_CENTER
        ws5.cell(row=idx, column=2, value=rdata[1]).font = FONT_BOLD
        ws5.cell(row=idx, column=2).alignment = ALIGN_LEFT
        ws5.merge_cells(f"C{idx}:D{idx}")
        ws5.cell(row=idx, column=3, value=rdata[2]).font = FONT_REGULAR
        ws5.cell(row=idx, column=3).alignment = ALIGN_LEFT
        for c in range(3, 5): ws5.cell(row=idx, column=c).border = BORDER_CELL
        ws5.merge_cells(f"E{idx}:F{idx}")
        ws5.cell(row=idx, column=5, value=rdata[3]).font = FONT_REGULAR
        ws5.cell(row=idx, column=5).alignment = ALIGN_LEFT
        for c in range(5, 7): ws5.cell(row=idx, column=c).border = BORDER_CELL
        for c in range(1, 3): ws5.cell(row=idx, column=c).border = BORDER_CELL

    # Signatures
    r_sig5 = r_ttcm_req + len(ttcm_reqs) + 2
    ws5.merge_cells(f"A{r_sig5}:C{r_sig5}")
    ws5[f"A{r_sig5}"] = "TỔ TRƯỞNG CHUYÊN MÔN"
    ws5[f"A{r_sig5}"].font = FONT_BOLD
    ws5[f"A{r_sig5}"].alignment = ALIGN_CENTER

    ws5.merge_cells(f"D{r_sig5}:F{r_sig5}")
    ws5[f"D{r_sig5}"] = "HIỆU TRƯỞNG PHÊ DUYỆT"
    ws5[f"D{r_sig5}"].font = FONT_BOLD
    ws5[f"D{r_sig5}"].alignment = ALIGN_CENTER

    # ==============================================================================
    # SHEET 6: BẢN MÔ TẢ VTVL TỔ PHÓ CHUYÊN MÔN (TPCM-THPT-01)
    # ==============================================================================
    ws6 = wb.create_sheet(title="6. VTVL Tổ phó CM")
    ws6.column_dimensions['A'].width = 8
    ws6.column_dimensions['B'].width = 26
    ws6.column_dimensions['C'].width = 38
    ws6.column_dimensions['D'].width = 24
    ws6.column_dimensions['E'].width = 14
    ws6.column_dimensions['F'].width = 28

    ws6.merge_cells("A1:F1")
    ws6["A1"] = "BẢN MÔ TẢ CÔNG VIỆC VÀ KHUNG NĂNG LỰC VỊ TRÍ VIỆC LÀM QUẢN LÝ"
    ws6["A1"].font = FONT_TITLE_MAIN
    ws6["A1"].alignment = ALIGN_CENTER

    ws6.merge_cells("A2:F2")
    ws6["A2"] = "VỊ TRÍ: TỔ PHÓ CHUYÊN MÔN (MÃ SỐ: TPCM-THPT-01)"
    ws6["A2"].font = Font(name=FONT_FAMILY, size=13, bold=True, color="1E3A8A")
    ws6["A2"].alignment = ALIGN_CENTER

    ws6.merge_cells("A3:F3")
    ws6["A3"] = "Theo Mẫu Phụ lục IV (Phần B) - Nghị định số 232/2026/NĐ-CP ngày 26/6/2026 của Chính phủ"
    ws6["A3"].font = FONT_ITALIC
    ws6["A3"].alignment = ALIGN_CENTER

    # I. THÔNG TIN CHUNG
    ws6.merge_cells("A5:F5")
    ws6["A5"] = "I. THÔNG TIN CHUNG VỀ VỊ TRÍ VIỆC LÀM"
    apply_row_styles(ws6, 5, font=FONT_SECTION, fill=SECTION_FILL)

    tpcm_info = [
        ("1.1", "Tên vị trí việc làm:", "Tổ phó chuyên môn", "1.2", "Mã số vị trí:", "TPCM-THPT-01"),
        ("1.3", "Nhóm vị trí việc làm:", "Viên chức quản lý (Mục 8 Phụ lục I NĐ 232/2026)", "1.4", "Cấp quản lý:", "Cấp tổ chuyên môn"),
        ("1.5", "Bậc nghề nghiệp áp dụng:", "Bậc 3 - Bậc 4 (GV THPT Hạng III, Hạng II)", "1.6", "Người quản lý trực tiếp:", "Tổ trưởng chuyên môn"),
        ("1.7", "Đối tượng phối hợp quản lý:", "Giáo viên và nhân viên trong tổ chuyên môn", "1.8", "Số lượng người đảm nhiệm:", "01 người/tổ (Tổng 04 - 06 người toàn trường)"),
        ("1.9", "Phụ cấp chức vụ lãnh đạo:", "Hệ số 0.15", "1.10", "Định mức giảm tiết giảng dạy:", "Giảm 01 tiết dạy/tuần (dạy 16 tiết/tuần)"),
    ]
    for idx, rdata in enumerate(tpcm_info, 6):
        ws6.cell(row=idx, column=1, value=rdata[0]).alignment = ALIGN_CENTER
        ws6.cell(row=idx, column=2, value=rdata[1]).font = FONT_BOLD
        ws6.cell(row=idx, column=3, value=rdata[2]).font = FONT_REGULAR
        ws6.cell(row=idx, column=4, value=rdata[3]).alignment = ALIGN_CENTER
        ws6.cell(row=idx, column=5, value=rdata[4]).font = FONT_BOLD
        ws6.cell(row=idx, column=6, value=rdata[5]).font = FONT_REGULAR
        for col_c in range(1, 7): ws6.cell(row=idx, column=col_c).border = BORDER_CELL

    # II. MỤC TIÊU VỊ TRÍ
    r_tpcm_obj = 12
    ws6.merge_cells(f"A{r_tpcm_obj}:F{r_tpcm_obj}")
    ws6[f"A{r_tpcm_obj}"] = "II. MỤC TIÊU CỦA VỊ TRÍ VIỆC LÀM"
    apply_row_styles(ws6, r_tpcm_obj, font=FONT_SECTION, fill=SECTION_FILL)

    ws6.merge_cells(f"A{r_tpcm_obj+1}:F{r_tpcm_obj+1}")
    ws6[f"A{r_tpcm_obj+1}"] = "Giúp việc cho Tổ trưởng chuyên môn trong việc tổ chức, đôn đốc và giám sát các hoạt động chuyên môn của tổ; trực tiếp phụ trách các mảng công việc chuyên môn được phân công (theo dõi tiến độ chương trình, nề nếp hồ sơ sổ sách, chuyển đổi số, thiết bị dạy học); thay mặt điều hành hoạt động của tổ khi Tổ trưởng vắng mặt hoặc được ủy quyền."
    ws6[f"A{r_tpcm_obj+1}"].font = FONT_REGULAR
    ws6[f"A{r_tpcm_obj+1}"].alignment = ALIGN_LEFT
    apply_row_styles(ws6, r_tpcm_obj+1, border=BORDER_CELL)

    # III. CÔNG VIỆC, KẾT QUẢ, SẢN PHẨM QUẢN LÝ
    r_tpcm_duties = 15
    ws6.merge_cells(f"A{r_tpcm_duties}:F{r_tpcm_duties}")
    ws6[f"A{r_tpcm_duties}"] = "III. CÔNG VIỆC, KẾT QUẢ, SẢN PHẨM QUẢN LÝ VÀ CHUYÊN MÔN (ĐÍNH KÈM TỶ TRỌNG & KPI)"
    apply_row_styles(ws6, r_tpcm_duties, font=FONT_SECTION, fill=SECTION_FILL)

    for col_i, h in enumerate(headers_duties, 1):
        cell = ws6.cell(row=r_tpcm_duties+1, column=col_i, value=h)
        cell.font = FONT_HEADER_WHITE
        cell.fill = BLUE_HEADER
        cell.alignment = ALIGN_CENTER
        cell.border = BORDER_HEADER

    tpcm_duties = [
        ("1", "Hỗ trợ xây dựng kế hoạch và theo dõi tiến độ thực hiện",
         "- Tham gia xây dựng Kế hoạch dạy học của tổ và kế hoạch giáo dục môn học.\n- Theo dõi tiến độ dạy học, lịch báo giảng của GV hàng tuần, đôn đốc dạy bù/dạy thay.\n- Báo cáo định kỳ với Tổ trưởng về tình hình thực hiện chương trình GDPT 2018.",
         "- Kế hoạch tổ hoàn thành đúng hạn.\n- 100% GV thực hiện đúng tiến độ, không để chậm trễ chương trình.",
         "20%",
         "- Dự thảo kế hoạch dạy học môn học.\n- Sổ theo dõi tiến độ chương trình của tổ.\n- Báo cáo tiến độ chương trình hàng tháng."),
        ("2", "Phụ trách hồ sơ chuyên môn, sổ sách điện tử & ký số",
         "- Kiểm tra việc cập nhật sổ điểm điện tử, học bạ điện tử, ký số của GV trong tổ.\n- Lập biên bản các buổi sinh hoạt chuyên môn, lưu trữ hồ sơ số hóa trên hệ thống.\n- Quản lý kho đề kiểm tra định kỳ, bài kiểm tra của các bộ môn trong tổ.",
         "- 100% biên bản sinh hoạt được lập đầy đủ, ký duyệt và số hóa.\n- Nhắc nhở và đôn đốc kịp thời GV hoàn thành ký số sổ điểm đúng hạn.",
         "20%",
         "- Tập biên bản họp tổ chuyên môn số hóa.\n- Bảng theo dõi ký số học bạ, sổ điểm của tổ.\n- Hồ sơ lưu trữ đề kiểm tra định kỳ."),
        ("3", "Tham gia dự giờ, kiểm tra chuyên môn & hỗ trợ đồng nghiệp",
         "- Dự giờ giáo viên trong tổ (tối thiểu 1-2 tiết/tháng); tham gia góp ý giờ dạy.\n- Hỗ trợ giáo viên trẻ, giáo viên mới tiếp cận chương trình và phương pháp dạy học mới.\n- Tham gia kiểm tra hồ sơ giáo án, kế hoạch bài dạy của GV theo phân công của TTCM.",
         "- Dự giờ đủ số tiết theo quy định.\n- Nhận xét mang tính xây dựng, hỗ trợ đồng nghiệp nâng cao chất lượng bài giảng.",
         "15%",
         "- Phiếu dự giờ đồng nghiệp.\n- Biên bản kiểm tra kế hoạch bài dạy.\n- Báo cáo kết quả hỗ trợ giáo viên trẻ."),
        ("4", "Phụ trách thiết bị dạy học, phòng thí nghiệm & hoạt động STEM",
         "- Theo dõi tình hình sử dụng thiết bị dạy học, phòng thực hành, hóa chất thí nghiệm của tổ.\n- Đề xuất bổ sung, sửa chữa thiết bị dạy học hỏng hóc hoặc thiếu thốn.\n- Phối hợp tổ chức hoạt động trải nghiệm, giáo dục STEM/STEAM, câu lạc bộ bộ môn.",
         "- Thiết bị dạy học được bảo quản tốt, khai thác đúng công năng.\n- Tổ chức ít nhất 01 hoạt động trải nghiệm/STEM môn học trong năm.",
         "15%",
         "- Sổ theo dõi mượn trả TBDH và phòng bộ môn.\n- Phiếu đề xuất mua sắm/tu sửa thiết bị.\n- Kế hoạch và báo cáo hoạt động STEM môn học."),
        ("5", "Trực tiếp tham gia giảng dạy trên lớp",
         "- Giảng dạy môn học theo đúng phân công chuyên môn của nhà trường.\n- Định mức tiết dạy: 16 tiết/tuần (đã giảm 01 tiết nhiệm vụ TPCM theo quy định).\n- Soạn bài, kiểm tra đánh giá đúng quy chế chuyên môn.",
         "- Dạy đủ 100% số tiết quy định.\n- Hồ sơ giáo án đạt chuẩn, chất lượng giảng dạy tốt.",
         "20%",
         "- Kế hoạch bài dạy (giáo án) cá nhân.\n- Sổ theo dõi đánh giá học sinh.\n- Kết quả kiểm tra, đánh giá học sinh trên lớp."),
        ("6", "Điều hành tổ khi được ủy quyền & phối hợp công tác đoàn thể",
         "- Thay mặt Tổ trưởng điều hành các cuộc họp tổ, xử lý công việc đột xuất khi TTCM vắng mặt.\n- Phối hợp với Công đoàn, Đoàn thanh niên tổ chức các phong trào thi đua trong tổ.",
         "- Hoàn thành tốt công việc điều hành khi được ủy quyền, không để ách tắc công việc.",
         "10%",
         "- Biên bản họp tổ do TPCM chủ trì.\n- Báo cáo công việc giải quyết theo ủy quyền."),
    ]

    for idx, rdata in enumerate(tpcm_duties, r_tpcm_duties+2):
        ws6.cell(row=idx, column=1, value=rdata[0]).alignment = ALIGN_CENTER
        ws6.cell(row=idx, column=2, value=rdata[1]).font = FONT_BOLD
        ws6.cell(row=idx, column=2).alignment = ALIGN_LEFT
        ws6.cell(row=idx, column=3, value=rdata[2]).font = FONT_REGULAR
        ws6.cell(row=idx, column=3).alignment = ALIGN_LEFT
        ws6.cell(row=idx, column=4, value=rdata[3]).font = FONT_REGULAR
        ws6.cell(row=idx, column=4).alignment = ALIGN_LEFT
        ws6.cell(row=idx, column=5, value=rdata[4]).font = FONT_BOLD
        ws6.cell(row=idx, column=5).alignment = ALIGN_CENTER
        ws6.cell(row=idx, column=6, value=rdata[5]).font = FONT_REGULAR
        ws6.cell(row=idx, column=6).alignment = ALIGN_LEFT
        for col_c in range(1, 7): ws6.cell(row=idx, column=col_c).border = BORDER_CELL

    # IV. TIÊU CHUẨN NĂNG LỰC TỔ PHÓ CM
    r_tpcm_req = r_tpcm_duties + len(tpcm_duties) + 3
    ws6.merge_cells(f"A{r_tpcm_req}:F{r_tpcm_req}")
    ws6[f"A{r_tpcm_req}"] = "IV. YÊU CẦU TIÊU CHUẨN, NĂNG LỰC VỊ TRÍ TỔ PHÓ CHUYÊN MÔN (CHUẨN BẬC 3 - BẬC 4)"
    apply_row_styles(ws6, r_tpcm_req, font=FONT_SECTION, fill=SECTION_FILL)

    tpcm_reqs = [
        ("1", "Trình độ chuyên môn đào tạo", "Có bằng Cử nhân trở lên ngành sư phạm (hoặc cử nhân chuyên ngành kèm chứng chỉ NVSP).", "Bằng tốt nghiệp đại học; Bảng điểm"),
        ("2", "Tiêu chuẩn CDNN & Ngạch bậc", "Được bổ nhiệm CDNN Giáo viên THPT Hạng III trở lên; có chứng chỉ bồi dưỡng CDNN.", "Quyết định bổ nhiệm CDNN; Chứng chỉ CDNN"),
        ("3", "Năng lực chuyên môn & Điều phối", "Vững vàng về chuyên môn môn học; có kỹ năng tổng hợp thông tin, lập biên bản, theo dõi tiến độ và giao tiếp tốt.", "Đánh giá Chuẩn nghề nghiệp GV (Mức Khá/Tốt); Quyết định bổ nhiệm"),
        ("4", "Năng lực tin học & Chuyển đổi số", "Thành thạo kỹ năng tin học văn phòng (Word, Excel), sử dụng phần mềm quản trị trường học số, hỗ trợ số hóa hồ sơ tổ.", "Sản phẩm hồ sơ số hóa, chứng nhận tập huấn CNTT"),
        ("5", "Phẩm chất đạo đức & Trách nhiệm", "Có tinh thần trách nhiệm cao, trung thực, nhiệt tình, phối hợp chặt chẽ với Tổ trưởng và đồng nghiệp.", "Phiếu đánh giá viên chức hàng năm"),
    ]
    for idx, rdata in enumerate(tpcm_reqs, r_tpcm_req+1):
        ws6.cell(row=idx, column=1, value=rdata[0]).alignment = ALIGN_CENTER
        ws6.cell(row=idx, column=2, value=rdata[1]).font = FONT_BOLD
        ws6.cell(row=idx, column=2).alignment = ALIGN_LEFT
        ws6.merge_cells(f"C{idx}:D{idx}")
        ws6.cell(row=idx, column=3, value=rdata[2]).font = FONT_REGULAR
        ws6.cell(row=idx, column=3).alignment = ALIGN_LEFT
        for c in range(3, 5): ws6.cell(row=idx, column=c).border = BORDER_CELL
        ws6.merge_cells(f"E{idx}:F{idx}")
        ws6.cell(row=idx, column=5, value=rdata[3]).font = FONT_REGULAR
        ws6.cell(row=idx, column=5).alignment = ALIGN_LEFT
        for c in range(5, 7): ws6.cell(row=idx, column=c).border = BORDER_CELL
        for c in range(1, 3): ws6.cell(row=idx, column=c).border = BORDER_CELL

    # Signatures
    r_sig6 = r_tpcm_req + len(tpcm_reqs) + 2
    ws6.merge_cells(f"A{r_sig6}:C{r_sig6}")
    ws6[f"A{r_sig6}"] = "TỔ PHÓ CHUYÊN MÔN"
    ws6[f"A{r_sig6}"].font = FONT_BOLD
    ws6[f"A{r_sig6}"].alignment = ALIGN_CENTER

    ws6.merge_cells(f"D{r_sig6}:F{r_sig6}")
    ws6[f"D{r_sig6}"] = "HIỆU TRƯỞNG PHÊ DUYỆT"
    ws6[f"D{r_sig6}"].font = FONT_BOLD
    ws6[f"D{r_sig6}"].alignment = ALIGN_CENTER

    # ==============================================================================
    # SHEET 7: KHUNG NĂNG LỰC & BỘ TIÊU CHÍ KPI ĐO LƯỜNG
    # ==============================================================================
    ws7 = wb.create_sheet(title="7. Khung năng lực & KPI")
    ws7.column_dimensions['A'].width = 8
    ws7.column_dimensions['B'].width = 22
    ws7.column_dimensions['C'].width = 32
    ws7.column_dimensions['D'].width = 14
    ws7.column_dimensions['E'].width = 14
    ws7.column_dimensions['F'].width = 14
    ws7.column_dimensions['G'].width = 14
    ws7.column_dimensions['H'].width = 26

    ws7.merge_cells("A1:H1")
    ws7["A1"] = "BẢNG KHUNG NĂNG LỰC VÀ MA TRẬN KPI VIÊN CHỨC QUẢN LÝ"
    ws7["A1"].font = FONT_TITLE_MAIN
    ws7["A1"].alignment = ALIGN_CENTER

    ws7.merge_cells("A2:H2")
    ws7["A2"] = "Áp dụng đánh giá xếp loại Hiệu trưởng, Phó Hiệu trưởng, Tổ trưởng CM, Tổ phó CM theo NĐ 232/2026/NĐ-CP"
    ws7["A2"].font = Font(name=FONT_FAMILY, size=11, italic=True)
    ws7["A2"].alignment = ALIGN_CENTER

    # BẢNG 1: KHUNG NĂNG LỰC CHUẨN (3 NHÓM)
    ws7.merge_cells("A4:H4")
    ws7["A4"] = "PHẦN I. MA TRẬN KHUNG NĂNG LỰC (CẤP ĐỘ TỪ 1 ĐẾN 5 THEO NGHỊ ĐỊNH 232/2026)"
    apply_row_styles(ws7, 4, font=FONT_SECTION, fill=SECTION_FILL)

    headers_knl = [
        "STT", "Nhóm năng lực", "Tên năng lực cụ thể", "Hiệu trưởng (HT)", "Phó HT (PHT)",
        "Tổ trưởng (TTCM)", "Tổ phó (TPCM)", "Mô tả chuẩn hành vi yêu cầu"
    ]
    for col_i, h in enumerate(headers_knl, 1):
        cell = ws7.cell(row=5, column=col_i, value=h)
        cell.font = FONT_HEADER_WHITE
        cell.fill = PURPLE_HEADER
        cell.alignment = ALIGN_CENTER
        cell.border = BORDER_HEADER

    knl_data = [
        # Năng lực chung
        ("1", "Năng lực chung", "Phẩm chất chính trị & Đạo đức công vụ", "Cấp độ 5", "Cấp độ 4-5", "Cấp độ 4", "Cấp độ 3-4", "Tuyệt đối trung thành, gương mẫu, liêm chính, tôn trọng pháp luật và quy chế."),
        ("2", "Năng lực chung", "Giao tiếp, ứng xử & Phục vụ cộng đồng", "Cấp độ 5", "Cấp độ 4-5", "Cấp độ 4", "Cấp độ 3-4", "Giao tiếp sư phạm chuẩn mực, lắng nghe phụ huynh, giải quyết thỏa đáng vướng mắc."),
        ("3", "Năng lực chung", "Đổi mới, sáng tạo & Thích ứng", "Cấp độ 5", "Cấp độ 4-5", "Cấp độ 4", "Cấp độ 3-4", "Chủ động đề xuất phương pháp mới, cải tiến mô hình giáo dục, thích ứng thay đổi."),
        ("4", "Năng lực chung", "Ứng dụng CNTT, Kỹ năng số & AI", "Cấp độ 4-5", "Cấp độ 4-5", "Cấp độ 4", "Cấp độ 3-4", "Khai thác thành thạo hệ thống quản trị số, học bạ điện tử, công cụ AI dạy học."),
        # Năng lực quản lý
        ("5", "Năng lực quản lý", "Tư duy chiến lược & Lập kế hoạch", "Cấp độ 5", "Cấp độ 4-5", "Cấp độ 4", "Cấp độ 3", "Xây dựng chiến lược trường học, kế hoạch chuyên môn khả thi và tầm nhìn dài hạn."),
        ("6", "Năng lực quản lý", "Tổ chức điều hành & Giám sát", "Cấp độ 5", "Cấp độ 4-5", "Cấp độ 4", "Cấp độ 3-4", "Phân công việc khoa học, kiểm tra đôn đốc tiến độ, đo lường kết quả công bằng."),
        ("7", "Năng lực quản lý", "Quản trị nhân sự & Phát triển đội ngũ", "Cấp độ 5", "Cấp độ 4", "Cấp độ 4", "Cấp độ 3", "Động viên khích lệ, bồi dưỡng GV trẻ, xây dựng khối đoàn kết nội bộ vững chắc."),
        ("8", "Năng lực quản lý", "Quản trị rủi ro & Xử lý tình huống", "Cấp độ 5", "Cấp độ 4", "Cấp độ 3-4", "Cấp độ 3", "Nhận diện nguy cơ mất an toàn trường học, giải quyết xung đột thấu tình đạt lý."),
        # Năng lực chuyên môn
        ("9", "Năng lực chuyên môn", "Xây dựng & phát triển chương trình", "Cấp độ 4-5", "Cấp độ 5", "Cấp độ 4-5", "Cấp độ 3-4", "Nắm vững CT GDPT 2018, chỉ đạo phân phối chương trình và tích hợp liên môn."),
        ("10", "Năng lực chuyên môn", "Phương pháp dạy học tích cực & STEM", "Cấp độ 4", "Cấp độ 4-5", "Cấp độ 4-5", "Cấp độ 4", "Vận dụng và hướng dẫn đồng nghiệp các kỹ thuật dạy học hiện đại, bài học STEM."),
        ("11", "Năng lực chuyên môn", "Khảo thí & Kiểm tra đánh giá phẩm chất", "Cấp độ 4-5", "Cấp độ 5", "Cấp độ 4-5", "Cấp độ 4", "Xây dựng ma trận đặc tả, ngân hàng đề chuẩn, phân tích phổ điểm khoa học."),
    ]

    for r_i, kdata in enumerate(knl_data, 6):
        for c_i, val in enumerate(kdata, 1):
            cell = ws7.cell(row=r_i, column=c_i, value=val)
            cell.font = FONT_BOLD if c_i in [2, 3] else FONT_REGULAR
            cell.border = BORDER_CELL
            cell.alignment = ALIGN_CENTER if c_i in [1, 4, 5, 6, 7] else ALIGN_LEFT

    # BẢNG 2: BỘ TIÊU CHÍ KPI ĐÁNH GIÁ KẾT QUẢ ĐẦU RA HÀNG THÁNG/KỲ/NĂM
    r_kpi_start = len(knl_data) + 8
    ws7.merge_cells(f"A{r_kpi_start}:H{r_kpi_start}")
    ws7[f"A{r_kpi_start}"] = "PHẦN II. BỘ TIÊU CHÍ KPI ĐO LƯỜNG HIỆU QUẢ CÔNG VIỆC THEO KẾT QUẢ ĐẦU RA"
    apply_row_styles(ws7, r_kpi_start, font=FONT_SECTION, fill=SECTION_FILL)

    headers_kpi = [
        "STT", "Vị trí áp dụng", "Chỉ số KPI chính", "Mục tiêu định lượng", "Tần suất đánh giá",
        "Phương pháp thu thập", "Trọng số điểm", "Mức độ hoàn thành xuất sắc"
    ]
    for col_i, h in enumerate(headers_kpi, 1):
        cell = ws7.cell(row=r_kpi_start+1, column=col_i, value=h)
        cell.font = FONT_HEADER_WHITE
        cell.fill = AMBER_HEADER
        cell.alignment = ALIGN_CENTER
        cell.border = BORDER_HEADER

    kpi_items = [
        ("1", "Hiệu trưởng", "Tỷ lệ tốt nghiệp THPT toàn trường", ">= 98.5%", "Hàng năm", "Kết quả thi chính thức của Bộ GD&ĐT", "25 điểm", ">= 99.5% và điểm TB top đầu toàn tỉnh"),
        ("2", "Hiệu trưởng", "Quyết toán ngân sách & Quản trị tài chính", "100% đúng hạn, không sai phạm", "Hàng quý/năm", "Báo cáo duyệt quyết toán của Sở GD&ĐT", "20 điểm", "Tiết kiệm chi, minh bạch 100%, kiểm toán tốt"),
        ("3", "Hiệu trưởng", "Chất lượng kiểm định & Trường học an toàn", "Mức 2 trở lên; 0 vụ việc mất an toàn", "Hàng năm", "Biên bản kiểm tra KĐCLGD & Công an", "20 điểm", "Đạt chuẩn quốc gia Mức 2, an toàn xuất sắc"),
        ("4", "Phó Hiệu trưởng", "Tiến độ chương trình & Ký số học bạ", "100% đúng hạn, không cắt xén", "Hàng tháng/kỳ", "Hệ thống quản lý điểm số điện tử", "25 điểm", "Hoàn thành sớm hơn quy định 2 ngày"),
        ("5", "Phó Hiệu trưởng", "Kết quả thi HSG & Phổ điểm thi TN THPT", "Đạt chỉ tiêu tỉnh giao", "Hàng năm", "Quyết định công nhận giải thưởng Sở GD&ĐT", "25 điểm", "Vượt chỉ tiêu số lượng giải tỉnh >= 15%"),
        ("6", "Phó Hiệu trưởng", "Kiểm tra chuyên môn & Dự giờ GV", "100% GV theo kế hoạch", "Hàng kỳ", "Biên bản kiểm tra & Phiếu dự giờ", "20 điểm", "Kiểm tra chuyên môn sâu, có chuyên đề đột phá"),
        ("7", "Tổ trưởng CM", "Tỷ lệ sinh hoạt chuyên môn theo NCBL", ">= 02 chuyên đề/học kỳ", "Hàng tháng", "Biên bản số hóa trên hệ thống trường", "25 điểm", "Có 03 chuyên đề chất lượng cao, chia sẻ cấp trường"),
        ("8", "Tổ trưởng CM", "Chỉ đạo chất lượng học tập bộ môn", "Học sinh đạt chuẩn >= 95%", "Hàng kỳ/năm", "Thống kê điểm trung bình môn trên hệ thống", "25 điểm", "Học sinh giỏi bộ môn tăng >= 10%"),
        ("9", "Tổ phó CM", "Kiểm tra nề nếp sổ sách & Ký số của tổ", "100% thành viên đúng hạn", "Hàng tháng", "Báo cáo kiểm tra định kỳ của tổ", "25 điểm", "Không có bất kỳ trường hợp nào nộp trễ"),
        ("10", "Tổ phó CM", "Khai thác thiết bị dạy học & phòng bộ môn", ">= 90% tiết có TBDH", "Hàng tháng", "Sổ theo dõi mượn trả thiết bị & phần mềm", "25 điểm", "Đạt 95% trở lên, bảo quản tốt không thất thoát"),
    ]

    for r_i, kpdata in enumerate(kpi_items, r_kpi_start+2):
        for c_i, val in enumerate(kpdata, 1):
            cell = ws7.cell(row=r_i, column=c_i, value=val)
            cell.font = FONT_BOLD if c_i in [2, 3] else FONT_REGULAR
            cell.border = BORDER_CELL
            cell.alignment = ALIGN_CENTER if c_i in [1, 4, 5, 7] else ALIGN_LEFT

    # Signatures
    r_sig7 = r_kpi_start + len(kpi_items) + 3
    ws7.merge_cells(f"A{r_sig7}:D{r_sig7}")
    ws7[f"A{r_sig7}"] = "HỘI ĐỒNG ĐÁNH GIÁ XẾP LOẠI NHÀ TRƯỜNG"
    ws7[f"A{r_sig7}"].font = FONT_BOLD
    ws7[f"A{r_sig7}"].alignment = ALIGN_CENTER

    ws7.merge_cells(f"E{r_sig7}:H{r_sig7}")
    ws7[f"E{r_sig7}"] = "HIỆU TRƯỞNG PHÊ DUYỆT"
    ws7[f"E{r_sig7}"].font = FONT_BOLD
    ws7[f"E{r_sig7}"].alignment = ALIGN_CENTER

    # Auto-fit all sheets
    for ws_item in [ws1, ws2, ws3, ws4, ws5, ws6, ws7]:
        auto_fit_columns(ws_item)

    # Save to all target paths
    for p in output_paths:
        os.makedirs(os.path.dirname(os.path.abspath(p)), exist_ok=True)
        wb.save(p)
        print(f"✓ Đã tạo thành công file Excel: {p}")

if __name__ == "__main__":
    paths = [
        r"d:\Du-an-web\web-vtvl-ttcm-tpcm-phuc-hoa\Bang-mo-ta-De-an-VTVL-THPT-Phuc-Hoa-ND232-2026.xlsx",
        r"d:\Du-an-web\web-vtvl-ttcm-tpcm-phuc-hoa\Bang-mo-ta-VTVL-TTCM-TPCM-THPT-Phuc-Hoa.xlsx",
        r"D:\Desktop\Bang-mo-ta-De-an-VTVL-THPT-Phuc-Hoa-ND232-2026.xlsx",
        r"D:\Desktop\Bang-mo-ta-VTVL-TTCM-TPCM-THPT-Phuc-Hoa.xlsx",
    ]
    create_vtvl_excel(paths)
