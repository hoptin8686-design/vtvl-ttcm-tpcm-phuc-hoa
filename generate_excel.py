# -*- coding: utf-8 -*-
"""
Tạo bảng Excel mô tả vị trí việc làm TTCM & TPCM - THPT Phục Hòa
Căn cứ Nghị định số 232/2026/NĐ-CP
"""

import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

wb = openpyxl.Workbook()
# Remove default sheet
wb.remove(wb.active)

# Color definitions
NAVY_FILL = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid") # #1E3A8A
BLUE_HEADER = PatternFill(start_color="2563EB", end_color="2563EB", fill_type="solid") # #2563EB
TEAL_HEADER = PatternFill(start_color="0D9488", end_color="0D9488", fill_type="solid") # #0D9488
LIGHT_BLUE_FILL = PatternFill(start_color="EFF6FF", end_color="EFF6FF", fill_type="solid") # #EFF6FF
LIGHT_GRAY_FILL = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid") # #F8FAFC
SECTION_FILL = PatternFill(start_color="DBEAFE", end_color="DBEAFE", fill_type="solid") # #DBEAFE
SECTION_FILL_GREEN = PatternFill(start_color="CCFBF1", end_color="CCFBF1", fill_type="solid") # #CCFBF1

FONT_FAMILY = "Times New Roman"

FONT_TITLE_MAIN = Font(name=FONT_FAMILY, size=15, bold=True, color="1E3A8A")
FONT_TITLE_SUB = Font(name=FONT_FAMILY, size=12, bold=True, color="1F2937")
FONT_HEADER_WHITE = Font(name=FONT_FAMILY, size=11, bold=True, color="FFFFFF")
FONT_SECTION = Font(name=FONT_FAMILY, size=11, bold=True, color="1E3A8A")
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

def auto_fit_columns(ws, max_cols=None, max_width_limit=60):
    ws.views.sheetView[0].showGridLines = True
    cols = max_cols if max_cols else ws.max_column
    for col in range(1, cols + 1):
        col_letter = get_column_letter(col)
        max_len = 0
        for row in range(1, ws.max_row + 1):
            val = ws.cell(row=row, column=col).value
            if val is not None:
                # avoid merged cells calculation error
                lines = str(val).split("\n")
                line_max = max(len(l) for l in lines) if lines else 0
                if line_max > max_len and line_max < 100:
                    max_len = line_max
        adjusted_width = min(max(max_len + 3, 10), max_width_limit)
        ws.column_dimensions[col_letter].width = adjusted_width


# ==============================================================================
# SHEET 1: TỔNG QUAN DANH MỤC & CĂN CỨ PHÁP LÝ
# ==============================================================================
ws1 = wb.create_sheet(title="1. Tổng quan & Căn cứ")
ws1.column_dimensions['A'].width = 8
ws1.column_dimensions['B'].width = 25
ws1.column_dimensions['C'].width = 30
ws1.column_dimensions['D'].width = 20
ws1.column_dimensions['E'].width = 22
ws1.column_dimensions['F'].width = 30

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
ws1["A4"] = "BẢNG TỔNG HỢP DANH MỤC VỊ TRÍ VIỆC LÀM VIÊN CHỨC QUẢN LÝ"
ws1["A4"].font = FONT_TITLE_MAIN
ws1["A4"].alignment = ALIGN_CENTER

ws1.merge_cells("A5:F5")
ws1["A5"] = "Căn cứ Nghị định số 232/2026/NĐ-CP ngày 26/6/2026 của Chính phủ"
ws1["A5"].font = Font(name=FONT_FAMILY, size=11, italic=True)
ws1["A5"].alignment = ALIGN_CENTER

ws1.merge_cells("A6:F6")
ws1["A6"] = "Năm học: 2026 - 2027 | Đơn vị: Trường THPT Phục Hòa, huyện Phục Hòa, tỉnh Cao Bằng"
ws1["A6"].font = FONT_ITALIC
ws1["A6"].alignment = ALIGN_CENTER

# Legal section
ws1.merge_cells("A8:F8")
ws1["A8"] = "I. CƠ SỞ PHÁP LÝ XÂY DỰNG VỊ TRÍ VIỆC LÀM"
ws1["A8"].font = FONT_SECTION
ws1["A8"].fill = SECTION_FILL
apply_row_styles(ws1, 8, font=FONT_SECTION, fill=SECTION_FILL)

legal_rows = [
    ("1", "Nghị định số 232/2026/NĐ-CP", "Chính phủ", "26/06/2026", "01/07/2026", "Quy định về vị trí việc làm viên chức trong đơn vị sự nghiệp công lập"),
    ("2", "Thông tư số 15/2026/TT-BGDĐT", "Bộ GD&ĐT", "24/03/2026", "10/05/2026", "Ban hành Điều lệ trường trung học cơ sở, trường THPT và trường phổ thông có nhiều cấp học"),
    ("3", "Luật Viên chức 2010 (sửa đổi 2019)", "Quốc hội", "15/11/2010", "01/07/2020", "Quy định quyền, nghĩa vụ, tuyển dụng, sử dụng và quản lý viên chức"),
    ("4", "Thông tư số 04/2021/TT-BGDĐT & TT 08/2023/TT-BGDĐT", "Bộ GD&ĐT", "02/02/2021", "20/03/2021", "Quy định mã số, tiêu chuẩn chức danh nghề nghiệp và bổ nhiệm, xếp lương giáo viên THPT"),
    ("5", "Kế hoạch phát triển giáo dục năm học 2026-2027", "THPT Phục Hòa", "15/08/2026", "01/09/2026", "Kế hoạch phân công nhiệm vụ và chỉ tiêu năm học 2026-2027"),
]

headers_legal = ["STT", "Văn bản căn cứ", "Cơ quan ban hành", "Ngày ban hành", "Hiệu lực", "Nội dung điều chỉnh chính"]
for col_i, h in enumerate(headers_legal, 1):
    cell = ws1.cell(row=9, column=col_i, value=h)
    cell.font = FONT_HEADER_WHITE
    cell.fill = BLUE_HEADER
    cell.alignment = ALIGN_CENTER
    cell.border = BORDER_HEADER

for r_i, ldata in enumerate(legal_rows, 10):
    for c_i, val in enumerate(ldata, 1):
        cell = ws1.cell(row=r_i, column=c_i, value=val)
        cell.font = FONT_REGULAR
        cell.border = BORDER_CELL
        cell.alignment = ALIGN_CENTER if c_i in [1, 4, 5] else ALIGN_LEFT

# Summary table section
r_start = 17
ws1.merge_cells(f"A{r_start}:F{r_start}")
ws1[f"A{r_start}"] = "II. DANH MỤC VỊ TRÍ VIỆC LÀM VIÊN CHỨC QUẢN LÝ TỔ CHUYÊN MÔN"
apply_row_styles(ws1, r_start, font=FONT_SECTION, fill=SECTION_FILL)

headers_vtvl = ["STT", "Mã số VTVL", "Tên Vị trí việc làm", "Nhóm vị trí", "Hạng CDNN tối thiểu", "Ghi chú phân loại (NĐ 232/2026)"]
for col_i, h in enumerate(headers_vtvl, 1):
    cell = ws1.cell(row=r_start+1, column=col_i, value=h)
    cell.font = FONT_HEADER_WHITE
    cell.fill = NAVY_FILL
    cell.alignment = ALIGN_CENTER
    cell.border = BORDER_HEADER

vtvl_data = [
    ("1", "TTCM-THPT-01", "Tổ trưởng chuyên môn", "Viên chức quản lý", "Giáo viên THPT Hạng II (Mã: V.07.05.14)", "Viên chức quản lý cấp tổ trường THPT theo NĐ 232/2026. Phụ cấp chức vụ: 0.25"),
    ("2", "TPCM-THPT-01", "Tổ phó chuyên môn", "Viên chức quản lý", "Giáo viên THPT Hạng III (Mã: V.07.05.15)", "Viên chức quản lý cấp tổ trường THPT theo NĐ 232/2026. Phụ cấp chức vụ: 0.15"),
]

for r_i, vdata in enumerate(vtvl_data, r_start+2):
    for c_i, val in enumerate(vdata, 1):
        cell = ws1.cell(row=r_i, column=c_i, value=val)
        cell.font = FONT_BOLD if c_i in [2, 3] else FONT_REGULAR
        cell.border = BORDER_CELL
        cell.alignment = ALIGN_CENTER if c_i in [1, 2, 4, 5] else ALIGN_LEFT

# Signatures
r_sig = r_start + 6
ws1.merge_cells(f"A{r_sig}:C{r_sig}")
ws1[f"A{r_sig}"] = "NGƯỜI LẬP BIỂU"
ws1[f"A{r_sig}"].font = FONT_BOLD
ws1[f"A{r_sig}"].alignment = ALIGN_CENTER

ws1.merge_cells(f"D{r_sig}:F{r_sig}")
ws1[f"D{r_sig}"] = "HIỆU TRƯỞNG PHÊ DUYỆT"
ws1[f"D{r_sig}"].font = FONT_BOLD
ws1[f"D{r_sig}"].alignment = ALIGN_CENTER

ws1.merge_cells(f"A{r_sig+1}:C{r_sig+1}")
ws1[f"A{r_sig+1}"] = "(Ký và ghi rõ họ tên)"
ws1[f"A{r_sig+1}"].font = FONT_ITALIC
ws1[f"A{r_sig+1}"].alignment = ALIGN_CENTER

ws1.merge_cells(f"D{r_sig+1}:F{r_sig+1}")
ws1[f"D{r_sig+1}"] = "(Ký, đóng dấu và ghi rõ họ tên)"
ws1[f"D{r_sig+1}"].font = FONT_ITALIC
ws1[f"D{r_sig+1}"].alignment = ALIGN_CENTER


# ==============================================================================
# SHEET 2: BẢN MÔ TẢ VTVL TỔ TRƯỞNG CHUYÊN MÔN (TTCM)
# ==============================================================================
ws2 = wb.create_sheet(title="2. VTVL Tổ trưởng CM")
ws2.column_dimensions['A'].width = 8
ws2.column_dimensions['B'].width = 26
ws2.column_dimensions['C'].width = 38
ws2.column_dimensions['D'].width = 24
ws2.column_dimensions['E'].width = 14
ws2.column_dimensions['F'].width = 28

# Title block
ws2.merge_cells("A1:F1")
ws2["A1"] = "BẢN MÔ TẢ VỊ TRÍ VIỆC LÀM VIÊN CHỨC QUẢN LÝ"
ws2["A1"].font = FONT_TITLE_MAIN
ws2["A1"].alignment = ALIGN_CENTER

ws2.merge_cells("A2:F2")
ws2["A2"] = "VỊ TRÍ: TỔ TRƯỞNG CHUYÊN MÔN (MÃ SỐ: TTCM-THPT-01)"
ws2["A2"].font = Font(name=FONT_FAMILY, size=13, bold=True, color="1E3A8A")
ws2["A2"].alignment = ALIGN_CENTER

ws2.merge_cells("A3:F3")
ws2["A3"] = "Áp dụng tại: Trường THPT Phục Hòa - Theo Nghị định số 232/2026/NĐ-CP"
ws2["A3"].font = FONT_ITALIC
ws2["A3"].alignment = ALIGN_CENTER

# Part I: General Information
ws2.merge_cells("A5:F5")
ws2["A5"] = "PHẦN I. THÔNG TIN CHUNG VỀ VỊ TRÍ VIỆC LÀM"
apply_row_styles(ws2, 5, font=FONT_SECTION, fill=SECTION_FILL)

ttcm_info = [
    ("1.1", "Tên vị trí việc làm:", "Tổ trưởng chuyên môn", "1.2", "Mã số vị trí:", "TTCM-THPT-01"),
    ("1.3", "Nhóm vị trí việc làm:", "Viên chức quản lý (theo Điều 3 & Phụ lục I NĐ 232/2026/NĐ-CP)", "1.4", "Cấp quản lý:", "Cấp tổ chuyên môn"),
    ("1.5", "Đơn vị công tác:", "Trường THPT Phục Hòa - Sở GD&ĐT Cao Bằng", "1.6", "Người quản lý trực tiếp:", "Hiệu trưởng / Phó Hiệu trưởng phụ trách chuyên môn"),
    ("1.7", "Đối tượng quản lý trực tiếp:", "Tổ phó chuyên môn, giáo viên và nhân viên trực thuộc tổ", "1.8", "Chức danh nghề nghiệp tối thiểu:", "Giáo viên THPT Hạng II (Mã: V.07.05.14)"),
    ("1.9", "Phụ cấp chức vụ lãnh đạo:", "Hệ số 0.25 (theo quy định hiện hành đối với trường hạng II/III)", "1.10", "Định mức giảm tiết giảng dạy:", "Giảm 03 tiết dạy/tuần theo quy định"),
]

for idx, rdata in enumerate(ttcm_info, 6):
    ws2.cell(row=idx, column=1, value=rdata[0]).alignment = ALIGN_CENTER
    ws2.cell(row=idx, column=2, value=rdata[1]).font = FONT_BOLD
    ws2.cell(row=idx, column=3, value=rdata[2]).font = FONT_REGULAR
    ws2.cell(row=idx, column=4, value=rdata[3]).alignment = ALIGN_CENTER
    ws2.cell(row=idx, column=5, value=rdata[4]).font = FONT_BOLD
    ws2.cell(row=idx, column=6, value=rdata[5]).font = FONT_REGULAR
    for col_c in range(1, 7):
        ws2.cell(row=idx, column=col_c).border = BORDER_CELL

# Part II: Objectives
r_obj = 12
ws2.merge_cells(f"A{r_obj}:F{r_obj}")
ws2[f"A{r_obj}"] = "PHẦN II. MỤC TIÊU VÀ SỨ MỆNH CỦA VỊ TRÍ VIỆC LÀM"
apply_row_styles(ws2, r_obj, font=FONT_SECTION, fill=SECTION_FILL)

ws2.merge_cells(f"A{r_obj+1}:F{r_obj+1}")
ws2[f"A{r_obj+1}"] = "Lãnh đạo, điều hành toàn diện công tác chuyên môn của tổ; tổ chức triển khai hiệu quả Chương trình Giáo dục phổ thông 2018; nâng cao chất lượng dạy và học; quản lý, bồi dưỡng và phát triển năng lực đội ngũ giáo viên trong tổ; tiên phong đổi mới phương pháp, ứng dụng chuyển đổi số và công nghệ AI trong giáo dục; chịu trách nhiệm trước Hiệu trưởng về toàn bộ hoạt động chuyên môn của tổ."
ws2[f"A{r_obj+1}"].alignment = ALIGN_LEFT
ws2[f"A{r_obj+1}"].font = FONT_REGULAR
apply_row_styles(ws2, r_obj+1, border=BORDER_CELL)

# Part III: Duties & Responsibilities
r_duties = 15
ws2.merge_cells(f"A{r_duties}:F{r_duties}")
ws2[f"A{r_duties}"] = "PHẦN III. NHIỆM VỤ, TRÁCH NHIỆM VÀ SẢN PHẨM ĐẦU RA CỤ THỂ"
apply_row_styles(ws2, r_duties, font=FONT_SECTION, fill=SECTION_FILL)

headers_duties = ["STT", "Nhiệm vụ chính", "Hoạt động công việc cụ thể", "Tiêu chí đánh giá hoàn thành", "Tỷ trọng (%)", "Sản phẩm / Kết quả đầu ra"]
for col_i, h in enumerate(headers_duties, 1):
    cell = ws2.cell(row=r_duties+1, column=col_i, value=h)
    cell.font = FONT_HEADER_WHITE
    cell.fill = BLUE_HEADER
    cell.alignment = ALIGN_CENTER
    cell.border = BORDER_HEADER

ttcm_duties = [
    ("1", "Xây dựng và triển khai kế hoạch giáo dục của tổ", 
     "- Xây dựng Kế hoạch dạy học các môn học/HĐTN của tổ theo CT GDPT 2018.\n- Phân phối chương trình, xây dựng ma trận đặc tả đề kiểm tra định kỳ.\n- Trình Ban Giám hiệu phê duyệt và giám sát việc thực hiện đúng tiến độ.",
     "- Kế hoạch hoàn thành trước ngày 25/8 hàng năm.\n- 100% giáo viên thực hiện đúng kế hoạch.\n- Được Hiệu trưởng phê duyệt.",
     "20%",
     "- Kế hoạch giáo dục tổ chuyên môn năm học.\n- Kế hoạch dạy học từng môn lớp 10, 11, 12.\n- Phân phối chương trình chi tiết."),
     
    ("2", "Tổ chức sinh hoạt chuyên môn & đổi mới phương pháp",
     "- Tổ chức sinh hoạt chuyên môn định kỳ ít nhất 2 lần/tháng.\n- Chỉ đạo sinh hoạt chuyên môn theo nghiên cứu bài học (tối thiểu 02 chuyên đề/học kỳ).\n- Triển khai phương pháp dạy học tích cực, giáo dục STEM/STEAM.\n- Đổi mới kiểm tra đánh giá theo định hướng phẩm chất năng lực học sinh.",
     "- Đủ số buổi sinh hoạt theo quy định (>= 2 lần/tháng).\n- Có biên bản và hồ sơ sinh hoạt chuyên môn lưu trữ số hóa.\n- 100% GV tham gia đầy đủ.",
     "20%",
     "- Biên bản sinh hoạt chuyên môn số hóa.\n- Báo cáo chuyên đề nghiên cứu bài học.\n- Kế hoạch bài dạy (giáo án) mẫu của tổ."),
     
    ("3", "Quản lý, kiểm tra nội bộ và bồi dưỡng giáo viên",
     "- Dự giờ, thăm lớp (tối thiểu 2-3 tiết/GV/học kỳ).\n- Kiểm tra hồ sơ giáo án, sổ điểm, đề kiểm tra của giáo viên theo kế hoạch.\n- Tham gia đánh giá chuẩn nghề nghiệp GV và xếp loại thi đua cuối kỳ/năm học.\n- Bồi dưỡng GV mới, GV tham gia thi GVDG các cấp.",
     "- 100% GV trong tổ được dự giờ và kiểm tra hồ sơ đúng định kỳ.\n- Nhận xét, đánh giá khách quan, công tâm, đúng quy trình.",
     "15%",
     "- Phiếu đánh giá dự giờ giáo viên.\n- Biên bản kiểm tra chuyên môn nội bộ.\n- Bảng tổng hợp đánh giá chuẩn nghề nghiệp GV."),
     
    ("4", "Chỉ đạo công tác mũi nhọn & phụ đạo học sinh",
     "- Xây dựng kế hoạch bồi dưỡng học sinh giỏi cấp trường, cấp tỉnh.\n- Phân công GV phụ trách đội tuyển HSG, theo dõi và đánh giá định kỳ.\n- Tổ chức phụ đạo học sinh có học lực chưa đạt chuẩn, giảm tỷ lệ yếu kém.\n- Hướng dẫn học sinh tham gia cuộc thi KHKT, STEM cấp trường, cấp tỉnh.",
     "- Đạt và vượt chỉ tiêu giải HSG cấp tỉnh do Nhà trường giao.\n- Tỷ lệ tốt nghiệp môn thi đạt mặt bằng chung toàn tỉnh.",
     "15%",
     "- Kế hoạch bồi dưỡng HSG & phụ đạo yếu kém.\n- Danh sách đội tuyển và kết quả thi HSG.\n- Báo cáo kết quả dự thi KHKT cấp tỉnh."),
     
    ("5", "Trực tiếp tham gia giảng dạy trên lớp",
     "- Thực hiện giảng dạy môn học theo đúng phân công chuyên môn của trường.\n- Định mức tiết dạy: 14 tiết/tuần (đã giảm 3 tiết nhiệm vụ TTCM theo quy định).\n- Đảm bảo chất lượng bài dạy, soạn bài và chấm trả bài đúng quy chế.",
     "- Đạt 100% số tiết dạy theo định mức.\n- Hồ sơ giáo án đạt chuẩn, đúng quy định chuyên môn.\n- Tỷ lệ HS đạt yêu cầu bộ môn >= 95%.",
     "20%",
     "- Kế hoạch bài dạy (giáo án) cá nhân.\n- Sổ theo dõi đánh giá học sinh.\n- Kết quả kiểm tra, đánh giá học sinh trên lớp."),
     
    ("6", "Chuyển đổi số & ứng dụng AI trong dạy học",
     "- Triển khai sử dụng sổ điểm điện tử, học bạ số, ký số trên phần mềm quản lý.\n- Hướng dẫn GV tổ ứng dụng công nghệ thông tin, phần mềm mô phỏng, công cụ AI.\n- Xây dựng kho học liệu số, ngân hàng câu hỏi trắc nghiệm của tổ.",
     "- 100% GV thực hiện sổ sách điện tử, ký số đúng hạn.\n- Đóng góp tối thiểu 20 học liệu số/học kỳ vào kho dùng chung.",
     "10%",
     "- Kho học liệu số của tổ chuyên môn.\n- Ngân hàng đề thi/kiểm tra số hóa.\n- Báo cáo ứng dụng CNTT và AI của tổ."),
]

for idx, rdata in enumerate(ttcm_duties, r_duties+2):
    ws2.cell(row=idx, column=1, value=rdata[0]).alignment = ALIGN_CENTER
    ws2.cell(row=idx, column=2, value=rdata[1]).font = FONT_BOLD
    ws2.cell(row=idx, column=2).alignment = ALIGN_LEFT
    ws2.cell(row=idx, column=3, value=rdata[2]).font = FONT_REGULAR
    ws2.cell(row=idx, column=3).alignment = ALIGN_LEFT
    ws2.cell(row=idx, column=4, value=rdata[3]).font = FONT_REGULAR
    ws2.cell(row=idx, column=4).alignment = ALIGN_LEFT
    ws2.cell(row=idx, column=5, value=rdata[4]).font = FONT_BOLD
    ws2.cell(row=idx, column=5).alignment = ALIGN_CENTER
    ws2.cell(row=idx, column=6, value=rdata[5]).font = FONT_REGULAR
    ws2.cell(row=idx, column=6).alignment = ALIGN_LEFT
    for col_c in range(1, 7):
        ws2.cell(row=idx, column=col_c).border = BORDER_CELL

# Part IV: Requirements / Qualifications
r_req = r_duties + len(ttcm_duties) + 3
ws2.merge_cells(f"A{r_req}:F{r_req}")
ws2[f"A{r_req}"] = "PHẦN IV. YÊU CẦU TIÊU CHUẨN, NĂNG LỰC ĐỐI VỚI VỊ TRÍ VIỆC LÀM"
apply_row_styles(ws2, r_req, font=FONT_SECTION, fill=SECTION_FILL)

headers_req = ["STT", "Nhóm tiêu chuẩn", "Yêu cầu chi tiết theo Nghị định 232/2026/NĐ-CP & TT Bộ GD&ĐT", "Minh chứng / Hồ sơ yêu cầu", "", ""]
for col_i in range(1, 4):
    cell = ws2.cell(row=r_req+1, column=col_i, value=headers_req[col_i-1])
    cell.font = FONT_HEADER_WHITE
    cell.fill = BLUE_HEADER
    cell.alignment = ALIGN_CENTER
    cell.border = BORDER_HEADER
ws2.merge_cells(f"C{r_req+1}:D{r_req+1}")
ws2.merge_cells(f"E{r_req+1}:F{r_req+1}")
ws2.cell(row=r_req+1, column=5, value="Minh chứng / Hồ sơ yêu cầu").font = FONT_HEADER_WHITE
ws2.cell(row=r_req+1, column=5).fill = BLUE_HEADER
ws2.cell(row=r_req+1, column=5).alignment = ALIGN_CENTER
ws2.cell(row=r_req+1, column=5).border = BORDER_HEADER
ws2.cell(row=r_req+1, column=6).border = BORDER_HEADER

req_data = [
    ("1", "Trình độ chuyên môn đào tạo", "Có bằng Cử nhân trở lên thuộc ngành đào tạo giáo viên hoặc có bằng cử nhân chuyên ngành phù hợp kèm chứng chỉ bồi dưỡng nghiệp vụ sư phạm.", "Bằng tốt nghiệp đại học / Thạc sĩ; Bảng điểm; Chứng chỉ NVSP"),
    ("2", "Tiêu chuẩn CDNN & Ngạch bậc", "Được bổ nhiệm chức danh nghề nghiệp Giáo viên THPT Hạng II (Mã: V.07.05.14) trở lên; có Chứng chỉ bồi dưỡng theo tiêu chuẩn CDNN giáo viên THPT.", "Quyết định bổ nhiệm CDNN Hạng II; Chứng chỉ bồi dưỡng CDNN"),
    ("3", "Năng lực quản lý điều hành", "Có năng lực lập kế hoạch, tổ chức, điều hành, giải quyết xung đột; có kỹ năng quản lý tổ nhóm chuyên môn; nắm vững quy chế chuyên môn trường THPT.", "Biên bản đánh giá năng lực; Quyết định bổ nhiệm TTCM"),
    ("4", "Năng lực tin học & Chuyển đổi số", "Có kỹ năng sử dụng công nghệ thông tin cơ bản; sử dụng thành thạo phần mềm quản trị nhà trường, hệ thống quản lý học tập (LMS), ứng dụng AI trong giảng dạy.", "Chứng chỉ CNTT / Sản phẩm số hóa thực tế / Học bạ điện tử"),
    ("5", "Phẩm chất đạo đức & Chính trị", "Đảng viên Đảng Cộng sản Việt Nam (khuyến khích/ưu tiên); có phẩm chất chính trị vững vàng, đạo đức nhà giáo mẫu mực; đạt Chuẩn nghề nghiệp mức Tốt.", "Bản kiểm điểm Đảng viên; Phiếu đánh giá Chuẩn nghề nghiệp GV"),
]

for idx, rdata in enumerate(req_data, r_req+2):
    ws2.cell(row=idx, column=1, value=rdata[0]).alignment = ALIGN_CENTER
    ws2.cell(row=idx, column=2, value=rdata[1]).font = FONT_BOLD
    ws2.cell(row=idx, column=2).alignment = ALIGN_LEFT
    ws2.merge_cells(f"C{idx}:D{idx}")
    ws2.cell(row=idx, column=3, value=rdata[2]).font = FONT_REGULAR
    ws2.cell(row=idx, column=3).alignment = ALIGN_LEFT
    ws2.merge_cells(f"E{idx}:F{idx}")
    ws2.cell(row=idx, column=5, value=rdata[3]).font = FONT_REGULAR
    ws2.cell(row=idx, column=5).alignment = ALIGN_LEFT
    for col_c in range(1, 7):
        ws2.cell(row=idx, column=col_c).border = BORDER_CELL

# Signatures for TTCM
r_sig2 = r_req + len(req_data) + 3
ws2.merge_cells(f"A{r_sig2}:C{r_sig2}")
ws2[f"A{r_sig2}"] = "TỔ TRƯỞNG CHUYÊN MÔN"
ws2[f"A{r_sig2}"].font = FONT_BOLD
ws2[f"A{r_sig2}"].alignment = ALIGN_CENTER

ws2.merge_cells(f"D{r_sig2}:F{r_sig2}")
ws2[f"D{r_sig2}"] = "HIỆU TRƯỞNG DUYỆT"
ws2[f"D{r_sig2}"].font = FONT_BOLD
ws2[f"D{r_sig2}"].alignment = ALIGN_CENTER

ws2.merge_cells(f"A{r_sig2+1}:C{r_sig2+1}")
ws2[f"A{r_sig2+1}"] = "(Ký và ghi rõ họ tên)"
ws2[f"A{r_sig2+1}"].font = FONT_ITALIC
ws2[f"A{r_sig2+1}"].alignment = ALIGN_CENTER

ws2.merge_cells(f"D{r_sig2+1}:F{r_sig2+1}")
ws2[f"D{r_sig2+1}"] = "(Ký, đóng dấu và ghi rõ họ tên)"
ws2[f"D{r_sig2+1}"].font = FONT_ITALIC
ws2[f"D{r_sig2+1}"].alignment = ALIGN_CENTER


# ==============================================================================
# SHEET 3: BẢN MÔ TẢ VTVL TỔ PHÓ CHUYÊN MÔN (TPCM)
# ==============================================================================
ws3 = wb.create_sheet(title="3. VTVL Tổ phó CM")
ws3.column_dimensions['A'].width = 8
ws3.column_dimensions['B'].width = 26
ws3.column_dimensions['C'].width = 38
ws3.column_dimensions['D'].width = 24
ws3.column_dimensions['E'].width = 14
ws3.column_dimensions['F'].width = 28

# Title block
ws3.merge_cells("A1:F1")
ws3["A1"] = "BẢN MÔ TẢ VỊ TRÍ VIỆC LÀM VIÊN CHỨC QUẢN LÝ"
ws3["A1"].font = FONT_TITLE_MAIN
ws3["A1"].alignment = ALIGN_CENTER

ws3.merge_cells("A2:F2")
ws3["A2"] = "VỊ TRÍ: TỔ PHÓ CHUYÊN MÔN (MÃ SỐ: TPCM-THPT-01)"
ws3["A2"].font = Font(name=FONT_FAMILY, size=13, bold=True, color="0D9488")
ws3["A2"].alignment = ALIGN_CENTER

ws3.merge_cells("A3:F3")
ws3["A3"] = "Áp dụng tại: Trường THPT Phục Hòa - Theo Nghị định số 232/2026/NĐ-CP"
ws3["A3"].font = FONT_ITALIC
ws3["A3"].alignment = ALIGN_CENTER

# Part I: General Info
ws3.merge_cells("A5:F5")
ws3["A5"] = "PHẦN I. THÔNG TIN CHUNG VỀ VỊ TRÍ VIỆC LÀM"
apply_row_styles(ws3, 5, font=FONT_SECTION_GREEN, fill=SECTION_FILL_GREEN)

tpcm_info = [
    ("1.1", "Tên vị trí việc làm:", "Tổ phó chuyên môn", "1.2", "Mã số vị trí:", "TPCM-THPT-01"),
    ("1.3", "Nhóm vị trí việc làm:", "Viên chức quản lý (theo Điều 3 & Phụ lục I NĐ 232/2026/NĐ-CP)", "1.4", "Cấp quản lý:", "Cấp tổ chuyên môn"),
    ("1.5", "Đơn vị công tác:", "Trường THPT Phục Hòa - Sở GD&ĐT Cao Bằng", "1.6", "Người quản lý trực tiếp:", "Tổ trưởng chuyên môn & Hiệu trưởng"),
    ("1.7", "Quan hệ phối hợp:", "Phối hợp với TTCM điều hành giáo viên trong phân môn/khối phụ trách", "1.8", "Chức danh nghề nghiệp tối thiểu:", "Giáo viên THPT Hạng III (Mã: V.07.05.15)"),
    ("1.9", "Phụ cấp chức vụ lãnh đạo:", "Hệ số 0.15 (theo quy định hiện hành đối với trường hạng II/III)", "1.10", "Định mức giảm tiết giảng dạy:", "Giảm 01 tiết dạy/tuần theo quy định"),
]

for idx, rdata in enumerate(tpcm_info, 6):
    ws3.cell(row=idx, column=1, value=rdata[0]).alignment = ALIGN_CENTER
    ws3.cell(row=idx, column=2, value=rdata[1]).font = FONT_BOLD
    ws3.cell(row=idx, column=3, value=rdata[2]).font = FONT_REGULAR
    ws3.cell(row=idx, column=4, value=rdata[3]).alignment = ALIGN_CENTER
    ws3.cell(row=idx, column=5, value=rdata[4]).font = FONT_BOLD
    ws3.cell(row=idx, column=6, value=rdata[5]).font = FONT_REGULAR
    for col_c in range(1, 7):
        ws3.cell(row=idx, column=col_c).border = BORDER_CELL

# Part II: Objectives
r_obj3 = 12
ws3.merge_cells(f"A{r_obj3}:F{r_obj3}")
ws3[f"A{r_obj3}"] = "PHẦN II. MỤC TIÊU VÀ SỨ MỆNH CỦA VỊ TRÍ VIỆC LÀM"
apply_row_styles(ws3, r_obj3, font=FONT_SECTION_GREEN, fill=SECTION_FILL_GREEN)

ws3.merge_cells(f"A{r_obj3+1}:F{r_obj3+1}")
ws3[f"A{r_obj3+1}"] = "Trợ giúp Tổ trưởng chuyên môn quản lý, điều hành các hoạt động chuyên môn theo lĩnh vực phân công; quản lý hồ sơ chuyên môn, theo dõi tiến độ thực hiện chương trình; phụ trách công tác thiết bị thí nghiệm, học liệu số và các hoạt động trải nghiệm, STEM; thay mặt Tổ trưởng điều hành tổ khi Tổ trưởng đi vắng; chịu trách nhiệm trước Tổ trưởng và Hiệu trưởng về lĩnh vực được giao."
ws3[f"A{r_obj3+1}"].alignment = ALIGN_LEFT
ws3[f"A{r_obj3+1}"].font = FONT_REGULAR
apply_row_styles(ws3, r_obj3+1, border=BORDER_CELL)

# Part III: Duties
r_duties3 = 15
ws3.merge_cells(f"A{r_duties3}:F{r_duties3}")
ws3[f"A{r_duties3}"] = "PHẦN III. NHIỆM VỤ, TRÁCH NHIỆM VÀ SẢN PHẨM ĐẦU RA CỤ THỂ"
apply_row_styles(ws3, r_duties3, font=FONT_SECTION_GREEN, fill=SECTION_FILL_GREEN)

for col_i, h in enumerate(headers_duties, 1):
    cell = ws3.cell(row=r_duties3+1, column=col_i, value=h)
    cell.font = FONT_HEADER_WHITE
    cell.fill = TEAL_HEADER
    cell.alignment = ALIGN_CENTER
    cell.border = BORDER_HEADER

tpcm_duties = [
    ("1", "Giúp việc Tổ trưởng điều hành chuyên môn theo phân công",
     "- Cùng Tổ trưởng xây dựng kế hoạch giáo dục, kế hoạch dạy học năm học.\n- Trực tiếp phụ trách một nhóm môn hoặc một khối lớp theo phân công của tổ.\n- Theo dõi, đôn đốc thực hiện quy chế chuyên môn, nền nếp dạy học.",
     "- Kế hoạch khối/nhóm môn hoàn thành đúng hạn.\n- Đôn đốc 100% GV thực hiện đúng tiến độ chương trình.",
     "20%",
     "- Kế hoạch hoạt động nhóm chuyên môn/khối lớp.\n- Bảng theo dõi tiến độ chương trình hàng tháng."),
     
    ("2", "Quản lý hồ sơ chuyên môn & tiến độ chương trình",
     "- Kiểm tra việc vào điểm, ghi sổ điểm điện tử, ghi sổ đầu bài của giáo viên.\n- Rà soát tiến độ giảng dạy, kịp thời phát hiện dạy bù, dạy chậm chương trình.\n- Tổng hợp báo cáo chuyên môn định kỳ hàng tháng/học kỳ gửi TTCM.",
     "- Báo cáo tiến độ chính xác, gửi trước ngày 25 hàng tháng.\n- Không để xảy ra tình trạng cắt xén hoặc chậm chương trình.",
     "15%",
     "- Báo cáo tiến độ chương trình tháng/học kỳ.\n- Biên bản kiểm tra sổ đầu bài, sổ điểm điện tử."),
     
    ("3", "Quản lý thiết bị dạy học, phòng bộ môn & học liệu số",
     "- Lập kế hoạch sử dụng và bảo quản thiết bị dạy học, phòng thí nghiệm thực hành.\n- Kiểm tra, đôn đốc giáo viên sử dụng đồ dùng dạy học theo bài dạy.\n- Tham gia kiểm kê, đề xuất mua sắm, bổ sung thiết bị dạy học hàng năm.",
     "- 100% tiết thực hành thí nghiệm được thực hiện đầy đủ theo phân phối.\n- Sổ mượn trả thiết bị được ghi chép rõ ràng, đúng quy định.",
     "15%",
     "- Kế hoạch sử dụng thiết bị dạy học năm học.\n- Sổ theo dõi mượn trả thiết bị, phòng bộ môn.\n- Biên bản kiểm kê tài sản, thiết bị cuối năm."),
     
    ("4", "Phụ trách hoạt động STEM, trải nghiệm & phong trào thi đua",
     "- Tổ chức ngày hội STEM, hoạt động trải nghiệm hướng nghiệp cấp tổ.\n- Theo dõi các phong trào thi đua 'Dạy tốt - Học tốt', thao giảng, hội giảng.\n- Tổng hợp số liệu thi đua của giáo viên trong tổ phục vụ bình xét cuối kỳ.",
     "- Tổ chức ít nhất 01 hoạt động STEM/trải nghiệm cấp tổ mỗi học kỳ.\n- Hồ sơ thi đua minh bạch, công khai, chính xác.",
     "15%",
     "- Kế hoạch và kịch bản hoạt động STEM/trải nghiệm.\n- Bảng tổng hợp thi đua chuyên môn của tổ."),
     
    ("5", "Trực tiếp tham gia giảng dạy trên lớp",
     "- Thực hiện giảng dạy môn học theo đúng phân công chuyên môn của nhà trường.\n- Định mức tiết dạy: 16 tiết/tuần (đã giảm 1 tiết nhiệm vụ TPCM theo quy định).\n- Đảm bảo chất lượng giáo dục bộ môn, hoàn thành tốt kế hoạch bài dạy.",
     "- Đạt 100% định mức tiết dạy.\n- Hồ sơ giáo án đạt loại Tốt.\n- Tỷ lệ HS đạt yêu cầu bộ môn >= 95%.",
     "25%",
     "- Kế hoạch bài dạy cá nhân.\n- Sổ theo dõi đánh giá học sinh.\n- Đề kiểm tra và đáp án các môn trực tiếp dạy."),
     
    ("6", "Thực hiện nhiệm vụ ủy quyền khi TTCM vắng mặt",
     "- Điều hành sinh hoạt tổ, ký duyệt hồ sơ khi được Tổ trưởng ủy quyền bằng văn bản.\n- Báo cáo kịp thời tình hình đột xuất cho Ban Giám hiệu khi Tổ trưởng vắng mặt.",
     "- Điều hành công việc thông suốt, không gián đoạn hoạt động của tổ.",
     "10%",
     "- Biên bản sinh hoạt chuyên môn (trong thời gian ủy quyền).\n- Báo cáo công việc đột xuất cho Ban Giám hiệu."),
]

for idx, rdata in enumerate(tpcm_duties, r_duties3+2):
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
    for col_c in range(1, 7):
        ws3.cell(row=idx, column=col_c).border = BORDER_CELL

# Part IV: Requirements for TPCM
r_req3 = r_duties3 + len(tpcm_duties) + 3
ws3.merge_cells(f"A{r_req3}:F{r_req3}")
ws3[f"A{r_req3}"] = "PHẦN IV. YÊU CẦU TIÊU CHUẨN, NĂNG LỰC ĐỐI VỚI VỊ TRÍ VIỆC LÀM"
apply_row_styles(ws3, r_req3, font=FONT_SECTION_GREEN, fill=SECTION_FILL_GREEN)

for col_i in range(1, 4):
    cell = ws3.cell(row=r_req3+1, column=col_i, value=headers_req[col_i-1])
    cell.font = FONT_HEADER_WHITE
    cell.fill = TEAL_HEADER
    cell.alignment = ALIGN_CENTER
    cell.border = BORDER_HEADER
ws3.merge_cells(f"C{r_req3+1}:D{r_req3+1}")
ws3.merge_cells(f"E{r_req3+1}:F{r_req3+1}")
ws3.cell(row=r_req3+1, column=5, value="Minh chứng / Hồ sơ yêu cầu").font = FONT_HEADER_WHITE
ws3.cell(row=r_req3+1, column=5).fill = TEAL_HEADER
ws3.cell(row=r_req3+1, column=5).alignment = ALIGN_CENTER
ws3.cell(row=r_req3+1, column=5).border = BORDER_HEADER
ws3.cell(row=r_req3+1, column=6).border = BORDER_HEADER

req_tpcm = [
    ("1", "Trình độ chuyên môn đào tạo", "Có bằng Cử nhân trở lên thuộc ngành đào tạo giáo viên hoặc có bằng cử nhân chuyên ngành phù hợp kèm chứng chỉ bồi dưỡng nghiệp vụ sư phạm.", "Bằng tốt nghiệp đại học / Thạc sĩ; Bảng điểm; Chứng chỉ NVSP"),
    ("2", "Tiêu chuẩn CDNN & Ngạch bậc", "Được bổ nhiệm chức danh nghề nghiệp Giáo viên THPT Hạng III (Mã: V.07.05.15) trở lên (khuyến khích Hạng II); có Chứng chỉ bồi dưỡng theo tiêu chuẩn CDNN.", "Quyết định bổ nhiệm CDNN; Chứng chỉ bồi dưỡng CDNN"),
    ("3", "Năng lực phối hợp & Quản lý", "Có tinh thần trách nhiệm cao, năng lực phối hợp tốt; phương pháp làm việc khoa học, tỉ mỉ; có khả năng bao quát quản lý thiết bị, hồ sơ sổ sách.", "Đánh giá xếp loại viên chức hàng năm; Quyết định bổ nhiệm TPCM"),
    ("4", "Năng lực tin học & Thiết bị", "Sử dụng thành thạo máy vi tính, phần mềm văn phòng, quản lý thiết bị dạy học và các ứng dụng giáo dục trực tuyến.", "Chứng chỉ CNTT / Kế hoạch số hóa thiết bị phòng thí nghiệm"),
    ("5", "Phẩm chất đạo đức nghề nghiệp", "Gương mẫu, tận tụy, có tinh thần cầu tiến và đoàn kết nội bộ; đạt Chuẩn nghề nghiệp giáo viên từ mức Khá trở lên.", "Phiếu đánh giá Chuẩn nghề nghiệp GV hàng năm"),
]

for idx, rdata in enumerate(req_tpcm, r_req3+2):
    ws3.cell(row=idx, column=1, value=rdata[0]).alignment = ALIGN_CENTER
    ws3.cell(row=idx, column=2, value=rdata[1]).font = FONT_BOLD
    ws3.cell(row=idx, column=2).alignment = ALIGN_LEFT
    ws3.merge_cells(f"C{idx}:D{idx}")
    ws3.cell(row=idx, column=3, value=rdata[2]).font = FONT_REGULAR
    ws3.cell(row=idx, column=3).alignment = ALIGN_LEFT
    ws3.merge_cells(f"E{idx}:F{idx}")
    ws3.cell(row=idx, column=5, value=rdata[3]).font = FONT_REGULAR
    ws3.cell(row=idx, column=5).alignment = ALIGN_LEFT
    for col_c in range(1, 7):
        ws3.cell(row=idx, column=col_c).border = BORDER_CELL

# Signatures for TPCM
r_sig3 = r_req3 + len(req_tpcm) + 3
ws3.merge_cells(f"A{r_sig3}:B{r_sig3}")
ws3[f"A{r_sig3}"] = "TỔ PHÓ CHUYÊN MÔN"
ws3[f"A{r_sig3}"].font = FONT_BOLD
ws3[f"A{r_sig3}"].alignment = ALIGN_CENTER

ws3.merge_cells(f"C{r_sig3}:D{r_sig3}")
ws3[f"C{r_sig3}"] = "TỔ TRƯỞNG XÁC NHẬN"
ws3[f"C{r_sig3}"].font = FONT_BOLD
ws3[f"C{r_sig3}"].alignment = ALIGN_CENTER

ws3.merge_cells(f"E{r_sig3}:F{r_sig3}")
ws3[f"E{r_sig3}"] = "HIỆU TRƯỞNG PHÊ DUYỆT"
ws3[f"E{r_sig3}"].font = FONT_BOLD
ws3[f"E{r_sig3}"].alignment = ALIGN_CENTER

ws3.merge_cells(f"A{r_sig3+1}:B{r_sig3+1}")
ws3[f"A{r_sig3+1}"] = "(Ký và ghi rõ họ tên)"
ws3[f"A{r_sig3+1}"].font = FONT_ITALIC
ws3[f"A{r_sig3+1}"].alignment = ALIGN_CENTER

ws3.merge_cells(f"C{r_sig3+1}:D{r_sig3+1}")
ws3[f"C{r_sig3+1}"] = "(Ký và ghi rõ họ tên)"
ws3[f"C{r_sig3+1}"].font = FONT_ITALIC
ws3[f"C{r_sig3+1}"].alignment = ALIGN_CENTER

ws3.merge_cells(f"E{r_sig3+1}:F{r_sig3+1}")
ws3[f"E{r_sig3+1}"] = "(Ký, đóng dấu và ghi rõ họ tên)"
ws3[f"E{r_sig3+1}"].font = FONT_ITALIC
ws3[f"E{r_sig3+1}"].alignment = ALIGN_CENTER


# ==============================================================================
# SHEET 4: KHUNG NĂNG LỰC & TIÊU CHÍ KPI ĐÁNH GIÁ (NGHỊ ĐỊNH 232/2026)
# ==============================================================================
ws4 = wb.create_sheet(title="4. Khung năng lực & KPI")
ws4.column_dimensions['A'].width = 8
ws4.column_dimensions['B'].width = 24
ws4.column_dimensions['C'].width = 36
ws4.column_dimensions['D'].width = 16
ws4.column_dimensions['E'].width = 16
ws4.column_dimensions['F'].width = 26

ws4.merge_cells("A1:F1")
ws4["A1"] = "KHUNG NĂNG LỰC VÀ BỘ TIÊU CHÍ KPI ĐÁNH GIÁ VIÊN CHỨC QUẢN LÝ"
ws4["A1"].font = FONT_TITLE_MAIN
ws4["A1"].alignment = ALIGN_CENTER

ws4.merge_cells("A2:F2")
ws4["A2"] = "Theo định hướng Nghị định số 232/2026/NĐ-CP - Trường THPT Phục Hòa"
ws4["A2"].font = FONT_ITALIC
ws4["A2"].alignment = ALIGN_CENTER

# Table 1: Competency Framework
ws4.merge_cells("A4:F4")
ws4["A4"] = "I. KHUNG NĂNG LỰC VỊ TRÍ VIỆC LÀM VIÊN CHỨC QUẢN LÝ CẤP TỔ"
apply_row_styles(ws4, 4, font=FONT_SECTION, fill=SECTION_FILL)

headers_knl = ["STT", "Nhóm năng lực", "Tên năng lực & Định nghĩa chi tiết", "Cấp độ yêu cầu TTCM", "Cấp độ yêu cầu TPCM", "Phương thức đánh giá"]
for col_i, h in enumerate(headers_knl, 1):
    cell = ws4.cell(row=5, column=col_i, value=h)
    cell.font = FONT_HEADER_WHITE
    cell.fill = NAVY_FILL
    cell.alignment = ALIGN_CENTER
    cell.border = BORDER_HEADER

knl_data = [
    ("1", "Năng lực chung cốt lõi", "1.1 Bản lĩnh chính trị, tư tưởng kiên định, chấp hành pháp luật", "Mức 4/4 (Thành thạo)", "Mức 4/4 (Thành thạo)", "Phiếu đánh giá Đảng viên, viên chức"),
    ("2", "Năng lực chung cốt lõi", "1.2 Đạo đức nhà giáo mẫu mực, tác phong sư phạm chuẩn mực", "Mức 4/4 (Xuất sắc)", "Mức 4/4 (Xuất sắc)", "Đánh giá chuẩn nghề nghiệp GV"),
    ("3", "Năng lực chuyên môn", "2.1 Năng lực nắm vững & triển khai CT GDPT 2018", "Mức 4/4 (Chuyên gia)", "Mức 3/4 (Thành thạo)", "Hồ sơ dạy học, bài giảng mẫu"),
    ("4", "Năng lực chuyên môn", "2.2 Năng lực đổi mới PPDH và kiểm tra đánh giá theo năng lực", "Mức 4/4 (Chuyên gia)", "Mức 3/4 (Thành thạo)", "Kết quả khảo sát tiết dạy, dự giờ"),
    ("5", "Năng lực chuyên môn", "2.3 Năng lực nghiên cứu khoa học sư phạm & hướng dẫn KHKT", "Mức 3/4 (Thành thạo)", "Mức 3/4 (Thành thạo)", "Sản phẩm SKKN, dự án KHKT"),
    ("6", "Năng lực quản lý điều hành", "3.1 Lập kế hoạch giáo dục tổ chuyên môn chiến lược & năm học", "Mức 4/4 (Xuất sắc)", "Mức 3/4 (Thành thạo)", "Kế hoạch giáo dục tổ được duyệt"),
    ("7", "Năng lực quản lý điều hành", "3.2 Phân công, điều phối công việc và kiểm tra giám sát", "Mức 4/4 (Thành thạo)", "Mức 3/4 (Thành thạo)", "Tiến độ và hiệu quả của các thành viên"),
    ("8", "Năng lực quản lý điều hành", "3.3 Động viên, bồi dưỡng phát triển đội ngũ và giải quyết xung đột", "Mức 4/4 (Thành thạo)", "Mức 3/4 (Thành thạo)", "Đoàn kết nội bộ, số GV đạt GVDG"),
    ("9", "Năng lực chuyển đổi số & AI", "4.1 Ứng dụng CNTT, phần mềm quản lý, ký số, học bạ điện tử", "Mức 4/4 (Thành thạo)", "Mức 4/4 (Thành thạo)", "Sổ sách điện tử 100% đúng hạn"),
    ("10", "Năng lực chuyển đổi số & AI", "4.2 Ứng dụng công nghệ AI, phần mềm dạy học vào bài giảng số", "Mức 3/4 (Thành thạo)", "Mức 3/4 (Thành thạo)", "Kho học liệu số, ngân hàng đề thi"),
]

for idx, rdata in enumerate(knl_data, 6):
    ws4.cell(row=idx, column=1, value=rdata[0]).alignment = ALIGN_CENTER
    ws4.cell(row=idx, column=2, value=rdata[1]).font = FONT_BOLD
    ws4.cell(row=idx, column=2).alignment = ALIGN_LEFT
    ws4.cell(row=idx, column=3, value=rdata[2]).font = FONT_REGULAR
    ws4.cell(row=idx, column=3).alignment = ALIGN_LEFT
    ws4.cell(row=idx, column=4, value=rdata[3]).font = FONT_REGULAR
    ws4.cell(row=idx, column=4).alignment = ALIGN_CENTER
    ws4.cell(row=idx, column=5, value=rdata[4]).font = FONT_REGULAR
    ws4.cell(row=idx, column=5).alignment = ALIGN_CENTER
    ws4.cell(row=idx, column=6, value=rdata[5]).font = FONT_REGULAR
    ws4.cell(row=idx, column=6).alignment = ALIGN_LEFT
    for col_c in range(1, 7):
        ws4.cell(row=idx, column=col_c).border = BORDER_CELL

# Table 2: KPI System
r_kpi = len(knl_data) + 8
ws4.merge_cells(f"A{r_kpi}:F{r_kpi}")
ws4[f"A{r_kpi}"] = "II. BỘ CHỈ SỐ ĐÁNH GIÁ KẾT QUẢ ĐẦU RA (KPI) VIÊN CHỨC QUẢN LÝ TỔ"
apply_row_styles(ws4, r_kpi, font=FONT_SECTION, fill=SECTION_FILL)

headers_kpi = ["STT", "Nhóm chỉ số KPI", "Chỉ số đo lường hiệu quả cụ thể", "Đơn vị tính", "Chỉ tiêu năm học", "Trọng số (%)"]
for col_i, h in enumerate(headers_kpi, 1):
    cell = ws4.cell(row=r_kpi+1, column=col_i, value=h)
    cell.font = FONT_HEADER_WHITE
    cell.fill = BLUE_HEADER
    cell.alignment = ALIGN_CENTER
    cell.border = BORDER_HEADER

kpi_data = [
    ("KPI-01", "Chất lượng kế hoạch giáo dục", "Xây dựng Kế hoạch tổ và kế hoạch môn học được BGH duyệt đúng hạn", "Thời gian & Chất lượng", "Hoàn thành trước 25/8; Đạt chuẩn", "15%"),
    ("KPI-02", "Duy trì nền nếp sinh hoạt CM", "Tổ chức sinh hoạt chuyên môn theo nghiên cứu bài học", "Số lần / học kỳ", ">= 2 lần/tháng; >= 2 chuyên đề NCBH/kỳ", "15%"),
    ("KPI-03", "Chất lượng giáo dục đại trà", "Tỷ lệ học sinh đạt yêu cầu học tập môn học toàn trường", "Tỷ lệ %", ">= 95% đạt yêu cầu trở lên", "20%"),
    ("KPI-04", "Kết quả học sinh giỏi mũi nhọn", "Số lượng giải HSG cấp tỉnh và dự án KHKT đoạt giải", "Số giải", "Đạt hoặc vượt chỉ tiêu giao đầu năm", "15%"),
    ("KPI-05", "Định mức giảng dạy cá nhân", "Số tiết dạy trực tiếp trên lớp theo định mức đã giảm trừ", "Tiết dạy", "100% định mức (TTCM: 14 tiết; TPCM: 16 tiết)", "15%"),
    ("KPI-06", "Kiểm tra, dự giờ bồi dưỡng GV", "Tỷ lệ giáo viên trong tổ được dự giờ và kiểm tra hồ sơ", "Tỷ lệ %", "100% GV được kiểm tra đúng quy định", "10%"),
    ("KPI-07", "Chuyển đổi số & Học liệu số", "Tỷ lệ hồ sơ số hóa, sổ điểm điện tử ký duyệt đúng hạn; số học liệu số", "Tỷ lệ % / Sản phẩm", "100% đúng hạn; >= 20 học liệu/học kỳ", "10%"),
]

for idx, rdata in enumerate(kpi_data, r_kpi+2):
    ws4.cell(row=idx, column=1, value=rdata[0]).alignment = ALIGN_CENTER
    ws4.cell(row=idx, column=2, value=rdata[1]).font = FONT_BOLD
    ws4.cell(row=idx, column=2).alignment = ALIGN_LEFT
    ws4.cell(row=idx, column=3, value=rdata[2]).font = FONT_REGULAR
    ws4.cell(row=idx, column=3).alignment = ALIGN_LEFT
    ws4.cell(row=idx, column=4, value=rdata[3]).font = FONT_REGULAR
    ws4.cell(row=idx, column=4).alignment = ALIGN_CENTER
    ws4.cell(row=idx, column=5, value=rdata[4]).font = FONT_REGULAR
    ws4.cell(row=idx, column=5).alignment = ALIGN_CENTER
    ws4.cell(row=idx, column=6, value=rdata[5]).font = FONT_BOLD
    ws4.cell(row=idx, column=6).alignment = ALIGN_CENTER
    for col_c in range(1, 7):
        ws4.cell(row=idx, column=col_c).border = BORDER_CELL

# Enable grid lines for all sheets
for s in [ws1, ws2, ws3, ws4]:
    s.views.sheetView[0].showGridLines = True

# Save workbook
excel_path = "d:\\Du-an-web\\web-vtvl-ttcm-tpcm-phuc-hoa\\Bang-mo-ta-VTVL-TTCM-TPCM-THPT-Phuc-Hoa.xlsx"
wb.save(excel_path)
print(f"XLSX created successfully at: {excel_path}")
