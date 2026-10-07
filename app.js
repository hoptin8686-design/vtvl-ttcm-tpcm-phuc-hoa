/* ==============================================================================
   app.js – Đề Án Vị Trí Việc Làm Viên Chức – THPT Phục Hòa
   Căn cứ Nghị định số 232/2026/NĐ-CP & Phụ lục I Danh mục VTVL của Trường
   Hỗ trợ tương tác 35 Biên chế & BẢN MÔ TẢ CHUẨN MẪU NGHỊ ĐỊNH 232 (PHẦN B PHỤ LỤC IV/VI)
   ============================================================================== */

// DỮ LIỆU BẢN MÔ TẢ CHUẨN CÔNG VỤ NGHỊ ĐỊNH 232/2026/NĐ-CP CHO 23 CHỨC DANH
const POSITIONS_DATA = {
  // 13 MÔN HỌC CHUYÊN MÔN
  gv_van: {
    ten: "Giáo viên trung học phổ thông môn Ngữ văn",
    nhom: "Chuyên môn, nghiệp vụ",
    ma: "GV-THPT-01",
    bac: "Bậc 3 đến Bậc 4",
    linhVuc: "Giáo dục THPT – Giảng dạy và giáo dục môn Ngữ văn (CT GDPT 2018)",
    soLuong: "03 người",
    dinhMuc: "17 tiết/tuần",
    capQuanLy: "Tổ trưởng Chuyên môn Ngữ văn, Ban Giám hiệu trường THPT Phục Hòa",
    viTriLienQuan: "Tổ trưởng CM, Giáo viên chủ nhiệm, Thư viện viên, Giáo vụ",
    mucTieu: "Trực tiếp giảng dạy và giáo dục môn Ngữ văn cho học sinh THPT nhằm phát triển toàn diện năng lực ngôn ngữ, văn học và bồi dưỡng phẩm chất nhân văn; nâng cao tỷ lệ tốt nghiệp THPT và bồi dưỡng học sinh giỏi môn Ngữ văn cho trường THPT Phục Hòa.",
    congViec: [
      {
        bac: "Bậc 3",
        noiDung: "1. Giảng dạy môn Ngữ văn 10, 11, 12 theo CT GDPT 2018 (định mức 17 tiết/tuần).\n2. Xây dựng Kế hoạch bài dạy (giáo án), đổi mới PPDH tích cực, rèn 4 kỹ năng Đọc - Viết - Nói - Nghe.\n3. Kiểm tra đánh giá học sinh thường xuyên, định kỳ, cập nhật điểm số điện tử đúng hạn.\n4. Thực hiện công tác giáo viên chủ nhiệm, phối hợp chặt chẽ với cha mẹ học sinh.\n5. Bồi dưỡng học sinh có nguy cơ chưa đạt chuẩn môn học, phụ đạo ôn thi tốt nghiệp.",
        sanPham: "- 100% tiết dạy có kế hoạch bài dạy chuẩn CT 2018.\n- Sổ điểm điện tử, học bạ số ký nhận xét đúng hạn.\n- Tỷ lệ HS đạt chuẩn môn Ngữ văn ≥ 98%.\n- Học sinh tích cực tham gia các diễn đàn văn học.",
        tieuChi: "Chất lượng giờ dạy; tính chuẩn mực sư phạm; sự tiến bộ của học sinh; đúng tiến độ chương trình."
      },
      {
        bac: "Bậc 4",
        noiDung: "1. Thực hiện giảng dạy chuyên đề học tập Ngữ văn nâng cao, ôn thi tốt nghiệp THPT khối 12.\n2. Chủ trì hoặc nòng cốt sinh hoạt chuyên môn theo NCBL cấp tổ và cấp trường.\n3. Xây dựng ma trận đề, ngân hàng câu hỏi đề thi kiểm tra định kỳ môn Văn.\n4. Trực tiếp phát hiện, tuyển chọn và bồi dưỡng đội tuyển HSG Ngữ văn dự thi cấp tỉnh.\n5. Đổi mới phương pháp: Có sáng kiến kinh nghiệm, bài giảng số hoặc ứng dụng AI được nghiệm thu.",
        sanPham: "- Ngân hàng đề thi trắc nghiệm & tự luận chuẩn quy chế.\n- Có học sinh đạt giải HSG môn Ngữ văn cấp tỉnh.\n- Điểm trung bình thi TN môn Văn của trường trong top đầu tỉnh.\n- 01 Chuyên đề đổi mới PPDH cấp trường/năm.",
        tieuChi: "Khả năng dẫn dắt chuyên môn tổ; chất lượng học sinh giỏi; hiệu quả đổi mới phương pháp và ứng dụng AI."
      }
    ],
    khungNangLuc: [
      "1. Năng lực chung: Phẩm chất đạo đức nhà giáo, gương mẫu (Cấp độ 5); Kỷ luật trách nhiệm (Cấp độ 4); Giao tiếp ứng xử sư phạm chuẩn mực (Cấp độ 4); Ứng dụng CNTT, phần mềm học liệu số và AI (Cấp độ 4).",
      "2. Năng lực chuyên môn: Làm chủ chương trình Ngữ văn 2018 (Cấp độ 4-5); Năng lực thẩm mỹ và cảm thụ văn học sâu sắc; Phương pháp dạy học phát triển năng lực (Cấp độ 4); Kỹ năng kiểm tra đánh giá theo ma trận chuẩn (Cấp độ 4)."
    ],
    quanHe: {
      trong: "Báo cáo Tổ trưởng CM, Phó Hiệu trưởng; phối hợp với GV bộ môn, GV chủ nhiệm, cán bộ Thư viện, Thiết bị, Giáo vụ.",
      ngoai: "Quan hệ phối hợp với Ban đại diện Cha mẹ học sinh; Hội Khuyến học địa phương; Sở GD&ĐT khi tham gia chấm thi, tập huấn."
    },
    quyenHan: "Chủ động lựa chọn ngữ liệu ngoài SGK phù hợp chuẩn đầu ra; đánh giá cho điểm và xếp loại học sinh theo quy chế; đề xuất khen thưởng hoặc kỷ luật học sinh; được bảo đảm điều kiện cơ sở vật chất và thời gian nghiên cứu chuyên môn.",
    yeuCau: {
      daoTao: "Bằng Cử nhân (Đại học) sư phạm chuyên ngành Ngữ văn trở lên.",
      boiDuong: "Chứng chỉ bồi dưỡng tiêu chuẩn chức danh nghề nghiệp giáo viên THPT theo quy định.",
      kinhNghiem: "Đã trực tiếp giảng dạy chương trình THPT, có kinh nghiệm ôn thi tốt nghiệp hoặc bồi dưỡng học sinh giỏi.",
      phamChat: "Tâm huyết, mẫu mực, truyền cảm hứng văn học, trách nhiệm và thương yêu học sinh.",
      khac: "Năng lực ngoại ngữ và tin học đạt chuẩn; sử dụng thành thạo phần mềm ký số học bạ và công cụ AI hỗ trợ soạn bài."
    }
  },

  gv_toan: {
    ten: "Giáo viên trung học phổ thông môn Toán",
    nhom: "Chuyên môn, nghiệp vụ",
    ma: "GV-THPT-01",
    bac: "Bậc 3 đến Bậc 4",
    linhVuc: "Giáo dục THPT – Giảng dạy và giáo dục môn Toán (CT GDPT 2018)",
    soLuong: "03 người",
    dinhMuc: "17 tiết/tuần",
    capQuanLy: "Tổ trưởng Chuyên môn Toán, Ban Giám hiệu trường THPT Phục Hòa",
    viTriLienQuan: "Tổ trưởng CM, Giáo viên bộ môn KHTN, Giáo vụ, Thiết bị thí nghiệm",
    mucTieu: "Trực tiếp giảng dạy và giáo dục môn Toán học nhằm phát triển tư duy logic, mô hình hóa toán học và kỹ năng giải quyết bài toán thực tế cho học sinh; bảo đảm kết quả thi tốt nghiệp THPT và bồi dưỡng HSG môn Toán cho trường THPT Phục Hòa.",
    congViec: [
      {
        bac: "Bậc 3",
        noiDung: "1. Giảng dạy môn Toán khối 10, 11, 12 (Đại số, Hình học, Xác suất thống kê) đúng tiến độ 17 tiết/tuần.\n2. Thiết kế kế hoạch bài dạy có tích hợp thực tiễn, giáo dục STEM và ứng dụng GeoGebra.\n3. Kiểm tra đánh giá học sinh theo đúng quy chế của Bộ GD&ĐT, cập nhật điểm số điện tử.\n4. Thực hiện công tác chủ nhiệm lớp, theo dõi nề nếp và sự tiến bộ học tập môn Toán của học sinh.\n5. Phụ đạo, bồi dưỡng học sinh yếu kém môn Toán có nguy cơ hổng kiến thức.",
        sanPham: "- Kế hoạch bài dạy Toán chuẩn CT 2018.\n- Sổ điểm điện tử môn Toán chính xác 100%.\n- Tỷ lệ học sinh đạt chuẩn môn Toán ≥ 96%.\n- Học sinh ứng dụng tốt máy tính Casio và phần mềm hình học.",
        tieuChi: "Chất lượng giờ dạy; tính khoa học và chính xác logic; sự chuyển biến của học sinh yếu kém."
      },
      {
        bac: "Bậc 4",
        noiDung: "1. Giảng dạy chuyên đề Toán nâng cao và trực tiếp ôn thi tốt nghiệp THPT khối 12.\n2. Chủ trì xây dựng ma trận đề kiểm tra định kỳ, ngân hàng đề thi trắc nghiệm theo 4 mức độ nhận thức.\n3. Nòng cốt bồi dưỡng đội tuyển HSG Toán dự thi cấp tỉnh.\n4. Tổ chức chuyên đề sinh hoạt chuyên môn theo NCBL, tích hợp bài học STEM cấp trường.\n5. Ứng dụng AI sinh câu hỏi trắc nghiệm Toán và sơ đồ hóa bài giảng điện tử.",
        sanPham: "- Ngân hàng đề thi thử tốt nghiệp THPT môn Toán chuẩn form đề minh họa Bộ GD&ĐT.\n- Có học sinh đạt giải HSG môn Toán cấp tỉnh.\n- Điểm trung bình thi TN môn Toán nằm trong top trường vùng cao.\n- Có sáng kiến kinh nghiệm cấp cơ sở được nghiệm thu.",
        tieuChi: "Hiệu quả ôn thi TN THPT; năng lực hướng dẫn đội tuyển HSG; khả năng đổi mới sáng tạo trong phương pháp dạy Toán."
      }
    ],
    khungNangLuc: [
      "1. Năng lực chung: Gương mẫu, kỷ luật (Cấp độ 5); Ứng dụng CNTT, phần mềm GeoGebra và AI sinh đề thi (Cấp độ 4); Giao tiếp ứng xử sư phạm (Cấp độ 4).",
      "2. Năng lực chuyên môn: Nắm vững kiến thức Toán THPT chuyên sâu (Cấp độ 4-5); Tư duy mô hình hóa toán học; Phương pháp giáo dục STEM liên môn (Cấp độ 4); Kiểm tra đánh giá trắc nghiệm chuẩn mực (Cấp độ 4)."
    ],
    quanHe: {
      trong: "Báo cáo Tổ trưởng CM, Ban Giám hiệu; phối hợp với GV Tin học, GV Lý, Hóa, Sinh, Giáo vụ.",
      ngoai: "Ban đại diện CMHS; giáo viên cốt cán môn Toán các trường trong cụm chuyên môn tỉnh Cao Bằng."
    },
    quyenHan: "Đánh giá, xếp loại kết quả học tập môn Toán của học sinh; đề xuất các hình thức khen thưởng, bồi dưỡng; lựa chọn phần mềm dạy học hỗ trợ.",
    yeuCau: {
      daoTao: "Bằng Cử nhân (Đại học) sư phạm chuyên ngành Toán học trở lên.",
      boiDuong: "Chứng chỉ tiêu chuẩn chức danh nghề nghiệp GV THPT.",
      kinhNghiem: "Có kinh nghiệm giảng dạy môn Toán CT 2018, bồi dưỡng HSG hoặc luyện thi TN THPT.",
      phamChat: "Cẩn trọng, logic, nhiệt huyết, tâm lý và kiên trì với học sinh.",
      khac: "Sử dụng thành thạo GeoGebra, Mathtype, máy tính cầm tay, công cụ AI tạo đề thi trắc nghiệm."
    }
  },

  gv_anh: {
    ten: "Giáo viên trung học phổ thông môn Ngoại ngữ (Tiếng Anh)",
    nhom: "Chuyên môn, nghiệp vụ",
    ma: "GV-THPT-01",
    bac: "Bậc 3 đến Bậc 4",
    linhVuc: "Giáo dục THPT – Giảng dạy Tiếng Anh hệ 10 năm theo CT 2018",
    soLuong: "02 người",
    dinhMuc: "17 tiết/tuần",
    capQuanLy: "Tổ trưởng Chuyên môn Ngoại ngữ, Ban Giám hiệu trường THPT Phục Hòa",
    viTriLienQuan: "Tổ trưởng CM, Giáo viên chủ nhiệm, Quản trị mạng phòng học ngoại ngữ, Giáo vụ",
    mucTieu: "Trực tiếp giảng dạy và phát triển năng lực giao tiếp tiếng Anh (Nghe - Nói - Đọc - Viết) cho học sinh THPT; nâng cao chất lượng thi tốt nghiệp THPT môn Ngoại ngữ và phát triển môi trường học tập ngoại ngữ trong nhà trường.",
    congViec: [
      {
        bac: "Bậc 3",
        noiDung: "1. Giảng dạy môn Tiếng Anh 10, 11, 12 theo CT 2018 hệ 10 năm (17 tiết/tuần).\n2. Xây dựng môi trường giao tiếp bằng tiếng Anh trên lớp; rèn phát âm chuẩn và ngữ pháp cơ bản.\n3. Khai thác phòng học Ngoại ngữ tương tác đa phương tiện và app AI luyện nói cho học sinh.\n4. Thực hiện kiểm tra đánh giá định kỳ 4 kỹ năng ngôn ngữ; cập nhật điểm số điện tử.\n5. Tổ chức phụ đạo học sinh yếu môn Tiếng Anh, xóa bỏ tâm lý ngại giao tiếp ngoại ngữ.",
        sanPham: "- 100% tiết dạy có bài giảng điện tử sinh động, phát âm chuẩn.\n- Tỷ lệ học sinh đạt chuẩn môn Tiếng Anh tăng đều qua từng kỳ.\n- Học sinh tích cực tương tác hội thoại cơ bản.",
        tieuChi: "Phương pháp dạy học tích cực; khả năng lôi cuốn học sinh; sự tiến bộ kỹ năng nghe nói của học sinh vùng cao."
      },
      {
        bac: "Bậc 4",
        noiDung: "1. Dạy chuyên đề ngữ pháp nâng cao, đọc hiểu và luyện thi tốt nghiệp THPT môn Tiếng Anh.\n2. Phụ trách Câu lạc bộ Tiếng Anh (English Club) nhà trường, tổ chức Ngày hội ngoại ngữ.\n3. Bồi dưỡng đội tuyển HSG Tiếng Anh tham gia kỳ thi chọn HSG cấp tỉnh Cao Bằng.\n4. Xây dựng ngân hàng câu hỏi đề thi trắc nghiệm Tiếng Anh chuẩn format đề thi của Bộ GD&ĐT.\n5. Nghiên cứu giải pháp nâng cao phổ điểm thi tốt nghiệp THPT môn Tiếng Anh.",
        sanPham: "- Ngân hàng đề thi tốt nghiệp THPT môn Tiếng Anh có giải thích chi tiết.\n- Có học sinh đạt giải HSG Tiếng Anh cấp tỉnh.\n- Tỷ lệ điểm trung bình thi TN môn Anh tăng ≥ 5% so với năm trước.\n- Hoạt động CLB Tiếng Anh duy trì đều đặn hàng tháng.",
        tieuChi: "Chất lượng ôn thi TN; kết quả giải HSG cấp tỉnh; năng lực lan tỏa phong trào học ngoại ngữ toàn trường."
      }
    ],
    khungNangLuc: [
      "1. Năng lực chung: Phẩm chất đạo đức nhà giáo (Cấp độ 5); Năng lực số, ứng dụng app AI hội thoại (Cấp độ 4); Giao tiếp tương tác linh hoạt (Cấp độ 4).",
      "2. Năng lực chuyên môn: Năng lực ngôn ngữ Tiếng Anh chuẩn bậc 5 theo KNLNN Việt Nam (C1); Phương pháp giảng dạy ngôn ngữ giao tiếp tích cực (Cấp độ 4-5); Thiết kế đề kiểm tra trắc nghiệm chuẩn ma trận (Cấp độ 4)."
    ],
    quanHe: {
      trong: "Báo cáo Tổ trưởng CM, Ban Giám hiệu; phối hợp với Đoàn Thanh niên, GV chủ nhiệm, GV Tin học.",
      ngoai: "Ban đại diện CMHS; giáo viên ngoại ngữ các trường THPT trong cụm chuyên môn."
    },
    quyenHan: "Chủ động lựa chọn học liệu mở và video luyện nghe quốc tế phù hợp; đánh giá năng lực ngôn ngữ của học sinh theo thang chuẩn; phụ trách CLB Ngoại ngữ trường.",
    yeuCau: {
      daoTao: "Bằng Cử nhân Sư phạm Tiếng Anh hoặc Cử nhân Ngôn ngữ Anh kèm chứng chỉ nghiệp vụ sư phạm.",
      boiDuong: "Chứng chỉ chức danh nghề nghiệp GV THPT; chứng chỉ năng lực tiếng Anh Bậc 5 (C1).",
      kinhNghiem: "Có kinh nghiệm giảng dạy tiếng Anh THPT và ôn luyện thi tốt nghiệp.",
      phamChat: "Năng động, nhiệt huyết, phát âm chuẩn, kiên nhẫn và sáng tạo.",
      khac: "Thành thạo các phần mềm học tiếng Anh, app AI luyện phát âm (ELSA, Duolingo, ChatGPT)."
    }
  },

  // 06 VIÊN CHỨC HỖ TRỢ
  ht_kt: {
    ten: "Kế toán / Phụ trách kế toán",
    nhom: "Hỗ trợ, phục vụ",
    ma: "KT-THPT-01",
    bac: "Bậc 1 đến Bậc 4",
    linhVuc: "Quản lý tài chính, kế toán ngân sách nhà nước và tài sản công trường học",
    soLuong: "01 người",
    dinhMuc: "40 giờ/tuần",
    capQuanLy: "Hiệu trưởng trường THPT Phục Hòa (Chủ tài khoản)",
    viTriLienQuan: "Thủ quỹ, Văn thư, Hiệu trưởng, Công đoàn, Tổ chuyên môn",
    mucTieu: "Thực hiện toàn diện công tác tài chính, kế toán và quản lý ngân sách nhà nước nhằm bảo đảm quản lý thu chi minh bạch, đúng chế độ chính sách và phục vụ kịp thời các hoạt động dạy học của trường THPT Phục Hòa.",
    congViec: [
      {
        bac: "Bậc 2",
        noiDung: "1. Thực hiện các nghiệp vụ ghi chép sổ sách kế toán hàng ngày trên phần mềm MISA Mimosa.\n2. Tập hợp, kiểm tra tính hợp pháp, hợp lệ của chứng từ thanh toán thu - chi.\n3. Lập hồ sơ chứng từ thanh toán chế độ thường xuyên, thanh toán tiền điện, nước, văn phòng phẩm qua Kho bạc.\n4. Thực hiện các thủ tục đóng bảo hiểm xã hội, bảo hiểm y tế, công đoàn cho cán bộ giáo viên đúng hạn.",
        sanPham: "- Hệ thống sổ kế toán cập nhật đầy đủ, ngăn nắp.\n- Chứng từ kế toán được kiểm tra 100% hợp lệ trước khi trình ký.\n- Nộp tiền BHXH đúng kỳ hạn không để phát sinh nợ đọng.",
        tieuChi: "Tính chính xác, kịp thời; tuân thủ đúng quy trình kế toán công; không để xảy ra sai sót số liệu."
      },
      {
        bac: "Bậc 3",
        noiDung: "1. Xây dựng dự toán thu - chi ngân sách nhà nước đầu năm học và dự toán các nguồn thu hợp pháp khác.\n2. Tính toán lập bảng thanh toán tiền lương, phụ cấp chức vụ, thâm niên, tăng lương và chi trả thừa giờ đúng hạn.\n3. Thực hiện giao dịch điện tử và đối chiếu khớp số liệu hàng tháng, quý với Kho bạc Nhà nước Phục Hòa.\n4. Lập báo cáo quyết toán tài chính quý và năm học gửi Sở GD&ĐT và cơ quan tài chính kiểm toán.",
        sanPham: "- Dự toán ngân sách năm được cấp có thẩm quyền phê duyệt.\n- Bảng lương chuyển khoản đến tài khoản giáo viên trước ngày 05 hàng tháng.\n- Báo cáo tài chính quý/năm được Sở GD&ĐT thẩm định đạt chuẩn 100%.\n- Biên bản đối chiếu số dư Kho bạc khớp 100%.",
        tieuChi: "Chính xác tuyệt đối; đúng Luật Ngân sách và Luật Kế toán; hoàn thành đúng thời hạn quy định."
      }
    ],
    khungNangLuc: [
      "1. Năng lực chung: Liêm chính tuyệt đối, đạo đức công vụ trong sáng (Cấp độ 5); Kỷ luật cao (Cấp độ 4); Bảo mật dữ liệu tài chính (Cấp độ 4); Thành thạo phần mềm kế toán số (Cấp độ 4).",
      "2. Năng lực chuyên môn: Nắm vững Luật Kế toán, Luật Ngân sách, chế độ tài chính trường học (Cấp độ 4); Nghiệp vụ kho bạc điện tử và thanh toán số; Năng lực phân tích và lập kế hoạch tài chính (Cấp độ 3-4)."
    ],
    quanHe: {
      trong: "Chịu sự chỉ đạo trực tiếp của Hiệu trưởng; phối hợp với Thủ quỹ, Văn thư, Công đoàn, Tổ chuyên môn.",
      ngoai: "Kho bạc Nhà nước, Cơ quan Thuế, Phòng Kế hoạch - Tài chính Sở GD&ĐT, Bảo hiểm xã hội."
    },
    quyenHan: "Độc lập về chuyên môn nghiệp vụ kế toán; có quyền từ chối thanh toán các khoản chi sai nguyên tắc, không đúng dự toán hoặc chứng từ không hợp lệ; yêu cầu các bộ phận cung cấp đầy đủ chứng từ khi thanh toán.",
    yeuCau: {
      daoTao: "Bằng Đại học chuyên ngành Tài chính - Kế toán hoặc Kiểm toán trở lên.",
      boiDuong: "Chứng chỉ bồi dưỡng kế toán viên hành chính sự nghiệp / Kế toán trưởng.",
      kinhNghiem: "Có kinh nghiệm làm kế toán đơn vị sự nghiệp giáo dục công lập.",
      phamChat: "Trung thực, cẩn trọng tuyệt đối, khách quan, liêm chính.",
      khac: "Thành thạo phần mềm MISA Mimosa Online, Dịch vụ công Kho bạc Nhà nước, chữ ký số token."
    }
  },

  ht_vt: {
    ten: "Văn thư trường học",
    nhom: "Hỗ trợ, phục vụ",
    ma: "VT-THPT-01",
    bac: "Bậc 1 đến Bậc 4",
    linhVuc: "Quản lý văn bản, con dấu, công tác lưu trữ và hành chính trường học",
    soLuong: "01 người",
    dinhMuc: "40 giờ/tuần",
    capQuanLy: "Hiệu trưởng, Phó Hiệu trưởng phụ trách hành chính văn phòng",
    viTriLienQuan: "Hiệu trưởng, Ban Giám hiệu, Kế toán, Giáo vụ, Tổ chuyên môn",
    mucTieu: "Tiếp nhận, phát hành, quản lý hệ thống văn bản đi - đến và con dấu của nhà trường theo đúng quy định pháp luật; lập hồ sơ lưu trữ khoa học nhằm bảo đảm tính thông suốt, bảo mật và tra cứu kịp thời cho trường THPT Phục Hòa.",
    congViec: [
      {
        bac: "Bậc 2",
        noiDung: "1. Tiếp nhận, đăng ký văn bản đi - đến trên hệ thống Quản lý văn bản và điều hành điện tử (iOffice).\n2. Chuyển giao công văn đến đúng người có thẩm quyền xử lý trong ngày làm việc.\n3. Quản lý, đóng dấu cơ quan và ký số văn bản phát hành theo đúng thẩm quyền và thể thức.\n4. Đánh máy, in sao tài liệu, giấy triệu tập, giấy giới thiệu phục vụ công tác điều hành của Ban Giám hiệu.\n5. Sắp xếp, bảo quản hồ sơ tài liệu lưu trữ hiện hành đúng quy chuẩn.",
        sanPham: "- 100% công văn đến được vào sổ điện tử và xử lý trong ngày.\n- Quản lý con dấu an toàn tuyệt đối, không thất lạc.\n- Văn bản phát hành đúng thể thức Nghị định 30/2020/NĐ-CP.\n- Hồ sơ lưu trữ được số hóa khoa học.",
        tieuChi: "Tính kịp thời; độ chính xác; tuân thủ nghiêm ngặt quy chế quản lý con dấu và văn bản."
      }
    ],
    khungNangLuc: [
      "1. Năng lực chung: Cẩn thận, trách nhiệm (Cấp độ 4); Giữ gìn bí mật công vụ (Cấp độ 5); Kỹ năng giao tiếp hành chính chuẩn mực (Cấp độ 3-4); Sử dụng thành thạo máy tính và phần mềm số (Cấp độ 3).",
      "2. Năng lực chuyên môn: Nghiệp vụ văn thư lưu trữ, thể thức văn bản hành chính theo NĐ 30/2020 (Cấp độ 3-4); Quản lý và khai thác hệ thống iOffice, chữ ký số tổ chức (Cấp độ 3-4)."
    ],
    quanHe: {
      trong: "Chịu sự phân công của Hiệu trưởng; phối hợp chặt chẽ với Kế toán, Giáo vụ, các Tổ chuyên môn.",
      ngoai: "Văn phòng Sở GD&ĐT, Bưu điện, UBND huyện, các cơ quan ban ngành gửi văn bản đến."
    },
    quyenHan: "Được quyền từ chối đóng dấu vào các văn bản không đúng thể thức hoặc không có chữ ký của người có thẩm quyền; quản lý trực tiếp con dấu và chứng thư số của nhà trường.",
    yeuCau: {
      daoTao: "Tốt nghiệp Trung cấp trở lên chuyên ngành Văn thư - Lưu trữ hoặc ngành Luật, Hành chính.",
      boiDuong: "Chứng chỉ bồi dưỡng nghiệp vụ văn thư lưu trữ; chứng chỉ tin học ứng dụng.",
      kinhNghiem: "Có kinh nghiệm quản lý văn bản và con dấu trong cơ quan nhà nước.",
      phamChat: "Cẩn thận, bảo mật, ngăn nắp, trung thực và tận tụy.",
      khac: "Thành thạo phần mềm iOffice, máy scan, máy in, phần mềm số hóa hồ sơ."
    }
  },

  ht_tbtn: {
    ten: "Thiết bị, thí nghiệm",
    nhom: "Hỗ trợ, phục vụ",
    ma: "TBTN-THPT-01",
    bac: "Bậc 2 đến Bậc 3",
    linhVuc: "Quản lý thiết bị dạy học, hóa chất, máy móc và phòng thực hành bộ môn",
    soLuong: "01 người",
    dinhMuc: "40 giờ/tuần",
    capQuanLy: "Phó Hiệu trưởng phụ trách CSVC, Hiệu trưởng",
    viTriLienQuan: "Giáo viên bộ môn Vật lý, Hóa học, Sinh học, Công nghệ, Kế toán",
    mucTieu: "Quản lý, bảo quản an toàn toàn bộ trang thiết bị dạy học, hóa chất và phòng thí nghiệm; chuẩn bị đầy đủ dụng cụ phục vụ 100% tiết thực hành môn học theo kế hoạch giáo dục của trường THPT Phục Hòa.",
    congViec: [
      {
        bac: "Bậc 2",
        noiDung: "1. Quản lý, bảo quản toàn bộ thiết bị dạy học, máy móc thí nghiệm và hóa chất phòng bộ môn Lý, Hóa, Sinh, Công nghệ.\n2. Căn cứ lịch báo giảng hàng tuần, chuẩn bị trước thiết bị, mẫu vật, hóa chất cho tiết dạy thực hành của giáo viên.\n3. Hỗ trợ giáo viên trong giờ thực hành; theo dõi học sinh thực hiện đúng quy tắc an toàn phòng thí nghiệm.\n4. Lập sổ theo dõi mượn - trả thiết bị dạy học; ghi chép sổ nhật ký phòng thực hành đầy đủ.\n5. Thực hiện kiểm kê định kỳ tài sản thiết bị cuối kỳ/năm; lập biên bản thanh lý thiết bị hỏng.",
        sanPham: "- 100% tiết thực hành có đầy đủ thiết bị, hóa chất chuẩn bị sẵn sàng.\n- Sổ mượn trả TBDH được ghi chép cập nhật hàng tuần.\n- Không để xảy ra mất mát, hỏng hóc thiết bị do thiếu trách nhiệm.\n- Bảo đảm an toàn tuyệt đối PCCC và không xảy ra sự cố hóa chất.",
        tieuChi: "Tính chu đáo; chuẩn bị đúng hẹn; bảo đảm tuyệt đối an toàn phòng thí nghiệm và phòng chống cháy nổ."
      }
    ],
    khungNangLuc: [
      "1. Năng lực chung: Cẩn thận, trách nhiệm (Cấp độ 4); Tuân thủ quy định an toàn lao động (Cấp độ 5); Phối hợp hỗ trợ đồng nghiệp chu đáo (Cấp độ 4).",
      "2. Năng lực chuyên môn: Nắm vững quy trình vận hành thiết bị thí nghiệm Lý - Hóa - Sinh (Cấp độ 3-4); Quy chuẩn bảo quản hóa chất độc hại và an toàn PCCC (Cấp độ 4); Quản lý sổ sách thiết bị (Cấp độ 3)."
    ],
    quanHe: {
      trong: "Báo cáo Phó Hiệu trưởng; phối hợp trực tiếp với GV bộ môn KHTN, Kế toán, Lao công.",
      ngoai: "Các đơn vị cung ứng thiết bị giáo dục, cơ quan kiểm định an toàn phòng thí nghiệm."
    },
    quyenHan: "Được quyền từ chối cho mượn thiết bị nếu giáo viên/học sinh không tuân thủ quy tắc an toàn; lập biên bản xử lý học sinh làm vỡ hỏng dụng cụ do thiếu ý thức; đề xuất mua sắm, sửa chữa thiết bị.",
    yeuCau: {
      daoTao: "Tốt nghiệp Cao đẳng trở lên chuyên ngành Thiết bị dạy học hoặc Sư phạm Lý, Hóa, Sinh, Công nghệ.",
      boiDuong: "Chứng chỉ nghiệp vụ thiết bị thí nghiệm; chứng chỉ an toàn hóa chất và PCCC.",
      kinhNghiem: "Có kinh nghiệm quản lý phòng thực hành thí nghiệm trường học.",
      phamChat: "Ngăn nắp, cẩn trọng, kỷ luật, có tinh thần trách nhiệm cao.",
      khac: "Kỹ năng sửa chữa bảo trì đơn giản các mô hình học cụ, máy chiếu, bảng tương tác."
    }
  },

  ht_gvu: {
    ten: "Giáo vụ trường học",
    nhom: "Hỗ trợ, phục vụ",
    ma: "GVU-THPT-01",
    bac: "Bậc 2 đến Bậc 3",
    linhVuc: "Quản lý hồ sơ học sinh, sổ đăng bộ, tuyển sinh, thi tốt nghiệp và văn bằng",
    soLuong: "01 người",
    dinhMuc: "40 giờ/tuần",
    capQuanLy: "Phó Hiệu trưởng phụ trách chuyên môn, Hiệu trưởng",
    viTriLienQuan: "Giáo viên chủ nhiệm, Quản trị CSDL ngành, Ban Giám hiệu, Văn thư",
    mucTieu: "Quản lý chính xác hệ thống sổ đăng bộ, hồ sơ tuyển sinh lớp 10, hồ sơ thi tốt nghiệp THPT, cơ sở dữ liệu học sinh và công tác cấp phát văn bằng chứng chỉ cho học sinh trường THPT Phục Hòa.",
    congViec: [
      {
        bac: "Bậc 2",
        noiDung: "1. Quản lý, ghi chép và bảo quản sổ đăng bộ của nhà trường theo đúng quy chế.\n2. Thu nhận, kiểm tra và hoàn thiện hồ sơ trúng tuyển học sinh lớp 10; hồ sơ học sinh chuyển trường.\n3. Quản lý và kiểm tra dữ liệu học bạ số, sổ điểm điện tử của các lớp; đối chiếu thông tin cá nhân học sinh.\n4. Lập hồ sơ đăng ký dự thi tốt nghiệp THPT cho học sinh khối 12; in ấn phiếu đăng ký và thẻ dự thi.\n5. Quản lý sổ cấp phát bằng tốt nghiệp THPT; thực hiện cấp phát bằng cho học sinh đúng quy định.",
        sanPham: "- Sổ đăng bộ cập nhật đầy đủ, chính xác 100% học sinh toàn trường.\n- Hồ sơ tuyển sinh 10 và thi tốt nghiệp nộp về Sở GD&ĐT đúng thời hạn.\n- Dữ liệu học bạ số khớp 100% với giấy khai sinh và hồ sơ gốc.\n- Cấp phát bằng tốt nghiệp đúng quy trình, không để xảy ra sai sót.",
        tieuChi: "Chính xác tuyệt đối về thông tin nhân thân; bảo mật dữ liệu; tuân thủ quy chế thi và cấp văn bằng."
      }
    ],
    khungNangLuc: [
      "1. Năng lực chung: Cẩn thận, chi tiết (Cấp độ 4); Trách nhiệm bảo mật thông tin (Cấp độ 4); Phục vụ học sinh và phụ huynh tận tình (Cấp độ 4).",
      "2. Năng lực chuyên môn: Nắm vững quy chế thi tốt nghiệp THPT, quy chế tuyển sinh và quản lý văn bằng (Cấp độ 3-4); Sử dụng thành thạo phần mềm Quản lý thi và CSDL ngành GDĐT (Cấp độ 3-4)."
    ],
    quanHe: {
      trong: "Báo cáo Phó Hiệu trưởng; phối hợp với GV chủ nhiệm các khối, Văn thư, Quản trị CSDL.",
      ngoai: "Phòng Quản lý chất lượng & Khảo thí Sở GD&ĐT Cao Bằng; phụ huynh và cựu học sinh liên hệ rút học bạ, nhận bằng."
    },
    quyenHan: "Kiểm tra tính hợp lệ của hồ sơ học sinh; yêu cầu GV chủ nhiệm bổ sung thông tin học bạ còn thiếu; ký xác nhận vào sổ cấp phát văn bằng theo phân công.",
    yeuCau: {
      daoTao: "Tốt nghiệp Trung cấp trở lên ngành Quản lý giáo dục, Tin học ứng dụng, Hành chính hoặc Văn thư.",
      boiDuong: "Chứng chỉ quản lý CSDL ngành giáo dục; nghiệp vụ giáo vụ trường học.",
      kinhNghiem: "Có kinh nghiệm quản lý hồ sơ sổ sách học sinh THPT.",
      phamChat: "Cẩn thận, trung thực, nhẹ nhàng, chu đáo và kiên nhẫn.",
      khac: "Sử dụng thành thạo Excel, Word, phần mềm VnEdu/SMAS, hệ thống thi Bộ GD&ĐT."
    }
  },

  ht_yt: {
    ten: "Y tế trường học",
    nhom: "Hỗ trợ, phục vụ",
    ma: "YT-THPT-01",
    bac: "Bậc 1 đến Bậc 3",
    linhVuc: "Chăm sóc sức khỏe ban đầu, sơ cấp cứu, phòng chống dịch bệnh và vệ sinh học đường",
    soLuong: "01 người",
    dinhMuc: "40 giờ/tuần",
    capQuanLy: "Phó Hiệu trưởng phụ trách nề nếp, Hiệu trưởng",
    viTriLienQuan: "Giáo viên GDTC, Giáo viên chủ nhiệm, Đoàn Thanh niên, Nhân viên căng tin",
    mucTieu: "Chăm sóc sức khỏe ban đầu, sơ cấp cứu kịp thời các tai nạn thương tích và phòng chống dịch bệnh học đường nhằm bảo đảm môi trường giáo dục an toàn, khỏe mạnh cho toàn thể học sinh và giáo viên trường THPT Phục Hòa.",
    congViec: [
      {
        bac: "Bậc 2",
        noiDung: "1. Thường trực tại phòng y tế; sơ cứu, cấp cứu kịp thời học sinh và cán bộ giáo viên bị tai nạn thương tích hoặc ốm đau đột xuất.\n2. Quản lý tủ thuốc y tế học đường; định kỳ kiểm tra hạn dùng, bổ sung cơ số thuốc thiết yếu và bông băng gạc.\n3. Lập sổ theo dõi sức khỏe học sinh; phối hợp với TTYT huyện tổ chức khám sức khỏe định kỳ cho học sinh toàn trường.\n4. Tuyên truyền giáo dục sức khỏe, phòng chống dịch bệnh theo mùa (cúm, sốt xuất huyết, đau mắt đỏ, nha học đường).\n5. Kiểm tra vệ sinh môi trường học đường, an toàn nguồn nước uống và an toàn thực phẩm căng tin trường học.",
        sanPham: "- 100% ca tai nạn thương tích hoặc đau ốm được sơ cứu kịp thời và ghi nhật ký y tế.\n- Tủ thuốc luôn đầy đủ cơ số cấp cứu ban đầu theo quy định Bộ Y tế.\n- Hồ sơ theo dõi sức khỏe 100% học sinh được hoàn thiện sau đợt khám định kỳ.\n- Không để bùng phát dịch bệnh lây lan hoặc ngộ độc thực phẩm trong trường.",
        tieuChi: "Kịp thời, chính xác trong sơ cứu; tận tình, chu đáo; tuân thủ đúng quy chuẩn chuyên môn y tế."
      }
    ],
    khungNangLuc: [
      "1. Năng lực chung: Y đức, tình thương yêu học sinh (Cấp độ 5); Tinh thần trách nhiệm thường trực (Cấp độ 4); Kỹ năng giao tiếp động viên học sinh (Cấp độ 4).",
      "2. Năng lực chuyên môn: Kỹ năng chẩn đoán ban đầu và sơ cấp cứu tai nạn thương tích học đường (Cấp độ 3-4); Quản lý thuốc và vệ sinh dịch tễ trường học (Cấp độ 3); Nghiệp vụ y tế học đường (Cấp độ 3)."
    ],
    quanHe: {
      trong: "Báo cáo Ban Giám hiệu; phối hợp với GV Thể dục, GV chủ nhiệm, Đoàn Thanh niên.",
      ngoai: "Trung tâm Y tế huyện Phục Hòa, Bệnh viện Đa khoa, Trạm Y tế xã/thị trấn, phụ huynh học sinh."
    },
    quyenHan: "Chủ động xử trí sơ cấp cứu; quyết định chuyển tuyến cấp cứu đối với trường hợp vượt quá khả năng sơ cứu; đình chỉ hoạt động căng tin nếu phát hiện vi phạm nghiêm trọng an toàn thực phẩm.",
    yeuCau: {
      daoTao: "Tốt nghiệp Trung cấp Y sĩ hoặc Cao đẳng Điều dưỡng trở lên.",
      boiDuong: "Chứng chỉ hành nghề y tế; chứng chỉ bồi dưỡng công tác y tế trường học.",
      kinhNghiem: "Có kinh nghiệm thực tế về sơ cấp cứu ban đầu hoặc y tế cơ sở.",
      phamChat: "Nhẹ nhàng, cẩn trọng, tận tụy, có tinh thần y đức.",
      khac: "Kỹ năng tuyên truyền phòng chống dịch bệnh và xử lý rác thải y tế đúng quy chuẩn."
    }
  },

  ht_tq: {
    ten: "Thủ quỹ trường học",
    nhom: "Hỗ trợ, phục vụ",
    ma: "TQ-THPT-01",
    bac: "Bậc 1 đến Bậc 3",
    linhVuc: "Quản lý két sắt, quỹ tiền mặt và thực hiện thu chi ngân quỹ theo quy định",
    soLuong: "01 người",
    dinhMuc: "40 giờ/tuần",
    capQuanLy: "Hiệu trưởng (Chủ tài khoản), Kế toán trưởng",
    viTriLienQuan: "Kế toán, Hiệu trưởng, Văn thư, Công đoàn",
    mucTieu: "Quản lý an toàn tuyệt đối két sắt và quỹ tiền mặt của nhà trường; thực hiện việc thu, chi tiền mặt chính xác, kịp thời theo đúng phiếu thu, phiếu chi hợp lệ có đầy đủ chữ ký theo quy định tại trường THPT Phục Hòa.",
    congViec: [
      {
        bac: "Bậc 2",
        noiDung: "1. Trực tiếp quản lý khóa và mã số két sắt quỹ tiền mặt; bảo đảm an toàn tuyệt đối tiền mặt của cơ quan.\n2. Thực hiện thu tiền mặt theo đúng phiếu thu hợp lệ có đủ chữ ký của Kế toán và người nộp tiền.\n3. Thực hiện chi trả tiền mặt theo đúng phiếu chi hợp lệ đã được Hiệu trưởng duyệt và Kế toán ký.\n4. Mở sổ quỹ tiền mặt; hàng ngày ghi chép đầy đủ các khoản thu, chi và tính số dư tồn quỹ cuối ngày.\n5. Định kỳ và đột xuất cùng Kế toán đối chiếu khớp số dư sổ quỹ tiền mặt với sổ kế toán; lập biên bản kiểm kê quỹ.",
        sanPham: "- Số dư quỹ tiền mặt thực tế khớp 100% với sổ quỹ và sổ kế toán.\n- 100% chứng từ thu - chi có đầy đủ chữ ký hợp lệ trước khi xuất/nhập tiền.\n- Sổ quỹ tiền mặt được ghi chép hàng ngày, rõ ràng, không tẩy xóa.\n- Biên bản kiểm kê quỹ tiền mặt cuối tháng, quý đầy đủ.",
        tieuChi: "Tuyệt đối trung thực, chính xác; không để xảy ra thiếu hụt, chênh lệch hay thất thoát tiền mặt."
      }
    ],
    khungNangLuc: [
      "1. Năng lực chung: Trung thực tuyệt đối, liêm chính (Cấp độ 5); Kỷ luật cao (Cấp độ 4); Cẩn trọng, tỉ mỉ (Cấp độ 5); Bảo mật an toàn ngân quỹ (Cấp độ 5).",
      "2. Năng lực chuyên môn: Nghiệp vụ quản lý quỹ tiền mặt (Cấp độ 3); Kỹ năng kiểm đếm, phân biệt tiền thật/tiền giả (Cấp độ 4); Ghi chép và đối chiếu sổ quỹ theo quy chuẩn kế toán (Cấp độ 3)."
    ],
    quanHe: {
      trong: "Chịu sự kiểm soát trực tiếp của Hiệu trưởng và Kế toán; phối hợp với Văn thư, các cá nhân nhận tiền.",
      ngoai: "Kho bạc Nhà nước Phục Hòa, Ngân hàng khi thực hiện rút tiền mặt về nhập quỹ."
    },
    quyenHan: "Có quyền từ chối chi tiền nếu phiếu chi thiếu chữ ký của Hiệu trưởng hoặc Kế toán, hoặc chứng từ bị sửa chữa, tẩy xóa; yêu cầu người nhận tiền đếm kỹ tiền mặt tại bàn thu ngân trước khi rời đi.",
    yeuCau: {
      daoTao: "Tốt nghiệp Trung cấp trở lên ngành Tài chính - Kế toán, Ngân hàng hoặc Kinh tế.",
      boiDuong: "Chứng chỉ bồi dưỡng nghiệp vụ thủ quỹ, kiểm đếm tiền mặt.",
      kinhNghiem: "Có kinh nghiệm thực tế về quản lý quỹ tiền mặt hoặc kế toán viên.",
      phamChat: "Tuyệt đối trung thực, cẩn trọng, nguyên tắc và bảo mật.",
      khac: "Kỹ năng sử dụng máy đếm tiền, máy soi tiền giả, phần mềm quản lý thu chi."
    }
  },

  // VIÊN CHỨC QUẢN LÝ
  ql_ht: {
    ten: "Hiệu trưởng trường trung học phổ thông",
    nhom: "Viên chức quản lý",
    ma: "HT-THPT-01",
    bac: "Bậc 3 đến Bậc 5",
    linhVuc: "Lãnh đạo, quản lý và điều hành toàn diện trường THPT Phục Hòa",
    soLuong: "01 người",
    dinhMuc: "02 tiết/tuần (Phụ cấp chức vụ: 0.70)",
    capQuanLy: "Giám đốc Sở Giáo dục và Đào tạo tỉnh Cao Bằng, UBND tỉnh",
    viTriLienQuan: "Phó Hiệu trưởng, Tổ trưởng CM, Kế toán, Văn thư, Công đoàn, Đoàn TN",
    mucTieu: "Lãnh đạo, quản lý và điều hành toàn diện mọi hoạt động của trường THPT Phục Hòa theo đúng chủ trương đường lối của Đảng, pháp luật của Nhà nước và Điều lệ trường THPT; chịu trách nhiệm trước Giám đốc Sở GD&ĐT và pháp luật về chất lượng giáo dục, tài chính, nhân sự và an toàn trường học.",
    congViec: [
      {
        bac: "Bậc 5",
        noiDung: "1. Xây dựng Chiến lược phát triển trường 2026-2030 và Kế hoạch giáo dục nhà trường hàng năm theo CT 2018.\n2. Quản trị tổ chức bộ máy, bổ nhiệm TTCM/TPCM, phân công nhiệm vụ và đánh giá xếp loại viên chức.\n3. Là chủ tài khoản, quản lý tài chính, ngân sách và tài sản công đúng Luật Ngân sách và Kế toán.\n4. Chỉ đạo nâng cao chất lượng dạy học, kỳ thi tốt nghiệp THPT, thi tuyển sinh 10, kiểm định CLGD duy trì chuẩn quốc gia.\n5. Lãnh đạo công tác chuyển đổi số, an ninh an toàn trường học và trực tiếp giảng dạy 02 tiết/tuần.",
        sanPham: "- Kế hoạch giáo dục nhà trường được Sở GD&ĐT phê duyệt.\n- Tỷ lệ tốt nghiệp THPT toàn trường đạt ≥ 98.5%.\n- Quyết toán ngân sách minh bạch, kiểm toán tốt 100%.\n- 100% học bạ số được ký đúng hạn; trường đạt chuẩn quốc gia.",
        tieuChi: "Tầm nhìn chiến lược; hiệu quả điều hành; kỷ cương nề nếp; chất lượng giáo dục toàn diện của nhà trường."
      }
    ],
    khungNangLuc: [
      "1. Năng lực chung: Phẩm chất chính trị vững vàng, liêm chính, công bằng (Cấp độ 5); Kỷ luật cao (Cấp độ 5); Giao tiếp đối ngoại chuẩn mực (Cấp độ 5); Chuyển đổi số và ứng dụng AI trong quản trị (Cấp độ 4).",
      "2. Năng lực quản lý: Tư duy chiến lược và quy hoạch phát triển (Cấp độ 5); Quản trị nhân sự và phát triển đội ngũ (Cấp độ 5); Quản trị tài chính công (Cấp độ 4-5); Chỉ đạo chuyên môn và kiểm tra nội bộ (Cấp độ 5)."
    ],
    quanHe: {
      trong: "Lãnh đạo toàn thể cán bộ, giáo viên, nhân viên và học sinh trường THPT Phục Hòa.",
      ngoai: "Sở GD&ĐT Cao Bằng, Huyện ủy, UBND huyện Phục Hòa, Đảng ủy khối, các cơ quan ban ngành và CMHS."
    },
    quyenHan: "Người đứng đầu đơn vị sự nghiệp công lập; chủ tài khoản; ký các quyết định quản lý nhân sự, ban hành quy chế nội bộ, ký học bạ và văn bằng chứng chỉ tốt nghiệp; bổ nhiệm các chức danh tổ trưởng, tổ phó.",
    yeuCau: {
      daoTao: "Bằng Thạc sĩ hoặc Cử nhân Đại học Sư phạm trở lên.",
      boiDuong: "Chứng chỉ bồi dưỡng cán bộ quản lý giáo dục; chứng chỉ Lý luận chính trị; tiêu chuẩn chức danh Hiệu trưởng.",
      kinhNghiem: "Có ít nhất 05 năm trực tiếp giảng dạy và có kinh nghiệm quản lý giáo dục.",
      phamChat: "Bản lĩnh chính trị vững vàng, có uy tín cao, công tâm, gương mẫu, tận tụy.",
      khac: "Năng lực ngoại ngữ và tin học quản lý đạt chuẩn; làm chủ chữ ký số và hệ thống quản trị số."
    }
  },

  ql_pht: {
    ten: "Phó Hiệu trưởng trường trung học phổ thông",
    nhom: "Viên chức quản lý",
    ma: "PHT-THPT-01",
    bac: "Bậc 3 đến Bậc 5",
    linhVuc: "Quản lý chuyên môn dạy học, khảo thí, nề nếp và cơ sở vật chất trường học",
    soLuong: "02 người",
    dinhMuc: "04 tiết/tuần (Phụ cấp chức vụ: 0.50)",
    capQuanLy: "Hiệu trưởng trường THPT Phục Hòa, Giám đốc Sở GD&ĐT",
    viTriLienQuan: "Hiệu trưởng, Tổ trưởng CM, Tổ phó CM, Giáo vụ, Thiết bị TN, GV chủ nhiệm",
    mucTieu: "Giúp Hiệu trưởng quản lý và điều hành công tác chuyên môn dạy học, khảo thí kiểm định, chuyển đổi số, cơ sở vật chất và nề nếp học sinh; điều hành hoạt động của nhà trường khi được ủy quyền.",
    congViec: [
      {
        bac: "Bậc 4",
        noiDung: "1. Chỉ đạo trực tiếp công tác chuyên môn dạy học; duyệt kế hoạch giáo dục tổ chuyên môn; xếp thời khóa biểu.\n2. Chỉ đạo khảo thí, kiểm tra đánh giá, bồi dưỡng HSG và ôn tập thi tốt nghiệp THPT khối 12.\n3. Tổ chức kiểm tra chuyên môn nội bộ, dự giờ thăm lớp, bồi dưỡng đội ngũ giáo viên dạy giỏi.\n4. Quản lý khai thác cơ sở vật chất, phòng bộ môn, thiết bị dạy học và chỉ đạo công tác giáo vụ.\n5. Thực hiện ký số học bạ điện tử, điều hành khi được ủy quyền và trực tiếp giảng dạy 04 tiết/tuần.",
        sanPham: "- Thời khóa biểu khoa học, 100% tiết dạy đúng phân phối chương trình.\n- Ngân hàng đề thi kiểm tra định kỳ chuẩn quy chế.\n- Vượt chỉ tiêu số lượng giải học sinh giỏi cấp tỉnh.\n- 100% học bạ số ký đúng hạn; hoàn thành 04 tiết dạy/tuần.",
        tieuChi: "Chất lượng chuyên môn môn học; tiến độ chương trình; năng lực xử lý tình huống nghiệp vụ sư phạm."
      }
    ],
    khungNangLuc: [
      "1. Năng lực chung: Gương mẫu, chuẩn mực (Cấp độ 5); Kỷ luật cao (Cấp độ 5); Ứng dụng CNTT và AI (Cấp độ 4).",
      "2. Năng lực quản lý: Tổ chức điều hành chuyên môn (Cấp độ 4-5); Kiểm tra đánh giá và khảo thí (Cấp độ 4-5); Quản lý cơ sở vật chất trường học (Cấp độ 4)."
    ],
    quanHe: {
      trong: "Chịu sự phân công của Hiệu trưởng; chỉ đạo các Tổ chuyên môn, bộ phận Giáo vụ, Thiết bị, Y tế.",
      ngoai: "Sở GD&ĐT Cao Bằng, Ban đại diện CMHS, các trường THPT trong cụm chuyên môn."
    },
    quyenHan: "Ký duyệt giáo án, kế hoạch dạy học của tổ chuyên môn; ký duyệt lịch báo giảng, duyệt đề kiểm tra; điều hành nhà trường khi được Hiệu trưởng ủy quyền bằng văn bản.",
    yeuCau: {
      daoTao: "Bằng Thạc sĩ hoặc Cử nhân Đại học Sư phạm trở lên.",
      boiDuong: "Chứng chỉ bồi dưỡng cán bộ quản lý giáo dục; chứng chỉ Lý luận chính trị.",
      kinhNghiem: "Có ít nhất 03 năm giảng dạy THPT và có năng lực chuyên môn vững vàng.",
      phamChat: "Công tâm, trách nhiệm, nhiệt huyết, có uy tín với đồng nghiệp.",
      khac: "Sử dụng thành thạo phần mềm thời khóa biểu, CSDL ngành, chữ ký số."
    }
  },

  ql_ttcm: {
    ten: "Tổ trưởng chuyên môn trường THPT",
    nhom: "Viên chức quản lý",
    ma: "TTCM-THPT-01",
    bac: "Bậc 3 đến Bậc 4",
    linhVuc: "Quản lý và điều hành toàn diện công tác chuyên môn tổ bộ môn",
    soLuong: "02 người",
    dinhMuc: "14 tiết/tuần (Giảm 3 tiết; Phụ cấp chức vụ: 0.25)",
    capQuanLy: "Ban Giám hiệu trường THPT Phục Hòa",
    viTriLienQuan: "Phó Hiệu trưởng, Tổ phó CM, Giáo viên trong tổ, Thiết bị TN",
    mucTieu: "Quản lý, điều hành toàn diện hoạt động chuyên môn của tổ bộ môn; xây dựng kế hoạch dạy học môn học theo CT 2018; tổ chức sinh hoạt chuyên môn NCBL và bồi dưỡng nâng cao năng lực đội ngũ giáo viên trong tổ.",
    congViec: [
      {
        bac: "Bậc 4",
        noiDung: "1. Xây dựng kế hoạch giáo dục môn học của tổ (Phụ lục 1, 2, 3) trình Hiệu trưởng phê duyệt.\n2. Tổ chức sinh hoạt chuyên môn theo nghiên cứu bài học định kỳ ≥ 02 lần/tháng.\n3. Phân công nhiệm vụ, kiểm tra hồ sơ giáo án, dự giờ thăm lớp 2-3 tiết/kỳ/thành viên.\n4. Tổ chức công tác bồi dưỡng học sinh giỏi môn học và phụ đạo học sinh yếu kém.\n5. Thực hiện giảng dạy 14 tiết/tuần và ứng dụng AI đổi mới phương pháp dạy học.",
        sanPham: "- Kế hoạch giáo dục tổ chuyên môn được duyệt trước 30/8.\n- Biên bản sinh hoạt tổ chuyên môn đầy đủ; có ≥ 02 chuyên đề đổi mới PPDH/kỳ.\n- Đội tuyển HSG của tổ đạt giải cấp tỉnh; 100% thành viên hoàn thành nhiệm vụ.",
        tieuChi: "Chất lượng sinh hoạt tổ; tinh thần đoàn kết nội bộ; kết quả môn học của tổ phụ trách."
      }
    ],
    khungNangLuc: [
      "1. Năng lực chung: Gương mẫu chuyên môn (Cấp độ 5); Kỷ luật trách nhiệm (Cấp độ 4); Đổi mới sáng tạo và AI (Cấp độ 4).",
      "2. Năng lực quản lý: Điều hành sinh hoạt chuyên môn (Cấp độ 4); Kiểm tra đánh giá đồng nghiệp (Cấp độ 4); Bồi dưỡng thế hệ kế cận (Cấp độ 4)."
    ],
    quanHe: {
      trong: "Chịu sự chỉ đạo trực tiếp của Phó Hiệu trưởng và Hiệu trưởng; quản lý các thành viên trong tổ.",
      ngoai: "Tổ chuyên môn các trường THPT trong cụm chuyên môn của tỉnh Cao Bằng."
    },
    quyenHan: "Phân công giảng dạy và công tác trong tổ; kiểm tra hồ sơ chuyên môn của giáo viên trong tổ; đề xuất khen thưởng hoặc kỷ luật thành viên trong tổ; duyệt đề thi của tổ.",
    yeuCau: {
      daoTao: "Bằng Cử nhân Đại học Sư phạm đúng chuyên ngành trở lên.",
      boiDuong: "Chứng chỉ chức danh nghề nghiệp GV THPT; là Giáo viên dạy giỏi cấp trường trở lên.",
      kinhNghiem: "Có uy tín chuyên môn cao trong tổ, có khả năng tập hợp và dẫn dắt đồng nghiệp.",
      phamChat: "Nhiệt tình, khách quan, công bằng, trách nhiệm cao.",
      khac: "Làm chủ các phần mềm soạn bài giảng, tạo ma trận đề thi và ứng dụng AI."
    }
  },

  ql_tpcm: {
    ten: "Tổ phó chuyên môn trường THPT",
    nhom: "Viên chức quản lý",
    ma: "TPCM-THPT-01",
    bac: "Bậc 3 đến Bậc 4",
    linhVuc: "Giúp việc cho Tổ trưởng chuyên môn; theo dõi tiến độ và nề nếp tổ bộ môn",
    soLuong: "02 người",
    dinhMuc: "16 tiết/tuần (Giảm 1 tiết; Phụ cấp chức vụ: 0.15)",
    capQuanLy: "Tổ trưởng chuyên môn, Ban Giám hiệu trường THPT Phục Hòa",
    viTriLienQuan: "Tổ trưởng CM, Giáo viên trong tổ, Thiết bị TN, Giáo vụ",
    mucTieu: "Giúp Tổ trưởng theo dõi sát tiến độ chương trình, nề nếp hồ sơ sổ sách, lịch báo giảng điện tử và ký số học bạ của các thành viên trong tổ; điều hành tổ khi Tổ trưởng vắng mặt.",
    congViec: [
      {
        bac: "Bậc 3",
        noiDung: "1. Theo dõi tiến độ thực hiện chương trình môn học hàng tuần của các thành viên trong tổ.\n2. Đôn đốc việc lập lịch báo giảng điện tử, kế hoạch dạy bù, dạy thay của giáo viên.\n3. Kiểm tra nề nếp vào điểm, ký số học bạ điện tử; theo dõi việc sử dụng thiết bị dạy học.\n4. Thực hiện giảng dạy 16 tiết/tuần và thay mặt Tổ trưởng điều hành tổ khi được ủy quyền.",
        sanPham: "- Báo cáo tiến độ chương trình môn học hàng tháng không có trường hợp dạy chậm.\n- 100% thành viên cập nhật lịch báo giảng trước thứ Hai hàng tuần.\n- Hoàn thành đầy đủ 16 tiết dạy/tuần theo phân công.",
        tieuChi: "Tính đôn đốc kịp thời; theo dõi sát sao; tinh thần trách nhiệm và tương trợ đồng nghiệp."
      }
    ],
    khungNangLuc: [
      "1. Năng lực chung: Trách nhiệm, cần mẫn (Cấp độ 4); Giao tiếp nội bộ tốt (Cấp độ 4); Ứng dụng CNTT (Cấp độ 4).",
      "2. Năng lực quản lý: Kiểm tra nề nếp sổ sách (Cấp độ 3-4); Điều phối công việc nhóm (Cấp độ 3)."
    ],
    quanHe: {
      trong: "Phối hợp chặt chẽ với Tổ trưởng CM; hỗ trợ các giáo viên trong tổ; báo cáo Ban Giám hiệu khi cần.",
      ngoai: "Tham gia các buổi sinh hoạt cụm chuyên môn cùng Tổ trưởng."
    },
    quyenHan: "Kiểm tra tiến độ chương trình và lịch báo giảng của các thành viên; ký biên bản cuộc họp tổ khi được ủy quyền; thay mặt Tổ trưởng điều hành tổ khi vắng mặt.",
    yeuCau: {
      daoTao: "Bằng Cử nhân Đại học Sư phạm đúng chuyên ngành trở lên.",
      boiDuong: "Chứng chỉ chức danh nghề nghiệp GV THPT.",
      kinhNghiem: "Có kinh nghiệm giảng dạy vững vàng, cẩn thận và có tinh thần trách nhiệm.",
      phamChat: "Nhiệt tình, chu đáo, thẳng thắn, hỗ trợ đồng nghiệp.",
      khac: "Thành thạo các phần mềm quản lý sổ sách điện tử và học bạ số."
    }
  }
};

// Map default fallback for remaining subjects
const DEFAULT_SUBJECTS = {
  gv_su: "Lịch sử", gv_gdtc: "Giáo dục thể chất", gv_gdqp: "GDQP-AN",
  gv_dia: "Địa lí", gv_ktpl: "GD KT&PL", gv_ly: "Vật lí",
  gv_hoa: "Hóa học", gv_sinh: "Sinh học", gv_cn: "Công nghệ", gv_tin: "Tin học"
};

// Generate subject configs dynamically if not explicitly defined
Object.keys(DEFAULT_SUBJECTS).forEach(key => {
  if (!POSITIONS_DATA[key]) {
    const subName = DEFAULT_SUBJECTS[key];
    const isSpecial = key === 'gv_tin' ? 'lập trình Python và công nghệ số' :
                     key === 'gv_ly' ? 'thí nghiệm Vật lý và giáo dục STEM' :
                     key === 'gv_hoa' ? 'thực hành hóa chất và an toàn thí nghiệm' :
                     key === 'gv_sinh' ? 'thực hành kính hiển vi và di truyền sinh thái' :
                     key === 'gv_gdtc' ? 'rèn luyện thể lực và Hội khỏe Phù Đổng' :
                     key === 'gv_gdqp' ? 'huấn luyện quân sự, điều lệnh và an ninh quốc gia' : 'năng lực môn học theo CT 2018';

    POSITIONS_DATA[key] = {
      ten: `Giáo viên trung học phổ thông môn ${subName}`,
      nhom: "Chuyên môn, nghiệp vụ",
      ma: "GV-THPT-01",
      bac: "Bậc 3 đến Bậc 4",
      linhVuc: `Giáo dục THPT – Giảng dạy và giáo dục môn ${subName} (CT GDPT 2018)`,
      soLuong: key === 'gv_su' || key === 'gv_dia' || key === 'gv_gdtc' || key === 'gv_cn' ? "02 người" : "01 người",
      dinhMuc: "17 tiết/tuần",
      capQuanLy: `Tổ trưởng Chuyên môn, Ban Giám hiệu trường THPT Phục Hòa`,
      viTriLienQuan: "Tổ trưởng CM, GV chủ nhiệm, Thiết bị thí nghiệm, Giáo vụ",
      mucTieu: `Trực tiếp giảng dạy và giáo dục môn ${subName} theo CT GDPT 2018 nhằm phát triển ${isSpecial} cho học sinh; hoàn thành chỉ tiêu chất lượng bộ môn và bồi dưỡng học sinh giỏi môn ${subName} cho trường THPT Phục Hòa.`,
      congViec: [
        {
          bac: "Bậc 3",
          noiDung: `1. Giảng dạy môn ${subName} theo phân phối chương trình GDPT 2018 (định mức 17 tiết/tuần).\n2. Xây dựng Kế hoạch bài dạy (giáo án), đổi mới PPDH tích cực, khai thác tốt phòng bộ môn và thiết bị dạy học.\n3. Kiểm tra đánh giá học sinh thường xuyên, định kỳ; vào điểm số điện tử đúng hạn.\n4. Thực hiện công tác giáo viên chủ nhiệm, phối hợp chặt chẽ với cha mẹ học sinh.\n5. Bồi dưỡng học sinh yếu, phụ đạo học sinh có nguy cơ chưa đạt chuẩn môn ${subName}.`,
          sanPham: `- 100% tiết dạy có giáo án chuẩn bị kỹ lưỡng trước khi lên lớp.\n- Sổ điểm điện tử, học bạ số ký nhận xét đúng hạn.\n- Tỷ lệ học sinh đạt chuẩn môn ${subName} ≥ 96%.\n- 100% tiết thực hành bảo đảm an toàn thiết bị.`,
          tieuChi: "Chất lượng giờ dạy; tính chuẩn mực khoa học; khả năng thu hút học sinh; hoàn thành đúng tiến độ."
        },
        {
          bac: "Bậc 4",
          noiDung: `1. Giảng dạy chuyên đề nâng cao, dạy ôn thi tốt nghiệp THPT hoặc hội thao, ngày hội chuyên môn.\n2. Nòng cốt sinh hoạt chuyên môn NCBL, đổi mới phương pháp dạy học môn ${subName} cấp trường.\n3. Xây dựng ma trận đề, ngân hàng câu hỏi đề kiểm tra định kỳ chuẩn quy chế.\n4. Bồi dưỡng đội tuyển học sinh giỏi môn ${subName} dự thi cấp tỉnh Cao Bằng.\n5. Đổi mới sáng tạo: Có sáng kiến kinh nghiệm hoặc bài giảng số ứng dụng AI được nghiệm thu.`,
          sanPham: `- Ngân hàng đề kiểm tra đánh giá chuẩn mực.\n- Có học sinh đạt giải HSG môn ${subName} cấp tỉnh.\n- Kết quả thi tốt nghiệp / kiểm tra môn học đạt tỷ lệ cao.\n- Có chuyên đề chuyên môn đổi mới PPDH được chia sẻ trong tổ.`,
          tieuChi: "Khả năng dẫn dắt chuyên môn; kết quả học sinh giỏi cấp tỉnh; mức độ đổi mới sáng tạo."
        }
      ],
      khungNangLuc: [
        "1. Năng lực chung: Phẩm chất đạo đức nhà giáo (Cấp độ 5); Kỷ luật cao (Cấp độ 4); Ứng dụng CNTT và AI (Cấp độ 4); Giao tiếp sư phạm chuẩn mực (Cấp độ 4).",
        `2. Năng lực chuyên môn: Nắm vững kiến thức chuyên sâu môn ${subName} CT 2018 (Cấp độ 4-5); Kỹ năng khai thác phòng bộ môn và thiết bị chuyên dùng; Kiểm tra đánh giá phát triển năng lực học sinh (Cấp độ 4).`
      ],
      quanHe: {
        trong: "Báo cáo Tổ trưởng CM, Ban Giám hiệu; phối hợp với GV bộ môn khác, Thiết bị thí nghiệm, Giáo vụ.",
        ngoai: "Ban đại diện CMHS; giáo viên cốt cán bộ môn trong cụm chuyên môn tỉnh Cao Bằng."
      },
      quyenHan: `Đánh giá, xếp loại kết quả học tập môn ${subName} của học sinh; đề xuất khen thưởng; chủ động lựa chọn phương pháp dạy học và học liệu phù hợp.`,
      yeuCau: {
        daoTao: `Bằng Cử nhân (Đại học) sư phạm chuyên ngành ${subName} trở lên.`,
        boiDuong: "Chứng chỉ tiêu chuẩn chức danh nghề nghiệp GV THPT.",
        kinhNghiem: `Có kinh nghiệm giảng dạy môn ${subName} CT 2018 và bồi dưỡng học sinh.`,
        phamChat: "Nhiệt huyết, yêu nghề, tận tâm, có tinh thần trách nhiệm.",
        khac: "Sử dụng thành thạo máy chiếu, phần mềm chuyên ngành và ký số học bạ."
      }
    };
  }
});

// HÀM RENDER BẢN MÔ TẢ CHUẨN CÔNG VỤ NGHỊ ĐỊNH 232/2026/NĐ-CP
function renderOfficialDoc(posKey) {
  const data = POSITIONS_DATA[posKey];
  if (!data) return;

  const paper = document.getElementById('officialDocPaper');
  if (!paper) return;

  const isQuanLy = data.nhom.includes('quản lý');
  const isHoTro = data.nhom.includes('Hỗ trợ');
  const isChuyenMon = data.nhom.includes('Chuyên môn');

  let tableHtml = '';
  data.congViec.forEach(item => {
    tableHtml += `
      <tr>
        <td style="text-align: center; font-weight: bold; width: 80px;">${item.bac}</td>
        <td style="white-space: pre-line;">${item.noiDung}</td>
        <td style="white-space: pre-line;">${item.sanPham}</td>
        <td>${item.tieuChi}</td>
      </tr>
    `;
  });

  const html = `
    <!-- Quốc hiệu & Tiêu ngữ -->
    <div class="doc-header-grid">
      <div class="doc-header-left">
        <p>SỞ GIÁO DỤC VÀ ĐÀO TẠO CAO BẰNG</p>
        <strong>TRƯỜNG THPT PHỤC HÒA</strong>
        <div class="doc-divider"></div>
      </div>
      <div class="doc-header-right">
        <strong>CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM</strong>
        <p style="font-weight: bold; font-style: italic;">Độc lập – Tự do – Hạnh phúc</p>
        <div class="doc-divider"></div>
      </div>
    </div>

    <!-- Tiêu đề chính -->
    <div class="doc-main-heading">
      <h3>BẢN MÔ TẢ CÔNG VIỆC VÀ KHUNG NĂNG LỰC VỊ TRÍ VIỆC LÀM</h3>
      <p>(Ban hành kèm theo Nghị định số 232/2026/NĐ-CP ngày 26 tháng 6 năm 2026 của Chính phủ)</p>
    </div>

    <!-- I. THÔNG TIN CHUNG -->
    <h4 class="doc-sec-title">I. THÔNG TIN CHUNG</h4>
    <ul class="doc-info-list">
      <li><strong>1. Tên vị trí việc làm:</strong> <span style="font-weight: bold; color: #1e3a8a;">${data.ten}</span></li>
      <li><strong>2. Nhóm vị trí việc làm:</strong> 
        Quản lý [${isQuanLy ? '☒' : '☐'}] &nbsp;&nbsp;&nbsp;&nbsp;
        Chuyên môn, nghiệp vụ [${isChuyenMon ? '☒' : '☐'}] &nbsp;&nbsp;&nbsp;&nbsp;
        Hỗ trợ, phục vụ [${isHoTro ? '☒' : '☐'}]
      </li>
      <li><strong>3. Bậc nghề nghiệp sử dụng:</strong> ${data.bac}</li>
      <li><strong>4. Lĩnh vực hoạt động:</strong> ${data.linhVuc}</li>
      <li><strong>5. Đơn vị công tác:</strong> Trường THPT Phục Hòa, huyện Phục Hòa, tỉnh Cao Bằng</li>
      <li><strong>6. Cấp quản lý trực tiếp:</strong> ${data.capQuanLy}</li>
      <li><strong>7. Vị trí việc làm liên quan:</strong> ${data.viTriLienQuan}</li>
      <li><strong>8. Số lượng biên chế đảm nhiệm:</strong> <strong>${data.soLuong}</strong> (Định mức: ${data.dinhMuc})</li>
    </ul>

    <!-- II. MỤC TIÊU VỊ TRÍ VIỆC LÀM -->
    <h4 class="doc-sec-title">II. MỤC TIÊU VỊ TRÍ VIỆC LÀM</h4>
    <p style="text-align: justify; padding: 6px 10px; background: #f8fafc; border-left: 3px solid #1e3a8a;">
      ${data.mucTieu}
    </p>

    <!-- III / IV / V. CÔNG VIỆC, KẾT QUẢ, SẢN PHẨM -->
    <h4 class="doc-sec-title">
      ${isChuyenMon ? 'III. CÔNG VIỆC, KẾT QUẢ, SẢN PHẨM CHUYÊN MÔN' : 
        isHoTro ? 'V. CÔNG VIỆC, KẾT QUẢ, SẢN PHẨM HỖ TRỢ' : 
        'IV. CÔNG VIỆC, KẾT QUẢ, SẢN PHẨM QUẢN LÝ'}
    </h4>
    <table class="doc-tbl">
      <thead>
        <tr>
          <th>Bậc</th>
          <th style="width: 38%;">Nội dung công việc</th>
          <th style="width: 32%;">Sản phẩm, Kết quả đầu ra</th>
          <th>Tiêu chí đánh giá</th>
        </tr>
      </thead>
      <tbody>
        ${tableHtml}
      </tbody>
    </table>

    <!-- VI. KHUNG NĂNG LỰC -->
    <h4 class="doc-sec-title">VI. KHUNG NĂNG LỰC CỦA VỊ TRÍ VIỆC LÀM</h4>
    <div style="font-size: 0.95rem;">
      ${data.khungNangLuc.map(line => `<p style="margin-bottom: 6px;">${line}</p>`).join('')}
    </div>

    <!-- VII. MỐI QUAN HỆ CÔNG TÁC -->
    <h4 class="doc-sec-title">VII. MỐI QUAN HỆ CÔNG TÁC</h4>
    <ul class="doc-info-list">
      <li><strong>1. Bên trong:</strong> ${data.quanHe.trong}</li>
      <li><strong>2. Bên ngoài:</strong> ${data.quanHe.ngoai}</li>
    </ul>

    <!-- VIII. PHẠM VI QUYỀN HẠN -->
    <h4 class="doc-sec-title">VIII. PHẠM VI QUYỀN HẠN</h4>
    <p style="text-align: justify; margin-bottom: 8px;">
      ${data.quyenHan}
    </p>

    <!-- IX. YÊU CẦU VỀ TRÌNH ĐỘ, KINH NGHIỆM, PHẨM CHẤT -->
    <h4 class="doc-sec-title">IX. YÊU CẦU VỀ TRÌNH ĐỘ, KINH NGHIỆM, PHẨM CHẤT</h4>
    <table class="doc-tbl">
      <thead>
        <tr>
          <th style="width: 25%;">Nhóm yêu cầu</th>
          <th>Yêu cầu cụ thể của vị trí việc làm</th>
        </tr>
      </thead>
      <tbody>
        <tr>
          <td><strong>Trình độ đào tạo</strong></td>
          <td>${data.yeuCau.daoTao}</td>
        </tr>
        <tr>
          <td><strong>Bồi dưỡng, chứng chỉ</strong></td>
          <td>${data.yeuCau.boiDuong}</td>
        </tr>
        <tr>
          <td><strong>Kinh nghiệm chuyên môn</strong></td>
          <td>${data.yeuCau.kinhNghiem}</td>
        </tr>
        <tr>
          <td><strong>Phẩm chất cá nhân</strong></td>
          <td>${data.yeuCau.phamChat}</td>
        </tr>
        <tr>
          <td><strong>Yêu cầu khác (CNTT, AI, Ngoại ngữ)</strong></td>
          <td>${data.yeuCau.khac}</td>
        </tr>
      </tbody>
    </table>

    <!-- X. KÝ DUYỆT -->
    <div class="doc-sign-grid">
      <div class="doc-sign-col">
        <p>NGƯỜI LẬP BIỂU<br/><small>(Ký và ghi rõ họ tên)</small></p>
      </div>
      <div class="doc-sign-col">
        <p>Phục Hòa, ngày 07 tháng 10 năm 2026<br/><strong>HIỆU TRƯỞNG</strong><br/><small>(Ký tên, đóng dấu)</small></p>
      </div>
    </div>
  `;

  paper.innerHTML = html;
}

document.addEventListener('DOMContentLoaded', () => {
  // Initial render official doc
  renderOfficialDoc('gv_van');

  // Selector change handler
  const docSelect = document.getElementById('docPositionSelect');
  if (docSelect) {
    docSelect.addEventListener('change', (e) => {
      renderOfficialDoc(e.target.value);
      const text = e.target.options[e.target.selectedIndex].text;
      showToast(`Đã nạp Bản mô tả chuẩn NĐ 232: ${text}`);
    });
  }

  // ---- Scroll-top button ----
  const scrollBtn = document.getElementById('scrollTopBtn');
  window.addEventListener('scroll', () => {
    if (window.scrollY > 400) {
      scrollBtn.classList.add('visible');
    } else {
      scrollBtn.classList.remove('visible');
    }
  });

  // ---- Intersection Observer – animate-in elements ----
  const observerOptions = {
    threshold: 0.08,
    rootMargin: '0px 0px -40px 0px'
  };

  const observer = new IntersectionObserver((entries) => {
    entries.forEach(entry => {
      if (entry.isIntersecting) {
        entry.target.classList.add('visible');
        observer.unobserve(entry.target);
      }
    });
  }, observerOptions);

  const animTargets = document.querySelectorAll(
    '.block-box, .subject-card, .info-card, .task-item, .product-card, .kpi-card, .stat-card, .comp-table-wrap, .category-header, .master-table-card'
  );
  animTargets.forEach((el, i) => {
    el.classList.add('animate-in');
    el.style.transitionDelay = `${Math.min(i * 0.02, 0.2)}s`;
    observer.observe(el);
  });

  // ---- Toast Message Helper ----
  const toastMsg = document.getElementById('toastMsg');
  const toastText = document.getElementById('toastText');
  function showToast(text) {
    if (!toastMsg) return;
    toastText.textContent = text;
    toastMsg.classList.add('show');
    setTimeout(() => {
      toastMsg.classList.remove('show');
    }, 3500);
  }

  // Hook download buttons
  const downloadBtns = document.querySelectorAll('#btn-excel-hero, #btn-download-excel-banner, #floatingExcelBtn');
  downloadBtns.forEach(btn => {
    btn.addEventListener('click', () => {
      showToast('✓ Đang tải xuống trọn bộ 07 Sheet Đề án 35 Biên chế (.xlsx)...');
    });
  });

  // ---- Category Tabs Switcher / Filter ----
  const roleTabBtns = document.querySelectorAll('.role-tab-btn[data-filter]');
  const catBlocks = document.querySelectorAll('.category-block, .official-doc-section');

  roleTabBtns.forEach(btn => {
    btn.addEventListener('click', () => {
      roleTabBtns.forEach(b => b.classList.remove('active'));
      btn.classList.add('active');

      const filter = btn.getAttribute('data-filter');

      if (filter === 'all') {
        catBlocks.forEach(blk => {
          blk.style.display = 'block';
        });
        showToast('Hiển thị toàn cảnh Đề án: Đủ 35 Biên chế (03 Danh mục VTVL)');
      } else {
        catBlocks.forEach(blk => {
          if (blk.getAttribute('data-cat') === filter) {
            blk.style.display = 'block';
            blk.scrollIntoView({ behavior: 'smooth', block: 'start' });
          } else {
            blk.style.display = 'none';
          }
        });
        const catNames = {
          'doc-mau': 'Bản mô tả chuẩn mẫu công vụ: Phần B Phụ lục IV/VI Nghị định 232/2026/NĐ-CP',
          'master': 'Bảng tổng hợp Phụ lục I: 35 Biên chế người làm việc',
          'dm1': 'Danh mục 1: Vị trí Quản lý (07 người: Hiệu trưởng, 2 Phó HT, 2 TTCM, 2 TPCM)',
          'dm2': 'Danh mục 2: Vị trí Chuyên môn (22 GV: Đầy đủ 13 môn học CT 2018)',
          'dm3': 'Danh mục 3: Vị trí Hỗ trợ (06 người: Kế toán, Văn thư, TB-TN, Giáo vụ, Y tế, Thủ quỹ)'
        };
        showToast(`Đang lọc: ${catNames[filter] || filter}`);
      }
    });
  });

  // ---- 13 Subjects Sub-Filter ----
  const subjFilterBtns = document.querySelectorAll('.subj-filter-btn[data-sub]');
  const subjectCards = document.querySelectorAll('.subject-card[data-sub-key]');

  subjFilterBtns.forEach(btn => {
    btn.addEventListener('click', () => {
      subjFilterBtns.forEach(b => b.classList.remove('active'));
      btn.classList.add('active');

      const subKey = btn.getAttribute('data-sub');

      if (subKey === 'all') {
        subjectCards.forEach(card => card.style.display = 'flex');
        showToast('Hiển thị đầy đủ 13 môn học giảng dạy (22 GV)');
      } else {
        subjectCards.forEach(card => {
          if (card.getAttribute('data-sub-key') === subKey) {
            card.style.display = 'flex';
            card.scrollIntoView({ behavior: 'smooth', block: 'center' });
          } else {
            card.style.display = 'none';
          }
        });
        showToast(`Đang hiển thị môn: ${btn.textContent.trim()}`);
      }
    });
  });

  // ---- Live Search Filter ----
  const searchInput = document.getElementById('liveSearchInput');
  if (searchInput) {
    searchInput.addEventListener('input', (e) => {
      const query = e.target.value.trim().toLowerCase();
      const taskItems = document.querySelectorAll('.task-item');
      const kpiCards = document.querySelectorAll('.kpi-card');
      const blockBoxes = document.querySelectorAll('.block-box');
      const subCards = document.querySelectorAll('.subject-card');
      const tableRows = document.querySelectorAll('.master-table tbody tr:not(.group-header-row):not(.group-total-row)');

      if (!query) {
        catBlocks.forEach(blk => blk.style.display = 'block');
        taskItems.forEach(item => {
          item.style.display = 'flex';
          item.style.backgroundColor = '';
        });
        kpiCards.forEach(card => card.style.display = 'block');
        blockBoxes.forEach(box => box.style.display = 'block');
        subCards.forEach(card => card.style.display = 'flex');
        tableRows.forEach(row => row.style.display = '');
        return;
      }

      // Show all category blocks while searching
      catBlocks.forEach(blk => blk.style.display = 'block');
      roleTabBtns.forEach(b => b.classList.remove('active'));
      document.querySelector('.role-tab-btn[data-filter="all"]')?.classList.add('active');

      // Filter subject cards
      subCards.forEach(card => {
        const text = card.textContent.toLowerCase();
        if (text.includes(query)) {
          card.style.display = 'flex';
        } else {
          card.style.display = 'none';
        }
      });

      // Filter block boxes
      blockBoxes.forEach(box => {
        const text = box.textContent.toLowerCase();
        if (text.includes(query)) {
          box.style.display = 'block';
        } else {
          box.style.display = 'none';
        }
      });

      // Filter table rows
      tableRows.forEach(row => {
        const text = row.textContent.toLowerCase();
        if (text.includes(query)) {
          row.style.display = '';
          row.style.backgroundColor = 'rgba(254, 240, 138, 0.35)';
        } else {
          row.style.display = 'none';
          row.style.backgroundColor = '';
        }
      });

      // Filter KPI cards
      kpiCards.forEach(card => {
        const text = card.textContent.toLowerCase();
        if (text.includes(query)) {
          card.style.display = 'block';
        } else {
          card.style.display = 'none';
        }
      });
    });
  }

  // ---- Table row hover effects ----
  document.querySelectorAll('.comp-table tbody tr, .master-table tbody tr').forEach(row => {
    row.addEventListener('mouseenter', () => {
      row.style.background = 'rgba(245, 158, 11, 0.08)';
    });
    row.addEventListener('mouseleave', () => {
      row.style.background = '';
    });
  });

  console.log(
    '%c🏛️ TRƯỜNG THPT PHỤC HÒA – CAO BẰNG\n' +
    '%cĐỀ ÁN VỊ TRÍ VIỆC LÀM VIÊN CHỨC (35 BIÊN CHẾ - 03 DANH MỤC)\n' +
    '%cCăn cứ Nghị định số 232/2026/NĐ-CP & Mẫu bản mô tả chuẩn Phần B\n' +
    '%cTự động đẩy lên GitHub & Vercel | Năm học 2026–2027',
    'color: #f59e0b; font-size: 16px; font-weight: bold;',
    'color: #3b82f6; font-size: 14px; font-weight: bold;',
    'color: #10b981; font-size: 12px;',
    'color: #94a3b8; font-size: 11px;'
  );
});
