/*************************************************
 * IMPORT_SOQUY_CU.GS
 * Nhập liệu 1 LẦN DUY NHẤT dữ liệu Sổ Quỹ / Sổ Cơm / Sổ Quỹ
 * Công đoàn từ sổ Excel cũ, ghi thẳng bằng Date object thật
 * (new Date(năm, tháng, ngày)) — KHÔNG đi qua paste/CSV nên
 * không thể bị Google Sheets hiểu sai ngày/tháng theo locale.
 *
 * CÁCH DÙNG: dán file này vào Apps Script Editor, chọn hàm
 * "importDuLieuSoQuyCu" ở dropdown trên cùng, bấm Run.
 * Chỉ chạy 1 LẦN — chạy lại sẽ nhập trùng dữ liệu.
 * Trước khi chạy: xoá sạch các dòng dữ liệu (từ dòng 2 trở
 * xuống, giữ lại dòng tiêu đề) trong PHIEU_THU, PHIEU_CHI,
 * SO_COM nếu trước đó đã lỡ dán CSV bị sai.
 *************************************************/

function importDuLieuSoQuyCu() {
  const nguoiLap = "Nhập liệu sổ cũ";
  const now = new Date();
  const shThu = _sheet(SHEET_PHIEU_THU);
  const shChi = _sheet(SHEET_PHIEU_CHI);
  const shCom = _sheet(SHEET_SO_COM);

  // Ép cột chứa số phiếu / mã dạng text để Sheets không tự ý
  // đổi thành ngày tháng khi ghi (dù setValues ít khi bị vậy,
  // set trước cho chắc).
  shThu.getRange(2, 2, 500, 1).setNumberFormat("@");
  shThu.getRange(2, 21, 500, 3).setNumberFormat("@");
  shChi.getRange(2, 2, 500, 1).setNumberFormat("@");
  shChi.getRange(2, 21, 500, 3).setNumberFormat("@");

  const thuRows = [];
  thuRows.push({
    soPhieu: "01/09.26", ngayCT: new Date(2026, 8, 1), ngayHT: new Date(2026, 8, 1),
    nguoi: "NGÔ THỊ KIỀU NƯƠNG", noiDung: "Thu tiền cân thuê 1039,1040,1041", soTien: 250000,
    loaiQuy: "Quỹ tiền mặt", tkDoiUng: "7113", maNV: "01.006.001",
    chiNhanh: "", ghiChu: "Nhập từ Sổ Quỹ tổng (Excel cũ)"
  });
  thuRows.push({
    soPhieu: "02/09.26", ngayCT: new Date(2026, 8, 3), ngayHT: new Date(2026, 8, 3),
    nguoi: "NGÔ THỊ KIỀU NƯƠNG", noiDung: "Thu tiền HT TX nhập quỹ", soTien: 21590000,
    loaiQuy: "Quỹ tiền mặt", tkDoiUng: "6426", maNV: "01.006.001",
    chiNhanh: "", ghiChu: "Nhập từ Sổ Quỹ tổng (Excel cũ)"
  });
  thuRows.push({
    soPhieu: "03/09.26", ngayCT: new Date(2026, 8, 4), ngayHT: new Date(2026, 8, 4),
    nguoi: "NGÔ THỊ KIỀU NƯƠNG", noiDung: "Thu tiền cân thuê 1042,1043,1044", soTien: 200000,
    loaiQuy: "Quỹ tiền mặt", tkDoiUng: "7113", maNV: "01.006.001",
    chiNhanh: "", ghiChu: "Nhập từ Sổ Quỹ tổng (Excel cũ)"
  });
  thuRows.push({
    soPhieu: "04/09.26", ngayCT: new Date(2026, 8, 5), ngayHT: new Date(2026, 8, 5),
    nguoi: "NGÔ THỊ KIỀU NƯƠNG", noiDung: "Thu tiền cân thuê 04,05,1045,1046,1047,1048,1049", soTien: 450000,
    loaiQuy: "Quỹ tiền mặt", tkDoiUng: "7113", maNV: "01.006.001",
    chiNhanh: "", ghiChu: "Nhập từ Sổ Quỹ tổng (Excel cũ)"
  });
  thuRows.push({
    soPhieu: "05/09.26", ngayCT: new Date(2026, 8, 7), ngayHT: new Date(2026, 8, 7),
    nguoi: "NGÔ THỊ KIỀU NƯƠNG", noiDung: "Thu tiền cân thuê 06,07,08,09,10,12,13,14", soTien: 500000,
    loaiQuy: "Quỹ tiền mặt", tkDoiUng: "7113", maNV: "01.006.001",
    chiNhanh: "", ghiChu: "Nhập từ Sổ Quỹ tổng (Excel cũ)"
  });
  thuRows.push({
    soPhieu: "06/09.26", ngayCT: new Date(2026, 8, 7), ngayHT: new Date(2026, 8, 7),
    nguoi: "VÕ VĂN PHỤNG", noiDung: "Nhận tiền mượn sếp", soTien: 50000000,
    loaiQuy: "Quỹ tiền mặt", tkDoiUng: "1388NB", maNV: "01.001.001",
    chiNhanh: "", ghiChu: "Nhập từ Sổ Quỹ tổng (Excel cũ)"
  });
  thuRows.push({
    soPhieu: "07/09.26", ngayCT: new Date(2026, 8, 7), ngayHT: new Date(2026, 8, 7),
    nguoi: "DƯƠNG THỊ THANH NGA", noiDung: "Nhận tiền BIDV QN", soTien: 200000000,
    loaiQuy: "Quỹ tiền mặt", tkDoiUng: "11211", maNV: "",
    chiNhanh: "", ghiChu: "Nhập từ Sổ Quỹ tổng (Excel cũ)"
  });
  thuRows.push({
    soPhieu: "08/09.26", ngayCT: new Date(2026, 8, 8), ngayHT: new Date(2026, 8, 8),
    nguoi: "NGÔ THỊ KIỀU NƯƠNG", noiDung: "Thu tiền cân thuê 12,14,15,16,17,18,19,20", soTien: 550000,
    loaiQuy: "Quỹ tiền mặt", tkDoiUng: "7113", maNV: "01.006.001",
    chiNhanh: "", ghiChu: "Nhập từ Sổ Quỹ tổng (Excel cũ)"
  });
  thuRows.push({
    soPhieu: "09/09.26", ngayCT: new Date(2026, 8, 9), ngayHT: new Date(2026, 8, 9),
    nguoi: "NGÔ THỊ KIỀU NƯƠNG", noiDung: "Thu tiền cân thuê 19,21,22,23,24,25", soTien: 350000,
    loaiQuy: "Quỹ tiền mặt", tkDoiUng: "7113", maNV: "01.006.001",
    chiNhanh: "", ghiChu: "Nhập từ Sổ Quỹ tổng (Excel cũ)"
  });
  thuRows.push({
    soPhieu: "10/09.26", ngayCT: new Date(2026, 8, 9), ngayHT: new Date(2026, 8, 9),
    nguoi: "DƯƠNG THỊ THANH NGA", noiDung: "Nhận tiền BIDV QN", soTien: 1334000000,
    loaiQuy: "Quỹ tiền mặt", tkDoiUng: "1121", maNV: "",
    chiNhanh: "", ghiChu: "Nhập từ Sổ Quỹ tổng (Excel cũ)"
  });
  thuRows.push({
    soPhieu: "11/09.26", ngayCT: new Date(2026, 8, 10), ngayHT: new Date(2026, 8, 10),
    nguoi: "NGÔ THỊ KIỀU NƯƠNG", noiDung: "Thu tiền cân thuê 26,28,29,30", soTien: 300000,
    loaiQuy: "Quỹ tiền mặt", tkDoiUng: "7113", maNV: "01.006.001",
    chiNhanh: "", ghiChu: "Nhập từ Sổ Quỹ tổng (Excel cũ)"
  });
  thuRows.push({
    soPhieu: "12/09.26", ngayCT: new Date(2026, 8, 10), ngayHT: new Date(2026, 8, 10),
    nguoi: "ĐINH THỊ THANH SƯƠNG", noiDung: "Thu tiền CK bảo hiểm 2 xe 06732,06371", soTien: 1920000,
    loaiQuy: "Quỹ tiền mặt", tkDoiUng: "7113", maNV: "01.006.002",
    chiNhanh: "", ghiChu: "Nhập từ Sổ Quỹ tổng (Excel cũ)"
  });
  thuRows.push({
    soPhieu: "CD-01/02", ngayCT: new Date(2026, 1, 3), ngayHT: new Date(2026, 1, 3),
    nguoi: "", noiDung: "Nhận tiền thưởng CĐ cấp trên", soTien: 1400000,
    loaiQuy: "Quỹ công đoàn", tkDoiUng: "", maNV: "",
    chiNhanh: "", ghiChu: "Nhập từ Sổ Quỹ Công đoàn (Excel cũ)"
  });
  thuRows.push({
    soPhieu: "CD-01/03", ngayCT: new Date(2026, 2, 12), ngayHT: new Date(2026, 2, 12),
    nguoi: "", noiDung: "Nhận tiền KPCĐ cấp trên cấp T01+T02", soTien: 8728200,
    loaiQuy: "Quỹ công đoàn", tkDoiUng: "", maNV: "",
    chiNhanh: "", ghiChu: "Nhập từ Sổ Quỹ Công đoàn (Excel cũ)"
  });
  thuRows.push({
    soPhieu: "CD-01/03", ngayCT: new Date(2026, 2, 14), ngayHT: new Date(2026, 2, 14),
    nguoi: "", noiDung: "Nhận tiền ĐPCĐ tháng 01+02 nhập quỹ CĐ", soTien: 2909400,
    loaiQuy: "Quỹ công đoàn", tkDoiUng: "", maNV: "",
    chiNhanh: "", ghiChu: "Nhập từ Sổ Quỹ Công đoàn (Excel cũ)"
  });
  thuRows.push({
    soPhieu: "CD-01/03", ngayCT: new Date(2026, 3, 13), ngayHT: new Date(2026, 3, 13),
    nguoi: "", noiDung: "Nhận tiền KPCĐ cấp trên cấp T03", soTien: 4400325,
    loaiQuy: "Quỹ công đoàn", tkDoiUng: "", maNV: "",
    chiNhanh: "", ghiChu: "Nhập từ Sổ Quỹ Công đoàn (Excel cũ)"
  });
  thuRows.push({
    soPhieu: "CD-02/03", ngayCT: new Date(2026, 3, 13), ngayHT: new Date(2026, 3, 13),
    nguoi: "", noiDung: "Nhận tiền ĐPCĐ tháng 03 nhập quỹ CĐ", soTien: 1466775,
    loaiQuy: "Quỹ công đoàn", tkDoiUng: "", maNV: "",
    chiNhanh: "", ghiChu: "Nhập từ Sổ Quỹ Công đoàn (Excel cũ)"
  });
  thuRows.push({
    soPhieu: "CD-01/05", ngayCT: new Date(2026, 4, 18), ngayHT: new Date(2026, 4, 18),
    nguoi: "", noiDung: "Nhận tiền KPCĐ cấp trên cấp Tháng 04/2026", soTien: 4168500,
    loaiQuy: "Quỹ công đoàn", tkDoiUng: "", maNV: "",
    chiNhanh: "", ghiChu: "Nhập từ Sổ Quỹ Công đoàn (Excel cũ)"
  });
  thuRows.push({
    soPhieu: "CD-02/05", ngayCT: new Date(2026, 4, 28), ngayHT: new Date(2026, 4, 28),
    nguoi: "", noiDung: "Nhận tiền ĐPCĐ Tháng 04 nhập quỹ", soTien: 1389000,
    loaiQuy: "Quỹ công đoàn", tkDoiUng: "", maNV: "",
    chiNhanh: "", ghiChu: "Nhập từ Sổ Quỹ Công đoàn (Excel cũ)"
  });
  thuRows.push({
    soPhieu: "CD-01/06", ngayCT: new Date(2026, 5, 8), ngayHT: new Date(2026, 5, 8),
    nguoi: "", noiDung: "Nhận tiền KPCĐ cấp trên cấp Tháng 05/2026", soTien: 4164600,
    loaiQuy: "Quỹ công đoàn", tkDoiUng: "", maNV: "",
    chiNhanh: "", ghiChu: "Nhập từ Sổ Quỹ Công đoàn (Excel cũ)"
  });
  thuRows.push({
    soPhieu: "CD-02/06", ngayCT: new Date(2026, 5, 17), ngayHT: new Date(2026, 5, 17),
    nguoi: "", noiDung: "Nhận tiền ĐPCĐ Tháng 05 nhập quỹ", soTien: 1388200,
    loaiQuy: "Quỹ công đoàn", tkDoiUng: "", maNV: "",
    chiNhanh: "", ghiChu: "Nhập từ Sổ Quỹ Công đoàn (Excel cũ)"
  });
  thuRows.push({
    soPhieu: "CD-01/07", ngayCT: new Date(2026, 6, 22), ngayHT: new Date(2026, 6, 22),
    nguoi: "", noiDung: "Nhận tiền KPCĐ cấp trên cấp Tháng 06/2026", soTien: 3930675,
    loaiQuy: "Quỹ công đoàn", tkDoiUng: "", maNV: "",
    chiNhanh: "", ghiChu: "Nhập từ Sổ Quỹ Công đoàn (Excel cũ)"
  });
  thuRows.push({
    soPhieu: "CD-02/07", ngayCT: new Date(2026, 6, 22), ngayHT: new Date(2026, 6, 22),
    nguoi: "", noiDung: "Nhận tiền ĐPCĐ Tháng 06 nhập quỹ", soTien: 1310225,
    loaiQuy: "Quỹ công đoàn", tkDoiUng: "", maNV: "",
    chiNhanh: "", ghiChu: "Nhập từ Sổ Quỹ Công đoàn (Excel cũ)"
  });
  thuRows.push({
    soPhieu: "CD-01/08", ngayCT: new Date(2026, 7, 28), ngayHT: new Date(2026, 7, 28),
    nguoi: "", noiDung: "Nhận tiền KPCĐ cấp trên cấp tháng 07/2026", soTien: 4005225,
    loaiQuy: "Quỹ công đoàn", tkDoiUng: "", maNV: "",
    chiNhanh: "", ghiChu: "Nhập từ Sổ Quỹ Công đoàn (Excel cũ)"
  });
  thuRows.push({
    soPhieu: "CD-02/08", ngayCT: new Date(2026, 7, 28), ngayHT: new Date(2026, 7, 28),
    nguoi: "", noiDung: "Nhận tiền ĐPCĐ T07/2026 nhập quỹ CĐ", soTien: 1335075,
    loaiQuy: "Quỹ công đoàn", tkDoiUng: "", maNV: "",
    chiNhanh: "", ghiChu: "Nhập từ Sổ Quỹ Công đoàn (Excel cũ)"
  });

  const chiRows = [];
  chiRows.push({
    soPhieu: "01/09.26", ngayCT: new Date(2026, 8, 1), ngayHT: new Date(2026, 8, 1),
    nguoi: "NGÔ THỊ KIỀU NƯƠNG", noiDung: "Chi tiền thưởng lễ 2/9 cho Nguyên + Hằng", soTien: 600000,
    loaiQuy: "Quỹ tiền mặt", tkDoiUng: "3348", maNV: "01.006.001",
    chiNhanh: "", ghiChu: "Nhập từ Sổ Quỹ tổng (Excel cũ)"
  });
  chiRows.push({
    soPhieu: "02/09.26", ngayCT: new Date(2026, 8, 1), ngayHT: new Date(2026, 8, 1),
    nguoi: "ĐINH TRẦN KHƯƠNG", noiDung: "Chi tiền bơm 4 lốp xe 22841", soTien: 20000,
    loaiQuy: "Quỹ tiền mặt", tkDoiUng: "6426", maNV: "01.009.001",
    chiNhanh: "", ghiChu: "Nhập từ Sổ Quỹ tổng (Excel cũ)"
  });
  chiRows.push({
    soPhieu: "03/09.26", ngayCT: new Date(2026, 8, 1), ngayHT: new Date(2026, 8, 1),
    nguoi: "ĐINH TRẦN KHƯƠNG", noiDung: "Chi tiền anh Khương mua VT ( đồng hồ nhiệt xe xúc lật)", soTien: 380000,
    loaiQuy: "Quỹ tiền mặt", tkDoiUng: "6273", maNV: "01.009.001",
    chiNhanh: "", ghiChu: "Nhập từ Sổ Quỹ tổng (Excel cũ)"
  });
  chiRows.push({
    soPhieu: "04/09.26", ngayCT: new Date(2026, 8, 1), ngayHT: new Date(2026, 8, 1),
    nguoi: "NGUYỄN MỘC QUẾ ANH", noiDung: "Chi tiền HT Quế Anh làm chứng từ keo Quế Cường", soTien: 1200000,
    loaiQuy: "Quỹ tiền mặt", tkDoiUng: "6426", maNV: "01.005.001",
    chiNhanh: "", ghiChu: "Nhập từ Sổ Quỹ tổng (Excel cũ)"
  });
  chiRows.push({
    soPhieu: "05/09.26", ngayCT: new Date(2026, 8, 1), ngayHT: new Date(2026, 8, 1),
    nguoi: "VÕ VĂN PHỤNG", noiDung: "Chi tiền cho sếp Phụng", soTien: 15000000,
    loaiQuy: "Quỹ tiền mặt", tkDoiUng: "6426", maNV: "01.001.001",
    chiNhanh: "", ghiChu: "Nhập từ Sổ Quỹ tổng (Excel cũ)"
  });
  chiRows.push({
    soPhieu: "06/09.26", ngayCT: new Date(2026, 8, 3), ngayHT: new Date(2026, 8, 3),
    nguoi: "ĐẶNG CÔNG NGHĨA", noiDung: "Chi tiền đổ xăng xe 43616", soTien: 1500000,
    loaiQuy: "Quỹ tiền mặt", tkDoiUng: "33111", maNV: "04.006.002",
    chiNhanh: "", ghiChu: "Nhập từ Sổ Quỹ tổng (Excel cũ)"
  });
  chiRows.push({
    soPhieu: "07/09.26", ngayCT: new Date(2026, 8, 5), ngayHT: new Date(2026, 8, 5),
    nguoi: "NGÔ THỊ KIỀU NƯƠNG", noiDung: "Chi tiền bd tổ công nhân Đức", soTien: 50000,
    loaiQuy: "Quỹ tiền mặt", tkDoiUng: "6426", maNV: "01.006.001",
    chiNhanh: "", ghiChu: "Nhập từ Sổ Quỹ tổng (Excel cũ)"
  });
  chiRows.push({
    soPhieu: "08/09.26", ngayCT: new Date(2026, 8, 7), ngayHT: new Date(2026, 8, 7),
    nguoi: "ĐINH TRẦN KHƯƠNG", noiDung: "Chi tiền anh Khương mua VT ( ổ cắm, nối hơi)", soTien: 220000,
    loaiQuy: "Quỹ tiền mặt", tkDoiUng: "6273", maNV: "01.009.001",
    chiNhanh: "", ghiChu: "Nhập từ Sổ Quỹ tổng (Excel cũ)"
  });
  chiRows.push({
    soPhieu: "09/09.26", ngayCT: new Date(2026, 8, 7), ngayHT: new Date(2026, 8, 7),
    nguoi: "NGUYỄN THỊ NGA", noiDung: "Chi tiền chị Nga mua đồ vệ sinh VP", soTien: 383000,
    loaiQuy: "Quỹ tiền mặt", tkDoiUng: "33111", maNV: "03.004.006",
    chiNhanh: "", ghiChu: "Nhập từ Sổ Quỹ tổng (Excel cũ)"
  });
  chiRows.push({
    soPhieu: "10/09.26", ngayCT: new Date(2026, 8, 7), ngayHT: new Date(2026, 8, 7),
    nguoi: "NGUYỄN THỊ HUỲNH GIANG", noiDung: "Chi tiền căn tin ứng cơm đợt 1 T09", soTien: 5000000,
    loaiQuy: "Quỹ tiền mặt", tkDoiUng: "3348", maNV: "",
    chiNhanh: "", ghiChu: "Nhập từ Sổ Quỹ tổng (Excel cũ)"
  });
  chiRows.push({
    soPhieu: "11/09.26", ngayCT: new Date(2026, 8, 7), ngayHT: new Date(2026, 8, 7),
    nguoi: "NGUYỄN THỊ HUỲNH GIANG", noiDung: "Chi tiền cơm căn tin T09/2026", soTien: 6100000,
    loaiQuy: "Quỹ tiền mặt", tkDoiUng: "3348", maNV: "",
    chiNhanh: "", ghiChu: "Nhập từ Sổ Quỹ tổng (Excel cũ)"
  });
  chiRows.push({
    soPhieu: "12/09.26", ngayCT: new Date(2026, 8, 7), ngayHT: new Date(2026, 8, 7),
    nguoi: "NGUYỄN THỊ HUỲNH GIANG", noiDung: "Chi tiền cơm sếp", soTien: 190000,
    loaiQuy: "Quỹ tiền mặt", tkDoiUng: "6426", maNV: "",
    chiNhanh: "", ghiChu: "Nhập từ Sổ Quỹ tổng (Excel cũ)"
  });
  chiRows.push({
    soPhieu: "13/09.26", ngayCT: new Date(2026, 8, 7), ngayHT: new Date(2026, 8, 7),
    nguoi: "ĐẶNG CÔNG NGHĨA", noiDung: "Chi tiền đổ xăng xe 43616", soTien: 1500000,
    loaiQuy: "Quỹ tiền mặt", tkDoiUng: "33111", maNV: "04.006.002",
    chiNhanh: "", ghiChu: "Nhập từ Sổ Quỹ tổng (Excel cũ)"
  });
  chiRows.push({
    soPhieu: "14/09.26", ngayCT: new Date(2026, 8, 7), ngayHT: new Date(2026, 8, 7),
    nguoi: "NGÔ THỊ KIỀU NƯƠNG", noiDung: "Chi tiền lương tiền mặt T08/2026", soTien: 17460000,
    loaiQuy: "Quỹ tiền mặt", tkDoiUng: "3341", maNV: "01.006.001",
    chiNhanh: "", ghiChu: "Nhập từ Sổ Quỹ tổng (Excel cũ)"
  });
  chiRows.push({
    soPhieu: "15/09.26", ngayCT: new Date(2026, 8, 7), ngayHT: new Date(2026, 8, 7),
    nguoi: "VÕ VĂN PHỤNG", noiDung: "Chi tiền cho sếp Phụng", soTien: 10000000,
    loaiQuy: "Quỹ tiền mặt", tkDoiUng: "6426", maNV: "01.001.001",
    chiNhanh: "", ghiChu: "Nhập từ Sổ Quỹ tổng (Excel cũ)"
  });
  chiRows.push({
    soPhieu: "16/09.26", ngayCT: new Date(2026, 8, 7), ngayHT: new Date(2026, 8, 7),
    nguoi: "DƯƠNG THỊ THANH NGA", noiDung: "Chi tiền cho Quế Sơn", soTien: 100000000,
    loaiQuy: "Quỹ tiền mặt", tkDoiUng: "1361", maNV: "",
    chiNhanh: "", ghiChu: "Nhập từ Sổ Quỹ tổng (Excel cũ)"
  });
  chiRows.push({
    soPhieu: "17/09.26", ngayCT: new Date(2026, 8, 8), ngayHT: new Date(2026, 8, 8),
    nguoi: "VÕ VĂN PHỤNG", noiDung: "Chi tiền trả mượn của sếp Phụng ngày 21/08 ( cơn nợ 100.000.000)", soTien: 100000000,
    loaiQuy: "Quỹ tiền mặt", tkDoiUng: "1388NB", maNV: "01.001.001",
    chiNhanh: "", ghiChu: "Nhập từ Sổ Quỹ tổng (Excel cũ)"
  });
  chiRows.push({
    soPhieu: "18/09.26", ngayCT: new Date(2026, 8, 8), ngayHT: new Date(2026, 8, 8),
    nguoi: "ĐINH THỊ THANH SƯƠNG", noiDung: "Chi tiền phí cầu đường vào tk VETC HAK", soTien: 500000,
    loaiQuy: "Quỹ tiền mặt", tkDoiUng: "33111", maNV: "01.006.002",
    chiNhanh: "", ghiChu: "Nhập từ Sổ Quỹ tổng (Excel cũ)"
  });
  chiRows.push({
    soPhieu: "19/09.26", ngayCT: new Date(2026, 8, 9), ngayHT: new Date(2026, 8, 9),
    nguoi: "NGUYỄN THỊ NGA", noiDung: "Chi tiền mua xăng dọn bãi", soTien: 120000,
    loaiQuy: "Quỹ tiền mặt", tkDoiUng: "6426", maNV: "03.004.006",
    chiNhanh: "", ghiChu: "Nhập từ Sổ Quỹ tổng (Excel cũ)"
  });
  chiRows.push({
    soPhieu: "20/09.26", ngayCT: new Date(2026, 8, 9), ngayHT: new Date(2026, 8, 9),
    nguoi: "ĐINH TRẦN KHƯƠNG", noiDung: "Chi tiền anh Khương mua VT ( bulong 20 -8, bộ phốt)", soTien: 450000,
    loaiQuy: "Quỹ tiền mặt", tkDoiUng: "6273", maNV: "01.009.001",
    chiNhanh: "", ghiChu: "Nhập từ Sổ Quỹ tổng (Excel cũ)"
  });
  chiRows.push({
    soPhieu: "21/09.26", ngayCT: new Date(2026, 8, 9), ngayHT: new Date(2026, 8, 9),
    nguoi: "ANH TUÂN", noiDung: "Chi tiền sửa camera ( thay nguồn do ko ghi lại được)", soTien: 350000,
    loaiQuy: "Quỹ tiền mặt", tkDoiUng: "6426", maNV: "",
    chiNhanh: "", ghiChu: "Nhập từ Sổ Quỹ tổng (Excel cũ)"
  });
  chiRows.push({
    soPhieu: "22/09.26", ngayCT: new Date(2026, 8, 9), ngayHT: new Date(2026, 8, 9),
    nguoi: "VÕ THỊ NHƯ NGỌC", noiDung: "Chi tiền trả sếp Ngọc", soTien: 1334000000,
    loaiQuy: "Quỹ tiền mặt", tkDoiUng: "33882", maNV: "13884A.007",
    chiNhanh: "", ghiChu: "Nhập từ Sổ Quỹ tổng (Excel cũ)"
  });
  chiRows.push({
    soPhieu: "23/09.26", ngayCT: new Date(2026, 8, 9), ngayHT: new Date(2026, 8, 9),
    nguoi: "TRƯƠNG QUỐC THÀNH", noiDung: "Chi tiền làm bảng nhựa ( nội dung nơi cất đồ vệ sinh)", soTien: 387000,
    loaiQuy: "Quỹ tiền mặt", tkDoiUng: "33111", maNV: "04.003.001",
    chiNhanh: "", ghiChu: "Nhập từ Sổ Quỹ tổng (Excel cũ)"
  });
  chiRows.push({
    soPhieu: "24/09.26", ngayCT: new Date(2026, 8, 10), ngayHT: new Date(2026, 8, 10),
    nguoi: "NGUYỄN THỊ NGA", noiDung: "Chi tiền mua đồ cúng mùng 1", soTien: 341000,
    loaiQuy: "Quỹ tiền mặt", tkDoiUng: "6426", maNV: "03.004.006",
    chiNhanh: "", ghiChu: "Nhập từ Sổ Quỹ tổng (Excel cũ)"
  });
  chiRows.push({
    soPhieu: "CD-01/01", ngayCT: new Date(2026, 0, 7), ngayHT: new Date(2022, 0, 6),
    nguoi: "", noiDung: "Chi tiền thăm ốm cô khương mổ mắt", soTien: 500000,
    loaiQuy: "Quỹ công đoàn", tkDoiUng: "", maNV: "",
    chiNhanh: "", ghiChu: "Nhập từ Sổ Quỹ Công đoàn (Excel cũ)"
  });
  chiRows.push({
    soPhieu: "CD-02/01", ngayCT: new Date(2026, 0, 30), ngayHT: new Date(2026, 0, 30),
    nguoi: "", noiDung: "Chi tiền thăm ốm chú trương cơ khí nhập viện 21-27", soTien: 500000,
    loaiQuy: "Quỹ công đoàn", tkDoiUng: "", maNV: "",
    chiNhanh: "", ghiChu: "Nhập từ Sổ Quỹ Công đoàn (Excel cũ)"
  });
  chiRows.push({
    soPhieu: "CD-01/02", ngayCT: new Date(2026, 1, 5), ngayHT: new Date(2026, 1, 5),
    nguoi: "", noiDung: "Chi tiền quà tết CĐ CNV năm 2025                     ( 800.000/người)", soTien: 48800000,
    loaiQuy: "Quỹ công đoàn", tkDoiUng: "", maNV: "",
    chiNhanh: "", ghiChu: "Nhập từ Sổ Quỹ Công đoàn (Excel cũ)"
  });
  chiRows.push({
    soPhieu: "CD-01/03", ngayCT: new Date(2026, 2, 7), ngayHT: new Date(2026, 2, 7),
    nguoi: "", noiDung: "Chi tiền lễ 8/3 cho cnv nữ", soTien: 1600000,
    loaiQuy: "Quỹ công đoàn", tkDoiUng: "", maNV: "",
    chiNhanh: "", ghiChu: "Nhập từ Sổ Quỹ Công đoàn (Excel cũ)"
  });
  chiRows.push({
    soPhieu: "CD-02/03", ngayCT: new Date(2026, 2, 12), ngayHT: new Date(2026, 2, 12),
    nguoi: "", noiDung: "Chi tiền nộp TK CĐ để duy trì số dư", soTien: 500000,
    loaiQuy: "Quỹ công đoàn", tkDoiUng: "", maNV: "",
    chiNhanh: "", ghiChu: "Nhập từ Sổ Quỹ Công đoàn (Excel cũ)"
  });
  chiRows.push({
    soPhieu: "CD-03/03", ngayCT: new Date(2026, 2, 12), ngayHT: new Date(2026, 2, 12),
    nguoi: "", noiDung: "Chi tiền nộp ĐPCĐ T01+02 lên cấp trên", soTien: 872820,
    loaiQuy: "Quỹ công đoàn", tkDoiUng: "", maNV: "",
    chiNhanh: "", ghiChu: "Nhập từ Sổ Quỹ Công đoàn (Excel cũ)"
  });
  chiRows.push({
    soPhieu: "CD-04/03", ngayCT: new Date(2026, 2, 25), ngayHT: new Date(2026, 2, 25),
    nguoi: "", noiDung: "Chi tiền đi đám ma mẹ chị Lệ CN", soTien: 500000,
    loaiQuy: "Quỹ công đoàn", tkDoiUng: "", maNV: "",
    chiNhanh: "", ghiChu: "Nhập từ Sổ Quỹ Công đoàn (Excel cũ)"
  });
  chiRows.push({
    soPhieu: "CD-01/04", ngayCT: new Date(2026, 3, 13), ngayHT: new Date(2026, 3, 13),
    nguoi: "", noiDung: "Chi tiền nộp ĐPCĐ T03 lên cấp trên", soTien: 440033,
    loaiQuy: "Quỹ công đoàn", tkDoiUng: "", maNV: "",
    chiNhanh: "", ghiChu: "Nhập từ Sổ Quỹ Công đoàn (Excel cũ)"
  });
  chiRows.push({
    soPhieu: "CD-02/04", ngayCT: new Date(2026, 3, 16), ngayHT: new Date(2026, 3, 16),
    nguoi: "", noiDung: "Chi tiền thăm ốm chị Lý CN", soTien: 500000,
    loaiQuy: "Quỹ công đoàn", tkDoiUng: "", maNV: "",
    chiNhanh: "", ghiChu: "Nhập từ Sổ Quỹ Công đoàn (Excel cũ)"
  });
  chiRows.push({
    soPhieu: "CD-01/05", ngayCT: new Date(2026, 4, 4), ngayHT: new Date(2026, 4, 4),
    nguoi: "", noiDung: "Chi thăm ốm anh Đạt TX", soTien: 500000,
    loaiQuy: "Quỹ công đoàn", tkDoiUng: "", maNV: "",
    chiNhanh: "", ghiChu: "Nhập từ Sổ Quỹ Công đoàn (Excel cũ)"
  });
  chiRows.push({
    soPhieu: "CD-02/05", ngayCT: new Date(2026, 4, 16), ngayHT: new Date(2026, 4, 16),
    nguoi: "", noiDung: "Chi tiền thăm ốm Hà Điền Hưng TX", soTien: 500000,
    loaiQuy: "Quỹ công đoàn", tkDoiUng: "", maNV: "",
    chiNhanh: "", ghiChu: "Nhập từ Sổ Quỹ Công đoàn (Excel cũ)"
  });
  chiRows.push({
    soPhieu: "CD-03/05", ngayCT: new Date(2026, 4, 18), ngayHT: new Date(2026, 4, 18),
    nguoi: "", noiDung: "Chi tiền nộp ĐPCĐ T04/2026 lên cấp trên", soTien: 416850,
    loaiQuy: "Quỹ công đoàn", tkDoiUng: "", maNV: "",
    chiNhanh: "", ghiChu: "Nhập từ Sổ Quỹ Công đoàn (Excel cũ)"
  });
  chiRows.push({
    soPhieu: "CD-04/05", ngayCT: new Date(2026, 4, 27), ngayHT: new Date(2026, 4, 27),
    nguoi: "", noiDung: "Chi tiền đi đám ma mẹ anh Khương", soTien: 500000,
    loaiQuy: "Quỹ công đoàn", tkDoiUng: "", maNV: "",
    chiNhanh: "", ghiChu: "Nhập từ Sổ Quỹ Công đoàn (Excel cũ)"
  });
  chiRows.push({
    soPhieu: "CD-05/05", ngayCT: new Date(2026, 4, 27), ngayHT: new Date(2026, 4, 27),
    nguoi: "", noiDung: "Chi tiền quà Tết thiếu nhi (01/06) 100.000/người x 55 người", soTien: 5500000,
    loaiQuy: "Quỹ công đoàn", tkDoiUng: "", maNV: "",
    chiNhanh: "", ghiChu: "Nhập từ Sổ Quỹ Công đoàn (Excel cũ)"
  });
  chiRows.push({
    soPhieu: "CD-01/06", ngayCT: new Date(2026, 5, 8), ngayHT: new Date(2026, 5, 8),
    nguoi: "", noiDung: "Chi tiền thăm chị Lý công nhân", soTien: 500000,
    loaiQuy: "Quỹ công đoàn", tkDoiUng: "", maNV: "",
    chiNhanh: "", ghiChu: "Nhập từ Sổ Quỹ Công đoàn (Excel cũ)"
  });
  chiRows.push({
    soPhieu: "CD-02/06", ngayCT: new Date(2026, 5, 8), ngayHT: new Date(2026, 5, 8),
    nguoi: "", noiDung: "Chi tiền nộp ĐPCĐ T05/2026 lên cấp trên", soTien: 416460,
    loaiQuy: "Quỹ công đoàn", tkDoiUng: "", maNV: "",
    chiNhanh: "", ghiChu: "Nhập từ Sổ Quỹ Công đoàn (Excel cũ)"
  });
  chiRows.push({
    soPhieu: "CD-01/07", ngayCT: new Date(2026, 6, 22), ngayHT: new Date(2026, 6, 22),
    nguoi: "", noiDung: "Chi tiền nộp ĐPCĐ T06/2026 lên cấp trên", soTien: 393068,
    loaiQuy: "Quỹ công đoàn", tkDoiUng: "", maNV: "",
    chiNhanh: "", ghiChu: "Nhập từ Sổ Quỹ Công đoàn (Excel cũ)"
  });
  chiRows.push({
    soPhieu: "CD-01/08", ngayCT: new Date(2026, 7, 13), ngayHT: new Date(2026, 7, 13),
    nguoi: "", noiDung: "Chi tiền thăm Quang CN tổ Đức bị tai nạn", soTien: 500000,
    loaiQuy: "Quỹ công đoàn", tkDoiUng: "", maNV: "",
    chiNhanh: "", ghiChu: "Nhập từ Sổ Quỹ Công đoàn (Excel cũ)"
  });
  chiRows.push({
    soPhieu: "CD-02/08", ngayCT: new Date(2026, 7, 24), ngayHT: new Date(2026, 7, 24),
    nguoi: "", noiDung: "Chi tiền thăm chị Hạnh công nhân nữ", soTien: 500000,
    loaiQuy: "Quỹ công đoàn", tkDoiUng: "", maNV: "",
    chiNhanh: "", ghiChu: "Nhập từ Sổ Quỹ Công đoàn (Excel cũ)"
  });
  chiRows.push({
    soPhieu: "CD-03/08", ngayCT: new Date(2026, 7, 28), ngayHT: new Date(2026, 7, 28),
    nguoi: "", noiDung: "Chi tiền nộp ĐPCĐ T07/2026 lên cấp trên", soTien: 400523,
    loaiQuy: "Quỹ công đoàn", tkDoiUng: "", maNV: "",
    chiNhanh: "", ghiChu: "Nhập từ Sổ Quỹ Công đoàn (Excel cũ)"
  });

  const comRows = [];
  comRows.push({
    ngay: new Date(2026, 8, 1), trua: 23, toi: 18, tong: 41, dg: 25000, thanhTien: 1025000,
    ngayTamUng: null, soTamUng: 0
  });
  comRows.push({
    ngay: new Date(2026, 8, 2), trua: 0, toi: 0, tong: 0, dg: 25000, thanhTien: 0,
    ngayTamUng: null, soTamUng: 0
  });
  comRows.push({
    ngay: new Date(2026, 8, 3), trua: 23, toi: 16, tong: 39, dg: 25000, thanhTien: 975000,
    ngayTamUng: null, soTamUng: 0
  });
  comRows.push({
    ngay: new Date(2026, 8, 4), trua: 25, toi: 19, tong: 44, dg: 25000, thanhTien: 1100000,
    ngayTamUng: null, soTamUng: 0
  });
  comRows.push({
    ngay: new Date(2026, 8, 5), trua: 25, toi: 18, tong: 43, dg: 25000, thanhTien: 1075000,
    ngayTamUng: null, soTamUng: 0
  });
  comRows.push({
    ngay: new Date(2026, 8, 6), trua: 21, toi: 14, tong: 35, dg: 25000, thanhTien: 875000,
    ngayTamUng: null, soTamUng: 0
  });
  comRows.push({
    ngay: new Date(2026, 8, 7), trua: 23, toi: 18, tong: 41, dg: 25000, thanhTien: 1025000,
    ngayTamUng: new Date(2026, 8, 7), soTamUng: 5000000
  });
  comRows.push({
    ngay: new Date(2026, 8, 8), trua: 26, toi: 18, tong: 44, dg: 25000, thanhTien: 1100000,
    ngayTamUng: null, soTamUng: 0
  });
  comRows.push({
    ngay: new Date(2026, 8, 9), trua: 26, toi: 19, tong: 45, dg: 25000, thanhTien: 1125000,
    ngayTamUng: null, soTamUng: 0
  });
  comRows.push({
    ngay: new Date(2026, 8, 10), trua: 25, toi: 20, tong: 45, dg: 25000, thanhTien: 1125000,
    ngayTamUng: null, soTamUng: 0
  });

  let tsSeq = 0;
  function nextTs(ngayCT) {
    tsSeq++;
    const base = new Date(ngayCT.getFullYear(), ngayCT.getMonth(), ngayCT.getDate(), 8, 0, 0);
    return new Date(base.getTime() + tsSeq * 1000);
  }

  thuRows.forEach(function (r, idx) {
    const ts = nextTs(r.ngayCT);
    shThu.appendRow([
      "PT_IMPORT2_" + String(idx + 1).padStart(4, "0"), r.soPhieu, r.ngayCT, "", "Thu thường",
      r.nguoi, "", r.noiDung, "", r.soTien, "", r.ghiChu,
      nguoiLap, ts, "", "", "Hợp lệ", "",
      r.loaiQuy, r.ngayHT, "1111", r.tkDoiUng, r.maNV, r.chiNhanh
    ]);
  });

  chiRows.forEach(function (r, idx) {
    const ts = nextTs(r.ngayCT);
    shChi.appendRow([
      "PC_IMPORT2_" + String(idx + 1).padStart(4, "0"), r.soPhieu, r.ngayCT, "", "Chi thường",
      r.nguoi, "", r.noiDung, "", r.soTien, "", r.ghiChu,
      nguoiLap, ts, "", "", "Hợp lệ", "",
      r.loaiQuy, r.ngayHT, "1111", r.tkDoiUng, r.maNV, r.chiNhanh
    ]);
  });

  comRows.forEach(function (r, idx) {
    const ts = new Date(r.ngay.getFullYear(), r.ngay.getMonth(), r.ngay.getDate(), 8, 1, idx);
    shCom.appendRow([
      "COM_IMPORT2_" + String(idx + 1).padStart(2, "0"), r.ngay, r.trua, r.toi, r.tong, r.dg, r.thanhTien,
      r.ngayTamUng, r.soTamUng, nguoiLap, ts, ""
    ]);
  });

  SpreadsheetApp.getUi().alert(
    "Đã nhập xong dữ liệu sổ cũ!\n\n" +
    "Phiếu Thu: " + thuRows.length + "\n" +
    "Phiếu Chi: " + chiRows.length + "\n" +
    "Sổ Cơm: " + comRows.length + " ngày\n\n" +
    "Vào Sổ Quỹ / Sổ Quỹ Công Đoàn / Sổ Cơm để kiểm tra lại."
  );
}
