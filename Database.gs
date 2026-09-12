/*************************************************
 * DATABASE.GS
 * Tạo cấu trúc Database ban đầu (12 sheet).
 *
 * CÁCH CHẠY: Trong Apps Script Editor, chọn hàm
 * "setupDatabase" ở dropdown trên cùng > bấm Run.
 * Chỉ chạy 1 LẦN DUY NHẤT lúc khởi tạo — chạy lại sẽ
 * XÓA SẠCH dữ liệu cũ trong các sheet đã tồn tại.
 *************************************************/

function setupDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  _setupUsersSheet(ss);
  _setupDoiTuongSheet(ss);
  _setupLoaiThuSheet(ss);
  _setupLoaiChiSheet(ss);
  _setupPhieuThuSheet(ss);
  _setupPhieuChiSheet(ss);
  _setupCongNoSheet(ss);
  _setupThanhToanCongNoSheet(ss);
  _setupKiemKeQuySheet(ss);
  _setupKhoaSoSheet(ss);
  _setupAuditLogSheet(ss);
  _setupCauHinhSheet(ss);
  _setupSoComSheet(ss);

  const defaultSheet = ss.getSheetByName('Sheet1');
  if (defaultSheet && ss.getSheets().length > 1) {
    ss.deleteSheet(defaultSheet);
  }

  SpreadsheetApp.getUi().alert(
    'Đã tạo xong Database (13 sheet)!\n\n' +
    'Tài khoản đăng nhập mẫu:\n' +
    'Username: admin\n' +
    'Password: admin123\n\n' +
    'Hãy đổi mật khẩu này sau khi triển khai thật.'
  );
}

function _getOrCreateSheet(ss, name) {
  let sh = ss.getSheetByName(name);
  if (sh) {
    sh.clear();
  } else {
    sh = ss.insertSheet(name);
  }
  return sh;
}

function _styleHeader(sh, numCols) {
  const headerRange = sh.getRange(1, 1, 1, numCols);
  headerRange.setBackground('#1e40af');
  headerRange.setFontColor('#ffffff');
  headerRange.setFontWeight('bold');
  sh.setFrozenRows(1);
  sh.autoResizeColumns(1, numCols);
}

/* ---------- 1. USERS ---------- */
function _setupUsersSheet(ss) {
  const sh = _getOrCreateSheet(ss, SHEET_USERS);
  const headers = ['user_id', 'username', 'password_hash', 'full_name', 'role', 'status', 'created_at'];
  sh.getRange(1, 1, 1, headers.length).setValues([headers]);

  const adminHash = _hashPassword('admin123');
  sh.getRange(2, 1, 1, headers.length).setValues([[
    'USR001', 'admin', adminHash, 'Quản trị viên', ROLE_ADMIN, USER_STATUS_ACTIVE, new Date()
  ]]);

  _styleHeader(sh, headers.length);
}

/* ---------- 2. DM_DOITUONG ---------- */
function _setupDoiTuongSheet(ss) {
  const sh = _getOrCreateSheet(ss, SHEET_DOITUONG);
  const headers = ['ma_doi_tuong', 'ten_doi_tuong', 'loai_doi_tuong', 'so_dien_thoai', 'dia_chi', 'mst', 'nguoi_lien_he', 'ghi_chu', 'trang_thai'];
  sh.getRange(1, 1, 1, headers.length).setValues([headers]);

  const sample = [
    ['DT001', 'Nguyễn Văn A', 'Khách hàng', '0901234567', '', '', '', '', 'Hoạt động'],
    ['DT002', 'Công ty TNHH B', 'Nhà cung cấp', '0912345678', '', '', '', '', 'Hoạt động']
  ];
  sh.getRange(2, 1, sample.length, headers.length).setValues(sample);

  _styleHeader(sh, headers.length);
}

/* ---------- 3. DM_LOAI_THU ---------- */
function _setupLoaiThuSheet(ss) {
  const sh = _getOrCreateSheet(ss, SHEET_LOAI_THU);
  const headers = ['ma_loai_thu', 'ten_loai_thu', 'trang_thai'];
  sh.getRange(1, 1, 1, headers.length).setValues([headers]);

  const sample = [
    ['LT001', 'Thu bán hàng', 'Hoạt động'],
    ['LT002', 'Thu hồi công nợ', 'Hoạt động'],
    ['LT003', 'Thu khác', 'Hoạt động']
  ];
  sh.getRange(2, 1, sample.length, headers.length).setValues(sample);

  _styleHeader(sh, headers.length);
}

/* ---------- 4. DM_LOAI_CHI ---------- */
function _setupLoaiChiSheet(ss) {
  const sh = _getOrCreateSheet(ss, SHEET_LOAI_CHI);
  const headers = ['ma_loai_chi', 'ten_loai_chi', 'trang_thai'];
  sh.getRange(1, 1, 1, headers.length).setValues([headers]);

  const sample = [
    ['LC001', 'Chi mua hàng/vật tư', 'Hoạt động'],
    ['LC002', 'Trả công nợ', 'Hoạt động'],
    ['LC003', 'Chi phí vận hành', 'Hoạt động'],
    ['LC004', 'Chi khác', 'Hoạt động']
  ];
  sh.getRange(2, 1, sample.length, headers.length).setValues(sample);

  _styleHeader(sh, headers.length);
}

/* ---------- 5. PHIEU_THU ----------
 * Các cột loai_quy/ngay_hach_toan/tai_khoan/tk_doi_ung/ma_nhan_vien/chi_nhanh
 * được thêm ở CUỐI để phục vụ Sổ Quỹ theo mẫu kế toán (Ngày hạch toán,
 * Tài khoản, TK đối ứng...) và để 1 bộ Phiếu Thu/Chi có thể ghi cho
 * nhiều Sổ Quỹ khác nhau (Quỹ tiền mặt / Quỹ Công đoàn) qua cột loai_quy.
 */
function _setupPhieuThuSheet(ss) {
  const sh = _getOrCreateSheet(ss, SHEET_PHIEU_THU);
  const headers = [
    'id', 'so_phieu_thu', 'ngay_thu', 'gio_thu', 'loai_giao_dich',
    'nguoi_nop_tien', 'ma_doi_tuong', 'noi_dung_thu', 'ma_loai_thu', 'so_tien',
    'chung_tu_lien_quan', 'ghi_chu', 'nguoi_lap', 'thoi_gian_lap',
    'nguoi_sua_cuoi', 'thoi_gian_sua_cuoi', 'trang_thai', 'ly_do_huy',
    'loai_quy', 'ngay_hach_toan', 'tai_khoan', 'tk_doi_ung', 'ma_nhan_vien', 'chi_nhanh'
  ];
  sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  sh.getRange(2, 3, sh.getMaxRows() - 1, 1).setNumberFormat('yyyy-mm-dd');
  sh.getRange(2, 10, sh.getMaxRows() - 1, 1).setNumberFormat('#,##0');
  sh.getRange(2, 20, sh.getMaxRows() - 1, 1).setNumberFormat('yyyy-mm-dd');
  _styleHeader(sh, headers.length);
}

/* ---------- 6. PHIEU_CHI ---------- */
function _setupPhieuChiSheet(ss) {
  const sh = _getOrCreateSheet(ss, SHEET_PHIEU_CHI);
  const headers = [
    'id', 'so_phieu_chi', 'ngay_chi', 'gio_chi', 'loai_giao_dich',
    'nguoi_nhan_tien', 'ma_doi_tuong', 'noi_dung_chi', 'ma_loai_chi', 'so_tien',
    'chung_tu_lien_quan', 'ghi_chu', 'nguoi_lap', 'thoi_gian_lap',
    'nguoi_sua_cuoi', 'thoi_gian_sua_cuoi', 'trang_thai', 'ly_do_huy',
    'loai_quy', 'ngay_hach_toan', 'tai_khoan', 'tk_doi_ung', 'ma_nhan_vien', 'chi_nhanh'
  ];
  sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  sh.getRange(2, 3, sh.getMaxRows() - 1, 1).setNumberFormat('yyyy-mm-dd');
  sh.getRange(2, 10, sh.getMaxRows() - 1, 1).setNumberFormat('#,##0');
  sh.getRange(2, 20, sh.getMaxRows() - 1, 1).setNumberFormat('yyyy-mm-dd');
  _styleHeader(sh, headers.length);
}

/* ---------- 7. CONG_NO ---------- */
function _setupCongNoSheet(ss) {
  const sh = _getOrCreateSheet(ss, SHEET_CONG_NO);
  const headers = [
    'ma_cong_no', 'ngay_phat_sinh', 'ma_doi_tuong', 'loai_cong_no', 'noi_dung',
    'so_tien_phat_sinh', 'han_thanh_toan', 'da_thanh_toan', 'con_lai',
    'trang_thai', 'nguoi_lap', 'thoi_gian_lap'
  ];
  sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  sh.getRange(2, 2, sh.getMaxRows() - 1, 1).setNumberFormat('yyyy-mm-dd');
  sh.getRange(2, 6, sh.getMaxRows() - 1, 4).setNumberFormat('#,##0');
  _styleHeader(sh, headers.length);
}

/* ---------- 8. THANH_TOAN_CONG_NO ---------- */
function _setupThanhToanCongNoSheet(ss) {
  const sh = _getOrCreateSheet(ss, SHEET_THANH_TOAN_CONG_NO);
  const headers = [
    'id', 'ma_cong_no', 'ngay_thanh_toan', 'so_tien_thanh_toan', 'loai',
    'so_phieu_lien_quan', 'nguoi_lap', 'thoi_gian_lap'
  ];
  sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  sh.getRange(2, 3, sh.getMaxRows() - 1, 1).setNumberFormat('yyyy-mm-dd');
  sh.getRange(2, 4, sh.getMaxRows() - 1, 1).setNumberFormat('#,##0');
  _styleHeader(sh, headers.length);
}

/* ---------- 9. KIEM_KE_QUY ---------- */
function _setupKiemKeQuySheet(ss) {
  const sh = _getOrCreateSheet(ss, SHEET_KIEM_KE_QUY);
  const headers = [
    'id', 'ngay_kiem_ke', 'ton_he_thong', 'tien_thuc_te', 'chenh_lech',
    'ly_do', 'nguoi_kiem_ke', 'thoi_gian_lap'
  ];
  sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  sh.getRange(2, 2, sh.getMaxRows() - 1, 1).setNumberFormat('yyyy-mm-dd');
  sh.getRange(2, 3, sh.getMaxRows() - 1, 3).setNumberFormat('#,##0');
  _styleHeader(sh, headers.length);
}

/* ---------- 10. KHOA_SO ---------- */
function _setupKhoaSoSheet(ss) {
  const sh = _getOrCreateSheet(ss, SHEET_KHOA_SO);
  const headers = [
    'ngay_khoa', 'nguoi_khoa', 'thoi_gian_khoa', 'trang_thai',
    'nguoi_mo_khoa', 'thoi_gian_mo_khoa', 'ly_do_mo_khoa'
  ];
  sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  sh.getRange(2, 1, sh.getMaxRows() - 1, 1).setNumberFormat('yyyy-mm-dd');
  _styleHeader(sh, headers.length);
}

/* ---------- 11. AUDIT_LOG ---------- */
function _setupAuditLogSheet(ss) {
  const sh = _getOrCreateSheet(ss, SHEET_AUDIT_LOG);
  const headers = [
    'id', 'thoi_gian', 'nguoi_thao_tac', 'chuc_nang', 'loai_thao_tac',
    'ma_chung_tu', 'du_lieu_truoc', 'du_lieu_sau'
  ];
  sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  _styleHeader(sh, headers.length);
}

/* ---------- 12. CAU_HINH ---------- */
function _setupCauHinhSheet(ss) {
  const sh = _getOrCreateSheet(ss, SHEET_CAU_HINH);
  const headers = ['key', 'value', 'ghi_chu'];
  sh.getRange(1, 1, 1, headers.length).setValues([headers]);

  const sample = [
    ['SO_PHIEU_THU_TIEP_THEO', 1, 'Số thứ tự Phiếu Thu kế tiếp (dùng để sinh PT000001...)'],
    ['SO_PHIEU_CHI_TIEP_THEO', 1, 'Số thứ tự Phiếu Chi kế tiếp (dùng để sinh PC000001...)'],
    ['SO_DU_QUY_KHOI_TAO', 0, 'Số dư quỹ tiền mặt tại thời điểm bắt đầu dùng hệ thống'],
    ['SO_DU_QUY_CONG_DOAN_KHOI_TAO', 0, 'Số dư Quỹ Công đoàn tại thời điểm bắt đầu dùng hệ thống'],
    ['SO_DU_QUY_COM_KHOI_TAO', 0, 'Số dư (tạm ứng còn lại) Quỹ Cơm tại thời điểm bắt đầu dùng hệ thống'],
    ['DON_GIA_COM_MAC_DINH', 25000, 'Đơn giá 1 suất ăn mặc định (đồng/suất), dùng để gợi ý khi nhập Sổ Cơm']
  ];
  sh.getRange(2, 1, sample.length, headers.length).setValues(sample);

  _styleHeader(sh, headers.length);
}

/* ---------- 13. SO_COM ----------
 * Sổ theo dõi suất ăn (trưa/tối) và tạm ứng tiền cơm hàng ngày,
 * mô phỏng theo mẫu "cơm" trong file Excel: mỗi dòng là 1 ngày,
 * có thể kèm theo 1 lần tạm ứng tiền cơm nếu ngày đó nhận tạm ứng.
 * Tồn quỹ cơm = Số dư khởi tạo + lũy kế tạm ứng - lũy kế thành tiền.
 */
function _setupSoComSheet(ss) {
  const sh = _getOrCreateSheet(ss, SHEET_SO_COM);
  const headers = [
    'id', 'ngay', 'buoi_trua', 'buoi_toi', 'tong_suat', 'don_gia', 'thanh_tien',
    'ngay_tam_ung', 'so_tien_tam_ung', 'nguoi_lap', 'thoi_gian_lap', 'ghi_chu'
  ];
  sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  sh.getRange(2, 2, sh.getMaxRows() - 1, 1).setNumberFormat('yyyy-mm-dd');
  sh.getRange(2, 6, sh.getMaxRows() - 1, 2).setNumberFormat('#,##0');
  sh.getRange(2, 8, sh.getMaxRows() - 1, 1).setNumberFormat('yyyy-mm-dd');
  sh.getRange(2, 9, sh.getMaxRows() - 1, 1).setNumberFormat('#,##0');
  _styleHeader(sh, headers.length);
}

/*************************************************
 * NÂNG CẤP KHÔNG XÓA DỮ LIỆU
 * Dùng hàm này (thay vì setupDatabase) khi hệ thống ĐÃ CÓ dữ liệu
 * thật và chỉ cần bổ sung: cột mở rộng cho Sổ Quỹ ở PHIEU_THU/
 * PHIEU_CHI, sheet SO_COM, và các dòng cấu hình mới cho Quỹ
 * Công đoàn / Quỹ Cơm. KHÔNG xóa bất kỳ dữ liệu nào đã có.
 *
 * CÁCH CHẠY: Trong Apps Script Editor, chọn hàm "migrateSoQuyMoRong"
 * ở dropdown trên cùng > bấm Run > chạy 1 lần duy nhất.
 *************************************************/
function migrateSoQuyMoRong() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cotMoi = ['loai_quy', 'ngay_hach_toan', 'tai_khoan', 'tk_doi_ung', 'ma_nhan_vien', 'chi_nhanh'];

  _themCotConThieu(ss, SHEET_PHIEU_THU, cotMoi);
  _themCotConThieu(ss, SHEET_PHIEU_CHI, cotMoi);
  _dienGiaTriMacDinhCotMoiSoQuy(ss, SHEET_PHIEU_THU, 'ngay_thu');
  _dienGiaTriMacDinhCotMoiSoQuy(ss, SHEET_PHIEU_CHI, 'ngay_chi');

  if (!ss.getSheetByName(SHEET_SO_COM)) _setupSoComSheet(ss);

  _themCauHinhNeuChua('SO_DU_QUY_CONG_DOAN_KHOI_TAO', 0, 'Số dư Quỹ Công đoàn tại thời điểm bắt đầu dùng hệ thống');
  _themCauHinhNeuChua('SO_DU_QUY_COM_KHOI_TAO', 0, 'Số dư (tạm ứng còn lại) Quỹ Cơm tại thời điểm bắt đầu dùng hệ thống');
  _themCauHinhNeuChua('DON_GIA_COM_MAC_DINH', 25000, 'Đơn giá 1 suất ăn mặc định (đồng/suất), dùng để gợi ý khi nhập Sổ Cơm');

  SpreadsheetApp.getUi().alert(
    'Đã nâng cấp xong!\n\n' +
    '- PHIEU_THU / PHIEU_CHI: đã thêm cột loai_quy, ngay_hach_toan, tai_khoan, tk_doi_ung, ma_nhan_vien, chi_nhanh.\n' +
    '- Đã tạo sheet SO_COM (nếu chưa có).\n' +
    '- Đã bổ sung cấu hình Quỹ Công đoàn / Quỹ Cơm trong CAU_HINH.\n\n' +
    'Toàn bộ dữ liệu cũ được giữ nguyên. Hãy vào sheet CAU_HINH để cập nhật số dư đầu kỳ thật của Quỹ Công đoàn / Quỹ Cơm nếu cần.'
  );
}

/*************************************************
 * NÂNG CẤP: BÁO CÁO NGÀY (KEO) + EMAIL
 * Thiết lập sẵn trong CAU_HINH: ID của 2 Google Sheet nguồn dữ liệu
 * (PhanTichNhapTT_DRAFT, PhieuCan_DN) + ô email nhận báo cáo (để
 * trống, thiết lập sau trong menu Quản Trị). KHÔNG xóa dữ liệu nào.
 *
 * CÁCH CHẠY: Trong Apps Script Editor, chọn hàm
 * "migrateBaoCaoNgayKeo" ở dropdown trên cùng > bấm Run > chạy 1
 * lần duy nhất.
 *************************************************/
function migrateBaoCaoNgayKeo() {
  _themCauHinhNeuChua('ID_SHEET_PHANTICH_NHAP_TT', '1kjYne-hIpnHs7UfXUiokbyyEe5Y02rKaCe-kfoVwuFc',
    'ID Google Sheet chứa sheet PhanTichNhapTT_DRAFT (dùng cho Báo cáo ngày Keo)');
  _themCauHinhNeuChua('ID_SHEET_PHIEU_CAN_DN', '1vqMVxccBA7zlAMHrGsVBydGFwZJ6QuDZW10zJ74V29g',
    'ID Google Sheet chứa sheet PhieuCan_DN (dùng cho Báo cáo ngày Keo)');
  _themCauHinhNeuChua('EMAIL_BAO_CAO_NGAY', '',
    'Danh sách email nhận Báo cáo ngày Keo, cách nhau bởi dấu phẩy - thiết lập ở Quản Trị');

  SpreadsheetApp.getUi().alert(
    'Đã thiết lập xong Báo Cáo Ngày (Keo)!\n\n' +
    '- Đã lưu ID 2 Google Sheet nguồn dữ liệu (PhanTichNhapTT_DRAFT, PhieuCan_DN) vào CAU_HINH.\n' +
    '- Vào menu Báo Cáo > "Báo cáo ngày (Keo)" để xem, xuất Excel, hoặc gửi email.\n' +
    '- Vào menu Quản Trị > "Báo cáo ngày (Keo)" để cập nhật lại ID 2 sheet trên hoặc thiết lập email nhận báo cáo.\n\n' +
    'Lưu ý: Tài khoản Google đang chạy Apps Script này PHẢI có quyền xem (Viewer trở lên) trên 2 Google Sheet nguồn đó thì mới đọc được dữ liệu.'
  );
}

function _themCotConThieu(ss, sheetName, newCols) {
  const sh = ss.getSheetByName(sheetName);
  if (!sh) return;
  const lastCol = sh.getLastColumn();
  const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h).trim());
  let nextCol = lastCol;
  newCols.forEach(col => {
    if (headers.indexOf(col) === -1) {
      nextCol++;
      sh.getRange(1, nextCol).setValue(col);
    }
  });
}

/**
 * Điền giá trị mặc định cho các dòng dữ liệu cũ (trước khi có cột mới):
 * loai_quy -> Quỹ tiền mặt, tai_khoan -> 1111, ngay_hach_toan -> = ngày chứng từ.
 */
function _dienGiaTriMacDinhCotMoiSoQuy(ss, sheetName, colNgayChungTu) {
  const sh = ss.getSheetByName(sheetName);
  if (!sh) return;
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return;

  const lastCol = sh.getLastColumn();
  const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h).trim());
  const idxLoaiQuy = headers.indexOf('loai_quy');
  const idxTaiKhoan = headers.indexOf('tai_khoan');
  const idxNgayHachToan = headers.indexOf('ngay_hach_toan');
  const idxNgayCT = headers.indexOf(colNgayChungTu);

  const range = sh.getRange(2, 1, lastRow - 1, lastCol);
  const values = range.getValues();
  let changed = false;

  for (let i = 0; i < values.length; i++) {
    if (idxLoaiQuy > -1 && !values[i][idxLoaiQuy]) { values[i][idxLoaiQuy] = QUY_TIEN_MAT; changed = true; }
    if (idxTaiKhoan > -1 && !values[i][idxTaiKhoan]) { values[i][idxTaiKhoan] = '1111'; changed = true; }
    if (idxNgayHachToan > -1 && idxNgayCT > -1 && !values[i][idxNgayHachToan] && values[i][idxNgayCT]) {
      values[i][idxNgayHachToan] = values[i][idxNgayCT];
      changed = true;
    }
  }

  if (changed) range.setValues(values);
}

function _themCauHinhNeuChua(key, value, ghiChu) {
  const existing = _getCauHinh(key);
  if (existing === null || existing === undefined) {
    _sheet(SHEET_CAU_HINH).appendRow([key, value, ghiChu]);
  }
}
