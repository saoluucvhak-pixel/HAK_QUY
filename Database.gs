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

  const defaultSheet = ss.getSheetByName('Sheet1');
  if (defaultSheet && ss.getSheets().length > 1) {
    ss.deleteSheet(defaultSheet);
  }

  SpreadsheetApp.getUi().alert(
    'Đã tạo xong Database (12 sheet)!\n\n' +
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

/* ---------- 5. PHIEU_THU ---------- */
function _setupPhieuThuSheet(ss) {
  const sh = _getOrCreateSheet(ss, SHEET_PHIEU_THU);
  const headers = [
    'id', 'so_phieu_thu', 'ngay_thu', 'gio_thu', 'loai_giao_dich',
    'nguoi_nop_tien', 'ma_doi_tuong', 'noi_dung_thu', 'ma_loai_thu', 'so_tien',
    'chung_tu_lien_quan', 'ghi_chu', 'nguoi_lap', 'thoi_gian_lap',
    'nguoi_sua_cuoi', 'thoi_gian_sua_cuoi', 'trang_thai', 'ly_do_huy'
  ];
  sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  sh.getRange(2, 3, sh.getMaxRows() - 1, 1).setNumberFormat('yyyy-mm-dd');
  sh.getRange(2, 10, sh.getMaxRows() - 1, 1).setNumberFormat('#,##0');
  _styleHeader(sh, headers.length);
}

/* ---------- 6. PHIEU_CHI ---------- */
function _setupPhieuChiSheet(ss) {
  const sh = _getOrCreateSheet(ss, SHEET_PHIEU_CHI);
  const headers = [
    'id', 'so_phieu_chi', 'ngay_chi', 'gio_chi', 'loai_giao_dich',
    'nguoi_nhan_tien', 'ma_doi_tuong', 'noi_dung_chi', 'ma_loai_chi', 'so_tien',
    'chung_tu_lien_quan', 'ghi_chu', 'nguoi_lap', 'thoi_gian_lap',
    'nguoi_sua_cuoi', 'thoi_gian_sua_cuoi', 'trang_thai', 'ly_do_huy'
  ];
  sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  sh.getRange(2, 3, sh.getMaxRows() - 1, 1).setNumberFormat('yyyy-mm-dd');
  sh.getRange(2, 10, sh.getMaxRows() - 1, 1).setNumberFormat('#,##0');
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
    ['SO_DU_QUY_KHOI_TAO', 0, 'Số dư quỹ tiền mặt tại thời điểm bắt đầu dùng hệ thống']
  ];
  sh.getRange(2, 1, sample.length, headers.length).setValues(sample);

  _styleHeader(sh, headers.length);
}
