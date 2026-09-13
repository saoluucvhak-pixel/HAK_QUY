/*************************************************
 * QUANTRI.GS
 * Các chức năng dành cho menu QUẢN TRỊ — TẤT CẢ đều
 * chỉ ADMIN mới được phép: quản lý người dùng, khóa sổ /
 * mở khóa, xem Nhật ký hệ thống (Audit Log).
 * currentUser = { username, full_name }.
 *************************************************/

/*************************************************
 * ===================== KHÓA SỔ =====================
 *************************************************/

/**
 * Khóa sổ 1 ngày cụ thể. Nếu ngày đó trước đây từng bị khóa rồi
 * được mở ra, khóa lại sẽ ghi đè (xóa lịch sử mở khóa cũ của
 * chính ngày đó) — Audit Log vẫn giữ đầy đủ lịch sử thao tác.
 */
function khoaSo(ngayKhoa, currentUser) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);

    if (!currentUser || !currentUser.username) throw new Error('Thiếu thông tin người dùng.');
    _yeuCauQuyen(currentUser.username, [ROLE_ADMIN]);

    if (!ngayKhoa) throw new Error('Vui lòng chọn ngày cần khóa.');
    const ngayKey = Utilities.formatDate(new Date(ngayKhoa), Session.getScriptTimeZone(), 'yyyy-MM-dd');

    const sh = _sheet(SHEET_KHOA_SO);
    const data = sh.getDataRange().getValues();
    const headers = data[0].map(h => String(h).trim());
    const idxNgay = headers.indexOf('ngay_khoa');
    const idxTrangThai = headers.indexOf('trang_thai');

    let rowIndex = -1;
    for (let i = 1; i < data.length; i++) {
      if (!data[i][idxNgay]) continue;
      const rowKey = Utilities.formatDate(new Date(data[i][idxNgay]), Session.getScriptTimeZone(), 'yyyy-MM-dd');
      if (rowKey === ngayKey) { rowIndex = i; break; }
    }

    if (rowIndex !== -1 && data[rowIndex][idxTrangThai] === KHOA_SO_DA_KHOA) {
      throw new Error('Ngày ' + ngayKey + ' đã được khóa sổ từ trước.');
    }

    const now = new Date();
    const nguoiLap = currentUser.full_name || 'N/A';
    const rowValues = [new Date(ngayKhoa), nguoiLap, now, KHOA_SO_DA_KHOA, '', '', ''];

    if (rowIndex !== -1) {
      sh.getRange(rowIndex + 1, 1, 1, rowValues.length).setValues([rowValues]);
    } else {
      sh.appendRow(rowValues);
    }

    _writeAuditLog(nguoiLap, 'Khóa Sổ', 'Thêm', ngayKey, '', 'Khóa sổ ngày ' + ngayKey);

    return _jsonOk({ ngay_khoa: ngayKey });

  } catch (err) {
    return _jsonErr(err);
  } finally {
    lock.releaseLock();
  }
}

/**
 * Mở khóa 1 ngày đã khóa — bắt buộc phải nhập lý do.
 */
function moKhoaSo(ngayKhoa, lyDoMoKhoa, currentUser) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);

    if (!currentUser || !currentUser.username) throw new Error('Thiếu thông tin người dùng.');
    _yeuCauQuyen(currentUser.username, [ROLE_ADMIN]);

    if (!ngayKhoa) throw new Error('Thiếu ngày cần mở khóa.');
    if (!lyDoMoKhoa) throw new Error('Vui lòng nhập lý do mở khóa.');

    const ngayKey = Utilities.formatDate(new Date(ngayKhoa), Session.getScriptTimeZone(), 'yyyy-MM-dd');

    const sh = _sheet(SHEET_KHOA_SO);
    const data = sh.getDataRange().getValues();
    const headers = data[0].map(h => String(h).trim());
    const idxNgay = headers.indexOf('ngay_khoa');
    const idxTrangThai = headers.indexOf('trang_thai');
    const idxNguoiMo = headers.indexOf('nguoi_mo_khoa');
    const idxThoiGianMo = headers.indexOf('thoi_gian_mo_khoa');
    const idxLyDoMo = headers.indexOf('ly_do_mo_khoa');

    let rowIndex = -1;
    for (let i = 1; i < data.length; i++) {
      if (!data[i][idxNgay]) continue;
      const rowKey = Utilities.formatDate(new Date(data[i][idxNgay]), Session.getScriptTimeZone(), 'yyyy-MM-dd');
      if (rowKey === ngayKey) { rowIndex = i; break; }
    }
    if (rowIndex === -1 || data[rowIndex][idxTrangThai] !== KHOA_SO_DA_KHOA) {
      throw new Error('Ngày ' + ngayKey + ' hiện không ở trạng thái Đã khóa.');
    }

    const now = new Date();
    const nguoiLap = currentUser.full_name || 'N/A';

    sh.getRange(rowIndex + 1, idxTrangThai + 1).setValue(KHOA_SO_DA_MO);
    sh.getRange(rowIndex + 1, idxNguoiMo + 1).setValue(nguoiLap);
    sh.getRange(rowIndex + 1, idxThoiGianMo + 1).setValue(now);
    sh.getRange(rowIndex + 1, idxLyDoMo + 1).setValue(lyDoMoKhoa);

    _writeAuditLog(nguoiLap, 'Khóa Sổ', 'Sửa', ngayKey, 'Đã khóa', 'Mở khóa - Lý do: ' + lyDoMoKhoa);

    return _jsonOk({ ngay_khoa: ngayKey });

  } catch (err) {
    return _jsonErr(err);
  } finally {
    lock.releaseLock();
  }
}

/**
 * Danh sách toàn bộ lịch sử khóa/mở khóa sổ.
 */
function getDanhSachKhoaSo(currentUser) {
  try {
    if (!currentUser || !currentUser.username) throw new Error('Thiếu thông tin người dùng.');
    _yeuCauQuyen(currentUser.username, [ROLE_ADMIN]);

    const list = _sheetToObjects(SHEET_KHOA_SO).map(k => ({
      ngay_khoa: _fmtDate(k.ngay_khoa),
      nguoi_khoa: k.nguoi_khoa,
      thoi_gian_khoa: _fmtDateTime(k.thoi_gian_khoa),
      trang_thai: k.trang_thai,
      nguoi_mo_khoa: k.nguoi_mo_khoa,
      thoi_gian_mo_khoa: k.thoi_gian_mo_khoa ? _fmtDateTime(k.thoi_gian_mo_khoa) : '',
      ly_do_mo_khoa: k.ly_do_mo_khoa,
      __row: k.__row
    }));
    list.sort((a, b) => b.__row - a.__row);
    return _jsonOk(list);
  } catch (err) {
    return _jsonErr(err);
  }
}

/*************************************************
 * ================= QUẢN LÝ NGƯỜI DÙNG =================
 *************************************************/

function getUserList(currentUser) {
  try {
    if (!currentUser || !currentUser.username) throw new Error('Thiếu thông tin người dùng.');
    _yeuCauQuyen(currentUser.username, [ROLE_ADMIN]);

    const list = _sheetToObjects(SHEET_USERS).map(u => ({
      username: u.username,
      full_name: u.full_name,
      role: u.role,
      status: u.status
    }));
    return _jsonOk(list);
  } catch (err) {
    return _jsonErr(err);
  }
}

function addUserByAdmin(payload, currentUser) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);

    if (!currentUser || !currentUser.username) throw new Error('Thiếu thông tin người dùng.');
    _yeuCauQuyen(currentUser.username, [ROLE_ADMIN]);

    if (!payload) throw new Error('Thiếu dữ liệu người dùng.');
    const username = String(payload.username || '').trim();
    if (!username) throw new Error('Vui lòng nhập Tên đăng nhập.');
    if (!payload.password || payload.password.length < 6) throw new Error('Mật khẩu tối thiểu 6 ký tự.');
    if (!payload.full_name) throw new Error('Vui lòng nhập Họ tên.');
    if ([ROLE_ADMIN, ROLE_THU_QUY, ROLE_XEM].indexOf(payload.role) === -1) throw new Error('Vai trò không hợp lệ.');

    const existed = _sheetToObjects(SHEET_USERS).find(u => String(u.username).toLowerCase() === username.toLowerCase());
    if (existed) throw new Error('Tên đăng nhập "' + username + '" đã tồn tại.');

    _sheet(SHEET_USERS).appendRow([
      'USR' + new Date().getTime(),
      username,
      _hashPassword(payload.password),
      payload.full_name,
      payload.role,
      USER_STATUS_ACTIVE,
      new Date()
    ]);

    _writeAuditLog(currentUser.full_name, 'Quản Trị', 'Thêm', username, '', 'Tạo tài khoản vai trò ' + payload.role);

    return _jsonOk({ username: username });

  } catch (err) {
    return _jsonErr(err);
  } finally {
    lock.releaseLock();
  }
}

/**
 * Khóa/Mở khóa 1 tài khoản. Không cho tự khóa chính mình để
 * tránh tình huống ADMIN duy nhất tự khóa mất quyền truy cập.
 */
function toggleUserStatus(targetUsername, currentUser) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);

    if (!currentUser || !currentUser.username) throw new Error('Thiếu thông tin người dùng.');
    _yeuCauQuyen(currentUser.username, [ROLE_ADMIN]);

    if (String(targetUsername).toLowerCase() === String(currentUser.username).toLowerCase()) {
      throw new Error('Không thể tự khóa chính tài khoản đang đăng nhập.');
    }

    const sh = _sheet(SHEET_USERS);
    const data = sh.getDataRange().getValues();
    const headers = data[0].map(h => String(h).trim());
    const idxUsername = headers.indexOf('username');
    const idxStatus = headers.indexOf('status');

    let rowIndex = -1;
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][idxUsername]).toLowerCase() === String(targetUsername).toLowerCase()) { rowIndex = i; break; }
    }
    if (rowIndex === -1) throw new Error('Không tìm thấy tài khoản: ' + targetUsername);

    const trangThaiCu = data[rowIndex][idxStatus];
    const trangThaiMoi = trangThaiCu === USER_STATUS_ACTIVE ? USER_STATUS_LOCKED : USER_STATUS_ACTIVE;
    sh.getRange(rowIndex + 1, idxStatus + 1).setValue(trangThaiMoi);

    _writeAuditLog(currentUser.full_name, 'Quản Trị', 'Sửa', targetUsername, trangThaiCu, trangThaiMoi);

    return _jsonOk({ username: targetUsername, status: trangThaiMoi });

  } catch (err) {
    return _jsonErr(err);
  } finally {
    lock.releaseLock();
  }
}

/*************************************************
 * ============ CẤU HÌNH BÁO CÁO NGÀY (KEO) + EMAIL ============
 *************************************************/

function getCauHinhBaoCaoNgay(currentUser) {
  try {
    if (!currentUser || !currentUser.username) throw new Error('Thiếu thông tin người dùng.');
    _yeuCauQuyen(currentUser.username, [ROLE_ADMIN]);

    return _jsonOk({
      id_sheet_phantich: _getCauHinh('ID_SHEET_PHANTICH_NHAP_TT') || '',
      id_sheet_phieucan: _getCauHinh('ID_SHEET_PHIEU_CAN_DN') || '',
      email_bao_cao: _getCauHinh('EMAIL_BAO_CAO_NGAY') || '',
      dong_bo_luc: _getCauHinh('KEO_DONGBO_LUC') || ''
    });
  } catch (err) {
    return _jsonErr(err);
  }
}

function luuCauHinhBaoCaoNgay(payload, currentUser) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);

    if (!currentUser || !currentUser.username) throw new Error('Thiếu thông tin người dùng.');
    _yeuCauQuyen(currentUser.username, [ROLE_ADMIN]);
    if (!payload) throw new Error('Thiếu dữ liệu cấu hình.');

    // Cho phép dán nguyên đường link Google Sheet thay vì chỉ ID —
    // _rutGonIdSheet tự tách lấy đúng phần ID để lưu, tránh lỗi
    // "Illegal spreadsheet id or key" khi mở sheet sau này.
    _setCauHinh('ID_SHEET_PHANTICH_NHAP_TT', _rutGonIdSheet(payload.id_sheet_phantich));
    _setCauHinh('ID_SHEET_PHIEU_CAN_DN', _rutGonIdSheet(payload.id_sheet_phieucan));
    _setCauHinh('EMAIL_BAO_CAO_NGAY', String(payload.email_bao_cao || '').trim());

    _writeAuditLog(currentUser.full_name, 'Quản Trị', 'Sửa', 'CAU_HINH_BAO_CAO_NGAY', '', 'Cập nhật cấu hình Báo cáo ngày (Keo) + email nhận báo cáo');

    return _jsonOk({});

  } catch (err) {
    return _jsonErr(err);
  } finally {
    lock.releaseLock();
  }
}

/*************************************************
 * ================= NHẬT KÝ HỆ THỐNG (AUDIT LOG) =================
 *************************************************/

/**
 * Xem Audit Log — giới hạn 200 dòng gần nhất khớp bộ lọc để
 * tránh tải toàn bộ log (có thể rất lớn theo thời gian) mỗi lần xem.
 */
function getAuditLogList(filters, currentUser) {
  try {
    if (!currentUser || !currentUser.username) throw new Error('Thiếu thông tin người dùng.');
    _yeuCauQuyen(currentUser.username, [ROLE_ADMIN]);

    filters = filters || {};

    let list = _sheetToObjects(SHEET_AUDIT_LOG).map(l => ({
      thoi_gian: _fmtDateTime(l.thoi_gian),
      nguoi_thao_tac: l.nguoi_thao_tac,
      chuc_nang: l.chuc_nang,
      loai_thao_tac: l.loai_thao_tac,
      ma_chung_tu: l.ma_chung_tu,
      du_lieu_truoc: l.du_lieu_truoc,
      du_lieu_sau: l.du_lieu_sau,
      __row: l.__row
    }));

    if (filters.chucNang) list = list.filter(l => String(l.chuc_nang).toLowerCase().includes(String(filters.chucNang).toLowerCase()));
    if (filters.nguoiThaoTac) list = list.filter(l => String(l.nguoi_thao_tac).toLowerCase().includes(String(filters.nguoiThaoTac).toLowerCase()));
    if (filters.maChungTu) list = list.filter(l => String(l.ma_chung_tu).toLowerCase().includes(String(filters.maChungTu).toLowerCase()));

    list.sort((a, b) => b.__row - a.__row);
    list = list.slice(0, 200);

    return _jsonOk(list);

  } catch (err) {
    return _jsonErr(err);
  }
}
