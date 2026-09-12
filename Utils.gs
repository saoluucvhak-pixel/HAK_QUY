/*************************************************
 * UTILS.GS
 * Các hàm tiện ích dùng chung cho toàn bộ hệ thống.
 *************************************************/

function _ss() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

function _sheet(name) {
  const sh = _ss().getSheetByName(name);
  if (!sh) throw new Error('Không tìm thấy sheet: ' + name);
  return sh;
}

/**
 * Loại bỏ các ký tự điều khiển/line-separator ẩn (đặc biệt U+2028, U+2029)
 * hay lọt vào dữ liệu khi copy/paste từ Excel, Word, PDF... Các ký tự này
 * không hiện ra khi nhìn ô tính, nhưng có thể khiến phản hồi của
 * google.script.run bị hỏng ngầm trong lúc truyền về trình duyệt — server
 * chạy đúng và trả dữ liệu đúng, nhưng client lại nhận về null.
 */
function _cleanCell(v) {
  if (typeof v !== 'string') return v;
  return v.replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F\u2028\u2029]/g, ' ');
}

/**
 * Đọc toàn bộ 1 sheet, trả về mảng object (key = tên cột ở dòng header).
 * __row = vị trí dòng thật trên sheet (1-based), dùng khi cần sửa lại dòng đó.
 */
function _sheetToObjects(sheetName) {
  return _sheetToObjectsFromSheetObj(_sheet(sheetName));
}

/**
 * Giống _sheetToObjects nhưng nhận thẳng 1 đối tượng Sheet thay vì tên
 * sheet trong spreadsheet đang chạy — dùng để đọc dữ liệu từ 1 Google
 * Sheet KHÁC (ví dụ PhieuCan_DN, PhanTichNhapTT_DRAFT) mở qua
 * SpreadsheetApp.openById().
 */
function _sheetToObjectsFromSheetObj(sh) {
  const values = sh.getDataRange().getValues();
  if (values.length < 2) return [];
  const headers = values[0].map(h => String(h).trim());
  const rows = [];
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    if (row.every(c => c === '' || c === null)) continue; // bỏ dòng trống
    const obj = {};
    headers.forEach((h, idx) => obj[h] = _cleanCell(row[idx]));
    obj.__row = i + 1;
    rows.push(obj);
  }
  return rows;
}

/**
 * Đặt độ rộng cột hợp lý dựa theo tên tiêu đề cột, thay cho
 * autoResizeColumns() — vốn tính độ rộng KHÔNG đáng tin cậy khi
 * dòng 1 phía trên là 1 ô đã merge() ngang qua nhiều cột (tiêu đề
 * báo cáo), dễ khiến các cột dữ liệu bên dưới bị co hẹp/lệch,
 * trông không chuẩn. Dùng cho mọi sheet Excel xuất ra.
 */
function _datDoRongCotTheoTieuDe(sh, headers, cotBatDau) {
  const batDau = cotBatDau || 1;
  headers.forEach((h, i) => {
    const label = String(h === null || h === undefined ? '' : h);
    let w = 90;
    if (label === '') w = 24;
    else if (/diễn giải|nội dung/i.test(label)) w = 260;
    else if (/người|khách hàng|đối tượng|biển số/i.test(label)) w = 170;
    else if (/ngày/i.test(label)) w = 100;
    else if (/tồn|thành tiền|giá trị|^nợ$|^có$|^thu$|^chi$/i.test(label)) w = 115;
    else if (/số phiếu/i.test(label)) w = 100;
    else if (/tổng.*\(kg\)|đơn giá/i.test(label)) w = 110;
    sh.setColumnWidth(batDau + i, w);
  });
}

/**
 * Chuyển 1 giá trị đọc từ ô Sheet (có thể là Date thật, hoặc chuỗi text
 * "yyyy-MM-dd" như trong PhanTichNhapTT_DRAFT) thành khóa ngày dạng
 * "yyyy-MM-dd" để so sánh/lọc, không phụ thuộc việc ô đó được lưu dưới
 * dạng Date hay Text.
 */
function _ngayKeyLinhHoat(v) {
  if (!v) return '';
  if (Object.prototype.toString.call(v) === '[object Date]') {
    return _isValidDate(v) ? Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd') : '';
  }
  return String(v).trim();
}

function _jsonOk(data) {
  return { success: true, data: data };
}
function _jsonErr(err) {
  return { success: false, message: (err && err.message) ? err.message : String(err) };
}

/**
 * Kiểm tra 1 giá trị có phải Date object HỢP LỆ hay không (loại trừ
 * "Invalid Date" — trường hợp Google Sheets không parse được 1 ô text
 * tưởng là ngày, ví dụ do lệch định dạng ngôn ngữ/khu vực lúc dán dữ
 * liệu). Utilities.formatDate() trên 1 Invalid Date sẽ ném lỗi, và một
 * Date không hợp lệ lọt vào phản hồi trả về trình duyệt có thể khiến
 * google.script.run âm thầm hỏng, khiến client nhận về null dù server
 * chạy thành công.
 */
function _isValidDate(d) {
  return Object.prototype.toString.call(d) === '[object Date]' && !isNaN(d.getTime());
}

/**
 * Ép 1 giá trị đọc từ Sheet thành CHUỖI an toàn để trả về client.
 * Dùng cho các cột lẽ ra là text tự do (số phiếu, mã, ghi chú...) nhưng
 * có thể lỡ bị Google Sheets tự động hiểu nhầm thành Ngày/Số khi nhập
 * liệu (ví dụ số phiếu "01/02" bị hiểu thành ngày 01/02). Nếu vô tình
 * vẫn còn 1 ô như vậy, hàm này chuyển về dạng hiển thị dd/MM/yyyy thay
 * vì để lọt 1 Date/Invalid Date thô ra ngoài (có thể làm hỏng ngầm
 * phản hồi của google.script.run).
 */
function _safeText(v) {
  if (v === null || v === undefined) return '';
  if (Object.prototype.toString.call(v) === '[object Date]') {
    return _isValidDate(v) ? Utilities.formatDate(v, Session.getScriptTimeZone(), 'dd/MM/yyyy') : '';
  }
  return String(v);
}

function _fmtDate(d) {
  if (!d) return '';
  if (_isValidDate(d)) {
    return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  if (Object.prototype.toString.call(d) === '[object Date]') return ''; // Invalid Date
  return String(d);
}

function _fmtDateTime(d) {
  if (!d) return '';
  if (_isValidDate(d)) {
    return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  }
  if (Object.prototype.toString.call(d) === '[object Date]') return ''; // Invalid Date
  return String(d);
}

/**
 * Băm mật khẩu bằng SHA-256 (KHÔNG lưu mật khẩu dạng thô trong Sheet).
 */
function _hashPassword(password) {
  const rawHash = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, password, Utilities.Charset.UTF_8);
  return rawHash.map(b => (b < 0 ? b + 256 : b).toString(16).padStart(2, '0')).join('');
}

/**
 * Lấy giá trị cấu hình từ sheet CAU_HINH theo key.
 */
function _getCauHinh(key) {
  const list = _sheetToObjects(SHEET_CAU_HINH);
  const found = list.find(c => String(c.key) === String(key));
  return found ? found.value : null;
}

/**
 * Cập nhật hoặc tạo mới 1 dòng cấu hình trong CAU_HINH.
 */
function _setCauHinh(key, value) {
  const sh = _sheet(SHEET_CAU_HINH);
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(key)) {
      sh.getRange(i + 1, 2).setValue(value);
      return;
    }
  }
  sh.appendRow([key, value, '']);
}

/**
 * Sinh số phiếu tự động dạng PT000001 / PC000001...
 */
function _generateNextNumber(configKey, prefix) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    let current = Number(_getCauHinh(configKey));
    if (!current || isNaN(current)) current = 1;
    const soPhieu = prefix + String(current).padStart(6, '0');
    _setCauHinh(configKey, current + 1);
    return soPhieu;
  } finally {
    lock.releaseLock();
  }
}

/**
 * Ghi Audit Log cho mọi thao tác quan trọng (Thêm/Sửa/Hủy).
 */
function _writeAuditLog(nguoiThaoTac, chucNang, loaiThaoTac, maChungTu, duLieuTruoc, duLieuSau) {
  try {
    const sh = _sheet(SHEET_AUDIT_LOG);
    sh.appendRow([
      'LOG' + new Date().getTime(),
      new Date(),
      nguoiThaoTac || 'N/A',
      chucNang || '',
      loaiThaoTac || '',
      maChungTu || '',
      duLieuTruoc || '',
      duLieuSau || ''
    ]);
  } catch (err) {
    Logger.log('Lỗi ghi Audit Log: ' + err.message);
  }
}

/**
 * Kiểm tra 1 ngày đã bị khóa sổ hay chưa.
 */
function _isDateLocked(dateKey) {
  const list = _sheetToObjects(SHEET_KHOA_SO);
  const found = list.find(k => _fmtDate(k.ngay_khoa) === dateKey);
  return found ? found.trang_thai === KHOA_SO_DA_KHOA : false;
}

/*************************************************
 * PHÂN QUYỀN
 * Luôn tra cứu vai trò THẬT của user từ sheet USERS
 * (không tin trực tiếp giá trị role do client gửi lên),
 * để đảm bảo phân quyền có hiệu lực thật sự chứ không
 * chỉ ẩn/hiện trên giao diện.
 *************************************************/

/**
 * @param {string} username
 * @param {string[]} allowedRoles - ví dụ [ROLE_ADMIN, ROLE_THU_QUY]
 * @returns {boolean}
 */
function _laVaiTroHopLe(username, allowedRoles) {
  if (!username) return false;
  const users = _sheetToObjects(SHEET_USERS);
  const user = users.find(u => String(u.username).toLowerCase() === String(username).toLowerCase());
  if (!user) return false;
  if (user.status !== USER_STATUS_ACTIVE) return false;
  return allowedRoles.indexOf(user.role) !== -1;
}

/**
 * Chặn thao tác nếu user không đủ quyền. Ném lỗi rõ ràng để
 * hiển thị cho người dùng biết vì sao bị từ chối.
 */
function _yeuCauQuyen(username, allowedRoles) {
  if (!_laVaiTroHopLe(username, allowedRoles)) {
    throw new Error('Bạn không có quyền thực hiện thao tác này.');
  }
}
