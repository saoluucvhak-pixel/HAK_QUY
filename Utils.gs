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
    else if (/diễn giải|nội dung/i.test(label)) w = 340;
    else if (/người|khách hàng|đối tượng|biển số/i.test(label)) w = 170;
    else if (/^nợ$|^có$/i.test(label)) w = 170;
    else if (/ngày/i.test(label)) w = 100;
    else if (/tồn|thành tiền|giá trị|^thu$|^chi$/i.test(label)) w = 115;
    else if (/số phiếu/i.test(label)) w = 100;
    else if (/tổng.*\((kg|tấn)\)|đơn giá/i.test(label)) w = 110;
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

/**
 * Tương tự _ngayKeyLinhHoat() nhưng cho cột GIỜ (vd gio_thu/gio_chi
 * lưu "HH:mm"). Google Sheets cũng tự ý hiểu chuỗi "14:35" là 1 giá
 * trị Giờ và lưu thành Date thật y hệt cách nó tự hiểu ngày tháng —
 * nếu không ép về lại chuỗi "HH:mm" trước khi trả cho client, 1 Date
 * lọt thẳng vào JSON trả về của google.script.run có thể khiến toàn
 * bộ phản hồi bị hỏng, client nhận về null dù server chạy đúng.
 */
function _gioLinhHoat(v) {
  if (!v) return '';
  if (Object.prototype.toString.call(v) === '[object Date]') {
    return _isValidDate(v) ? Utilities.formatDate(v, Session.getScriptTimeZone(), 'HH:mm') : '';
  }
  return String(v).trim();
}

function _jsonOk(data) {
  return { success: true, data: data };
}
function _jsonErr(err) {
  return { success: false, message: _ngayVN((err && err.message) ? err.message : String(err)) };
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

/**
 * Chuẩn hoá tên người dùng để ghép vào tên file Excel xuất ra — thay
 * khoảng trắng bằng "_" và bỏ các ký tự không hợp lệ trong tên file
 * (/ \ : * ? " < > |). Thiếu currentUser/full_name thì trả về
 * 'KhongRo' thay vì để tên file thiếu 1 đoạn hoặc ném lỗi.
 */
function _tenNguoiDungChoFile(currentUser) {
  const ten = (currentUser && (currentUser.full_name || currentUser.username)) || 'KhongRo';
  return String(ten).trim().replace(/[\\/:*?"<>|]/g, '').replace(/\s+/g, '_');
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
 *
 * Đọc từ CUỐI sheet lên (không phải từ đầu xuống) — nếu vì lý do nào
 * đó (sửa tay, lỗi cũ...) có nhiều dòng cùng trùng 1 key, dòng nằm ở
 * dưới luôn là dòng được ghi gần đây nhất nên đáng tin hơn dòng phía
 * trên. _setCauHinh() bên dưới cũng tự dọn dòng trùng mỗi khi ghi, nên
 * theo thời gian sheet sẽ tự hết trùng, nhưng hàm đọc vẫn cần an toàn
 * cho cả những dòng trùng cũ chưa kịp dọn.
 */
function _getCauHinh(key) {
  const list = _sheetToObjects(SHEET_CAU_HINH);
  for (let i = list.length - 1; i >= 0; i--) {
    if (String(list[i].key) === String(key)) return _chuanHoaGiaTriCauHinh(list[i].value);
  }
  return null;
}

/**
 * Google Sheets có thể tự động diễn giải 1 chuỗi trông giống ngày giờ
 * (vd giá trị KEO_DONGBO_LUC dạng "yyyy-MM-dd HH:mm:ss" do
 * _setCauHinh() ghi bằng setValue()) thành 1 ô kiểu Date thật sự —
 * dù cột value trong CAU_HINH vốn chỉ để chứa chuỗi. Khi đọc lại,
 * _sheetToObjects() trả về nguyên object Date đó thay vì chuỗi. Nếu
 * để lọt 1 Date (nhất là Date không hợp lệ) vào JSON trả về trình
 * duyệt qua google.script.run, request có thể âm thầm hỏng và client
 * nhận về null dù server chạy thành công (xem thêm _isValidDate()) —
 * nên luôn chuẩn hoá về chuỗi ngay tại nguồn đọc, cho mọi nơi gọi
 * _getCauHinh() dùng chung, thay vì phải tự phòng thủ riêng lẻ.
 */
function _chuanHoaGiaTriCauHinh(v) {
  if (Object.prototype.toString.call(v) === '[object Date]') {
    return _isValidDate(v) ? Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss') : '';
  }
  return v;
}

/**
 * Cập nhật hoặc tạo mới 1 dòng cấu hình trong CAU_HINH.
 *
 * Nếu phát hiện NHIỀU dòng cùng key (dữ liệu trùng lặp tồn đọng), chỉ
 * ghi giá trị mới vào dòng CUỐI CÙNG rồi xoá các dòng trùng phía trên —
 * tự dọn dẹp, không để tích tụ thêm, khớp với thứ tự đọc của
 * _getCauHinh() ở trên (luôn ưu tiên dòng dưới cùng).
 */
function _setCauHinh(key, value) {
  const sh = _sheet(SHEET_CAU_HINH);
  const data = sh.getDataRange().getValues();
  const dongTrung = [];
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(key)) dongTrung.push(i + 1); // số dòng thật (1-based)
  }
  if (dongTrung.length > 0) {
    const dongGhi = dongTrung[dongTrung.length - 1];
    // Đặt định dạng ô về "Văn bản thuần" (@) TRƯỚC khi ghi giá trị —
    // nếu ghi trước rồi mới đổi định dạng, Sheets đã kịp tự diễn giải
    // chuỗi thành Date/Number mất rồi, đổi định dạng sau không cứu lại
    // được chuỗi gốc. Xem thêm chú thích ở _chuanHoaGiaTriCauHinh().
    sh.getRange(dongGhi, 2).setNumberFormat('@').setValue(value);
    dongTrung.slice(0, -1).sort((a, b) => b - a).forEach(dong => sh.deleteRow(dong));
    return;
  }
  const dongMoi = sh.getLastRow() + 1;
  sh.getRange(dongMoi, 2).setNumberFormat('@');
  sh.getRange(dongMoi, 1, 1, 3).setValues([[key, value, '']]);
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
    _appendRowVN(sh, [
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

/*************************************************
 * ĐỊNH DẠNG CHUẨN VIỆT NAM — dùng chung cho ghi Sheet dữ liệu và
 * kết xuất báo cáo Excel:
 *   1. Số: canh PHẢI, có phân cách hàng nghìn (#,##0)
 *   2. Chuỗi: canh TRÁI
 *   3. Ngày: canh GIỮA
 *   4. Ngày: dd/MM/yyyy (ngày giờ: dd/MM/yyyy HH:mm:ss)
 * Nội bộ server/client vẫn trao đổi khóa ngày 'yyyy-MM-dd' để lọc/so
 * sánh (sắp xếp chuỗi đúng thứ tự thời gian); chỉ tầng hiển thị/ghi
 * ô mới đổi sang dd/MM/yyyy.
 *************************************************/
const DINH_DANG_NGAY_VN = 'dd/MM/yyyy';
const DINH_DANG_NGAY_GIO_VN = 'dd/MM/yyyy HH:mm:ss';
const DINH_DANG_GIO_VN = 'HH:mm';
const DINH_DANG_SO_VN = '#,##0';
const DINH_DANG_SO_LE_VN = '#,##0.00';

/** Đổi mọi khóa ngày 'yyyy-MM-dd' nằm trong 1 chuỗi (hoặc 1 Date) sang 'dd/MM/yyyy'. */
function _ngayVN(v) {
  if (v === null || v === undefined || v === '') return '';
  if (_isValidDate(v)) return Utilities.formatDate(v, Session.getScriptTimeZone(), DINH_DANG_NGAY_VN);
  return String(v).replace(/\b(\d{4})-(\d{2})-(\d{2})\b/g, '$3/$2/$1');
}

/**
 * Chuẩn hóa 1 giá trị trước khi ghi vào ô: chuỗi 'yyyy-MM-dd' /
 * 'yyyy-MM-dd HH:mm(:ss)' được đổi thành Date thật để định dạng
 * dd/MM/yyyy và sắp xếp được trong Excel.
 */
function _giaTriOVN(v) {
  if (typeof v !== 'string') return v;
  const t = v.trim();
  const tz = Session.getScriptTimeZone();
  if (/^\d{4}-\d{2}-\d{2}$/.test(t)) return Utilities.parseDate(t, tz, 'yyyy-MM-dd');
  if (/^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}$/.test(t)) return Utilities.parseDate(t, tz, 'yyyy-MM-dd HH:mm:ss');
  if (/^\d{4}-\d{2}-\d{2} \d{2}:\d{2}$/.test(t)) return Utilities.parseDate(t, tz, 'yyyy-MM-dd HH:mm');
  return v;
}

function _coGioTrongNgay(d) {
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'HH:mm:ss') !== '00:00:00';
}

/**
 * Ghi 1 bảng dữ liệu vào range (kích thước = rows) theo chuẩn VN:
 * tự nhận diện kiểu từng ô (Date / số / chuỗi) để đặt định dạng số và
 * canh lề. Định dạng được đặt TRƯỚC khi ghi giá trị để chuỗi dạng mã
 * (vd tài khoản "1111", số phiếu "000123") giữ nguyên là chữ, không
 * bị Sheets tự đổi thành số.
 * @param {Range} range
 * @param {Array<Array>} rows
 * @param {{cotTien?:number[], cotThapPhan?:number[]}} [opts] chỉ số cột 0-based
 */
function _ghiBangVN(range, rows, opts) {
  opts = opts || {};
  const cotTien = {}, cotLe = {};
  (opts.cotTien || []).forEach(function (c) { cotTien[c] = true; });
  (opts.cotThapPhan || []).forEach(function (c) { cotLe[c] = true; });

  const values = [], formats = [], aligns = [];
  rows.forEach(function (r) {
    const v = [], f = [], a = [];
    r.forEach(function (cell, c) {
      const x = _giaTriOVN(cell);
      if (_isValidDate(x)) {
        v.push(x); f.push(_coGioTrongNgay(x) ? DINH_DANG_NGAY_GIO_VN : DINH_DANG_NGAY_VN); a.push('center');
      } else if (typeof x === 'number' && isFinite(x)) {
        v.push(x);
        f.push(cotLe[c] ? DINH_DANG_SO_LE_VN : (Math.round(x) === x || cotTien[c] ? DINH_DANG_SO_VN : DINH_DANG_SO_LE_VN));
        a.push('right');
      } else if (x === '' || x === null || x === undefined) {
        v.push('');
        f.push(cotLe[c] ? DINH_DANG_SO_LE_VN : cotTien[c] ? DINH_DANG_SO_VN : '@');
        a.push(cotLe[c] || cotTien[c] ? 'right' : 'left');
      } else {
        v.push(typeof x === 'boolean' ? (x ? 'Có' : 'Không') : String(x));
        f.push('@'); a.push('left');
      }
    });
    values.push(v); formats.push(f); aligns.push(a);
  });
  if (!values.length) return range;
  range.setNumberFormats(formats);
  range.setValues(values);
  range.setHorizontalAlignments(aligns);
  return range;
}

/*************************************************
 * ĐỊNH DẠNG CỘT CÁC SHEET DỮ LIỆU (PHIEU_THU, PHIEU_CHI, CONG_NO...)
 * Kiểu cột suy ra từ TÊN CỘT ở dòng tiêu đề.
 *************************************************/
const _COT_SO_SHEET = {
  so_tien: 1, so_tien_phat_sinh: 1, da_thanh_toan: 1, con_lai: 1, so_tien_thanh_toan: 1,
  ton_he_thong: 1, tien_thuc_te: 1, chenh_lech: 1, buoi_trua: 1, buoi_toi: 1, tong_suat: 1,
  don_gia: 1, thanh_tien: 1, so_tien_tam_ung: 1
};

function _kieuCotSheet(tenCot) {
  const h = String(tenCot || '').trim();
  if (/^(ngay|han_)/.test(h)) return 'ngay';
  if (/^thoi_gian|^created_at$|^updated_at$/.test(h)) return 'ngaygio';
  if (/^gio_/.test(h)) return 'gio';
  if (_COT_SO_SHEET[h]) return 'so';
  return 'chu';
}

function _dinhDangTheoKieu(kieu) {
  if (kieu === 'ngay') return { f: DINH_DANG_NGAY_VN, a: 'center' };
  if (kieu === 'ngaygio') return { f: DINH_DANG_NGAY_GIO_VN, a: 'center' };
  if (kieu === 'gio') return { f: DINH_DANG_GIO_VN, a: 'center' };
  if (kieu === 'so') return { f: DINH_DANG_SO_VN, a: 'right' };
  return { f: null, a: 'left' };
}

/** Định dạng toàn bộ cột dữ liệu (từ dòng 2 tới hết lưới) của 1 sheet. */
function _dinhDangSheetVN(sh) {
  if (!sh || sh.getName() === SHEET_CAU_HINH) return;
  const soCot = sh.getLastColumn();
  const soDong = sh.getMaxRows() - 1;
  if (soCot < 1 || soDong < 1) return;
  const headers = sh.getRange(1, 1, 1, soCot).getValues()[0];
  headers.forEach(function (h, i) {
    const dd = _dinhDangTheoKieu(_kieuCotSheet(h));
    const rg = sh.getRange(2, i + 1, soDong, 1);
    if (dd.f) rg.setNumberFormat(dd.f);
    rg.setHorizontalAlignment(dd.a);
  });
  sh.getRange(1, 1, 1, soCot).setHorizontalAlignment('center');
}

/** Định dạng 1 dòng dữ liệu vừa ghi (dùng sau appendRow / ghi đè dòng). */
function _dinhDangDongVN(sh, dong) {
  if (!sh || sh.getName() === SHEET_CAU_HINH || dong < 2) return;
  const soCot = sh.getLastColumn();
  const headers = sh.getRange(1, 1, 1, soCot).getValues()[0];
  const rg = sh.getRange(dong, 1, 1, soCot);
  const fmtCu = rg.getNumberFormats()[0];
  const f = [], a = [];
  headers.forEach(function (h, i) {
    const dd = _dinhDangTheoKieu(_kieuCotSheet(h));
    f.push(dd.f || fmtCu[i]);
    a.push(dd.a);
  });
  rg.setNumberFormats([f]);
  rg.setHorizontalAlignments([a]);
}

/**
 * Thêm 1 dòng theo TÊN CỘT (khớp dòng tiêu đề) thay vì theo vị trí cột.
 * Ghi theo vị trí dễ lệch cột khi sheet thật có thứ tự/số cột khác mẫu
 * (vd cột doi_soat được thêm sau bằng migrate, có cột chèn thêm) — khi
 * đó giá trị rơi vào sai cột, đọc lại (theo tên cột) thành rỗng. Cột
 * nào chưa có trên sheet sẽ được tự thêm vào cuối dòng tiêu đề.
 */
function _appendRowTheoTenCot(sh, duLieu) {
  const lastCol = Math.max(sh.getLastColumn(), 1);
  const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h).trim());
  Object.keys(duLieu).forEach(function (k) {
    if (headers.indexOf(k) === -1) {
      headers.push(k);
      sh.getRange(1, headers.length).setValue(k);
    }
  });
  const row = headers.map(function (h) {
    return Object.prototype.hasOwnProperty.call(duLieu, h) ? duLieu[h] : '';
  });
  _appendRowVN(sh, row);
}

/** appendRow + định dạng chuẩn VN cho dòng vừa thêm. */
function _appendRowVN(sh, rowValues) {
  sh.appendRow(rowValues);
  try {
    _dinhDangDongVN(sh, sh.getLastRow());
  } catch (err) {
    Logger.log('Lỗi định dạng dòng mới (' + sh.getName() + '): ' + err.message);
  }
}

/**
 * CHẠY 1 LẦN (Apps Script Editor > chọn hàm > Run, hoặc menu "HAK_QUY"
 * trên Google Sheet): định dạng lại toàn bộ dữ liệu đã có theo chuẩn
 * VN — ngày dd/MM/yyyy canh giữa, số #,##0 canh phải, chữ canh trái.
 * Không thay đổi bất kỳ giá trị nào, chỉ đổi cách hiển thị.
 */
function dinhDangLaiDuLieuVN() {
  const ss = _ss();
  [SHEET_USERS, SHEET_DOITUONG, SHEET_LOAI_THU, SHEET_LOAI_CHI, SHEET_PHIEU_THU, SHEET_PHIEU_CHI,
    SHEET_CONG_NO, SHEET_THANH_TOAN_CONG_NO, SHEET_KIEM_KE_QUY, SHEET_KHOA_SO, SHEET_AUDIT_LOG, SHEET_SO_COM
  ].forEach(function (ten) {
    const sh = ss.getSheetByName(ten);
    if (sh) _dinhDangSheetVN(sh);
  });
  ss.setSpreadsheetLocale('vi_VN');
  try {
    SpreadsheetApp.getUi().alert('Đã định dạng lại dữ liệu: ngày dd/mm/yyyy (canh giữa), số có phân cách (canh phải), chữ canh trái.');
  } catch (e) { /* chạy từ trigger/không có UI */ }
}

/** Thêm menu "HAK_QUY" trên Google Sheet dữ liệu. */
function onOpen() {
  try {
    SpreadsheetApp.getUi().createMenu('HAK_QUY')
      .addItem('Định dạng lại dữ liệu (dd/mm/yyyy, số có phân cách)', 'dinhDangLaiDuLieuVN')
      .addToUi();
  } catch (e) { /* không có UI */ }
}
