/*************************************************
 * THUCHI.GS
 * Xử lý nghiệp vụ Phiếu Thu, Phiếu Chi, Sổ Quỹ.
 *
 * TỪ GIAI ĐOẠN 5: addPhieuThu/addPhieuChi nhận
 * currentUser dạng OBJECT { username, full_name }
 * thay vì chuỗi tên như trước — username dùng để
 * kiểm tra quyền THẬT từ sheet USERS, full_name dùng
 * để ghi vào cột nguoi_lap cho báo cáo dễ đọc.
 * Quyền lập Phiếu Thu/Chi: ADMIN hoặc THU_QUY.
 *************************************************/

/*************************************************
 * API: LẤY DANH MỤC CHO FORM (Đối tượng, Loại thu, Loại chi)
 *************************************************/
function getFormOptions() {
  try {
    const doiTuong = _sheetToObjects(SHEET_DOITUONG)
      .filter(d => d.trang_thai === 'Hoạt động')
      .map(d => ({ ma: d.ma_doi_tuong, ten: d.ten_doi_tuong }));

    const loaiThu = _sheetToObjects(SHEET_LOAI_THU)
      .filter(l => l.trang_thai === 'Hoạt động')
      .map(l => ({ ma: l.ma_loai_thu, ten: l.ten_loai_thu }));

    const loaiChi = _sheetToObjects(SHEET_LOAI_CHI)
      .filter(l => l.trang_thai === 'Hoạt động')
      .map(l => ({ ma: l.ma_loai_chi, ten: l.ten_loai_chi }));

    return _jsonOk({ doiTuong: doiTuong, loaiThu: loaiThu, loaiChi: loaiChi });
  } catch (err) {
    return _jsonErr(err);
  }
}

/**
 * Chuẩn hóa loai_quy: dữ liệu cũ (trước khi có cột này) coi như
 * thuộc Quỹ tiền mặt để đảm bảo tương thích ngược.
 */
function _chuanHoaLoaiQuy(loaiQuy) {
  return loaiQuy === QUY_CONG_DOAN ? QUY_CONG_DOAN : QUY_TIEN_MAT;
}

function _cauHinhSoDuKhoiTaoTheoQuy(loaiQuy) {
  return loaiQuy === QUY_CONG_DOAN ? 'SO_DU_QUY_CONG_DOAN_KHOI_TAO' : 'SO_DU_QUY_KHOI_TAO';
}

/**
 * Duyệt TOÀN BỘ lịch sử giao dịch hợp lệ CỦA CÙNG 1 QUỸ (kể cả giao dịch
 * giả định mới, nếu truyền vào), tính tồn quỹ lũy kế theo đúng thứ tự
 * thời gian, trả về giá trị tồn THẤP NHẤT từng xuất hiện trong toàn bộ
 * dòng thời gian. Mỗi quỹ (Quỹ tiền mặt / Quỹ Công đoàn) có tồn riêng,
 * không ảnh hưởng lẫn nhau.
 */
function _timTonThapNhat(giaoDichGiaDinh, loaiQuy) {
  loaiQuy = _chuanHoaLoaiQuy(loaiQuy);
  const soDuKhoiTao = Number(_getCauHinh(_cauHinhSoDuKhoiTaoTheoQuy(loaiQuy))) || 0;
  const list = [];

  _sheetToObjects(SHEET_PHIEU_THU).forEach(p => {
    if (p.trang_thai === TRANG_THAI_HOP_LE && _chuanHoaLoaiQuy(p.loai_quy) === loaiQuy) {
      list.push({
        ngay: _fmtDate(p.ngay_thu), gio: p.gio_thu || '00:00',
        thoiGianLap: p.thoi_gian_lap, soTien: Number(p.so_tien) || 0, loai: 'Thu'
      });
    }
  });
  _sheetToObjects(SHEET_PHIEU_CHI).forEach(p => {
    if (p.trang_thai === TRANG_THAI_HOP_LE && _chuanHoaLoaiQuy(p.loai_quy) === loaiQuy) {
      list.push({
        ngay: _fmtDate(p.ngay_chi), gio: p.gio_chi || '00:00',
        thoiGianLap: p.thoi_gian_lap, soTien: Number(p.so_tien) || 0, loai: 'Chi'
      });
    }
  });

  if (giaoDichGiaDinh) {
    list.push({
      ngay: giaoDichGiaDinh.ngay, gio: '23:59',
      thoiGianLap: new Date(8640000000000000),
      soTien: giaoDichGiaDinh.soTien, loai: giaoDichGiaDinh.loai
    });
  }

  list.sort((a, b) => {
    const keyA = a.ngay + ' ' + a.gio;
    const keyB = b.ngay + ' ' + b.gio;
    if (keyA !== keyB) return keyA < keyB ? -1 : 1;
    const tA = a.thoiGianLap ? new Date(a.thoiGianLap).getTime() : 0;
    const tB = b.thoiGianLap ? new Date(b.thoiGianLap).getTime() : 0;
    return tA - tB;
  });

  let running = soDuKhoiTao;
  let min = running;
  list.forEach(t => {
    running += (t.loai === 'Thu' ? t.soTien : -t.soTien);
    if (running < min) min = running;
  });

  return min;
}

/**
 * Tồn quỹ HIỆN TẠI (không phải mức thấp nhất lịch sử) của 1 quỹ,
 * dùng cho Dashboard / các chỗ chỉ cần biết số dư mới nhất.
 */
function _tonQuyHienTai(loaiQuy) {
  loaiQuy = _chuanHoaLoaiQuy(loaiQuy);
  let total = Number(_getCauHinh(_cauHinhSoDuKhoiTaoTheoQuy(loaiQuy))) || 0;

  _sheetToObjects(SHEET_PHIEU_THU).forEach(p => {
    if (p.trang_thai === TRANG_THAI_HOP_LE && _chuanHoaLoaiQuy(p.loai_quy) === loaiQuy) total += Number(p.so_tien) || 0;
  });
  _sheetToObjects(SHEET_PHIEU_CHI).forEach(p => {
    if (p.trang_thai === TRANG_THAI_HOP_LE && _chuanHoaLoaiQuy(p.loai_quy) === loaiQuy) total -= Number(p.so_tien) || 0;
  });

  return total;
}

/*************************************************
 * API: THÊM PHIẾU THU
 * currentUser = { username, full_name }
 *************************************************/
function addPhieuThu(payload, currentUser) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);

    if (!currentUser || !currentUser.username) throw new Error('Thiếu thông tin người dùng.');
    _yeuCauQuyen(currentUser.username, [ROLE_ADMIN, ROLE_THU_QUY]);

    if (!payload) throw new Error('Thiếu dữ liệu phiếu thu.');
    if (!payload.ngay_thu) throw new Error('Vui lòng chọn Ngày thu.');

    const ngayKey = Utilities.formatDate(new Date(payload.ngay_thu), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    if (_isDateLocked(ngayKey)) throw new Error('Ngày ' + ngayKey + ' đã bị khóa sổ, không thể thêm phiếu.');

    const loaiQuy = _chuanHoaLoaiQuy(payload.loai_quy);
    const taiKhoan = String(payload.tai_khoan || '1111').trim() || '1111';
    const ngayHachToan = payload.ngay_hach_toan ? new Date(payload.ngay_hach_toan) : new Date(payload.ngay_thu);

    const loaiGiaoDich = payload.loai_giao_dich || 'Thu thường';
    if (loaiGiaoDich !== 'Thu thường' && loaiGiaoDich !== 'Thu công nợ') {
      throw new Error('Loại giao dịch không hợp lệ.');
    }

    const soTien = Number(payload.so_tien);
    if (!soTien || soTien <= 0) throw new Error('Số tiền không hợp lệ.');
    if (!payload.ma_loai_thu) throw new Error('Vui lòng chọn Loại thu.');
    if (!payload.nguoi_nop_tien) throw new Error('Vui lòng nhập Người nộp tiền.');

    if (loaiGiaoDich === 'Thu công nợ') {
      if (!payload.ma_doi_tuong) throw new Error('Vui lòng chọn Đối tượng khi Thu công nợ.');
      if (!payload.ma_cong_no) throw new Error('Vui lòng chọn Khoản công nợ cần thu.');

      const congNoRow = _sheetToObjects(SHEET_CONG_NO).find(c => c.ma_cong_no === payload.ma_cong_no);
      if (!congNoRow) throw new Error('Không tìm thấy khoản công nợ: ' + payload.ma_cong_no);
      if (congNoRow.loai_cong_no !== 'Phải thu') throw new Error('Khoản công nợ này không phải loại "Phải thu".');
      if (congNoRow.ma_doi_tuong !== payload.ma_doi_tuong) throw new Error('Khoản công nợ không thuộc đối tượng đã chọn.');
      if (congNoRow.trang_thai !== 'Còn nợ') throw new Error('Khoản công nợ này đã được tất toán.');

      const conLaiHienTai = Number(congNoRow.con_lai) || 0;
      if (soTien > conLaiHienTai) {
        throw new Error('Số tiền thu (' + soTien.toLocaleString('vi-VN') + ' đ) vượt quá số còn nợ (' +
          conLaiHienTai.toLocaleString('vi-VN') + ' đ) của khoản ' + payload.ma_cong_no + '.');
      }
    }

    const soPhieuThu = _generateNextNumber('SO_PHIEU_THU_TIEP_THEO', 'PT');
    const now = new Date();
    const nguoiLap = currentUser.full_name || 'N/A';

    _sheet(SHEET_PHIEU_THU).appendRow([
      'PT_' + now.getTime(), soPhieuThu, new Date(payload.ngay_thu),
      payload.gio_thu || Utilities.formatDate(now, Session.getScriptTimeZone(), 'HH:mm'),
      loaiGiaoDich, payload.nguoi_nop_tien, payload.ma_doi_tuong || '',
      payload.noi_dung_thu || '', payload.ma_loai_thu, soTien,
      payload.chung_tu_lien_quan || '', payload.ghi_chu || '',
      nguoiLap, now, '', '', TRANG_THAI_HOP_LE, '',
      loaiQuy, ngayHachToan, taiKhoan, payload.tk_doi_ung || '',
      payload.ma_nhan_vien || '', payload.chi_nhanh || ''
    ]);

    _writeAuditLog(nguoiLap, 'Phiếu Thu', 'Thêm', soPhieuThu, '',
      'Số tiền: ' + soTien.toLocaleString('vi-VN') + ' đ | Nội dung: ' + (payload.noi_dung_thu || ''));

    if (loaiGiaoDich === 'Thu công nợ') {
      _capNhatCongNoSauThanhToan(payload.ma_cong_no, soTien, 'Thu nợ', soPhieuThu, nguoiLap);
      _writeAuditLog(nguoiLap, 'Công Nợ', 'Sửa', payload.ma_cong_no, '',
        'Thu nợ ' + soTien.toLocaleString('vi-VN') + ' đ qua phiếu ' + soPhieuThu);
    }

    return _jsonOk({ so_phieu_thu: soPhieuThu });

  } catch (err) {
    return _jsonErr(err);
  } finally {
    lock.releaseLock();
  }
}

/*************************************************
 * API: THÊM PHIẾU CHI
 * currentUser = { username, full_name }
 *************************************************/
function addPhieuChi(payload, currentUser) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);

    if (!currentUser || !currentUser.username) throw new Error('Thiếu thông tin người dùng.');
    _yeuCauQuyen(currentUser.username, [ROLE_ADMIN, ROLE_THU_QUY]);

    if (!payload) throw new Error('Thiếu dữ liệu phiếu chi.');
    if (!payload.ngay_chi) throw new Error('Vui lòng chọn Ngày chi.');

    const ngayKey = Utilities.formatDate(new Date(payload.ngay_chi), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    if (_isDateLocked(ngayKey)) throw new Error('Ngày ' + ngayKey + ' đã bị khóa sổ, không thể thêm phiếu.');

    const loaiQuy = _chuanHoaLoaiQuy(payload.loai_quy);
    const taiKhoan = String(payload.tai_khoan || '1111').trim() || '1111';
    const ngayHachToan = payload.ngay_hach_toan ? new Date(payload.ngay_hach_toan) : new Date(payload.ngay_chi);

    const loaiGiaoDich = payload.loai_giao_dich || 'Chi thường';
    if (loaiGiaoDich !== 'Chi thường' && loaiGiaoDich !== 'Trả công nợ') {
      throw new Error('Loại giao dịch không hợp lệ.');
    }

    const soTien = Number(payload.so_tien);
    if (!soTien || soTien <= 0) throw new Error('Số tiền không hợp lệ.');
    if (!payload.ma_loai_chi) throw new Error('Vui lòng chọn Loại chi.');
    if (!payload.nguoi_nhan_tien) throw new Error('Vui lòng nhập Người nhận tiền.');

    if (loaiGiaoDich === 'Trả công nợ') {
      if (!payload.ma_doi_tuong) throw new Error('Vui lòng chọn Đối tượng khi Trả công nợ.');
      if (!payload.ma_cong_no) throw new Error('Vui lòng chọn Khoản công nợ cần trả.');

      const congNoRow = _sheetToObjects(SHEET_CONG_NO).find(c => c.ma_cong_no === payload.ma_cong_no);
      if (!congNoRow) throw new Error('Không tìm thấy khoản công nợ: ' + payload.ma_cong_no);
      if (congNoRow.loai_cong_no !== 'Phải trả') throw new Error('Khoản công nợ này không phải loại "Phải trả".');
      if (congNoRow.ma_doi_tuong !== payload.ma_doi_tuong) throw new Error('Khoản công nợ không thuộc đối tượng đã chọn.');
      if (congNoRow.trang_thai !== 'Còn nợ') throw new Error('Khoản công nợ này đã được tất toán.');

      const conLaiHienTai = Number(congNoRow.con_lai) || 0;
      if (soTien > conLaiHienTai) {
        throw new Error('Số tiền trả (' + soTien.toLocaleString('vi-VN') + ' đ) vượt quá số còn nợ (' +
          conLaiHienTai.toLocaleString('vi-VN') + ' đ) của khoản ' + payload.ma_cong_no + '.');
      }
    }

    const tonThapNhatNeuChi = _timTonThapNhat({ ngay: ngayKey, soTien: soTien, loai: 'Chi' }, loaiQuy);
    if (tonThapNhatNeuChi < 0) {
      throw new Error(
        'Không thể lưu: giao dịch này sẽ khiến tồn quỹ bị ÂM tại một thời điểm nào đó ' +
        '(mức thấp nhất: ' + tonThapNhatNeuChi.toLocaleString('vi-VN') + ' đ). ' +
        'Vui lòng kiểm tra lại ngày chi hoặc số tiền.'
      );
    }

    const soPhieuChi = _generateNextNumber('SO_PHIEU_CHI_TIEP_THEO', 'PC');
    const now = new Date();
    const nguoiLap = currentUser.full_name || 'N/A';

    _sheet(SHEET_PHIEU_CHI).appendRow([
      'PC_' + now.getTime(), soPhieuChi, new Date(payload.ngay_chi),
      payload.gio_chi || Utilities.formatDate(now, Session.getScriptTimeZone(), 'HH:mm'),
      loaiGiaoDich, payload.nguoi_nhan_tien, payload.ma_doi_tuong || '',
      payload.noi_dung_chi || '', payload.ma_loai_chi, soTien,
      payload.chung_tu_lien_quan || '', payload.ghi_chu || '',
      nguoiLap, now, '', '', TRANG_THAI_HOP_LE, '',
      loaiQuy, ngayHachToan, taiKhoan, payload.tk_doi_ung || '',
      payload.ma_nhan_vien || '', payload.chi_nhanh || ''
    ]);

    _writeAuditLog(nguoiLap, 'Phiếu Chi', 'Thêm', soPhieuChi, '',
      'Số tiền: ' + soTien.toLocaleString('vi-VN') + ' đ | Nội dung: ' + (payload.noi_dung_chi || ''));

    if (loaiGiaoDich === 'Trả công nợ') {
      _capNhatCongNoSauThanhToan(payload.ma_cong_no, soTien, 'Trả nợ', soPhieuChi, nguoiLap);
      _writeAuditLog(nguoiLap, 'Công Nợ', 'Sửa', payload.ma_cong_no, '',
        'Trả nợ ' + soTien.toLocaleString('vi-VN') + ' đ qua phiếu ' + soPhieuChi);
    }

    return _jsonOk({ so_phieu_chi: soPhieuChi });

  } catch (err) {
    return _jsonErr(err);
  } finally {
    lock.releaseLock();
  }
}

/*************************************************
 * API: HỦY PHIẾU THU / PHIẾU CHI
 * Theo đúng nguyên tắc kế toán: KHÔNG sửa/xóa phiếu đã lập — chỉ
 * đổi trạng thái sang "Đã hủy" kèm lý do, giữ nguyên lịch sử để
 * truy vết (Audit Log + số phiếu cũ vẫn còn trên Sổ Quỹ, chỉ không
 * còn được tính vào tồn quỹ). Muốn ghi nhận đúng thì lập phiếu mới.
 * Quyền: ADMIN hoặc THU_QUY (giống quyền lập phiếu).
 *************************************************/
function huyPhieuThu(soPhieuThu, lyDoHuy, currentUser) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);

    if (!currentUser || !currentUser.username) throw new Error('Thiếu thông tin người dùng.');
    _yeuCauQuyen(currentUser.username, [ROLE_ADMIN, ROLE_THU_QUY]);

    if (!soPhieuThu) throw new Error('Thiếu số phiếu thu cần hủy.');
    if (!lyDoHuy) throw new Error('Vui lòng nhập lý do hủy.');

    const sh = _sheet(SHEET_PHIEU_THU);
    const data = sh.getDataRange().getValues();
    const headers = data[0].map(h => String(h).trim());
    const idxSoPhieu = headers.indexOf('so_phieu_thu');
    const idxNgay = headers.indexOf('ngay_thu');
    const idxLoaiGD = headers.indexOf('loai_giao_dich');
    const idxTrangThai = headers.indexOf('trang_thai');
    const idxLyDoHuy = headers.indexOf('ly_do_huy');
    const idxNguoiSua = headers.indexOf('nguoi_sua_cuoi');
    const idxThoiGianSua = headers.indexOf('thoi_gian_sua_cuoi');

    let rowIndex = -1;
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][idxSoPhieu]) === String(soPhieuThu)) { rowIndex = i; break; }
    }
    if (rowIndex === -1) throw new Error('Không tìm thấy phiếu thu: ' + soPhieuThu);
    if (data[rowIndex][idxTrangThai] === TRANG_THAI_DA_HUY) throw new Error('Phiếu này đã bị hủy trước đó.');

    const ngayKey = _fmtDate(data[rowIndex][idxNgay]);
    if (_isDateLocked(ngayKey)) throw new Error('Ngày ' + ngayKey + ' đã bị khóa sổ, không thể hủy phiếu.');

    const loaiGD = data[rowIndex][idxLoaiGD];
    const nguoiLap = currentUser.full_name || 'N/A';
    const now = new Date();

    sh.getRange(rowIndex + 1, idxTrangThai + 1).setValue(TRANG_THAI_DA_HUY);
    sh.getRange(rowIndex + 1, idxLyDoHuy + 1).setValue(lyDoHuy);
    sh.getRange(rowIndex + 1, idxNguoiSua + 1).setValue(nguoiLap);
    sh.getRange(rowIndex + 1, idxThoiGianSua + 1).setValue(now);

    if (loaiGD === 'Thu công nợ') {
      const maCongNoDaHoanTac = _huyThanhToanCongNoTheoPhieu(soPhieuThu, nguoiLap);
      maCongNoDaHoanTac.forEach(ma => {
        _writeAuditLog(nguoiLap, 'Công Nợ', 'Sửa', ma, '', 'Hoàn tác thu nợ do hủy phiếu ' + soPhieuThu);
      });
    }

    _writeAuditLog(nguoiLap, 'Phiếu Thu', 'Hủy', soPhieuThu, TRANG_THAI_HOP_LE, 'Lý do: ' + lyDoHuy);

    return _jsonOk({ so_phieu_thu: soPhieuThu });

  } catch (err) {
    return _jsonErr(err);
  } finally {
    lock.releaseLock();
  }
}

function huyPhieuChi(soPhieuChi, lyDoHuy, currentUser) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);

    if (!currentUser || !currentUser.username) throw new Error('Thiếu thông tin người dùng.');
    _yeuCauQuyen(currentUser.username, [ROLE_ADMIN, ROLE_THU_QUY]);

    if (!soPhieuChi) throw new Error('Thiếu số phiếu chi cần hủy.');
    if (!lyDoHuy) throw new Error('Vui lòng nhập lý do hủy.');

    const sh = _sheet(SHEET_PHIEU_CHI);
    const data = sh.getDataRange().getValues();
    const headers = data[0].map(h => String(h).trim());
    const idxSoPhieu = headers.indexOf('so_phieu_chi');
    const idxNgay = headers.indexOf('ngay_chi');
    const idxLoaiGD = headers.indexOf('loai_giao_dich');
    const idxTrangThai = headers.indexOf('trang_thai');
    const idxLyDoHuy = headers.indexOf('ly_do_huy');
    const idxNguoiSua = headers.indexOf('nguoi_sua_cuoi');
    const idxThoiGianSua = headers.indexOf('thoi_gian_sua_cuoi');

    let rowIndex = -1;
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][idxSoPhieu]) === String(soPhieuChi)) { rowIndex = i; break; }
    }
    if (rowIndex === -1) throw new Error('Không tìm thấy phiếu chi: ' + soPhieuChi);
    if (data[rowIndex][idxTrangThai] === TRANG_THAI_DA_HUY) throw new Error('Phiếu này đã bị hủy trước đó.');

    const ngayKey = _fmtDate(data[rowIndex][idxNgay]);
    if (_isDateLocked(ngayKey)) throw new Error('Ngày ' + ngayKey + ' đã bị khóa sổ, không thể hủy phiếu.');

    const loaiGD = data[rowIndex][idxLoaiGD];
    const nguoiLap = currentUser.full_name || 'N/A';
    const now = new Date();

    sh.getRange(rowIndex + 1, idxTrangThai + 1).setValue(TRANG_THAI_DA_HUY);
    sh.getRange(rowIndex + 1, idxLyDoHuy + 1).setValue(lyDoHuy);
    sh.getRange(rowIndex + 1, idxNguoiSua + 1).setValue(nguoiLap);
    sh.getRange(rowIndex + 1, idxThoiGianSua + 1).setValue(now);

    if (loaiGD === 'Trả công nợ') {
      const maCongNoDaHoanTac = _huyThanhToanCongNoTheoPhieu(soPhieuChi, nguoiLap);
      maCongNoDaHoanTac.forEach(ma => {
        _writeAuditLog(nguoiLap, 'Công Nợ', 'Sửa', ma, '', 'Hoàn tác trả nợ do hủy phiếu ' + soPhieuChi);
      });
    }

    _writeAuditLog(nguoiLap, 'Phiếu Chi', 'Hủy', soPhieuChi, TRANG_THAI_HOP_LE, 'Lý do: ' + lyDoHuy);

    return _jsonOk({ so_phieu_chi: soPhieuChi });

  } catch (err) {
    return _jsonErr(err);
  } finally {
    lock.releaseLock();
  }
}

/*************************************************
 * API: SỔ QUỸ (tính động, có lọc) — không giới hạn quyền xem
 *
 * filters.loaiQuy chọn quỹ cần xem: 'Quỹ tiền mặt' (mặc định) hoặc
 * 'Quỹ công đoàn' — mỗi quỹ có số dư đầu kỳ và tồn lũy kế RIÊNG,
 * dùng chung cho cả Sổ Quỹ và Sổ Quỹ Công Đoàn ở giao diện.
 * Các cột trả về đầy đủ theo mẫu sổ kế toán chi tiết quỹ tiền mặt:
 * ngày hạch toán, ngày chứng từ, số phiếu thu/chi riêng cột, tài
 * khoản, tài khoản đối ứng, người nhận/nộp, mã nhân viên, chi nhánh.
 *************************************************/
function getSoQuy(filters) {
  try {
    filters = filters || {};
    const loaiQuy = _chuanHoaLoaiQuy(filters.loaiQuy);
    const soDuKhoiTao = Number(_getCauHinh(_cauHinhSoDuKhoiTaoTheoQuy(loaiQuy))) || 0;

    const doiTuongMap = {};
    _sheetToObjects(SHEET_DOITUONG).forEach(d => doiTuongMap[d.ma_doi_tuong] = d.ten_doi_tuong);

    const list = [];

    _sheetToObjects(SHEET_PHIEU_THU).forEach(p => {
      if (p.trang_thai !== TRANG_THAI_HOP_LE) return;
      if (_chuanHoaLoaiQuy(p.loai_quy) !== loaiQuy) return;
      const soPhieuThu = _safeText(p.so_phieu_thu);
      list.push({
        ngay: _fmtDate(p.ngay_thu), gio: p.gio_thu || '00:00',
        ngay_hach_toan: p.ngay_hach_toan ? _fmtDate(p.ngay_hach_toan) : _fmtDate(p.ngay_thu),
        so_phieu_thu: soPhieuThu, so_phieu_chi: '', so_phieu: soPhieuThu,
        noi_dung: _safeText(p.noi_dung_thu),
        tai_khoan: _safeText(p.tai_khoan) || '1111', tk_doi_ung: _safeText(p.tk_doi_ung),
        ma_doi_tuong: p.ma_doi_tuong,
        doi_tuong: doiTuongMap[p.ma_doi_tuong] || p.ma_doi_tuong || '',
        thu: Number(p.so_tien) || 0, chi: 0, loai: 'Thu',
        nguoi_nhan_nop: _safeText(p.nguoi_nop_tien),
        ma_nhan_vien: _safeText(p.ma_nhan_vien), chi_nhanh: _safeText(p.chi_nhanh),
        nguoi_lap: p.nguoi_lap, thoi_gian_lap: p.thoi_gian_lap
      });
    });

    _sheetToObjects(SHEET_PHIEU_CHI).forEach(p => {
      if (p.trang_thai !== TRANG_THAI_HOP_LE) return;
      if (_chuanHoaLoaiQuy(p.loai_quy) !== loaiQuy) return;
      const soPhieuChi = _safeText(p.so_phieu_chi);
      list.push({
        ngay: _fmtDate(p.ngay_chi), gio: p.gio_chi || '00:00',
        ngay_hach_toan: p.ngay_hach_toan ? _fmtDate(p.ngay_hach_toan) : _fmtDate(p.ngay_chi),
        so_phieu_thu: '', so_phieu_chi: soPhieuChi, so_phieu: soPhieuChi,
        noi_dung: _safeText(p.noi_dung_chi),
        tai_khoan: _safeText(p.tai_khoan) || '1111', tk_doi_ung: _safeText(p.tk_doi_ung),
        ma_doi_tuong: p.ma_doi_tuong,
        doi_tuong: doiTuongMap[p.ma_doi_tuong] || p.ma_doi_tuong || '',
        thu: 0, chi: Number(p.so_tien) || 0, loai: 'Chi',
        nguoi_nhan_nop: _safeText(p.nguoi_nhan_tien),
        ma_nhan_vien: _safeText(p.ma_nhan_vien), chi_nhanh: _safeText(p.chi_nhanh),
        nguoi_lap: p.nguoi_lap, thoi_gian_lap: p.thoi_gian_lap
      });
    });

    list.sort((a, b) => {
      const keyA = a.ngay + ' ' + a.gio;
      const keyB = b.ngay + ' ' + b.gio;
      if (keyA !== keyB) return keyA < keyB ? -1 : 1;
      const tA = a.thoi_gian_lap ? new Date(a.thoi_gian_lap).getTime() : 0;
      const tB = b.thoi_gian_lap ? new Date(b.thoi_gian_lap).getTime() : 0;
      return tA - tB;
    });

    let running = soDuKhoiTao;
    list.forEach(t => {
      running += t.thu - t.chi;
      t.ton = running;
    });

    let filtered = list;
    if (filters.tuNgay) filtered = filtered.filter(t => t.ngay >= filters.tuNgay);
    if (filters.denNgay) filtered = filtered.filter(t => t.ngay <= filters.denNgay);
    if (filters.loaiGiaoDich) filtered = filtered.filter(t => t.loai === filters.loaiGiaoDich);
    if (filters.nguoiLap) filtered = filtered.filter(t => String(t.nguoi_lap).toLowerCase().includes(String(filters.nguoiLap).toLowerCase()));
    if (filters.maDoiTuong) filtered = filtered.filter(t => t.ma_doi_tuong === filters.maDoiTuong);

    // Bỏ trường thoi_gian_lap (Date object thô, chỉ dùng để sắp xếp ở
    // trên) trước khi trả về client — 1 Date không hợp lệ lọt vào đây
    // có thể khiến phản hồi google.script.run bị hỏng ngầm (client
    // nhận về null dù server chạy đúng).
    const ketQua = filtered.map(t => {
      const o = Object.assign({}, t);
      delete o.thoi_gian_lap;
      return o;
    });

    return _jsonOk(ketQua);

  } catch (err) {
    return _jsonErr(err);
  }
}
