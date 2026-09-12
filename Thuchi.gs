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
 * Duyệt TOÀN BỘ lịch sử giao dịch hợp lệ (kể cả giao dịch giả định mới,
 * nếu truyền vào), tính tồn quỹ lũy kế theo đúng thứ tự thời gian,
 * trả về giá trị tồn THẤP NHẤT từng xuất hiện trong toàn bộ dòng thời gian.
 */
function _timTonThapNhat(giaoDichGiaDinh) {
  const soDuKhoiTao = Number(_getCauHinh('SO_DU_QUY_KHOI_TAO')) || 0;
  const list = [];

  _sheetToObjects(SHEET_PHIEU_THU).forEach(p => {
    if (p.trang_thai === TRANG_THAI_HOP_LE) {
      list.push({
        ngay: _fmtDate(p.ngay_thu), gio: p.gio_thu || '00:00',
        thoiGianLap: p.thoi_gian_lap, soTien: Number(p.so_tien) || 0, loai: 'Thu'
      });
    }
  });
  _sheetToObjects(SHEET_PHIEU_CHI).forEach(p => {
    if (p.trang_thai === TRANG_THAI_HOP_LE) {
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
      nguoiLap, now, '', '', TRANG_THAI_HOP_LE, ''
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

    const tonThapNhatNeuChi = _timTonThapNhat({ ngay: ngayKey, soTien: soTien, loai: 'Chi' });
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
      nguoiLap, now, '', '', TRANG_THAI_HOP_LE, ''
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
 * API: SỔ QUỸ (tính động, có lọc) — không giới hạn quyền xem
 *************************************************/
function getSoQuy(filters) {
  try {
    filters = filters || {};
    const soDuKhoiTao = Number(_getCauHinh('SO_DU_QUY_KHOI_TAO')) || 0;

    const doiTuongMap = {};
    _sheetToObjects(SHEET_DOITUONG).forEach(d => doiTuongMap[d.ma_doi_tuong] = d.ten_doi_tuong);

    const list = [];

    _sheetToObjects(SHEET_PHIEU_THU).forEach(p => {
      if (p.trang_thai !== TRANG_THAI_HOP_LE) return;
      list.push({
        ngay: _fmtDate(p.ngay_thu), gio: p.gio_thu || '00:00',
        so_phieu: p.so_phieu_thu, noi_dung: p.noi_dung_thu,
        ma_doi_tuong: p.ma_doi_tuong,
        doi_tuong: doiTuongMap[p.ma_doi_tuong] || p.ma_doi_tuong || '',
        thu: Number(p.so_tien) || 0, chi: 0, loai: 'Thu',
        nguoi_lap: p.nguoi_lap, thoi_gian_lap: p.thoi_gian_lap
      });
    });

    _sheetToObjects(SHEET_PHIEU_CHI).forEach(p => {
      if (p.trang_thai !== TRANG_THAI_HOP_LE) return;
      list.push({
        ngay: _fmtDate(p.ngay_chi), gio: p.gio_chi || '00:00',
        so_phieu: p.so_phieu_chi, noi_dung: p.noi_dung_chi,
        ma_doi_tuong: p.ma_doi_tuong,
        doi_tuong: doiTuongMap[p.ma_doi_tuong] || p.ma_doi_tuong || '',
        thu: 0, chi: Number(p.so_tien) || 0, loai: 'Chi',
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

    return _jsonOk(filtered);

  } catch (err) {
    return _jsonErr(err);
  }
}
