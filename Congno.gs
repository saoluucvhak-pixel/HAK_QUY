/*************************************************
 * CONGNO.GS
 * Quản lý Công Nợ Phải Thu / Phải Trả.
 *
 * TỪ GIAI ĐOẠN 5: Tạo khoản công nợ mới CHỈ ADMIN
 * được phép (THỦ QUỸ chỉ được lập Phiếu Thu/Chi, không
 * được quản lý Công Nợ trực tiếp — theo đúng bảng phân
 * quyền mục 16). currentUser = { username, full_name }.
 *************************************************/

/*************************************************
 * API: DANH SÁCH CÔNG NỢ (có lọc) — không giới hạn quyền xem
 *************************************************/
function getCongNoList(filters) {
  try {
    filters = filters || {};
    const doiTuongMap = {};
    _sheetToObjects(SHEET_DOITUONG).forEach(d => doiTuongMap[d.ma_doi_tuong] = d.ten_doi_tuong);

    let list = _sheetToObjects(SHEET_CONG_NO).map(c => ({
      ma_cong_no: c.ma_cong_no,
      ngay_phat_sinh: _fmtDate(c.ngay_phat_sinh),
      ma_doi_tuong: c.ma_doi_tuong,
      doi_tuong: doiTuongMap[c.ma_doi_tuong] || c.ma_doi_tuong,
      loai_cong_no: c.loai_cong_no,
      noi_dung: c.noi_dung,
      so_tien_phat_sinh: Number(c.so_tien_phat_sinh) || 0,
      han_thanh_toan: c.han_thanh_toan ? _fmtDate(c.han_thanh_toan) : '',
      da_thanh_toan: Number(c.da_thanh_toan) || 0,
      con_lai: Number(c.con_lai) || 0,
      trang_thai: c.trang_thai,
      nguoi_lap: c.nguoi_lap,
      __row: c.__row
    }));

    if (filters.loaiCongNo) list = list.filter(c => c.loai_cong_no === filters.loaiCongNo);
    if (filters.trangThai) list = list.filter(c => c.trang_thai === filters.trangThai);
    if (filters.maDoiTuong) list = list.filter(c => c.ma_doi_tuong === filters.maDoiTuong);

    list.sort((a, b) => b.__row - a.__row);

    return _jsonOk(list);
  } catch (err) {
    return _jsonErr(err);
  }
}

/*************************************************
 * API: TẠO KHOẢN CÔNG NỢ MỚI — CHỈ ADMIN
 *************************************************/
function addCongNo(payload, currentUser) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);

    if (!currentUser || !currentUser.username) throw new Error('Thiếu thông tin người dùng.');
    _yeuCauQuyen(currentUser.username, [ROLE_ADMIN]);

    if (!payload) throw new Error('Thiếu dữ liệu công nợ.');
    if (!payload.ngay_phat_sinh) throw new Error('Vui lòng chọn Ngày phát sinh.');
    if (!payload.ma_doi_tuong) throw new Error('Vui lòng chọn Đối tượng.');

    const loai = payload.loai_cong_no;
    if (loai !== 'Phải thu' && loai !== 'Phải trả') throw new Error('Loại công nợ không hợp lệ.');

    const soTien = Number(payload.so_tien_phat_sinh);
    if (!soTien || soTien <= 0) throw new Error('Số tiền phát sinh không hợp lệ.');

    const maCongNo = _generateNextNumber('SO_CONG_NO_TIEP_THEO', 'CN');
    const now = new Date();
    const nguoiLap = currentUser.full_name || 'N/A';

    _sheet(SHEET_CONG_NO).appendRow([
      maCongNo,
      new Date(payload.ngay_phat_sinh),
      payload.ma_doi_tuong,
      loai,
      payload.noi_dung || '',
      soTien,
      payload.han_thanh_toan ? new Date(payload.han_thanh_toan) : '',
      0,
      soTien,
      'Còn nợ',
      nguoiLap,
      now
    ]);

    _writeAuditLog(nguoiLap, 'Công Nợ', 'Thêm', maCongNo, '',
      'Loại: ' + loai + ' | Số tiền: ' + soTien.toLocaleString('vi-VN') + ' đ | Nội dung: ' + (payload.noi_dung || ''));

    return _jsonOk({ ma_cong_no: maCongNo });

  } catch (err) {
    return _jsonErr(err);
  } finally {
    lock.releaseLock();
  }
}

/*************************************************
 * API: LẤY DANH SÁCH CÔNG NỢ CÒN NỢ CỦA 1 ĐỐI TƯỢNG
 * (không giới hạn quyền xem — THU_QUY cần xem để lập Phiếu Thu/Chi)
 *************************************************/
function getCongNoMoOptions(maDoiTuong, loaiCanTim) {
  try {
    if (!maDoiTuong) return _jsonOk([]);
    const list = _sheetToObjects(SHEET_CONG_NO)
      .filter(c => c.ma_doi_tuong === maDoiTuong && c.loai_cong_no === loaiCanTim && c.trang_thai === 'Còn nợ')
      .map(c => ({
        ma_cong_no: c.ma_cong_no,
        noi_dung: c.noi_dung,
        con_lai: Number(c.con_lai) || 0,
        ngay_phat_sinh: _fmtDate(c.ngay_phat_sinh)
      }));
    return _jsonOk(list);
  } catch (err) {
    return _jsonErr(err);
  }
}

/**
 * Cập nhật 1 khoản công nợ sau khi có thanh toán. Hàm nội bộ — luôn
 * được gọi từ ThuChi.gs, nằm trong lock đã giữ sẵn nên KHÔNG tự
 * tạo lock riêng. Quyền đã được kiểm tra ở addPhieuThu/addPhieuChi
 * (ADMIN hoặc THU_QUY được lập phiếu, kể cả loại Thu/Trả công nợ).
 */
function _capNhatCongNoSauThanhToan(maCongNo, soTienThanhToan, loaiThanhToan, soPhieuLienQuan, nguoiLap) {
  const sh = _sheet(SHEET_CONG_NO);
  const data = sh.getDataRange().getValues();
  const headers = data[0].map(h => String(h).trim());
  const idxMa = headers.indexOf('ma_cong_no');
  const idxDaTT = headers.indexOf('da_thanh_toan');
  const idxConLai = headers.indexOf('con_lai');
  const idxTrangThai = headers.indexOf('trang_thai');

  let rowIndex = -1;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idxMa]) === String(maCongNo)) { rowIndex = i; break; }
  }
  if (rowIndex === -1) throw new Error('Không tìm thấy khoản công nợ: ' + maCongNo);

  const conLaiHienTai = Number(data[rowIndex][idxConLai]) || 0;
  const daTTMoi = (Number(data[rowIndex][idxDaTT]) || 0) + soTienThanhToan;
  const conLaiMoi = conLaiHienTai - soTienThanhToan;
  const trangThaiMoi = conLaiMoi <= 0 ? 'Đã tất toán' : 'Còn nợ';

  sh.getRange(rowIndex + 1, idxDaTT + 1).setValue(daTTMoi);
  sh.getRange(rowIndex + 1, idxConLai + 1).setValue(conLaiMoi);
  sh.getRange(rowIndex + 1, idxTrangThai + 1).setValue(trangThaiMoi);

  _sheet(SHEET_THANH_TOAN_CONG_NO).appendRow([
    'TT' + new Date().getTime(),
    maCongNo,
    new Date(),
    soTienThanhToan,
    loaiThanhToan,
    soPhieuLienQuan,
    nguoiLap,
    new Date()
  ]);

  return conLaiMoi;
}

/**
 * Hoàn tác lại khoản công nợ đã thanh toán qua 1 phiếu Thu/Chi, dùng
 * khi Hủy phiếu đó (loại giao dịch "Thu công nợ"/"Trả công nợ"). Cộng
 * lại đúng số tiền vào "còn lại", trừ lại "đã thanh toán", và ghi 1
 * dòng đảo ngược vào THANH_TOAN_CONG_NO để giữ đầy đủ lịch sử (không
 * xóa dòng thanh toán cũ). Hàm nội bộ — được gọi từ trong lock đã giữ
 * sẵn ở huyPhieuThu/huyPhieuChi.
 */
function _huyThanhToanCongNoTheoPhieu(soPhieu, nguoiLap) {
  const danhSachThanhToan = _sheetToObjects(SHEET_THANH_TOAN_CONG_NO)
    .filter(t => t.so_phieu_lien_quan === soPhieu);
  if (danhSachThanhToan.length === 0) return [];

  const shCongNo = _sheet(SHEET_CONG_NO);
  const data = shCongNo.getDataRange().getValues();
  const headers = data[0].map(h => String(h).trim());
  const idxMa = headers.indexOf('ma_cong_no');
  const idxDaTT = headers.indexOf('da_thanh_toan');
  const idxConLai = headers.indexOf('con_lai');
  const idxTrangThai = headers.indexOf('trang_thai');

  const shThanhToan = _sheet(SHEET_THANH_TOAN_CONG_NO);
  const maCongNoDaXuLy = [];

  danhSachThanhToan.forEach(t => {
    const soTien = Number(t.so_tien_thanh_toan) || 0;
    let rowIndex = -1;
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][idxMa]) === String(t.ma_cong_no)) { rowIndex = i; break; }
    }
    if (rowIndex === -1) return;

    const daTTMoi = (Number(data[rowIndex][idxDaTT]) || 0) - soTien;
    const conLaiMoi = (Number(data[rowIndex][idxConLai]) || 0) + soTien;
    shCongNo.getRange(rowIndex + 1, idxDaTT + 1).setValue(daTTMoi);
    shCongNo.getRange(rowIndex + 1, idxConLai + 1).setValue(conLaiMoi);
    shCongNo.getRange(rowIndex + 1, idxTrangThai + 1).setValue(conLaiMoi > 0 ? 'Còn nợ' : 'Đã tất toán');
    data[rowIndex][idxDaTT] = daTTMoi;
    data[rowIndex][idxConLai] = conLaiMoi;

    shThanhToan.appendRow([
      'TT' + new Date().getTime(), t.ma_cong_no, new Date(), -soTien,
      'Hủy ' + t.loai, soPhieu + ' (hủy)', nguoiLap, new Date()
    ]);
    maCongNoDaXuLy.push(t.ma_cong_no);
  });

  return maCongNoDaXuLy;
}
