/*************************************************
 * SOCOM.GS
 * Sổ Cơm — theo dõi số suất ăn (trưa/tối) và tạm ứng tiền cơm
 * hàng ngày, mô phỏng đúng mẫu "cơm" trong file Excel gốc.
 *
 * Đây là một Quỹ riêng (không đi qua PHIEU_THU/PHIEU_CHI):
 * - "Thu" của quỹ cơm = các lần tạm ứng tiền cơm (so_tien_tam_ung).
 * - "Chi" của quỹ cơm = thành tiền suất ăn phát sinh mỗi ngày.
 * Tồn quỹ cơm = Số dư khởi tạo (CAU_HINH: SO_DU_QUY_COM_KHOI_TAO)
 * + lũy kế tạm ứng - lũy kế thành tiền, tính theo đúng thứ tự ngày.
 *
 * Quyền: ADMIN/THU_QUY được nhập/sửa; XEM chỉ xem (giống Sổ Quỹ).
 *************************************************/

function getDonGiaComMacDinh() {
  try {
    return _jsonOk({ don_gia: Number(_getCauHinh('DON_GIA_COM_MAC_DINH')) || 0 });
  } catch (err) {
    return _jsonErr(err);
  }
}

/**
 * Lấy toàn bộ Sổ Cơm của 1 tháng/năm, kèm tồn quỹ cơm lũy kế tính
 * trên TOÀN BỘ lịch sử (không chỉ riêng tháng đang xem) để đảm bảo
 * cột Tồn luôn đúng, giống nguyên tắc của Sổ Quỹ tiền mặt.
 */
function getSoComThang(nam, thang) {
  try {
    nam = Number(nam);
    thang = Number(thang);
    if (!nam || !thang || thang < 1 || thang > 12) throw new Error('Vui lòng chọn Tháng/Năm hợp lệ.');

    const donGiaMacDinh = Number(_getCauHinh('DON_GIA_COM_MAC_DINH')) || 0;
    const soDuKhoiTao = Number(_getCauHinh('SO_DU_QUY_COM_KHOI_TAO')) || 0;

    const all = _sheetToObjects(SHEET_SO_COM).map(r => ({
      ngay: _fmtDate(r.ngay),
      buoi_trua: Number(r.buoi_trua) || 0,
      buoi_toi: Number(r.buoi_toi) || 0,
      tong_suat: Number(r.tong_suat) || 0,
      don_gia: Number(r.don_gia) || donGiaMacDinh,
      thanh_tien: Number(r.thanh_tien) || 0,
      ngay_tam_ung: r.ngay_tam_ung ? _fmtDate(r.ngay_tam_ung) : '',
      so_tien_tam_ung: Number(r.so_tien_tam_ung) || 0,
      nguoi_lap: r.nguoi_lap || '',
      ghi_chu: r.ghi_chu || ''
    })).filter(r => r.ngay);

    all.sort((a, b) => a.ngay < b.ngay ? -1 : (a.ngay > b.ngay ? 1 : 0));

    let running = soDuKhoiTao;
    all.forEach(r => {
      running += r.so_tien_tam_ung - r.thanh_tien;
      r.ton = running;
    });

    const thangKey = String(nam) + '-' + String(thang).padStart(2, '0');
    const rows = all.filter(r => r.ngay.slice(0, 7) === thangKey);

    const tongSuat = rows.reduce((s, r) => s + r.tong_suat, 0);
    const tongThanhTien = rows.reduce((s, r) => s + r.thanh_tien, 0);
    const tongTamUng = rows.reduce((s, r) => s + r.so_tien_tam_ung, 0);

    return _jsonOk({
      don_gia_mac_dinh: donGiaMacDinh,
      rows: rows,
      tong_suat: tongSuat,
      tong_thanh_tien: tongThanhTien,
      tong_tam_ung: tongTamUng,
      ton_hien_tai: running
    });

  } catch (err) {
    return _jsonErr(err);
  }
}

/**
 * Thêm mới hoặc cập nhật (nếu ngày đã tồn tại) 1 dòng Sổ Cơm.
 * payload = { ngay, buoi_trua, buoi_toi, don_gia, ngay_tam_ung,
 *             so_tien_tam_ung, ghi_chu }
 */
function upsertSoComNgay(payload, currentUser) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);

    if (!currentUser || !currentUser.username) throw new Error('Thiếu thông tin người dùng.');
    _yeuCauQuyen(currentUser.username, [ROLE_ADMIN, ROLE_THU_QUY]);

    if (!payload || !payload.ngay) throw new Error('Vui lòng chọn Ngày.');
    const ngayKey = Utilities.formatDate(new Date(payload.ngay), Session.getScriptTimeZone(), 'yyyy-MM-dd');

    const buoiTrua = Number(payload.buoi_trua) || 0;
    const buoiToi = Number(payload.buoi_toi) || 0;
    if (buoiTrua < 0 || buoiToi < 0) throw new Error('Số suất ăn không hợp lệ.');

    const donGiaMacDinh = Number(_getCauHinh('DON_GIA_COM_MAC_DINH')) || 0;
    const donGia = Number(payload.don_gia) || donGiaMacDinh;
    if (donGia < 0) throw new Error('Đơn giá không hợp lệ.');

    const tongSuat = buoiTrua + buoiToi;
    const thanhTien = tongSuat * donGia;

    const soTienTamUng = Number(payload.so_tien_tam_ung) || 0;
    if (soTienTamUng < 0) throw new Error('Số tiền tạm ứng không hợp lệ.');
    const ngayTamUngKey = payload.ngay_tam_ung
      ? Utilities.formatDate(new Date(payload.ngay_tam_ung), Session.getScriptTimeZone(), 'yyyy-MM-dd')
      : (soTienTamUng > 0 ? ngayKey : '');
    if (soTienTamUng > 0 && !ngayTamUngKey) throw new Error('Vui lòng chọn Ngày tạm ứng.');

    const nguoiLap = currentUser.full_name || 'N/A';
    const now = new Date();

    const sh = _sheet(SHEET_SO_COM);
    const data = sh.getDataRange().getValues();
    const headers = data[0].map(h => String(h).trim());
    const idxId = headers.indexOf('id');
    const idxNgay = headers.indexOf('ngay');

    let rowIndex = -1;
    for (let i = 1; i < data.length; i++) {
      if (!data[i][idxNgay]) continue;
      if (_fmtDate(data[i][idxNgay]) === ngayKey) { rowIndex = i; break; }
    }

    const id = rowIndex === -1 ? ('COM_' + now.getTime()) : data[rowIndex][idxId];
    const rowValues = [
      id, new Date(ngayKey), buoiTrua, buoiToi, tongSuat, donGia, thanhTien,
      ngayTamUngKey ? new Date(ngayTamUngKey) : '', soTienTamUng,
      nguoiLap, now, payload.ghi_chu || ''
    ];

    if (rowIndex === -1) {
      sh.appendRow(rowValues);
    } else {
      sh.getRange(rowIndex + 1, 1, 1, rowValues.length).setValues([rowValues]);
    }

    _writeAuditLog(nguoiLap, 'Sổ Cơm', rowIndex === -1 ? 'Thêm' : 'Sửa', ngayKey, '',
      'Suất: ' + tongSuat + ' (trưa ' + buoiTrua + ', tối ' + buoiToi + ') | Thành tiền: ' +
      thanhTien.toLocaleString('vi-VN') + ' đ' +
      (soTienTamUng ? ' | Tạm ứng: ' + soTienTamUng.toLocaleString('vi-VN') + ' đ' : ''));

    return _jsonOk({ ngay: ngayKey });

  } catch (err) {
    return _jsonErr(err);
  } finally {
    lock.releaseLock();
  }
}
