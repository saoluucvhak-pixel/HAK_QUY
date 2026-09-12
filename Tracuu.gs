/*************************************************
 * TRACUU.GS
 * Tra cứu Phiếu Thu/Chi theo nhiều tiêu chí, xem chi tiết
 * 1 giao dịch kèm toàn bộ lịch sử thao tác (từ Audit Log).
 * Không giới hạn quyền xem — đây thuần túy là tra cứu/xem
 * lại, giống Sổ Quỹ, dành cho cả ADMIN/THU_QUY/XEM.
 *************************************************/

function traCuuGiaoDich(filters) {
  try {
    filters = filters || {};
    const doiTuongMap = {};
    _sheetToObjects(SHEET_DOITUONG).forEach(d => doiTuongMap[d.ma_doi_tuong] = d.ten_doi_tuong);

    const list = [];

    _getPhieuThuData().forEach(p => {
      list.push({
        so_phieu: p.so_phieu_thu, loai: 'Thu', ngay: _fmtDate(p.ngay_thu),
        noi_dung: p.noi_dung_thu, ma_doi_tuong: p.ma_doi_tuong,
        doi_tuong: doiTuongMap[p.ma_doi_tuong] || '',
        so_tien: Number(p.so_tien) || 0, nguoi_lap: p.nguoi_lap, trang_thai: p.trang_thai
      });
    });
    _getPhieuChiData().forEach(p => {
      list.push({
        so_phieu: p.so_phieu_chi, loai: 'Chi', ngay: _fmtDate(p.ngay_chi),
        noi_dung: p.noi_dung_chi, ma_doi_tuong: p.ma_doi_tuong,
        doi_tuong: doiTuongMap[p.ma_doi_tuong] || '',
        so_tien: Number(p.so_tien) || 0, nguoi_lap: p.nguoi_lap, trang_thai: p.trang_thai
      });
    });

    let filtered = list;
    if (filters.soPhieu) {
      const kw = String(filters.soPhieu).toLowerCase();
      filtered = filtered.filter(t => String(t.so_phieu).toLowerCase().includes(kw));
    }
    if (filters.tuNgay) filtered = filtered.filter(t => t.ngay >= filters.tuNgay);
    if (filters.denNgay) filtered = filtered.filter(t => t.ngay <= filters.denNgay);
    if (filters.maDoiTuong) filtered = filtered.filter(t => t.ma_doi_tuong === filters.maDoiTuong);
    if (filters.noiDung) {
      const kw = String(filters.noiDung).toLowerCase();
      filtered = filtered.filter(t => String(t.noi_dung).toLowerCase().includes(kw));
    }
    if (filters.soTien) filtered = filtered.filter(t => t.so_tien === Number(filters.soTien));
    if (filters.nguoiLap) {
      const kw = String(filters.nguoiLap).toLowerCase();
      filtered = filtered.filter(t => String(t.nguoi_lap).toLowerCase().includes(kw));
    }

    filtered.sort((a, b) => (a.ngay < b.ngay) ? 1 : (a.ngay > b.ngay ? -1 : 0));

    // Giới hạn 100 kết quả để tránh phản hồi quá lớn khi tìm kiếm quá rộng
    const tongSoKetQua = filtered.length;
    filtered = filtered.slice(0, 100);

    return _jsonOk({ ket_qua: filtered, tong_so: tongSoKetQua });

  } catch (err) {
    return _jsonErr(err);
  }
}

function getChiTietGiaoDich(soPhieu, loai) {
  try {
    if (!soPhieu || !loai) throw new Error('Thiếu thông tin tra cứu.');

    const doiTuongMap = {};
    _sheetToObjects(SHEET_DOITUONG).forEach(d => doiTuongMap[d.ma_doi_tuong] = d.ten_doi_tuong);

    let phieu = null;

    if (loai === 'Thu') {
      const row = _getPhieuThuData().find(p => p.so_phieu_thu === soPhieu);
      if (row) {
        phieu = {
          so_phieu: row.so_phieu_thu, loai: 'Thu', ngay: _fmtDate(row.ngay_thu), gio: row.gio_thu,
          loai_giao_dich: row.loai_giao_dich, nguoi_nop_nhan: row.nguoi_nop_tien,
          doi_tuong: doiTuongMap[row.ma_doi_tuong] || '', noi_dung: row.noi_dung_thu,
          so_tien: Number(row.so_tien) || 0, chung_tu: row.chung_tu_lien_quan, ghi_chu: row.ghi_chu,
          nguoi_lap: row.nguoi_lap, thoi_gian_lap: _fmtDateTime(row.thoi_gian_lap),
          trang_thai: row.trang_thai, ly_do_huy: row.ly_do_huy
        };
      }
    } else if (loai === 'Chi') {
      const row = _getPhieuChiData().find(p => p.so_phieu_chi === soPhieu);
      if (row) {
        phieu = {
          so_phieu: row.so_phieu_chi, loai: 'Chi', ngay: _fmtDate(row.ngay_chi), gio: row.gio_chi,
          loai_giao_dich: row.loai_giao_dich, nguoi_nop_nhan: row.nguoi_nhan_tien,
          doi_tuong: doiTuongMap[row.ma_doi_tuong] || '', noi_dung: row.noi_dung_chi,
          so_tien: Number(row.so_tien) || 0, chung_tu: row.chung_tu_lien_quan, ghi_chu: row.ghi_chu,
          nguoi_lap: row.nguoi_lap, thoi_gian_lap: _fmtDateTime(row.thoi_gian_lap),
          trang_thai: row.trang_thai, ly_do_huy: row.ly_do_huy
        };
      }
    } else {
      throw new Error('Loại giao dịch không hợp lệ.');
    }

    if (!phieu) throw new Error('Không tìm thấy chứng từ: ' + soPhieu);

    const lichSu = _sheetToObjects(SHEET_AUDIT_LOG)
      .filter(l => l.ma_chung_tu === soPhieu)
      .map(l => ({
        thoi_gian: _fmtDateTime(l.thoi_gian),
        nguoi_thao_tac: l.nguoi_thao_tac,
        loai_thao_tac: l.loai_thao_tac,
        du_lieu_truoc: l.du_lieu_truoc,
        du_lieu_sau: l.du_lieu_sau,
        __row: l.__row
      }))
      .sort((a, b) => b.__row - a.__row);

    return _jsonOk({ phieu: phieu, lich_su: lichSu });

  } catch (err) {
    return _jsonErr(err);
  }
}
