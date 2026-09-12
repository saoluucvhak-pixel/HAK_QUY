/*************************************************
 * BAOCAO.GS
 * Báo cáo Thu-Chi và Công nợ.
 *
 * NGUYÊN TẮC: Tồn đầu kỳ/cuối kỳ luôn tính trên TOÀN BỘ
 * giao dịch hợp lệ (không bị ảnh hưởng bởi bộ lọc phụ
 * như loại/đối tượng/người lập) để đảm bảo công thức
 * Tồn cuối kỳ = Tồn đầu kỳ + Tổng thu - Tổng chi luôn
 * đúng. Bộ lọc phụ chỉ áp dụng cho phần "Chi tiết".
 *************************************************/

/*************************************************
 * API: BÁO CÁO THU - CHI
 * filters = { tuNgay, denNgay, loaiGiaoDich, maDoiTuong,
 *             nguoiLap, maLoai }
 *************************************************/
function getBaoCaoThuChi(filters) {
  try {
    filters = filters || {};
    if (!filters.tuNgay || !filters.denNgay) throw new Error('Vui lòng chọn khoảng ngày báo cáo.');

    const soDuKhoiTao = Number(_getCauHinh('SO_DU_QUY_KHOI_TAO')) || 0;
    const doiTuongMap = {};
    _sheetToObjects(SHEET_DOITUONG).forEach(d => doiTuongMap[d.ma_doi_tuong] = d.ten_doi_tuong);

    const allThu = _sheetToObjects(SHEET_PHIEU_THU).filter(p => p.trang_thai === TRANG_THAI_HOP_LE);
    const allChi = _sheetToObjects(SHEET_PHIEU_CHI).filter(p => p.trang_thai === TRANG_THAI_HOP_LE);

    // Tồn đầu kỳ: TOÀN BỘ giao dịch trước tuNgay (không lọc phụ)
    let tonDauKy = soDuKhoiTao;
    allThu.forEach(p => { if (_fmtDate(p.ngay_thu) < filters.tuNgay) tonDauKy += Number(p.so_tien) || 0; });
    allChi.forEach(p => { if (_fmtDate(p.ngay_chi) < filters.tuNgay) tonDauKy -= Number(p.so_tien) || 0; });

    // Tổng thu/chi trong kỳ: TOÀN BỘ giao dịch trong khoảng ngày (không lọc phụ)
    let tongThuTrongKy = 0, tongChiTrongKy = 0;
    allThu.forEach(p => {
      const d = _fmtDate(p.ngay_thu);
      if (d >= filters.tuNgay && d <= filters.denNgay) tongThuTrongKy += Number(p.so_tien) || 0;
    });
    allChi.forEach(p => {
      const d = _fmtDate(p.ngay_chi);
      if (d >= filters.tuNgay && d <= filters.denNgay) tongChiTrongKy += Number(p.so_tien) || 0;
    });

    const tonCuoiKy = tonDauKy + tongThuTrongKy - tongChiTrongKy;

    // Chi tiết: áp dụng ĐẦY ĐỦ bộ lọc phụ
    const chiTiet = [];
    allThu.forEach(p => {
      const d = _fmtDate(p.ngay_thu);
      if (d < filters.tuNgay || d > filters.denNgay) return;
      if (filters.loaiGiaoDich && filters.loaiGiaoDich !== 'Thu') return;
      if (filters.maDoiTuong && p.ma_doi_tuong !== filters.maDoiTuong) return;
      if (filters.nguoiLap && !String(p.nguoi_lap).toLowerCase().includes(String(filters.nguoiLap).toLowerCase())) return;
      if (filters.maLoai && p.ma_loai_thu !== filters.maLoai) return;

      chiTiet.push({
        ngay: d, so_phieu: p.so_phieu_thu, loai: 'Thu',
        noi_dung: p.noi_dung_thu, doi_tuong: doiTuongMap[p.ma_doi_tuong] || '',
        thu: Number(p.so_tien) || 0, chi: 0, nguoi_lap: p.nguoi_lap
      });
    });
    allChi.forEach(p => {
      const d = _fmtDate(p.ngay_chi);
      if (d < filters.tuNgay || d > filters.denNgay) return;
      if (filters.loaiGiaoDich && filters.loaiGiaoDich !== 'Chi') return;
      if (filters.maDoiTuong && p.ma_doi_tuong !== filters.maDoiTuong) return;
      if (filters.nguoiLap && !String(p.nguoi_lap).toLowerCase().includes(String(filters.nguoiLap).toLowerCase())) return;
      if (filters.maLoai && p.ma_loai_chi !== filters.maLoai) return;

      chiTiet.push({
        ngay: d, so_phieu: p.so_phieu_chi, loai: 'Chi',
        noi_dung: p.noi_dung_chi, doi_tuong: doiTuongMap[p.ma_doi_tuong] || '',
        thu: 0, chi: Number(p.so_tien) || 0, nguoi_lap: p.nguoi_lap
      });
    });

    chiTiet.sort((a, b) => a.ngay < b.ngay ? -1 : (a.ngay > b.ngay ? 1 : 0));

    return _jsonOk({
      tu_ngay: filters.tuNgay, den_ngay: filters.denNgay,
      ton_dau_ky: tonDauKy, tong_thu: tongThuTrongKy, tong_chi: tongChiTrongKy, ton_cuoi_ky: tonCuoiKy,
      chi_tiet: chiTiet
    });

  } catch (err) {
    return _jsonErr(err);
  }
}

/*************************************************
 * API: BÁO CÁO CÔNG NỢ TỔNG HỢP (theo 1 loại: Phải thu HOẶC Phải trả)
 *************************************************/
function getBaoCaoCongNoTongHop(loaiCongNo, tuNgay, denNgay) {
  try {
    if (!tuNgay || !denNgay) throw new Error('Vui lòng chọn khoảng ngày báo cáo.');
    if (loaiCongNo !== 'Phải thu' && loaiCongNo !== 'Phải trả') throw new Error('Loại công nợ không hợp lệ.');

    const doiTuongMap = {};
    _sheetToObjects(SHEET_DOITUONG).forEach(d => doiTuongMap[d.ma_doi_tuong] = d.ten_doi_tuong);

    const congNoList = _sheetToObjects(SHEET_CONG_NO).filter(c => c.loai_cong_no === loaiCongNo);
    const congNoMap = {};
    congNoList.forEach(c => congNoMap[c.ma_cong_no] = c.ma_doi_tuong);

    const thanhToanList = _sheetToObjects(SHEET_THANH_TOAN_CONG_NO)
      .filter(t => congNoMap.hasOwnProperty(t.ma_cong_no));

    const acc = {};
    function ensure(ma) {
      if (!acc[ma]) acc[ma] = { dauKy: 0, phatSinh: 0, thanhToan: 0 };
      return acc[ma];
    }

    congNoList.forEach(c => {
      const ngay = _fmtDate(c.ngay_phat_sinh);
      const soTien = Number(c.so_tien_phat_sinh) || 0;
      const a = ensure(c.ma_doi_tuong);
      if (ngay < tuNgay) a.dauKy += soTien;
      else if (ngay <= denNgay) a.phatSinh += soTien;
    });

    thanhToanList.forEach(t => {
      const ngay = _fmtDate(t.ngay_thanh_toan);
      const soTien = Number(t.so_tien_thanh_toan) || 0;
      const maDoiTuong = congNoMap[t.ma_cong_no];
      const a = ensure(maDoiTuong);
      if (ngay < tuNgay) a.dauKy -= soTien;
      else if (ngay <= denNgay) a.thanhToan += soTien;
    });

    const result = Object.keys(acc)
      .map(ma => {
        const a = acc[ma];
        return {
          ma_doi_tuong: ma,
          doi_tuong: doiTuongMap[ma] || ma,
          dau_ky: a.dauKy,
          phat_sinh: a.phatSinh,
          da_thu_tra: a.thanhToan,
          cuoi_ky: a.dauKy + a.phatSinh - a.thanhToan
        };
      })
      .filter(r => r.dau_ky !== 0 || r.phat_sinh !== 0 || r.da_thu_tra !== 0 || r.cuoi_ky !== 0);

    result.sort((a, b) => b.cuoi_ky - a.cuoi_ky);

    return _jsonOk(result);

  } catch (err) {
    return _jsonErr(err);
  }
}

/*************************************************
 * API: CHI TIẾT CÔNG NỢ CỦA 1 ĐỐI TƯỢNG TRONG KỲ
 *************************************************/
function getBaoCaoCongNoChiTiet(maDoiTuong, loaiCongNo, tuNgay, denNgay) {
  try {
    if (!maDoiTuong || !loaiCongNo || !tuNgay || !denNgay) throw new Error('Thiếu tham số báo cáo chi tiết.');

    const congNoList = _sheetToObjects(SHEET_CONG_NO)
      .filter(c => c.ma_doi_tuong === maDoiTuong && c.loai_cong_no === loaiCongNo);
    const congNoMap = {};
    congNoList.forEach(c => congNoMap[c.ma_cong_no] = true);

    const thanhToanList = _sheetToObjects(SHEET_THANH_TOAN_CONG_NO)
      .filter(t => congNoMap.hasOwnProperty(t.ma_cong_no));

    let dauKy = 0;
    const phatSinhTrongKy = [];
    congNoList.forEach(c => {
      const ngay = _fmtDate(c.ngay_phat_sinh);
      const soTien = Number(c.so_tien_phat_sinh) || 0;
      if (ngay < tuNgay) dauKy += soTien;
      else if (ngay <= denNgay) phatSinhTrongKy.push({ ma_cong_no: c.ma_cong_no, ngay: ngay, noi_dung: c.noi_dung, so_tien: soTien });
    });

    const thanhToanTrongKy = [];
    thanhToanList.forEach(t => {
      const ngay = _fmtDate(t.ngay_thanh_toan);
      const soTien = Number(t.so_tien_thanh_toan) || 0;
      if (ngay < tuNgay) dauKy -= soTien;
      else if (ngay <= denNgay) thanhToanTrongKy.push({ ngay: ngay, so_tien: soTien, so_phieu: t.so_phieu_lien_quan, loai: t.loai });
    });

    const tongPhatSinh = phatSinhTrongKy.reduce((s, x) => s + x.so_tien, 0);
    const tongThanhToan = thanhToanTrongKy.reduce((s, x) => s + x.so_tien, 0);
    const cuoiKy = dauKy + tongPhatSinh - tongThanhToan;

    phatSinhTrongKy.sort((a, b) => a.ngay < b.ngay ? -1 : 1);
    thanhToanTrongKy.sort((a, b) => a.ngay < b.ngay ? -1 : 1);

    return _jsonOk({ dau_ky: dauKy, phat_sinh: phatSinhTrongKy, thanh_toan: thanhToanTrongKy, cuoi_ky: cuoiKy });

  } catch (err) {
    return _jsonErr(err);
  }
}

/*************************************************
 * API: CÔNG NỢ QUÁ HẠN
 *************************************************/
function getBaoCaoCongNoQuaHan() {
  try {
    const todayKey = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    const doiTuongMap = {};
    _sheetToObjects(SHEET_DOITUONG).forEach(d => doiTuongMap[d.ma_doi_tuong] = d.ten_doi_tuong);

    const list = _sheetToObjects(SHEET_CONG_NO)
      .filter(c => c.trang_thai === 'Còn nợ' && c.han_thanh_toan && _fmtDate(c.han_thanh_toan) < todayKey)
      .map(c => {
        const hanKey = _fmtDate(c.han_thanh_toan);
        const soNgayQuaHan = Math.floor((new Date(todayKey) - new Date(hanKey)) / (1000 * 60 * 60 * 24));
        return {
          ma_cong_no: c.ma_cong_no,
          doi_tuong: doiTuongMap[c.ma_doi_tuong] || c.ma_doi_tuong,
          loai_cong_no: c.loai_cong_no,
          noi_dung: c.noi_dung,
          ngay_phat_sinh: _fmtDate(c.ngay_phat_sinh),
          han_thanh_toan: hanKey,
          so_ngay_qua_han: soNgayQuaHan,
          con_lai: Number(c.con_lai) || 0
        };
      });

    list.sort((a, b) => b.so_ngay_qua_han - a.so_ngay_qua_han);

    return _jsonOk(list);

  } catch (err) {
    return _jsonErr(err);
  }
}
