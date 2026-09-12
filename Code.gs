/*************************************************
 * CODE.GS
 * Entry point của Web App + API cho Dashboard.
 * Các API nghiệp vụ khác (Thu/Chi, Công nợ, Báo cáo)
 * nằm ở ThuChi.gs, CongNo.gs, BaoCao.gs — sẽ bổ sung
 * ở các Giai đoạn tiếp theo.
 *************************************************/

function doGet() {
  return HtmlService.createTemplateFromFile('Index').evaluate()
    .setTitle('Quản lý Thu - Chi - Tồn Quỹ Tiền Mặt')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * Cho phép Index.html include CSS.html và JS.html.
 * Cách dùng trong Index.html: <?!= include('CSS'); ?>
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/*************************************************
 * API: DỮ LIỆU DASHBOARD
 * Tính động từ PHIEU_THU + PHIEU_CHI + CONG_NO,
 * chỉ tính các dòng trạng thái = "Hợp lệ".
 *************************************************/
function getDashboardData() {
  try {
    const todayKey = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');

    // Dashboard chỉ phản ánh Quỹ tiền mặt chính — Quỹ Công đoàn / Quỹ Cơm
    // là các quỹ riêng, có tồn/Sổ Quỹ tách biệt (xem _tonQuyHienTai bên dưới).
    const allThu = _sheetToObjects(SHEET_PHIEU_THU).filter(p => p.trang_thai === TRANG_THAI_HOP_LE && _chuanHoaLoaiQuy(p.loai_quy) === QUY_TIEN_MAT);
    const allChi = _sheetToObjects(SHEET_PHIEU_CHI).filter(p => p.trang_thai === TRANG_THAI_HOP_LE && _chuanHoaLoaiQuy(p.loai_quy) === QUY_TIEN_MAT);

    // Tồn đầu ngày = Số dư khởi tạo + tổng Thu - tổng Chi của TẤT CẢ các ngày TRƯỚC hôm nay
    const soDuKhoiTao = Number(_getCauHinh('SO_DU_QUY_KHOI_TAO')) || 0;
    let tonDauNgay = soDuKhoiTao;

    allThu.forEach(p => {
      if (_fmtDate(p.ngay_thu) < todayKey) tonDauNgay += Number(p.so_tien) || 0;
    });
    allChi.forEach(p => {
      if (_fmtDate(p.ngay_chi) < todayKey) tonDauNgay -= Number(p.so_tien) || 0;
    });

    // Thu/chi hôm nay
    let tongThuHomNay = 0, soPhieuThuHomNay = 0;
    allThu.forEach(p => {
      if (_fmtDate(p.ngay_thu) === todayKey) {
        tongThuHomNay += Number(p.so_tien) || 0;
        soPhieuThuHomNay++;
      }
    });

    let tongChiHomNay = 0, soPhieuChiHomNay = 0;
    allChi.forEach(p => {
      if (_fmtDate(p.ngay_chi) === todayKey) {
        tongChiHomNay += Number(p.so_tien) || 0;
        soPhieuChiHomNay++;
      }
    });

    const tonCuoiNgay = tonDauNgay + tongThuHomNay - tongChiHomNay;

    // Công nợ tổng hợp
    const congNoList = _sheetToObjects(SHEET_CONG_NO);
    let tongPhaiThu = 0, tongPhaiTra = 0, quaHanCount = 0;

    congNoList.forEach(c => {
      const conLai = (Number(c.so_tien_phat_sinh) || 0) - (Number(c.da_thanh_toan) || 0);
      if (conLai <= 0) return; // đã tất toán, bỏ qua

      if (c.loai_cong_no === 'Phải thu') tongPhaiThu += conLai;
      if (c.loai_cong_no === 'Phải trả') tongPhaiTra += conLai;

      if (c.han_thanh_toan) {
        const hanKey = _fmtDate(c.han_thanh_toan);
        if (hanKey && hanKey < todayKey) quaHanCount++;
      }
    });

    return _jsonOk({
      date: todayKey,
      ton_dau_ngay: tonDauNgay,
      tong_thu_hom_nay: tongThuHomNay,
      tong_chi_hom_nay: tongChiHomNay,
      ton_cuoi_ngay: tonCuoiNgay,
      so_phieu_thu_hom_nay: soPhieuThuHomNay,
      so_phieu_chi_hom_nay: soPhieuChiHomNay,
      tong_phai_thu: tongPhaiThu,
      tong_phai_tra: tongPhaiTra,
      cong_no_qua_han: quaHanCount,
      ton_quy_cong_doan: _tonQuyHienTai(QUY_CONG_DOAN)
    });

  } catch (err) {
    return _jsonErr(err);
  }
}
