/*************************************************
 * BAOCAONGAYKEO.GS
 * Báo cáo ngày cho nghiệp vụ mua Keo (gỗ nguyên liệu), tổng hợp từ
 * 3 nguồn:
 *   - Sheet "PhanTichNhapTT_DRAFT" (Google Sheet ngoài, cấu hình ở
 *     CAU_HINH key ID_SHEET_PHANTICH_NHAP_TT): tổng Nhập/Thanh toán
 *     trong ngày, chia theo Nguồn gốc (NG) và theo Đại lý chính
 *     (DL: DT/QT/KL) — đúng như 2 bảng "ĐẠI LÝ"/"NGUỒN GỐC" trong
 *     sheet Keo nhập / CK KEO của file Excel mẫu, KHÔNG tách chi
 *     tiết theo từng đại lý phụ (theo yêu cầu người dùng).
 *   - Sheet "PhieuCan_DN" (Google Sheet ngoài, cấu hình ở CAU_HINH
 *     key ID_SHEET_PHIEU_CAN_DN): danh sách chi tiết từng phiếu cân
 *     phát sinh trong ngày, để đối chiếu/tra cứu.
 *   - Sổ Quỹ tiền mặt (PHIEU_THU/PHIEU_CHI trong chính spreadsheet
 *     này): tồn quỹ đầu ngày / cuối ngày.
 *
 * Cả 2 Google Sheet ngoài PHẢI được chia sẻ (ít nhất quyền Xem) cho
 * tài khoản Google đang đứng tên chạy Apps Script này.
 *************************************************/

const TEN_SHEET_PHANTICH_NHAP_TT = 'PhanTichNhapTT_DRAFT';
const TEN_SHEET_PHIEU_CAN_DN = 'PhieuCan_DN';

/**
 * Cho phép cấu hình dán ID thô ("1kjYne-...") HOẶC nguyên đường link
 * Google Sheet ("https://docs.google.com/spreadsheets/d/1kjYne-.../edit?usp=sharing")
 * đều dùng được — SpreadsheetApp.openById() chỉ nhận đúng ID, dán
 * nguyên link vào sẽ báo lỗi "Illegal spreadsheet id or key".
 */
function _rutGonIdSheet(giaTri) {
  if (!giaTri) return '';
  const s = String(giaTri).trim();
  const m = s.match(/\/d\/([a-zA-Z0-9_-]+)/);
  return m ? m[1] : s;
}

/**
 * Mở 1 Google Sheet ngoài theo ID lưu trong CAU_HINH (key cauHinhKey)
 * và trả về đúng sheet con tên tenSheet. Ném lỗi tiếng Việt rõ ràng
 * nếu chưa cấu hình / không mở được / không tìm thấy sheet con.
 */
function _moSheetNgoaiTheoTen(cauHinhKey, tenSheet) {
  const id = _rutGonIdSheet(_getCauHinh(cauHinhKey));
  if (!id) {
    throw new Error('Chưa cấu hình ID Google Sheet nguồn dữ liệu (' + cauHinhKey +
      '). Vào Quản Trị > "Báo cáo ngày (Keo)" để thiết lập.');
  }
  let ssNgoai;
  try {
    ssNgoai = SpreadsheetApp.openById(id);
  } catch (e) {
    throw new Error('Không mở được Google Sheet nguồn dữ liệu (' + cauHinhKey +
      '). Kiểm tra lại ID hoặc quyền chia sẻ. Chi tiết: ' + e.message);
  }
  const sh = ssNgoai.getSheetByName(tenSheet);
  if (!sh) {
    throw new Error('Không tìm thấy sheet "' + tenSheet + '" trong Google Sheet nguồn (' + cauHinhKey + ').');
  }
  return sh;
}

/**
 * Kiểm tra 1 sheet ngoài có đầy đủ các cột bắt buộc hay không, ném lỗi
 * rõ ràng nếu thiếu — thay vì để lọt qua rồi âm thầm đọc ra undefined/0
 * cho từng dòng (rất khó phát hiện, dễ làm sai số liệu báo cáo mà
 * không ai nhận ra). Sheet nguồn là sheet "nháp" người dùng tự chỉnh
 * sửa nên cấu trúc cột có thể thay đổi theo thời gian.
 */
function _kiemTraCotBatBuoc(sh, cacCotCanCo) {
  const lastCol = sh.getLastColumn();
  if (lastCol === 0) throw new Error('Sheet "' + sh.getName() + '" trống, không có dữ liệu.');
  const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h).trim());
  const thieu = cacCotCanCo.filter(c => headers.indexOf(c) === -1);
  if (thieu.length > 0) {
    throw new Error('Sheet "' + sh.getName() + '" thiếu cột: ' + thieu.join(', ') +
      '. Kiểm tra lại cấu trúc sheet nguồn (có thể tên cột đã bị đổi).');
  }
}

/*************************************************
 * ĐỒNG BỘ DỮ LIỆU KEO VÀO SHEET NỘI BỘ
 * Mở 2 Google Sheet ngoài (PhanTichNhapTT_DRAFT, PhieuCan_DN) và mở
 * KHÔNG hề rẻ (mỗi lần mất 1-3 giây + phải đọc lại hàng nghìn dòng) —
 * đây là phần chậm nhất của Báo cáo ngày/tháng (Keo). Thay vì đọc lại
 * từ nguồn ngoài mỗi lần xem báo cáo, ta ĐỒNG BỘ (copy) toàn bộ 2
 * sheet đó vào 2 sheet nội bộ (cùng spreadsheet này) theo lịch (xem
 * thietLapDongBoHangNgay) hoặc bấm tay — mọi hàm báo cáo sau đó chỉ
 * đọc từ sheet nội bộ, nhanh như đọc Sổ Quỹ bình thường.
 *************************************************/
const SHEET_KEO_PHANTICH_DONGBO = 'KEO_PHANTICH_DONGBO';
const SHEET_KEO_PHIEUCAN_DONGBO = 'KEO_PHIEUCAN_DONGBO';

/**
 * Ghi đè 1 sheet nội bộ bằng dữ liệu mới đồng bộ — xoá sạch nội dung
 * cũ rồi ghi lại từ đầu (không tăng dần) để không bao giờ lẫn dữ liệu
 * đã xoá/sửa ở nguồn. cotNgayIndex (0-based, có thể bỏ trống) được ép
 * định dạng Văn bản (@) TRƯỚC khi ghi để giảm rủi ro Google Sheets tự
 * hiểu nhầm chuỗi "yyyy-MM-dd" thành 1 Date thật; dù vậy phía đọc vẫn
 * luôn dùng _ngayKeyLinhHoat() để an toàn cho cả 2 trường hợp.
 */
function _ghiDeSheetDongBo(ss, tenSheet, headers, rows, cotNgayIndex) {
  let sh = ss.getSheetByName(tenSheet);
  if (!sh) sh = ss.insertSheet(tenSheet);
  sh.clear();
  sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  if (rows.length > 0) {
    if (cotNgayIndex !== null && cotNgayIndex !== undefined) {
      sh.getRange(2, cotNgayIndex + 1, rows.length, 1).setNumberFormat('@');
    }
    sh.getRange(2, 1, rows.length, headers.length).setValues(rows);
  }
}

/**
 * API: chạy đồng bộ ngay — gọi từ Apps Script Editor (thiết lập lần
 * đầu / kiểm tra), từ trigger theo lịch (thietLapDongBoHangNgay), hoặc
 * từ nút "🔄 Đồng bộ dữ liệu Keo ngay" trong Quản Trị.
 */
function dongBoDuLieuKeo() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const shPhanTich = _moSheetNgoaiTheoTen('ID_SHEET_PHANTICH_NHAP_TT', TEN_SHEET_PHANTICH_NHAP_TT);
  _kiemTraCotBatBuoc(shPhanTich, ['Ngày', 'Loại', 'PhanLoai', 'Ten', 'KhoiLuongKg', 'GiaTri']);
  const headersA = ['Ngày', 'Loại', 'PhanLoai', 'Ten', 'KhoiLuongKg', 'GiaTri'];
  const rowsA = _sheetToObjectsFromSheetObj(shPhanTich).map(r => ([
    _ngayKeyLinhHoat(r['Ngày']), r['Loại'], r['PhanLoai'], _safeText(r['Ten']),
    Number(r['KhoiLuongKg']) || 0, Number(r['GiaTri']) || 0
  ]));
  _ghiDeSheetDongBo(ss, SHEET_KEO_PHANTICH_DONGBO, headersA, rowsA, 0);

  const headersB = ['Số phiếu', 'Ngày cân 1', 'Biển số 1', 'Khách hàng', 'ĐL', 'NG', 'KL hàng (KG)', 'Đơn giá_TC', 'Thành tiền', 'Trạng thái'];
  const shPhieuCan = _moSheetNgoaiTheoTen('ID_SHEET_PHIEU_CAN_DN', TEN_SHEET_PHIEU_CAN_DN);
  _kiemTraCotBatBuoc(shPhieuCan, headersB);
  const rowsB = _sheetToObjectsFromSheetObj(shPhieuCan).map(r => headersB.map(h => {
    if (h === 'Ngày cân 1') return _ngayKeyLinhHoat(r[h]);
    if (h === 'KL hàng (KG)' || h === 'Đơn giá_TC' || h === 'Thành tiền') return Number(r[h]) || 0;
    return _safeText(r[h]);
  }));
  _ghiDeSheetDongBo(ss, SHEET_KEO_PHIEUCAN_DONGBO, headersB, rowsB, 1);

  _setCauHinh('KEO_DONGBO_LUC', Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss'));

  return { so_dong_phantich: rowsA.length, so_dong_phieucan: rowsB.length };
}

/**
 * API: đồng bộ ngay, gọi từ giao diện Quản Trị (chỉ ADMIN).
 */
function dongBoDuLieuKeoTuGiaoDien(currentUser) {
  try {
    if (!currentUser || !currentUser.username) throw new Error('Thiếu thông tin người dùng.');
    _yeuCauQuyen(currentUser.username, [ROLE_ADMIN]);
    const ketQua = dongBoDuLieuKeo();
    return _jsonOk({
      so_dong_phantich: ketQua.so_dong_phantich,
      so_dong_phieucan: ketQua.so_dong_phieucan,
      luc: _getCauHinh('KEO_DONGBO_LUC')
    });
  } catch (err) {
    return _jsonErr(err);
  }
}

/**
 * CÁCH CHẠY: Trong Apps Script Editor, chọn hàm "thietLapDongBoHangNgay"
 * ở dropdown trên cùng > bấm Run > chạy 1 lần duy nhất để bật đồng bộ
 * tự động mỗi ngày (khoảng 1 giờ sáng). Muốn đồng bộ ngay lập tức bất
 * cứ lúc nào, dùng nút "🔄 Đồng bộ dữ liệu Keo ngay" trong Quản Trị.
 */
function thietLapDongBoHangNgay() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'dongBoDuLieuKeo') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('dongBoDuLieuKeo').timeBased().everyDays(1).atHour(1).create();

  SpreadsheetApp.getUi().alert(
    'Đã thiết lập đồng bộ dữ liệu Keo tự động — chạy 1 lần mỗi ngày vào khoảng 1 giờ sáng.\n\n' +
    'Nếu trong ngày có cập nhật dữ liệu Keo cần xem báo cáo ngay, vào Quản Trị > "Báo cáo ngày (Keo)" > bấm "🔄 Đồng bộ dữ liệu Keo ngay" để đồng bộ thủ công bất cứ lúc nào.'
  );
}

/**
 * Đọc dữ liệu PhanTichNhapTT_DRAFT đã đồng bộ sẵn trong sheet nội bộ
 * — nhanh, không phải mở Google Sheet ngoài mỗi lần gọi.
 */
function _docPhanTichNhapTTDaDongBo() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_KEO_PHANTICH_DONGBO);
  if (!sh) {
    throw new Error('Chưa có dữ liệu Keo được đồng bộ. Vào Quản Trị > "Báo cáo ngày (Keo)" > bấm "🔄 Đồng bộ dữ liệu Keo ngay".');
  }
  return _sheetToObjectsFromSheetObj(sh).map(r => ({
    'Ngày': _ngayKeyLinhHoat(r['Ngày']),
    'Loại': r['Loại'],
    'PhanLoai': r['PhanLoai'],
    'Ten': r['Ten'],
    'KhoiLuongKg': Number(r['KhoiLuongKg']) || 0,
    'GiaTri': Number(r['GiaTri']) || 0
  }));
}

/**
 * Đọc dữ liệu PhieuCan_DN đã đồng bộ sẵn trong sheet nội bộ.
 */
function _docPhieuCanDaDongBo() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_KEO_PHIEUCAN_DONGBO);
  if (!sh) {
    throw new Error('Chưa có dữ liệu Keo được đồng bộ. Vào Quản Trị > "Báo cáo ngày (Keo)" > bấm "🔄 Đồng bộ dữ liệu Keo ngay".');
  }
  return _sheetToObjectsFromSheetObj(sh);
}

/**
 * Tồn quỹ đầu ngày / cuối ngày của 1 loại quỹ, tính trên TOÀN BỘ
 * lịch sử giao dịch (không phụ thuộc bộ lọc), để luôn đúng kể cả
 * những ngày không phát sinh giao dịch nào.
 */
function _tonDauCuoiNgay(loaiQuy, ngay) {
  const res = getSoQuy({ loaiQuy: loaiQuy });
  const soDuKhoiTao = Number(_getCauHinh(_cauHinhSoDuKhoiTaoTheoQuy(loaiQuy))) || 0;
  if (!res.success) return { dau: soDuKhoiTao, cuoi: soDuKhoiTao };

  const list = res.data;
  let cuoi = null, dau = null;
  for (let i = list.length - 1; i >= 0; i--) {
    if (list[i].ngay <= ngay) { cuoi = list[i].ton; break; }
  }
  for (let i = list.length - 1; i >= 0; i--) {
    if (list[i].ngay < ngay) { dau = list[i].ton; break; }
  }
  if (dau === null) dau = soDuKhoiTao;
  if (cuoi === null) cuoi = dau;
  return { dau: dau, cuoi: cuoi };
}

/**
 * Quy đổi khối lượng từ KG (đơn vị lưu trong sheet nguồn) sang TẤN,
 * làm tròn 2 chữ số thập phân — mọi báo cáo Keo hiển thị khối lượng
 * theo TẤN, không phải KG.
 */
function _kgSangTan(kg) {
  return Math.round((Number(kg) || 0) / 10) / 100;
}

/*************************************************
 * API: LẤY DỮ LIỆU BÁO CÁO NGÀY (KEO)
 *************************************************/
function getBaoCaoNgayKeo(ngay) {
  try {
    if (!ngay) throw new Error('Vui lòng chọn ngày báo cáo.');

    const rowsPhanTich = _docPhanTichNhapTTDaDongBo().filter(r => r['Ngày'] === ngay);

    function gomNhom(loai, phanLoai) {
      return rowsPhanTich
        .filter(r => r['Loại'] === loai && r['PhanLoai'] === phanLoai)
        .map(r => ({ ten: _safeText(r['Ten']), kl: _kgSangTan(r['KhoiLuongKg']), gt: Number(r['GiaTri']) || 0 }))
        .sort((a, b) => b.gt - a.gt);
    }
    function tongCua(loai) {
      const t = rowsPhanTich.find(r => r['Loại'] === loai && r['PhanLoai'] === 'TONG');
      return { kl: t ? _kgSangTan(t['KhoiLuongKg']) : 0, gt: t ? Number(t['GiaTri']) || 0 : 0 };
    }

    const nhap = { theo_nguon_goc: gomNhom('NHAP', 'NG'), theo_dai_ly: gomNhom('NHAP', 'DL'), tong: tongCua('NHAP') };
    const thanhToan = { theo_nguon_goc: gomNhom('THANHTOAN', 'NG'), theo_dai_ly: gomNhom('THANHTOAN', 'DL'), tong: tongCua('THANHTOAN') };

    const chiTietPhieuCan = _docPhieuCanDaDongBo()
      .filter(r => _ngayKeyLinhHoat(r['Ngày cân 1']) === ngay)
      .map(r => ({
        so_phieu: _safeText(r['Số phiếu']),
        bien_so: _safeText(r['Biển số 1']),
        khach_hang: _safeText(r['Khách hàng']),
        dai_ly: _safeText(r['ĐL']),
        nguon_goc: _safeText(r['NG']),
        kl: _kgSangTan(r['KL hàng (KG)']),
        don_gia: Number(r['Đơn giá_TC']) || 0,
        thanh_tien: Number(r['Thành tiền']) || 0,
        trang_thai: _safeText(r['Trạng thái'])
      }))
      .sort((a, b) => Number(a.so_phieu) - Number(b.so_phieu));

    const soDu = _tonDauCuoiNgay(QUY_TIEN_MAT, ngay);

    return _jsonOk({
      ngay: ngay,
      nhap: nhap,
      thanh_toan: thanhToan,
      chi_tiet_phieu_can: chiTietPhieuCan,
      ton_quy_dau_ngay: soDu.dau,
      ton_quy_cuoi_ngay: soDu.cuoi
    });

  } catch (err) {
    return _jsonErr(err);
  }
}

/*************************************************
 * API: XUẤT BÁO CÁO NGÀY (KEO) RA EXCEL
 *************************************************/
function xuatBaoCaoNgayKeoExcel(ngay) {
  try {
    const res = getBaoCaoNgayKeo(ngay);
    if (!res.success) throw new Error(res.message);
    const d = res.data;

    function bangTong(nhom, tieuCot) {
      const rows = nhom.map(x => [x.ten, x.kl, x.gt]);
      return { headers: [tieuCot, 'Khối lượng (tấn)', 'Thành tiền'], rows: rows, condoTien: [2], soThapPhan: [1] };
    }

    const tongQuan = {
      tenSheet: 'Tổng quan',
      tieuDe: 'BÁO CÁO NGÀY - THU MUA KEO',
      phuDe: 'Ngày ' + ngay,
      headers: ['Chỉ tiêu', 'Khối lượng (tấn)', 'Giá trị (đ)'],
      rows: [
        ['Tổng NHẬP trong ngày', d.nhap.tong.kl, d.nhap.tong.gt],
        ['Tổng THANH TOÁN trong ngày', d.thanh_toan.tong.kl, d.thanh_toan.tong.gt],
        ['Tồn quỹ tiền mặt đầu ngày', '', d.ton_quy_dau_ngay],
        ['Tồn quỹ tiền mặt cuối ngày', '', d.ton_quy_cuoi_ngay]
      ],
      condoTien: [2],
      soThapPhan: [1]
    };

    const nhapNguonGoc = Object.assign({ tenSheet: 'Nhập - Nguồn gốc' }, bangTong(d.nhap.theo_nguon_goc, 'Nguồn gốc'));
    const nhapDaiLy = Object.assign({ tenSheet: 'Nhập - Đại lý' }, bangTong(d.nhap.theo_dai_ly, 'Đại lý'));
    const ttNguonGoc = Object.assign({ tenSheet: 'Thanh toán - Nguồn gốc' }, bangTong(d.thanh_toan.theo_nguon_goc, 'Nguồn gốc'));
    const ttDaiLy = Object.assign({ tenSheet: 'Thanh toán - Đại lý' }, bangTong(d.thanh_toan.theo_dai_ly, 'Đại lý'));

    const chiTiet = {
      tenSheet: 'Chi tiết phiếu cân',
      tieuDe: 'CHI TIẾT PHIẾU CÂN NGÀY ' + ngay,
      headers: ['Số phiếu', 'Biển số', 'Khách hàng', 'Đại lý', 'Nguồn gốc', 'KL (tấn)', 'Đơn giá', 'Thành tiền', 'Trạng thái'],
      rows: d.chi_tiet_phieu_can.map(r => [r.so_phieu, r.bien_so, r.khach_hang, r.dai_ly, r.nguon_goc, r.kl, r.don_gia, r.thanh_tien, r.trang_thai]),
      condoTien: [6, 7],
      soThapPhan: [5]
    };

    const file = _taoFileExcel('BaoCaoNgayKeo_' + ngay, [tongQuan, nhapNguonGoc, nhapDaiLy, ttNguonGoc, ttDaiLy, chiTiet]);

    return _jsonOk(file);
  } catch (err) {
    return _jsonErr(err);
  }
}

/*************************************************
 * API: GỬI BÁO CÁO NGÀY (KEO) QUA EMAIL
 * Người nhận lấy từ CAU_HINH key EMAIL_BAO_CAO_NGAY (thiết lập ở
 * Quản Trị), nhiều email cách nhau bởi dấu phẩy hoặc chấm phẩy.
 *************************************************/
function guiBaoCaoNgayKeoEmail(ngay, currentUser) {
  try {
    if (!currentUser || !currentUser.username) throw new Error('Thiếu thông tin người dùng.');
    _yeuCauQuyen(currentUser.username, [ROLE_ADMIN, ROLE_THU_QUY]);

    const emailCauHinh = _getCauHinh('EMAIL_BAO_CAO_NGAY');
    const danhSachEmail = String(emailCauHinh || '').split(/[,;]/).map(e => e.trim()).filter(Boolean);
    if (danhSachEmail.length === 0) {
      throw new Error('Chưa thiết lập email nhận báo cáo. Vào Quản Trị > "Báo cáo ngày (Keo)" để thiết lập.');
    }

    const fileRes = xuatBaoCaoNgayKeoExcel(ngay);
    if (!fileRes.success) throw new Error(fileRes.message);

    const blob = Utilities.newBlob(
      Utilities.base64Decode(fileRes.data.base64),
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      fileRes.data.filename
    );

    const dataRes = getBaoCaoNgayKeo(ngay);
    const d = dataRes.success ? dataRes.data : null;
    let noiDung = 'Kính gửi Anh/Chị,\n\nHệ thống HAK_QUY gửi Báo cáo ngày (thu mua Keo) ngày ' + ngay + ', chi tiết xem file Excel đính kèm.';
    if (d) {
      noiDung += '\n\n- Tổng nhập: ' + Math.round(d.nhap.tong.gt).toLocaleString('vi-VN') + ' đ (' + d.nhap.tong.kl.toLocaleString('vi-VN') + ' kg)' +
        '\n- Tổng thanh toán: ' + Math.round(d.thanh_toan.tong.gt).toLocaleString('vi-VN') + ' đ (' + d.thanh_toan.tong.kl.toLocaleString('vi-VN') + ' kg)' +
        '\n- Tồn quỹ tiền mặt cuối ngày: ' + Math.round(d.ton_quy_cuoi_ngay).toLocaleString('vi-VN') + ' đ';
    }
    noiDung += '\n\n(Email được gửi tự động từ hệ thống HAK_QUY)';

    GmailApp.sendEmail(danhSachEmail.join(','), 'Báo cáo ngày Keo - ' + ngay, noiDung, {
      attachments: [blob],
      name: 'HAK_QUY - Báo cáo tự động'
    });

    _writeAuditLog(currentUser.full_name, 'Báo Cáo Ngày', 'Gửi Email', ngay, '', 'Đã gửi tới: ' + danhSachEmail.join(', '));

    return _jsonOk({ da_gui_toi: danhSachEmail });

  } catch (err) {
    return _jsonErr(err);
  }
}
