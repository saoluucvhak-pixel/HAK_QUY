/*************************************************
 * XUATEXCEL.GS
 * Xuất báo cáo / sổ sách ra file Excel (.xlsx) THẬT, có định dạng
 * đẹp (tiêu đề, header tô màu, số tiền có dấu phẩy ngăn cách, kẻ
 * khung, cố định dòng tiêu đề) — thay cho việc chỉ xuất CSV thô.
 *
 * Cách làm: tạo 1 Google Sheet tạm, đổ dữ liệu + định dạng vào đó,
 * dùng URL export sẵn có của Google Sheets để lấy đúng file .xlsx,
 * rồi xoá Sheet tạm đi ngay. Trả về nội dung file dạng base64 để
 * trình duyệt tải xuống trực tiếp — không để lại gì trên Drive.
 *************************************************/

/**
 * @param {string} tenFile tên file xuất ra (không cần đuôi .xlsx)
 * @param {Array<{tenSheet:string, tieuDe?:string, phuDe?:string,
 *   headers:string[], rows:Array<Array>, condoTien?:number[]}>} dsSheet
 * @returns {{base64:string, filename:string}}
 */
function _taoFileExcel(tenFile, dsSheet) {
  const ss = SpreadsheetApp.create(tenFile);

  dsSheet.forEach(function (sd, idx) {
    const sh = idx === 0 ? ss.getSheets()[0].setName(sd.tenSheet) : ss.insertSheet(sd.tenSheet);
    const soCot = sd.headers.length;
    let dong = 1;

    if (sd.tieuDe) {
      sh.getRange(dong, 1, 1, soCot).merge().setValue(sd.tieuDe)
        .setFontWeight('bold').setFontSize(13).setHorizontalAlignment('center');
      dong++;
    }
    if (sd.phuDe) {
      sh.getRange(dong, 1, 1, soCot).merge().setValue(sd.phuDe)
        .setFontStyle('italic').setFontSize(10).setHorizontalAlignment('center').setFontColor('#555555');
      dong++;
    }
    dong++; // dòng trống ngăn cách tiêu đề với bảng dữ liệu

    const dongHeader = dong;
    sh.getRange(dongHeader, 1, 1, soCot).setValues([sd.headers])
      .setFontWeight('bold').setBackground('#1e40af').setFontColor('#ffffff')
      .setHorizontalAlignment('center').setVerticalAlignment('middle');
    sh.setRowHeight(dongHeader, 26);
    dong++;

    if (sd.rows.length > 0) {
      sh.getRange(dong, 1, sd.rows.length, soCot).setValues(sd.rows);
      (sd.condoTien || []).forEach(function (colIdx) {
        sh.getRange(dong, colIdx + 1, sd.rows.length, 1).setNumberFormat('#,##0');
      });
      sh.getRange(dong, 1, sd.rows.length, soCot)
        .setBorder(true, true, true, true, true, true, '#d1d5db', SpreadsheetApp.BorderStyle.SOLID);
    }

    sh.getRange(dongHeader, 1, 1, soCot)
      .setBorder(true, true, true, true, true, true, '#d1d5db', SpreadsheetApp.BorderStyle.SOLID);

    sh.setFrozenRows(dongHeader);
    sh.autoResizeColumns(1, soCot);
  });

  return _xuatVaXoaFile(ss, tenFile);
}

/**
 * Xuất 1 Spreadsheet tạm (đã đổ đủ dữ liệu/định dạng) ra file .xlsx
 * dạng base64, rồi xoá Spreadsheet tạm đó khỏi Drive ngay — dùng
 * chung cho _taoFileExcel() và mọi hàm tự dựng sheet thủ công khác
 * (ví dụ báo cáo tổng hợp nhiều sheet với layout phức tạp, không
 * theo khuôn 1-tiêu-đề/1-bảng của _taoFileExcel).
 */
function _xuatVaXoaFile(ss, tenFile) {
  SpreadsheetApp.flush();
  const fileId = ss.getId();

  const token = ScriptApp.getOAuthToken();
  const url = 'https://docs.google.com/spreadsheets/d/' + fileId + '/export?format=xlsx';
  const response = UrlFetchApp.fetch(url, { headers: { Authorization: 'Bearer ' + token } });
  const blob = response.getBlob();
  const base64 = Utilities.base64Encode(blob.getBytes());

  DriveApp.getFileById(fileId).setTrashed(true);

  return { base64: base64, filename: tenFile + '.xlsx' };
}

/*************************************************
 * API: XUẤT BÁO CÁO THU - CHI RA EXCEL
 *************************************************/
function xuatBaoCaoThuChiExcel(filters) {
  try {
    const res = getBaoCaoThuChi(filters);
    if (!res.success) throw new Error(res.message);
    const d = res.data;

    const headers = ['Ngày', 'Số phiếu', 'Loại', 'Nội dung', 'Đối tượng', 'Thu', 'Chi', 'Người lập'];
    const rows = d.chi_tiet.map(function (t) {
      return [t.ngay, t.so_phieu, t.loai, t.noi_dung || '', t.doi_tuong || '', t.thu || '', t.chi || '', t.nguoi_lap || ''];
    });
    rows.push(['', '', '', '', 'TỔNG CỘNG', d.tong_thu, d.tong_chi, '']);

    const phuDe = 'Từ ngày ' + d.tu_ngay + ' đến ' + d.den_ngay +
      '  |  Tồn đầu kỳ: ' + Math.round(d.ton_dau_ky).toLocaleString('vi-VN') + ' đ' +
      '  |  Tồn cuối kỳ: ' + Math.round(d.ton_cuoi_ky).toLocaleString('vi-VN') + ' đ';

    const file = _taoFileExcel('BaoCaoThuChi_' + d.tu_ngay + '_' + d.den_ngay, [{
      tenSheet: 'Báo cáo Thu Chi',
      tieuDe: 'BÁO CÁO THU - CHI',
      phuDe: phuDe,
      headers: headers,
      rows: rows,
      condoTien: [5, 6]
    }]);

    return _jsonOk(file);
  } catch (err) {
    return _jsonErr(err);
  }
}

/*************************************************
 * API: XUẤT BÁO CÁO CÔNG NỢ TỔNG HỢP RA EXCEL
 *************************************************/
function xuatBaoCaoCongNoExcel(loaiCongNo, tuNgay, denNgay) {
  try {
    const res = getBaoCaoCongNoTongHop(loaiCongNo, tuNgay, denNgay);
    if (!res.success) throw new Error(res.message);
    const list = res.data;

    const headers = ['Đối tượng', 'Đầu kỳ', 'Phát sinh', 'Đã thu/trả', 'Cuối kỳ'];
    const rows = list.map(function (r) {
      return [r.doi_tuong, r.dau_ky, r.phat_sinh, r.da_thu_tra, r.cuoi_ky];
    });
    const tongCuoiKy = list.reduce(function (s, r) { return s + r.cuoi_ky; }, 0);
    rows.push(['TỔNG CỘNG', '', '', '', tongCuoiKy]);

    const file = _taoFileExcel('BaoCaoCongNo_' + loaiCongNo.replace(/\s+/g, '') + '_' + tuNgay + '_' + denNgay, [{
      tenSheet: 'Công nợ ' + loaiCongNo,
      tieuDe: 'BÁO CÁO CÔNG NỢ ' + loaiCongNo.toUpperCase(),
      phuDe: 'Từ ngày ' + tuNgay + ' đến ' + denNgay,
      headers: headers,
      rows: rows,
      condoTien: [1, 2, 3, 4]
    }]);

    return _jsonOk(file);
  } catch (err) {
    return _jsonErr(err);
  }
}

/*************************************************
 * API: XUẤT BÁO CÁO CÔNG NỢ QUÁ HẠN RA EXCEL
 *************************************************/
function xuatBaoCaoQuaHanExcel() {
  try {
    const res = getBaoCaoCongNoQuaHan();
    if (!res.success) throw new Error(res.message);
    const list = res.data;

    const headers = ['Mã CN', 'Đối tượng', 'Loại', 'Nội dung', 'Ngày phát sinh', 'Hạn TT', 'Số ngày quá hạn', 'Còn lại'];
    const rows = list.map(function (r) {
      return [r.ma_cong_no, r.doi_tuong, r.loai_cong_no, r.noi_dung || '', r.ngay_phat_sinh, r.han_thanh_toan, r.so_ngay_qua_han, r.con_lai];
    });

    const file = _taoFileExcel('BaoCaoCongNoQuaHan_' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd'), [{
      tenSheet: 'Công nợ quá hạn',
      tieuDe: 'BÁO CÁO CÔNG NỢ QUÁ HẠN',
      phuDe: 'Tính đến ngày ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy'),
      headers: headers,
      rows: rows,
      condoTien: [7]
    }]);

    return _jsonOk(file);
  } catch (err) {
    return _jsonErr(err);
  }
}

/*************************************************
 * API: XUẤT SỔ QUỸ (Quỹ tiền mặt / Quỹ Công đoàn) RA EXCEL
 * Trình bày theo đúng mẫu sổ kế toán chi tiết quỹ tiền mặt.
 *************************************************/
function xuatSoQuyExcel(filters) {
  try {
    const res = getSoQuy(filters);
    if (!res.success) throw new Error(res.message);
    const list = res.data;
    const loaiQuy = _chuanHoaLoaiQuy(filters && filters.loaiQuy);

    const headers = ['Ngày HT', 'Ngày CT', 'Số PT', 'Số PC', 'Diễn giải', 'TK', 'TK đối ứng', 'Thu', 'Chi', 'Tồn', 'Người nhận/nộp'];
    const rows = list.map(function (t) {
      return [t.ngay_hach_toan, t.ngay, t.so_phieu_thu, t.so_phieu_chi, t.noi_dung || '',
        t.tai_khoan, t.tk_doi_ung, t.thu || '', t.chi || '', t.ton, t.nguoi_nhan_nop || ''];
    });

    let phuDe = 'Tài khoản: 1111';
    if (filters && filters.tuNgay) phuDe += '  |  Từ ngày ' + filters.tuNgay;
    if (filters && filters.denNgay) phuDe += ' đến ' + filters.denNgay;

    const file = _taoFileExcel('SoQuy_' + loaiQuy.replace(/\s+/g, ''), [{
      tenSheet: 'Sổ kế toán chi tiết',
      tieuDe: 'SỔ KẾ TOÁN CHI TIẾT ' + loaiQuy.toUpperCase(),
      phuDe: phuDe,
      headers: headers,
      rows: rows,
      condoTien: [7, 8, 9]
    }]);

    return _jsonOk(file);
  } catch (err) {
    return _jsonErr(err);
  }
}
