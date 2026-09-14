/*************************************************
 * XUATTONGHOPEXCEL.GS
 * Xuất 1 file Excel TỔNG HỢP theo đúng bố cục các sheet trong file
 * Excel mẫu gốc (QUỸ TIỀN MẶT...xlsx) mà người dùng cung cấp, đặt
 * tên sheet theo yêu cầu riêng của người dùng:
 *   - "Quy_TamUng"  : Sổ kế toán chi tiết Quỹ tiền mặt, theo THÁNG
 *   - "KEO_NHAP"    : bảng lịch theo ngày (1 dòng/ngày) — bên trái
 *                    theo Đại lý (DT/QT/KL), bên phải theo Nguồn gốc
 *   - "CK_KEO"      : như "KEO_NHAP" nhưng cho phần Thanh toán (Chi)
 *   - "COM"         : Sổ Cơm theo THÁNG
 *   - "QUY_CongDoan": Sổ kế toán chi tiết Quỹ Công đoàn, theo CẢ NĂM
 *                    (đúng như file mẫu gốc lấy cả năm cho sheet này)
 *
 * KHÔNG có sheet "NHân viên" (theo yêu cầu) và "Kangatang" (sheet
 * rỗng trong file gốc, không có dữ liệu để tái tạo).
 *
 * "Keo nhập"/"CK KEO" lấy dữ liệu từ sheet PhanTichNhapTT_DRAFT
 * (Google Sheet ngoài, xem BaoCaoNgayKeo.gs) — vẫn theo đúng mức
 * đơn giản hoá đã thống nhất trước đó (đại lý chỉ DT/QT/KL, không
 * tách theo từng đại lý phụ), chỉ khác là trình bày theo LỊCH (1
 * dòng/ngày trong tháng) thay vì 1 ngày đơn lẻ như "Báo cáo ngày".
 *************************************************/

function _ddmmyyyy(ngayKey) {
  const p = String(ngayKey).split('-');
  return p[2] + '/' + p[1] + '/' + p[0];
}

/**
 * Chuyển 1 khóa ngày "yyyy-MM-dd" (chuỗi text getSoQuy() trả về)
 * thành 1 Date thật, để ghi vào Excel dưới dạng Ngày Tháng có thể
 * định dạng/sắp xếp — thay vì để lọt ra 1 chuỗi text thô "2026-09-01"
 * trông không giống một cột Ngày chuẩn.
 */
function _ngayTuKey(ngayKey) {
  if (!ngayKey) return '';
  const p = String(ngayKey).split('-');
  if (p.length !== 3) return ngayKey;
  return new Date(Number(p[0]), Number(p[1]) - 1, Number(p[2]));
}

/*************************************************
 * API: XUẤT FILE EXCEL TỔNG HỢP
 *************************************************/
function xuatTongHopExcel(nam, thang) {
  try {
    nam = Number(nam);
    thang = Number(thang);
    if (!nam || !thang || thang < 1 || thang > 12) throw new Error('Vui lòng chọn Tháng/Năm hợp lệ.');

    const tenFile = 'BaoCaoTongHop_' + nam + '_' + String(thang).padStart(2, '0');
    const ss = SpreadsheetApp.create(tenFile);

    let rowsPhanTich = [];
    let loiNguonKeo = '';
    try {
      rowsPhanTich = _docPhanTichNhapTTDaDongBo();
    } catch (e) {
      loiNguonKeo = e.message;
    }

    _veSheetQuyTong(ss, 0, nam, thang);
    _veSheetKeoCalendar(ss, 1, 'KEO_NHAP', 'NHAP', nam, thang, rowsPhanTich, loiNguonKeo);
    _veSheetKeoCalendar(ss, 2, 'CK_KEO', 'THANHTOAN', nam, thang, rowsPhanTich, loiNguonKeo);
    _veSheetCom(ss, 3, nam, thang);
    _veSheetCongDoanNam(ss, 4, nam);

    const file = _xuatVaXoaFile(ss, tenFile);
    return _jsonOk(file);

  } catch (err) {
    return _jsonErr(err);
  }
}

/*************************************************
 * API: TÓM TẮT BÁO CÁO THÁNG (dùng cho màn hình xem nhanh trước khi
 * xuất Excel / gửi email — không dựng sheet, chỉ tổng hợp số liệu)
 *************************************************/
function getBaoCaoThangTongHop(nam, thang) {
  try {
    nam = Number(nam);
    thang = Number(thang);
    if (!nam || !thang || thang < 1 || thang > 12) throw new Error('Vui lòng chọn Tháng/Năm hợp lệ.');

    const cuoiThang = new Date(nam, thang, 0).getDate();
    const tuNgay = nam + '-' + String(thang).padStart(2, '0') + '-01';
    const denNgay = nam + '-' + String(thang).padStart(2, '0') + '-' + String(cuoiThang).padStart(2, '0');

    const tonTienMat = _tonDauCuoiNgay(QUY_TIEN_MAT, denNgay).cuoi;
    const tonCongDoan = _tonDauCuoiNgay(QUY_CONG_DOAN, denNgay).cuoi;

    const comRes = getSoComThang(nam, thang);
    const com = comRes.success ? comRes.data : { tong_suat: 0, tong_thanh_tien: 0, tong_tam_ung: 0 };

    let keoNhapKg = 0, keoNhapGt = 0, keoTTKg = 0, keoTTGt = 0, loiKeo = '';
    try {
      const monthPrefix = nam + '-' + String(thang).padStart(2, '0');
      _docPhanTichNhapTTDaDongBo().forEach(r => {
        if (r['PhanLoai'] !== 'TONG') return;
        if (r['Ngày'].slice(0, 7) !== monthPrefix) return;
        if (r['Loại'] === 'NHAP') { keoNhapKg += Number(r['KhoiLuongKg']) || 0; keoNhapGt += Number(r['GiaTri']) || 0; }
        else if (r['Loại'] === 'THANHTOAN') { keoTTKg += Number(r['KhoiLuongKg']) || 0; keoTTGt += Number(r['GiaTri']) || 0; }
      });
    } catch (e) {
      loiKeo = e.message;
    }

    return _jsonOk({
      nam: nam, thang: thang, tu_ngay: tuNgay, den_ngay: denNgay,
      ton_quy_tien_mat_cuoi_thang: tonTienMat,
      ton_quy_cong_doan_cuoi_thang: tonCongDoan,
      keo_nhap: { kl: _kgSangTan(keoNhapKg), gt: keoNhapGt },
      keo_thanh_toan: { kl: _kgSangTan(keoTTKg), gt: keoTTGt },
      loi_keo: loiKeo,
      com_tong_suat: com.tong_suat,
      com_thanh_tien: com.tong_thanh_tien,
      com_tam_ung: com.tong_tam_ung
    });

  } catch (err) {
    return _jsonErr(err);
  }
}

/*************************************************
 * API: GỬI BÁO CÁO THÁNG QUA EMAIL
 * Dùng chung danh sách email nhận báo cáo với Báo cáo ngày (Keo),
 * cấu hình ở CAU_HINH key EMAIL_BAO_CAO_NGAY (Quản Trị).
 *************************************************/
function guiBaoCaoThangEmail(nam, thang, currentUser) {
  try {
    if (!currentUser || !currentUser.username) throw new Error('Thiếu thông tin người dùng.');
    _yeuCauQuyen(currentUser.username, [ROLE_ADMIN, ROLE_THU_QUY]);

    const emailCauHinh = _getCauHinh('EMAIL_BAO_CAO_NGAY');
    const danhSachEmail = String(emailCauHinh || '').split(/[,;]/).map(e => e.trim()).filter(Boolean);
    if (danhSachEmail.length === 0) {
      throw new Error('Chưa thiết lập email nhận báo cáo. Vào Quản Trị > "Báo cáo ngày (Keo)" để thiết lập.');
    }

    const fileRes = xuatTongHopExcel(nam, thang);
    if (!fileRes.success) throw new Error(fileRes.message);

    const blob = Utilities.newBlob(
      Utilities.base64Decode(fileRes.data.base64),
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      fileRes.data.filename
    );

    const dataRes = getBaoCaoThangTongHop(nam, thang);
    const d = dataRes.success ? dataRes.data : null;
    let noiDung = 'Kính gửi Anh/Chị,\n\nHệ thống HAK_QUY gửi Báo cáo tháng ' + thang + '/' + nam + ', chi tiết xem file Excel đính kèm.';
    if (d) {
      noiDung += '\n\n- Tổng nhập Keo: ' + d.keo_nhap.kl.toFixed(2) + ' tấn (' + Math.round(d.keo_nhap.gt).toLocaleString('vi-VN') + ' đ)' +
        '\n- Tổng thanh toán Keo: ' + d.keo_thanh_toan.kl.toFixed(2) + ' tấn (' + Math.round(d.keo_thanh_toan.gt).toLocaleString('vi-VN') + ' đ)' +
        '\n- Tồn quỹ tiền mặt cuối tháng: ' + Math.round(d.ton_quy_tien_mat_cuoi_thang).toLocaleString('vi-VN') + ' đ' +
        '\n- Tồn quỹ công đoàn cuối tháng: ' + Math.round(d.ton_quy_cong_doan_cuoi_thang).toLocaleString('vi-VN') + ' đ' +
        '\n- Sổ cơm: ' + d.com_tong_suat.toLocaleString('vi-VN') + ' suất, thành tiền ' + Math.round(d.com_thanh_tien).toLocaleString('vi-VN') + ' đ';
    }
    noiDung += '\n\n(Email được gửi tự động từ hệ thống HAK_QUY)';

    GmailApp.sendEmail(danhSachEmail.join(','), 'Báo cáo tháng ' + thang + '/' + nam, noiDung, {
      attachments: [blob],
      name: 'HAK_QUY - Báo cáo tự động'
    });

    _writeAuditLog(currentUser.full_name, 'Báo Cáo Tháng', 'Gửi Email', thang + '/' + nam, '', 'Đã gửi tới: ' + danhSachEmail.join(', '));

    return _jsonOk({ da_gui_toi: danhSachEmail });

  } catch (err) {
    return _jsonErr(err);
  }
}

/**
 * Chèn 1 dòng "Cộng phát sinh trong ngày" sau mỗi ngày (nhóm theo
 * cfg.dayKeys, cùng thứ tự với cfg.dataRows — vốn đã được sắp xếp
 * theo ngày từ getSoQuy) — cộng tổng Nợ/Có phát sinh trong ngày đó
 * và chốt lại Tồn quỹ cuối ngày, đúng thông lệ sổ quỹ kế toán.
 * Trả về { rows, dongDamDaiIdx } — dongDamDaiIdx là các vị trí
 * (0-based, tính trong mảng rows trả về) cần in đậm/tô nền.
 */
function _chenDongCongPhatSinhNgay(cfg) {
  const soCot = cfg.headers.length;
  const rows = [];
  const dongDamIdx = [];
  let i = 0;
  while (i < cfg.dataRows.length) {
    const ngay = cfg.dayKeys[i];
    let tongNo = 0, tongCo = 0, tonCuoiNgay = '';
    while (i < cfg.dataRows.length && cfg.dayKeys[i] === ngay) {
      const hang = cfg.dataRows[i];
      tongNo += Number(hang[cfg.idxNo - 1]) || 0;
      tongCo += Number(hang[cfg.idxNo]) || 0;
      tonCuoiNgay = hang[cfg.idxTon - 1];
      rows.push(hang);
      i++;
    }
    const congNgay = new Array(soCot).fill('');
    congNgay[cfg.idxDienGiai - 1] = 'Cộng phát sinh ngày ' + _ddmmyyyy(ngay);
    congNgay[cfg.idxNo - 1] = tongNo;
    congNgay[cfg.idxNo] = tongCo;
    congNgay[cfg.idxTon - 1] = tonCuoiNgay;
    dongDamIdx.push(rows.length);
    rows.push(congNgay);
  }
  return { rows: rows, dongDamIdx: dongDamIdx };
}

/**
 * Dựng khung "Sổ kế toán chi tiết" 2 dòng tiêu đề cột (dòng "Số phát
 * sinh" gộp ngang, tách "Nợ"/"Có" ở dòng dưới) — dùng chung cho sheet
 * "Quỹ tổng" và "Công đoàn", chỉ khác bộ cột. Sau mỗi ngày có 1 dòng
 * "Cộng phát sinh ngày..." chốt tổng Nợ/Có và Tồn quỹ cuối ngày.
 */
function _veSheetSoKeToan(ss, idx, cfg) {
  const soCot = cfg.headers.length;
  const sh = idx === 0 ? ss.getSheets()[0].setName(cfg.tenSheet) : ss.insertSheet(cfg.tenSheet);

  const ketQuaCong = _chenDongCongPhatSinhNgay(cfg);
  const dataRowsCoCong = ketQuaCong.rows;
  const tongSoDong = 4 + 1 + dataRowsCoCong.length; // 2 dòng tiêu đề + 2 dòng header + 1 dòng tồn đầu kỳ + dữ liệu

  sh.getRange(1, 1, 1, soCot).merge().setValue(cfg.tieuDe)
    .setFontWeight('bold').setFontSize(13).setHorizontalAlignment('center');
  sh.getRange(2, 1, 1, soCot).merge().setValue(cfg.phuDe)
    .setFontStyle('italic').setFontSize(10).setHorizontalAlignment('center').setFontColor('#555555');

  sh.getRange(3, 1, 1, soCot).setValues([cfg.headers]);
  sh.getRange(3, cfg.idxNo, 1, 2).merge().setValue('Số phát sinh');
  sh.getRange(4, cfg.idxNo).setValue('Nợ');
  sh.getRange(4, cfg.idxNo + 1).setValue('Có');
  for (let c = 1; c <= soCot; c++) {
    if (c === cfg.idxNo || c === cfg.idxNo + 1) continue;
    sh.getRange(3, c, 2, 1).merge();
  }
  sh.getRange(3, 1, 2, soCot)
    .setFontWeight('bold').setBackground('#1e40af').setFontColor('#ffffff')
    .setHorizontalAlignment('center').setVerticalAlignment('middle').setWrap(true);
  sh.setRowHeight(3, 22);
  sh.setRowHeight(4, 22);

  const dong = 5;
  const toanBoDuLieu = [cfg.tonDauKyRow].concat(dataRowsCoCong);
  sh.getRange(dong, 1, toanBoDuLieu.length, soCot).setValues(toanBoDuLieu);
  sh.getRange(dong, 1, 1, soCot).setFontWeight('bold').setBackground('#eef2ff');
  sh.getRange(dong, 1, toanBoDuLieu.length, soCot).setVerticalAlignment('middle');
  (cfg.condoTienCols || []).forEach(c => {
    sh.getRange(dong, c, toanBoDuLieu.length, 1).setNumberFormat('#,##0').setHorizontalAlignment('right');
  });
  cfg.ngayCols.forEach(c => sh.getRange(dong, c, toanBoDuLieu.length, 1).setHorizontalAlignment('center').setNumberFormat('dd/MM/yyyy'));

  // Dòng "Cộng phát sinh ngày..." in đậm, tô nền nhạt để dễ nhận biết
  ketQuaCong.dongDamIdx.forEach(i0 => {
    sh.getRange(dong + 1 + i0, 1, 1, soCot).setFontWeight('bold').setFontStyle('italic').setBackground('#f3f4f6');
  });

  sh.getRange(3, 1, tongSoDong - 2, soCot)
    .setBorder(true, true, true, true, true, true, '#999999', SpreadsheetApp.BorderStyle.SOLID);

  sh.setFrozenRows(4);
  _datDoRongCotTheoTieuDe(sh, cfg.headers);
  // Cột Nợ/Có (dòng con, không nằm trong cfg.headers ở dòng 3 nên
  // _datDoRongCotTheoTieuDe không nhận diện được) — đặt tay cho chắc
  sh.setColumnWidth(cfg.idxNo, 170);
  sh.setColumnWidth(cfg.idxNo + 1, 170);
  return sh;
}

function _veSheetQuyTong(ss, idx, nam, thang) {
  const cuoiThang = new Date(nam, thang, 0).getDate();
  const tuNgay = nam + '-' + String(thang).padStart(2, '0') + '-01';
  const denNgay = nam + '-' + String(thang).padStart(2, '0') + '-' + String(cuoiThang).padStart(2, '0');

  const res = getSoQuy({ loaiQuy: QUY_TIEN_MAT, tuNgay: tuNgay, denNgay: denNgay });
  if (!res.success) throw new Error(res.message);
  const list = res.data;
  const tonDauKy = _tonDauCuoiNgay(QUY_TIEN_MAT, tuNgay).dau;

  const headers = ['Ngày hạch toán', 'Ngày chứng từ', 'Số phiếu thu', 'Số phiếu chi', 'Diễn giải',
    'Tài khoản', 'TK đối ứng', 'Số phát sinh', '', 'Số tồn', 'Người nhận/Người nộp', 'Mã NV'];
  const tonDauKyRow = ['', '', '', '', 'Số tồn đầu kỳ', '1111', '', 0, 0, tonDauKy, '', ''];
  const dataRows = list.map(t => [_ngayTuKey(t.ngay_hach_toan), _ngayTuKey(t.ngay), t.so_phieu_thu, t.so_phieu_chi, t.noi_dung,
    t.tai_khoan, t.tk_doi_ung, t.thu || '', t.chi || '', t.ton, t.nguoi_nhan_nop, t.ma_nhan_vien]);

  _veSheetSoKeToan(ss, idx, {
    tenSheet: 'Quy_TamUng',
    tieuDe: 'SỔ KẾ TOÁN CHI TIẾT QUỸ TIỀN MẶT',
    phuDe: 'Loại tiền: Tổng hợp; Tài khoản: 1111; Từ ngày ' + _ddmmyyyy(tuNgay) + ' đến ngày ' + _ddmmyyyy(denNgay),
    headers: headers,
    idxNo: 8,
    idxDienGiai: 5,
    idxTon: 10,
    ngayCols: [1, 2],
    tonDauKyRow: tonDauKyRow,
    dataRows: dataRows,
    dayKeys: list.map(t => t.ngay),
    condoTienCols: [8, 9, 10]
  });
}

function _veSheetCongDoanNam(ss, idx, nam) {
  const tuNgay = nam + '-01-01';
  const denNgay = nam + '-12-31';

  const res = getSoQuy({ loaiQuy: QUY_CONG_DOAN, tuNgay: tuNgay, denNgay: denNgay });
  if (!res.success) throw new Error(res.message);
  const list = res.data;
  const tonDauKy = _tonDauCuoiNgay(QUY_CONG_DOAN, tuNgay).dau;

  const headers = ['Ngày hạch toán', 'Ngày chứng từ', 'Số phiếu thu', 'Số phiếu chi', 'Diễn giải',
    'Số phát sinh', '', 'Số tồn', 'Người nhận/Người nộp', 'Chi nhánh'];
  const tonDauKyRow = ['', '', '', '', 'Số tồn đầu kỳ', 0, 0, tonDauKy, '', ''];
  const dataRows = list.map(t => [_ngayTuKey(t.ngay_hach_toan), _ngayTuKey(t.ngay), t.so_phieu_thu, t.so_phieu_chi, t.noi_dung,
    t.thu || '', t.chi || '', t.ton, t.nguoi_nhan_nop, t.chi_nhanh]);

  _veSheetSoKeToan(ss, idx, {
    tenSheet: 'QUY_CongDoan',
    tieuDe: 'SỔ KẾ TOÁN CHI TIẾT QUỸ CÔNG ĐOÀN',
    phuDe: 'Loại tiền: VND; Từ ngày ' + _ddmmyyyy(tuNgay) + ' đến ngày ' + _ddmmyyyy(denNgay),
    headers: headers,
    idxNo: 6,
    idxDienGiai: 5,
    idxTon: 8,
    ngayCols: [1, 2],
    tonDauKyRow: tonDauKyRow,
    dataRows: dataRows,
    dayKeys: list.map(t => t.ngay),
    condoTienCols: [6, 7, 8]
  });
}

function _veSheetCom(ss, idx, nam, thang) {
  const res = getSoComThang(nam, thang);
  if (!res.success) throw new Error(res.message);
  const d = res.data;
  const cuoiThang = new Date(nam, thang, 0).getDate();

  const sh = idx === 0 ? ss.getSheets()[0].setName('COM') : ss.insertSheet('COM');
  const headers = ['NGÀY', 'Trưa', 'Tối', 'Tổng', 'Đơn giá', 'Thành tiền', 'Ngày tạm ứng', 'Số tiền tạm ứng'];
  const soCot = headers.length;

  sh.getRange(1, 1, 1, soCot).merge().setValue('THÁNG ' + String(thang).padStart(2, '0') + '/' + nam)
    .setFontWeight('bold').setFontSize(13).setHorizontalAlignment('center');
  sh.getRange(2, 1, 1, soCot).merge().setValue('(Từ 01-' + cuoiThang + '/' + String(thang).padStart(2, '0') + ')')
    .setFontStyle('italic').setFontSize(10).setHorizontalAlignment('center').setFontColor('#555555');
  sh.getRange(3, 1, 1, soCot).setValues([headers])
    .setFontWeight('bold').setBackground('#1e40af').setFontColor('#ffffff').setHorizontalAlignment('center');

  const byDay = {};
  d.rows.forEach(r => { byDay[Number(r.ngay.slice(8, 10))] = r; });

  const dataRows = [];
  for (let day = 1; day <= cuoiThang; day++) {
    const r = byDay[day];
    if (r) {
      dataRows.push([day, r.buoi_trua || '', r.buoi_toi || '', r.tong_suat || 0, r.don_gia || '', r.thanh_tien || '',
        r.ngay_tam_ung ? Number(r.ngay_tam_ung.slice(8, 10)) : '', r.so_tien_tam_ung || '']);
    } else {
      dataRows.push([day, '', '', 0, '', '', '', '']);
    }
  }

  sh.getRange(4, 1, dataRows.length, soCot).setValues(dataRows).setVerticalAlignment('middle');
  sh.getRange(4, 1, dataRows.length, 4).setHorizontalAlignment('center');
  sh.getRange(4, 5, dataRows.length, 2).setNumberFormat('#,##0').setHorizontalAlignment('right');
  sh.getRange(4, 7, dataRows.length, 1).setHorizontalAlignment('center');
  sh.getRange(4, 8, dataRows.length, 1).setNumberFormat('#,##0').setHorizontalAlignment('right');
  sh.getRange(3, 1, dataRows.length + 1, soCot).setBorder(true, true, true, true, true, true, '#999999', SpreadsheetApp.BorderStyle.SOLID);

  sh.setFrozenRows(3);
  _datDoRongCotTheoTieuDe(sh, headers);
}

/**
 * Bảng lịch theo ngày cho "Keo nhập" (loai = 'NHAP') / "CK KEO"
 * (loai = 'THANHTOAN'): bên trái theo Đại lý, bên phải theo Nguồn
 * gốc, mỗi dòng là 1 ngày trong tháng — đúng bố cục file mẫu, chỉ
 * khác là danh sách đại lý/nguồn gốc lấy động từ dữ liệu thật
 * (không cố định cứng theo file mẫu cũ) và KHÔNG tách chi tiết theo
 * từng đại lý phụ (đã thống nhất với người dùng trước đó).
 */
function _veSheetKeoCalendar(ss, idx, tenSheet, loai, nam, thang, rowsPhanTich, loiNguonKeo) {
  const sh = idx === 0 ? ss.getSheets()[0].setName(tenSheet) : ss.insertSheet(tenSheet);

  if (loiNguonKeo) {
    sh.getRange(1, 1).setValue('Không lấy được dữ liệu từ sheet nguồn PhanTichNhapTT_DRAFT: ' + loiNguonKeo)
      .setFontColor('#b91c1c').setFontWeight('bold');
    return;
  }

  const cuoiThang = new Date(nam, thang, 0).getDate();
  const monthPrefix = nam + '-' + String(thang).padStart(2, '0');

  const rows = rowsPhanTich.filter(r => r['Loại'] === loai && _ngayKeyLinhHoat(r['Ngày']).slice(0, 7) === monthPrefix);

  const dealerSet = {}, nguonGocSet = {};
  rows.forEach(r => {
    if (r['PhanLoai'] === 'DL') dealerSet[_safeText(r['Ten'])] = true;
    else if (r['PhanLoai'] === 'NG') nguonGocSet[_safeText(r['Ten'])] = true;
  });
  const dealerList = Object.keys(dealerSet).sort();
  const nguonGocList = Object.keys(nguonGocSet).sort();

  const nhanLabel = loai === 'NHAP' ? 'nhập' : 'thanh toán';
  const leftHeaders = ['Ngày', 'Tổng ' + nhanLabel + ' (tấn)', 'Thành tiền'].concat(dealerList);
  const rightHeaders = ['Ngày', 'Tổng ' + nhanLabel + ' (tấn)', 'Thành tiền'].concat(nguonGocList);
  const soCotTrai = leftHeaders.length;
  const soCotPhai = rightHeaders.length;
  const soCot = soCotTrai + 1 + soCotPhai;

  const byDay = {};
  rows.forEach(r => {
    const d = _ngayKeyLinhHoat(r['Ngày']);
    if (!byDay[d]) byDay[d] = {};
    byDay[d][r['PhanLoai'] + '|' + _safeText(r['Ten'])] = { kl: _kgSangTan(r['KhoiLuongKg']), gt: Number(r['GiaTri']) || 0 };
  });

  sh.getRange(1, 1, 1, soCot).merge().setValue(
    (loai === 'NHAP' ? 'BÁO CÁO NHẬP KEO' : 'BÁO CÁO THANH TOÁN (CHI) KEO') +
    ' THÁNG ' + String(thang).padStart(2, '0') + '/' + nam
  ).setFontWeight('bold').setFontSize(13).setHorizontalAlignment('center');

  const headerRow = leftHeaders.concat(['']).concat(rightHeaders);
  sh.getRange(2, 1, 1, soCot).setValues([headerRow])
    .setFontWeight('bold').setBackground('#1e40af').setFontColor('#ffffff').setHorizontalAlignment('center');

  const dataRows = [];
  let tongKlAll = 0, tongGtAll = 0;
  for (let day = 1; day <= cuoiThang; day++) {
    const dKey = monthPrefix + '-' + String(day).padStart(2, '0');
    const tong = (byDay[dKey] && byDay[dKey]['TONG|Tổng cộng']) || { kl: 0, gt: 0 };
    tongKlAll += tong.kl;
    tongGtAll += tong.gt;

    const leftRow = [day, tong.kl || '', tong.gt || ''].concat(dealerList.map(name => {
      const c = byDay[dKey] && byDay[dKey]['DL|' + name];
      return c ? c.kl : '';
    }));
    const rightRow = [day, tong.kl || '', tong.gt || ''].concat(nguonGocList.map(name => {
      const c = byDay[dKey] && byDay[dKey]['NG|' + name];
      return c ? c.kl : '';
    }));
    dataRows.push(leftRow.concat(['']).concat(rightRow));
  }

  const totalRow = new Array(soCot).fill('');
  totalRow[0] = 'Tổng cộng';
  totalRow[1] = tongKlAll;
  totalRow[2] = tongGtAll;
  totalRow[soCotTrai + 2] = tongKlAll;
  totalRow[soCotTrai + 3] = tongGtAll;
  dataRows.push(totalRow);

  sh.getRange(3, 1, dataRows.length, soCot).setValues(dataRows).setVerticalAlignment('middle');

  // Cột "Ngày" (trái: cột 1, phải: cột soCotTrai+2) căn giữa
  sh.getRange(3, 1, dataRows.length, 1).setHorizontalAlignment('center');
  sh.getRange(3, soCotTrai + 2, dataRows.length, 1).setHorizontalAlignment('center');

  // Cột khối lượng (kg) — có thể lẻ, giữ 2 số thập phân: Tổng KL +
  // từng cột Đại lý/Nguồn gốc, cả bên trái lẫn bên phải
  const cotKlTrai = [2].concat(dealerList.map((_, i) => 4 + i));
  const cotKlPhai = [soCotTrai + 3].concat(nguonGocList.map((_, i) => soCotTrai + 5 + i));
  cotKlTrai.concat(cotKlPhai).forEach(c => {
    sh.getRange(3, c, dataRows.length, 1).setNumberFormat('#,##0.00').setHorizontalAlignment('right');
  });

  // Cột "Thành tiền" (trái: cột 3, phải: cột soCotTrai+4) — tiền, không lẻ
  sh.getRange(3, 3, dataRows.length, 1).setNumberFormat('#,##0').setHorizontalAlignment('right');
  sh.getRange(3, soCotTrai + 4, dataRows.length, 1).setNumberFormat('#,##0').setHorizontalAlignment('right');

  sh.getRange(dataRows.length + 2, 1, 1, soCot).setFontWeight('bold').setBackground('#eef2ff');
  sh.getRange(2, 1, dataRows.length + 1, soCot).setBorder(true, true, true, true, true, true, '#999999', SpreadsheetApp.BorderStyle.SOLID);

  sh.setFrozenRows(2);
  _datDoRongCotTheoTieuDe(sh, headerRow);
}
