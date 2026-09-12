/*************************************************
 * DOITUONG.GS
 * Quản lý danh mục Đối Tượng (Khách hàng/NCC/Nhân viên/Khác).
 * Thêm/Sửa chỉ ADMIN — đúng nguyên tắc "Quản lý danh mục"
 * ở mục 16. Không xóa vật lý — chỉ đổi trạng thái Hoạt động/
 * Ngừng hoạt động khi không còn dùng nữa.
 *************************************************/

const LOAI_DOI_TUONG_HOP_LE = ['Khách hàng', 'Nhà cung cấp', 'Nhân viên', 'Khác'];

function getDoiTuongList(filters) {
  try {
    filters = filters || {};
    let list = _sheetToObjects(SHEET_DOITUONG).map(d => ({
      ma_doi_tuong: d.ma_doi_tuong,
      ten_doi_tuong: d.ten_doi_tuong,
      loai_doi_tuong: d.loai_doi_tuong,
      so_dien_thoai: d.so_dien_thoai,
      dia_chi: d.dia_chi,
      mst: d.mst,
      nguoi_lien_he: d.nguoi_lien_he,
      ghi_chu: d.ghi_chu,
      trang_thai: d.trang_thai,
      __row: d.__row
    }));

    if (filters.loaiDoiTuong) list = list.filter(d => d.loai_doi_tuong === filters.loaiDoiTuong);
    if (filters.trangThai) list = list.filter(d => d.trang_thai === filters.trangThai);
    if (filters.tuKhoa) {
      const kw = filters.tuKhoa.toLowerCase();
      list = list.filter(d => String(d.ten_doi_tuong).toLowerCase().includes(kw) || String(d.ma_doi_tuong).toLowerCase().includes(kw));
    }

    list.sort((a, b) => b.__row - a.__row);
    return _jsonOk(list);
  } catch (err) {
    return _jsonErr(err);
  }
}

function addDoiTuong(payload, currentUser) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);

    if (!currentUser || !currentUser.username) throw new Error('Thiếu thông tin người dùng.');
    _yeuCauQuyen(currentUser.username, [ROLE_ADMIN]);

    if (!payload) throw new Error('Thiếu dữ liệu đối tượng.');
    if (!payload.ten_doi_tuong) throw new Error('Vui lòng nhập Tên đối tượng.');
    if (LOAI_DOI_TUONG_HOP_LE.indexOf(payload.loai_doi_tuong) === -1) throw new Error('Loại đối tượng không hợp lệ.');

    const maDoiTuong = _generateNextNumber('SO_DOI_TUONG_TIEP_THEO', 'DT');

    _sheet(SHEET_DOITUONG).appendRow([
      maDoiTuong,
      payload.ten_doi_tuong,
      payload.loai_doi_tuong,
      payload.so_dien_thoai || '',
      payload.dia_chi || '',
      payload.mst || '',
      payload.nguoi_lien_he || '',
      payload.ghi_chu || '',
      'Hoạt động'
    ]);

    _writeAuditLog(currentUser.full_name, 'Đối Tượng', 'Thêm', maDoiTuong, '',
      'Tên: ' + payload.ten_doi_tuong + ' | Loại: ' + payload.loai_doi_tuong);

    return _jsonOk({ ma_doi_tuong: maDoiTuong });

  } catch (err) {
    return _jsonErr(err);
  } finally {
    lock.releaseLock();
  }
}

function updateDoiTuong(payload, currentUser) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);

    if (!currentUser || !currentUser.username) throw new Error('Thiếu thông tin người dùng.');
    _yeuCauQuyen(currentUser.username, [ROLE_ADMIN]);

    if (!payload || !payload.ma_doi_tuong) throw new Error('Thiếu mã đối tượng cần sửa.');
    if (!payload.ten_doi_tuong) throw new Error('Vui lòng nhập Tên đối tượng.');
    if (LOAI_DOI_TUONG_HOP_LE.indexOf(payload.loai_doi_tuong) === -1) throw new Error('Loại đối tượng không hợp lệ.');

    const sh = _sheet(SHEET_DOITUONG);
    const data = sh.getDataRange().getValues();
    const headers = data[0].map(h => String(h).trim());
    const idxMa = headers.indexOf('ma_doi_tuong');

    let rowIndex = -1;
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][idxMa]) === String(payload.ma_doi_tuong)) { rowIndex = i; break; }
    }
    if (rowIndex === -1) throw new Error('Không tìm thấy đối tượng: ' + payload.ma_doi_tuong);

    const tenCu = data[rowIndex][headers.indexOf('ten_doi_tuong')];

    const rowValues = [
      payload.ma_doi_tuong,
      payload.ten_doi_tuong,
      payload.loai_doi_tuong,
      payload.so_dien_thoai || '',
      payload.dia_chi || '',
      payload.mst || '',
      payload.nguoi_lien_he || '',
      payload.ghi_chu || '',
      payload.trang_thai || 'Hoạt động'
    ];
    sh.getRange(rowIndex + 1, 1, 1, rowValues.length).setValues([rowValues]);

    _writeAuditLog(currentUser.full_name, 'Đối Tượng', 'Sửa', payload.ma_doi_tuong,
      'Tên: ' + tenCu, 'Tên: ' + payload.ten_doi_tuong + ' | Trạng thái: ' + (payload.trang_thai || 'Hoạt động'));

    return _jsonOk({ ma_doi_tuong: payload.ma_doi_tuong });

  } catch (err) {
    return _jsonErr(err);
  } finally {
    lock.releaseLock();
  }
}
