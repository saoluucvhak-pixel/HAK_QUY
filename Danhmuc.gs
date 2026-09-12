/*************************************************
 * DANHMUC.GS
 * Quản lý danh mục Loại Thu / Loại Chi. Thêm/Sửa chỉ ADMIN.
 * Không xóa vật lý — chỉ đổi trạng thái Hoạt động/Ngừng
 * hoạt động (mục nào Ngừng hoạt động sẽ không hiện trong
 * dropdown khi lập Phiếu Thu/Chi mới, nhưng dữ liệu cũ
 * vẫn giữ nguyên).
 *************************************************/

function _sheetDanhMuc(loai) {
  return loai === 'Thu' ? SHEET_LOAI_THU : SHEET_LOAI_CHI;
}
function _cauHinhSoDanhMuc(loai) {
  return loai === 'Thu' ? 'SO_LOAI_THU_TIEP_THEO' : 'SO_LOAI_CHI_TIEP_THEO';
}
function _prefixDanhMuc(loai) {
  return loai === 'Thu' ? 'LT' : 'LC';
}
function _colMaDanhMuc(loai) {
  return loai === 'Thu' ? 'ma_loai_thu' : 'ma_loai_chi';
}
function _colTenDanhMuc(loai) {
  return loai === 'Thu' ? 'ten_loai_thu' : 'ten_loai_chi';
}

function getDanhMucList(loai) {
  try {
    if (loai !== 'Thu' && loai !== 'Chi') throw new Error('Loại danh mục không hợp lệ.');
    const maCol = _colMaDanhMuc(loai);
    const tenCol = _colTenDanhMuc(loai);

    const list = _sheetToObjects(_sheetDanhMuc(loai)).map(d => ({
      ma: d[maCol],
      ten: d[tenCol],
      trang_thai: d.trang_thai,
      __row: d.__row
    }));
    list.sort((a, b) => b.__row - a.__row);

    return _jsonOk(list);
  } catch (err) {
    return _jsonErr(err);
  }
}

function addDanhMuc(loai, tenDanhMuc, currentUser) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);

    if (!currentUser || !currentUser.username) throw new Error('Thiếu thông tin người dùng.');
    _yeuCauQuyen(currentUser.username, [ROLE_ADMIN]);

    if (loai !== 'Thu' && loai !== 'Chi') throw new Error('Loại danh mục không hợp lệ.');
    if (!tenDanhMuc) throw new Error('Vui lòng nhập tên danh mục.');

    const ma = _generateNextNumber(_cauHinhSoDanhMuc(loai), _prefixDanhMuc(loai));
    _sheet(_sheetDanhMuc(loai)).appendRow([ma, tenDanhMuc, 'Hoạt động']);

    _writeAuditLog(currentUser.full_name, 'Danh Mục ' + (loai === 'Thu' ? 'Loại Thu' : 'Loại Chi'), 'Thêm', ma, '', tenDanhMuc);

    return _jsonOk({ ma: ma });

  } catch (err) {
    return _jsonErr(err);
  } finally {
    lock.releaseLock();
  }
}

function updateDanhMuc(loai, ma, tenMoi, trangThaiMoi, currentUser) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);

    if (!currentUser || !currentUser.username) throw new Error('Thiếu thông tin người dùng.');
    _yeuCauQuyen(currentUser.username, [ROLE_ADMIN]);

    if (loai !== 'Thu' && loai !== 'Chi') throw new Error('Loại danh mục không hợp lệ.');
    if (!ma) throw new Error('Thiếu mã danh mục cần sửa.');
    if (!tenMoi) throw new Error('Vui lòng nhập tên danh mục.');

    const sh = _sheet(_sheetDanhMuc(loai));
    const data = sh.getDataRange().getValues();
    const headers = data[0].map(h => String(h).trim());
    const idxMa = headers.indexOf(_colMaDanhMuc(loai));
    const idxTen = headers.indexOf(_colTenDanhMuc(loai));
    const idxTrangThai = headers.indexOf('trang_thai');

    let rowIndex = -1;
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][idxMa]) === String(ma)) { rowIndex = i; break; }
    }
    if (rowIndex === -1) throw new Error('Không tìm thấy danh mục: ' + ma);

    const tenCu = data[rowIndex][idxTen];

    sh.getRange(rowIndex + 1, idxTen + 1).setValue(tenMoi);
    sh.getRange(rowIndex + 1, idxTrangThai + 1).setValue(trangThaiMoi || 'Hoạt động');

    _writeAuditLog(currentUser.full_name, 'Danh Mục ' + (loai === 'Thu' ? 'Loại Thu' : 'Loại Chi'), 'Sửa', ma,
      tenCu, tenMoi + ' | ' + (trangThaiMoi || 'Hoạt động'));

    return _jsonOk({ ma: ma });

  } catch (err) {
    return _jsonErr(err);
  } finally {
    lock.releaseLock();
  }
}
