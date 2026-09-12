/*************************************************
 * AUTH.GS
 * Xử lý đăng nhập. Phân quyền chi tiết (chặn theo
 * role ở từng chức năng) sẽ hoàn thiện ở GIAI ĐOẠN 5.
 * Giai đoạn này: xác thực đúng tài khoản + trả về role
 * để giao diện hiển thị đúng menu.
 *************************************************/

/**
 * Kiểm tra đăng nhập.
 * @param {string} username
 * @param {string} password
 * @returns {{success:boolean, data?:object, message?:string}}
 */
function loginUser(username, password) {
  try {
    if (!username || !password) {
      throw new Error('Vui lòng nhập đầy đủ Tên đăng nhập và Mật khẩu.');
    }

    const users = _sheetToObjects(SHEET_USERS);
    const user = users.find(u => String(u.username).toLowerCase() === String(username).toLowerCase().trim());

    if (!user) throw new Error('Tên đăng nhập không tồn tại.');
    if (String(user.status) !== USER_STATUS_ACTIVE) throw new Error('Tài khoản này đã bị khóa. Liên hệ ADMIN.');

    const hash = _hashPassword(password);
    if (String(user.password_hash) !== hash) throw new Error('Mật khẩu không đúng.');

    return _jsonOk({
      user_id: user.user_id,
      username: user.username,
      full_name: user.full_name,
      role: user.role
    });

  } catch (err) {
    return _jsonErr(err);
  }
}

/**
 * Hàm tiện ích cho ADMIN tạo người dùng mới trực tiếp từ Apps Script Editor
 * (chức năng tạo user qua giao diện sẽ làm ở phần QUẢN TRỊ - Giai đoạn 5).
 * Cách dùng: sửa các giá trị bên dưới rồi chạy hàm này 1 lần.
 */
function _taoNguoiDungMoi_ChayThuCong() {
  const sh = _sheet(SHEET_USERS);
  const newUser = {
    user_id: 'USR' + new Date().getTime(),
    username: 'thuquy1',           // <-- sửa tên đăng nhập tại đây
    password: '123456',            // <-- sửa mật khẩu tại đây
    full_name: 'Nguyễn Thị Thủ Quỹ', // <-- sửa họ tên tại đây
    role: ROLE_THU_QUY              // <-- ROLE_ADMIN / ROLE_THU_QUY / ROLE_XEM
  };

  sh.appendRow([
    newUser.user_id,
    newUser.username,
    _hashPassword(newUser.password),
    newUser.full_name,
    newUser.role,
    USER_STATUS_ACTIVE,
    new Date()
  ]);

  Logger.log('Đã tạo user: ' + newUser.username + ' / role: ' + newUser.role);
}
