/*************************************************
 * CONFIG.GS
 * Nơi khai báo TOÀN BỘ tên Sheet, hằng số dùng chung.
 * Muốn đổi tên 1 sheet, chỉ cần sửa ở đây, không cần
 * sửa rải rác trong các file khác.
 *************************************************/

// ----- Tên các Sheet trong Database -----
const SHEET_USERS               = 'USERS';
const SHEET_DOITUONG             = 'DM_DOITUONG';
const SHEET_LOAI_THU             = 'DM_LOAI_THU';
const SHEET_LOAI_CHI             = 'DM_LOAI_CHI';
const SHEET_PHIEU_THU            = 'PHIEU_THU';
const SHEET_PHIEU_CHI            = 'PHIEU_CHI';
const SHEET_CONG_NO              = 'CONG_NO';
const SHEET_THANH_TOAN_CONG_NO   = 'THANH_TOAN_CONG_NO';
const SHEET_KIEM_KE_QUY          = 'KIEM_KE_QUY';
const SHEET_KHOA_SO              = 'KHOA_SO';
const SHEET_AUDIT_LOG            = 'AUDIT_LOG';
const SHEET_CAU_HINH             = 'CAU_HINH';

// ----- Vai trò người dùng -----
const ROLE_ADMIN   = 'ADMIN';
const ROLE_THU_QUY = 'THU_QUY';
const ROLE_XEM     = 'XEM';

// ----- Trạng thái chứng từ -----
const TRANG_THAI_HOP_LE = 'Hợp lệ';
const TRANG_THAI_DA_HUY = 'Đã hủy';

// ----- Trạng thái tài khoản người dùng -----
const USER_STATUS_ACTIVE = 'Hoạt động';
const USER_STATUS_LOCKED = 'Khóa';

// ----- Trạng thái khóa sổ -----
const KHOA_SO_DA_KHOA = 'Đã khóa';
const KHOA_SO_DA_MO   = 'Đã mở';
