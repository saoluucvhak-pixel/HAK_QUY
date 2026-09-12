function myFunction() {
  /*************************************************
 * FIX_SOPHIEU_CONGDOAN.GS
 * Sửa lỗi: số phiếu thu/chi của Quỹ Công đoàn (dạng "01/02",
 * "01/03"...) bị Google Sheets tự động hiểu nhầm thành NGÀY
 * THÁNG khi importDuLieuSoQuyCu() ghi vào sheet bằng appendRow
 * (dù đã ép định dạng Plain text trước khi ghi, Sheets vẫn tự
 * suy luận lại). Không ảnh hưởng số tiền/ngày chứng từ/tồn quỹ
 * — chỉ ảnh hưởng CỘT SỐ PHIẾU hiển thị.
 *
 * Cách sửa: đổi số phiếu Công đoàn sang dạng không thể bị hiểu
 * nhầm thành ngày (thêm tiền tố "CD-", ví dụ "01/02" -> "CD-01/02").
 *
 * CÁCH DÙNG: dán file này vào Apps Script Editor, chọn hàm
 * "suaSoPhieuCongDoanBiSai" ở dropdown trên cùng, bấm Run.
 * Chỉ cần chạy 1 lần.
 *************************************************/

function suaSoPhieuCongDoanBiSai() {
  const thuMap = {
    "PT_IMPORT2_0013": "CD-01/02",
    "PT_IMPORT2_0014": "CD-01/03",
    "PT_IMPORT2_0015": "CD-01/03",
    "PT_IMPORT2_0016": "CD-01/03",
    "PT_IMPORT2_0017": "CD-02/03",
    "PT_IMPORT2_0018": "CD-01/05",
    "PT_IMPORT2_0019": "CD-02/05",
    "PT_IMPORT2_0020": "CD-01/06",
    "PT_IMPORT2_0021": "CD-02/06",
    "PT_IMPORT2_0022": "CD-01/07",
    "PT_IMPORT2_0023": "CD-02/07",
    "PT_IMPORT2_0024": "CD-01/08",
    "PT_IMPORT2_0025": "CD-02/08"
  };
  const chiMap = {
    "PC_IMPORT2_0025": "CD-01/01",
    "PC_IMPORT2_0026": "CD-02/01",
    "PC_IMPORT2_0027": "CD-01/02",
    "PC_IMPORT2_0028": "CD-01/03",
    "PC_IMPORT2_0029": "CD-02/03",
    "PC_IMPORT2_0030": "CD-03/03",
    "PC_IMPORT2_0031": "CD-04/03",
    "PC_IMPORT2_0032": "CD-01/04",
    "PC_IMPORT2_0033": "CD-02/04",
    "PC_IMPORT2_0034": "CD-01/05",
    "PC_IMPORT2_0035": "CD-02/05",
    "PC_IMPORT2_0036": "CD-03/05",
    "PC_IMPORT2_0037": "CD-04/05",
    "PC_IMPORT2_0038": "CD-05/05",
    "PC_IMPORT2_0039": "CD-01/06",
    "PC_IMPORT2_0040": "CD-02/06",
    "PC_IMPORT2_0041": "CD-01/07",
    "PC_IMPORT2_0042": "CD-01/08",
    "PC_IMPORT2_0043": "CD-02/08",
    "PC_IMPORT2_0044": "CD-03/08"
  };

  function suaMotSheet(sh, idMap) {
    const data = sh.getDataRange().getValues();
    const headers = data[0].map(h => String(h).trim());
    const idxId = headers.indexOf("id");
    const idxSoPhieu = 1; // cột B, ngay sau id
    let soDaSua = 0;
    for (let i = 1; i < data.length; i++) {
      const id = data[i][idxId];
      if (idMap.hasOwnProperty(id)) {
        const cell = sh.getRange(i + 1, idxSoPhieu + 1);
        cell.setNumberFormat("@");
        cell.setValue(idMap[id]);
        soDaSua++;
      }
    }
    return soDaSua;
  }

  const soThuDaSua = suaMotSheet(_sheet(SHEET_PHIEU_THU), thuMap);
  const soChiDaSua = suaMotSheet(_sheet(SHEET_PHIEU_CHI), chiMap);

  SpreadsheetApp.getUi().alert(
    "Đã sửa số phiếu Quỹ Công đoàn!\n\n" +
    "Phiếu Thu đã sửa: " + soThuDaSua + "\n" +
    "Phiếu Chi đã sửa: " + soChiDaSua
  );
}

}
