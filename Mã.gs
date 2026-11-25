const MAT_KHAU_ADMIN = "123456"; // Đổi PIN tùy ý
const MUI_GIO = "Asia/Ho_Chi_Minh";

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// Đổi chuỗi "yyyy-MM-dd[THH:mm[:ss]]" (giờ VN) thành Date
function parseGio_(s) {
  if (!s) return "";
  try {
    var t = String(s).replace("T", " ");
    var fmt = t.length <= 10 ? "yyyy-MM-dd" : (t.length <= 16 ? "yyyy-MM-dd HH:mm" : "yyyy-MM-dd HH:mm:ss");
    return Utilities.parseDate(t, MUI_GIO, fmt);
  } catch (err) {
    return String(s);
  }
}

// Xác định hành động in/out
function chuanHanhDong_(hd, kq) {
  var a = String(hd || "").trim().toLowerCase();
  if (a === "in" || a === "out") return a;
  var k = String(kq || "").toUpperCase();
  if (k.indexOf("CHO RA") > -1) return "out";
  if (k.indexOf("CHO VÀO") > -1) return "in";
  return "";
}

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    // ==========================================
    // ĐĂNG NHẬP QUẢN LÝ
    // ==========================================
    if (data.action === "login") {
      if (data.pin === MAT_KHAU_ADMIN) return json_({ "status": "success" });
      return json_({ "status": "error", "message": "Mã PIN không chính xác" });
    }

    // ==========================================
    // QUẢN LÝ XÁC NHẬN "ĐÃ RA"
    // ==========================================
    if (data.action === "force_out") {
      if (data.pin !== MAT_KHAU_ADMIN) {
        return json_({ "status": "error", "message": "Mã PIN không chính xác. Bạn không có quyền!" });
      }
      var fLog = ss.getSheetByName("LOG_RA_VAO");
      if (!fLog) return json_({ "status": "error", "message": "Chưa có sheet LOG_RA_VAO" });
      fLog.appendRow([new Date(), data.maID, data.hoTen || "", "out", "Hợp lệ - Cho RA", "Quản lý xác nhận ra (quên check-out)"]);
      return json_({ "status": "success" });
    }

    // ==========================================
    // QUÉT MÃ (Ghi log ra vào)
    // ==========================================
    if (data.action === "scan") {
      var logSheet = ss.getSheetByName("LOG_RA_VAO");
      if (!logSheet) {
        logSheet = ss.insertSheet("LOG_RA_VAO");
        logSheet.appendRow(["Thời Gian", "Mã ID", "Họ Tên", "Hành Động", "Kết Quả", "Ghi Chú"]);
      }
      logSheet.appendRow([new Date(), data.maID, data.hoTen, data.hanhDong || "", data.ketQua, data.ghiChu || ""]);
      return json_({ "status": "success" });
    }

    // ==========================================
    // ĐĂNG KÝ MỚI (YÊU CẦU MÃ PIN) — KHÔNG LƯU ẢNH, CÓ SỐ ĐIỆN THOẠI
    // ==========================================
    if (data.action === "register") {
      if (data.pin !== MAT_KHAU_ADMIN) {
        return json_({
          "status": "error",
          "message": "Mã PIN không chính xác. Bạn không có quyền đăng ký!"
        });
      }

      var dmSheet = ss.getSheetByName("DM_NHAN_SU");
      if (!dmSheet) {
        dmSheet = ss.insertSheet("DM_NHAN_SU");
        dmSheet.appendRow([
          "Mã ID", "Họ Tên", "Đơn Vị", "Chứng Chỉ AT", "Trạng Thái",
          "Số Điện Thoại", "Hạn Mã QR", "Loại", "Mục Đích",
          "Biển Số", "Hiệu Lực Từ", "CCCD"
        ]);
      }

      // Thứ tự cột (KHÔNG còn ảnh):
      // A=0 maID | B=1 hoTen | C=2 donVi | D=3 chungChi | E=4 trangThai
      // F=5 soDienThoai | G=6 validTo | H=7 loai | I=8 mucDich
      // J=9 bienSo | K=10 validFrom | L=11 cccd
      dmSheet.appendRow([
        data.maID,
        data.hoTen,
        data.donVi,
        data.chungChi,
        "Đang làm",
        data.soDienThoai || "",
        parseGio_(data.validTo),
        data.loai || "worker",
        data.mucDich || "",
        data.bienSo || "",
        parseGio_(data.validFrom),
        data.cccd || ""
      ]);

      var qrUrl = "https://api.qrserver.com/v1/create-qr-code/?size=300x300&data=" + data.maID;

      return json_({
        "status": "success",
        "message": "Đăng ký thành công",
        "qrCodeUrl": qrUrl
      });
    }

    return json_({ "status": "error", "message": "Action không hợp lệ" });

  } catch (error) {
    return json_({ "status": "error", "message": error.toString() });
  }
}

// ==========================================
// LỊCH SỬ RA VÀO
// ==========================================
function layLichSu_(from, to) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("LOG_RA_VAO");
  if (!sh) return { history: [], quenCheckout: [] };

  var homNay = Utilities.formatDate(new Date(), MUI_GIO, "yyyy-MM-dd");
  var tuNgay = from || homNay;
  var denNgay = to || homNay;
  var rows = sh.getDataRange().getValues().slice(1);
  var history = [];
  var last = {};

  rows.forEach(function (r) {
    var t = r[0];
    if (!(t instanceof Date)) t = new Date(t);
    if (isNaN(t.getTime())) return;

    var ngay = Utilities.formatDate(t, MUI_GIO, "yyyy-MM-dd");
    var maID = String(r[1]);
    var hanhDong = chuanHanhDong_(r[3], r[4]);

    if (String(r[4]).indexOf("Hợp lệ") === 0 && hanhDong) {
      last[maID] = {
        hanhDong: hanhDong,
        hoTen: r[2],
        thoiGian: Utilities.formatDate(t, MUI_GIO, "HH:mm:ss dd/MM/yyyy"),
        ngay: Utilities.formatDate(t, MUI_GIO, "dd/MM/yyyy"),
        ngayISO: ngay
      };
    }

    if (ngay >= tuNgay && ngay <= denNgay) {
      history.push({
        thoiGian: Utilities.formatDate(t, MUI_GIO, "HH:mm:ss dd/MM/yyyy"),
        ts: t.getTime(),
        homNay: ngay === homNay,
        maID: maID,
        hoTen: r[2],
        hanhDong: r[3],
        ketQua: r[4],
        ghiChu: r[5]
      });
    }
  });

  var quen = [];
  Object.keys(last).forEach(function (id) {
    var x = last[id];
    if (x.hanhDong === "in" && x.ngayISO < homNay) {
      quen.push({ maID: id, hoTen: x.hoTen, thoiGian: x.thoiGian, ngay: x.ngay });
    }
  });

  return { history: history, quenCheckout: quen };
}

function doGet(e) {
  try {
    if (e && e.parameter && e.parameter.action === "history") {
      var from = e.parameter.from, to = e.parameter.to;
      if (from || to) {
        if (e.parameter.pin !== MAT_KHAU_ADMIN) {
          return json_({ "status": "error", "message": "Mã PIN không chính xác. Bạn không có quyền xem lịch sử." });
        }
        var re = /^\d{4}-\d{2}-\d{2}$/;
        if ((from && !re.test(from)) || (to && !re.test(to))) {
          return json_({ "status": "error", "message": "Ngày không hợp lệ (cần dạng yyyy-MM-dd)" });
        }
      }
      return json_(layLichSu_(from, to));
    }

    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("DM_NHAN_SU");
    var data = sheet.getDataRange().getValues();
    return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify([])).setMimeType(ContentService.MimeType.JSON);
  }
}