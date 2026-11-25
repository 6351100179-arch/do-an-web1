const MAT_KHAU_ADMIN = "123456"; // Bạn có thể đổi mã PIN tùy ý ở đây (nên đặt dài hơn 6 số)
const MUI_GIO = "Asia/Ho_Chi_Minh";

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// Đổi chuỗi "yyyy-MM-dd[THH:mm[:ss]]" (giờ Việt Nam) thành Date để Sheet lưu đúng giờ
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

// Xác định hành động in/out từ cột "Hành Động" hoặc từ chữ trong "Kết Quả"
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
    // ĐĂNG NHẬP QUẢN LÝ: kiểm tra mã PIN trên server
    // ==========================================
    if (data.action === "login") {
      if (data.pin === MAT_KHAU_ADMIN) return json_({ "status": "success" });
      return json_({ "status": "error", "message": "Mã PIN không chính xác" });
    }

    // ==========================================
    // QUẢN LÝ XÁC NHẬN "ĐÃ RA" (cho người quên check-out) - YÊU CẦU MÃ PIN
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
    // TRƯỜNG HỢP 1: QUÉT MÃ (Ghi log ra vào) - Ai cũng làm được
    // ==========================================
    if (data.action === "scan") {
      var logSheet = ss.getSheetByName("LOG_RA_VAO");
      if(!logSheet) {
        logSheet = ss.insertSheet("LOG_RA_VAO");
        logSheet.appendRow(["Thời Gian", "Mã ID", "Họ Tên", "Hành Động", "Kết Quả", "Ghi Chú"]);
      }
      logSheet.appendRow([new Date(), data.maID, data.hoTen, data.hanhDong || "", data.ketQua, data.ghiChu || ""]);
      return ContentService.createTextOutput(JSON.stringify({"status": "success"})).setMimeType(ContentService.MimeType.JSON);
    }
    
    // ==========================================
    // TRƯỜNG HỢP 2: CẬP NHẬT SỐ LẦN DÙNG
    // ==========================================
    if (data.action === "update_usage") {
      var dmSheet = ss.getSheetByName("DM_NHAN_SU");
      var lastRow = dmSheet.getLastRow();
      var dataRange = dmSheet.getRange(1, 1, lastRow, 10).getValues(); 
      
      for (var i = 0; i < dataRange.length; i++) {
        if (dataRange[i][0] == data.maID) {
          var rowIndex = i + 1;
          var currentSoLan = parseInt(dataRange[i][9]) || 0;
          if (currentSoLan > 0) {
            dmSheet.getRange(rowIndex, 10).setValue(currentSoLan - 1);
          }
          break;
        }
      }
      return ContentService.createTextOutput(JSON.stringify({"status": "success"})).setMimeType(ContentService.MimeType.JSON);
    }
    
    // ==========================================
    // TRƯỜNG HỢP 3: ĐĂNG KÝ MỚI (YÊU CẦU MÃ PIN)
    // ==========================================
    if (data.action === "register") {
      // KIỂM TRA MÃ PIN BẢO MẬT
      if (data.pin !== MAT_KHAU_ADMIN) {
        return ContentService.createTextOutput(JSON.stringify({
          "status": "error", 
          "message": "Mã PIN không chính xác. Bạn không có quyền đăng ký!"
        })).setMimeType(ContentService.MimeType.JSON);
      }

      var imageUrl = "";
      
      // Xử lý lưu ảnh thẻ (nếu có)
      if (data.image && data.image.includes(',')) {
        try {
          var folderName = "Anh_Nhan_Su_Cong_Truong";
          var folders = DriveApp.getFoldersByName(folderName);
          var folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(folderName);
          
          var splitBase = data.image.split(',');
          var contentType = splitBase[0].split(':')[1].split(';')[0];
          var bytes = Utilities.base64Decode(splitBase[1]);
          var blob = Utilities.newBlob(bytes, contentType, data.maID + "_" + data.hoTen + ".jpg");
          var file = folder.createFile(blob);
          file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
          imageUrl = file.getUrl();
        } catch (imgError) {
          console.error("Lỗi lưu ảnh: " + imgError);
        }
      }
      
      var dmSheet = ss.getSheetByName("DM_NHAN_SU");
      
      // Lưu hồ sơ (cột G = hết hiệu lực, cột K = có hiệu lực từ; lưu dạng ngày giờ thật)
      dmSheet.appendRow([
        data.maID, 
        data.hoTen, 
        data.donVi, 
        data.chungChi, 
        "Đang làm", 
        imageUrl, 
        parseGio_(data.validTo), 
        data.loai || "worker", 
        data.mucDich || "",
        data.bienSo || "",
        parseGio_(data.validFrom)
      ]);
      
      var qrUrl = "https://api.qrserver.com/v1/create-qr-code/?size=300x300&data=" + data.maID;
      
      return ContentService.createTextOutput(JSON.stringify({
        "status": "success", 
        "message": "Đăng ký thành công",
        "imageUrl": imageUrl, 
        "qrCodeUrl": qrUrl
      })).setMimeType(ContentService.MimeType.JSON);
    }
    
    return ContentService.createTextOutput(JSON.stringify({"status": "error", "message": "Action không hợp lệ"})).setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({"status": "error", "message": error.toString()})).setMimeType(ContentService.MimeType.JSON);
  }
}

// ==========================================
// LỊCH SỬ RA VÀO từ sheet LOG_RA_VAO
// Cột: A=Thời Gian, B=Mã ID, C=Họ Tên, D=Hành Động, E=Kết Quả, F=Ghi Chú
// - Mặc định: chỉ trả lượt quét của HÔM NAY (nhẹ, cho điện thoại quét)
// - from/to (yyyy-MM-dd): trả lượt quét trong khoảng ngày
// - quenCheckout: người có lượt hợp lệ gần nhất là VÀO của ngày trước hôm nay
// ==========================================
function layLichSu_(from, to) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("LOG_RA_VAO");
  if (!sh) return { history: [], quenCheckout: [] };

  var homNay = Utilities.formatDate(new Date(), MUI_GIO, "yyyy-MM-dd");
  var tuNgay = from || homNay;
  var denNgay = to || homNay;
  var rows = sh.getDataRange().getValues().slice(1); // bỏ dòng tiêu đề
  var history = [];
  var last = {}; // maID -> lượt hợp lệ gần nhất

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

  return { history: history, quenCheckout: quen }; // history theo thứ tự cũ -> mới
}

function doGet(e) {
  try {
    // Lịch sử: ?action=history  (hôm nay)  |  ?action=history&from=...&to=...&pin=...  (khoảng ngày, cần PIN)
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

    // Mặc định: trả danh sách nhân sự (như cũ)
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("DM_NHAN_SU");
    var data = sheet.getDataRange().getValues();
    return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify([])).setMimeType(ContentService.MimeType.JSON);
  }
}
