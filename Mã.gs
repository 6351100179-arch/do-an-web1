const MAT_KHAU_ADMIN = "123456"; // Bạn có thể đổi mã PIN tùy ý ở đây

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
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
      
      // Lưu hồ sơ
      dmSheet.appendRow([
        data.maID, 
        data.hoTen, 
        data.donVi, 
        data.chungChi, 
        "Đang làm", 
        imageUrl, 
        data.validTo || "", 
        data.loai || "worker", 
        data.mucDich || "",
        data.bienSo || ""
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

function doGet(e) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("DM_NHAN_SU");
    var data = sheet.getDataRange().getValues();
    return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify([])).setMimeType(ContentService.MimeType.JSON);
  }
}