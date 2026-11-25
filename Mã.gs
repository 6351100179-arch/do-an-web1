// ==========================================
// CẤU HÌNH BẢO MẬT
// ==========================================
const MAT_KHAU_ADMIN = "123456";        // PIN quản lý
const MAT_KHAU_BAOVE = "888888";        // PIN bảo vệ (ĐỔI TÙY Ý)
const SO_LAN_SAI_TOI_DA = 5;            // Sai 5 lần → khóa
const PHUT_KHOA = 5;                    // Khóa 5 phút
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

// Tìm cột theo tên tiêu đề (bỏ dấu, bỏ _ và khoảng trắng, không phân biệt hoa thường)
function chuanTen_(t) {
  return String(t || "").normalize("NFD").replace(/[\u0300-\u036f]/g, "")
    .replace(/đ/g, "d").replace(/Đ/g, "d").toLowerCase().replace(/[^a-z0-9]/g, "");
}
function timCot_(header, ten) {
  for (var i = 0; i < header.length; i++) if (chuanTen_(header[i]) === ten) return i;
  return -1;
}

// ==========================================
// BẢO MẬT BẢO VỆ
// ==========================================
function laySheetBaoMat_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName("BAO_MAT_BAOVE");
  if (!sh) {
    sh = ss.insertSheet("BAO_MAT_BAOVE");
    sh.appendRow(["Thời Gian", "Loại", "Chi Tiết", "Kết Quả"]);
  }
  return sh;
}

function kiemTraKhoa_() {
  var sh = laySheetBaoMat_();
  var rows = sh.getDataRange().getValues().slice(1);
  var bayGio = new Date().getTime();
  var soLanSai = 0;
  var lanSaiCuoi = 0;
  rows.forEach(function (r) {
    if (String(r[1]) === "login_fail") {
      var t = r[0] instanceof Date ? r[0].getTime() : new Date(r[0]).getTime();
      if (bayGio - t < PHUT_KHOA * 60 * 1000) {
        soLanSai++;
        if (t > lanSaiCuoi) lanSaiCuoi = t;
      }
    }
  });
  if (soLanSai >= SO_LAN_SAI_TOI_DA) {
    var conLai = Math.ceil((lanSaiCuoi + PHUT_KHOA * 60 * 1000 - bayGio) / 60000);
    return { locked: true, conLai: conLai > 0 ? conLai : 1 };
  }
  return { locked: false };
}

function ghiLogBaoMat_(loai, chiTiet, ketQua) {
  laySheetBaoMat_().appendRow([new Date(), loai, chiTiet || "", ketQua || ""]);
}

function xacThucBaoVe_(pin) {
  var khoa = kiemTraKhoa_();
  if (khoa.locked) {
    return { ok: false, message: "Đã khóa do sai quá " + SO_LAN_SAI_TOI_DA + " lần. Thử lại sau " + khoa.conLai + " phút." };
  }
  if (pin === MAT_KHAU_BAOVE) {
    return { ok: true };
  }
  ghiLogBaoMat_("login_fail", "Sai PIN bảo vệ", "FAIL");
  return { ok: false, message: "Mã PIN bảo vệ không chính xác." };
}

// ==========================================
// doPost
// ==========================================
function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    // ĐĂNG NHẬP QUẢN LÝ
    if (data.action === "login") {
      if (data.pin === MAT_KHAU_ADMIN) return json_({ "status": "success" });
      return json_({ "status": "error", "message": "Mã PIN không chính xác" });
    }

    // ĐĂNG NHẬP BẢO VỆ
    if (data.action === "login_guard") {
      var kq = xacThucBaoVe_(String(data.pin || ""));
      if (!kq.ok) return json_({ "status": "error", "message": kq.message });
      ghiLogBaoMat_("login_success", "Bảo vệ đăng nhập", "OK");
      return json_({ "status": "success", "message": "Đăng nhập bảo vệ thành công" });
    }

    // KIỂM TRA PHIÊN BẢO VỆ
    if (data.action === "check_guard_session") {
      var kqSess = xacThucBaoVe_(String(data.pin || ""));
      if (!kqSess.ok) return json_({ "status": "error", "message": kqSess.message, "needRelogin": true });
      return json_({ "status": "success" });
    }

    // QUẢN LÝ XÁC NHẬN "ĐÃ RA"
    if (data.action === "force_out") {
      if (data.pin !== MAT_KHAU_ADMIN) {
        return json_({ "status": "error", "message": "Mã PIN không chính xác. Bạn không có quyền!" });
      }
      var fLog = ss.getSheetByName("LOG_RA_VAO");
      if (!fLog) return json_({ "status": "error", "message": "Chưa có sheet LOG_RA_VAO" });
      fLog.appendRow([new Date(), data.maID, data.hoTen || "", "out", "Hợp lệ - Cho RA", "Quản lý xác nhận ra (quên check-out)"]);
      return json_({ "status": "success" });
    }

    // QUÉT MÃ (YÊU CẦU PIN BẢO VỆ)
    if (data.action === "scan") {
      var kqScan = xacThucBaoVe_(String(data.pin || ""));
      if (!kqScan.ok) {
        ghiLogBaoMat_("scan_fail", "Quét không có PIN hợp lệ: " + (data.maID || "?"), "BLOCK");
        return json_({ "status": "error", "message": kqScan.message, "needRelogin": true });
      }
      var logSheet = ss.getSheetByName("LOG_RA_VAO");
      if (!logSheet) {
        logSheet = ss.insertSheet("LOG_RA_VAO");
        logSheet.appendRow(["Thời Gian", "Mã ID", "Họ Tên", "Hành Động", "Kết Quả", "Ghi Chú"]);
      }
      logSheet.appendRow([new Date(), data.maID, data.hoTen, data.hanhDong || "", data.ketQua, data.ghiChu || ""]);
      return json_({ "status": "success" });
    }

    // LẤY THÔNG TIN CÁ NHÂN (CỔNG CÔNG NHÂN - HỖ TRỢ MÃ ID HOẶC SĐT)
    if (data.action === "my_info") {
      var query = String(data.query || data.maID || "").trim(); 
      if (!query) return json_({ "status": "error", "message": "Thiếu thông tin tra cứu (Mã ID hoặc SĐT)" });

      var dm = ss.getSheetByName("DM_NHAN_SU");
      if (!dm) return json_({ "status": "error", "message": "Chưa có sheet DM_NHAN_SU" });

      var dmData = dm.getDataRange().getValues();
      var header = dmData[0];
      var nhanSu = null;

      // Tìm theo Mã ID (cột A) hoặc SĐT (cột F) - đọc theo VỊ TRÍ CỘT
      for (var i = 1; i < dmData.length; i++) {
        var rowData = dmData[i];
        var maID = String(rowData[0]).trim();
        var sdt = String(rowData[5] || "").trim();

        if (maID.toLowerCase() === query.toLowerCase() ||
            (sdt !== "" && (sdt === query || "0" + sdt === query || sdt === "0" + query))) {
          nhanSu = rowData;
          break;
        }
      }

      if (!nhanSu) return json_({ "status": "error", "message": "Không tìm thấy Mã ID hoặc Số điện thoại này" });

      var maIDThucTe = String(nhanSu[0] || "").trim();
      var cotHieuLuc = timCot_(header, "hieuluctu");
      var hieuLucRaw = cotHieuLuc > -1 ? nhanSu[cotHieuLuc] : "";

      var log = ss.getSheetByName("LOG_RA_VAO");
      var lichSu = [];
      var homNay = Utilities.formatDate(new Date(), MUI_GIO, "yyyy-MM-dd");
      var baMuoiNgayTruoc = Utilities.formatDate(new Date(Date.now() - 30 * 24 * 3600 * 1000), MUI_GIO, "yyyy-MM-dd");

      if (log) {
        var logData = log.getDataRange().getValues().slice(1);
        logData.forEach(function (r) {
          if (String(r[1]).trim() !== maIDThucTe) return;
          var t = r[0] instanceof Date ? r[0] : new Date(r[0]);
          if (isNaN(t.getTime())) return;
          var ngay = Utilities.formatDate(t, MUI_GIO, "yyyy-MM-dd");
          if (ngay < baMuoiNgayTruoc) return;
          lichSu.push({
            thoiGian: Utilities.formatDate(t, MUI_GIO, "HH:mm:ss dd/MM/yyyy"),
            ts: t.getTime(),
            ngay: ngay,
            homNay: ngay === homNay,
            hanhDong: chuanHanhDong_(r[3], r[4]),
            ketQua: String(r[4] || ""),
            ghiChu: String(r[5] || "")
          });
        });
      }

      var trangThaiHienTai = "ra";
      var gioVaoGanNhat = null;
      for (var j = lichSu.length - 1; j >= 0; j--) {
        var h = lichSu[j];
        if (h.ketQua.indexOf("Hợp lệ") === 0) {
          if (h.hanhDong === "in") { trangThaiHienTai = "trong"; gioVaoGanNhat = h.thoiGian; break; }
          if (h.hanhDong === "out") { trangThaiHienTai = "ra"; break; }
        }
      }

      var soNgayCong = {};
      var tongMs = 0;
      var lastIn = null;
      var soCanhBao = 0;
      lichSu.forEach(function (h) {
        if (h.ketQua.indexOf("Hợp lệ") !== 0) {
          if (h.ketQua.indexOf("Cảnh báo") > -1) soCanhBao++;
          return;
        }
        if (h.hanhDong === "in") { lastIn = h.ts; soNgayCong[h.ngay] = true; }
        else if (h.hanhDong === "out" && lastIn !== null) {
          tongMs += h.ts - lastIn;
          lastIn = null;
        }
      });

      return json_({
        status: "success",
        nhanSu: {
          maID: maIDThucTe,
          hoTen: nhanSu[1] || "",
          donVi: nhanSu[2] || "",
          chungChi: nhanSu[3] || "",
          trangThai: nhanSu[4] || "",
          soDienThoai: nhanSu[5] || "",
          hanQR: nhanSu[6] ? Utilities.formatDate(new Date(nhanSu[6]), MUI_GIO, "dd/MM/yyyy") : "",
          hanQRRaw: nhanSu[6] || "",
          loai: nhanSu[7] || "worker",
          hieuLucTuRaw: hieuLucRaw || "",
          hieuLucTu: hieuLucRaw ? Utilities.formatDate(new Date(hieuLucRaw), MUI_GIO, "dd/MM/yyyy HH:mm") : ""
        },
        trangThaiHienTai: trangThaiHienTai,
        gioVaoGanNhat: gioVaoGanNhat,
        thongKe: {
          soNgayCong: Object.keys(soNgayCong).length,
          tongGio: Math.round(tongMs / 36000) / 100,
          soCanhBao: soCanhBao,
          soLuot: lichSu.length
        },
        lichSu: lichSu.slice(-50)
      });
    }

    // ĐĂNG KÝ MỚI
    if (data.action === "register") {
      if (data.pin !== MAT_KHAU_ADMIN) {
        return json_({ "status": "error", "message": "Mã PIN không chính xác. Bạn không có quyền đăng ký!" });
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
      return json_({ "status": "success", "message": "Đăng ký thành công", "qrCodeUrl": qrUrl });
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