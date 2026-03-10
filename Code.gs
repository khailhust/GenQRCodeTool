function taoHoSoDongDangChon() {
  // --- CẤU HÌNH (Đã giữ nguyên ID của bạn) ---
  const TEMPLATE_ID = '1dM7Yj-YdBkev99LlAo4opwELQ4fF3cEisId8mJEUE5E'; // ID file mẫu Docs
  const DATA_FOLDER_ID = '1Mcg3wG8HtUwSvGxFD0Pfktqlj4qUJkDc'; // Thư mục tổng chứa các thư mục con
  // ----------------

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();

  // 1. XÁC ĐỊNH DÒNG ĐANG CHỌN
  const activeRange = sheet.getActiveRange();
  const rowIndex = activeRange.getRow();

  if (rowIndex === 1) {
    ss.toast("Vui lòng chọn dòng khách hàng (không chọn tiêu đề).", "⚠️ Chọn sai dòng", 3);
    return;
  }

  const rowData = sheet.getRange(rowIndex, 4, 1, 6).getValues()[0];

  let tenKH = rowData[0];     // Cột D
  let loaiHS = rowData[1];    // Cột E
  let diaChi = rowData[2];    // Cột F
  let linkDrive = rowData[3]; // Cột G
  let status = rowData[4];    // Cột H

  if (!tenKH) {
    ss.toast("Dòng này không có tên khách hàng!", "⚠️ Dữ liệu trống", 3);
    return;
  }

  ss.toast("Đang xử lý hồ sơ: " + tenKH + "...", "⏳ Đang chạy", -1);

  const rootDataFolder = DriveApp.getFolderById(DATA_FOLDER_ID);

  try {
    let targetFolder;

    // --- BƯỚC 1: XÁC ĐỊNH THƯ MỤC LƯU TRỮ ---
    if (!linkDrive || linkDrive === "") {
      let folderName = tenKH + "-" + diaChi + " - Hồ sơ";
      targetFolder = rootDataFolder.createFolder(folderName);
      targetFolder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      linkDrive = targetFolder.getUrl();
      sheet.getRange(rowIndex, 7).setValue(linkDrive);
      SpreadsheetApp.flush();
    } else {
      let folderId = trichXuatFolderId(linkDrive);
      if (folderId) {
        targetFolder = DriveApp.getFolderById(folderId);
      } else {
        throw new Error("Link Drive không hợp lệ, không tìm thấy Folder.");
      }
    }

    // --- BƯỚC 2: TẠO FILE DOCS TẠM THỜI ---
    let tenFileMoi = `${loaiHS} - ${tenKH} - QR`;
    let templateFile = DriveApp.getFileById(TEMPLATE_ID);
    let tempFile = templateFile.makeCopy(tenFileMoi, targetFolder);
    let tempDoc = DocumentApp.openById(tempFile.getId());
    let body = tempDoc.getBody();

    body.replaceText("{{TEN_KHACH_HANG}}", tenKH);
    body.replaceText("{{LOAI_HO_SO}}", loaiHS);
    body.replaceText("{{DIA_CHI}}", diaChi);
    body.replaceText("{{LINK_DRIVE}}", linkDrive);

    // Tạo QR Code
    xuLyChenQR(body, "QR_BAN_VE", linkDrive);

    tempDoc.saveAndClose();

    // --- BƯỚC 3: DỌN RÁC VÀ XUẤT PDF ---
    dondepFileCu(targetFolder, tenFileMoi);
    let pdfBlob = tempFile.getAs(MimeType.PDF);
    let pdfFile = targetFolder.createFile(pdfBlob).setName(tenFileMoi + ".pdf");
    tempFile.setTrashed(true);

    // --- BƯỚC 4: CẬP NHẬT TRẠNG THÁI ---
    sheet.getRange(rowIndex, 8).setValue("Đã tạo");
    sheet.getRange(rowIndex, 9).setValue(pdfFile.getUrl());

    // --- THÊM MỚI: GHI LOG THÀNH CÔNG ---
    ghiLogHeThong(tenKH, "✅ Tạo mới thành công", "Link PDF: " + pdfFile.getUrl());

    ss.toast("Đã xong hồ sơ cho: " + tenKH, "✅ Hoàn tất", 5);

  } catch (e) {
    Logger.log("Lỗi: " + e.toString());
    sheet.getRange(rowIndex, 8).setValue("Lỗi: " + e.message);
    ss.toast("Gặp lỗi: " + e.message, "❌ Thất bại", 5);

    // --- THÊM MỚI: GHI LOG LỖI ---
    ghiLogHeThong(tenKH, "❌ LỖI TẠO MỚI: " + e.message, e.stack);
  }
}

// ======================================================================
// --- HÀM XỬ LÝ QR (Cập nhật: Thu thập và báo cáo chi tiết lỗi từng Server) ---
// ======================================================================
function xuLyChenQR(body, altText, linkRaw) {
  let cleanLink = chuanHoaLinkHienThi(linkRaw);
  if (!cleanLink || cleanLink === "") {
    throw new Error("Link Drive trống, không thể tạo QR.");
  }

  // Danh sách 3 Server tạo QR
  const qrApis = [
    `https://quickchart.io/qr?text=${encodeURIComponent(cleanLink)}&size=300&margin=1&ecLevel=M`,
    `https://api.qrserver.com/v1/create-qr-code/?size=300x300&data=${encodeURIComponent(cleanLink)}&margin=2`,
    `https://chart.googleapis.com/chart?chs=300x300&cht=qr&chl=${encodeURIComponent(cleanLink)}`
  ];

  let blob = null;
  let danhSachLoi = []; // Biến chứa "hồ sơ bệnh án" của các server bị lỗi

  // Thử lần lượt từng API
  for (let i = 0; i < qrApis.length; i++) {
    try {
      let response = UrlFetchApp.fetch(qrApis[i], { muteHttpExceptions: true });

      if (response.getResponseCode() === 200) {
        blob = response.getBlob();
        break; // Thành công thì lập tức thoát vòng lặp
      } else {
        // Nếu Server phản hồi nhưng từ chối (Vd: Lỗi 403 Forbidden, Lỗi 429 Quá tải...)
        let errorMsg = response.getContentText().substring(0, 100); // Lấy 100 ký tự đầu cho đỡ rác log
        danhSachLoi.push(`Server ${i + 1} (Mã HTTP ${response.getResponseCode()}): ${errorMsg}`);
      }
    } catch (e) {
      // Nếu Server chết hẳn, không phản hồi hoặc báo timeout
      danhSachLoi.push(`Server ${i + 1} (Không phản hồi): ${e.message}`);
    }
  }

  // --- NẾU CẢ 3 SERVER ĐỀU CHẾT ---
  if (!blob) {
    // Gom tất cả lỗi thành một thông báo dài và ném ra ngoài để hàm chính ghi vào Sheet Logs
    let thongBaoLoiTongHop = "Cả 3 server QR đều thất bại:\n- " + danhSachLoi.join("\n- ");
    throw new Error(thongBaoLoiTongHop);
  }

  // --- TIẾN HÀNH CHÈN ẢNH VÀO DOCS ---
  const images = body.getImages();
  let isReplaced = false;

  for (let img of images) {
    if (img.getAltDescription() === altText || img.getAltTitle() === altText) {
      let parent = img.getParent();
      let index = parent.getChildIndex(img);

      let newImg = parent.asParagraph().insertInlineImage(index, blob);
      newImg.setWidth(img.getWidth()).setHeight(img.getHeight());

      img.removeFromParent();
      isReplaced = true;
      return;
    }
  }

  // Nếu không tìm thấy ảnh để thay thế
  if (!isReplaced) {
    throw new Error("Template bị lỗi: Không tìm thấy ảnh nháp nào có Alt Text là '" + altText + "'.");
  }
}

function trichXuatFolderId(link) {
  if (!link) return null;
  let strLink = String(link).trim();
  let match = strLink.match(/[-\w]{25,}/);
  if (match) return match[0];
  return null;
}

function chuanHoaLinkHienThi(link) {
  let id = trichXuatFolderId(link);
  if (id) {
    return "https://drive.google.com/drive/folders/" + id;
  }
  return String(link).trim();
}

function regenToanBoHoSo() {
  let ui = SpreadsheetApp.getUi();

  // --- 🔒 CHỐT CHẶN BẢO MẬT ADMIN ---
  let passPrompt = ui.prompt(
    '🔒 Yêu cầu quyền Admin',
    'Tính năng này chỉ dành cho Admin.\nVui lòng nhập mã PIN để tiếp tục:',
    ui.ButtonSet.OK_CANCEL
  );

  if (passPrompt.getSelectedButton() !== ui.Button.OK) { return; }
  if (passPrompt.getResponseText() !== "686868") {
    ui.alert("❌ Sai mã PIN!", "Bạn không có quyền sử dụng tính năng này.", ui.ButtonSet.OK);
    return;
  }

  const TEMPLATE_ID = '1dM7Yj-YdBkev99LlAo4opwELQ4fF3cEisId8mJEUE5E';
  const DATA_FOLDER_ID = '1Mcg3wG8HtUwSvGxFD0Pfktqlj4qUJkDc';

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const data = sheet.getDataRange().getValues();

  let response = ui.alert('Xác nhận Tạo lại', 'Bạn có chắc chắn muốn TẠO LẠI (Regen) toàn bộ hồ sơ trong Sheet này không? Quá trình này sẽ chạy qua tất cả các dòng có tên khách hàng.', ui.ButtonSet.YES_NO);
  if (response != ui.Button.YES) return;

  let totalValid = 0;
  for (let i = 1; i < data.length; i++) {
    if (data[i][3]) totalValid++;
  }

  if (totalValid === 0) {
    ss.toast("Không tìm thấy khách hàng nào để xử lý!", "⚠️ Trống", 3);
    return;
  }

  ss.toast("Bắt đầu quét " + totalValid + " hồ sơ...", "⏳ Đang khởi động", 3);

  let rootDataFolder;
  try {
    rootDataFolder = DriveApp.getFolderById(DATA_FOLDER_ID);
  } catch(e) {
    ss.toast("Lỗi truy cập Folder Data Tổng.", "❌ Lỗi", 5);
    return;
  }

  let count = 0;

  for (let i = 1; i < data.length; i++) {
    let tenKH = data[i][3];
    let loaiHS = data[i][4];
    let diaChi = data[i][5];
    let linkDrive = data[i][6];

    if (!tenKH) continue;

    count++;
    ss.toast(`Đang xử lý: ${tenKH} (${count}/${totalValid})`, "🔄 Tiến trình", 4);

    sheet.getRange(i + 1, 8).setValue("⏳ Đang xử lý...");
    SpreadsheetApp.flush();

    try {
      let targetFolder;

      if (!linkDrive || linkDrive === "") {
        let folderName = tenKH + "-" + diaChi + " - Hồ sơ";
        targetFolder = rootDataFolder.createFolder(folderName);
        targetFolder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        linkDrive = targetFolder.getUrl();
        sheet.getRange(i + 1, 7).setValue(linkDrive);
      } else {
        let folderId = trichXuatFolderId(linkDrive);
        if (folderId) {
          targetFolder = DriveApp.getFolderById(folderId);
        } else {
          throw new Error("Link Drive cũ không hợp lệ.");
        }
      }

      let tenFileMoi = `${loaiHS} - ${tenKH} - QR`;
      let templateFile = DriveApp.getFileById(TEMPLATE_ID);
      let tempFile = templateFile.makeCopy(tenFileMoi, targetFolder);
      let tempDoc = DocumentApp.openById(tempFile.getId());
      let body = tempDoc.getBody();

      body.replaceText("{{TEN_KHACH_HANG}}", tenKH);
      body.replaceText("{{LOAI_HO_SO}}", loaiHS);
      body.replaceText("{{DIA_CHI}}", diaChi);
      body.replaceText("{{LINK_DRIVE}}", linkDrive);

      xuLyChenQR(body, "QR_BAN_VE", linkDrive);
      tempDoc.saveAndClose();

      dondepFileCu(targetFolder, tenFileMoi);
      let pdfBlob = tempFile.getAs(MimeType.PDF);
      let pdfFile = targetFolder.createFile(pdfBlob).setName(tenFileMoi + ".pdf");
      tempFile.setTrashed(true);

      sheet.getRange(i + 1, 8).setValue("✅ Đã Regen");
      sheet.getRange(i + 1, 9).setValue(pdfFile.getUrl());
      SpreadsheetApp.flush();

      // --- THÊM MỚI: GHI LOG THÀNH CÔNG ---
      ghiLogHeThong(tenKH, "✅ Regen thành công", "Link PDF mới: " + pdfFile.getUrl());

    } catch (e) {
      Logger.log("Lỗi dòng " + (i+1) + ": " + e.toString());
      sheet.getRange(i + 1, 8).setValue("❌ Lỗi: " + e.message);
      SpreadsheetApp.flush();

      // --- THÊM MỚI: GHI LOG LỖI ---
      ghiLogHeThong(tenKH, "❌ LỖI REGEN: " + e.message, e.stack);
    }
  }

  ss.toast("Đã hoàn tất " + count + " hồ sơ!", "🎉 Xong toàn bộ", 10);
}

function dondepFileCu(folder, tenFile) {
  let files = folder.getFiles();

  while (files.hasNext()) {
    let oldFile = files.next();
    let fileName = oldFile.getName();

    if (fileName.includes(tenFile)) {
      try {
        oldFile.setTrashed(true);
      } catch (e) {
        if (!fileName.includes("[CŨ CẦN XÓA]")) {
          let randomNum = Math.floor(Math.random() * 1000);
          oldFile.setName(`[CŨ CẦN XÓA - ${randomNum}] ` + fileName);
        }
      }
    }
  }
}

// ======================================================================
// --- THÊM MỚI: HÀM TẠO NHẬT KÝ (LOGS) HỆ THỐNG ---
// ======================================================================
function ghiLogHeThong(tenKhachHang, thongBao, chiTietLoi) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const LOG_SHEET_NAME = "Logs_Hệ_Thống";
  let logSheet = ss.getSheetByName(LOG_SHEET_NAME);

  if (!logSheet) {
    logSheet = ss.insertSheet(LOG_SHEET_NAME);
    logSheet.appendRow(["Thời gian", "Người chạy", "Khách hàng", "Trạng thái / Thông báo", "Chi tiết Code (Stack Trace)"]);
    logSheet.getRange("A1:E1").setFontWeight("bold").setBackground("#d9ead3");
    logSheet.setColumnWidth(1, 150);
    logSheet.setColumnWidth(3, 150);
    logSheet.setColumnWidth(4, 300);
    logSheet.setColumnWidth(5, 400);
  }

  let thoiGian = new Date();
  let nguoiChay = Session.getActiveUser().getEmail();
  if (!nguoiChay) nguoiChay = "Người dùng ẩn danh";

  logSheet.appendRow([
    thoiGian,
    nguoiChay,
    tenKhachHang,
    thongBao,
    chiTietLoi || ""
  ]);
}

// --- ĐÃ DỌN DẸP: XÓA HÀM onOpen BỊ LẶP, CHỈ GIỮ 1 BẢN CHUẨN ---
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('👉 CÔNG CỤ QR')
      .addItem('⚡ Tạo hồ sơ cho dòng ĐANG CHỌN', 'taoHoSoDongDangChon')
      .addSeparator()
      .addItem('🔄 TẠO LẠI (REGEN) TOÀN BỘ', 'regenToanBoHoSo')
      .addToUi();
}