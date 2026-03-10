function taoHoSoDongDangChon() {
  // --- CẤU HÌNH (Đã giữ nguyên ID của bạn) ---
  const TEMPLATE_ID = '1dM7Yj-YdBkev99LlAo4opwELQ4fF3cEisId8mJEUE5E'; // ID file mẫu Docs
  const DATA_FOLDER_ID = '1Mcg3wG8HtUwSvGxFD0Pfktqlj4qUJkDc'; // Thư mục tổng chứa các thư mục con
  // ----------------

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet(); // Lấy sheet đang mở (tránh lỗi sai tên sheet)

  // 1. XÁC ĐỊNH DÒNG ĐANG CHỌN
  const activeRange = sheet.getActiveRange();
  const rowIndex = activeRange.getRow(); // Lấy số thứ tự dòng

  // Chặn nếu chọn vào dòng tiêu đề (dòng 1)
  if (rowIndex === 1) {
    ss.toast("Vui lòng chọn dòng khách hàng (không chọn tiêu đề).", "⚠️ Chọn sai dòng", 3);
    return;
  }

  // Lấy dữ liệu của dòng đó (Cột A đến G -> 7 cột)
  // rowData sẽ là mảng chứa dữ liệu của dòng đang chọn
  const rowData = sheet.getRange(rowIndex, 4, 1, 6).getValues()[0];

  let tenKH = rowData[0];     // Cột A
  let loaiHS = rowData[1];    // Cột B
  let diaChi = rowData[2];    // Cột C
  let linkDrive = rowData[3]; // Cột D
  //let linkMaps = rowData[7];  // Cột E
  let status = rowData[4];    // Cột F

  // Kiểm tra tên
  if (!tenKH) {
    ss.toast("Dòng này không có tên khách hàng!", "⚠️ Dữ liệu trống", 3);
    return;
  }

  ss.toast("Đang xử lý hồ sơ: " + tenKH + "...", "⏳ Đang chạy", -1);

  // Thư mục gốc chứa toàn bộ data
  const rootDataFolder = DriveApp.getFolderById(DATA_FOLDER_ID);

  try {
    let targetFolder; // Biến lưu thư mục đích

    // --- BƯỚC 1: XÁC ĐỊNH THƯ MỤC LƯU TRỮ ---
    if (!linkDrive || linkDrive === "") {
      // TRƯỜNG HỢP A: Chưa có folder -> Tự tạo mới
      let folderName = tenKH + "-" + diaChi + " - Hồ sơ";
      targetFolder = rootDataFolder.createFolder(folderName);

      // Chia sẻ công khai
      targetFolder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

      // Cập nhật link vào Sheets (Cột D - index 4)
      linkDrive = targetFolder.getUrl();
      sheet.getRange(rowIndex, 7).setValue(linkDrive);
      SpreadsheetApp.flush(); // Lưu ngay lập tức

    } else {
      // TRƯỜNG HỢP B: Đã có link -> Lấy ID
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

    // Tạo bản sao Doc ngay trong thư mục của khách
    let tempFile = templateFile.makeCopy(tenFileMoi, targetFolder);
    let tempDoc = DocumentApp.openById(tempFile.getId());
    let body = tempDoc.getBody();

    // Điền thông tin chữ (Giữ nguyên các placeholder như code cũ của bạn)
    body.replaceText("{{TEN_KHACH_HANG}}", tenKH);
    body.replaceText("{{LOAI_HO_SO}}", loaiHS);
    body.replaceText("{{DIA_CHI}}", diaChi);
    body.replaceText("{{LINK_DRIVE}}", linkDrive);
    //body.replaceText("{{LINK_MAPS}}", linkMaps);

    // Tạo QR Code
    xuLyChenQR(body, "QR_BAN_VE", linkDrive);
    //xuLyChenQR(body, "QR_VI_TRI", linkMaps);

    tempDoc.saveAndClose();

    // --- BƯỚC 3: DỌN RÁC VÀ XUẤT PDF ---
      // Dọn dẹp các file PDF cũ bị trùng tên trong folder này
      dondepFileCu(targetFolder, tenFileMoi);

      // Chuyển đổi Doc sang PDF
      let pdfBlob = tempFile.getAs(MimeType.PDF);

      // Tạo file PDF mới (tôi thêm hẳn đuôi .pdf vào tên file để Google Drive quản lý chuẩn hơn)
      let pdfFile = targetFolder.createFile(pdfBlob).setName(tenFileMoi + ".pdf");

      // Xóa file Doc nháp
      tempFile.setTrashed(true);

    // --- BƯỚC 4: CẬP NHẬT TRẠNG THÁI ---
    // Cập nhật đúng dòng đang chọn (rowIndex)
    sheet.getRange(rowIndex, 8).setValue("Đã tạo");
    sheet.getRange(rowIndex, 9).setValue(pdfFile.getUrl());

    ss.toast("Đã xong hồ sơ cho: " + tenKH, "✅ Hoàn tất", 5);

  } catch (e) {
    Logger.log("Lỗi: " + e.toString());
    sheet.getRange(rowIndex, 8).setValue("Lỗi: " + e.message);
    ss.toast("Gặp lỗi: " + e.message, "❌ Thất bại", 5);
  }
}

// --- CÁC HÀM PHỤ TRỢ (GIỮ NGUYÊN) ---

function xuLyChenQR(body, altText, linkRaw) {
  try {
    let cleanLink = chuanHoaLinkHienThi(linkRaw);
    if (!cleanLink) {
      Logger.log("Link trống cho: " + altText);
      return;
    }

    // Đã đổi sang API của QRServer (hoạt động cực kỳ ổn định, dự phòng khi QuickChart lỗi)
    const qrUrl = `https://api.qrserver.com/v1/create-qr-code/?size=300x300&data=${encodeURIComponent(cleanLink)}&margin=2`;

    // muteHttpExceptions giúp code không bị sập cứng nếu API bên thứ 3 bảo trì
    const response = UrlFetchApp.fetch(qrUrl, {muteHttpExceptions: true});
    if (response.getResponseCode() !== 200) {
      Logger.log("Lỗi Server QR: " + response.getContentText());
      return;
    }

    const blob = response.getBlob();
    const images = body.getImages();

    for (let img of images) {
      if (img.getAltDescription() === altText || img.getAltTitle() === altText) {
        let parent = img.getParent();
        let index = parent.getChildIndex(img);
        let newImg = parent.asParagraph().insertInlineImage(index, blob);
        newImg.setWidth(img.getWidth()).setHeight(img.getHeight());
        img.removeFromParent();
        return;
      }
    }
  } catch (err) { Logger.log("Lỗi chèn QR " + altText + ": " + err); }
}

// Hàm trích xuất ID siêu mạnh: Tự động quét và tìm đúng chuỗi 25-35 ký tự của Drive ID
function trichXuatFolderId(link) {
  if (!link) return null;
  let strLink = String(link).trim();

  // Regex mới: Bắt chính xác ID Drive bất chấp mọi cấu trúc link thừa thãi
  let match = strLink.match(/[-\w]{25,}/);
  if (match) return match[0];
  return null;
}

function chuanHoaLinkHienThi(link) {
  let id = trichXuatFolderId(link);
  // Chỉ ráp thành link folder chuẩn nếu thực sự tìm thấy ID
  if (id) {
    return "https://drive.google.com/drive/folders/" + id;
  }
  return String(link).trim();
}

function regenToanBoHoSo() {
  let ui = SpreadsheetApp.getUi(); // Khởi tạo giao diện UI trước

  // --- 🔒 CHỐT CHẶN BẢO MẬT ADMIN ---
  let passPrompt = ui.prompt(
    '🔒 Yêu cầu quyền Admin',
    'Tính năng này chỉ dành cho Admin.\nVui lòng nhập mã PIN để tiếp tục:',
    ui.ButtonSet.OK_CANCEL
  );

  // Nếu người dùng bấm Cancel hoặc nút X
  if (passPrompt.getSelectedButton() !== ui.Button.OK) {
    return;
  }

  // Kiểm tra mật khẩu (Bạn hãy tự đổi "686868" thành số bạn muốn)
  if (passPrompt.getResponseText() !== "686868") {
    ui.alert("❌ Sai mã PIN!", "Bạn không có quyền sử dụng tính năng này.", ui.ButtonSet.OK);
    return;
  }
  // ----------------------------------

  // --- CẤU HÌNH (Đã đồng bộ ID với hàm taoHoSoDongDangChon) ---
  const TEMPLATE_ID = '1dM7Yj-YdBkev99LlAo4opwELQ4fF3cEisId8mJEUE5E';
  const DATA_FOLDER_ID = '1Mcg3wG8HtUwSvGxFD0Pfktqlj4qUJkDc';
  // ----------------

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const data = sheet.getDataRange().getValues();

  // --- Xác nhận trước khi chạy (Đoạn code cũ của bạn tiếp tục ở đây) ---
  let response = ui.alert('Xác nhận Tạo lại', 'Bạn có chắc chắn muốn TẠO LẠI (Regen) toàn bộ hồ sơ trong Sheet này không? Quá trình này sẽ chạy qua tất cả các dòng có tên khách hàng.', ui.ButtonSet.YES_NO);

  if (response != ui.Button.YES) return;

  // --- BƯỚC MỚI: ĐẾM TỔNG SỐ HỒ SƠ CẦN CHẠY ---
  let totalValid = 0;
  for (let i = 1; i < data.length; i++) {
    if (data[i][3]) totalValid++; // data[i][3] là Cột D (Tên KH)
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

  // Chạy vòng lặp
  for (let i = 1; i < data.length; i++) {
    let tenKH = data[i][3];     // Cột D
    let loaiHS = data[i][4];    // Cột E
    let diaChi = data[i][5];    // Cột F
    let linkDrive = data[i][6]; // Cột G

    // Bỏ qua nếu dòng này không có tên
    if (!tenKH) continue;

    count++; // Tăng biến đếm

    // --- HIỂN THỊ TIẾN TRÌNH TRƯỚC KHI LÀM ---
    // Hiện thông báo góc màn hình (hiện 4 giây rồi tắt để nhường chỗ cho thông báo tiếp theo)
    ss.toast(`Đang xử lý: ${tenKH} (${count}/${totalValid})`, "🔄 Tiến trình", 4);

    // In trạng thái đang chạy ra màn hình Excel
    sheet.getRange(i + 1, 8).setValue("⏳ Đang xử lý...");
    SpreadsheetApp.flush(); // ÉP GOOGLE SHEETS CẬP NHẬT GIAO DIỆN NGAY LẬP TỨC

    try {
      let targetFolder;

      // XÁC ĐỊNH FOLDER
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

      // TẠO FILE DOCS
      let tenFileMoi = `${loaiHS} - ${tenKH} - QR`;
      let templateFile = DriveApp.getFileById(TEMPLATE_ID);
      let tempFile = templateFile.makeCopy(tenFileMoi, targetFolder);
      let tempDoc = DocumentApp.openById(tempFile.getId());
      let body = tempDoc.getBody();

      body.replaceText("{{TEN_KHACH_HANG}}", tenKH);
      body.replaceText("{{LOAI_HO_SO}}", loaiHS);
      body.replaceText("{{DIA_CHI}}", diaChi);
      body.replaceText("{{LINK_DRIVE}}", linkDrive);

      // Tạo QR
      xuLyChenQR(body, "QR_BAN_VE", linkDrive);

      tempDoc.saveAndClose();

      // --- DỌN RÁC VÀ XUẤT PDF ---
      // Dọn dẹp các file PDF cũ bị trùng tên trong folder này
      dondepFileCu(targetFolder, tenFileMoi);
      // Chuyển đổi Doc sang PDF
      let pdfBlob = tempFile.getAs(MimeType.PDF);
      // Tạo file PDF mới (tôi thêm hẳn đuôi .pdf vào tên file để Google Drive quản lý chuẩn hơn)
      let pdfFile = targetFolder.createFile(pdfBlob).setName(tenFileMoi + ".pdf");
      // Xóa file Doc nháp
      tempFile.setTrashed(true);

      // --- CẬP NHẬT TRẠNG THÁI HOÀN THÀNH ---
      sheet.getRange(i + 1, 8).setValue("✅ Đã Regen");
      sheet.getRange(i + 1, 9).setValue(pdfFile.getUrl());
      SpreadsheetApp.flush(); // Cập nhật màn hình ngay lập tức để người dùng thấy tích xanh

    } catch (e) {
      Logger.log("Lỗi dòng " + (i+1) + ": " + e.toString());
      sheet.getRange(i + 1, 8).setValue("❌ Lỗi: " + e.message);
      SpreadsheetApp.flush(); // Hiển thị lỗi ngay lập tức
    }
  }

  ss.toast("Đã hoàn tất " + count + " hồ sơ!", "🎉 Xong toàn bộ", 10);
}

// Hàm tìm, xóa file cũ và dọn cả rác do người khác để lại (Đã fix lỗi nhận diện tên)
function dondepFileCu(folder, tenFile) {
  // Lấy TOÀN BỘ file đang có trong thư mục của khách hàng này
  let files = folder.getFiles();

  while (files.hasNext()) {
    let oldFile = files.next();
    let fileName = oldFile.getName();

    // Kiểm tra xem tên file CÓ CHỨA chuỗi tên gốc (tenFile) không?
    // Bằng cách này, dù file tên là "[CŨ CẦN XÓA - 123] Hồ sơ.pdf" thì vẫn bị nhận diện
    if (fileName.includes(tenFile)) {
      try {
        // Cố gắng đưa vào thùng rác (Chủ sở hữu chạy sẽ thành công)
        oldFile.setTrashed(true);
      } catch (e) {
        // Nếu bị lỗi quyền (do nhân viên chạy), chỉ đổi tên nếu file đó CHƯA bị đổi tên
        if (!fileName.includes("[CŨ CẦN XÓA]")) {
          let randomNum = Math.floor(Math.random() * 1000);
          oldFile.setName(`[CŨ CẦN XÓA - ${randomNum}] ` + fileName);
        }
      }
    }
  }
}

// Hàm cập nhật menu
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('👉 CÔNG CỤ QR')
      .addItem('⚡ Tạo hồ sơ cho dòng ĐANG CHỌN', 'taoHoSoDongDangChon')
      .addToUi();
}

// Hàm cập nhật menu
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('👉 CÔNG CỤ QR')
      .addItem('⚡ Tạo hồ sơ cho dòng ĐANG CHỌN', 'taoHoSoDongDangChon')
      .addSeparator() // Đường kẻ ngang phân cách
      .addItem('🔄 TẠO LẠI (REGEN) TOÀN BỘ', 'regenToanBoHoSo')
      .addToUi();
}