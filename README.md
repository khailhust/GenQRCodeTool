# 📱 GenQRCode - Hệ Thống Tạo Hồ Sơ QR Tự Động

Dự án Google Apps Script tự động tạo hồ sơ khách hàng với mã QR liên kết tới Google Drive.

## 🎯 Mục Đích

Tự động hóa quy trình tạo hồ sơ khách hàng bao gồm:
- ✅ Tạo thư mục trên Google Drive
- ✅ Sao chép mẫu tài liệu Google Docs
- ✅ Điền thông tin khách hàng vào mẫu
- ✅ Tạo mã QR liên kết tới thư mục Drive
- ✅ Xuất PDF
- ✅ Ghi log hệ thống chi tiết

## ⚙️ Cấu Hình Ban Đầu

Trước khi sử dụng, cập nhật các ID sau trong `Code.gs`:

```javascript
const TEMPLATE_ID = 'YOUR_TEMPLATE_DOC_ID';      // ID của mẫu Google Docs
const DATA_FOLDER_ID = 'YOUR_DATA_FOLDER_ID';    // ID thư mục chứa hồ sơ
```

### Cách Lấy ID:
- **Google Docs Template**: Mở file → URL chứa: `https://docs.google.com/document/d/{TEMPLATE_ID}/edit`
- **Google Drive Folder**: Mở thư mục → URL chứa: `https://drive.google.com/drive/folders/{FOLDER_ID}`

## 📋 Cấu Trúc Spreadsheet

Spreadsheet phải có các cột sau (từ cột A):

| Cột | Tên | Mô Tả |
|-----|-----|-------|
| A | (Tùy chọn) | - |
| B | (Tùy chọn) | - |
| C | (Tùy chọn) | - |
| D | Tên Khách Hàng | Bắt buộc |
| E | Loại Hồ Sơ | Vd: "Hồ sơ A", "Hồ sơ B" |
| F | Địa Chỉ | Địa chỉ khách hàng |
| G | Link Drive | Sẽ được tự động điền nếu trống |
| H | Trạng Thái | Hiển thị: ✅ Đã tạo / ❌ Lỗi |
| I | Link PDF | URL của PDF được tạo |

## 🚀 Cách Sử Dụng

### 1. Tạo Hồ Sơ Cho Một Dòng
1. Nhấp vào cell bất kỳ trong dòng khách hàng muốn tạo
2. Vào menu: **👉 CÔNG CỤ QR** → **⚡ Tạo hồ sơ cho dòng ĐANG CHỌN**
3. Hệ thống sẽ:
   - Tạo thư mục trên Drive nếu chưa có
   - Sao chép mẫu Docs
   - Điền dữ liệu vào mẫu
   - Tạo mã QR
   - Xuất PDF
   - Ghi log kết quả


### 2. Tạo Lại Toàn Bộ Hồ Sơ
1. Vào menu: **👉 CÔNG CỤ QR** → **🔄 TẠO LẠI (REGEN) TOÀN BỘ**
2. Nhập PIN Admin
3. Xác nhận thao tác
4. Hệ thống sẽ xử lý tất cả dòng có tên khách hàng

**⚠️ Lưu ý**: Tính năng này yêu cầu PIN bảo mật để tránh xoá dữ liệu nhầm

## 📖 Tài Liệu Hàm

### `taoHoSoDongDangChon()`
Tạo hồ sơ cho dòng được chọn hiện tại.

**Quy trình:**
1. Kiểm tra dòng không phải tiêu đề
2. Lấy dữ liệu từ cột D-H
3. Tạo/xác định thư mục trên Drive
4. Sao chép và chỉnh sửa mẫu
5. Tạo QR code
6. Xuất PDF
7. Cập nhật trạng thái

### `xuLyChenQR(body, altText, linkRaw)`
Tạo mã QR và chèn vào tài liệu Google Docs.

**Tham số:**
- `body`: Body của DocumentApp
- `altText`: Tên Alt Text của ảnh placeholder (VD: "QR_BAN_VE")
- `linkRaw`: URL thư mục Drive cần tạo QR

**Server QR (3 lựa chọn):**
1. QuickChart.io
2. QR Server API
3. Google Chart API

Hệ thống sẽ thử lần lượt tới khi thành công.

**Return:**
- Blob của ảnh QR
- Báo cáo chi tiết server nào được dùng

### `regenToanBoHoSo()`
Tạo lại toàn bộ hồ sơ trong sheet (yêu cầu PIN).

**Quy trình:**
1. Xác thực PIN
2. Đếm tổng dòng khách hàng
3. Lặp qua từng dòng và gọi `taoHoSoDongDangChon()`
4. Ghi log tiến trình

### `dondepFileCu(folder, tenFile)`
Xoá các file cũ cùng tên trong thư mục.

**Cách xử lý:**
- Nếu xoá được: Đưa vào thùng rác
- Nếu ko xoá được: Đổi tên có prefix "[CŨ CẦN XÓA]"

### `ghiLogHeThong(tenKhachHang, thongBao, chiTietLoi)`
Ghi log hệ thống vào sheet "Logs_Hệ_Thống".

**Thông tin ghi:**
- Thời gian chạy
- Email người chạy
- Tên khách hàng
- Trạng thái / Thông báo
- Chi tiết lỗi (nếu có)

### `trichXuatFolderId(link)`
Trích xuất ID thư mục từ Google Drive link.

**Input:** URL hoặc ID
**Output:** Folder ID (25+ ký tự)

### `chuanHoaLinkHienThi(link)`
Chuẩn hóa link Drive sang định dạng standard.

**Input:** URL hoặc ID
**Output:** `https://drive.google.com/drive/folders/{ID}`

### `onOpen()`
Gọi khi spreadsheet mở - tạo menu tùy chỉnh.

## 🛠️ Yêu Cầu

### Google Services
- ✅ Google Sheets API
- ✅ Google Drive API
- ✅ Google Docs API
- ✅ URL Fetch API

### Quyền Hạn Cần
- Truy cập Spreadsheet hiện tại
- Truy cập Google Drive
- Truy cập Google Docs
- Kết nối Internet (tạo QR code)

## ⚠️ Xử Lý Lỗi

| Lỗi | Nguyên Nhân | Giải Pháp |
|-----|-----------|----------|
| "Dòng này không có tên khách hàng!" | Cột D trống | Nhập tên khách hàng vào cột D |
| "Link Drive không hợp lệ" | Folder ID sai | Kiểm tra link Drive, xoá cột G để tạo folder mới |
| "Cả 3 server QR đều thất bại" | Không có kết nối mạng | Kiểm tra kết nối Internet, thử lại |
| "Template bị lỗi: Không tìm thấy ảnh" | Alt Text sai | Kiểm tra mẫu có ảnh placeholder với Alt Text "QR_BAN_VE" |

**Mọi lỗi sẽ tự động ghi vào sheet "Logs_Hệ_Thống"**

## 📊 Sheet Logs_Hệ_Thống

Hệ thống tự động tạo sheet log với cấu trúc:

| Thời gian | Người chạy | Khách hàng | Trạng thái | Chi tiết Code |
|-----------|-----------|-----------|-----------|---------------|
| ... | email@gmail.com | Nguyễn Văn A | ✅ Tạo mới thành công | Link PDF + Báo cáo QR |

## 🔒 Bảo Mật

- PIN Admin: `686868` (dùng cho REGEN toàn bộ)
- ⚠️ **Khuyến cáo**: Thay đổi PIN tại hàm `regenToanBoHoSo()` trước khi dùng sản xuất

## 📝 Mẫu Google Docs

Mẫu phải chứa các **Replacement Text** tại vị trí cần:

```
{{TEN_KHACH_HANG}}    → Tên khách hàng
{{LOAI_HO_SO}}        → Loại hồ sơ
{{DIA_CHI}}           → Địa chỉ khách hàng
{{LINK_DRIVE}}        → Link Drive thư mục hồ sơ
```

Mẫu cũng phải có **ảnh placeholder** với Alt Text: `QR_BAN_VE` (QR code sẽ thay thế ảnh này)

## 🚨 Troubleshooting

### Script chạy chậm
- Google có rate limit
- Dùng REGEN cho nhiều hồ sơ sẽ chậm hơn (tùy số lượng)
- Kiểm tra tốc độ mạng

### QR code không hiển thị
- Kiểm tra mẫu có Alt Text `QR_BAN_VE`
- Kiểm tra kết nối Internet
- Xem logs sheet để biết server nào lỗi

### Pin Admin sai
- Nhập lại PIN chính xác: `686868`
- Nếu quên, sửa trực tiếp trong code

## 📞 Hỗ Trợ

- Xem chi tiết lỗi trong sheet **"Logs_Hệ_Thống"**
- Kiểm tra Google Drive có quyền truy cập folder
- Kiểm tra template ID và folder ID có chính xác

## 📄 Giấy phép

Dự án nội bộ - Sử dụng cho mục đích quản lý hồ sơ khách hàng.

---

**Phiên bản**: 1.0
**Cập nhật lần cuối**: Tháng 3, 2026
