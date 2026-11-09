# UyOffice Info

Công cụ web của **Uy** giúp sửa thông tin **File → Info** của file Word, PowerPoint, Excel (`.docx`, `.pptx`, `.xlsx`, kèm `.docm`, `.dotx`, `.pptm`, `.ppsx`, `.xlsm`...).

**File được xử lý 100% trên trình duyệt**, không tải lên server nào.

## Sửa được những gì

| Nhóm | Thông tin |
|---|---|
| Thuộc tính | Tiêu đề, Chủ đề, Thẻ (Tags), Thể loại, Trạng thái, Chú thích, Công ty, Hyperlink base, Mẫu (Template) |
| Người liên quan | Tác giả, Người sửa cuối, Quản lý |
| Ngày tháng | Ngày tạo, Sửa lần cuối, In lần cuối, Tổng thời gian chỉnh sửa |
| Thống kê | Số trang/slide/từ/ký tự/dòng/đoạn... |
| Nâng cao | Số lần sửa (Revision), Ứng dụng, Phiên bản, Ngôn ngữ... |
| Tùy chỉnh | Thêm / sửa / xóa Custom properties (văn bản, số, ngày, có/không) |

Ngoài ra còn có:

- Mở nhiều file cùng lúc, nút **Áp dụng cho tất cả file**, tải tất cả về dưới dạng `.zip`
- Nút **Xóa thông tin cá nhân**: xóa Tác giả, Người sửa cuối, Công ty, Quản lý, Ngày in

Tool chỉ ghi lại những trường bạn đã sửa, nội dung tài liệu được giữ nguyên.

> Lưu ý: nếu sau đó bạn mở file bằng Office rồi **lưu lại**, Office sẽ tự cập nhật *Sửa lần cuối*, *Người sửa cuối*, số trang/từ...

## Chạy trên máy

```bash
npm install
npm run dev
```

Build: `npm run build` (output nằm trong `dist/`). Dự án dùng Vite + [JSZip](https://stuk.github.io/jszip/).
