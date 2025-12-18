# FolderStatsTool - Công cụ thống kê thư mục

<u>Version:</u> 1.0.1  
<u>Author:</u> ThanhRòm  
<u>Release Date:</u> 2025-12-18

## 📋**GIỚI THIỆU**

FolderStatsTool là ứng dụng desktop (GUI) giúp thống kê nhanh dung lượng, số lượng file của các thư mục và hỗ trợ đổi tên hàng loạt file/folder thông qua Excel. Sử dụng thuật toán **Bottom-Up Dynamic Programming** để đạt tốc độ quét cực nhanh.

## ✨**TÍNH NĂNG**

### <u>Tính năng chính:</u>

- 🚀 **Quét siêu nhanh** - Thuật toán Bottom-Up, tốc độ tăng gấp 100 lần
- 🌳 **TreeView trực quan** - Hiển thị cấu trúc cây thư mục
- ✏️ **Inline Editing** - Double-click để sửa tên trực tiếp trong bảng
- 📥 **Import Excel** - Nhập danh sách đổi tên từ file Excel
- 📤 **Export Excel** - Xuất dữ liệu với mẫu điền sẵn
- ⚡ **Batch Rename** - Đổi tên hàng loạt file/folder
- 📊 **Thống kê chi tiết** - Dung lượng, số file, loại file

## 📚**HƯỚNG DẪN**

### <u>Yêu cầu hệ thống</u>

- **Python:** 3.7 trở lên
- **Hệ điều hành:** Windows, macOS, Linux

### <u>Thư viện sử dụng</u>

- **tkinter:** GUI framework (có sẵn trong Python)
- **pandas:** Xử lý dữ liệu dạng bảng
- **openpyxl:** Đọc/ghi file Excel (.xlsx)

### <u>Cài đặt</u>

```bash
pip install pandas openpyxl
```

> **Lưu ý:** Ứng dụng sẽ tự động cài đặt thư viện nếu thiếu khi chạy lần đầu.

### <u>Chạy ứng dụng</u>

```bash
python main.py
```

### <u>Sử dụng</u>

1. Chọn thư mục cần quét bằng nút "Chọn Thư mục..."
2. Bật/tắt "Quét sâu" tùy nhu cầu
3. Nhấn "BẮT ĐẦU QUÉT"
4. Xem kết quả trong TreeView
5. Đổi tên: Double-click cột "Tên mới" hoặc dùng Excel

## **KIẾN TRÚC**

```
FolderStatsTool/
├── main.py          # File chính chứa toàn bộ logic
└── README.md        # Tài liệu này
```

### <u>Cấu trúc code</u>

| Phần | Mô tả |
|------|-------|
| Phần 0 | Import thư viện & Bootstrap |
| Phần 0.2 | Cấu hình thông báo khởi động |
| Phần 1 | Utils (format_size, format_duration) |
| Phần 2 | GUI Class (FolderStatsTool) |
| Phần 3 | Core Logic (quét thư mục) |
| Phần 4 | Inline Editing |
| Phần 5 | Export/Import Excel |
| Phần 6 | Copy utilities & Filter |
| Phần 7 | File Operations (Xóa, Tạo mới, Di chuyển) |
| Phần 8 | Khởi chạy ứng dụng |

## **LỊCH SỬ PHÁT TRIỂN**

### **Version 1.0.1** - Update by ThanhRòm (2025-12-18)

<details>
  <summary>MOVE, DELETE & CREATE</summary>

1. Di chuyển hàng loạt file/folder:
   - Hỗ trợ browse và nhập đường dẫn trực tiếp
   - Tự động tạo thư mục mới nếu chưa tồn tại
2. Xóa hàng loạt với xác nhận
3. Tạo mới file/folder ngay trong ứng dụng
</details>

---
<details>
  <summary>MULTI-COLUMN FILTER</summary>

1. Lọc song song nhiều cột đồng thời
2. Header hiển thị biểu tượng 🔍 khi có filter
3. Lọc theo tên, loại, dung lượng, số file
4. Hỗ trợ lọc = 0 (file rỗng, folder trống)
</details>

---
<details>
  <summary>UI IMPROVEMENTS</summary>

1. Cửa sổ tự động resize theo màn hình
2. Cột Tên và Di chuyển đến co dãn thông minh
3. TreeView mặc định mở rộng
4. Popup căn giữa ứng dụng
</details>

---

### **Version 1.0.0** - Initial Release by ThanhRòm (2025-12-16)

<details>
  <summary>CORE SCRIPT</summary>

- Loại: Python Script chạy ẩn
- Chức năng: Quét đệ quy, tính toán kích thước và xuất Excel
- Hạn chế: Tốc độ chậm với thư mục lớn
</details>

---
<details>
  <summary>HIGH PERFORMANCE & ASYNC</summary>

1. Thuật toán "Bottom-Up Dynamic Programming":
   - Sử dụng os.walk(topdown=False) thay vì quét đi quét lại
   - Tính dung lượng thư mục con sâu nhất trước, cộng dồn lên cha
   - Tốc độ tăng gấp hàng trăm lần với thư mục sâu

2. Giao diện phản hồi (Responsive GUI):
   - Tách biệt luồng xử lý khỏi luồng giao diện
   - Thanh tiến trình và Log thời gian thực
   - TreeView hiển thị trực quan

3. Tương tác dữ liệu:
   - Copy vùng chọn (Ctrl+C)
   - Dán đường dẫn trực tiếp

4. Báo cáo Excel thông minh:
   - Tự động định dạng độ rộng cột
</details>

---
<details>
  <summary>BATCH RENAME</summary>

1. Đổi tên file/folder hàng loạt
2. Import/Export Excel để đổi tên
3. Inline editing trong TreeView
</details>

## **BẢN QUYỀN**

© 2025 ThanhRòm.

Phần mềm được phát triển cho mục đích học tập và hỗ trợ công việc cá nhân.

## **LIÊN HỆ**

*Made with ❤️ by ThanhRòm*
- GitHub: [thanhlehoang0107](https://github.com/thanhlehoang0107)
- Email: thanh.lehoang0107@gmail.com
