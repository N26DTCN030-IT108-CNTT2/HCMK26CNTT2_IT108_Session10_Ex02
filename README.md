# HCMK26CNTT2_IT108_Session10_Ex02
Bài tập số 2
# Bộ Hồ Sơ Thiết Kế Dự Án PTIT-Learning (Session 10 / Session 12 Ex02)

Dự án thiết kế bộ tài liệu kỹ thuật SRS, quản lý dữ liệu học viên và trình bày slide giới thiệu cho nền tảng học trực tuyến "PTIT-Learning" của Startup EdTech.

---

# Phần 1. Phân Tích Cấu Trúc Sơ Đồ Cây Heading File SRS

Cấu trúc phân cấp tiêu đề được gán nhãn Heading Styles trong tài liệu `SRS_Lop_Hoc_Ao.docx` giúp kiểm soát tốt giao diện điều hướng (Navigation Pane) và cập nhật mục lục tự động (Auto TOC):

# 1. TỔNG QUAN HỆ THỐNG PTIT-LEARNING (Heading 1)
## 1.1. Bối cảnh dự án (Heading 2)
## 1.2. Mục tiêu hệ thống (Heading 2)

# 2. ĐẶC TẢ TÍNH NĂNG LỚP HỌC ẢO (Heading 1)
## 2.1. Quản lý phòng học trực tuyến (Heading 2)
## 2.2. Điểm danh sinh viên tự động (Heading 2)
## 2.3. Tương tác phát biểu và Giơ tay (Heading 2)
## 2.4. Chia sẻ màn hình và Bảng trắng (Heading 2)
## 2.5. Ghi hình buổi học (Cloud Recording) (Heading 2)

# 3. YÊU CẦU PHI CHỨC NĂNG (Heading 1)
## 3.1. Hiệu năng và Độ trễ truyền tải (Heading 2)
## 3.2. Bảo mật và Phân quyền truy cập (Heading 2)
## 3.3. Khả năng mở rộng hạ tầng (Heading 2)

---

# Phần 2. Mô Tả Triển Khai Bộ Tài Liệu Dự Án

## 1. File Word Đặc tả SRS (`SRS_Lop_Hoc_Ao.docx`)
- **Quản lý Navigation Pane:** Đã gán nhãn chuẩn cho 10 tiêu đề cấp 2 (`Heading 2`) nằm dưới 3 chương chính (`Heading 1`), giúp người quản trị dễ dàng di chuyển nhanh giữa các phần tài liệu.
- **Tích hợp Auto TOC:** Chèn trường mục lục tự động `Table of Contents` tại trang 1. Trước khi lưu và bàn giao, tài liệu đã được bấm `Update entire table` để đảm bảo số trang khớp 100% với nội dung thực tế.

---

## 2. File Excel Quản Lý Học Viên (`Danh_Sach_Sinh_Vien_PTIT.xlsx`)
- **Tuân thủ 5 Quy tắc Vàng Database:**
  - Dòng 1 là **Header Row duy nhất** chứa tên các trường dữ liệu (`MSSV`, `Ho_Ten`, `Khoa`, `Diem_Chuyen_Can`, `Diem_Chuyen_Mon`, `Trang_Thai`).
  - Tuyệt đối không chứa dòng tính tổng hay dòng trống xen kẽ làm đứt gãy vùng dữ liệu.
- **Thao tác AutoFilter đa điều kiện (`Ctrl + Shift + L`):**
  - Kích hoạt bộ lọc `AutoFilter` trên vùng bảng.
  - Lọc cột `Khoa`: Tích chọn duy nhất giá trị `"CNTT"`.
  - Lọc cột `Diem_Chuyen_Can`: Chọn `Number Filters` ➔ `Less Than` ➔ Nhập con số `7`.
  - **Kết quả:** Trích xuất chính xác danh sách các sinh viên thuộc khoa CNTT có điểm chuyên cần dưới 7 để gửi cảnh báo học tập.

---

## 3. File Slide Trình Bày (`Gioi_Thieu_PTIT_Learning.pptx`)
- **Phối màu Combo tối (High Contrast Ratio):**
  - **Nền chính:** Xanh Navy đậm (`#102C57`).
  - **Màu chữ:** Trắng ngà (`#FFFFFF`) đảm bảo độ tương phản cao, dễ đọc khi chiếu.
  - **Điểm nhấn (Accent Color):** Cam tươi (`#FF6B00`) tạo sự chú ý vào các nút tính năng chính và chỉ số tăng trưởng.
- **Bố cục 3 vùng chuẩn:** Vùng Header chứa Assertion Title, Vùng Trung tâm chứa Flowchart lộ trình học tập, Vùng Footer chứa thông tin liên hệ và số trang.

---

# Phần 3. Kiểm Thử & Giải Pháp Xử Lý Khoảng Trắng (Hàm TRIM)

## 1. Mô tả sự cố kiểm thử
Khi thực hiện lọc hoặc tìm kiếm sinh viên theo tên (ví dụ: tìm `"Trần Thị Bình"`), hệ thống Excel không trả về kết quả dù mắt thường vẫn nhìn thấy tên sinh viên đó nằm rõ ràng ở dòng 3.

## 2. Phân tích nguyên nhân kỹ thuật
- **Lỗi khoảng trắng ẩn (Trailing Spaces):** Giá trị trong ô Excel thực tế bị dính khoảng trắng thừa ở cuối chuỗi: `"Trần Thị Bình "` (15 ký tự).
- **Cơ chế so sánh:** Điều kiện tìm kiếm nhập vào là `"Trần Thị Bình"` (14 ký tự). Do mã Unicode của khoảng trắng khác với ký tự rỗng, thuật toán so sánh chuỗi xác định 2 chuỗi này không khớp nhau (`Exact Match = FALSE`).

## 3. Quy trình 3 bước xử lý triệt để bằng hàm TRIM()
1. **Bước 1:** Tạo thêm một cột phụ tên là `Ho_Ten_Chuan` ngay bên cạnh cột họ tên thô.
2. **Bước 2:** Nhập công thức `=TRIM(B2)` vào dòng đầu tiên và nhấp đôi Fill Handle để áp dụng cho toàn bộ danh sách. Hàm `TRIM()` sẽ tự động cắt bỏ sạch sẽ toàn bộ khoảng trắng thừa ở đầu, cuối và khoảng trắng kép giữa các từ.
3. **Bước 3:** Bật lại bộ lọc `AutoFilter` (`Ctrl + Shift + L`) trên cột `Ho_Ten_Chuan` mới làm sạch. Kết quả tìm kiếm và lọc dữ liệu sẽ chính xác 100% mà không bị bỏ sót bản ghi.

---

# Phần 4. Danh Sách Kiểm Tra Hoàn Thiện (Task List)

- [x] Phân tích cây Heading H2 và quy trình gỡ lỗi bằng TRIM trong `README.md`
- [x] Tạo file Word `SRS_Lop_Hoc_Ao.docx` gán chuẩn Navigation Pane 10 mục H2 & Update TOC
- [x] Tạo file Excel `Danh_Sach_Sinh_Vien_PTIT.xlsx` chuẩn Database và Filter (Khoa CNTT & Chuyên cần < 7)
- [x] Thiết kế Slide `Gioi_Thieu_PTIT_Learning.pptx` chuẩn Combo màu Navy/Trắng/Accent cam[cite: 1]
- [ ] Push toàn bộ sản phẩm lên GitHub Repository `HCMK26CNTT2_IT108_Session10_Ex02`[cite: 1, 2]
- [ ] Copy đường dẫn link nộp lên hệ thống LMS Portal[cite: 1]
