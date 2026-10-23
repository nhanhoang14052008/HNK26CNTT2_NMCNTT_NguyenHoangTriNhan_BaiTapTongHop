
# TÀI LIỆU ÔN TẬP KIẾN THỨC

**Terminal • Markdown • Microsoft Word • Excel • PowerPoint**

**Tác giả: Nguyễn Hoàng Trí Nhân**
**Lớp: CNTT2**

## MỤC TIÊU

Ôn tập các thao tác tin học văn phòng cơ bản, bao gồm quản lý tệp bằng Terminal, định dạng Markdown, soạn thảo Word, sử dụng VLOOKUP trong Excel và thiết kế PowerPoint.

## 1. TERMINAL

Terminal là công cụ nhập lệnh để quản lý tệp và thư mục.

Lệnh PowerShell

Chức năng

`pwd`

Xem thư mục hiện tại

`ls`

Liệt kê tệp và thư mục

`cd TenThuMuc`

Chuyển thư mục

`cd ..`

Quay lại thư mục cha

`mkdir docs`

Tạo thư mục

`New-Item note.txt -ItemType File`

Tạo tệp

`Copy-Item a.txt b.txt`

Sao chép tệp

`Move-Item a.txt docs\`

Di chuyển tệp

`Rename-Item a.txt b.txt`

Đổi tên tệp

`Remove-Item a.txt`

Xóa tệp

`Get-Content a.txt`

Đọc nội dung tệp

`Clear-Host`

Xóa màn hình

**Lệnh nâng cao:**

-   `Get-ChildItem -Recurse`: Liệt kê tệp trong các thư mục con.
-   `Get-Process`: Xem tiến trình đang chạy.
-   `Get-Help`: Xem hướng dẫn lệnh.
-   `Select-String`: Tìm nội dung trong tệp.
-   `|`: Chuyển kết quả giữa các lệnh.
-   `>`: Ghi đè tệp; `>>`: Ghi nối thêm.

**Lưu ý:** Kiểm tra đường dẫn trước khi xóa hoặc ghi đè dữ liệu.

## 2. MARKDOWN

Markdown là cú pháp đơn giản để định dạng văn bản.

Mục đích

Cú pháp

Tiêu đề

`# Tiêu đề`

In đậm

`**nội dung**`

In nghiêng

`*nội dung*`

Gạch ngang chữ

`~~nội dung~~`

Danh sách

`- Mục 1`

Danh sách số

`1. Bước một`

Liên kết

`[Tên](https://example.com)`

Mã trong dòng

`` `code` ``

Trích dẫn

`> Nội dung`

Đường kẻ ngang

`---`

Bảng

`| Tên | Điểm |`

Hộp kiểm

`- [ ] Công việc`

**Mẹo:** Thêm khoảng trắng sau dấu `#` và sử dụng tiêu đề theo thứ tự hợp lý.

## 3. MICROSOFT WORD

### Các thao tác cơ bản

-   **Tạo tiêu đề:** Home → Styles → Heading 1, Heading 2.
-   **Đánh số trang:** Insert → Page Number.
-   **Tạo mục lục:** References → Table of Contents.
-   **Tạo danh sách:** Home → Bullets hoặc Numbering.
-   **Ngắt trang:** `Ctrl + Enter`.
-   **Lưu tài liệu:** `Ctrl + S`.
-   **In tài liệu:** `Ctrl + P`.

### Phím tắt căn lề

-   `Ctrl + L`: Căn trái.
-   `Ctrl + E`: Căn giữa.
-   `Ctrl + R`: Căn phải.
-   `Ctrl + J`: Căn đều hai bên.

**Lưu ý:** Sử dụng Heading nhất quán để tạo mục lục tự động và kiểm tra lỗi chính tả trước khi nộp.

## 4. EXCEL – HÀM VLOOKUP

VLOOKUP dùng để tìm một giá trị trong cột đầu tiên của bảng và lấy dữ liệu từ cột khác trên cùng hàng.

**Cú pháp:**

`=VLOOKUP(lookup_value, table_array, col_index_num, FALSE)`

Trong đó:

-   `lookup_value`: Giá trị cần tìm.
-   `table_array`: Vùng dữ liệu tra cứu.
-   `col_index_num`: Số thứ tự cột lấy kết quả.
-   `FALSE`: Tìm kiếm chính xác.

**Ví dụ:** Bảng dữ liệu nằm trong vùng A2:C5, ô E2 chứa mã sinh viên.

Lấy họ tên:

`=VLOOKUP(E2,$A$2:$C$5,2,FALSE)`

Lấy lớp:

`=VLOOKUP(E2,$A$2:$C$5,3,FALSE)`

Xử lý lỗi:

`=IFERROR(VLOOKUP(E2,$A$2:$C$5,2,FALSE),"Không tìm thấy")`

**Lỗi thường gặp:**

-   `#N/A`: Không tìm thấy giá trị.
-   `#REF!`: Chỉ số cột không hợp lệ.
-   Dùng dấu `$` để cố định vùng dữ liệu khi sao chép công thức.

## 5. MICROSOFT POWERPOINT

### Nguyên tắc thiết kế

-   Mỗi slide tập trung vào một ý chính.
-   Nội dung ngắn gọn, dễ hiểu.
-   Sử dụng màu sắc và phông chữ thống nhất.
-   Đảm bảo chữ dễ đọc, hình ảnh rõ nét.
-   Kiểm tra chính tả, số liệu và nguồn hình ảnh.

### Quy tắc 10–20–30

-   Khoảng 10 slide.
-   Thời lượng khoảng 20 phút.
-   Cỡ chữ tối thiểu khoảng 30 pt.

Đây là quy tắc tham khảo, có thể thay đổi tùy bài thuyết trình.

### Quy trình thực hiện

1.  Xác định mục tiêu và người nghe.
2.  Lập dàn ý bài trình bày.
3.  Chọn mẫu thiết kế phù hợp.
4.  Trình bày nội dung ngắn gọn.
5.  Kiểm tra và luyện tập trước khi thuyết trình.

### Phím tắt hữu ích

-   `F5`: Trình chiếu từ đầu.
-   `Shift + F5`: Trình chiếu từ slide hiện tại.
-   `Ctrl + M`: Thêm slide mới.
-   `Ctrl + D`: Nhân đôi đối tượng hoặc slide đang chọn trong các ngữ cảnh hỗ trợ.
-   `Ctrl + S`: Lưu bài trình chiếu
