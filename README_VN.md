# Office Metadata Timestamp GUI v3

**Office Metadata Timestamp GUI v3** là công cụ Python/Tkinter chạy trên Windows, dùng để sửa lỗi ngày giờ file bị sai trong các kho tài liệu lớn. Công cụ đọc metadata nội bộ của file Office/PDF, lấy ngày tạo thật của tài liệu, sau đó set lại thời gian file trong Windows.

Công cụ được thiết kế cho các thư mục dữ liệu lớn như:

```text
D:\PharmSolu
```

Trường hợp thường gặp là file trong Windows Explorer hiển thị ngày tương lai như `2029`, `2062` hoặc các mốc thời gian không hợp lý.

---

## Tính năng chính

- Paste trực tiếp một hoặc nhiều đường dẫn file/folder vào GUI.
- Thêm file hoặc folder bằng nút chọn file/thư mục.
- Quét đệ quy toàn bộ thư mục lớn.
- Chỉ scan các file có ngày filesystem lớn hơn ngày hiện tại.
- Đọc metadata nội bộ từ Office và PDF.
- Ưu tiên ngày **Content created** của tài liệu.
- Từ chối các metadata có ngày tương lai bất thường.
- Có fallback lấy ngày từ tên file hoặc đường dẫn, ví dụ `17.05.2016`.
- Set lại các timestamp của Windows:
  - `CreationTime`
  - `ModifiedTime / LastWriteTime`
  - `AccessedTime`
- Ghi log CSV để kiểm tra và audit.
- Không cần `pywin32`; việc set timestamp dùng Windows API qua `ctypes`.

---

## Vì sao cần công cụ này?

Một số kho dữ liệu cũ có thể bị sai ngày giờ file do:

- máy tính từng bị sai ngày hệ thống;
- copy/backup/restore dữ liệu;
- đồng bộ qua ổ mạng hoặc ổ ngoài;
- lỗi phần mềm nén/giải nén;
- lỗi chuyển dữ liệu giữa nhiều máy chủ.

Ví dụ:

```text
Windows file timestamp : 2062-05-16 22:57
Office Content created : 2016-05-11 12:43
```

Trong trường hợp này, ngày nên dùng để sửa file là:

```text
2016-05-11 12:43
```

Công cụ sẽ scan metadata trước, hiển thị ngày mục tiêu trong bảng kết quả, sau đó chỉ sửa khi người dùng bấm **Apply timestamps**.

---

## File hỗ trợ

### File Office hiện đại

Các định dạng sau được đọc bằng thư viện chuẩn của Python, không cần cài thêm:

```text
.docx .docm .dotx .dotm
.xlsx .xlsm .xltx .xltm
.pptx .pptm .potx .potm
```

Metadata được đọc từ:

```text
docProps/core.xml
```

Các trường quan trọng:

```text
created
modified
creator
lastModifiedBy
```

---

### File Microsoft Office cũ

Các định dạng sau cần cài thêm `olefile`:

```text
.doc .xls .ppt
```

Metadata được đọc:

```text
create_time
last_saved_time
author
last_saved_by
```

Cài thư viện:

```bat
pip install olefile
```

---

### File PDF

Công cụ quét metadata PDF:

```text
/CreationDate
/ModDate
```

---

### Fallback từ tên file hoặc đường dẫn

Nếu file không có metadata hợp lệ, công cụ có thể lấy ngày từ tên file hoặc đường dẫn.

Ví dụ được hỗ trợ:

```text
17.05.2016
17-05-2016
17_05_2016
2016-05-17
2016.05.17
```

---

## Thiết lập khuyến nghị cho PharmSolu

Khi sửa dữ liệu trong thư mục lớn như `D:\PharmSolu`, nên dùng cấu hình sau:

```text
Scan folders recursively: ON
Only supported extensions: ON

Scan only files with filesystem date > current date: ON
Compare by date only, not exact time: ON
Show skipped non-future files: OFF

Prefer Content created date: ON
Allow Date last saved fallback: OFF
Reject future metadata dates: ON
Use filename/path date fallback: ON

CreationTime: ON
ModifiedTime / LastWriteTime: ON
AccessedTime: ON
```

Quan trọng:

```text
Allow Date last saved fallback: OFF
```

Không nên bật tùy chọn này mặc định vì nhiều file có trường `Date last saved` bị sai sang tương lai, ví dụ `2062`.

---

## Cấu trúc project

```text
office_metadata_timestamp_gui_v3/
├── office_metadata_timestamp_gui_v3.py
├── run_office_metadata_timestamp_gui_v3.bat
├── requirements.txt
├── README.md
├── README_VN.md
├── readme__office_metadata_timestamp_gui_v3.md
└── context_pack__office_metadata_timestamp_gui_v3.md
```

---

## Yêu cầu hệ thống

Bắt buộc:

```text
Python 3.9+
Tkinter
```

Tkinter thường có sẵn trong Python trên Windows.

Khuyến nghị:

```text
olefile>=0.47
```

Cài thư viện:

```bat
pip install -r requirements.txt
```

Hoặc:

```bat
pip install olefile
```

---

## Cách chạy

### Cách 1: Chạy bằng file BAT

Double-click:

```bat
run_office_metadata_timestamp_gui_v3.bat
```

File BAT sẽ tự tìm Portable Python 3.12 trong thư mục hiện tại và các thư mục cha:

```text
Python312\python.exe
..\Python312\python.exe
..\..\Python312\python.exe
..\..\..\Python312\python.exe
```

---

### Cách 2: Chạy trực tiếp bằng Python

```bat
python office_metadata_timestamp_gui_v3.py
```

---

## Ví dụ input

Paste một hoặc nhiều đường dẫn vào GUI:

```text
"D:\PharmSolu\SERVERCS15-GUEST\Kho Nguyen Lieu Bao Bi\Thien\Doc1.docx"
"D:\PharmSolu\bt_loan\Ploan\HD\hdSAGOPHA.doc"
"D:\PharmSolu\SERVERCS15-GUEST\Phong Nghien Cuu Phat Trien\THAO NCPT\GTCN du an san xuat san pham cong nghe 17.05.2016 .doc"
```

Có thể paste trực tiếp cả thư mục:

```text
"D:\PharmSolu"
```

Sau đó bấm:

```text
Scan metadata
```

Kiểm tra bảng kết quả và log CSV. Nếu đúng, bấm:

```text
Apply timestamps
```

---

## Ý nghĩa bộ lọc v3

Tùy chọn mới của v3:

```text
Scan only files with filesystem date > current date
```

Khi bật tùy chọn này, chương trình chỉ đọc metadata của các file có timestamp trong Windows lớn hơn ngày hiện tại. Các file bình thường sẽ bị bỏ qua trước khi đọc metadata, giúp quét thư mục lớn nhanh hơn và an toàn hơn.

Ví dụ nếu hôm nay là:

```text
2026-06-16
```

và bật:

```text
Compare by date only, not exact time
```

thì chỉ các file có ngày từ:

```text
2026-06-17
```

trở đi mới được scan.

---

## Log CSV

Công cụ tự động tạo log trong thư mục:

```text
office_timestamp_logs_v3
```

Các cột chính:

```text
path
status
message
metadata_source
target_source
future_fs_fields
fs_time_error
fs_creation_time
fs_modified_time
fs_accessed_time
metadata_created
metadata_modified
author
last_saved_by
target_time
applied
creation_set
modified_set
accessed_set
```

Nên kiểm tra log trước khi bấm **Apply timestamps** trên số lượng file lớn.

---

## Quy trình an toàn

Khuyến nghị sử dụng theo thứ tự:

1. Paste file hoặc folder path.
2. Bấm **Scan metadata**.
3. Kiểm tra bảng kết quả.
4. Mở CSV log để kiểm tra ngày mục tiêu.
5. Chỉ bấm **Apply timestamps** khi chắc chắn ngày đúng.
6. Với kho dữ liệu quan trọng, nên backup trước khi apply hàng loạt.

---

## Đưa project lên GitHub

### 1. Mở thư mục project

```bat
cd /d D:\path\to\office_metadata_timestamp_gui_v3
```

### 2. Khởi tạo Git

```bat
git init
```

### 3. Thêm toàn bộ file

```bat
git add .
```

### 4. Commit

```bat
git commit -m "Initial release: Office Metadata Timestamp GUI v3"
```

### 5. Đổi branch chính thành main

```bat
git branch -M main
```

### 6. Thêm remote GitHub

Thay `YOUR_USERNAME` và `YOUR_REPOSITORY` bằng tài khoản và repo của bạn.

```bat
git remote add origin https://github.com/YOUR_USERNAME/YOUR_REPOSITORY.git
```

### 7. Push lên GitHub

```bat
git push -u origin main
```

---

## Cập nhật repo GitHub sau này

Sau khi sửa code:

```bat
git status
git add .
git commit -m "Update Office Metadata Timestamp GUI v3"
git push
```

---

## `.gitignore` khuyến nghị

```gitignore
__pycache__/
*.pyc
*.pyo
*.pyd
*.log
*.tmp
.DS_Store
Thumbs.db

office_timestamp_logs_v3/
timestamp_setter_logs/
timestamp_repair_logs/

.venv/
venv/
env/
```

---

## Ghi chú kỹ thuật cho developer

### Set timestamp trên Windows

Công cụ dùng Windows API:

```text
kernel32.CreateFileW
kernel32.SetFileTime
```

thông qua Python `ctypes`.

Có thể set:

```text
CreationTime
LastWriteTime
LastAccessTime
```

mà không cần `pywin32`.

---

### Hỗ trợ đường dẫn dài

Đường dẫn được chuyển sang dạng an toàn cho Windows long path:

```text
\\?\C:\path\to\file
```

UNC path:

```text
\\?\UNC\server\share\path
```

---

### Thứ tự ưu tiên metadata

Mặc định:

```text
1. Office/PDF Content created date
2. Filename/path date fallback
3. Date last saved chỉ dùng nếu người dùng bật thủ công
```

---

## Cảnh báo

Công cụ này thay đổi timestamp của file. Hãy luôn scan và xem log CSV trước khi apply trên thư mục lớn. Với dữ liệu quan trọng, nên có bản backup.

---

## Tác giả / Dự án

**PharmSolu Timestamp Repair Utility**  
NGHIEN CUU THUOC // RnD PHARMA PLUS  

Website chính thức:

```text
https://www.nghiencuuthuoc.com
```

PharmApp demo:

```text
https://www.pharmapp.dev
```
