# AgribankV3

Ứng dụng desktop cho quá trình chuyển đổi add-in AgribankV2 sang Python.

Vertical slice đầu tiên đã hoạt động:

- Kết nối tới phiên Microsoft Excel đang mở.
- Hiển thị workbook, worksheet và vùng ô đang chọn.
- Chuyển vùng văn bản sang chữ hoa, chữ thường hoặc viết hoa tên.
- Giữ nguyên công thức, số và ô trống.
- Hoàn tác lần chuyển đổi gần nhất trong phiên AgribankV3.
- Tự động kết nối khi phát hiện Excel đang chạy.
- Cho phép chọn phiên bản Excel đã cài và tạo workbook mới nếu Excel chưa chạy.
- Bấm vào logo hoặc chữ **AgribankV3** để thu gọn/mở rộng sidebar; rê chuột
  lên vùng logo để xem hướng dẫn.
- Mặc định giữ Excel trong cửa sổ riêng để Ribbon, nhập liệu và hộp thoại hoạt
  động mượt và đúng chuẩn.
- Sidebar có nút **Mở Excel/Hiện Excel** để chuyển nhanh sang workbook.

## Chạy nhanh

Từ PowerShell:

```powershell
.\run.ps1
```

Lần chạy đầu tiên script sẽ tạo `.venv` và cài các thư viện trong
`pyproject.toml`.

Hoặc chạy thủ công:

```powershell
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install -e .
.\.venv\Scripts\python.exe -m agribank_v3
```

## Thử chức năng Excel

1. Mở Excel và mở một workbook thử nghiệm.
2. Chọn vùng ô có dữ liệu văn bản.
3. Chạy AgribankV3 và nhấn **Kết nối Excel**.
4. Mở **Chức năng → Chuyển kiểu chữ**.
5. Chọn kiểu chuyển đổi và nhấn **Áp dụng**.

Không nên thử lần đầu trên file dữ liệu thật. Hãy dùng một bản sao hoặc workbook
thử nghiệm để đối chiếu kết quả.

Nếu chưa mở Excel, nhấn **Mở Excel** trong sidebar, chọn phiên bản trong danh sách và
nhấn **Mở Excel và tạo workbook**. AgribankV3 sẽ đợi Excel sẵn sàng, mở workbook
trắng `AgribankV3-New.xlsx` và kết nối tự động.

AgribankV3 không thay đổi parent, style hoặc kích thước native của cửa sổ Excel.

## Trắc nghiệm nghiệp vụ

Mở **Trắc nghiệm → Kiểm tra nghiệp vụ** để:

- Chọn nhân viên, chuyên đề, nghiệp vụ, số câu và giới hạn thời gian.
- Tạo đề ngẫu nhiên theo từng chuyên đề thuộc nghiệp vụ; cấu hình số câu của
  từng chuyên đề được tự lưu.
- Xuất bộ câu hỏi ngẫu nhiên sang Excel theo đúng hạn mức đang chọn, kèm
  đáp án A/B/C/D, cột đáp án đúng và cột để người dùng tự chọn.
- Quản lý nghiệp vụ và các chuyên đề con trong tab
  **Quản Lý Dữ Liệu → Nghiệp Vụ & Chuyên Đề**.
- Tìm kiếm, thêm, sửa, xóa hoặc xóa hàng loạt câu hỏi trong tab
  **Quản Lý Dữ Liệu → Chỉnh Sửa Câu Hỏi**.
- Nhập hàng loạt câu hỏi từ file `.xlsx` hoặc `.xlsm`.
- Làm bài, kiểm tra từng đáp án, xem nguồn tham khảo và kết quả cuối.
- Lưu lịch sử lượt làm cùng chi tiết từng câu trả lời.

Dòng đầu của file nhập Excel cần các cột: `Mã nghiệp vụ`, `Nghiệp vụ`,
`Chuyên đề`, `Câu hỏi`, `Đáp án A`, `Đáp án B`, `Đáp án C`, `Đáp án D`,
`Đáp án đúng`, `Đáp án không đảo` và `Nguồn`. Các cột `ID`, `STT` là tùy
chọn. Nhập `A,C` tại cột `Đáp án không đảo` nếu A và C phải giữ nguyên vị trí.
Thứ tự A/B/C/D trong database luôn là thứ tự đề cương gốc; tùy chọn hoán đổi
chỉ tác động lên phiên làm bài. Nút **Tải Excel mẫu** tạo sẵn đúng cấu trúc này.

Dữ liệu vận hành nằm tại `data/agribank_v3.sqlite3`. Lần đầu sử dụng hoặc khi
file `Data/AgribankMenuData.mdb` thay đổi, ứng dụng so sánh SHA-256 và nhập lại
4.070 câu hỏi trong một transaction. Trước khi cập nhật, SQLite hiện tại được
sao lưu vào `data/backups`.

MDB chỉ được mở ở chế độ đọc. Có thể trỏ tới một bản MDB khác bằng biến môi
trường `AGRIBANKV3_ACCESS_DB`.

## Chẩn đoán kết nối

Nếu giao diện không kết nối được, giữ Excel và workbook đang mở rồi chạy:

```powershell
.\.venv\Scripts\python.exe .\tools\diagnose_excel.py
```

AgribankV3 dùng COM late-binding và tự nhận diện Excel 2007, 2010, 2013 và
Excel 2016 trở lên. Excel và AgribankV3 phải chạy cùng mức quyền; nếu một chương
trình chạy Administrator còn chương trình kia không chạy Administrator thì
Windows có thể chặn COM.
