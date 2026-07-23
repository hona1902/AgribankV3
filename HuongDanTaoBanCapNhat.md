# Hướng dẫn tạo bản cập nhật AgribankV3

Tài liệu này hướng dẫn người phát triển tạo gói cập nhật nội bộ cho AgribankV3 bằng ứng dụng `AgribankV3 Update Builder`.

Update Builder chỉ tạo:

- `AgribankV3_<version>.zip`
- `manifest.json`
- thư mục/file migration nếu có

Update Builder không đóng gói database người dùng và không lấy bản cập nhật từ internet. Bản cập nhật phải được đặt trong thư mục nội bộ do người dùng hoặc đơn vị cấu hình.

## Nguyên tắc bắt buộc

Luôn giữ nguyên dữ liệu người dùng. Không đưa các file database thật vào gói cập nhật.

Database chỉ được thay đổi bằng migration an toàn:

- Thêm bảng mới nếu chưa tồn tại.
- Thêm cột mới nếu chưa tồn tại.
- Thêm index mới nếu chưa tồn tại.
- Thêm dữ liệu mặc định bằng kiểu không ghi đè, ví dụ `INSERT OR IGNORE`.

Không dùng migration để xóa hoặc ghi đè dữ liệu người dùng:

- Không `DROP TABLE`.
- Không xóa cột.
- Không đổi kiểu dữ liệu bằng cách tạo lại bảng rồi copy thiếu dữ liệu.
- Không update hàng loạt setting/cấu hình người dùng đã chỉnh.
- Không copy `DuLieuV3.db`, `quiz.db`, `*.db`, `*.sqlite`, `*.mdb`, `*.accdb` vào gói update.

## Thư mục liên quan

Source chính:

```text
D:\Soft VBA\THAOTAC-EXCEL\src\5491-AgribankV2\AgribankV2-PyThon
```

Update Builder:

```text
tools/update_builder
```

App Update Builder đã build:

```text
dist/AgribankV3UpdateBuilder/AgribankV3UpdateBuilder.exe
```

Zip portable của Update Builder:

```text
dist/AgribankV3UpdateBuilder-portable.zip
```

Thư mục Update mặc định:

```text
X:\public\5491\06- PHONG KHKD\NAM\AgribankV3\Update
```

Có thể đổi thư mục Update trong app AgribankV3 hoặc trong Update Builder nếu đơn vị dùng vị trí khác.

## Chuẩn bị trước khi tạo bản cập nhật

1. Hoàn tất sửa code ở source AgribankV3.
2. Chạy test hoặc chạy thử app từ source.
3. Nếu có thay đổi database, chuẩn bị database mẫu bản cũ và bản mới để so sánh.
4. Xác định version mới, ví dụ `0.1.5`.
5. Viết ghi chú cập nhật ngắn gọn để người dùng biết bản mới thay đổi gì.
6. Không để file dữ liệu thật trong source nếu file đó không cần thiết cho code.

Nên dùng version tăng dần theo dạng:

```text
0.1.1
0.1.2
0.1.3
```

Version mới phải lớn hơn version hiện tại. Nếu version bằng hoặc thấp hơn, app người dùng sẽ không xem đó là bản cập nhật mới.

Trong Update Builder, cần hiểu rõ version hiện tại của người dùng khác với
version trong source:

- `Phiên bản trong source`: version của code developer đang đóng gói.
- `Phiên bản phát hành trước`: version app người dùng đang dùng hoặc bản đã có
  trong thư mục Update.
- `Phiên bản mới`: version gói update sẽ tạo.

`Phiên bản mới` được phép bằng `Phiên bản trong source`. Điều kiện quan trọng là
`Phiên bản mới` phải lớn hơn `Phiên bản phát hành trước`.

Ví dụ:

```text
Người dùng đang dùng: 0.1.0
Source developer: 0.1.1
Phiên bản gói update: 0.1.1
```

Trường hợp này hợp lệ vì `0.1.1 > 0.1.0`.

## Mở Update Builder

Có hai cách.

Cách 1: Chạy exe đã build:

```text
dist/AgribankV3UpdateBuilder/AgribankV3UpdateBuilder.exe
```

Cách 2: Chạy từ source:

```powershell
.\.venv\Scripts\python.exe .\tools\update_builder\update_builder_app.py
```

Giao diện có 3 tab:

- `Thông tin cập nhật`
- `Database migration`
- `Tạo gói & Log`

## Quy trình nhanh được khuyến nghị

Trong đa số trường hợp, dùng nút `Tạo bản cập nhật tự động`.

1. Mở tab `Thông tin cập nhật`.
2. Chọn thư mục source AgribankV3.
3. Bấm `Đọc phiên bản`.
4. Nhập `Phiên bản mới`.
5. Nhập `Ghi chú cập nhật`.
6. Chọn `Thư mục Update`.
7. Bấm `Tạo bản cập nhật tự động`.

Builder sẽ tự:

- đọc version hiện tại;
- tìm database dev hiện tại trong source;
- tìm schema snapshot của bản phát hành trước;
- so sánh schema để biết database có đổi hay không;
- tạo update không migration nếu database không đổi;
- tự tạo Python migration nếu database có thay đổi an toàn;
- thử migration trên bản copy;
- thêm migration vào manifest;
- tạo zip và `manifest.json`;
- lưu schema snapshot mới cho version vừa build.

Nếu chưa có snapshot bản trước, bấm `Thiết lập baseline database` và chọn
database của bản đang phát hành hiện tại. Sau đó nhập version baseline, ví dụ
`0.1.4`. Builder sẽ lưu:

```text
tools/update_builder/schema_snapshots/schema_0.1.4.json
```

Nếu Builder không đủ dữ liệu để tự kiểm tra database, app sẽ hỏi chọn một trong
ba hướng:

- chọn database bản cũ và database bản mới để tạo baseline;
- tiếp tục tạo update chỉ đổi code;
- hủy.

## Tạo bản cập nhật chỉ thay đổi code

Dùng khi bản mới chỉ sửa code, giao diện, nghiệp vụ, file Python, tài nguyên, nhưng không thay đổi cấu trúc database.

### Bước 1: Nhập thông tin cập nhật

Mở tab `Thông tin cập nhật`.

1. Ở `Thư mục source AgribankV3`, chọn thư mục source dự án.
2. Bấm `Đọc phiên bản`.
3. Kiểm tra `Phiên bản hiện tại`.
4. Nhập `Phiên bản mới`.
5. Nhập `Ngày phát hành`.
6. Tích `Yêu cầu khởi động lại` nếu muốn app người dùng khởi động lại sau cập nhật.
7. Tích `Tự cập nhật version trong source trước khi đóng gói` nếu muốn Builder tự ghi version mới vào source.
8. Chọn `Kiểu gói cập nhật`.
9. Nhập `Ghi chú cập nhật`, mỗi dòng là một ghi chú.
10. Chọn `Thư mục Update`, ví dụ thư mục nội bộ dùng chung.

Nếu người dùng chạy AgribankV3 bằng file `AgribankV3.exe`, phải build lại app
chính trước rồi chọn `Gói app đã build EXE`. Không dùng `Gói runtime tối thiểu`
cho bản exe, vì gói runtime chỉ thay source `src/agribank_v3` và không thay code
đã đóng bên trong `AgribankV3.exe`.

Ví dụ ghi chú:

```text
Sửa lỗi đọc phiên bản cập nhật.
Tối ưu giao diện phần cài đặt.
Bổ sung kiểm tra dữ liệu trước khi tạo gói.
```

### Bước 2: Không bật migration

Mở tab `Database migration`.

Đảm bảo không tích `Có thay đổi database` nếu bản cập nhật không thay đổi database.

### Bước 3: Kiểm tra và tạo gói

Mở tab `Tạo gói & Log`.

1. Bấm `Kiểm tra dữ liệu`.
2. Nếu hợp lệ, bấm `Tạo bản cập nhật`.
3. Xem log để biết zip và manifest đã tạo ở đâu.
4. Bấm `Mở thư mục Update` để kiểm tra kết quả.

Kết quả mong muốn:

```text
Update
|-- manifest.json
`-- AgribankV3_0.1.5.zip
```

## Tạo bản cập nhật có thay đổi database

Dùng khi bản mới cần thêm bảng, thêm cột, thêm index hoặc thêm dữ liệu mặc định.

Không được copy database mới vào gói cập nhật. Người dùng mỗi máy có database riêng, nên phải cập nhật bằng migration.

### Bước 1: Chuẩn bị database để so sánh

Cần có hai file database:

- Database phiên bản cũ: schema giống bản người dùng đang chạy.
- Database phiên bản mới: schema sau khi developer đã thêm bảng/cột/default mới.

Chỉ dùng bản copy để kiểm thử. Không dùng trực tiếp database thật của người dùng.

### Bước 2: So sánh database

Mở tab `Database migration`.

1. Tích `Có thay đổi database`.
2. Ở `Database phiên bản cũ`, chọn database cũ.
3. Ở `Database phiên bản mới`, chọn database mới/dev.
4. Nhập `Version migration`, thường trùng với version app mới.
5. Nhập `Description`.
6. Bấm `So sánh database`.

Kết quả so sánh sẽ chia thành:

- Bảng mới.
- Cột mới.
- Index mới.
- Dữ liệu mặc định mới trong `app_preferences`.
- Cần xử lý thủ công.

Nếu có mục trong `Cần xử lý thủ công`, phải đọc kỹ trước khi phát hành. Builder không tự xử lý các thay đổi nguy hiểm.

### Bước 3: Tạo migration gợi ý

1. Bấm `Tạo migration gợi ý`.
2. Kiểm tra file migration được tạo trong:

```text
tools/update_builder/generated_migrations/migrate_<version>.py
```

3. Mở file đó và đọc lại logic migration.

Migration hợp lệ phải có tính idempotent, nghĩa là chạy lại không gây lỗi và không ghi đè dữ liệu đã có.

Ví dụ nguyên tắc đúng:

```sql
CREATE TABLE IF NOT EXISTS ...
CREATE INDEX IF NOT EXISTS ...
INSERT OR IGNORE INTO app_preferences ...
```

Với Python migration, cần kiểm tra tồn tại bảng/cột trước khi thêm.

### Bước 4: Thử migration trên bản copy

Bấm `Thử chạy migration trên bản copy`.

Yêu cầu kết quả:

- Migration chạy thành công.
- Bảng/cột/index mới xuất hiện.
- Dữ liệu cũ vẫn còn.
- Setting người dùng đã chỉnh không bị ghi đè.
- Nếu lỗi, chỉ lỗi trên bản copy, không ảnh hưởng database thật.

Nếu test chưa đạt, sửa migration rồi thử lại.

### Bước 5: Đưa migration vào gói update

Có hai cách.

Cách 1: Python migration trong code:

1. Tích `Tự thêm migration Python vào src/agribank_v3/update/db_migrations.py` nếu muốn Builder tự chèn.
2. Builder sẽ backup `db_migrations.py`.
3. Builder chèn function migration và đăng ký vào `default_python_migrations()`.
4. Bấm `Thêm migration vào danh sách`.
5. Dòng migration sẽ có `Python migration` được tích.

Cách 2: SQL migration:

1. Tích `Có thay đổi database`.
2. Bấm `Thêm migration`.
3. Nhập `Version`.
4. Bấm `Chọn file .sql`.
5. Nhập `Description`.
6. Không tích `Python migration` nếu dùng file SQL.

Ưu tiên Python migration cho các thay đổi cần kiểm tra tồn tại bảng/cột trước khi chạy.

### Bước 6: Tạo gói update

Mở tab `Tạo gói & Log`.

1. Bấm `Kiểm tra dữ liệu`.
2. Đọc cảnh báo nếu có.
3. Nếu mọi thứ đúng, bấm `Tạo bản cập nhật`.
4. Kiểm tra log.

Kết quả mong muốn:

```text
Update
|-- manifest.json
|-- AgribankV3_0.1.5.zip
`-- migrations
    `-- 0.1.5.sql
```

Nếu dùng Python migration trong code, có thể không có file SQL trong thư mục `migrations`; manifest vẫn ghi migration theo version để app biết cần chạy Python migration tương ứng.

## Kiểm tra gói cập nhật trước khi phát hành

Trước khi cho người dùng cập nhật, phải kiểm tra các mục sau.

### Kiểm tra manifest.json

Mở `manifest.json` trong thư mục Update.

Ví dụ:

```json
{
  "latest_version": "0.1.5",
  "package": "AgribankV3_0.1.5.zip",
  "release_date": "2026-07-17",
  "required_app_restart": true,
  "notes": [
    "Bổ sung chức năng cập nhật phiên bản"
  ],
  "database_migrations": []
}
```

Cần kiểm tra:

- `latest_version` đúng version mới.
- `package` đúng tên file zip.
- `release_date` đúng ngày phát hành.
- `notes` dễ hiểu.
- `database_migrations` đúng với thay đổi database.

### Kiểm tra zip

Mở zip và kiểm tra không có database người dùng:

- Không có `DuLieuV3.db`.
- Không có `quiz.db`.
- Không có file `*.db`, `*.sqlite`, `*.sqlite3`.
- Không có file backup/log/temp không cần thiết.

Builder đã có danh sách loại trừ, nhưng vẫn nên kiểm tra lại trước khi phát hành.

### Kiểm tra cập nhật trên bản app copy

Nên tạo một bản copy app cũ và database cũ để test:

1. Copy app cũ sang thư mục test.
2. Copy database cũ sang thư mục test.
3. Cấu hình app test trỏ tới thư mục Update vừa tạo.
4. Mở app test.
5. Vào phần `Cập nhật phiên bản`.
6. Bấm kiểm tra cập nhật.
7. Thực hiện cập nhật.
8. Mở lại app nếu được yêu cầu.

Sau cập nhật, kiểm tra:

- Version app đã lên version mới.
- Chức năng mới hoạt động.
- Dữ liệu cũ vẫn còn.
- Database có bảng/cột/index mới nếu có migration.
- Bảng `app_schema_migrations` có version migration mới.
- Không mất dữ liệu người dùng.

## Phát hành vào thư mục nội bộ

Khi đã kiểm thử xong:

1. Copy hoặc để nguyên `manifest.json` trong thư mục Update nội bộ.
2. Copy hoặc để nguyên `AgribankV3_<version>.zip` trong thư mục Update nội bộ.
3. Nếu có SQL migration, đảm bảo thư mục `migrations` và file `.sql` nằm đúng vị trí.
4. Mở một máy trạm test để kiểm tra app nhìn thấy version mới.

Thư mục Update nội bộ cuối cùng thường có dạng:

```text
Update
|-- manifest.json
|-- AgribankV3_0.1.5.zip
`-- migrations
    `-- 0.1.5.sql
```

App người dùng sẽ đọc `manifest.json`, so sánh version hiện tại với `latest_version`, tải zip từ cùng thư mục nội bộ và chạy migration nếu cần.

## Chạy Update Builder bằng command line

Ngoài giao diện, có thể tạo update bằng command line.

Ví dụ:

```powershell
.\.venv\Scripts\python.exe .\tools\update_builder\update_builder_app.py `
  --source "D:\Soft VBA\THAOTAC-EXCEL\src\5491-AgribankV2\AgribankV2-PyThon" `
  --version 0.1.5 `
  --update-path "X:\public\5491\06- PHONG KHKD\NAM\AgribankV3\Update" `
  --notes "Sửa lỗi cập nhật phiên bản" `
  --notes "Tối ưu giao diện Update Builder"
```

Nếu tạo gói cho bản người dùng chạy exe, build AgribankV3 trước:

```powershell
.\build_portable.ps1
```

Sau đó chạy Builder với mode app/exe:

```powershell
.\.venv\Scripts\python.exe .\tools\update_builder\update_builder_app.py `
  --source "D:\Soft VBA\THAOTAC-EXCEL\src\5491-AgribankV2\AgribankV2-PyThon" `
  --version 0.1.5 `
  --package-mode app `
  --update-path "X:\public\5491\06- PHONG KHKD\NAM\AgribankV3\Update" `
  --notes "Sửa lỗi cập nhật phiên bản"
```

Hoặc dùng file config:

```powershell
.\.venv\Scripts\python.exe .\tools\update_builder\update_builder_app.py `
  --config .\tools\update_builder\build_config.json
```

Command line phù hợp khi đã có cấu hình ổn định. Với migration database, nên dùng giao diện để so sánh và thử migration trực quan hơn.

## Build lại Update Builder exe

Khi code của Update Builder thay đổi, build lại exe:

```powershell
.\build_update_builder.ps1
```

Kết quả:

```text
dist/AgribankV3UpdateBuilder/AgribankV3UpdateBuilder.exe
dist/AgribankV3UpdateBuilder-portable.zip
```

Sau khi build, chạy thử exe:

```text
dist/AgribankV3UpdateBuilder/AgribankV3UpdateBuilder.exe
```

Kiểm tra tối thiểu:

- Cửa sổ mở được.
- Có đủ 3 tab.
- Bấm `Đọc phiên bản` không lỗi.
- Bấm `Kiểm tra dữ liệu` không lỗi crash.
- Log hiển thị bình thường.
- Nếu có migration, thử được `So sánh database` và `Thử chạy migration trên bản copy`.

## Test nên chạy trước khi phát hành

Chạy các test liên quan:

```powershell
.\.venv\Scripts\python.exe -m unittest `
  tests.test_update_builder `
  tests.test_db_schema_diff `
  tests.test_update_manager `
  tests.test_settings -v
```

Nếu có thay đổi lớn ở UI hoặc update flow, nên chạy thêm toàn bộ test:

```powershell
.\.venv\Scripts\python.exe -m unittest discover -v
```

## Lỗi thường gặp

### App người dùng không thấy bản cập nhật

Kiểm tra:

- Đường dẫn Update trong app người dùng có đúng không.
- `manifest.json` có nằm trong thư mục Update không.
- `latest_version` trong manifest có lớn hơn version hiện tại không.
- File zip trong manifest có tồn tại đúng tên không.

### Tạo gói báo version không hợp lệ

Kiểm tra version mới có dạng số chấm số, ví dụ:

```text
0.1.5
1.0.0
```

Không dùng version rỗng hoặc version thấp hơn hiện tại.

### Migration không chạy

Kiểm tra:

- Manifest có `database_migrations` không.
- Version migration chưa có trong `app_schema_migrations`.
- Nếu dùng SQL, file `.sql` có nằm đúng trong thư mục `migrations` không.
- Nếu dùng Python migration, function migration đã được đăng ký trong `default_python_migrations()` chưa.

### Migration chạy lỗi

Không chạy tiếp trên database thật. Làm theo thứ tự:

1. Kiểm tra backup database được tạo trước migration.
2. Đọc log lỗi.
3. Sửa migration trên source.
4. Test lại trên bản copy database cũ.
5. Chỉ phát hành lại khi migration đã chạy thành công và giữ nguyên dữ liệu cũ.

### Zip có chứa database

Không phát hành gói đó.

Kiểm tra lại source và danh sách file bị exclude. Gói update đúng không được chứa database người dùng. Nếu database cần thay đổi, tạo migration thay vì đóng gói database.

## Checklist trước khi bấm tạo bản cập nhật

- Đã sửa xong code.
- Đã chạy thử app từ source.
- Nếu phát hành cho bản exe, đã chạy `.\build_portable.ps1` và chọn
  `Gói app đã build EXE`.
- Version mới lớn hơn version hiện tại.
- Release notes đã rõ ràng.
- Thư mục Update đúng là thư mục nội bộ muốn phát hành.
- Nếu không đổi database, không tích `Có thay đổi database`.
- Nếu đổi database, migration đã được tạo và thử trên bản copy.
- Không có thay đổi nguy hiểm chưa xử lý.

## Checklist trước khi cho người dùng cập nhật

- `manifest.json` đúng.
- Zip đúng tên.
- Zip không chứa database.
- Migration an toàn.
- Đã test cập nhật trên bản app copy.
- Dữ liệu người dùng trong bản test vẫn còn.
- App mở lại bình thường sau cập nhật.
- Thư mục Update nội bộ có đủ `manifest.json`, zip và migration nếu có.

## Quy trình khuyến nghị

Với bản cập nhật nhỏ, chỉ đổi code:

1. Sửa code.
2. Chạy test.
3. Mở Update Builder.
4. Nhập version mới và ghi chú.
5. Bấm `Kiểm tra dữ liệu`.
6. Bấm `Tạo bản cập nhật`.
7. Test trên bản app copy.
8. Phát hành vào thư mục Update nội bộ.

Với bản cập nhật có database:

1. Sửa code và schema dev.
2. Chuẩn bị database cũ và database mới.
3. So sánh database trong Update Builder.
4. Tạo migration gợi ý.
5. Đọc và chỉnh migration nếu cần.
6. Thử migration trên bản copy.
7. Thêm migration vào danh sách.
8. Tạo gói update.
9. Test cập nhật end-to-end trên bản app copy.
10. Phát hành vào thư mục Update nội bộ.
