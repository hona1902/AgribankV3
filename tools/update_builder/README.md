# AgribankV3 Update Builder

Ứng dụng này dành cho người phát triển AgribankV3 để tạo gói cập nhật nội bộ.
Nó tạo `AgribankV3_<version>.zip`, `manifest.json`, migration nếu cần và schema
snapshot cho lần build sau. App không chạy cập nhật cho người dùng cuối.

Update Builder không đóng gói database người dùng. Database dev chỉ được dùng để
đọc schema và tạo migration, không được đưa vào zip.

Thư mục Update mặc định:

```text
X:\public\5491\06- PHONG KHKD\NAM\AgribankV3\Update
```

## Quy trình khuyến nghị: Tạo bản cập nhật tự động

1. Mở Update Builder.
2. Chọn thư mục source AgribankV3.
3. Bấm `Đọc phiên bản`.
4. Nhập `Phiên bản mới`.
5. Chọn `Kiểu gói cập nhật`: mặc định nên giữ `Gói runtime tối thiểu`.
6. Nhập ghi chú cập nhật.
7. Chọn thư mục Update.
8. Bấm `Xem file sẽ đóng gói` để kiểm tra dung lượng và top file lớn.
9. Bấm `Tạo bản cập nhật tự động`.

Builder sẽ tự:

- đọc version hiện tại;
- tìm database dev hiện tại;
- tìm schema snapshot bản phát hành trước;
- so sánh schema;
- tạo update không migration nếu database không đổi;
- tự tạo Python migration nếu database có thay đổi an toàn;
- thử migration trên bản copy;
- chèn migration vào `src/agribank_v3/update/db_migrations.py`;
- thêm migration vào manifest;
- tạo zip và `manifest.json`;
- lưu schema snapshot mới.

Nếu phát hiện thay đổi nguy hiểm, Builder dừng và yêu cầu xử lý thủ công.

## Các kiểu gói cập nhật

Update Builder có 4 kiểu gói:

- `Gói runtime tối thiểu`: chế độ mặc định và khuyến nghị. Chỉ đóng gói các file
  cần để app chạy, như `src/agribank_v3/`, `pyproject.toml`, `requirements.txt`,
  `run.ps1`, `run.bat`, `README.md`, `assets/`, `resources/`, `templates/`,
  `MauBieu/`. Chỉ dùng khi AgribankV3 chạy từ source/Python.
- `Gói app đã build EXE`: dùng khi người dùng chạy `AgribankV3.exe`. Trước khi
  tạo gói này phải build AgribankV3 để có `dist/AgribankV3`. Zip sẽ chứa
  `AgribankV3.exe` và `_internal/`, nhưng vẫn loại trừ database, log, backup và
  dữ liệu người dùng.
- `Gói source đầy đủ`: dùng khi developer thật sự muốn đóng gói source rộng hơn.
  Chế độ này vẫn loại trừ dữ liệu lớn, database, file build, file log và thư mục
  Update.
- `Gói delta - chỉ file thay đổi`: so với snapshot file của bản phát hành trước
  và chỉ đóng gói file mới hoặc file đổi sha256. Manifest có `package_type:
  "delta"` và `base_version`.

Sau mỗi lần build thành công, Builder lưu snapshot file tại:

```text
tools/update_builder/release_snapshots/files_<version>.json
```

Snapshot này dùng cho lần tạo delta kế tiếp.

## Vì sao zip update bị nặng?

Zip update thường bị nặng nếu đóng gói cả project thay vì runtime, ví dụ:

- `dist/`, `build/`, `.venv/`;
- `tools/update_builder/` và file exe/zip của chính Update Builder;
- `tests/`, `docs/`, `DuLieuTEST/`, `sample_data/`;
- database `.db`, `.sqlite`, `.mdb`, `.accdb`;
- file Excel/Word/PDF test, file zip/exe cũ;
- `Update/`, `logs/`, `backups/`, `temp/`, `KetQua/`, `outputs/`.

Chế độ mặc định `Gói runtime tối thiểu` loại trừ các phần trên để tránh tạo zip
hàng trăm MB. Update Builder là công cụ developer, không phải thành phần runtime
của AgribankV3, nên không được đóng gói vào bản update AgribankV3.

Trước khi phát hành:

- bấm `Xem file sẽ đóng gói`;
- kiểm tra tổng dung lượng trước nén;
- kiểm tra top 20 file lớn và top 10 thư mục lớn;
- đảm bảo zip không chứa `dist`, `tools`, `tests`, database, log, `KetQua`,
  thư mục `Update` hoặc file build cũ.

## Hiểu đúng các loại version

Update Builder phân biệt 3 loại version:

- `Phiên bản trong source`: version đang ghi trong code hiện tại, ví dụ
  `src/agribank_v3/__init__.py`.
- `Phiên bản phát hành trước`: version app người dùng đang có hoặc bản đã phát
  hành trong thư mục Update.
- `Phiên bản mới`: version của gói update sẽ tạo.

Điều kiện đúng là `Phiên bản mới` phải lớn hơn `Phiên bản phát hành trước`, không
phải lớn hơn `Phiên bản trong source`.

Ví dụ hợp lệ:

```text
Người dùng đang dùng: 0.1.0
Source developer đã là: 0.1.1
Gói update sẽ tạo: 0.1.1
```

Trường hợp này hợp lệ vì `0.1.1 > 0.1.0`, dù version mới bằng version trong
source.

Nếu version mới lớn hơn version trong source, nên tích
`Tự cập nhật version trong source trước khi đóng gói` hoặc sửa source trước khi
build. Nếu version mới thấp hơn version trong source, Builder sẽ chặn để tránh
đóng gói nhầm bản thấp hơn.

Nếu thư mục Update đã có cùng version đang tạo lại để test, UI sẽ hỏi xác nhận
backup bản cũ và tạo lại cùng version.

## Chạy giao diện

```powershell
.\.venv\Scripts\python.exe .\tools\update_builder\update_builder_app.py
```

Hoặc chạy exe đã build:

```text
dist/AgribankV3UpdateBuilder/AgribankV3UpdateBuilder.exe
```

## Database dev và schema snapshot

Khi chọn source, Builder tự tìm database dev theo thứ tự:

1. `data/DuLieuV3.db`
2. `src/agribank_v3/data/DuLieuV3.db`
3. `DuLieuV3.db` trong source
4. File `DuLieuV3.db` tìm được trong source

Schema snapshot mặc định lưu tại:

```text
tools/update_builder/schema_snapshots/schema_<version>.json
```

Snapshot gồm schema bảng, cột, index, `app_preferences` default key và checksum.
Sau mỗi lần build thành công, Builder lưu snapshot cho version mới để lần sau
tự so sánh.

## Thiết lập baseline lần đầu

Nếu chưa có snapshot bản trước, Builder không tự đoán.

Cách thiết lập:

1. Mở tab `Thông tin cập nhật`.
2. Bấm `Thiết lập baseline database`.
3. Chọn database của bản đang phát hành hiện tại.
4. Nhập version baseline, ví dụ `0.1.4`.
5. Builder lưu `schema_0.1.4.json`.

Từ lần sau, chỉ cần bấm `Tạo bản cập nhật tự động`.

Nếu thiếu snapshot hoặc database dev khi tạo tự động, UI sẽ hỏi:

- chọn database bản cũ và bản mới để Builder tạo baseline;
- tiếp tục tạo update chỉ đổi code;
- hủy.

## Khi database không đổi

Manifest sẽ có:

```json
{
  "database_migrations": []
}
```

Zip vẫn loại trừ database và dữ liệu người dùng.

## Khi database có thay đổi an toàn

Builder tự xử lý các thay đổi sau:

- thêm bảng;
- thêm cột;
- thêm index;
- thêm key mặc định mới trong `app_preferences`.

Builder tạo file migration gợi ý tại:

```text
tools/update_builder/generated_migrations/migrate_<version>.py
```

Sau đó Builder chèn function migration vào:

```text
src/agribank_v3/update/db_migrations.py
```

Manifest sẽ có migration Python:

```json
{
  "version": "0.1.5",
  "file": "",
  "description": "Migration database tự động cho phiên bản 0.1.5"
}
```

## Khi database có thay đổi nguy hiểm

Builder dừng và không tạo update tự động nếu phát hiện:

- xóa bảng;
- xóa cột;
- đổi kiểu dữ liệu/ràng buộc/default của cột;
- đổi value của key `app_preferences` đã tồn tại;
- thay đổi cần xử lý thủ công.

Khi đó dùng tab `Database migration` để xem chi tiết, tự viết migration an toàn
và thử trên bản copy trước khi phát hành.

## Command line tự động

```powershell
.\.venv\Scripts\python.exe .\tools\update_builder\update_builder_app.py `
  --source "D:\Soft VBA\THAOTAC-EXCEL\src\5491-AgribankV2\AgribankV2-PyThon" `
  --version 0.1.6 `
  --update-path "X:\public\5491\06- PHONG KHKD\NAM\AgribankV3\Update" `
  --notes "Sửa lỗi bảng kê" `
  --auto-detect-db
```

Các option liên quan:

- `--auto-detect-db`: bật luồng tự động.
- `--dev-db "path\DuLieuV3.db"`: chỉ định database dev nếu Builder không tự tìm.
- `--baseline-db "path\DuLieuV3_old.db"`: chỉ định database bản cũ để so sánh.
- `--create-baseline`: chỉ tạo schema snapshot từ `--baseline-db`.
- `--snapshot-dir "path"`: chỉ định thư mục lưu snapshot.
- `--code-only-if-missing-baseline`: nếu thiếu snapshot/baseline thì vẫn tạo update chỉ đổi code.
- `--previous-release-version 0.1.0`: chỉ định version phát hành trước.
- `--allow-rebuild-same-version`: cho phép tạo lại cùng version để test.
- `--package-mode runtime|app|source|delta`: chọn kiểu gói cập nhật. Mặc định là
  `runtime`. Có thể dùng alias `exe`, `frozen`, `pyinstaller` cho mode `app`.

Ví dụ tạo baseline:

```powershell
.\.venv\Scripts\python.exe .\tools\update_builder\update_builder_app.py `
  --create-baseline `
  --baseline-db "D:\baseline\DuLieuV3.db" `
  --version 0.1.4
```

## Chế độ nâng cao: tạo update thủ công

Dùng khi cần tự kiểm soát migration hoặc xử lý thay đổi nguy hiểm.

1. Tab `Thông tin cập nhật`: nhập source, version, notes, thư mục Update.
2. Tab `Database migration`: tích `Có thay đổi database` nếu cần.
3. Thêm migration SQL hoặc dùng Python migration có sẵn.
4. Tab `Tạo gói & Log`: bấm `Kiểm tra dữ liệu`.
5. Bấm `Tạo bản cập nhật`.

## File không đưa vào zip

Builder loại trừ database và dữ liệu người dùng:

- `data/*.db`, `*.db`, `*.sqlite`, `*.sqlite3`, `*.mdb`, `*.accdb`
- `KetQua`, `outputs`, `logs`, `backups`, `temp`
- `.git`, `.venv`, `__pycache__`, `.pytest_cache`, `.mypy_cache`
- `dist`, `build`, `*.spec` khi đóng gói runtime/source. Riêng mode
  `Gói app đã build EXE` sẽ lấy nội dung trong `dist/AgribankV3` làm payload và
  vẫn bỏ database/dữ liệu người dùng.

Không đóng gói `data/DuLieuV3.db` hoặc `data/quiz.db`. Nếu database cần thay đổi,
hãy tạo migration an toàn, không `DROP TABLE`, không ghi đè dữ liệu người dùng.

## Kiểm thử trước khi phát hành

Chạy test:

```powershell
.\.venv\Scripts\python.exe -m unittest `
  tests.test_update_builder `
  tests.test_db_schema_diff `
  tests.test_update_manager `
  tests.test_settings -v
```

Kiểm tra gói update:

- Zip không chứa database.
- `manifest.json` đúng version và đúng package.
- Nếu có migration, migration đã chạy thử trên bản copy.
- Dữ liệu người dùng trong bản test vẫn còn.
- `app_schema_migrations` có version migration mới sau khi cập nhật.

## Build lại Update Builder exe

```powershell
.\build_update_builder.ps1
```

Kết quả:

```text
dist/AgribankV3UpdateBuilder/AgribankV3UpdateBuilder.exe
dist/AgribankV3UpdateBuilder-portable.zip
```
