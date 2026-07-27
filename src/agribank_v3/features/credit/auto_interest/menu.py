from __future__ import annotations

from agribank_v3.features.catalog import Feature


AUTO_INTEREST_TITLE = "Thu lãi bán tự động"
CREATE_INTEREST_FILE_TITLE = "Tạo file thu lãi"
REPORT_FOLDER_SETTINGS_TITLE = "Cài đặt thư mục báo cáo"
CREATE_REPORT_FILE_TITLE = "Tạo file báo cáo"


AUTO_INTEREST_FEATURES: tuple[Feature, ...] = (
    Feature(
        CREATE_INTEREST_FILE_TITLE,
        "Tạo file thu lãi bán tự động từ sao kê lãi dự kiến và sao kê tiền gửi msit/dpda.",
        "m09a.png",
    ),
    Feature(
        REPORT_FOLDER_SETTINGS_TITLE,
        "Thiết lập thư mục lưu file báo cáo thu nợ bán tự động.",
        "caidat.png",
    ),
    Feature(
        CREATE_REPORT_FILE_TITLE,
        "Tạo file báo cáo thu nợ bán tự động để cập nhật lên chương trình thống kê.",
        "file.png",
    ),
)
