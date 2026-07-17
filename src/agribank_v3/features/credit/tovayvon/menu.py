from __future__ import annotations

from agribank_v3.features.catalog import Feature


TOVAYVON_FEATURES: tuple[Feature, ...] = (
    Feature(
        "Quản lý tổ vay vốn",
        "Quản lý danh sách tổ vay vốn, import/export dữ liệu và cấu hình hoa hồng.",
        "tovayvon.svg",
    ),
    Feature(
        "Bảng kê thu lãi tổ vay vốn",
        "Tạo bảng kê thu lãi tổ vay vốn từ dữ liệu khoản vay/IPCAS.",
        "m09a.png",
    ),
    Feature(
        "Đề nghị thanh toán hoa hồng tổ vay vốn",
        "Xuất giấy đề nghị thanh toán hoa hồng hoặc chi phí liên quan đến tổ vay vốn.",
        "file.png",
    ),
    Feature(
        "Chỉnh sửa mẫu biểu Đề nghị thanh toán",
        "Mở file mẫu DeNghiThanhToan.docx để chỉnh sửa biểu mẫu đề nghị thanh toán hoa hồng tổ vay vốn.",
        "file.png",
    ),
    Feature(
        "Đối chiếu dư nợ theo tổ vay vốn",
        "Đối chiếu dư nợ, mã nhóm vay, số tiền lãi và thông tin khoản vay theo tổ.",
        "m09b.png",
    ),
    Feature(
        "Hướng dẫn tổ vay vốn",
        "Hướng dẫn tạo bảng kê thu lãi, đề nghị thanh toán và đối chiếu dư nợ tổ vay vốn.",
        "file.png",
    ),
)
