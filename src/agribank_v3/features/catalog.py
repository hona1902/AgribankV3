from __future__ import annotations

from dataclasses import dataclass


@dataclass(frozen=True, slots=True)
class Feature:
    title: str
    description: str
    icon: str


SECTIONS: dict[str, list[Feature]] = {
    "Cài đặt": [
        Feature("Kết nối Excel", "Kết nối tới phiên Excel đang hoạt động.", "caidat.png"),
        Feature("Thông tin chi nhánh", "Thiết lập chi nhánh và thông tin ứng dụng.", "inforcn.png"),
        Feature("Cơ sở dữ liệu", "Kiểm tra dữ liệu Access và cấu hình lưu trữ.", "access.png"),
        Feature("Sao lưu dữ liệu", "Sao lưu và phục hồi dữ liệu nghiệp vụ.", "file.png"),
    ],
    "Chức năng": [
        Feature("Chuyển kiểu chữ", "Đổi chữ hoa, chữ thường và viết hoa tên.", "case.png"),
        Feature("Chuyển bảng mã", "Chuyển đổi Unicode, VNI và TCVN3.", "conver.png"),
        Feature("Ngày và chuỗi", "Chuyển đổi ngày, chuỗi và định dạng dữ liệu.", "chchuoidate.png"),
        Feature("Ghép tệp Excel", "Ghép hoặc tách nhiều workbook và worksheet.", "gopsheet.png"),
        Feature("Sắp xếp dữ liệu", "Sắp xếp tiếng Việt và vùng dữ liệu.", "sort.png"),
        Feature("Bảo vệ workbook", "Quản lý bảo vệ sheet và workbook.", "protectwb.png"),
    ],
    "Dữ liệu": [
        Feature("Tìm kiếm", "Tìm kiếm và lọc dữ liệu theo nhiều điều kiện.", "file.png"),
        Feature("VLOOKUP mở rộng", "Đối chiếu dữ liệu giữa các bảng.", "m06.png"),
        Feature("Chuẩn hóa dữ liệu", "Làm sạch chuỗi, số và mã khách hàng.", "chchuoi.png"),
        Feature("Xuất báo cáo", "Xuất dữ liệu và thuộc tính workbook.", "printer.png"),
    ],
    "Tín dụng": [
        Feature("Bảng kê lãi", "Lập bảng kê lãi vay và lãi tồn nền.", "m09a.png"),
        Feature("Danh sách đến hạn", "Tạo danh sách khoản vay đến hạn.", "m09b.png"),
        Feature("Quản lý CBTD", "Quản lý danh sách cán bộ tín dụng.", "nv2.png"),
        Feature("Sao kê tín dụng", "Tạo sao kê và tổng hợp tín dụng.", "m15A.png"),
    ],
    "Kế toán": [
        Feature("Báo cáo kế toán", "Tạo báo cáo và bảng kê nghiệp vụ.", "qtkt.png"),
        Feature("Tạo file lương", "Tạo và kiểm tra file chi lương.", "ChiLuong.png"),
        Feature("Đối chiếu dữ liệu", "Đối chiếu tài khoản và số liệu kế toán.", "m20A.png"),
    ],
    "Quyết toán": [
        Feature("Quyết toán tín dụng", "Xử lý biểu quyết toán tín dụng.", "qttd.png"),
        Feature("Quyết toán kế toán", "Xử lý biểu quyết toán kế toán.", "qtkt.png"),
        Feature("Quyết toán tổng hợp", "Tổng hợp báo cáo quyết toán.", "qt.png"),
        Feature("Hướng dẫn", "Hướng dẫn tạo mẫu 30a và tổng hợp quyết toán.", "file.png"),
    ],
    "Trắc nghiệm": [
        Feature("Kiểm tra nghiệp vụ", "Mở bộ câu hỏi kiểm tra nghiệp vụ.", "tracnghiem.png"),
    ],
    "Trò chơi": [
        Feature("2048", "Trò chơi 2048 trong bộ công cụ cũ.", "BamChuot.png"),
        Feature("Bấm chuột", "Trò chơi phản xạ bấm chuột.", "BamChuot.png"),
    ],
}


QUYET_TOAN_TIN_DUNG_FEATURES: list[Feature] = [
    Feature(
        "Tạo Mẫu biểu 05/QT (From {MaCN}_rt05.csv)",
        "- Báo Cáo Kiểm Kê Hồ Sơ, Tài Sản Thế Chấp, Cầm Cố Của Khách Hàng",
        "QT/Mau_05_QT_text.svg",
    ),
    Feature(
        "Tạo Mẫu biểu 06/QT (From {MaCN}QT05.xls)",
        "- Báo Cáo Tổng Hợp Kiểm Kê Hồ Sơ, Tài Sản Thế Chấp, Cầm Cố Của Khách Hàng",
        "QT/Mau_06_QT_text.svg",
    ),
    Feature(
        "Tạo Mẫu biểu 15A/QT (From {MaCN}_rt15a.csv)",
        "- Sao Kê Chi Tiết Tài Khoản Cho Vay Khách Hàng Là Tổ Chức, Hộ Kinh Doanh Và Cá Nhân",
        "QT/Mau_15a_QT_text.svg",
    ),
    Feature(
        "Tạo Mẫu biểu 15B/QT (From {MaCN}_rt15b.csv)",
        "- Sao Kê Chi Tiết Tài Khoản Cho Vay Khách Hàng Là Tổ Chức, Hộ Kinh Doanh Và Cá Nhân "
        "Có Lãi Phải Thu Hạch Toán Ngoại Bảng (Tài Khoản 941)",
        "QT/Mau_15b_QT_text.svg",
    ),
    Feature(
        "Tạo Mẫu biểu 16/QT (From {MaCN}_rt16.csv)",
        "- Sao Kê Chi Tiết Khách Hàng Có Dư Nợ Từ 10 Tỷ Việt Nam Đồng Trở Lên",
        "QT/Mau_16_QT_text.svg",
    ),
    Feature(
        "Tạo Mẫu biểu 18/QT (From {MaCN}_rt18.csv)",
        "- Sao Kê Chi Tiết Tài Khoản 92, 93 : Bảo Lãnh Và Thư Tín Dụng Cho Khách Hàng",
        "QT/Mau_18_QT_text.svg",
    ),
    Feature(
        "Tạo Mẫu biểu 20a/QT (From {MaCN}_rt20.xls from lnlr20)",
        "- Báo Cáo Nợ Được Xử Lý Bằng Nguồn Dự Phòng",
        "QT/Mau_20a_QT_text.svg",
    ),
    Feature(
        "Tạo Mẫu biểu 30a/QT",
        "- Tạo mẫu biểu 30a Quyết toán năm",
        "qt.png",
    ),
]


QUYET_TOAN_KE_TOAN_FEATURES: list[Feature] = [
    Feature(
        "Tạo Mẫu biểu 04/QT (From IC_100435)",
        "- Báo Cáo Tình Hình Sử Dụng Ấn Chỉ Quan Trọng (From IC_100435)",
        "QT/Mau_04_QT_text.svg",
    ),
    Feature(
        "Tạo Mẫu biểu 07a/QT (From WT_100642)",
        "- Báo Cáo Kiểm Kê Công Cụ Dụng Cụ (From WT_100642)",
        "QT/Mau_07a_QT_text.svg",
    ),
    Feature(
        "Tạo Mẫu biểu 08/QT (From FA_100586)",
        "- Báo Cáo Kiểm Kê Tài Sản Cố Định (From FA_100586)",
        "QT/Mau_08_QT_text.svg",
    ),
    Feature(
        "Tạo Mẫu biểu 09a/QT (Màn hình mshr32 - xuất Excel báo cáo TMBCTC_TSCD001)",
        "- Báo Cáo Tình Hình Tăng, Giảm TSCĐ Hữu Hình",
        "QT/Mau_09a_QT_text.svg",
    ),
    Feature(
        "Tạo Mẫu biểu 09b/QT (Màn hình mshr32 - xuất Excel báo cáo TMBCTC_TSCD002)",
        "- Báo Cáo Tình Hình Tăng, Giảm TSCĐ Vô Hình",
        "QT/Mau_09b_QT_text.svg",
    ),
    Feature(
        "Tạo Mẫu biểu 09c/QT (Màn hình mshr32 - xuất Excel báo cáo TMBCTC_TSCD003)",
        "- Báo Cáo Tình Hình Tăng, Giảm TSCĐ Thuê Tài Chính",
        "QT/Mau_09c_QT_text.svg",
    ),
    Feature(
        "Tạo Mẫu biểu 13/QT (From {MaCN}_rt13.csv)",
        "- Sao Kê Chi Tiết Tiền Gửi Khách Hàng "
        "(Tiền Gửi Thanh Toán, Có Kỳ Hạn, Tiết Kiệm, Kỳ Phiếu, Trái Phiếu)",
        "QT/Mau_13_QT_text.svg",
    ),
    Feature(
        "Tạo Mẫu biểu 14/QT (From {MaCN}_rt14.csv)",
        "- Sao Kê Chi Tiết Số Dư Tiền Gửi, Tiết Kiệm, Kỳ Phiếu, Trái Phiếu "
        "Của Khách Hàng Có Số Dư Từ 10 Tỷ VNĐ Trở Lên",
        "QT/Mau_14_QT_text.svg",
    ),
    Feature(
        "Tạo Mẫu biểu 22/QT (From GL_glst34)",
        "- Sao Kê Chi Tiết Số Dư Tài Khoản Doanh Thu Và Chi Phí Chờ Phân Bổ "
        "Đến Ngày Quyết Toán",
        "QT/Mau_22_QT_text.svg",
    ),
    Feature(
        "Tạo Mẫu biểu 23/QT (From GL_glcb06)",
        "- Sao Kê Chi Tiết Tài Khoản Thu Nhập Và Chi Phí Bất Thường",
        "QT/Mau_23_QT_text.svg",
    ),
    Feature(
        "Tạo Mẫu biểu 24/QT (From {MaCN}_rt24.csv)",
        "- Sao Kê Chi Tiết Tài Khoản Phải Thu, Phải Trả",
        "QT/Mau_24_QT_text.svg",
    ),
    Feature(
        "Tạo Mẫu biểu 30a/QT",
        "- Tạo mẫu biểu 30a Quyết toán năm",
        "qt.png",
    ),
]


QUYET_TOAN_TONG_HOP_FEATURES: list[Feature] = [
    Feature(
        "Tổng hợp Mẫu biểu 05/QT (From {MaCN}_rt05.csv)",
        "- Tạo file Tổng hợp quyết toán mẫu 05 - TSĐB",
        "QT/Mau_05_QT_text.svg",
    ),
    Feature(
        "Tạo Mẫu biểu 06/QT (From {MaCN}QT05.xls)",
        "- Báo Cáo Tổng Hợp Kiểm Kê Hồ Sơ, Tài Sản Thế Chấp, Cầm Cố Của Khách Hàng",
        "QT/Mau_06_QT_text.svg",
    ),
    Feature(
        "Tổng hợp Mẫu biểu 13/QT (From {MaCN}_rt13.csv)",
        "- Tạo file Tổng hợp quyết toán mẫu 13 - Tiền gửi",
        "QT/Mau_13_QT_text.svg",
    ),
    Feature(
        "Tổng hợp Mẫu biểu 14/QT (From {MaCN}_rt14.csv)",
        "- Tạo file Tổng hợp quyết toán mẫu 14 - Tiền gửi (Từ 10 tỷ)",
        "QT/Mau_14_QT_text.svg",
    ),
    Feature(
        "Tổng hợp Mẫu biểu 15a/QT (From {MaCN}_rt15a.csv)",
        "- Tạo file Tổng hợp quyết toán mẫu 15a - Tín dụng",
        "QT/Mau_15a_QT_text.svg",
    ),
    Feature(
        "Tổng hợp Mẫu biểu 15b/QT (From {MaCN}_rt15b.csv)",
        "- Tạo file Tổng hợp quyết toán mẫu 15b - Tín dụng",
        "QT/Mau_15b_QT_text.svg",
    ),
    Feature(
        "Tổng hợp Mẫu biểu 16/QT (From {MaCN}_rt16.csv)",
        "- Tạo file Tổng hợp quyết toán mẫu 16 - Tín dụng (Từ 10 tỷ)",
        "QT/Mau_16_QT_text.svg",
    ),
    Feature(
        "Tổng hợp Mẫu biểu 18/QT (From {MaCN}_rt18.csv)",
        "- Tạo file Tổng hợp quyết toán mẫu 18 - Bảo lãnh và thư tín dụng",
        "QT/Mau_18_QT_text.svg",
    ),
    Feature(
        "Tạo Mẫu biểu 30a/QT - CN loại I",
        "- Tạo mẫu biểu 30a Quyết toán năm - Chi nhánh Loại I",
        "qt.png",
    ),
    Feature(
        "Hướng dẫn tổng hợp số liệu quyết toán",
        "- Hướng dẫn tổng hợp số liệu quyết toán",
        "file.png",
    ),
]
