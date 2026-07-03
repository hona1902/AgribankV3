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
        Feature("Thông tin đơn vị", "Thiết lập chi nhánh và thông tin ứng dụng.", "inforcn.png"),
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
        Feature("Quyết toán tổng hợp", "Tổng hợp báo cáo quyết toán.", "qt.png"),
        Feature("Quyết toán tín dụng", "Xử lý biểu quyết toán tín dụng.", "qttd.png"),
        Feature("Quyết toán kế toán", "Xử lý biểu quyết toán kế toán.", "qtkt.png"),
        Feature("Kiểm tra báo cáo", "Kiểm tra cấu trúc và số liệu báo cáo.", "chinhta.png"),
    ],
    "Trắc nghiệm": [
        Feature("Kiểm tra nghiệp vụ", "Mở bộ câu hỏi kiểm tra nghiệp vụ.", "tracnghiem.png"),
    ],
    "Trò chơi": [
        Feature("2048", "Trò chơi 2048 trong bộ công cụ cũ.", "BamChuot.png"),
        Feature("Bấm chuột", "Trò chơi phản xạ bấm chuột.", "BamChuot.png"),
    ],
}
