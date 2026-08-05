from __future__ import annotations

from PySide6.QtWidgets import QMessageBox, QWidget

from agribank_v3.features.credit.summary.credit_limit.models import CreditLimitBatchLookupState
from agribank_v3.features.credit.summary.services import Ln01DuplicateDecision, Ln01ImportContext


def ask_ln01_duplicate_decision(parent: QWidget | None, context: Ln01ImportContext) -> Ln01DuplicateDecision:
    if not context.resolution.requires_confirmation:
        return context.resolution.default_decision

    message = QMessageBox(parent)
    message.setWindowTitle("Import LN01")
    message.setIcon(_dialog_icon(context))
    message.setText(_dialog_text(context))
    message.setInformativeText(_dialog_detail(context))

    buttons: dict[object, Ln01DuplicateDecision] = {}
    default_button = None
    for label, decision, role in _dialog_actions(context):
        button = message.addButton(label, role)
        buttons[button] = decision
        if default_button is None and decision == context.resolution.default_decision:
            default_button = button
    if default_button is not None:
        message.setDefaultButton(default_button)
    message.exec()
    return buttons.get(message.clickedButton(), Ln01DuplicateDecision.CANCEL)


def _dialog_icon(context: Ln01ImportContext) -> QMessageBox.Icon:
    if context.credit_status.different_sha:
        return QMessageBox.Icon.Warning
    if context.hmhethan_status.state == CreditLimitBatchLookupState.FOUND_INVALID:
        return QMessageBox.Icon.Warning
    return QMessageBox.Icon.Question


def _dialog_text(context: Ln01ImportContext) -> str:
    batch = context.hmhethan_status.metadata
    if context.credit_status.different_sha:
        return f"Kỳ {context.period} đã được tạo từ một file LN01 khác."
    if context.credit_status.exists and batch is not None:
        return f"Kỳ {context.period} đã có dữ liệu và file LN01 này đã có batch Hạn mức."
    if context.credit_status.exists:
        return f"Kỳ {context.period} đã có dữ liệu Tổng hợp báo cáo nhưng chưa có batch Hạn mức tương ứng."
    if batch is not None:
        return f"Kỳ {context.period} đã có dữ liệu Hạn mức tín dụng."
    if context.hmhethan_status.state == CreditLimitBatchLookupState.FOUND_INVALID:
        return "Batch Hạn mức cùng nguồn đang lỗi hoặc không hợp lệ."
    return "Xác nhận import LN01."


def _dialog_detail(context: Ln01ImportContext) -> str:
    batch = context.hmhethan_status.metadata
    if context.credit_status.different_sha:
        old_hash = _short_hash(context.credit_status.source_sha256)
        new_hash = _short_hash(context.source_sha256)
        return (
            f"File cũ: {context.credit_status.source_file_name}\n"
            f"SHA cũ: {old_hash}\n"
            f"Thời gian import cũ: {context.credit_status.imported_at}\n\n"
            f"File mới: {context.source_file_name}\n"
            f"SHA mới: {new_hash}\n\n"
            "Chỉ ghi đè sau khi bạn xác nhận rõ."
        )
    if context.credit_status.exists and batch is not None:
        return (
            "Ghi đè dữ liệu kỳ báo cáo không xóa hoặc tạo lại batch Hạn mức đã có.\n"
            "Ghi đè cả hai sẽ thay thế atomically file batch Hạn mức hiện có."
        )
    if context.credit_status.exists:
        return (
            "Chỉ tạo batch Hạn mức sẽ không thay đổi Credit.db.\n"
            "Ghi đè kỳ và tạo batch Hạn mức sẽ thay thế phần dữ liệu LN01 của kỳ."
        )
    if batch is not None:
        imported_at = batch.imported_at.strftime("%Y-%m-%d %H:%M:%S") if batch.imported_at else ""
        return (
            f"Kỳ: {batch.period or context.period}\n"
            f"Chi nhánh: {batch.branch_code or context.branch_code}\n"
            f"Batch: {batch.file_name}\n"
            f"Ngày import: {imported_at}\n\n"
            f"Kỳ {context.period} hiện chưa có trong dữ liệu Tổng hợp báo cáo.\n"
            f"Bạn có muốn tạo lại dữ liệu kỳ {context.period} từ file đang chọn không?"
        )
    if context.hmhethan_status.state == CreditLimitBatchLookupState.FOUND_INVALID:
        return (
            f"File lỗi: {context.hmhethan_status.invalid_file_path or ''}\n"
            f"Lỗi: {context.hmhethan_status.error_message}\n\n"
            "Bạn có thể tạo lại batch Hạn mức và dữ liệu kỳ từ file nguồn đang chọn."
        )
    return ""


def _dialog_actions(
    context: Ln01ImportContext,
) -> tuple[tuple[str, Ln01DuplicateDecision, QMessageBox.ButtonRole], ...]:
    if context.credit_status.different_sha:
        if context.hmhethan_status.metadata is not None:
            return (
                (
                    "Ghi đè kỳ Hạn mức và dữ liệu báo cáo",
                    Ln01DuplicateDecision.OVERWRITE_BOTH,
                    QMessageBox.ButtonRole.DestructiveRole,
                ),
                (
                    "Chỉ ghi đè dữ liệu báo cáo",
                    Ln01DuplicateDecision.OVERWRITE_CREDIT,
                    QMessageBox.ButtonRole.AcceptRole,
                ),
                ("Hủy", Ln01DuplicateDecision.CANCEL, QMessageBox.ButtonRole.RejectRole),
            )
        return (
            (
                "Ghi đè kỳ bằng file mới",
                context.resolution.default_decision,
                QMessageBox.ButtonRole.DestructiveRole,
            ),
            ("Hủy", Ln01DuplicateDecision.CANCEL, QMessageBox.ButtonRole.RejectRole),
        )
    if context.credit_status.exists and context.hmhethan_status.metadata is not None:
        return (
            (
                "Ghi đè dữ liệu kỳ báo cáo",
                Ln01DuplicateDecision.OVERWRITE_CREDIT,
                QMessageBox.ButtonRole.AcceptRole,
            ),
            (
                "Ghi đè cả dữ liệu kỳ và batch Hạn mức",
                Ln01DuplicateDecision.OVERWRITE_BOTH,
                QMessageBox.ButtonRole.DestructiveRole,
            ),
            ("Hủy", Ln01DuplicateDecision.CANCEL, QMessageBox.ButtonRole.RejectRole),
        )
    if context.credit_status.exists:
        return (
            (
                "Ghi đè kỳ và tạo batch Hạn mức",
                Ln01DuplicateDecision.CREATE_BOTH,
                QMessageBox.ButtonRole.DestructiveRole,
            ),
            (
                "Chỉ tạo batch Hạn mức",
                Ln01DuplicateDecision.CREATE_HMHE_THAN_ONLY,
                QMessageBox.ButtonRole.AcceptRole,
            ),
            ("Hủy", Ln01DuplicateDecision.CANCEL, QMessageBox.ButtonRole.RejectRole),
        )
    if context.hmhethan_status.metadata is not None:
        return (
            (
                "Tạo lại dữ liệu kỳ",
                Ln01DuplicateDecision.CREATE_CREDIT_ONLY,
                QMessageBox.ButtonRole.AcceptRole,
            ),
            ("Hủy", Ln01DuplicateDecision.CANCEL, QMessageBox.ButtonRole.RejectRole),
        )
    return (
        (
            "Tạo lại batch Hạn mức và dữ liệu kỳ",
            Ln01DuplicateDecision.CREATE_BOTH,
            QMessageBox.ButtonRole.AcceptRole,
        ),
        ("Hủy", Ln01DuplicateDecision.CANCEL, QMessageBox.ButtonRole.RejectRole),
    )


def _short_hash(value: str) -> str:
    text = str(value or "").strip()
    if len(text) <= 16:
        return text
    return f"{text[:8]}...{text[-8:]}"
