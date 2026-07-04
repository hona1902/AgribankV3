from __future__ import annotations

from datetime import datetime
from html import escape
from importlib.resources import files
from pathlib import Path
import random
import shutil
import sqlite3
import time

from PySide6.QtCore import QEvent, Qt, QTimer, Signal
from PySide6.QtGui import QColor, QIcon, QPixmap
from PySide6.QtWidgets import (
    QButtonGroup,
    QCheckBox,
    QComboBox,
    QDialog,
    QFormLayout,
    QFrame,
    QFileDialog,
    QGridLayout,
    QGraphicsDropShadowEffect,
    QGroupBox,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QMessageBox,
    QProgressBar,
    QPushButton,
    QRadioButton,
    QScrollArea,
    QSizePolicy,
    QSplitter,
    QSpinBox,
    QStackedWidget,
    QTabWidget,
    QTableWidget,
    QTableWidgetItem,
    QTextEdit,
    QTreeWidget,
    QTreeWidgetItem,
    QVBoxLayout,
    QWidget,
)

from agribank_v3.quiz import (
    Question,
    QuizDatabase,
    QuizDatabaseError,
    QuizSession,
    export_questions_to_excel,
    import_questions_from_excel,
)


class ExpandingTextEdit(QTextEdit):
    """A text editor that grows with wrapped content inside a page scroll area."""

    def __init__(
        self,
        minimum_height: int,
        maximum_height: int,
        parent: QWidget | None = None,
    ) -> None:
        super().__init__(parent)
        self._minimum_content_height = minimum_height
        self._maximum_content_height = maximum_height
        self.document().documentLayout().documentSizeChanged.connect(
            self._fit_to_document
        )
        self.setAcceptRichText(False)
        self._fit_to_document()

    def _fit_to_document(self, *_: object) -> None:
        content_height = int(self.document().size().height()) + 14
        target_height = max(
            self._minimum_content_height,
            min(self._maximum_content_height, content_height),
        )
        if self.height() != target_height:
            self.setFixedHeight(target_height)
        self.setVerticalScrollBarPolicy(
            Qt.ScrollBarPolicy.ScrollBarAsNeeded
            if content_height > self._maximum_content_height
            else Qt.ScrollBarPolicy.ScrollBarAlwaysOff
        )

    def resizeEvent(self, event: object) -> None:
        super().resizeEvent(event)
        QTimer.singleShot(0, self._fit_to_document)


class AnswerLookupResultItem(QFrame):
    selected = Signal(int)
    double_clicked = Signal(int)

    def __init__(
        self,
        row: int,
        question: Question,
        correct_text: str,
        parent: QWidget | None = None,
    ) -> None:
        super().__init__(parent)
        self.row = row
        self.question_id = question.id
        self.setObjectName("AnswerResultItem")
        self.setProperty("selected", False)
        self.setCursor(Qt.CursorShape.PointingHandCursor)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(16, 14, 16, 14)
        layout.setSpacing(12)

        question_row = QHBoxLayout()
        question_row.setSpacing(16)
        question_chip = QLabel("●  Câu hỏi")
        question_chip.setObjectName("AnswerLookupChip")
        question_chip.setFixedWidth(132)
        question_text = QLabel(question.text)
        question_text.setObjectName("AnswerLookupQuestion")
        question_text.setWordWrap(True)
        question_text.setTextInteractionFlags(
            Qt.TextInteractionFlag.TextSelectableByMouse
        )
        question_row.addWidget(question_chip, alignment=Qt.AlignmentFlag.AlignTop)
        question_row.addWidget(question_text, stretch=1)
        layout.addLayout(question_row)

        divider = QFrame()
        divider.setObjectName("AnswerLookupDivider")
        divider.setFrameShape(QFrame.Shape.HLine)
        layout.addWidget(divider)

        answer_row = QHBoxLayout()
        answer_row.setSpacing(16)
        answer_chip = QLabel("✓  Đáp án đúng")
        answer_chip.setObjectName("AnswerLookupChip")
        answer_chip.setFixedWidth(132)
        answer_text = QLabel(
            f"Đáp án đúng {question.correct_answer}: {correct_text}"
        )
        answer_text.setObjectName("AnswerLookupAnswer")
        answer_text.setWordWrap(True)
        answer_text.setTextInteractionFlags(
            Qt.TextInteractionFlag.TextSelectableByMouse
        )
        answer_row.addWidget(answer_chip, alignment=Qt.AlignmentFlag.AlignTop)
        answer_row.addWidget(answer_text, stretch=1)
        layout.addLayout(answer_row)

        for child in (
            question_chip,
            question_text,
            divider,
            answer_chip,
            answer_text,
        ):
            child.setCursor(Qt.CursorShape.PointingHandCursor)
            child.installEventFilter(self)

    def set_selected(self, selected: bool) -> None:
        self.setProperty("selected", selected)
        self.style().unpolish(self)
        self.style().polish(self)

    def eventFilter(self, watched: object, event: object) -> bool:
        if isinstance(event, QEvent):
            if event.type() == QEvent.Type.MouseButtonPress:
                self.selected.emit(self.row)
                return False
            if event.type() == QEvent.Type.MouseButtonDblClick:
                self.selected.emit(self.row)
                self.double_clicked.emit(self.row)
                return True
        return super().eventFilter(watched, event)

    def mousePressEvent(self, event: object) -> None:
        self.selected.emit(self.row)
        super().mousePressEvent(event)

    def mouseDoubleClickEvent(self, event: object) -> None:
        self.double_clicked.emit(self.row)
        super().mouseDoubleClickEvent(event)


class QuestionReviewDialog(QDialog):
    def __init__(
        self,
        question: Question,
        selected_answer: str,
        display_number: int,
        parent: QWidget | None = None,
    ) -> None:
        super().__init__(parent)
        self.setWindowTitle(f"Đối chiếu câu hỏi {display_number}")
        self.setModal(True)
        self.setMinimumSize(760, 560)
        self.resize(900, 650)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(24, 22, 24, 22)
        layout.setSpacing(14)

        title = QLabel(f"Câu {display_number}")
        title.setObjectName("PageTitle")
        context = QLabel(
            f"Nghiệp vụ: {question.topic_name}  •  "
            f"Chuyên đề: {question.subject_name or 'Chưa phân loại'}"
        )
        context.setObjectName("MutedText")
        prompt = QLabel(question.text)
        prompt.setObjectName("QuizQuestion")
        prompt.setWordWrap(True)
        prompt.setTextInteractionFlags(
            Qt.TextInteractionFlag.TextSelectableByMouse
        )
        layout.addWidget(title)
        layout.addWidget(context)
        layout.addWidget(prompt)

        for letter, answer_text in question.options.items():
            option = QFrame()
            option_layout = QHBoxLayout(option)
            option_layout.setContentsMargins(14, 11, 14, 11)
            option_layout.setSpacing(12)
            badge = QLabel(letter)
            badge.setAlignment(Qt.AlignmentFlag.AlignCenter)
            badge.setFixedSize(30, 30)
            answer = QLabel(answer_text)
            answer.setWordWrap(True)
            status = QLabel("")
            status.setStyleSheet("font-weight: 700;")
            if letter == question.correct_answer:
                option.setStyleSheet(
                    "QFrame { background: #e8f8ee; border: 1px solid #38b66b; "
                    "border-radius: 9px; }"
                )
                badge.setStyleSheet(
                    "background: #23864c; color: white; border-radius: 15px; "
                    "font-weight: 700;"
                )
                status.setText(
                    "✓ Đáp án đúng · Bạn đã chọn"
                    if selected_answer == letter
                    else "✓ Đáp án đúng"
                )
                status.setStyleSheet("color: #17683a; font-weight: 700;")
            elif selected_answer == letter:
                option.setStyleSheet(
                    "QFrame { background: #fff0f0; border: 1px solid #e45b5b; "
                    "border-radius: 9px; }"
                )
                badge.setStyleSheet(
                    "background: #c93636; color: white; border-radius: 15px; "
                    "font-weight: 700;"
                )
                status.setText("✗ Đáp án bạn đã chọn")
                status.setStyleSheet("color: #a92929; font-weight: 700;")
            else:
                option.setStyleSheet(
                    "QFrame { background: white; border: 1px solid #dce3e8; "
                    "border-radius: 9px; }"
                )
                badge.setStyleSheet(
                    "background: #f0f3f6; color: #46515b; border-radius: 15px; "
                    "font-weight: 700;"
                )
                status.hide()
            option_layout.addWidget(badge)
            option_layout.addWidget(answer, stretch=1)
            option_layout.addWidget(status)
            layout.addWidget(option)

        if not selected_answer:
            unanswered = QLabel("Bạn chưa trả lời câu hỏi này.")
            unanswered.setStyleSheet("color: #9b2c2c; font-weight: 700;")
            layout.addWidget(unanswered)
        if question.source_reference:
            source = QLabel(f"Nguồn tham khảo: {question.source_reference}")
            source.setWordWrap(True)
            source.setStyleSheet(
                "background: #fff9df; color: #5f4b12; "
                "border: 1px solid #efd26a; border-radius: 8px; padding: 10px;"
            )
            layout.addWidget(source)

        layout.addStretch()
        close_button = QPushButton("Đóng")
        close_button.setObjectName("PrimaryButton")
        close_button.clicked.connect(self.accept)
        layout.addWidget(
            close_button,
            alignment=Qt.AlignmentFlag.AlignRight,
        )


class QuizWidget(QWidget):
    close_requested = Signal()

    def __init__(
        self,
        parent: QWidget | None = None,
        database: QuizDatabase | None = None,
    ) -> None:
        super().__init__(parent)

        self.database = database or QuizDatabase()
        self.sync_result = self.database.sync_from_access()
        self.session: QuizSession | None = None
        self.deadline = 0.0
        self.answer_letters: list[str] = []

        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        self.pages = QTabWidget()
        self.pages.setDocumentMode(True)
        self.pages.addTab(self._build_setup_page(), "Cài đặt đề cương")
        self.pages.addTab(self._build_quiz_page(), "Trả Lời Câu Hỏi")
        self.pages.addTab(self._build_result_page(), "Kết Quả Trả Lời")
        self.pages.addTab(self._build_answer_lookup_page(), "Tra Cứu Đáp Án")
        self.pages.addTab(self._build_data_management_page(), "Quản Lý Dữ Liệu")
        self.pages.setTabEnabled(1, False)
        self.pages.setTabEnabled(2, False)
        layout.addWidget(self.pages)

        self.timer = QTimer(self)
        self.timer.setInterval(1000)
        self.timer.timeout.connect(self.update_timer)
        self.refresh_setup_data()

    @staticmethod
    def _apply_card_shadow(widget: QWidget) -> None:
        shadow = QGraphicsDropShadowEffect(widget)
        shadow.setBlurRadius(22)
        shadow.setOffset(0, 3)
        shadow.setColor(QColor(45, 55, 65, 35))
        widget.setGraphicsEffect(shadow)

    @staticmethod
    def _set_button_icon(button: QPushButton, icon_name: str) -> None:
        button.setIcon(
            QIcon(
                str(
                    files("agribank_v3").joinpath(
                        "resources",
                        "icons",
                        icon_name,
                    )
                )
            )
        )

    @staticmethod
    def _section_header(
        icon_name: str,
        title: str,
        icon_size: int = 36,
    ) -> QWidget:
        header = QWidget()
        header.setObjectName("SectionHeader")
        layout = QHBoxLayout(header)
        layout.setContentsMargins(0, 0, 0, 4)
        layout.setSpacing(10)
        icon = QLabel()
        icon.setFixedSize(icon_size, icon_size)
        pixmap = QPixmap(
            str(
                files("agribank_v3").joinpath(
                    "resources",
                    "icons",
                    icon_name,
                )
            )
        )
        icon.setPixmap(
            pixmap.scaled(
                icon_size,
                icon_size,
                Qt.AspectRatioMode.KeepAspectRatio,
                Qt.TransformationMode.SmoothTransformation,
            )
        )
        label = QLabel(title)
        label.setObjectName("SectionTitle")
        layout.addWidget(icon)
        layout.addWidget(label)
        layout.addStretch()
        return header

    @classmethod
    def _setup_option_group(
        cls,
        icon_name: str,
        title: str,
    ) -> tuple[QGroupBox, QVBoxLayout]:
        group = QGroupBox()
        group.setObjectName("SetupGroup")
        root_layout = QVBoxLayout(group)
        root_layout.setContentsMargins(10, 8, 10, 8)
        root_layout.setSpacing(2)
        root_layout.addWidget(
            cls._section_header(icon_name, title, icon_size=30)
        )
        content_layout = QVBoxLayout()
        content_layout.setContentsMargins(38, 0, 0, 0)
        content_layout.setSpacing(3)
        root_layout.addLayout(content_layout)
        return group, content_layout

    def _build_setup_page(self) -> QWidget:
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(18, 16, 18, 16)
        layout.setSpacing(12)

        title = QLabel("ỨNG DỤNG ÔN TẬP KIẾN THỨC NGHIỆP VỤ")
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        title.setStyleSheet("color: #9d1d49; font-size: 23px; font-weight: 800;")
        layout.addWidget(title)

        card = QFrame()
        card_layout = QGridLayout(card)
        card_layout.setContentsMargins(6, 4, 6, 6)
        card_layout.setHorizontalSpacing(12)

        left_card = QFrame()
        left_card.setObjectName("WelcomeCard")
        left_card.setSizePolicy(
            QSizePolicy.Policy.Ignored,
            QSizePolicy.Policy.Expanding,
        )
        self._apply_card_shadow(left_card)
        left_layout = QGridLayout(left_card)
        left_layout.setContentsMargins(12, 10, 12, 10)
        left_layout.setHorizontalSpacing(10)
        left_layout.setVerticalSpacing(6)

        self.employee_combo = QComboBox()
        self.employee_combo.setEditable(True)
        self.employee_combo.setMinimumHeight(36)
        self.employee_combo.currentTextChanged.connect(self.update_topic_summary)
        left_layout.addWidget(QLabel("Họ tên nhân viên (Nếu có):"), 0, 0)
        left_layout.addWidget(self.employee_combo, 0, 1)

        mode_group, mode_layout = self._setup_option_group(
            "option_questions.svg",
            "Chọn nhóm câu hỏi trả lời",
        )
        self.topic_mode = QRadioButton(
            "Trả lời theo bộ câu hỏi của cả 1 chuyên đề"
        )
        self.random_mode = QRadioButton(
            "Tạo một đề kiểm tra ngẫu nhiên theo nghiệp vụ"
        )
        self.topic_mode.setChecked(True)
        mode_layout.addWidget(self.topic_mode)
        mode_layout.addWidget(self.random_mode)
        left_layout.addWidget(mode_group, 1, 0, 1, 2)

        time_group, time_layout = self._setup_option_group(
            "option_time.svg",
            "Chọn thời gian trả lời",
        )
        self.unlimited_time = QRadioButton("Không giới hạn thời gian trả lời")
        self.limited_time = QRadioButton("Giới hạn thời gian trả lời")
        self.unlimited_time.setChecked(True)
        time_row = QHBoxLayout()
        time_row.addWidget(self.limited_time)
        self.time_limit = QSpinBox()
        self.time_limit.setRange(1, 180)
        self.time_limit.setValue(60)
        self.time_limit.setSuffix(" phút")
        self.time_limit.setButtonSymbols(QSpinBox.ButtonSymbols.NoButtons)
        self.time_limit.setEnabled(False)
        self.time_limit.valueChanged.connect(self.update_topic_summary)
        self.limited_time.toggled.connect(self.time_limit.setEnabled)
        self.limited_time.toggled.connect(self.update_topic_summary)
        time_row.addWidget(self.time_limit)
        time_row.addStretch()
        time_layout.addWidget(self.unlimited_time)
        time_layout.addLayout(time_row)
        left_layout.addWidget(time_group, 2, 0, 1, 2)

        answer_group, answer_layout = self._setup_option_group(
            "option_answers.svg",
            "Chọn cách thể hiện đáp án",
        )
        self.source_answer_order = QRadioButton(
            "Thể hiện thứ tự đáp án theo đề cương"
        )
        self.shuffle_answer_order = QRadioButton(
            "Hoán đổi thứ tự đáp án so với đề cương"
        )
        self.source_answer_order.setChecked(True)
        self.source_answer_order.toggled.connect(self.update_topic_summary)
        self.shuffle_answer_order.setToolTip(
            "Chỉ hoán đổi khi hiển thị; thứ tự A/B/C/D gốc trong database được giữ nguyên."
        )
        answer_layout.addWidget(self.source_answer_order)
        answer_layout.addWidget(self.shuffle_answer_order)
        left_layout.addWidget(answer_group, 3, 0, 1, 2)

        audience_group, audience_layout = self._setup_option_group(
            "option_audience.svg",
            "Chọn đối tượng thi",
        )
        self.manager_audience = QRadioButton("Lãnh đạo quản lý")
        self.staff_audience = QRadioButton("Lao động chuyên môn nghiệp vụ")
        self.manager_audience.setChecked(True)
        audience_layout.addWidget(self.manager_audience)
        audience_layout.addWidget(self.staff_audience)
        left_layout.addWidget(audience_group, 4, 0, 1, 2)

        self.mode_pages = QStackedWidget()
        self.mode_pages.setSizePolicy(
            QSizePolicy.Policy.Expanding,
            QSizePolicy.Policy.Expanding,
        )
        self.mode_pages.addWidget(self._build_topic_setup())
        self.mode_pages.addWidget(self._build_professional_setup())
        card_layout.addWidget(left_card, 0, 0)
        card_layout.addWidget(self.mode_pages, 0, 1)
        card_layout.setColumnStretch(0, 9)
        card_layout.setColumnStretch(1, 11)
        self.topic_mode.toggled.connect(
            lambda checked: self.mode_pages.setCurrentIndex(0 if checked else 1)
        )

        self.data_status = QLabel("")
        self.data_status.setObjectName("MutedText")
        self.data_status.setWordWrap(True)
        layout.addWidget(card, stretch=1)
        layout.addWidget(self.data_status)

        actions = QHBoxLayout()
        about_button = QPushButton("Quản lý dữ liệu")
        about_button.setObjectName("SecondaryButton")
        self._set_button_icon(about_button, "button_data.svg")
        about_button.clicked.connect(lambda: self.pages.setCurrentIndex(3))
        actions.addWidget(about_button)

        export_button = QPushButton("Xuất Sang Excel")
        export_button.setObjectName("SecondaryButton")
        self._set_button_icon(export_button, "button_excel.svg")
        export_button.setToolTip(
            "Xuất theo chuyên đề hoặc hạn mức đề ngẫu nhiên đang chọn."
        )
        export_button.clicked.connect(self.export_quiz_excel)
        actions.addWidget(export_button)

        self.sync_button = QPushButton("Cập nhật từ Access")
        self.sync_button.setObjectName("SecondaryButton")
        self._set_button_icon(self.sync_button, "button_sync.svg")
        self.sync_button.clicked.connect(self.force_sync)
        actions.addWidget(self.sync_button)
        actions.addStretch()

        close_button = QPushButton("×  Thoát")
        close_button.setObjectName("SecondaryButton")
        close_button.clicked.connect(self.request_close)
        actions.addWidget(close_button)

        start_button = QPushButton("▶  Bắt Đầu Trả Lời")
        start_button.setObjectName("PrimaryButton")
        start_button.clicked.connect(self.start_quiz)
        actions.addWidget(start_button)
        layout.addLayout(actions)
        return page

    def _build_topic_setup(self) -> QWidget:
        page = QGroupBox()
        page.setObjectName("RightSetupCard")
        self._apply_card_shadow(page)
        layout = QVBoxLayout(page)
        layout.setContentsMargins(14, 12, 14, 14)
        layout.setSpacing(8)
        layout.addWidget(
            self._section_header(
                "section_clipboard.svg",
                "Tóm tắt lựa chọn",
            )
        )
        topic_caption = QLabel("Chuyên đề để kiểm tra:")
        topic_caption.setStyleSheet("font-weight: 700; color: #26313b;")
        layout.addWidget(topic_caption)
        self.topic_combo = QComboBox()
        self.topic_combo.setMinimumHeight(36)
        self.topic_combo.currentIndexChanged.connect(self.update_topic_summary)
        layout.addWidget(self.topic_combo)
        selected_caption = QLabel("Bạn đã chọn:")
        selected_caption.setStyleSheet(
            "font-weight: 700; color: #26313b; margin-top: 8px;"
        )
        layout.addWidget(selected_caption)
        self.topic_summary = QLabel("")
        self.topic_summary.setWordWrap(True)
        self.topic_summary.setTextFormat(Qt.TextFormat.RichText)
        self.topic_summary.setAlignment(
            Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignLeft
        )
        layout.addWidget(self.topic_summary, stretch=1)
        return page

    def _build_professional_setup(self) -> QWidget:
        page = QGroupBox()
        page.setObjectName("RightSetupCard")
        self._apply_card_shadow(page)
        layout = QVBoxLayout(page)
        layout.setContentsMargins(14, 12, 14, 12)
        layout.setSpacing(8)
        layout.addWidget(
            self._section_header(
                "section_list.svg",
                "Tạo đề từ các chuyên đề",
            )
        )
        intro = QLabel(
            "Chọn từng chuyên đề và số câu cần lấy. Cấu hình được tự động lưu "
            "sau lần tạo đề đầu tiên."
        )
        intro.setObjectName("MutedText")
        intro.setWordWrap(True)
        layout.addWidget(intro)
        self.quota_table = QTreeWidget()
        self.quota_table.setObjectName("QuotaTree")
        self.quota_table.setColumnCount(4)
        self.quota_table.setHeaderLabels(
            ["Chọn", "Nghiệp vụ", "Hiện có", "Số câu lấy"]
        )
        self.quota_table.setRootIsDecorated(True)
        self.quota_table.setAlternatingRowColors(False)
        self.quota_table.setSelectionMode(QTreeWidget.SelectionMode.NoSelection)
        self.quota_table.setFocusPolicy(Qt.FocusPolicy.NoFocus)
        self.quota_table.setHorizontalScrollBarPolicy(
            Qt.ScrollBarPolicy.ScrollBarAlwaysOff
        )
        self.quota_table.setIndentation(12)
        quota_header = self.quota_table.header()
        quota_header.setStretchLastSection(False)
        quota_header.setSectionResizeMode(0, QHeaderView.ResizeMode.Fixed)
        quota_header.setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        quota_header.setSectionResizeMode(2, QHeaderView.ResizeMode.Fixed)
        quota_header.setSectionResizeMode(3, QHeaderView.ResizeMode.Fixed)
        self.quota_table.setColumnWidth(0, 64)
        self.quota_table.setColumnWidth(2, 78)
        self.quota_table.setColumnWidth(3, 132)
        self.quota_table.itemChanged.connect(self._quota_item_changed)
        layout.addWidget(self.quota_table, stretch=1)
        self.total_quota_label = QLabel("")
        self.total_quota_label.setContentsMargins(4, 6, 0, 0)
        self.total_quota_label.setStyleSheet(
            "color: #65142f; font-weight: 800; font-size: 14px;"
        )
        layout.addWidget(self.total_quota_label)
        self.update_quota_total()
        return page

    def _build_quiz_page(self) -> QWidget:
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(30, 22, 30, 24)
        layout.setSpacing(14)

        toolbar = QHBoxLayout()
        self.question_progress_label = QLabel("Câu 0/0")
        self.question_progress_label.setObjectName("CardTitle")
        self.timer_label = QLabel("Không giới hạn")
        self.timer_label.setStyleSheet("color: #831f41; font-weight: 700;")
        jump_label = QLabel("Đi đến câu:")
        jump_label.setStyleSheet("color: #4c5862; font-weight: 600;")
        self.jump_question_spin = QSpinBox()
        self.jump_question_spin.setRange(1, 1)
        self.jump_question_spin.setButtonSymbols(
            QSpinBox.ButtonSymbols.NoButtons
        )
        self.jump_question_spin.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.jump_question_spin.setFixedWidth(62)
        self.jump_question_spin.lineEdit().returnPressed.connect(
            self.jump_to_question
        )
        jump_button = QPushButton("Đi")
        jump_button.setObjectName("SecondaryButton")
        jump_button.setToolTip("Chuyển tới số câu đã nhập")
        jump_button.clicked.connect(self.jump_to_question)
        finish_button = QPushButton("Kết thúc")
        finish_button.setObjectName("SecondaryButton")
        finish_button.clicked.connect(self.confirm_finish)
        toolbar.addWidget(self.question_progress_label)
        toolbar.addStretch()
        toolbar.addWidget(self.timer_label)
        toolbar.addSpacing(12)
        toolbar.addWidget(jump_label)
        toolbar.addWidget(self.jump_question_spin)
        toolbar.addWidget(jump_button)
        toolbar.addSpacing(8)
        toolbar.addWidget(finish_button)
        layout.addLayout(toolbar)

        self.answer_progress = QProgressBar()
        self.answer_progress.setTextVisible(True)
        self.answer_progress.setMinimumHeight(22)
        layout.addWidget(self.answer_progress)

        question_card = QFrame()
        question_card.setObjectName("WelcomeCard")
        question_layout = QVBoxLayout(question_card)
        question_layout.setContentsMargins(24, 20, 24, 20)
        question_layout.setSpacing(12)

        self.topic_label = QLabel("")
        self.topic_label.setObjectName("MutedText")
        self.topic_label.setSizePolicy(
            QSizePolicy.Policy.Expanding,
            QSizePolicy.Policy.Maximum,
        )
        question_header = QHBoxLayout()
        question_header.setSpacing(12)
        question_header.setAlignment(Qt.AlignmentFlag.AlignTop)
        self.question_number_badge = QLabel("1")
        self.question_number_badge.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.question_number_badge.setFixedSize(34, 34)
        self.question_number_badge.setStyleSheet(
            "background: #eaf2ff; color: #174ea6; border: 1px solid #c8dcff; "
            "border-radius: 17px; font-size: 14px; font-weight: 700;"
        )
        self.question_label = QLabel("")
        self.question_label.setObjectName("QuizQuestion")
        self.question_label.setWordWrap(True)
        self.question_label.setSizePolicy(
            QSizePolicy.Policy.Expanding,
            QSizePolicy.Policy.Maximum,
        )
        self.question_label.setTextInteractionFlags(
            Qt.TextInteractionFlag.TextSelectableByMouse
        )
        question_layout.addWidget(self.topic_label)
        question_header.addWidget(self.question_number_badge)
        question_header.addWidget(self.question_label, stretch=1)
        question_layout.addLayout(question_header)

        self.answer_group = QButtonGroup(self)
        self.answer_group.buttonClicked.connect(self.answer_selected)
        self.answer_buttons: list[QRadioButton] = []
        for index in range(4):
            button = QRadioButton()
            button.setObjectName("QuizOption")
            button.setMinimumHeight(52)
            button.setSizePolicy(
                QSizePolicy.Policy.Expanding,
                QSizePolicy.Policy.Fixed,
            )
            button.setProperty("answerIndex", index)
            self.answer_group.addButton(button, index)
            self.answer_buttons.append(button)
            question_layout.addWidget(button)

        self.feedback_label = QLabel("")
        self.feedback_label.setWordWrap(True)
        self.feedback_label.hide()
        question_layout.addWidget(self.feedback_label)

        self.hint_frame = QFrame()
        self.hint_frame.setSizePolicy(
            QSizePolicy.Policy.Expanding,
            QSizePolicy.Policy.Maximum,
        )
        self.hint_frame.setStyleSheet(
            "QFrame { background: #fffbea; border: 1px solid #f3d76b; "
            "border-radius: 8px; }"
        )
        hint_layout = QVBoxLayout(self.hint_frame)
        hint_layout.setContentsMargins(12, 11, 12, 11)
        self.source_label = QLabel("")
        self.source_label.setWordWrap(True)
        self.source_label.setStyleSheet(
            "color: #4a3d16; background: transparent; border: none;"
        )
        hint_layout.addWidget(self.source_label)
        self.hint_frame.hide()
        question_layout.addWidget(self.hint_frame)
        question_layout.addStretch(1)
        layout.addWidget(question_card, stretch=1)

        navigation = QHBoxLayout()
        self.previous_button = QPushButton("← Câu trước")
        self.previous_button.setObjectName("SecondaryButton")
        self.previous_button.clicked.connect(self.previous_question)
        self.check_button = QPushButton("Kiểm tra đáp án")
        self.check_button.setObjectName("PrimaryButton")
        self.check_button.clicked.connect(self.check_answer)
        self.next_button = QPushButton("Câu tiếp →")
        self.next_button.setObjectName("SecondaryButton")
        self.next_button.clicked.connect(self.next_question)
        navigation.addWidget(self.previous_button)
        navigation.addStretch()
        navigation.addWidget(self.check_button)
        navigation.addStretch()
        navigation.addWidget(self.next_button)
        layout.addLayout(navigation)
        return page

    def _build_result_page(self) -> QWidget:
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(30, 24, 30, 26)
        layout.setSpacing(14)

        title = QLabel("Kết quả làm bài")
        title.setObjectName("PageTitle")
        layout.addWidget(title)

        summary = QFrame()
        summary.setObjectName("WelcomeCard")
        summary_layout = QHBoxLayout(summary)
        summary_layout.setContentsMargins(24, 18, 24, 18)
        self.result_percentage = QLabel("0%")
        self.result_percentage.setObjectName("MetricValue")
        self.result_summary = QLabel("")
        self.result_summary.setWordWrap(True)
        summary_layout.addWidget(self.result_percentage)
        summary_layout.addSpacing(20)
        summary_layout.addWidget(self.result_summary, stretch=1)
        layout.addWidget(summary)

        self.review_table = QTableWidget(0, 4)
        self.review_table.setHorizontalHeaderLabels(
            ["Câu", "Đã chọn", "Đáp án", "Kết quả"]
        )
        self.review_table.verticalHeader().setVisible(False)
        self.review_table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.review_table.setSelectionBehavior(
            QTableWidget.SelectionBehavior.SelectRows
        )
        self.review_table.setToolTip(
            "Bấm vào một dòng để xem chi tiết câu hỏi và đối chiếu đáp án."
        )
        self.review_table.cellClicked.connect(self.open_result_question)
        header = self.review_table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(2, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(3, QHeaderView.ResizeMode.Stretch)
        layout.addWidget(self.review_table, stretch=1)

        actions = QHBoxLayout()
        new_button = QPushButton("Làm bài mới")
        new_button.setObjectName("SecondaryButton")
        new_button.clicked.connect(self.reset_to_setup)
        close_button = QPushButton("Đóng")
        close_button.setObjectName("PrimaryButton")
        close_button.clicked.connect(self.request_close)
        actions.addStretch()
        actions.addWidget(new_button)
        actions.addWidget(close_button)
        layout.addLayout(actions)
        return page

    def _build_answer_lookup_page(self) -> QWidget:
        page = QWidget()
        page.setObjectName("AnswerLookupPage")
        layout = QVBoxLayout(page)
        layout.setContentsMargins(24, 22, 24, 24)
        layout.setSpacing(16)

        header = QHBoxLayout()
        header.setSpacing(14)
        header_icon = QLabel("⌕")
        header_icon.setObjectName("AnswerLookupHeaderIcon")
        header_icon.setAlignment(Qt.AlignmentFlag.AlignCenter)
        header_icon.setFixedSize(48, 48)
        header_text = QVBoxLayout()
        header_text.setSpacing(2)
        title = QLabel("Tra cứu đáp án")
        title.setObjectName("AnswerLookupTitle")
        description = QLabel(
            "Nhập vài từ trong câu hỏi, đáp án hoặc trích dẫn để tìm nhanh "
            "trong ngân hàng câu hỏi."
        )
        description.setObjectName("AnswerLookupDescription")
        header_text.addWidget(title)
        header_text.addWidget(description)
        header.addWidget(header_icon)
        header.addLayout(header_text, stretch=1)
        layout.addLayout(header)

        search_card = QFrame()
        search_card.setObjectName("AnswerSearchCard")
        self._apply_card_shadow(search_card)
        search_layout = QHBoxLayout(search_card)
        search_layout.setContentsMargins(28, 24, 28, 24)
        search_layout.setSpacing(18)

        search_input_wrap = QFrame()
        search_input_wrap.setObjectName("AnswerSearchInputWrap")
        search_input_layout = QHBoxLayout(search_input_wrap)
        search_input_layout.setContentsMargins(16, 0, 16, 0)
        search_input_layout.setSpacing(12)
        search_icon = QLabel("⌕")
        search_icon.setObjectName("AnswerSearchInlineIcon")
        search_icon.setAlignment(Qt.AlignmentFlag.AlignCenter)
        search_icon.setFixedWidth(28)
        self.answer_lookup_edit = QLineEdit()
        self.answer_lookup_edit.setObjectName("AnswerLookupEdit")
        self.answer_lookup_edit.setPlaceholderText(
            "Ví dụ: Quy định 641, IPCAS, tra soát, tín dụng..."
        )
        self.answer_lookup_timer = QTimer(self)
        self.answer_lookup_timer.setSingleShot(True)
        self.answer_lookup_timer.setInterval(300)
        self.answer_lookup_timer.timeout.connect(self.search_answers)
        self.answer_lookup_edit.textChanged.connect(self.schedule_answer_lookup)
        self.answer_lookup_edit.returnPressed.connect(self.search_answers)
        self.answer_lookup_clear_button = QPushButton("Xóa tìm kiếm")
        self.answer_lookup_clear_button.setObjectName("AnswerLookupClearButton")
        self.answer_lookup_clear_button.clicked.connect(self.clear_answer_lookup)
        search_input_layout.addWidget(search_icon)
        search_input_layout.addWidget(self.answer_lookup_edit, stretch=1)
        search_layout.addWidget(search_input_wrap, stretch=1)
        search_layout.addWidget(self.answer_lookup_clear_button)
        layout.addWidget(search_card)

        self.answer_lookup_status = QLabel(
            "Nhập từ khóa để hệ thống tự tra cứu. Nhấp đúp vào dòng để xem chi tiết."
        )
        self.answer_lookup_status.setObjectName("AnswerLookupStatus")
        self.answer_lookup_status.setTextFormat(Qt.TextFormat.RichText)
        layout.addWidget(self.answer_lookup_status)

        result_card = QFrame()
        result_card.setObjectName("AnswerResultCard")
        self._apply_card_shadow(result_card)
        result_layout = QVBoxLayout(result_card)
        result_layout.setContentsMargins(22, 18, 22, 18)
        result_layout.setSpacing(16)

        result_header = QHBoxLayout()
        result_header.setSpacing(10)
        result_icon = QLabel("≡")
        result_icon.setObjectName("AnswerResultHeaderIcon")
        result_icon.setAlignment(Qt.AlignmentFlag.AlignCenter)
        result_icon.setFixedSize(28, 28)
        result_title = QLabel("Danh sách kết quả")
        result_title.setObjectName("AnswerResultTitle")
        self.answer_lookup_badge = QLabel("0 kết quả")
        self.answer_lookup_badge.setObjectName("AnswerResultBadge")
        self.answer_lookup_badge.setAlignment(Qt.AlignmentFlag.AlignCenter)
        result_header.addWidget(result_icon)
        result_header.addWidget(result_title)
        result_header.addStretch()
        result_header.addWidget(self.answer_lookup_badge)
        result_layout.addLayout(result_header)

        self.answer_lookup_scroll = QScrollArea()
        self.answer_lookup_scroll.setObjectName("AnswerResultScroll")
        self.answer_lookup_scroll.setWidgetResizable(True)
        self.answer_lookup_scroll.setFrameShape(QFrame.Shape.NoFrame)
        self.answer_lookup_scroll.setHorizontalScrollBarPolicy(
            Qt.ScrollBarPolicy.ScrollBarAlwaysOff
        )
        self.answer_lookup_results_widget = QWidget()
        self.answer_lookup_results_widget.setObjectName("AnswerResultsViewport")
        self.answer_lookup_results_layout = QVBoxLayout(
            self.answer_lookup_results_widget
        )
        self.answer_lookup_results_layout.setContentsMargins(0, 0, 0, 0)
        self.answer_lookup_results_layout.setSpacing(12)
        self.answer_lookup_empty_label = QLabel(
            "Chưa có kết quả. Nhập từ khóa để bắt đầu tra cứu."
        )
        self.answer_lookup_empty_label.setObjectName("AnswerLookupEmpty")
        self.answer_lookup_empty_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.answer_lookup_results_layout.addWidget(self.answer_lookup_empty_label)
        self.answer_lookup_results_layout.addStretch()
        self.answer_lookup_scroll.setWidget(self.answer_lookup_results_widget)
        result_layout.addWidget(self.answer_lookup_scroll, stretch=1)
        layout.addWidget(result_card, stretch=1)

        self.answer_lookup_questions: list[Question] = []
        self.answer_lookup_result_items: list[AnswerLookupResultItem] = []
        self.answer_lookup_selected_row = -1

        # Giữ control cũ ở dạng ẩn để không làm lỗi mã ngoài nếu có tham chiếu.
        self.answer_lookup_table = QTableWidget(0, 1)
        self.answer_lookup_table.setObjectName("AnswerLookupTable")
        self.answer_lookup_table.setHorizontalHeaderLabels(
            ["Câu hỏi và đáp án đúng"]
        )
        self.answer_lookup_table.hide()
        return page

    def _build_data_management_page(self) -> QWidget:
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(12, 10, 12, 12)
        self.data_admin_tabs = QTabWidget()
        self.data_admin_tabs.addTab(
            self._build_business_management_tab(),
            "Nghiệp Vụ & Chuyên Đề",
        )
        self.data_admin_tabs.addTab(
            self._build_question_management_tab(),
            "Chỉnh Sửa Câu Hỏi",
        )
        layout.addWidget(self.data_admin_tabs)
        return page

    def _build_business_management_tab(self) -> QWidget:
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(14, 12, 14, 14)
        layout.setSpacing(12)

        heading = QLabel("Quản lý nghiệp vụ và chuyên đề")
        heading.setObjectName("SectionTitle")
        description = QLabel(
            "Chọn một dòng để chỉnh sửa. Chuyên đề bên phải được lọc theo "
            "nghiệp vụ đang chọn."
        )
        description.setObjectName("MutedText")
        layout.addWidget(heading)
        layout.addWidget(description)

        content = QHBoxLayout()
        content.setSpacing(14)

        business_group = QGroupBox("1. Danh sách nghiệp vụ")
        business_group.setObjectName("AdminPanel")
        business_layout = QVBoxLayout(business_group)
        business_layout.setContentsMargins(14, 18, 14, 14)
        business_layout.setSpacing(10)

        business_toolbar = QHBoxLayout()
        business_hint = QLabel("Chọn nghiệp vụ để xem các chuyên đề")
        business_hint.setObjectName("MutedText")
        self.new_business_button = QPushButton("＋ Thêm nghiệp vụ")
        self.new_business_button.setObjectName("SecondaryButton")
        self.new_business_button.clicked.connect(self.new_business_topic)
        business_toolbar.addWidget(business_hint)
        business_toolbar.addStretch()
        business_toolbar.addWidget(self.new_business_button)
        business_layout.addLayout(business_toolbar)

        self.business_admin_table = QTableWidget(0, 3)
        self.business_admin_table.setHorizontalHeaderLabels(
            ["Mã", "Tên nghiệp vụ", "Số câu"]
        )
        self.business_admin_table.verticalHeader().setVisible(False)
        self.business_admin_table.setEditTriggers(
            QTableWidget.EditTrigger.NoEditTriggers
        )
        self.business_admin_table.setSelectionBehavior(
            QTableWidget.SelectionBehavior.SelectRows
        )
        self.business_admin_table.setSelectionMode(
            QTableWidget.SelectionMode.SingleSelection
        )
        self.business_admin_table.itemSelectionChanged.connect(
            self.load_selected_business_topic
        )
        business_header = self.business_admin_table.horizontalHeader()
        business_header.setSectionResizeMode(
            0, QHeaderView.ResizeMode.ResizeToContents
        )
        business_header.setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        business_header.setSectionResizeMode(
            2, QHeaderView.ResizeMode.ResizeToContents
        )
        business_layout.addWidget(self.business_admin_table, stretch=1)

        self.editing_business_code: str | None = None
        business_editor = QFrame()
        business_editor.setObjectName("AdminEditor")
        business_editor_layout = QVBoxLayout(business_editor)
        business_editor_layout.setContentsMargins(14, 12, 14, 12)
        business_editor_layout.setSpacing(9)
        self.business_editor_title = QLabel("Thông tin nghiệp vụ")
        self.business_editor_title.setObjectName("AdminEditorTitle")
        business_editor_layout.addWidget(self.business_editor_title)

        business_form = QFormLayout()
        business_form.setHorizontalSpacing(12)
        business_form.setVerticalSpacing(8)
        self.topic_code_edit = QLineEdit()
        self.topic_code_edit.setPlaceholderText("Ví dụ: TINDUNG")
        self.topic_name_edit = QLineEdit()
        self.topic_name_edit.setPlaceholderText("Nhập tên nghiệp vụ")
        self.topic_code_edit.returnPressed.connect(self.save_business_topic)
        self.topic_name_edit.returnPressed.connect(self.save_business_topic)
        business_form.addRow("Mã nghiệp vụ:", self.topic_code_edit)
        business_form.addRow("Tên nghiệp vụ:", self.topic_name_edit)
        business_editor_layout.addLayout(business_form)

        business_actions = QHBoxLayout()
        self.delete_business_button = QPushButton("Xóa")
        self.delete_business_button.setObjectName("DangerButton")
        self.delete_business_button.clicked.connect(self.delete_business_topic)
        self.cancel_business_button = QPushButton("Hủy")
        self.cancel_business_button.clicked.connect(self.cancel_business_edit)
        self.save_business_button = QPushButton("Lưu thay đổi")
        self.save_business_button.setObjectName("PrimaryButton")
        self.save_business_button.clicked.connect(self.save_business_topic)
        business_actions.addWidget(self.delete_business_button)
        business_actions.addStretch()
        business_actions.addWidget(self.cancel_business_button)
        business_actions.addWidget(self.save_business_button)
        business_editor_layout.addLayout(business_actions)
        business_layout.addWidget(business_editor)

        subject_group = QGroupBox("2. Chuyên đề của nghiệp vụ")
        subject_group.setObjectName("AdminPanel")
        subject_layout = QVBoxLayout(subject_group)
        subject_layout.setContentsMargins(14, 18, 14, 14)
        subject_layout.setSpacing(10)

        subject_toolbar = QHBoxLayout()
        self.subject_context_label = QLabel("Chưa chọn nghiệp vụ")
        self.subject_context_label.setObjectName("MutedText")
        self.new_subject_button = QPushButton("＋ Thêm chuyên đề")
        self.new_subject_button.setObjectName("SecondaryButton")
        self.new_subject_button.clicked.connect(self.new_question_topic)
        subject_toolbar.addWidget(self.subject_context_label)
        subject_toolbar.addStretch()
        subject_toolbar.addWidget(self.new_subject_button)
        subject_layout.addLayout(subject_toolbar)

        self.subject_admin_table = QTableWidget(0, 2)
        self.subject_admin_table.setHorizontalHeaderLabels(
            ["Tên chuyên đề", "Số câu"]
        )
        self.subject_admin_table.verticalHeader().setVisible(False)
        self.subject_admin_table.setEditTriggers(
            QTableWidget.EditTrigger.NoEditTriggers
        )
        self.subject_admin_table.setSelectionBehavior(
            QTableWidget.SelectionBehavior.SelectRows
        )
        self.subject_admin_table.setSelectionMode(
            QTableWidget.SelectionMode.SingleSelection
        )
        self.subject_admin_table.itemSelectionChanged.connect(
            self.load_selected_question_topic
        )
        subject_header = self.subject_admin_table.horizontalHeader()
        subject_header.setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        subject_header.setSectionResizeMode(
            1, QHeaderView.ResizeMode.ResizeToContents
        )
        subject_layout.addWidget(self.subject_admin_table, stretch=1)

        self.editing_subject_id: int | None = None
        subject_editor = QFrame()
        subject_editor.setObjectName("AdminEditor")
        subject_editor_layout = QVBoxLayout(subject_editor)
        subject_editor_layout.setContentsMargins(14, 12, 14, 12)
        subject_editor_layout.setSpacing(9)
        self.subject_editor_title = QLabel("Thông tin chuyên đề")
        self.subject_editor_title.setObjectName("AdminEditorTitle")
        subject_editor_layout.addWidget(self.subject_editor_title)

        subject_form = QFormLayout()
        subject_form.setHorizontalSpacing(12)
        subject_form.setVerticalSpacing(8)
        self.subject_business_combo = QComboBox()
        self.subject_name_edit = QLineEdit()
        self.subject_name_edit.setPlaceholderText(
            "Ví dụ: Ngân hàng điện tử, Điều lệ ngân hàng, Tín dụng..."
        )
        self.subject_name_edit.returnPressed.connect(self.save_question_topic)
        subject_form.addRow("Thuộc nghiệp vụ:", self.subject_business_combo)
        subject_form.addRow("Tên chuyên đề:", self.subject_name_edit)
        subject_editor_layout.addLayout(subject_form)

        subject_actions = QHBoxLayout()
        self.delete_subject_button = QPushButton("Xóa")
        self.delete_subject_button.setObjectName("DangerButton")
        self.delete_subject_button.clicked.connect(self.delete_question_topic)
        self.cancel_subject_button = QPushButton("Hủy")
        self.cancel_subject_button.clicked.connect(self.cancel_subject_edit)
        self.save_subject_button = QPushButton("Lưu thay đổi")
        self.save_subject_button.setObjectName("PrimaryButton")
        self.save_subject_button.clicked.connect(self.save_question_topic)
        subject_actions.addWidget(self.delete_subject_button)
        subject_actions.addStretch()
        subject_actions.addWidget(self.cancel_subject_button)
        subject_actions.addWidget(self.save_subject_button)
        subject_editor_layout.addLayout(subject_actions)
        subject_layout.addWidget(subject_editor)

        content.addWidget(business_group, stretch=1)
        content.addWidget(subject_group, stretch=1)
        layout.addLayout(content, stretch=1)
        return page

    def _build_question_management_tab(self) -> QWidget:
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(14, 14, 14, 14)

        search_row = QHBoxLayout()
        self.question_search_edit = QLineEdit()
        self.question_search_edit.setPlaceholderText(
            "Tìm trong câu hỏi, đáp án hoặc nguồn tham khảo..."
        )
        self.question_search_edit.returnPressed.connect(self.refresh_question_table)
        self.question_topic_filter = QComboBox()
        self.question_topic_filter.currentIndexChanged.connect(
            self.refresh_question_table
        )
        search_button = QPushButton("Tìm kiếm")
        search_button.setObjectName("SecondaryButton")
        search_button.clicked.connect(self.refresh_question_table)
        add_button = QPushButton("Thêm câu hỏi")
        add_button.setObjectName("PrimaryButton")
        add_button.clicked.connect(self.new_question)
        import_button = QPushButton("Nhập từ Excel")
        import_button.setObjectName("SecondaryButton")
        import_button.clicked.connect(self.import_questions_excel)
        template_button = QPushButton("Tải Excel mẫu")
        template_button.setObjectName("SecondaryButton")
        template_button.clicked.connect(self.download_excel_template)
        delete_topic_questions = QPushButton("Xóa câu hỏi theo nghiệp vụ")
        delete_topic_questions.setObjectName("SecondaryButton")
        delete_topic_questions.clicked.connect(self.delete_questions_for_business)
        delete_all_questions = QPushButton("Xóa toàn bộ câu hỏi")
        delete_all_questions.setStyleSheet(
            "color: #9b1c1c; background: #feecec; padding: 8px 12px;"
        )
        delete_all_questions.clicked.connect(self.delete_all_questions)
        search_row.addWidget(self.question_search_edit, stretch=1)
        search_row.addWidget(self.question_topic_filter)
        search_row.addWidget(search_button)
        search_row.addWidget(add_button)
        layout.addLayout(search_row)

        bulk_row = QHBoxLayout()
        bulk_row.addWidget(import_button)
        bulk_row.addWidget(template_button)
        bulk_row.addStretch()
        bulk_row.addWidget(delete_topic_questions)
        bulk_row.addWidget(delete_all_questions)
        layout.addLayout(bulk_row)

        question_content = QSplitter(Qt.Orientation.Horizontal)
        question_content.setChildrenCollapsible(False)

        question_list_panel = QWidget()
        question_list_layout = QVBoxLayout(question_list_panel)
        question_list_layout.setContentsMargins(0, 0, 6, 0)
        question_list_layout.setSpacing(6)
        list_hint = QLabel(
            "Danh sách câu hỏi · Nhấp đúp vào một dòng để chỉnh sửa"
        )
        list_hint.setObjectName("MutedText")
        question_list_layout.addWidget(list_hint)

        self.question_admin_table = QTableWidget(0, 5)
        self.question_admin_table.setHorizontalHeaderLabels(
            ["ID", "Nghiệp vụ", "Chuyên đề", "Câu hỏi", "Đáp án đúng"]
        )
        self.question_admin_table.verticalHeader().setVisible(False)
        self.question_admin_table.setEditTriggers(
            QTableWidget.EditTrigger.NoEditTriggers
        )
        self.question_admin_table.setSelectionBehavior(
            QTableWidget.SelectionBehavior.SelectRows
        )
        self.question_admin_table.doubleClicked.connect(self.edit_selected_question)
        admin_header = self.question_admin_table.horizontalHeader()
        admin_header.setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
        admin_header.setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
        admin_header.setSectionResizeMode(2, QHeaderView.ResizeMode.Stretch)
        admin_header.setSectionResizeMode(3, QHeaderView.ResizeMode.Stretch)
        admin_header.setSectionResizeMode(4, QHeaderView.ResizeMode.ResizeToContents)
        question_list_layout.addWidget(self.question_admin_table, stretch=1)
        question_content.addWidget(question_list_panel)

        self.question_editor = QGroupBox("Thêm / chỉnh sửa câu hỏi")
        editor = QGridLayout(self.question_editor)
        editor.setContentsMargins(16, 18, 16, 14)
        editor.setHorizontalSpacing(10)
        editor.setVerticalSpacing(8)
        self.editing_question_id: int | None = None
        self.editor_topic_combo = QComboBox()
        self.editor_topic_combo.currentIndexChanged.connect(
            self.refresh_editor_subjects
        )
        self.editor_subject_combo = QComboBox()
        self.editor_question = ExpandingTextEdit(82, 280)
        self.editor_question.setPlaceholderText("Nhập đầy đủ nội dung câu hỏi...")
        self.editor_options: dict[str, QTextEdit] = {}
        self.editor_correct = QComboBox()
        self.editor_correct.addItems(list("ABCD"))
        self.editor_locked_answers: dict[str, QCheckBox] = {}
        self.editor_source = ExpandingTextEdit(52, 140)
        self.editor_source.setPlaceholderText(
            "Nhập nguồn hoặc căn cứ của câu hỏi..."
        )
        editor.addWidget(QLabel("Nghiệp vụ:"), 0, 0)
        editor.addWidget(self.editor_topic_combo, 0, 1, 1, 3)
        editor.addWidget(QLabel("Chuyên đề:"), 1, 0)
        editor.addWidget(self.editor_subject_combo, 1, 1, 1, 3)
        editor.addWidget(QLabel("Đáp án đúng:"), 2, 0)
        editor.addWidget(self.editor_correct, 2, 1)
        locked_row = QHBoxLayout()
        locked_row.addWidget(QLabel("Đáp án giữ nguyên khi hoán đổi:"))
        for letter in "ABCD":
            checkbox = QCheckBox(letter)
            self.editor_locked_answers[letter] = checkbox
            locked_row.addWidget(checkbox)
        locked_row.addStretch()
        editor.addLayout(locked_row, 2, 2, 1, 2)
        editor.addWidget(QLabel("Câu hỏi:"), 3, 0)
        editor.addWidget(self.editor_question, 3, 1, 1, 3)
        for row, letter in enumerate("ABCD", start=4):
            field = ExpandingTextEdit(58, 220)
            field.setPlaceholderText(f"Nhập nội dung đáp án {letter}...")
            self.editor_options[letter] = field
            editor.addWidget(QLabel(f"Đáp án {letter}:"), row, 0)
            editor.addWidget(field, row, 1, 1, 3)
        editor.addWidget(QLabel("Nguồn:"), 8, 0)
        editor.addWidget(self.editor_source, 8, 1, 1, 3)
        editor_actions = QHBoxLayout()
        save_question_button = QPushButton("Lưu câu hỏi")
        save_question_button.setObjectName("PrimaryButton")
        save_question_button.clicked.connect(self.save_question)
        delete_question_button = QPushButton("Xóa câu hỏi")
        delete_question_button.setObjectName("SecondaryButton")
        delete_question_button.clicked.connect(self.delete_question)
        cancel_edit_button = QPushButton("Hủy")
        cancel_edit_button.clicked.connect(self.clear_question_editor)
        editor_actions.addStretch()
        editor_actions.addWidget(delete_question_button)
        editor_actions.addWidget(cancel_edit_button)
        editor_actions.addWidget(save_question_button)
        editor.addLayout(editor_actions, 9, 0, 1, 4)

        editor_scroll = QScrollArea()
        editor_scroll.setObjectName("QuestionEditorScroll")
        editor_scroll.setWidgetResizable(True)
        editor_scroll.setFrameShape(QFrame.Shape.NoFrame)
        editor_scroll.setHorizontalScrollBarPolicy(
            Qt.ScrollBarPolicy.ScrollBarAlwaysOff
        )
        editor_scroll.setWidget(self.question_editor)
        question_content.addWidget(editor_scroll)
        question_content.setStretchFactor(0, 2)
        question_content.setStretchFactor(1, 3)
        question_content.setSizes([430, 650])
        layout.addWidget(question_content, stretch=1)
        return page

    def refresh_setup_data(self) -> None:
        current_topic = self.topic_combo.currentData()
        topics = self.database.topics()
        self.topic_combo.blockSignals(True)
        self.topic_combo.clear()
        for topic, count in topics:
            self.topic_combo.addItem(f"{topic}: {count} câu", (topic, count))
        self.topic_combo.blockSignals(False)

        self.employee_combo.clear()
        self.employee_combo.addItems(self.database.employees())
        if current_topic is not None:
            for index in range(self.topic_combo.count()):
                if self.topic_combo.itemData(index) == current_topic:
                    self.topic_combo.setCurrentIndex(index)
                    break
        self.update_topic_summary()
        self.refresh_business_ui()
        self.refresh_random_topics()
        self.refresh_question_table()
        self.update_quota_total()
        question_count = self.database.question_count()
        business_count = len(self.database.business_topics())
        self.data_status.setText(
            "Hệ thống có "
            f"{question_count:,}".replace(",", ".")
            + " câu hỏi / "
            + f"{business_count:,}".replace(",", ".")
            + " nghiệp vụ."
        )

    def update_topic_summary(self) -> None:
        data = self.topic_combo.currentData()
        if not data:
            self.topic_summary.setText("")
            return
        topic, count = data
        employee = self.employee_combo.currentText().strip() or "(Chưa nhập)"
        time_text = (
            f"Giới hạn {self.time_limit.value()} phút"
            if self.limited_time.isChecked()
            else "Không giới hạn thời gian trả lời"
        )
        answer_text = (
            "Thể hiện thứ tự đáp án theo đề cương"
            if self.source_answer_order.isChecked()
            else "Hoán đổi thứ tự đáp án so với đề cương"
        )
        self.topic_summary.setText(
            "<div style='line-height: 1.55;'>"
            f"<p><span style='color:#a11f4d;'>•</span>&nbsp; Tên NV: {employee}</p>"
            f"<p><span style='color:#a11f4d;'>•</span>&nbsp; "
            f"Chuyên đề: {topic or 'Tất cả chuyên đề'}</p>"
            f"<p><span style='color:#a11f4d;'>•</span>&nbsp; "
            f"Tổng số câu hỏi: {count} câu</p>"
            f"<p><span style='color:#a11f4d;'>•</span>&nbsp; {time_text}</p>"
            f"<p><span style='color:#a11f4d;'>•</span>&nbsp; {answer_text}</p>"
            "</div>"
        )

    def update_quota_total(self) -> None:
        total = sum(self.random_quotas().values()) if hasattr(self, "quota_table") else 0
        self.total_quota_label.setText(f"Tổng cộng: {total} câu")

    def refresh_random_topics(self) -> None:
        saved = self.database.load_random_subject_setting()
        businesses = self.database.business_topics()
        subjects_by_business: dict[str, list[tuple[int, str, int]]] = {}
        for subject_id, business_code, subject_name, available in (
            self.database.question_topics()
        ):
            subjects_by_business.setdefault(business_code, []).append(
                (subject_id, subject_name, available)
            )
        self.quota_table.blockSignals(True)
        self.quota_table.clear()
        for business_code, business_name, _ in businesses:
            subjects = subjects_by_business.get(business_code, [])
            if not subjects:
                continue
            if len(subjects) == 1:
                subject_id, _, available = subjects[0]
                item = QTreeWidgetItem(self.quota_table)
                item.setText(1, business_name)
                item.setToolTip(1, business_name)
                self._configure_quota_item(
                    item,
                    subject_id,
                    available,
                    saved.get(subject_id, 0),
                )
                continue

            parent = QTreeWidgetItem(self.quota_table)
            parent.setText(1, business_name)
            parent.setToolTip(1, business_name)
            same_name_subject = next(
                (
                    subject
                    for subject in subjects
                    if subject[1].strip().casefold()
                    == business_name.strip().casefold()
                ),
                None,
            )
            if same_name_subject:
                subject_id, _, available = same_name_subject
                self._configure_quota_item(
                    parent,
                    subject_id,
                    available,
                    saved.get(subject_id, 0),
                    show_input=False,
                )
            else:
                parent.setText(2, str(sum(item[2] for item in subjects)))
                parent.setFlags(
                    Qt.ItemFlag.ItemIsEnabled
                    | Qt.ItemFlag.ItemIsUserCheckable
                )
                parent.setCheckState(0, Qt.CheckState.Unchecked)
            has_saved_selection = False
            for subject_id, subject_name, available in subjects:
                if same_name_subject and subject_id == same_name_subject[0]:
                    has_saved_selection = (
                        has_saved_selection or saved.get(subject_id, 0) > 0
                    )
                    continue
                child = QTreeWidgetItem(parent)
                child.setText(1, subject_name)
                child.setToolTip(1, subject_name)
                self._configure_quota_item(
                    child,
                    subject_id,
                    available,
                    saved.get(subject_id, 0),
                )
                has_saved_selection = (
                    has_saved_selection or saved.get(subject_id, 0) > 0
                )
            if not same_name_subject and has_saved_selection:
                parent.setCheckState(0, Qt.CheckState.Checked)
            parent.setExpanded(has_saved_selection)
        self.quota_table.blockSignals(False)
        self.update_quota_total()

    def _configure_quota_item(
        self,
        item: QTreeWidgetItem,
        subject_id: int,
        available: int,
        saved_value: int,
        show_input: bool = True,
    ) -> None:
        item.setFlags(
            Qt.ItemFlag.ItemIsEnabled | Qt.ItemFlag.ItemIsUserCheckable
        )
        item.setData(0, Qt.ItemDataRole.UserRole, subject_id)
        item.setCheckState(
            0,
            Qt.CheckState.Checked
            if saved_value > 0
            else Qt.CheckState.Unchecked,
        )
        item.setText(2, str(available))
        if not show_input:
            return
        spin = QSpinBox()
        spin.setRange(0, available)
        spin.setValue(min(saved_value, available))
        spin.setButtonSymbols(QSpinBox.ButtonSymbols.NoButtons)
        spin.setAlignment(Qt.AlignmentFlag.AlignRight)
        spin.setFixedWidth(72)
        spin.setFixedHeight(32)
        input_container = QWidget()
        input_container.setMinimumHeight(36)
        input_layout = QHBoxLayout(input_container)
        input_layout.setContentsMargins(2, 0, 2, 0)
        input_layout.setSpacing(4)
        input_layout.addWidget(spin)
        input_layout.addWidget(QLabel("câu"))
        input_layout.addStretch()
        spin.valueChanged.connect(
            lambda value, target=item: self._quota_value_changed(target, value)
        )
        spin.valueChanged.connect(self.update_quota_total)
        self.quota_table.setItemWidget(item, 3, input_container)

    def _quota_spin(self, item: QTreeWidgetItem) -> QSpinBox | None:
        widget = self.quota_table.itemWidget(item, 3)
        if isinstance(widget, QSpinBox):
            return widget
        return widget.findChild(QSpinBox) if widget is not None else None

    def _quota_value_changed(self, item: QTreeWidgetItem, value: int) -> None:
        item.setCheckState(
            0,
            Qt.CheckState.Checked if value > 0 else Qt.CheckState.Unchecked
        )

    def _quota_item_changed(
        self,
        item: QTreeWidgetItem,
        column: int,
    ) -> None:
        if column == 0:
            if item.childCount() > 0:
                item.setExpanded(
                    item.checkState(0) == Qt.CheckState.Checked
                )
            spin = self._quota_spin(item)
            if isinstance(spin, QSpinBox):
                if (
                    item.checkState(0) == Qt.CheckState.Checked
                    and spin.value() == 0
                    and spin.maximum() > 0
                ):
                    spin.setValue(1)
                elif (
                    item.checkState(0) == Qt.CheckState.Unchecked
                    and spin.value() > 0
                ):
                    spin.setValue(0)
        self.update_quota_total()

    def random_quotas(self) -> dict[int, int]:
        quotas: dict[int, int] = {}
        items: list[QTreeWidgetItem] = []
        for index in range(self.quota_table.topLevelItemCount()):
            top_level = self.quota_table.topLevelItem(index)
            items.append(top_level)
            items.extend(
                top_level.child(child_index)
                for child_index in range(top_level.childCount())
            )
        for item in items:
            subject_id = item.data(0, Qt.ItemDataRole.UserRole)
            spin = self._quota_spin(item)
            if (
                subject_id is not None
                and item.checkState(0) == Qt.CheckState.Checked
            ):
                value = spin.value() if isinstance(spin, QSpinBox) else int(
                    item.text(2) or 0
                )
                if value > 0:
                    quotas[int(subject_id)] = value
        return quotas

    def refresh_business_ui(self) -> None:
        topics = self.database.business_topics()
        current_code = self.editing_business_code
        for combo, include_all in (
            (self.subject_business_combo, False),
            (self.question_topic_filter, True),
            (self.editor_topic_combo, False),
        ):
            combo.blockSignals(True)
            combo.clear()
            if include_all:
                combo.addItem("Tất cả nghiệp vụ", None)
            for code, name, count in topics:
                combo.addItem(f"{name} ({count})", code)
            combo.blockSignals(False)

        self.business_admin_table.blockSignals(True)
        self.business_admin_table.setRowCount(len(topics))
        selected_row = -1
        for row, (code, name, count) in enumerate(topics):
            for column, value in enumerate((code, name, count)):
                item = QTableWidgetItem(str(value))
                item.setData(Qt.ItemDataRole.UserRole, code)
                self.business_admin_table.setItem(row, column, item)
            if code == current_code:
                selected_row = row
        self.business_admin_table.blockSignals(False)
        if selected_row < 0 and topics:
            selected_row = 0
        if selected_row >= 0:
            self.business_admin_table.selectRow(selected_row)
        else:
            self.new_business_topic()
        self.load_selected_business_topic()
        self.refresh_editor_subjects()

    def load_selected_business_topic(self) -> None:
        selected_items = self.business_admin_table.selectedItems()
        if not selected_items:
            return
        row = selected_items[0].row()
        previous_code = self.editing_business_code
        code = self.business_admin_table.item(row, 0).text()
        self.editing_business_code = code
        self.topic_code_edit.setText(code)
        self.topic_name_edit.setText(
            self.business_admin_table.item(row, 1).text()
        )
        self.business_editor_title.setText("Chỉnh sửa nghiệp vụ")
        self.save_business_button.setText("Lưu thay đổi")
        self.delete_business_button.setEnabled(True)
        self.cancel_business_button.setEnabled(True)
        if previous_code != code:
            self.editing_subject_id = None
        self.refresh_subject_admin_table()

    def refresh_subject_admin_table(self) -> None:
        if not hasattr(self, "subject_admin_table"):
            return
        business_code = self.editing_business_code
        subjects = (
            self.database.question_topics(business_code)
            if business_code
            else []
        )
        business_name = self.topic_name_edit.text().strip()
        self.subject_context_label.setText(
            f"Đang xem: {business_name}" if business_code else "Chưa chọn nghiệp vụ"
        )
        self.new_subject_button.setEnabled(bool(business_code))
        selected_row = -1
        self.subject_admin_table.blockSignals(True)
        self.subject_admin_table.setRowCount(len(subjects))
        for row, (subject_id, code, name, count) in enumerate(subjects):
            for column, value in enumerate((name, count)):
                item = QTableWidgetItem(str(value))
                item.setData(Qt.ItemDataRole.UserRole, subject_id)
                item.setData(Qt.ItemDataRole.UserRole + 1, code)
                self.subject_admin_table.setItem(row, column, item)
            if subject_id == self.editing_subject_id:
                selected_row = row
        self.subject_admin_table.blockSignals(False)
        if selected_row >= 0:
            self.subject_admin_table.selectRow(selected_row)
            self.load_selected_question_topic()
        else:
            self.new_question_topic()

    def load_selected_question_topic(self) -> None:
        selected_items = self.subject_admin_table.selectedItems()
        if not selected_items:
            return
        row = selected_items[0].row()
        self.editing_subject_id = int(
            self.subject_admin_table.item(row, 0).data(Qt.ItemDataRole.UserRole)
        )
        code = str(
            self.subject_admin_table.item(row, 0).data(
                Qt.ItemDataRole.UserRole + 1
            )
        )
        self.subject_business_combo.setCurrentIndex(
            self.subject_business_combo.findData(code)
        )
        self.subject_name_edit.setText(self.subject_admin_table.item(row, 0).text())
        self.subject_editor_title.setText("Chỉnh sửa chuyên đề")
        self.save_subject_button.setText("Lưu thay đổi")
        self.delete_subject_button.setEnabled(True)
        self.cancel_subject_button.setEnabled(True)
        self.subject_business_combo.setEnabled(True)
        self.subject_name_edit.setEnabled(True)
        self.save_subject_button.setEnabled(True)

    def new_question_topic(self) -> None:
        self.editing_subject_id = None
        self.subject_admin_table.blockSignals(True)
        self.subject_admin_table.clearSelection()
        self.subject_admin_table.setCurrentItem(None)
        self.subject_admin_table.blockSignals(False)
        code = self.editing_business_code
        if code:
            self.subject_business_combo.setCurrentIndex(
                self.subject_business_combo.findData(code)
            )
        self.subject_name_edit.clear()
        self.subject_editor_title.setText("Thêm chuyên đề mới")
        self.save_subject_button.setText("Thêm chuyên đề")
        self.delete_subject_button.setEnabled(False)
        self.cancel_subject_button.setEnabled(bool(code))
        self.subject_business_combo.setEnabled(bool(code))
        self.subject_name_edit.setEnabled(bool(code))
        self.save_subject_button.setEnabled(bool(code))
        if self.sender() is self.new_subject_button:
            self.subject_name_edit.setFocus()

    def cancel_subject_edit(self) -> None:
        if self.subject_admin_table.selectedItems():
            self.load_selected_question_topic()
        else:
            self.new_question_topic()

    def save_question_topic(self) -> None:
        business_code = str(self.subject_business_combo.currentData() or "")
        try:
            self.editing_subject_id = self.database.save_question_topic(
                business_code,
                self.subject_name_edit.text(),
                self.editing_subject_id,
            )
            self.editing_business_code = business_code
            self.refresh_setup_data()
        except (QuizDatabaseError, sqlite3.IntegrityError) as exc:
            QMessageBox.warning(self, "Không thể lưu chuyên đề", str(exc))

    def delete_question_topic(self) -> None:
        if self.editing_subject_id is None:
            return
        subject_name = self.subject_name_edit.text().strip()
        if QMessageBox.question(
            self,
            "Xóa chuyên đề",
            f"Xóa chuyên đề “{subject_name}”?\n\n"
            "Toàn bộ câu hỏi thuộc chuyên đề này sẽ ngừng được sử dụng. "
            "Thao tác này không thể hoàn tác.",
        ) != QMessageBox.StandardButton.Yes:
            return
        self.database.delete_question_topic(self.editing_subject_id)
        self.new_question_topic()
        self.refresh_setup_data()

    def new_business_topic(self) -> None:
        self.editing_business_code = None
        self.business_admin_table.blockSignals(True)
        self.business_admin_table.clearSelection()
        self.business_admin_table.setCurrentItem(None)
        self.business_admin_table.blockSignals(False)
        self.topic_code_edit.clear()
        self.topic_name_edit.clear()
        self.business_editor_title.setText("Thêm nghiệp vụ mới")
        self.save_business_button.setText("Thêm nghiệp vụ")
        self.delete_business_button.setEnabled(False)
        self.cancel_business_button.setEnabled(
            self.business_admin_table.rowCount() > 0
        )
        self.editing_subject_id = None
        self.refresh_subject_admin_table()
        if self.sender() is self.new_business_button:
            self.topic_code_edit.setFocus()

    def cancel_business_edit(self) -> None:
        if self.business_admin_table.rowCount() > 0:
            self.business_admin_table.selectRow(0)
            self.load_selected_business_topic()
        else:
            self.new_business_topic()

    def save_business_topic(self) -> None:
        code = self.topic_code_edit.text().strip()
        try:
            self.database.save_business_topic(
                code,
                self.topic_name_edit.text(),
                self.editing_business_code,
            )
            self.editing_business_code = code
            self.refresh_setup_data()
        except QuizDatabaseError as exc:
            QMessageBox.warning(self, "Không thể lưu nghiệp vụ", str(exc))

    def delete_business_topic(self) -> None:
        code = self.editing_business_code
        if not code:
            return
        business_name = self.topic_name_edit.text().strip()
        if QMessageBox.question(
            self,
            "Xóa nghiệp vụ",
            f"Xóa nghiệp vụ “{business_name}”?\n\n"
            "Toàn bộ chuyên đề và câu hỏi thuộc nghiệp vụ này sẽ ngừng được sử "
            "dụng. Thao tác này không thể hoàn tác.",
        ) != QMessageBox.StandardButton.Yes:
            return
        self.database.delete_business_topic(str(code))
        self.editing_business_code = None
        self.refresh_setup_data()

    def refresh_question_table(self) -> None:
        if not hasattr(self, "question_admin_table"):
            return
        questions = self.database.search_questions(
            self.question_search_edit.text(),
            self.question_topic_filter.currentData(),
        )
        self.question_admin_table.setRowCount(len(questions))
        for row, question in enumerate(questions):
            for column, value in enumerate(
                (
                    question.id,
                    question.topic_name,
                    question.subject_name,
                    question.text,
                    question.correct_answer,
                )
            ):
                item = QTableWidgetItem(str(value))
                item.setData(Qt.ItemDataRole.UserRole, question.id)
                self.question_admin_table.setItem(row, column, item)

    def new_question(self) -> None:
        self.clear_question_editor()
        self.editor_question.setFocus()

    def edit_selected_question(self) -> None:
        row = self.question_admin_table.currentRow()
        if row < 0:
            return
        question_id = int(
            self.question_admin_table.item(row, 0).data(Qt.ItemDataRole.UserRole)
        )
        question = self.database.question_by_id(question_id)
        if not question:
            return
        self.editing_question_id = question.id
        self.editor_topic_combo.setCurrentIndex(
            self.editor_topic_combo.findData(question.topic_code)
        )
        self.refresh_editor_subjects()
        self.editor_subject_combo.setCurrentIndex(
            self.editor_subject_combo.findData(question.subject_id)
        )
        self.editor_question.setPlainText(question.text)
        for letter, field in self.editor_options.items():
            field.setPlainText(question.options.get(letter, ""))
        self.editor_correct.setCurrentText(question.correct_answer)
        for letter, checkbox in self.editor_locked_answers.items():
            checkbox.setChecked(letter in question.locked_answers)
        self.editor_source.setPlainText(question.source_reference)

    def refresh_editor_subjects(self) -> None:
        if not hasattr(self, "editor_subject_combo"):
            return
        current_id = self.editor_subject_combo.currentData()
        code = self.editor_topic_combo.currentData()
        self.editor_subject_combo.blockSignals(True)
        self.editor_subject_combo.clear()
        for subject_id, _, name, count in self.database.question_topics(
            str(code) if code else None
        ):
            self.editor_subject_combo.addItem(f"{name} ({count})", subject_id)
        index = self.editor_subject_combo.findData(current_id)
        if index >= 0:
            self.editor_subject_combo.setCurrentIndex(index)
        self.editor_subject_combo.blockSignals(False)

    def clear_question_editor(self) -> None:
        self.editing_question_id = None
        self.editor_question.clear()
        for field in self.editor_options.values():
            field.clear()
        self.editor_correct.setCurrentText("A")
        for checkbox in self.editor_locked_answers.values():
            checkbox.setChecked(False)
        self.editor_source.clear()

    def save_question(self) -> None:
        if self.editor_subject_combo.currentData() is None:
            QMessageBox.warning(
                self,
                "Chưa chọn chuyên đề",
                "Hãy tạo hoặc chọn chuyên đề trước khi lưu câu hỏi.",
            )
            return
        try:
            self.database.save_question(
                question_id=self.editing_question_id,
                topic_code=str(self.editor_topic_combo.currentData() or ""),
                subject_id=self.editor_subject_combo.currentData(),
                text=self.editor_question.toPlainText(),
                options={
                    letter: field.toPlainText()
                    for letter, field in self.editor_options.items()
                },
                correct_answer=self.editor_correct.currentText(),
                locked_answers="".join(
                    letter
                    for letter, checkbox in self.editor_locked_answers.items()
                    if checkbox.isChecked()
                ),
                source_reference=self.editor_source.toPlainText(),
            )
            self.clear_question_editor()
            self.refresh_setup_data()
        except QuizDatabaseError as exc:
            QMessageBox.warning(self, "Không thể lưu câu hỏi", str(exc))

    def delete_question(self) -> None:
        if self.editing_question_id is None:
            return
        if QMessageBox.question(
            self,
            "Xóa câu hỏi",
            "Bạn có chắc chắn muốn xóa câu hỏi đang chỉnh sửa?",
        ) != QMessageBox.StandardButton.Yes:
            return
        self.database.delete_question(self.editing_question_id)
        self.clear_question_editor()
        self.refresh_setup_data()

    def delete_questions_for_business(self) -> None:
        code = self.question_topic_filter.currentData()
        if not code:
            QMessageBox.information(
                self,
                "Chọn nghiệp vụ",
                "Hãy chọn một nghiệp vụ trong bộ lọc trước khi xóa hàng loạt.",
            )
            return
        name = self.question_topic_filter.currentText()
        if QMessageBox.question(
            self,
            "Xóa câu hỏi theo nghiệp vụ",
            f"Xóa toàn bộ câu hỏi thuộc {name}? Thao tác này không thể hoàn tác.",
        ) != QMessageBox.StandardButton.Yes:
            return
        deleted = self.database.delete_questions_bulk(str(code))
        self.refresh_setup_data()
        QMessageBox.information(
            self,
            "Đã xóa",
            f"Đã xóa {deleted} câu hỏi.",
        )

    def delete_all_questions(self) -> None:
        if QMessageBox.question(
            self,
            "Xóa toàn bộ câu hỏi",
            "Bạn có chắc chắn muốn xóa TOÀN BỘ ngân hàng câu hỏi? "
            "Thao tác này không thể hoàn tác.",
        ) != QMessageBox.StandardButton.Yes:
            return
        if QMessageBox.question(
            self,
            "Xác nhận lần cuối",
            "Xác nhận xóa toàn bộ câu hỏi?",
        ) != QMessageBox.StandardButton.Yes:
            return
        deleted = self.database.delete_questions_bulk()
        self.refresh_setup_data()
        QMessageBox.information(
            self,
            "Đã xóa",
            f"Đã xóa {deleted} câu hỏi.",
        )

    def download_excel_template(self) -> None:
        filename, _ = QFileDialog.getSaveFileName(
            self,
            "Lưu file Excel mẫu",
            "AgribankV3-Mau-Cau-Hoi.xlsx",
            "Excel Workbook (*.xlsx)",
        )
        if not filename:
            return
        source = files("agribank_v3").joinpath(
            "resources",
            "templates",
            "AgribankV3-Mau-Cau-Hoi.xlsx",
        )
        try:
            with source.open("rb") as input_file, Path(filename).open("wb") as output:
                shutil.copyfileobj(input_file, output)
            QMessageBox.information(
                self,
                "Đã tạo file mẫu",
                f"Đã lưu file Excel mẫu tại:\n{filename}",
            )
        except OSError as exc:
            QMessageBox.warning(self, "Không thể lưu file mẫu", str(exc))

    def import_questions_excel(self) -> None:
        filename, _ = QFileDialog.getOpenFileName(
            self,
            "Chọn file câu hỏi",
            "",
            "Excel Workbook (*.xlsx *.xlsm)",
        )
        if not filename:
            return
        try:
            result = import_questions_from_excel(Path(filename), self.database)
            self.refresh_setup_data()
            QMessageBox.information(
                self,
                "Nhập dữ liệu hoàn tất",
                f"Đã nhập {result.imported} câu hỏi; bỏ qua {result.skipped} dòng.",
            )
        except (QuizDatabaseError, OSError, ValueError) as exc:
            QMessageBox.warning(self, "Không thể nhập file Excel", str(exc))

    def force_sync(self) -> None:
        self.sync_button.setEnabled(False)
        try:
            self.sync_result = self.database.sync_from_access(force=True)
            self.refresh_setup_data()
            QMessageBox.information(
                self,
                "Cập nhật hoàn tất",
                f"Đã cập nhật {self.sync_result.question_count:,} câu hỏi.",
            )
        except QuizDatabaseError as exc:
            QMessageBox.warning(self, "Không thể cập nhật", str(exc))
        finally:
            self.sync_button.setEnabled(True)

    def _questions_for_random_selection(self) -> list[Question] | None:
        quotas = self.random_quotas()
        if not quotas:
            QMessageBox.information(
                self,
                "Chưa chọn chuyên đề",
                "Hãy chọn ít nhất một chuyên đề và nhập số câu cần lấy.",
            )
            return None
        subject_rows = self.database.question_topics()
        available = {
            subject_id: count
            for subject_id, _, _, count in subject_rows
        }
        subject_names = {
            subject_id: name
            for subject_id, _, name, _ in subject_rows
        }
        shortages = [
            f"{subject_names.get(subject_id, str(subject_id))}: "
            f"cần {amount}, có {available.get(subject_id, 0)}"
            for subject_id, amount in quotas.items()
            if amount > available.get(subject_id, 0)
        ]
        if shortages:
            QMessageBox.warning(
                self,
                "Không đủ câu hỏi",
                "Hạn mức vượt quá dữ liệu hiện có:\n" + "\n".join(shortages),
            )
            return None
        self.database.save_random_subject_setting(quotas)
        return self.database.questions_by_subject_quotas(quotas)

    def export_quiz_excel(self) -> None:
        if self.topic_mode.isChecked():
            data = self.topic_combo.currentData()
            if not data:
                QMessageBox.warning(
                    self,
                    "Chưa chọn chuyên đề",
                    "Hãy chọn chuyên đề cần xuất.",
                )
                return
            topic = str(data[0])
            limit = int(data[1])
            questions = self.database.questions(
                topic_name=topic,
                limit=limit,
                randomize=False,
            )
            subtitle = f"Chuyên đề: {topic} ({len(questions)} Câu)"
            filename_prefix = "AgribankV3-Chuyen-De"
        else:
            questions = self._questions_for_random_selection()
            if questions is None:
                return
            subtitle = None
            filename_prefix = "AgribankV3-De-Ngau-Nhien"
        if not questions:
            QMessageBox.warning(
                self,
                "Không có dữ liệu",
                "Không tìm thấy câu hỏi để xuất.",
            )
            return
        if self.shuffle_answer_order.isChecked():
            questions = [
                self._shuffle_question_answers(question)
                for question in questions
            ]

        default_name = (
            f"{filename_prefix}-"
            f"{datetime.now().strftime('%Y%m%d-%H%M')}.xlsx"
        )
        filename, _ = QFileDialog.getSaveFileName(
            self,
            "Xuất bộ câu hỏi sang Excel",
            default_name,
            "Excel Workbook (*.xlsx)",
        )
        if not filename:
            return
        path = Path(filename)
        if path.suffix.lower() != ".xlsx":
            path = path.with_suffix(".xlsx")
        try:
            export_questions_to_excel(
                path,
                questions,
                subtitle=subtitle,
            )
            QMessageBox.information(
                self,
                "Xuất Excel hoàn tất",
                f"Đã xuất {len(questions)} câu hỏi:\n{path}",
            )
        except (OSError, PermissionError, ValueError) as exc:
            QMessageBox.warning(
                self,
                "Không thể xuất Excel",
                str(exc),
            )

    def start_quiz(self) -> None:
        if self.topic_mode.isChecked():
            data = self.topic_combo.currentData()
            topic = data[0] if data else None
            limit = int(data[1]) if data else 0
            questions = self.database.questions(topic, limit, False)
            session_topic = topic or "Tất cả chuyên đề"
        else:
            questions = self._questions_for_random_selection()
            if questions is None:
                return
            session_topic = "Đề ngẫu nhiên theo chuyên đề"
        if not questions:
            QMessageBox.warning(self, "Không có dữ liệu", "Không tìm thấy câu hỏi.")
            return
        if self.shuffle_answer_order.isChecked():
            questions = [self._shuffle_question_answers(question) for question in questions]
        time_limit = self.time_limit.value() if self.limited_time.isChecked() else 0
        self.session = QuizSession(
            questions=questions,
            employee_name=self.employee_combo.currentText(),
            topic_name=session_topic,
            time_limit_minutes=time_limit,
        )
        self.answer_progress.setRange(0, len(questions))
        self.jump_question_spin.setRange(1, len(questions))
        self.jump_question_spin.setValue(1)
        self.deadline = (
            time.monotonic() + time_limit * 60
            if time_limit
            else 0
        )
        self.pages.setTabEnabled(1, True)
        self.pages.setCurrentIndex(1)
        self.timer.start()
        self.load_question()

    @staticmethod
    def _shuffle_question_answers(question: Question) -> Question:
        locked = {
            letter
            for letter in question.locked_answers
            if letter in question.options
        }
        movable_letters = [
            letter for letter in question.options if letter not in locked
        ]
        shuffled_items = [
            (letter, question.options[letter]) for letter in movable_letters
        ]
        random.shuffle(shuffled_items)
        options = {
            letter: question.options[letter]
            for letter in locked
        }
        source_for_display = {letter: letter for letter in locked}
        for display_letter, (source_letter, source_value) in zip(
            movable_letters,
            shuffled_items,
            strict=True,
        ):
            options[display_letter] = source_value
            source_for_display[display_letter] = source_letter
        options = {
            letter: options[letter]
            for letter in question.options
        }
        correct_answer = next(
            display_letter
            for display_letter, source_letter in source_for_display.items()
            if source_letter == question.correct_answer
        )
        return Question(
            id=question.id,
            legacy_id=question.legacy_id,
            number=question.number,
            text=question.text,
            options=options,
            correct_answer=correct_answer,
            topic_code=question.topic_code,
            topic_name=question.topic_name,
            source_reference=question.source_reference,
            subject_id=question.subject_id,
            subject_name=question.subject_name,
            locked_answers=question.locked_answers,
        )

    def load_question(self) -> None:
        if self.session is None:
            return
        question = self.session.current_question
        total = len(self.session.questions)
        self.question_progress_label.setText(
            f"Câu {self.session.current_index + 1}/{total}"
        )
        self.question_number_badge.setText(str(self.session.current_index + 1))
        self.jump_question_spin.blockSignals(True)
        self.jump_question_spin.setValue(self.session.current_index + 1)
        self.jump_question_spin.blockSignals(False)
        self.topic_label.setText(
            f"Nghiệp vụ: {question.topic_name} • "
            f"Chuyên đề: {question.subject_name or 'Chưa phân loại'} • "
            f"Mã: {question.topic_code} "
            f"• Câu gốc: {question.number}"
        )
        self.question_label.setText(question.text)
        self.answer_progress.setValue(len(self.session.answers))
        self.answer_progress.setFormat(
            f"Đã trả lời {len(self.session.answers)}/{total} câu"
        )

        self.answer_group.setExclusive(False)
        for button in self.answer_buttons:
            button.setChecked(False)
            button.hide()
            button.setEnabled(True)
            button.setStyleSheet("")
        self.answer_group.setExclusive(True)

        self.answer_letters = list(question.options)
        selected = self.session.selected_answer(question)
        for index, letter in enumerate(self.answer_letters):
            button = self.answer_buttons[index]
            button.setText(question.options[letter])
            button.setProperty("answerLetter", letter)
            button.setChecked(selected == letter)
            button.show()

        checked = question.id in self.session.checked_questions
        self.check_button.setEnabled(not checked)
        self.feedback_label.hide()
        self.hint_frame.hide()
        if checked:
            for button in self.answer_buttons:
                if button.isHidden():
                    continue
                letter = str(button.property("answerLetter"))
                button.setEnabled(False)
                if letter == question.correct_answer:
                    button.setText(f"{question.options[letter]}    ✓")
                    button.setStyleSheet(
                        "QRadioButton { background: #dcfce7; color: #166534; "
                        "border: 1px solid #22c55e; border-radius: 8px; "
                        "padding: 11px 12px; font-size: 14px; font-weight: 500; }"
                    )
                elif letter == selected:
                    button.setStyleSheet(
                        "QRadioButton { background: #fef2f2; color: #991b1b; "
                        "border: 1px solid #ef4444; border-radius: 8px; "
                        "padding: 11px 12px; font-size: 14px; }"
                    )
            is_correct = selected == question.correct_answer
            self.feedback_label.setText(
                "✓ Trả lời đúng"
                if is_correct
                else f"✗ Chưa đúng. Đáp án đúng là {question.correct_answer}."
            )
            self.feedback_label.setStyleSheet(
                "color: #257047; font-weight: 700;"
                if is_correct
                else "color: #9b2c2c; font-weight: 700;"
            )
            if question.source_reference:
                self.source_label.setText(
                    f"Đáp án: {question.correct_answer} - "
                    f"{question.source_reference}"
                )
                self.hint_frame.show()

        self.previous_button.setEnabled(self.session.current_index > 0)
        self.next_button.setText(
            "Hoàn thành"
            if self.session.current_index == total - 1
            else "Câu tiếp →"
        )

    def answer_selected(self, button: QRadioButton) -> None:
        if self.session is None:
            return
        letter = str(button.property("answerLetter"))
        if letter:
            self.session.answer(letter)
            self.answer_progress.setValue(len(self.session.answers))
            self.answer_progress.setFormat(
                f"Đã trả lời {len(self.session.answers)}/"
                f"{len(self.session.questions)} câu"
            )

    def check_answer(self) -> None:
        if self.session is None:
            return
        result = self.session.check_current()
        if result is None:
            QMessageBox.information(
                self,
                "Chưa chọn đáp án",
                "Hãy chọn một đáp án trước khi kiểm tra.",
            )
            return
        self.load_question()

    def previous_question(self) -> None:
        if self.session and self.session.current_index > 0:
            self.session.current_index -= 1
            self.load_question()

    def jump_to_question(self) -> None:
        if self.session is None:
            return
        target_index = self.jump_question_spin.value() - 1
        if 0 <= target_index < len(self.session.questions):
            self.session.current_index = target_index
            self.load_question()

    def next_question(self) -> None:
        if self.session is None:
            return
        if self.session.current_index >= len(self.session.questions) - 1:
            self.confirm_finish()
            return
        self.session.current_index += 1
        self.load_question()

    def update_timer(self) -> None:
        if self.session is None:
            return
        if not self.deadline:
            elapsed = int(
                (datetime.now() - self.session.started_at).total_seconds()
            )
            self.timer_label.setText(f"Đã làm: {elapsed // 60:02}:{elapsed % 60:02}")
            return
        remaining = max(0, int(self.deadline - time.monotonic()))
        self.timer_label.setText(
            f"Còn lại: {remaining // 60:02}:{remaining % 60:02}"
        )
        if remaining <= 0:
            self.finish_quiz()

    def confirm_finish(self) -> None:
        if self.session is None:
            return
        unanswered = len(self.session.questions) - len(self.session.answers)
        message = "Bạn muốn kết thúc bài làm?"
        if unanswered:
            message += f"\nCòn {unanswered} câu chưa trả lời."
        if (
            QMessageBox.question(
                self,
                "Kết thúc bài",
                message,
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            )
            == QMessageBox.StandardButton.Yes
        ):
            self.finish_quiz()

    def finish_quiz(self) -> None:
        if self.session is None or self.pages.currentIndex() == 2:
            return
        self.timer.stop()
        result = self.session.finish()
        attempt_id = self.database.save_attempt(self.session, result)
        self.result_percentage.setText(f"{result.percentage:.2f}%")
        self.result_summary.setText(
            f"<b>Xếp loại: {result.rating}</b><br>"
            f"Đúng: {result.correct}/{result.total} • "
            f"Sai: {result.incorrect} • Chưa trả lời: {result.unanswered}<br>"
            f"Thời gian: {result.duration_seconds // 60:02}:"
            f"{result.duration_seconds % 60:02} • Mã lượt: {attempt_id}"
        )
        self.review_table.setRowCount(len(self.session.questions))
        for row, question in enumerate(self.session.questions):
            selected = self.session.answers.get(question.id, "")
            values = (
                str(row + 1),
                selected or "—",
                question.correct_answer,
                "Đúng" if selected == question.correct_answer else "Sai/Chưa trả lời",
            )
            for column, value in enumerate(values):
                item = QTableWidgetItem(value)
                if column == 3:
                    item.setForeground(
                        Qt.GlobalColor.darkGreen
                        if selected == question.correct_answer
                        else Qt.GlobalColor.darkRed
                    )
                self.review_table.setItem(row, column, item)
        self.pages.setTabEnabled(2, True)
        self.pages.setCurrentIndex(2)

    def open_result_question(self, row: int, _column: int) -> None:
        if self.session is None or not (0 <= row < len(self.session.questions)):
            return
        question = self.session.questions[row]
        dialog = QuestionReviewDialog(
            question,
            self.session.answers.get(question.id, ""),
            row + 1,
            self,
        )
        dialog.exec()

    def schedule_answer_lookup(self) -> None:
        if hasattr(self, "answer_lookup_timer"):
            self.answer_lookup_timer.start()

    def search_answers(self) -> None:
        if hasattr(self, "answer_lookup_timer"):
            self.answer_lookup_timer.stop()
        keyword = self.answer_lookup_edit.text().strip()
        if len(keyword) < 2:
            self._clear_answer_lookup_results()
            self.answer_lookup_status.setText(
                "Nhập ít nhất 2 ký tự để tra cứu đáp án."
            )
            self.answer_lookup_badge.setText("0 kết quả")
            return
        questions = self.database.search_questions(keyword, limit=200)
        self._render_answer_lookup_results(questions)
        safe_keyword = escape(keyword)
        count_text = f"{len(questions):,}".replace(",", ".")
        suffix = " Kết quả đang giới hạn 200 dòng." if len(questions) >= 200 else ""
        self.answer_lookup_status.setText(
            "Tìm thấy "
            f"<b style='color:#8B1743;'>{count_text}</b>"
            f" câu hỏi phù hợp với "
            f"<b style='color:#8B1743;'>“{safe_keyword}”</b>."
            + suffix
        )
        self.answer_lookup_badge.setText(f"{count_text} kết quả")

    def _clear_answer_lookup_results(self) -> None:
        self.answer_lookup_questions = []
        self.answer_lookup_result_items = []
        self.answer_lookup_selected_row = -1
        self.answer_lookup_table.setRowCount(0)
        while self.answer_lookup_results_layout.count():
            item = self.answer_lookup_results_layout.takeAt(0)
            widget = item.widget()
            if widget is not None:
                widget.deleteLater()
        self.answer_lookup_empty_label = QLabel(
            "Chưa có kết quả. Nhập từ khóa để bắt đầu tra cứu."
        )
        self.answer_lookup_empty_label.setObjectName("AnswerLookupEmpty")
        self.answer_lookup_empty_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.answer_lookup_results_layout.addWidget(self.answer_lookup_empty_label)
        self.answer_lookup_results_layout.addStretch()

    def _render_answer_lookup_results(self, questions: list[Question]) -> None:
        self._clear_answer_lookup_results()
        self.answer_lookup_questions = questions
        self.answer_lookup_table.setRowCount(len(questions))
        while self.answer_lookup_results_layout.count():
            item = self.answer_lookup_results_layout.takeAt(0)
            widget = item.widget()
            if widget is not None:
                widget.deleteLater()
        if not questions:
            no_result = QLabel("Không tìm thấy câu hỏi phù hợp.")
            no_result.setObjectName("AnswerLookupEmpty")
            no_result.setAlignment(Qt.AlignmentFlag.AlignCenter)
            self.answer_lookup_results_layout.addWidget(no_result)
            self.answer_lookup_results_layout.addStretch()
            return
        for row, question in enumerate(questions):
            correct_text = question.options.get(question.correct_answer, "")
            table_item = QTableWidgetItem(
                f"{question.text}\n\n"
                f"Đáp án đúng {question.correct_answer}: {correct_text}"
            )
            table_item.setData(Qt.ItemDataRole.UserRole, question.id)
            self.answer_lookup_table.setItem(row, 0, table_item)

            result_item = AnswerLookupResultItem(
                row,
                question,
                correct_text,
                self.answer_lookup_results_widget,
            )
            result_item.selected.connect(self.select_answer_lookup_result)
            result_item.double_clicked.connect(self.open_answer_lookup_question)
            self.answer_lookup_result_items.append(result_item)
            self.answer_lookup_results_layout.addWidget(result_item)
        self.answer_lookup_results_layout.addStretch()
        self.select_answer_lookup_result(0)

    def select_answer_lookup_result(self, row: int) -> None:
        if not (0 <= row < len(self.answer_lookup_result_items)):
            return
        self.answer_lookup_selected_row = row
        for index, item in enumerate(self.answer_lookup_result_items):
            item.set_selected(index == row)

    def clear_answer_lookup(self) -> None:
        if hasattr(self, "answer_lookup_timer"):
            self.answer_lookup_timer.stop()
        self.answer_lookup_edit.blockSignals(True)
        self.answer_lookup_edit.clear()
        self.answer_lookup_edit.blockSignals(False)
        self._clear_answer_lookup_results()
        self.answer_lookup_status.setText(
            "Nhập từ khóa để hệ thống tự tra cứu. Nhấp đúp vào dòng để xem chi tiết."
        )
        self.answer_lookup_badge.setText("0 kết quả")
        self.answer_lookup_edit.setFocus()

    def open_answer_lookup_question(self, row: int | None = None, *_: object) -> None:
        if row is None or not isinstance(row, int):
            row = self.answer_lookup_selected_row
        if not (0 <= row < len(self.answer_lookup_questions)):
            return
        self.select_answer_lookup_result(row)
        question_id = self.answer_lookup_questions[row].id
        question = self.database.question_by_id(int(question_id))
        if question is None:
            QMessageBox.warning(
                self,
                "Tra cứu đáp án",
                "Câu hỏi này không còn tồn tại trong dữ liệu hiện tại.",
            )
            return
        dialog = QuestionReviewDialog(
            question,
            "",
            row + 1,
            self,
        )
        dialog.exec()

    def reset_to_setup(self) -> None:
        self.timer.stop()
        self.session = None
        self.pages.setTabEnabled(1, False)
        self.pages.setTabEnabled(2, False)
        self.pages.setCurrentIndex(0)

    def request_close(self) -> None:
        if self.pages.currentIndex() == 1 and self.session is not None:
            if (
                QMessageBox.question(
                    self,
                    "Đóng trắc nghiệm",
                    "Bài làm chưa kết thúc. Bạn vẫn muốn đóng?",
                    QMessageBox.StandardButton.Yes
                    | QMessageBox.StandardButton.No,
                )
                != QMessageBox.StandardButton.Yes
            ):
                return
        self.timer.stop()
        self.close_requested.emit()


# Giữ tên cũ để mã tích hợp bên ngoài không bị lỗi khi nâng cấp.
QuizDialog = QuizWidget
