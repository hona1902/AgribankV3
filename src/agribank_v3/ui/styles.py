from importlib.resources import files


def _style_icon(name: str) -> str:
    return str(
        files("agribank_v3").joinpath("resources", "icons", name)
    ).replace("\\", "/")


APP_STYLESHEET = """
QWidget {
    color: #17202a;
}
QTabWidget::pane {
    background: #f7f9fb;
    border: 1px solid #e1e6eb;
    border-radius: 0 0 10px 10px;
}
QTabBar::tab {
    color: #66717c;
    background: #f4f6f8;
    border: 1px solid #e3e7eb;
    border-bottom: none;
    padding: 11px 20px;
    min-width: 110px;
}
QTabBar::tab:selected {
    color: #5f1732;
    background: white;
    border-bottom: 3px solid #a11f4d;
    font-weight: 700;
}
QTabBar::tab:hover:!selected {
    color: #831f41;
    background: #faedf2;
}
QGroupBox {
    background: white;
    border: 1px solid #dfe5ea;
    border-radius: 9px;
    margin-top: 12px;
    padding: 12px 10px 10px 10px;
    font-weight: 650;
}
QGroupBox::title {
    color: #26313b;
    subcontrol-origin: margin;
    left: 12px;
    padding: 0 6px;
}
QGroupBox#SetupGroup {
    border: 1px solid #e1e5ea;
    border-radius: 8px;
    margin-top: 0;
    padding: 0;
}
QGroupBox#SetupGroup::title {
    color: #27313b;
    left: 10px;
    padding: 0 5px;
    font-weight: 700;
}
QGroupBox#RightSetupCard {
    background: #ffffff;
    border: 1px solid #e1e5ea;
    border-radius: 10px;
    margin-top: 0;
    padding: 0;
}
QGroupBox#AdminPanel {
    background: #ffffff;
    border: 1px solid #dce3e8;
    border-radius: 10px;
    margin-top: 12px;
    padding: 8px;
}
QGroupBox#AdminPanel::title {
    color: #3d1626;
    font-size: 14px;
    font-weight: 750;
}
QFrame#AdminEditor {
    background: #f7f9fb;
    border: 1px solid #dfe5ea;
    border-radius: 8px;
}
QLabel#AdminEditorTitle {
    color: #3d1626;
    font-size: 14px;
    font-weight: 700;
}
QLabel#SectionTitle {
    color: #1f2937;
    font-size: 15px;
    font-weight: 800;
}
QLineEdit, QTextEdit, QComboBox, QSpinBox {
    background: white;
    border: 1px solid #d7dee5;
    border-radius: 6px;
    padding: 6px 9px;
    selection-background-color: #a72a53;
    min-height: 20px;
}
QLineEdit:focus, QTextEdit:focus, QComboBox:focus, QSpinBox:focus {
    border: 1px solid #a72a53;
}
QTableWidget, QTreeWidget {
    background: white;
    alternate-background-color: #fafbfc;
    border: 1px solid #dfe5ea;
    border-radius: 6px;
    gridline-color: #e6eaee;
}
QTreeWidget#QuotaTree::item {
    min-height: 38px;
    color: #1f2937;
}
QTreeWidget#QuotaTree::item:hover {
    background: #fff5f8;
}
QTreeWidget#QuotaTree::item:selected {
    color: #1f2937;
    background: #f9e8ee;
}
QTreeWidget#QuotaTree::branch {
    image: none;
    border-image: none;
}
QHeaderView::section {
    color: #38434d;
    background: #f3f5f7;
    border: none;
    border-right: 1px solid #dde3e8;
    border-bottom: 1px solid #d7dee5;
    padding: 7px 6px;
    font-weight: 700;
}
QRadioButton {
    spacing: 8px;
    padding: 4px 2px;
}
QRadioButton::indicator {
    width: 15px;
    height: 15px;
    border: 1px solid #aeb8c2;
    border-radius: 11px;
    background: white;
}
QRadioButton::indicator:hover {
    border-color: #a11f4d;
}
QRadioButton::indicator:checked {
    border: 4px solid #a11f4d;
    border-radius: 11px;
    background: white;
}
QRadioButton::indicator:unchecked {
    image: url("__CIRCLE_OFF__");
    border: none;
    background: transparent;
}
QRadioButton::indicator:checked {
    image: url("__RADIO_ON__");
    border: none;
    background: transparent;
}
QCheckBox::indicator, QTreeView::indicator {
    width: 18px;
    height: 18px;
}
QCheckBox::indicator:unchecked, QTreeView::indicator:unchecked {
    image: url("__CIRCLE_OFF__");
}
QCheckBox::indicator:checked, QTreeView::indicator:checked {
    image: url("__CHECK_ON__");
}
QToolTip {
    color: #26313b;
    background: white;
    border: 1px solid #cfd7de;
    padding: 5px;
}
QMainWindow, QWidget#AppRoot {
    background: #f6f7f9;
}
QFrame#Sidebar {
    background: #831f41;
    border: none;
}
QLabel#BrandTitle {
    color: white;
    font-size: 22px;
    font-weight: 700;
}
QLabel#BrandSubtitle {
    color: #f2ccd8;
    font-size: 11px;
}
QPushButton#NavButton {
    color: #f8edf1;
    background: transparent;
    border: none;
    border-radius: 8px;
    padding: 11px 14px;
    text-align: left;
    font-size: 13px;
}
QPushButton#NavButton:hover {
    background: rgba(255, 255, 255, 0.10);
}
QPushButton#NavButton:checked {
    color: white;
    background: #a72a53;
    font-weight: 600;
}
QLabel#PageTitle {
    color: #3d1626;
    font-size: 24px;
    font-weight: 700;
}
QLabel#MutedText {
    color: #68737d;
}
QFrame#FeatureCard, QFrame#MetricCard, QFrame#WelcomeCard {
    background: white;
    border: 1px solid #e1e6e9;
    border-radius: 12px;
}
QFrame#FeatureCard:hover {
    border: 1px solid #b75a79;
}
QLabel#CardTitle {
    color: #34131f;
    font-size: 15px;
    font-weight: 650;
}
QLabel#MetricValue {
    color: #831f41;
    font-size: 26px;
    font-weight: 750;
}
QLabel#QuizQuestion {
    color: #0f172a;
    font-size: 17px;
    font-weight: 700;
    padding: 3px 0 10px 0;
}
QRadioButton#QuizOption {
    background: #ffffff;
    border: 1px solid #dce3ea;
    border-radius: 8px;
    padding: 11px 12px;
    font-size: 14px;
}
QRadioButton#QuizOption:hover {
    background: #f8fbff;
    border-color: #93b4dd;
}
QRadioButton#QuizOption:checked {
    color: #123f7a;
    background: #dbeafe;
    border: 2px solid #2563eb;
    padding: 10px 11px;
    font-weight: 600;
}
QRadioButton#QuizOption::indicator:checked {
    image: none;
    border: 2px solid #ffffff;
    border-radius: 7px;
    background: #2563eb;
    width: 14px;
    height: 14px;
}
QPushButton#PrimaryButton {
    color: white;
    background: #931f49;
    border: none;
    border-radius: 7px;
    padding: 10px 18px;
    font-weight: 600;
    min-height: 18px;
}
QPushButton#DangerButton {
    color: #9b1c1c;
    background: #fff0f0;
    border: 1px solid #efc8c8;
    border-radius: 7px;
    padding: 9px 15px;
    font-weight: 600;
}
QPushButton#DangerButton:hover {
    color: #ffffff;
    background: #b42323;
    border-color: #b42323;
}
QPushButton#DangerButton:disabled {
    color: #aeb5bb;
    background: #f4f5f6;
    border-color: #e1e4e7;
}
QPushButton#PrimaryButton:hover {
    background: #ad2c57;
}
QPushButton#SecondaryButton {
    color: #831f41;
    background: white;
    border: 1px solid #d9dfe5;
    border-radius: 7px;
    padding: 9px 14px;
    font-weight: 600;
}
QPushButton#SecondaryButton:hover {
    background: #fff4f7;
    border-color: #bd6b87;
}
QFrame#BrandButton {
    background: transparent;
    border: none;
    border-radius: 8px;
    padding: 3px;
}
QFrame#BrandButton:hover {
    background: rgba(255, 255, 255, 0.12);
}
QPushButton#SidebarExcelButton {
    color: white;
    background: #b32b59;
    border: 1px solid rgba(255, 255, 255, 0.16);
    border-radius: 7px;
    min-height: 36px;
    padding: 0 10px;
    font-weight: 650;
}
QPushButton#SidebarExcelButton:hover {
    background: #c63867;
}
QStatusBar {
    background: white;
    color: #68737d;
    border-top: 1px solid #e3e7ea;
}
"""

APP_STYLESHEET = (
    APP_STYLESHEET
    .replace("__CIRCLE_OFF__", _style_icon("circle_off.svg"))
    .replace("__RADIO_ON__", _style_icon("radio_on.svg"))
    .replace("__CHECK_ON__", _style_icon("check_on.svg"))
)
