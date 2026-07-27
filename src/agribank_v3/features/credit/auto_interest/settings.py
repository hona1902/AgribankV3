from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

from agribank_v3.runtime_paths import application_root
from agribank_v3.settings import AppSettingsDatabase, SettingsDatabaseError


AUTO_INTEREST_REPORT_FOLDER_KEY = "auto_interest_report_folder"
AUTO_INTEREST_OUTPUT_FOLDER_KEY = "auto_interest_output_folder"
AUTO_INTEREST_LOAN_FOLDER_KEY = "auto_interest_loan_folder"
AUTO_INTEREST_DEPOSIT_FOLDER_KEY = "auto_interest_deposit_folder"
AUTO_INTEREST_BACKUP_FOLDER_KEY = "auto_interest_backup_folder"


@dataclass(frozen=True, slots=True)
class AutoInterestSettings:
    report_folder: Path
    output_folder: Path
    loan_folder: Path | None = None
    deposit_folder: Path | None = None
    backup_folder: Path | None = None


def default_auto_interest_settings() -> AutoInterestSettings:
    root = application_root()
    return AutoInterestSettings(
        report_folder=root / "KetQua" / "BaoCaoThuLaiBanTuDong",
        output_folder=root / "KetQua" / "ThuLaiBanTuDong",
        loan_folder=root / "DuLieu" / "ThuLaiBanTuDong" / "Loan",
        deposit_folder=root / "DuLieu" / "ThuLaiBanTuDong" / "TienGui",
        backup_folder=root / "data" / "backups" / "auto_interest",
    )


def load_auto_interest_settings(
    database_path: Path | None = None,
) -> AutoInterestSettings:
    defaults = default_auto_interest_settings()
    database = AppSettingsDatabase(database_path)
    report_folder = database.load_preference(
        AUTO_INTEREST_REPORT_FOLDER_KEY,
        str(defaults.report_folder),
    )
    output_folder = database.load_preference(
        AUTO_INTEREST_OUTPUT_FOLDER_KEY,
        str(defaults.output_folder),
    )
    loan_folder = database.load_preference(
        AUTO_INTEREST_LOAN_FOLDER_KEY,
        str(defaults.loan_folder),
    )
    deposit_folder = database.load_preference(
        AUTO_INTEREST_DEPOSIT_FOLDER_KEY,
        str(defaults.deposit_folder),
    )
    backup_folder = database.load_preference(
        AUTO_INTEREST_BACKUP_FOLDER_KEY,
        str(defaults.backup_folder),
    )
    return AutoInterestSettings(
        report_folder=Path(report_folder),
        output_folder=Path(output_folder),
        loan_folder=Path(loan_folder),
        deposit_folder=Path(deposit_folder),
        backup_folder=Path(backup_folder),
    )


def save_auto_interest_settings(
    settings: AutoInterestSettings,
    database_path: Path | None = None,
) -> AutoInterestSettings:
    database = AppSettingsDatabase(database_path)
    try:
        database.save_preference(
            AUTO_INTEREST_REPORT_FOLDER_KEY,
            str(settings.report_folder),
        )
        database.save_preference(
            AUTO_INTEREST_OUTPUT_FOLDER_KEY,
            str(settings.output_folder),
        )
        if settings.loan_folder is not None:
            database.save_preference(
                AUTO_INTEREST_LOAN_FOLDER_KEY,
                str(settings.loan_folder),
            )
        if settings.deposit_folder is not None:
            database.save_preference(
                AUTO_INTEREST_DEPOSIT_FOLDER_KEY,
                str(settings.deposit_folder),
            )
        if settings.backup_folder is not None:
            database.save_preference(
                AUTO_INTEREST_BACKUP_FOLDER_KEY,
                str(settings.backup_folder),
            )
    except SettingsDatabaseError:
        raise
    return settings
