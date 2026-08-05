from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

from agribank_v3.features.credit.summary.credit_report import CreditReportRepository
from agribank_v3.features.credit.summary.customer.repository import CustomerRepository
from agribank_v3.features.credit.summary.repository import SummaryRepository
from agribank_v3.settings import AppSettingsDatabase


class UserDatabaseEnsureError(RuntimeError):
    pass


@dataclass(frozen=True, slots=True)
class UserDatabaseEnsureResult:
    name: str
    path: Path
    existed_before: bool
    ok: bool
    error: str = ""


def ensure_user_databases(
    settings_database: AppSettingsDatabase | None = None,
    *,
    strict: bool = False,
) -> tuple[UserDatabaseEnsureResult, ...]:
    database = settings_database or AppSettingsDatabase()
    main_database_path = database.database_path
    results: list[UserDatabaseEnsureResult] = []
    for name, ensure in (
        ("CreditSummary.db", lambda: SummaryRepository(main_database_path).database_path),
        ("Customer.db", lambda: CustomerRepository(main_database_path).database_path),
        ("Credit.db", lambda: CreditReportRepository(main_database_path).database_path),
    ):
        expected_path = main_database_path.parent / name
        existed_before = expected_path.is_file()
        try:
            path = Path(ensure())
        except Exception as exc:
            if strict:
                raise UserDatabaseEnsureError(f"Khong the khoi tao {name}: {exc}") from exc
            results.append(
                UserDatabaseEnsureResult(
                    name=name,
                    path=expected_path,
                    existed_before=existed_before,
                    ok=False,
                    error=str(exc),
                )
            )
            continue
        results.append(
            UserDatabaseEnsureResult(
                name=name,
                path=path,
                existed_before=existed_before,
                ok=path.is_file(),
            )
        )
    return tuple(results)
