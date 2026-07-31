from __future__ import annotations

from contextlib import closing, contextmanager
from datetime import datetime
from pathlib import Path
import re
import sqlite3
from typing import Iterator

from agribank_v3.features.settings.unit_directory.legacy_seed import (
    LEGACY_BRANCHES,
    LEGACY_OFFICES,
)
from agribank_v3.features.settings.unit_directory.models import (
    AppUnitSettings,
    BranchDirectoryEntry,
    HEAD_OFFICE,
    OFFICE_TYPES,
    OfficeDirectoryEntry,
    OTHER_OFFICE,
    TRANSACTION_OFFICE,
)


class UnitDirectoryError(RuntimeError):
    pass


def normalize_branch_code(value: object) -> str:
    return "" if value is None else str(value).strip()


def normalize_trctcd(value: object) -> str:
    text = "" if value is None else str(value).strip().replace("'", "")
    if re.fullmatch(r"\d+\.0+", text):
        text = text.split(".", 1)[0]
    if text.isdigit() and len(text) < 2:
        text = text.zfill(2)
    return text


def build_office_code(branch_code: object, trctcd: object) -> str:
    branch = normalize_branch_code(branch_code)
    code = normalize_trctcd(trctcd)
    if not branch:
        return ""
    return f"{branch}-{code or 'UNKNOWN'}"


def default_office_type(trctcd: object) -> str:
    code = normalize_trctcd(trctcd)
    if not code:
        return OTHER_OFFICE
    return HEAD_OFFICE if code == "00" else TRANSACTION_OFFICE


def ensure_unit_directory_schema(database: sqlite3.Connection) -> None:
    database.executescript(
        """
        CREATE TABLE IF NOT EXISTS branch_directory (
            branch_code TEXT PRIMARY KEY,
            branch_name TEXT NOT NULL,
            short_name TEXT,
            display_name TEXT,
            province_name TEXT,
            is_active INTEGER NOT NULL DEFAULT 1 CHECK(is_active IN (0, 1)),
            sort_order INTEGER,
            created_at TEXT,
            updated_at TEXT,
            updated_by TEXT
        );

        CREATE TABLE IF NOT EXISTS office_directory (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            branch_code TEXT NOT NULL,
            trctcd TEXT NOT NULL,
            office_code TEXT NOT NULL,
            office_name TEXT NOT NULL,
            short_name TEXT,
            office_type TEXT NOT NULL,
            is_active INTEGER NOT NULL DEFAULT 1 CHECK(is_active IN (0, 1)),
            sort_order INTEGER,
            created_at TEXT,
            updated_at TEXT,
            updated_by TEXT,
            UNIQUE(branch_code, trctcd),
            UNIQUE(office_code),
            FOREIGN KEY(branch_code) REFERENCES branch_directory(branch_code)
        );

        CREATE INDEX IF NOT EXISTS idx_office_directory_branch
            ON office_directory(branch_code, is_active, sort_order, trctcd);

        CREATE TABLE IF NOT EXISTS app_unit_settings (
            id INTEGER PRIMARY KEY CHECK(id = 1),
            home_branch_code TEXT,
            default_office_code TEXT,
            organization_name TEXT,
            updated_at TEXT,
            updated_by TEXT
        );
        """
    )
    seed_legacy_unit_directory(database)


def seed_legacy_unit_directory(database: sqlite3.Connection) -> None:
    now = _now()
    for branch in LEGACY_BRANCHES:
        code = normalize_branch_code(branch["branch_code"])
        short_name = str(branch.get("short_name") or "").strip()
        display = f"{code} - {short_name}" if short_name else ""
        database.execute(
            """
            INSERT OR IGNORE INTO branch_directory(
                branch_code, branch_name, short_name, display_name,
                province_name, is_active, sort_order, created_at, updated_at, updated_by
            )
            VALUES (?, ?, ?, ?, ?, 1, ?, ?, ?, 'migration')
            """,
            (
                code,
                str(branch.get("branch_name") or "").strip(),
                short_name,
                display,
                str(branch.get("province_name") or "").strip(),
                branch.get("sort_order"),
                now,
                now,
            ),
        )
    for office in LEGACY_OFFICES:
        branch_code = normalize_branch_code(office["branch_code"])
        trctcd = normalize_trctcd(office["trctcd"])
        office_code = build_office_code(branch_code, trctcd)
        database.execute(
            """
            INSERT OR IGNORE INTO office_directory(
                branch_code, trctcd, office_code, office_name, short_name,
                office_type, is_active, sort_order, created_at, updated_at, updated_by
            )
            VALUES (?, ?, ?, ?, ?, ?, 1, ?, ?, ?, 'migration')
            """,
            (
                branch_code,
                trctcd,
                office_code,
                str(office.get("office_name") or "").strip(),
                str(office.get("short_name") or "").strip(),
                str(office.get("office_type") or default_office_type(trctcd)).strip(),
                office.get("sort_order"),
                now,
                now,
            ),
        )
    database.execute(
        """
        INSERT OR IGNORE INTO app_unit_settings(
            id, home_branch_code, default_office_code, organization_name, updated_at, updated_by
        )
        VALUES (1, '5491', '5491-00', 'Agribank Chi nhánh Lộc Phát Lâm Đồng', ?, 'migration')
        """,
        (now,),
    )


class UnitDirectoryRepository:
    def __init__(self, database_path: Path | str) -> None:
        self.database_path = Path(database_path)
        self.database_path.parent.mkdir(parents=True, exist_ok=True)
        self.initialize_schema()

    def connect(self) -> sqlite3.Connection:
        connection = sqlite3.connect(self.database_path, timeout=15)
        connection.row_factory = sqlite3.Row
        connection.execute("PRAGMA foreign_keys = ON")
        connection.execute("PRAGMA journal_mode = WAL")
        connection.execute("PRAGMA busy_timeout = 15000")
        return connection

    @contextmanager
    def _database(self) -> Iterator[sqlite3.Connection]:
        database = self.connect()
        try:
            with database:
                yield database
        finally:
            database.close()

    def initialize_schema(self) -> None:
        try:
            with closing(self.connect()) as database:
                ensure_unit_directory_schema(database)
                database.commit()
        except sqlite3.Error as exc:
            raise UnitDirectoryError(f"Không thể khởi tạo danh mục đơn vị: {exc}") from exc

    def list_branches(self, *, active_only: bool = False) -> list[BranchDirectoryEntry]:
        clause = "WHERE is_active = 1" if active_only else ""
        with self._database() as database:
            rows = database.execute(
                f"""
                SELECT *
                FROM branch_directory
                {clause}
                ORDER BY COALESCE(sort_order, 999999), branch_code COLLATE NOCASE
                """
            ).fetchall()
        return [_branch_from_row(row) for row in rows]

    def get_branch(self, branch_code: object) -> BranchDirectoryEntry | None:
        code = normalize_branch_code(branch_code)
        if not code:
            return None
        with self._database() as database:
            row = database.execute(
                "SELECT * FROM branch_directory WHERE branch_code = ?",
                (code,),
            ).fetchone()
        return _branch_from_row(row) if row is not None else None

    def save_branch(self, entry: BranchDirectoryEntry) -> BranchDirectoryEntry:
        code = _validate_branch_code(entry.branch_code)
        name = _validate_required(entry.branch_name, "Tên chi nhánh")
        now = _now()
        with self._database() as database:
            existing = database.execute(
                "SELECT created_at FROM branch_directory WHERE branch_code = ?",
                (code,),
            ).fetchone()
            created_at = str(existing["created_at"] or now) if existing else now
            database.execute(
                """
                INSERT INTO branch_directory(
                    branch_code, branch_name, short_name, display_name,
                    province_name, is_active, sort_order, created_at, updated_at, updated_by
                )
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                ON CONFLICT(branch_code) DO UPDATE SET
                    branch_name = excluded.branch_name,
                    short_name = excluded.short_name,
                    display_name = excluded.display_name,
                    province_name = excluded.province_name,
                    is_active = excluded.is_active,
                    sort_order = excluded.sort_order,
                    updated_at = excluded.updated_at,
                    updated_by = excluded.updated_by
                """,
                (
                    code,
                    name,
                    _clean(entry.short_name),
                    _clean(entry.display_name),
                    _clean(entry.province_name),
                    1 if entry.is_active else 0,
                    entry.sort_order,
                    created_at,
                    now,
                    _clean(entry.updated_by),
                ),
            )
        saved = self.get_branch(code)
        if saved is None:
            raise UnitDirectoryError("Không thể lưu chi nhánh.")
        return saved

    def create_branch(self, entry: BranchDirectoryEntry) -> BranchDirectoryEntry:
        code = _validate_branch_code(entry.branch_code)
        if self.get_branch(code) is not None:
            raise UnitDirectoryError("Mã chi nhánh đã tồn tại.")
        return self.save_branch(entry)

    def set_branch_active(self, branch_code: object, active: bool, *, updated_by: str = "") -> None:
        code = normalize_branch_code(branch_code)
        if not code:
            raise UnitDirectoryError("Mã chi nhánh không hợp lệ.")
        with self._database() as database:
            database.execute(
                """
                UPDATE branch_directory
                SET is_active = ?, updated_at = ?, updated_by = ?
                WHERE branch_code = ?
                """,
                (1 if active else 0, _now(), _clean(updated_by), code),
            )

    def list_offices(
        self,
        *,
        branch_code: object = "",
        office_type: str = "",
        active_only: bool = False,
    ) -> list[OfficeDirectoryEntry]:
        clauses: list[str] = []
        params: list[object] = []
        branch = normalize_branch_code(branch_code)
        if branch:
            clauses.append("branch_code = ?")
            params.append(branch)
        office_type = str(office_type or "").strip().upper()
        if office_type:
            clauses.append("office_type = ?")
            params.append(office_type)
        if active_only:
            clauses.append("is_active = 1")
        where = "WHERE " + " AND ".join(clauses) if clauses else ""
        with self._database() as database:
            rows = database.execute(
                f"""
                SELECT *
                FROM office_directory
                {where}
                ORDER BY branch_code COLLATE NOCASE,
                    COALESCE(sort_order, 999999),
                    CASE office_type
                        WHEN 'HEAD_OFFICE' THEN 0
                        WHEN 'TRANSACTION_OFFICE' THEN 1
                        ELSE 2
                    END,
                    trctcd COLLATE NOCASE
                """,
                params,
            ).fetchall()
        return [_office_from_row(row) for row in rows]

    def get_office(self, branch_code: object, trctcd: object) -> OfficeDirectoryEntry | None:
        branch = normalize_branch_code(branch_code)
        code = normalize_trctcd(trctcd)
        if not branch or not code:
            return None
        with self._database() as database:
            row = database.execute(
                "SELECT * FROM office_directory WHERE branch_code = ? AND trctcd = ?",
                (branch, code),
            ).fetchone()
        return _office_from_row(row) if row is not None else None

    def get_office_by_code(self, office_code: object) -> OfficeDirectoryEntry | None:
        code = _clean(office_code)
        if not code:
            return None
        with self._database() as database:
            row = database.execute(
                "SELECT * FROM office_directory WHERE office_code = ?",
                (code,),
            ).fetchone()
        return _office_from_row(row) if row is not None else None

    def save_office(self, entry: OfficeDirectoryEntry) -> OfficeDirectoryEntry:
        branch = _validate_branch_code(entry.branch_code)
        trctcd = _validate_required(normalize_trctcd(entry.trctcd), "Mã TRCTCD")
        office_name = _validate_required(entry.office_name, "Tên đơn vị")
        office_type = str(entry.office_type or default_office_type(trctcd)).strip().upper()
        if office_type not in OFFICE_TYPES:
            raise UnitDirectoryError("Loại đơn vị không hợp lệ.")
        office_code = build_office_code(branch, trctcd)
        now = _now()
        with self._database() as database:
            if database.execute(
                "SELECT 1 FROM branch_directory WHERE branch_code = ?",
                (branch,),
            ).fetchone() is None:
                raise UnitDirectoryError("Chi nhánh của Hội sở/PGD chưa tồn tại.")
            existing = database.execute(
                "SELECT created_at FROM office_directory WHERE branch_code = ? AND trctcd = ?",
                (branch, trctcd),
            ).fetchone()
            created_at = str(existing["created_at"] or now) if existing else now
            database.execute(
                """
                INSERT INTO office_directory(
                    branch_code, trctcd, office_code, office_name, short_name,
                    office_type, is_active, sort_order, created_at, updated_at, updated_by
                )
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                ON CONFLICT(branch_code, trctcd) DO UPDATE SET
                    office_code = excluded.office_code,
                    office_name = excluded.office_name,
                    short_name = excluded.short_name,
                    office_type = excluded.office_type,
                    is_active = excluded.is_active,
                    sort_order = excluded.sort_order,
                    updated_at = excluded.updated_at,
                    updated_by = excluded.updated_by
                """,
                (
                    branch,
                    trctcd,
                    office_code,
                    office_name,
                    _clean(entry.short_name),
                    office_type,
                    1 if entry.is_active else 0,
                    entry.sort_order,
                    created_at,
                    now,
                    _clean(entry.updated_by),
                ),
            )
        saved = self.get_office(branch, trctcd)
        if saved is None:
            raise UnitDirectoryError("Không thể lưu Hội sở/PGD.")
        return saved

    def create_office(self, entry: OfficeDirectoryEntry) -> OfficeDirectoryEntry:
        branch = _validate_branch_code(entry.branch_code)
        trctcd = _validate_required(normalize_trctcd(entry.trctcd), "Mã TRCTCD")
        if self.get_office(branch, trctcd) is not None:
            raise UnitDirectoryError("Mã TRCTCD đã tồn tại trong chi nhánh.")
        if self.get_office_by_code(build_office_code(branch, trctcd)) is not None:
            raise UnitDirectoryError("Mã đơn vị đã tồn tại.")
        return self.save_office(entry)

    def set_office_active(self, office_code: object, active: bool, *, updated_by: str = "") -> None:
        code = _clean(office_code)
        if not code:
            raise UnitDirectoryError("Mã đơn vị không hợp lệ.")
        with self._database() as database:
            database.execute(
                """
                UPDATE office_directory
                SET is_active = ?, updated_at = ?, updated_by = ?
                WHERE office_code = ?
                """,
                (1 if active else 0, _now(), _clean(updated_by), code),
            )

    def load_settings(self) -> AppUnitSettings:
        with self._database() as database:
            row = database.execute(
                "SELECT * FROM app_unit_settings WHERE id = 1"
            ).fetchone()
        if row is None:
            return AppUnitSettings()
        return AppUnitSettings(
            home_branch_code=str(row["home_branch_code"] or ""),
            default_office_code=str(row["default_office_code"] or ""),
            organization_name=str(row["organization_name"] or ""),
            updated_at=str(row["updated_at"] or ""),
            updated_by=str(row["updated_by"] or ""),
        )

    def save_settings(self, settings: AppUnitSettings) -> AppUnitSettings:
        home_branch = normalize_branch_code(settings.home_branch_code)
        default_office = _clean(settings.default_office_code)
        with self._database() as database:
            if home_branch and database.execute(
                "SELECT 1 FROM branch_directory WHERE branch_code = ?",
                (home_branch,),
            ).fetchone() is None:
                raise UnitDirectoryError("Chi nhánh đang sử dụng chưa tồn tại.")
            if default_office:
                row = database.execute(
                    "SELECT branch_code FROM office_directory WHERE office_code = ?",
                    (default_office,),
                ).fetchone()
                if row is None:
                    raise UnitDirectoryError("Đơn vị mặc định chưa tồn tại.")
                if home_branch and str(row["branch_code"] or "") != home_branch:
                    raise UnitDirectoryError("Đơn vị mặc định phải thuộc chi nhánh đang sử dụng.")
            now = _now()
            database.execute(
                """
                INSERT INTO app_unit_settings(
                    id, home_branch_code, default_office_code, organization_name,
                    updated_at, updated_by
                )
                VALUES (1, ?, ?, ?, ?, ?)
                ON CONFLICT(id) DO UPDATE SET
                    home_branch_code = excluded.home_branch_code,
                    default_office_code = excluded.default_office_code,
                    organization_name = excluded.organization_name,
                    updated_at = excluded.updated_at,
                    updated_by = excluded.updated_by
                """,
                (
                    home_branch,
                    default_office,
                    _clean(settings.organization_name),
                    now,
                    _clean(settings.updated_by),
                ),
            )
        return self.load_settings()

    def ensure_branch_placeholder(self, branch_code: object, *, updated_by: str = "system") -> bool:
        code = normalize_branch_code(branch_code)
        if not code:
            return False
        with self._database() as database:
            if database.execute(
                "SELECT 1 FROM branch_directory WHERE branch_code = ?",
                (code,),
            ).fetchone() is not None:
                return False
            now = _now()
            database.execute(
                """
                INSERT INTO branch_directory(
                    branch_code, branch_name, short_name, display_name,
                    province_name, is_active, sort_order, created_at, updated_at, updated_by
                )
                VALUES (?, 'Chưa khai báo', 'CN chưa khai báo', '', '', 1, NULL, ?, ?, ?)
                """,
                (code, now, now, _clean(updated_by)),
            )
            return True

    def ensure_office_placeholder(
        self,
        branch_code: object,
        trctcd: object,
        *,
        updated_by: str = "system",
    ) -> bool:
        branch = normalize_branch_code(branch_code)
        code = normalize_trctcd(trctcd)
        if not branch or not code:
            return False
        self.ensure_branch_placeholder(branch, updated_by=updated_by)
        with self._database() as database:
            if database.execute(
                "SELECT 1 FROM office_directory WHERE branch_code = ? AND trctcd = ?",
                (branch, code),
            ).fetchone() is not None:
                return False
            now = _now()
            office_type = default_office_type(code)
            short_name = "Hội sở" if code == "00" else f"PGD {code}"
            database.execute(
                """
                INSERT INTO office_directory(
                    branch_code, trctcd, office_code, office_name, short_name,
                    office_type, is_active, sort_order, created_at, updated_at, updated_by
                )
                VALUES (?, ?, ?, ?, ?, ?, 1, NULL, ?, ?, ?)
                """,
                (
                    branch,
                    code,
                    build_office_code(branch, code),
                    short_name,
                    short_name,
                    office_type,
                    now,
                    now,
                    _clean(updated_by),
                ),
            )
            return True


def _branch_from_row(row: sqlite3.Row) -> BranchDirectoryEntry:
    return BranchDirectoryEntry(
        branch_code=str(row["branch_code"] or ""),
        branch_name=str(row["branch_name"] or ""),
        short_name=str(row["short_name"] or ""),
        display_name=str(row["display_name"] or ""),
        province_name=str(row["province_name"] or ""),
        is_active=bool(row["is_active"]),
        sort_order=int(row["sort_order"]) if row["sort_order"] is not None else None,
        created_at=str(row["created_at"] or ""),
        updated_at=str(row["updated_at"] or ""),
        updated_by=str(row["updated_by"] or ""),
    )


def _office_from_row(row: sqlite3.Row) -> OfficeDirectoryEntry:
    return OfficeDirectoryEntry(
        id=int(row["id"]) if row["id"] is not None else None,
        branch_code=str(row["branch_code"] or ""),
        trctcd=str(row["trctcd"] or ""),
        office_code=str(row["office_code"] or ""),
        office_name=str(row["office_name"] or ""),
        short_name=str(row["short_name"] or ""),
        office_type=str(row["office_type"] or OTHER_OFFICE),
        is_active=bool(row["is_active"]),
        sort_order=int(row["sort_order"]) if row["sort_order"] is not None else None,
        created_at=str(row["created_at"] or ""),
        updated_at=str(row["updated_at"] or ""),
        updated_by=str(row["updated_by"] or ""),
    )


def _validate_branch_code(value: object) -> str:
    code = normalize_branch_code(value)
    if not code:
        raise UnitDirectoryError("Mã chi nhánh không được để trống.")
    if any(char.isspace() for char in code):
        raise UnitDirectoryError("Mã chi nhánh không được chứa khoảng trắng.")
    return code


def _validate_required(value: object, label: str) -> str:
    clean = _clean(value)
    if not clean:
        raise UnitDirectoryError(f"{label} không được để trống.")
    return clean


def _clean(value: object) -> str:
    return " ".join(str(value or "").strip().split())


def _now() -> str:
    return datetime.now().astimezone().isoformat(timespec="seconds")
