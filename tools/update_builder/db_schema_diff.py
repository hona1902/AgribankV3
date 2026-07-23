from __future__ import annotations

from contextlib import closing
from dataclasses import dataclass
from importlib import util
from pathlib import Path
import re
import shutil
import sqlite3
import sys
import tempfile


@dataclass(frozen=True, slots=True)
class SQLiteColumn:
    table: str
    name: str
    data_type: str
    not_null: bool
    default_value: str | None
    primary_key: int

    def definition_sql(self) -> str:
        pieces = [_quote_identifier(self.name)]
        if self.data_type:
            pieces.append(self.data_type)
        if self.not_null and self.default_value is not None:
            pieces.append("NOT NULL")
        if self.default_value is not None:
            pieces.append(f"DEFAULT {self.default_value}")
        return " ".join(pieces)


@dataclass(frozen=True, slots=True)
class SQLiteTable:
    name: str
    create_sql: str
    columns: dict[str, SQLiteColumn]


@dataclass(frozen=True, slots=True)
class SQLiteIndex:
    name: str
    table: str
    unique: bool
    columns: tuple[str, ...]
    create_sql: str


@dataclass(frozen=True, slots=True)
class SQLiteSchema:
    tables: dict[str, SQLiteTable]
    indexes: dict[str, SQLiteIndex]
    app_preferences: dict[str, tuple[str, str]]


@dataclass(frozen=True, slots=True)
class NewPreference:
    key: str
    value: str
    updated_at: str


@dataclass(frozen=True, slots=True)
class DangerousChange:
    kind: str
    name: str
    detail: str


@dataclass(frozen=True, slots=True)
class SchemaDiff:
    old_schema: SQLiteSchema
    new_schema: SQLiteSchema
    new_tables: tuple[SQLiteTable, ...]
    new_columns: tuple[SQLiteColumn, ...]
    new_indexes: tuple[SQLiteIndex, ...]
    new_app_preferences: tuple[NewPreference, ...]
    dangerous_changes: tuple[DangerousChange, ...]

    @property
    def has_safe_changes(self) -> bool:
        return bool(
            self.new_tables
            or self.new_columns
            or self.new_indexes
            or self.new_app_preferences
        )


@dataclass(frozen=True, slots=True)
class MigrationValidationResult:
    success: bool
    test_database_path: Path
    missing_items: tuple[str, ...]
    error: str = ""


def read_sqlite_schema(db_path: Path | str) -> SQLiteSchema:
    path = Path(db_path)
    if not path.is_file():
        raise FileNotFoundError(f"Không tìm thấy database: {path}")
    with closing(sqlite3.connect(path)) as connection:
        connection.row_factory = sqlite3.Row
        tables: dict[str, SQLiteTable] = {}
        for row in connection.execute(
            """
            SELECT name, sql
            FROM sqlite_master
            WHERE type = 'table'
              AND name NOT LIKE 'sqlite_%'
            ORDER BY name
            """
        ):
            table_name = str(row["name"])
            columns = _read_columns(connection, table_name)
            tables[table_name] = SQLiteTable(
                name=table_name,
                create_sql=str(row["sql"] or ""),
                columns=columns,
            )
        indexes: dict[str, SQLiteIndex] = {}
        for table_name in tables:
            for row in connection.execute(f"PRAGMA index_list({_quote_identifier(table_name)})"):
                index_name = str(row["name"])
                index_sql_row = connection.execute(
                    "SELECT sql FROM sqlite_master WHERE type = 'index' AND name = ?",
                    (index_name,),
                ).fetchone()
                create_sql = str(index_sql_row["sql"] or "") if index_sql_row else ""
                if not create_sql:
                    continue
                index_columns = tuple(
                    str(info["name"])
                    for info in connection.execute(
                        f"PRAGMA index_info({_quote_identifier(index_name)})"
                    )
                )
                indexes[index_name] = SQLiteIndex(
                    name=index_name,
                    table=table_name,
                    unique=bool(row["unique"]),
                    columns=index_columns,
                    create_sql=create_sql,
                )
        return SQLiteSchema(
            tables=tables,
            indexes=indexes,
            app_preferences=_read_app_preferences(connection, tables),
        )


def compare_sqlite_schema(
    old_db_path: Path | str,
    new_db_path: Path | str,
) -> SchemaDiff:
    old_schema = read_sqlite_schema(old_db_path)
    new_schema = read_sqlite_schema(new_db_path)
    dangerous: list[DangerousChange] = []
    new_tables = detect_new_tables(old_schema, new_schema)
    new_columns = list(detect_new_columns(old_schema, new_schema))
    safe_columns: list[SQLiteColumn] = []
    for column in new_columns:
        if column.not_null and column.default_value is None:
            dangerous.append(
                DangerousChange(
                    "new_not_null_column",
                    f"{column.table}.{column.name}",
                    "Cột mới NOT NULL không có DEFAULT; cần xử lý thủ công.",
                )
            )
        elif column.primary_key:
            dangerous.append(
                DangerousChange(
                    "new_primary_key_column",
                    f"{column.table}.{column.name}",
                    "Không tự thêm cột primary key vào bảng đã tồn tại.",
                )
            )
        else:
            safe_columns.append(column)
    dangerous.extend(detect_removed_tables(old_schema, new_schema))
    dangerous.extend(detect_removed_columns(old_schema, new_schema))
    dangerous.extend(detect_changed_columns(old_schema, new_schema))
    dangerous.extend(detect_changed_preferences(old_schema, new_schema))
    return SchemaDiff(
        old_schema=old_schema,
        new_schema=new_schema,
        new_tables=new_tables,
        new_columns=tuple(safe_columns),
        new_indexes=detect_new_indexes(old_schema, new_schema),
        new_app_preferences=detect_new_app_preferences(old_schema, new_schema),
        dangerous_changes=tuple(dangerous),
    )


def detect_new_tables(
    old_schema: SQLiteSchema,
    new_schema: SQLiteSchema,
) -> tuple[SQLiteTable, ...]:
    return tuple(
        table
        for name, table in new_schema.tables.items()
        if name not in old_schema.tables
    )


def detect_new_columns(
    old_schema: SQLiteSchema,
    new_schema: SQLiteSchema,
) -> tuple[SQLiteColumn, ...]:
    columns: list[SQLiteColumn] = []
    for table_name, new_table in new_schema.tables.items():
        old_table = old_schema.tables.get(table_name)
        if old_table is None:
            continue
        for column_name, column in new_table.columns.items():
            if column_name not in old_table.columns:
                columns.append(column)
    return tuple(columns)


def detect_new_indexes(
    old_schema: SQLiteSchema,
    new_schema: SQLiteSchema,
) -> tuple[SQLiteIndex, ...]:
    return tuple(
        index
        for name, index in new_schema.indexes.items()
        if name not in old_schema.indexes and index.table in old_schema.tables
    )


def detect_removed_tables(
    old_schema: SQLiteSchema,
    new_schema: SQLiteSchema,
) -> tuple[DangerousChange, ...]:
    return tuple(
        DangerousChange(
            "removed_table",
            table_name,
            "Bảng có trong DB cũ nhưng không có trong DB mới.",
        )
        for table_name in old_schema.tables
        if table_name not in new_schema.tables
    )


def detect_removed_columns(
    old_schema: SQLiteSchema,
    new_schema: SQLiteSchema,
) -> tuple[DangerousChange, ...]:
    changes: list[DangerousChange] = []
    for table_name, old_table in old_schema.tables.items():
        new_table = new_schema.tables.get(table_name)
        if new_table is None:
            continue
        for column_name in old_table.columns:
            if column_name not in new_table.columns:
                changes.append(
                    DangerousChange(
                        "removed_column",
                        f"{table_name}.{column_name}",
                        "Cột có trong DB cũ nhưng không có trong DB mới.",
                    )
                )
    return tuple(changes)


def detect_changed_columns(
    old_schema: SQLiteSchema,
    new_schema: SQLiteSchema,
) -> tuple[DangerousChange, ...]:
    changes: list[DangerousChange] = []
    for table_name, old_table in old_schema.tables.items():
        new_table = new_schema.tables.get(table_name)
        if new_table is None:
            continue
        for column_name, old_column in old_table.columns.items():
            new_column = new_table.columns.get(column_name)
            if new_column is None:
                continue
            old_signature = (
                old_column.data_type.casefold(),
                old_column.not_null,
                old_column.default_value,
                old_column.primary_key,
            )
            new_signature = (
                new_column.data_type.casefold(),
                new_column.not_null,
                new_column.default_value,
                new_column.primary_key,
            )
            if old_signature != new_signature:
                changes.append(
                    DangerousChange(
                        "changed_column",
                        f"{table_name}.{column_name}",
                        "Kiểu dữ liệu/ràng buộc/default thay đổi; không tự migration.",
                    )
                )
    return tuple(changes)


def detect_new_app_preferences(
    old_schema: SQLiteSchema,
    new_schema: SQLiteSchema,
) -> tuple[NewPreference, ...]:
    preferences: list[NewPreference] = []
    for key, (value, updated_at) in new_schema.app_preferences.items():
        if key not in old_schema.app_preferences:
            preferences.append(NewPreference(key=key, value=value, updated_at=updated_at))
    return tuple(preferences)


def detect_changed_preferences(
    old_schema: SQLiteSchema,
    new_schema: SQLiteSchema,
) -> tuple[DangerousChange, ...]:
    changes: list[DangerousChange] = []
    for key, (old_value, _old_updated_at) in old_schema.app_preferences.items():
        if key in new_schema.app_preferences and new_schema.app_preferences[key][0] != old_value:
            changes.append(
                DangerousChange(
                    "changed_app_preference",
                    key,
                    "Key app_preferences đã tồn tại nhưng value khác; không ghi đè cấu hình người dùng.",
                )
            )
    return tuple(changes)


def generate_python_migration(diff: SchemaDiff, version: str) -> str:
    function_name = migration_function_name(version)
    lines = [
        "from __future__ import annotations",
        "",
        "",
        "def table_exists(conn, table_name: str) -> bool:",
        "    return conn.execute(",
        "        \"SELECT 1 FROM sqlite_master WHERE type = 'table' AND name = ?\",",
        "        (table_name,),",
        "    ).fetchone() is not None",
        "",
        "",
        "def column_exists(conn, table_name: str, column_name: str) -> bool:",
        "    if not table_exists(conn, table_name):",
        "        return False",
        "    return column_name in {",
        "        str(row[1]) for row in conn.execute(f\"PRAGMA table_info({table_name})\")",
        "    }",
        "",
        "",
        "def index_exists(conn, index_name: str) -> bool:",
        "    return conn.execute(",
        "        \"SELECT 1 FROM sqlite_master WHERE type = 'index' AND name = ?\",",
        "        (index_name,),",
        "    ).fetchone() is not None",
        "",
        "",
        f"def {function_name}(conn):",
    ]
    body: list[str] = []
    for table in diff.new_tables:
        body.extend(
            [
                f"    if not table_exists(conn, {table.name!r}):",
                "        conn.execute(" + _triple_quote(table.create_sql, 12) + ")",
            ]
        )
    for column in diff.new_columns:
        body.extend(
            [
                (
                    f"    if table_exists(conn, {column.table!r}) and not "
                    f"column_exists(conn, {column.table!r}, {column.name!r}):"
                ),
                (
                    "        conn.execute("
                    + _triple_quote(
                        f"ALTER TABLE {_quote_identifier(column.table)} ADD COLUMN {column.definition_sql()}",
                        12,
                    )
                    + ")"
                ),
            ]
        )
    for index in diff.new_indexes:
        sql = _with_if_not_exists(index.create_sql, "INDEX", index.name)
        body.extend(
            [
                f"    if not index_exists(conn, {index.name!r}):",
                "        conn.execute(" + _triple_quote(sql, 12) + ")",
            ]
        )
    for preference in diff.new_app_preferences:
        body.extend(
            [
                "    if table_exists(conn, 'app_preferences'):",
                "        conn.execute(",
                "            \"\"\"",
                "            INSERT OR IGNORE INTO app_preferences(key, value, updated_at)",
                "            VALUES (?, ?, ?)",
                "            \"\"\",",
                f"            ({preference.key!r}, {preference.value!r}, {preference.updated_at!r}),",
                "        )",
            ]
        )
    if not body:
        body.append("    pass")
    lines.extend(body)
    lines.append("")
    return "\n".join(lines)


def generate_sql_reference(diff: SchemaDiff, version: str) -> str:
    statements: list[str] = [
        f"-- SQL tham khảo cho migration {version}.",
        "-- Với thêm cột, ưu tiên Python migration để kiểm tra column_exists.",
        "",
    ]
    for table in diff.new_tables:
        statements.append(_with_if_not_exists(table.create_sql, "TABLE", table.name) + ";")
    for index in diff.new_indexes:
        statements.append(_with_if_not_exists(index.create_sql, "INDEX", index.name) + ";")
    for preference in diff.new_app_preferences:
        statements.append(
            "INSERT OR IGNORE INTO app_preferences(key, value, updated_at) "
            f"VALUES ({_sql_literal(preference.key)}, {_sql_literal(preference.value)}, {_sql_literal(preference.updated_at)});"
        )
    return "\n\n".join(statements) + "\n"


def write_generated_python_migration(
    diff: SchemaDiff,
    version: str,
    output_dir: Path | str | None = None,
) -> Path:
    directory = (
        Path(output_dir)
        if output_dir is not None
        else _builder_output_root() / "generated_migrations"
    )
    directory.mkdir(parents=True, exist_ok=True)
    path = directory / f"migrate_{version.replace('.', '_')}.py"
    path.write_text(generate_python_migration(diff, version), encoding="utf-8")
    return path


def validate_generated_migration_on_copy(
    *,
    old_db_path: Path | str,
    new_db_path: Path | str,
    migration_file: Path | str,
    version: str,
) -> MigrationValidationResult:
    old_path = Path(old_db_path)
    with tempfile.TemporaryDirectory(prefix="agribankv3-migration-test-") as temp_dir:
        test_path = Path(temp_dir) / f"test_migration_{version}.db"
        shutil.copy2(old_path, test_path)
        try:
            _run_python_migration_file(test_path, Path(migration_file), version)
            remaining_diff = compare_sqlite_schema(test_path, new_db_path)
            missing = tuple(_missing_safe_items(remaining_diff))
            final_path = old_path.parent / f"test_migration_{version.replace('.', '_')}.db"
            shutil.copy2(test_path, final_path)
            return MigrationValidationResult(
                success=not missing,
                test_database_path=final_path,
                missing_items=missing,
            )
        except Exception as exc:
            final_path = old_path.parent / f"test_migration_{version.replace('.', '_')}_failed.db"
            if test_path.exists():
                shutil.copy2(test_path, final_path)
            return MigrationValidationResult(
                success=False,
                test_database_path=final_path,
                missing_items=(),
                error=str(exc),
            )


def migration_function_name(version: str) -> str:
    return f"migrate_{version.replace('.', '_').replace('-', '_')}"


def _read_columns(
    connection: sqlite3.Connection,
    table_name: str,
) -> dict[str, SQLiteColumn]:
    columns = {}
    for row in connection.execute(f"PRAGMA table_info({_quote_identifier(table_name)})"):
        column = SQLiteColumn(
            table=table_name,
            name=str(row["name"]),
            data_type=str(row["type"] or ""),
            not_null=bool(row["notnull"]),
            default_value=(
                str(row["dflt_value"])
                if row["dflt_value"] is not None
                else None
            ),
            primary_key=int(row["pk"] or 0),
        )
        columns[column.name] = column
    return columns


def _read_app_preferences(
    connection: sqlite3.Connection,
    tables: dict[str, SQLiteTable],
) -> dict[str, tuple[str, str]]:
    table = tables.get("app_preferences")
    if table is None or "key" not in table.columns or "value" not in table.columns:
        return {}
    return {
        str(row["key"]): (
            str(row["value"]),
            str(row["updated_at"]) if "updated_at" in row.keys() else _now_sql_literal(),
        )
        for row in connection.execute("SELECT * FROM app_preferences")
    }


def _run_python_migration_file(
    database_path: Path,
    migration_file: Path,
    version: str,
) -> None:
    module_name = f"generated_migration_{version.replace('.', '_')}"
    spec = util.spec_from_file_location(
        module_name,
        migration_file,
    )
    if spec is None or spec.loader is None:
        raise RuntimeError(f"Không import được migration: {migration_file}")
    module = util.module_from_spec(spec)
    sys.modules[module_name] = module
    try:
        spec.loader.exec_module(module)
        function = getattr(module, migration_function_name(version))
        with closing(sqlite3.connect(database_path)) as connection:
            with connection:
                function(connection)
    finally:
        sys.modules.pop(module_name, None)


def _missing_safe_items(diff: SchemaDiff) -> list[str]:
    missing: list[str] = []
    missing.extend(f"table:{table.name}" for table in diff.new_tables)
    missing.extend(f"column:{column.table}.{column.name}" for column in diff.new_columns)
    missing.extend(f"index:{index.name}" for index in diff.new_indexes)
    missing.extend(f"app_preferences:{item.key}" for item in diff.new_app_preferences)
    return missing


def _quote_identifier(value: str) -> str:
    return '"' + value.replace('"', '""') + '"'


def _sql_literal(value: str) -> str:
    return "'" + value.replace("'", "''") + "'"


def _triple_quote(sql: str, indent: int) -> str:
    prefix = "\n" + " " * indent
    body = prefix.join(line.rstrip() for line in sql.strip().splitlines())
    return f'"""{prefix}{body}\n{" " * (indent - 4)}"""'


def _with_if_not_exists(sql: str, object_type: str, name: str) -> str:
    if object_type == "INDEX":
        return re.sub(
            r"CREATE\s+(UNIQUE\s+)?INDEX\s+(?!IF\s+NOT\s+EXISTS)",
            lambda match: f"CREATE {match.group(1) or ''}INDEX IF NOT EXISTS ",
            sql,
            count=1,
            flags=re.IGNORECASE,
        )
    pattern = rf"CREATE\s+({object_type})\s+(?!IF\s+NOT\s+EXISTS)"
    return re.sub(pattern, rf"CREATE \1 IF NOT EXISTS ", sql, count=1, flags=re.IGNORECASE)


def _now_sql_literal() -> str:
    return "1970-01-01T00:00:00"


def _builder_output_root() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent
