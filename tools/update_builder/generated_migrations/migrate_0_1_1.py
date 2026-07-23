from __future__ import annotations


def table_exists(conn, table_name: str) -> bool:
    return conn.execute(
        "SELECT 1 FROM sqlite_master WHERE type = 'table' AND name = ?",
        (table_name,),
    ).fetchone() is not None


def column_exists(conn, table_name: str, column_name: str) -> bool:
    if not table_exists(conn, table_name):
        return False
    return column_name in {
        str(row[1]) for row in conn.execute(f"PRAGMA table_info({table_name})")
    }


def index_exists(conn, index_name: str) -> bool:
    return conn.execute(
        "SELECT 1 FROM sqlite_master WHERE type = 'index' AND name = ?",
        (index_name,),
    ).fetchone() is not None


def migrate_0_1_1(conn):
    if not table_exists(conn, 'credit_group_commission_rules'):
        conn.execute("""
            CREATE TABLE credit_group_commission_rules (
                            MaTo TEXT PRIMARY KEY,
                            use_custom_rule INTEGER DEFAULT 0,
                            updated_at TEXT
                        )
        """)
    if table_exists(conn, 'credit_groups') and not column_exists(conn, 'credit_groups', 'is_active'):
        conn.execute("""
            ALTER TABLE "credit_groups" ADD COLUMN "is_active" INTEGER DEFAULT 1
        """)
    if table_exists(conn, 'app_preferences'):
        conn.execute(
            """
            INSERT OR IGNORE INTO app_preferences(key, value, updated_at)
            VALUES (?, ?, ?)
            """,
            ('new_default_setting', 'default_value', '2026-07-20'),
        )
