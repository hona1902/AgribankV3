from __future__ import annotations

from contextlib import closing
from pathlib import Path
import sqlite3
from tempfile import TemporaryDirectory
import unittest

from agribank_v3.settings import AppSettingsDatabase
from agribank_v3.user_databases import ensure_user_databases


class UserDatabaseEnsureTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temporary_directory = TemporaryDirectory()
        self.root = Path(self.temporary_directory.name)
        self.database_path = self.root / "data" / "DuLieuV3.db"
        self.settings_database = AppSettingsDatabase(self.database_path)

    def tearDown(self) -> None:
        self.temporary_directory.cleanup()

    def test_ensure_user_databases_creates_new_databases_after_update(self) -> None:
        for name in ("CreditSummary.db", "Customer.db", "Credit.db"):
            self.assertFalse((self.database_path.parent / name).exists())

        results = ensure_user_databases(self.settings_database, strict=True)

        self.assertEqual([result.name for result in results], ["CreditSummary.db", "Customer.db", "Credit.db"])
        self.assertTrue(all(result.ok for result in results))
        self._assert_table_exists(self.database_path.parent / "CreditSummary.db", "summary_import_history")
        self._assert_table_exists(self.database_path.parent / "CreditSummary.db", "credit_summary_schema_migrations")
        self._assert_table_exists(self.database_path.parent / "Customer.db", "customer_master")
        self._assert_table_exists(self.database_path.parent / "Customer.db", "customer_schema_migrations")
        self._assert_table_exists(self.database_path.parent / "Credit.db", "credit_import_runs")
        self._assert_table_exists(self.database_path.parent / "Credit.db", "credit_schema_migrations")

    def test_ensure_user_databases_is_idempotent(self) -> None:
        first = ensure_user_databases(self.settings_database, strict=True)
        second = ensure_user_databases(self.settings_database, strict=True)

        self.assertTrue(all(result.ok for result in first))
        self.assertTrue(all(result.ok for result in second))
        self.assertTrue(all(result.existed_before for result in second))

    def _assert_table_exists(self, database_path: Path, table_name: str) -> None:
        self.assertTrue(database_path.is_file())
        with closing(sqlite3.connect(database_path)) as connection:
            row = connection.execute(
                "SELECT 1 FROM sqlite_master WHERE type = 'table' AND name = ?",
                (table_name,),
            ).fetchone()
        self.assertIsNotNone(row, f"{table_name} missing in {database_path.name}")


if __name__ == "__main__":
    unittest.main()
