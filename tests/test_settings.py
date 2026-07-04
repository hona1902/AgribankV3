from __future__ import annotations

from pathlib import Path
from contextlib import closing
import sqlite3
from tempfile import TemporaryDirectory
import unittest
import zipfile

from agribank_v3.settings import AddinMode, AppSettingsDatabase, BranchProfile


class SettingsDatabaseTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temporary_directory = TemporaryDirectory()
        self.database_path = (
            Path(self.temporary_directory.name) / "DuLieuV3.db"
        )
        self.database = AppSettingsDatabase(self.database_path)
        with closing(
            sqlite3.connect(self.database.quiz_database_path)
        ) as connection:
            connection.execute(
                """
                CREATE TABLE questions (
                    id INTEGER PRIMARY KEY,
                    question_text TEXT NOT NULL
                )
                """
            )
            connection.commit()

    def tearDown(self) -> None:
        self.temporary_directory.cleanup()

    def test_branch_profile_is_structured_versioned_and_durable(self) -> None:
        saved = self.database.save_branch_profile(
            BranchProfile(
                branch_code="  1234 ",
                branch_name="Agribank Chi nhánh Trung tâm",
                address="  01   Nguyễn Huệ  ",
                report_preparer="Nguyễn Văn An",
            )
        )
        updated = self.database.save_branch_profile(
            BranchProfile(
                branch_code="1234",
                branch_name="Agribank Chi nhánh Trung tâm",
                address="01 Nguyễn Huệ",
                report_preparer="Trần Thị Bình",
            )
        )

        loaded = self.database.load_branch_profile()
        self.assertEqual(saved.revision, 1)
        self.assertEqual(updated.revision, 2)
        self.assertEqual(loaded.address, "01 Nguyễn Huệ")
        self.assertEqual(loaded.report_preparer, "Trần Thị Bình")

        with closing(sqlite3.connect(self.database_path)) as connection:
            history_count = connection.execute(
                "SELECT COUNT(*) FROM branch_profile_history"
            ).fetchone()[0]
        self.assertEqual(history_count, 2)

    def test_backup_and_restore_include_branch_profile(self) -> None:
        self.database.save_branch_profile(
            BranchProfile(branch_code="1001", branch_name="Chi nhánh A")
        )
        backup = self.database.create_backup()
        self.database.save_branch_profile(
            BranchProfile(branch_code="2002", branch_name="Chi nhánh B")
        )

        safety_backup = self.database.restore_backup(backup)

        restored = self.database.load_branch_profile()
        self.assertEqual(restored.branch_code, "1001")
        self.assertEqual(restored.branch_name, "Chi nhánh A")
        self.assertTrue(safety_backup.is_file())
        self.assertEqual(self.database.status().integrity, "ok")

    def test_backup_bundle_includes_and_restores_both_databases(self) -> None:
        self.database.save_branch_profile(
            BranchProfile(branch_code="1001", branch_name="Chi nhánh A")
        )
        with closing(
            sqlite3.connect(self.database.quiz_database_path)
        ) as connection:
            connection.execute(
                """
                INSERT INTO questions(id, question_text)
                VALUES (1, 'Nội dung trước sao lưu')
                """
            )
            connection.commit()

        backup = self.database.create_backup()
        self.assertEqual(backup.suffix, ".zip")
        with zipfile.ZipFile(backup) as archive:
            self.assertEqual(
                set(archive.namelist()),
                {"manifest.json", "DuLieuV3.db", "quiz.db"},
            )

        self.database.save_branch_profile(
            BranchProfile(branch_code="2002", branch_name="Chi nhánh B")
        )
        with closing(
            sqlite3.connect(self.database.quiz_database_path)
        ) as connection:
            connection.execute(
                "UPDATE questions SET question_text = 'Nội dung đã thay đổi'"
            )
            connection.commit()

        self.database.restore_backup(backup)

        self.assertEqual(
            self.database.load_branch_profile().branch_code,
            "1001",
        )
        with closing(
            sqlite3.connect(self.database.quiz_database_path)
        ) as connection:
            restored_question = connection.execute(
                "SELECT question_text FROM questions WHERE id = 1"
            ).fetchone()[0]
        self.assertEqual(restored_question, "Nội dung trước sao lưu")

    def test_addin_mode_defaults_to_permanent_and_is_durable(self) -> None:
        self.assertEqual(
            self.database.load_addin_mode(),
            AddinMode.PERMANENT,
        )

        saved = self.database.save_addin_mode(AddinMode.SESSION)
        reopened = AppSettingsDatabase(self.database_path)

        self.assertEqual(saved, AddinMode.SESSION)
        self.assertEqual(reopened.load_addin_mode(), AddinMode.SESSION)

    def test_each_addin_enabled_state_is_durable_and_new_files_default_on(
        self,
    ) -> None:
        initial = self.database.load_addin_states(
            ["FunctionsA.xlam", "FunctionsB.xla"]
        )
        self.assertEqual(
            initial,
            {"FunctionsA.xlam": True, "FunctionsB.xla": True},
        )

        self.database.save_addin_enabled("FunctionsA.xlam", False)
        reopened = AppSettingsDatabase(self.database_path)

        self.assertEqual(
            reopened.load_addin_states(
                ["FunctionsA.xlam", "FunctionsB.xla", "NewFunctions.xlam"]
            ),
            {
                "FunctionsA.xlam": False,
                "FunctionsB.xla": True,
                "NewFunctions.xlam": True,
            },
        )


if __name__ == "__main__":
    unittest.main()
