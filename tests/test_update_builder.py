from __future__ import annotations

import json
from pathlib import Path
import sqlite3
import sys
import time
from contextlib import closing
from tempfile import TemporaryDirectory
import unittest
import zipfile

from PySide6.QtWidgets import QApplication, QMessageBox


BUILDER_DIR = Path(__file__).resolve().parents[1] / "tools" / "update_builder"
if str(BUILDER_DIR) not in sys.path:
    sys.path.insert(0, str(BUILDER_DIR))

from update_builder_core import (  # noqa: E402
    AutoBuildConfig,
    APP_BUILD_INFO_FILE_NAME,
    BuildUpdateConfig,
    MigrationItem,
    UpdateBuilder,
    UpdateBuilderError,
    auto_build_update,
    auto_detect_database_changes,
    backup_existing_update_files,
    build_manifest,
    collect_files,
    collect_package_files,
    create_package_plan,
    compare_versions,
    copy_migrations,
    create_update_zip,
    format_package_report,
    package_file_report,
    read_app_build_info,
    detect_previous_release_version,
    read_current_version,
    save_release_file_snapshot,
    save_schema_snapshot,
    source_has_python_migration,
    update_source_version,
    warn_dangerous_sql,
    write_manifest,
)
from update_builder_app import main as update_builder_main  # noqa: E402
from update_builder_ui import UpdateBuilderWindow, UpdateTaskWorker  # noqa: E402


class UpdateBuilderTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temporary_directory = TemporaryDirectory()
        self.root = Path(self.temporary_directory.name)
        self.source = self.root / "source"
        self.update = self.root / "Update"
        self._create_source(self.source, version="0.1.1")

    def tearDown(self) -> None:
        self.temporary_directory.cleanup()

    def test_read_current_version(self) -> None:
        self.assertEqual(read_current_version(self.source), "0.1.1")

    def test_compare_new_version_must_be_greater(self) -> None:
        self.assertLess(compare_versions("0.1.1", "0.1.2"), 0)
        self.assertLess(compare_versions("0.1.2", "0.1.10"), 0)

    def test_create_zip_excludes_database(self) -> None:
        output_zip = self.root / "AgribankV3_0.1.2.zip"

        create_update_zip(self.source, output_zip)

        with zipfile.ZipFile(output_zip) as archive:
            names = set(archive.namelist())
        self.assertIn("pyproject.toml", names)
        self.assertNotIn("data/DuLieuV3.db", names)
        self.assertNotIn("data/quiz.db", names)

    def test_create_zip_excludes_logs_backups_temp(self) -> None:
        output_zip = self.root / "AgribankV3_0.1.2.zip"

        create_update_zip(self.source, output_zip)

        with zipfile.ZipFile(output_zip) as archive:
            names = set(archive.namelist())
        self.assertFalse(any(name.startswith("logs/") for name in names))
        self.assertFalse(any(name.startswith("backups/") for name in names))
        self.assertFalse(any(name.startswith("temp/") for name in names))
        self.assertFalse(any(name.startswith("KetQua/") for name in names))

    def test_collect_files_excludes_large_dirs(self) -> None:
        large_dirs = [".venv", "venv", "dist", "build", "logs", "KetQua", "temp"]
        for folder in large_dirs:
            root = self.source / folder
            (root / "nested").mkdir(parents=True, exist_ok=True)
            (root / "nested" / "ignored.py").write_text("ignored", encoding="utf-8")

        included, excluded = collect_files(self.source)

        included_names = {item.relative_path.as_posix() for item in included}
        excluded_names = {item.relative_path.as_posix() for item in excluded}
        self.assertFalse(any("ignored.py" in name for name in included_names))
        for folder in large_dirs:
            self.assertIn(folder, excluded_names)

    def test_output_zip_not_included_when_update_path_inside_source(self) -> None:
        update_path = self.source / "Update"
        output_zip = update_path / "AgribankV3_0.1.2.zip"

        create_update_zip(
            self.source,
            output_zip,
            extra_excluded_paths=(update_path,),
        )

        with zipfile.ZipFile(output_zip) as archive:
            names = set(archive.namelist())
        self.assertFalse(any(name.startswith("Update/") for name in names))
        self.assertNotIn("Update/AgribankV3_0.1.2.zip", names)

    def test_runtime_package_excludes_tests_tools_dist(self) -> None:
        for folder in ("tests", "tools/update_builder", "dist/AgribankV3UpdateBuilder"):
            path = self.source / folder
            path.mkdir(parents=True, exist_ok=True)
            (path / "ignored.py").write_text("ignored", encoding="utf-8")

        included, _ = collect_package_files(self.source, "runtime")

        names = {item.relative_path.as_posix() for item in included}
        self.assertFalse(any(name.startswith("tests/") for name in names))
        self.assertFalse(any(name.startswith("tools/") for name in names))
        self.assertFalse(any(name.startswith("dist/") for name in names))

    def test_runtime_package_includes_src_agribank_v3(self) -> None:
        update_package = self.source / "src" / "agribank_v3" / "update"
        update_package.mkdir()
        (update_package / "db_migrations.py").write_text("MIGRATIONS = {}\n", encoding="utf-8")
        (self.source / "Update").mkdir()
        (self.source / "Update" / "old_package.zip").write_text("zip", encoding="utf-8")

        included, _ = collect_package_files(self.source, "runtime")

        names = {item.relative_path.as_posix() for item in included}
        self.assertIn("src/agribank_v3/__init__.py", names)
        self.assertIn("src/agribank_v3/update/db_migrations.py", names)
        self.assertNotIn("Update/old_package.zip", names)

    def test_update_builder_not_packaged_into_runtime_update(self) -> None:
        (self.source / "tools" / "update_builder").mkdir(parents=True)
        (self.source / "tools" / "update_builder" / "update_builder_app.py").write_text(
            "print('builder')",
            encoding="utf-8",
        )
        (self.source / "dist" / "AgribankV3UpdateBuilder").mkdir(parents=True)
        (
            self.source
            / "dist"
            / "AgribankV3UpdateBuilder"
            / "AgribankV3UpdateBuilder.exe"
        ).write_text("exe", encoding="utf-8")

        output_zip = self.root / "runtime.zip"
        included, _ = collect_package_files(self.source, "runtime")
        create_update_zip(self.source, output_zip, precollected_files=included)

        with zipfile.ZipFile(output_zip) as archive:
            names = set(archive.namelist())
        self.assertFalse(any(name.startswith("tools/update_builder/") for name in names))
        self.assertFalse(any(name.startswith("dist/AgribankV3UpdateBuilder/") for name in names))

    def test_zip_excludes_existing_zip_and_exe(self) -> None:
        package_dir = self.source / "src" / "agribank_v3"
        (package_dir / "old.zip").write_text("zip", encoding="utf-8")
        (package_dir / "tool.exe").write_text("exe", encoding="utf-8")

        included, _ = collect_package_files(self.source, "runtime")

        names = {item.relative_path.as_posix() for item in included}
        self.assertNotIn("src/agribank_v3/old.zip", names)
        self.assertNotIn("src/agribank_v3/tool.exe", names)

    def test_package_size_report_lists_largest_files(self) -> None:
        package_dir = self.source / "src" / "agribank_v3"
        (package_dir / "small.py").write_text("x", encoding="utf-8")
        (package_dir / "large.dat").write_bytes(b"x" * 4096)
        included, excluded = collect_package_files(self.source, "runtime")

        report = package_file_report(included, excluded)
        text = format_package_report(report)

        self.assertEqual(report.top_files[0].relative_path, Path("src/agribank_v3/large.dat"))
        self.assertIn("Top file lớn", text)

    def test_invalid_source_path_dist_rejected(self) -> None:
        invalid = self.root / "dist" / "AgribankV3UpdateBuilder"
        invalid.mkdir(parents=True)

        with self.assertRaisesRegex(Exception, "Thư mục source không hợp lệ"):
            UpdateBuilder().validate(
                BuildUpdateConfig(
                    source_path=invalid,
                    update_path=self.update,
                    new_version="0.1.2",
                    notes=("Test",),
                )
            )

    def test_delta_package_only_changed_files(self) -> None:
        baseline_files, _ = collect_package_files(self.source, "runtime")
        save_release_file_snapshot(self.source, "0.1.1", baseline_files)
        (self.source / "src" / "agribank_v3" / "__main__.py").write_text(
            "print('changed')\n",
            encoding="utf-8",
        )
        config = self._config(
            new_version="0.1.2",
            package_mode="delta",
            allow_rebuild_same_version=True,
        )

        result = UpdateBuilder().build(config)

        with zipfile.ZipFile(result.package_path) as archive:
            names = set(archive.namelist())
        self.assertEqual(names, {"src/agribank_v3/__main__.py"})

    def test_delta_manifest_has_package_type_and_base_version(self) -> None:
        baseline_files, _ = collect_package_files(self.source, "runtime")
        save_release_file_snapshot(self.source, "0.1.1", baseline_files)
        (self.source / "src" / "agribank_v3" / "__main__.py").write_text(
            "print('changed')\n",
            encoding="utf-8",
        )
        config = self._config(new_version="0.1.2", package_mode="delta")

        result = UpdateBuilder().build(config)

        manifest = json.loads(result.manifest_path.read_text(encoding="utf-8"))
        self.assertEqual(manifest["package_type"], "delta")
        self.assertEqual(manifest["base_version"], "0.1.1")

    def test_full_runtime_package_smaller_than_full_source(self) -> None:
        (self.source / "misc_big.bin").write_bytes(b"x" * 1024 * 1024)
        runtime_files, runtime_excluded = collect_package_files(self.source, "runtime")
        source_files, source_excluded = collect_package_files(self.source, "source")

        runtime_report = package_file_report(runtime_files, runtime_excluded)
        source_report = package_file_report(source_files, source_excluded)

        self.assertLess(runtime_report.total_size, source_report.total_size)

    def test_app_package_collects_built_exe_without_database(self) -> None:
        app_root = self.source / "dist" / "AgribankV3"
        (app_root / "_internal").mkdir(parents=True)
        (app_root / "data").mkdir()
        (app_root / "AgribankV3.exe").write_text("exe", encoding="utf-8")
        (app_root / APP_BUILD_INFO_FILE_NAME).write_text(
            '{"app":"AgribankV3","version":"0.1.2"}',
            encoding="utf-8",
        )
        (app_root / "_internal" / "library.dat").write_text("lib", encoding="utf-8")
        (app_root / "data" / "DuLieuV3.db").write_text("db", encoding="utf-8")
        (app_root / "data" / "AgribankMenuData.mdb").write_text("mdb", encoding="utf-8")

        files, excluded = collect_package_files(self.source, "app")

        names = {item.relative_path.as_posix() for item in files}
        excluded_names = {item.relative_path.as_posix() for item in excluded}
        self.assertIn("AgribankV3.exe", names)
        self.assertIn(APP_BUILD_INFO_FILE_NAME, names)
        self.assertIn("_internal/library.dat", names)
        self.assertNotIn("data/DuLieuV3.db", names)
        self.assertNotIn("data/AgribankMenuData.mdb", names)
        self.assertIn("data/DuLieuV3.db", excluded_names)
        self.assertIn("data/AgribankMenuData.mdb", excluded_names)

    def test_manifest_json_created(self) -> None:
        config = self._config(new_version="0.1.2", notes=("Sửa lỗi",))
        manifest = build_manifest(config, "AgribankV3_0.1.2.zip")

        manifest_path = write_manifest(self.update, manifest)

        loaded = json.loads(manifest_path.read_text(encoding="utf-8"))
        self.assertEqual(loaded["latest_version"], "0.1.2")
        self.assertEqual(loaded["package"], "AgribankV3_0.1.2.zip")
        self.assertEqual(loaded["notes"], ["Sửa lỗi"])
        self.assertEqual(loaded["database_migrations"], [])
        self.assertEqual(loaded["payload_layout"], "source")

    def test_app_manifest_uses_app_root_layout(self) -> None:
        config = self._config(new_version="0.1.2", package_mode="app")

        manifest = build_manifest(config, "AgribankV3_0.1.2.zip")

        self.assertEqual(manifest["payload_layout"], "app_root")

    def test_app_snapshot_uses_separate_file_name(self) -> None:
        app_root = self.source / "dist" / "AgribankV3"
        app_root.mkdir(parents=True)
        (app_root / "AgribankV3.exe").write_text("exe", encoding="utf-8")
        (app_root / APP_BUILD_INFO_FILE_NAME).write_text(
            '{"app":"AgribankV3","version":"0.1.1"}',
            encoding="utf-8",
        )
        files, _ = collect_package_files(self.source, "app")

        snapshot = save_release_file_snapshot(
            self.source,
            "0.1.1",
            files,
            package_mode="app",
        )

        self.assertEqual(snapshot.name, "files_0.1.1_app.json")
        self.assertFalse((snapshot.parent / "files_0.1.1.json").exists())

    def test_app_package_requires_build_info(self) -> None:
        app_root = self.source / "dist" / "AgribankV3"
        app_root.mkdir(parents=True)
        (app_root / "AgribankV3.exe").write_text("exe", encoding="utf-8")

        with self.assertRaisesRegex(UpdateBuilderError, "thiếu agribank_v3_build_info"):
            collect_package_files(self.source, "app")

    def test_app_package_rejects_build_version_mismatch(self) -> None:
        app_root = self.source / "dist" / "AgribankV3"
        app_root.mkdir(parents=True)
        (app_root / "AgribankV3.exe").write_text("exe", encoding="utf-8")
        (app_root / APP_BUILD_INFO_FILE_NAME).write_text(
            '{"app":"AgribankV3","version":"0.1.1"}',
            encoding="utf-8",
        )

        with self.assertRaisesRegex(UpdateBuilderError, "không khớp version gói"):
            UpdateBuilder().validate(
                self._config(new_version="0.1.2", package_mode="app")
            )

    def test_read_app_build_info(self) -> None:
        app_root = self.source / "dist" / "AgribankV3"
        app_root.mkdir(parents=True)
        (app_root / APP_BUILD_INFO_FILE_NAME).write_text(
            '{"app":"AgribankV3","version":"0.1.2"}',
            encoding="utf-8",
        )

        info = read_app_build_info(app_root)

        self.assertEqual(info.version, "0.1.2")

    def test_copy_migration_sql(self) -> None:
        migration = self.root / "0.1.2.sql"
        migration.write_text("CREATE TABLE IF NOT EXISTS demo(id INTEGER);", encoding="utf-8")
        config = self._config(
            new_version="0.1.2",
            migrations=(
                MigrationItem(
                    version="0.1.2",
                    source_file=migration,
                    description="Tạo bảng demo",
                ),
            ),
        )

        copied = copy_migrations(config)

        self.assertEqual(copied, [self.update / "migrations" / "0.1.2.sql"])
        self.assertTrue(copied[0].is_file())

    def test_warn_dangerous_sql(self) -> None:
        migration = self.root / "danger.sql"
        migration.write_text("DROP TABLE credit_groups;", encoding="utf-8")

        warnings = warn_dangerous_sql(migration)

        self.assertTrue(any("DROP" in warning for warning in warnings))

    def test_backup_existing_manifest(self) -> None:
        self.update.mkdir()
        (self.update / "manifest.json").write_text("old manifest", encoding="utf-8")
        (self.update / "AgribankV3_0.1.2.zip").write_text("old zip", encoding="utf-8")

        backup = backup_existing_update_files(self.update, "AgribankV3_0.1.2.zip")

        self.assertIsNotNone(backup)
        assert backup is not None
        self.assertEqual((backup / "manifest.json").read_text(encoding="utf-8"), "old manifest")
        self.assertEqual((backup / "AgribankV3_0.1.2.zip").read_text(encoding="utf-8"), "old zip")

    def test_update_source_version(self) -> None:
        updated = update_source_version(self.source, "0.1.2")

        self.assertIn(self.source / "src" / "agribank_v3" / "__init__.py", updated)
        self.assertIn(self.source / "pyproject.toml", updated)
        self.assertEqual(read_current_version(self.source), "0.1.2")
        self.assertTrue((self.source / "src" / "agribank_v3" / "__init__.py.bak").is_file())
        self.assertTrue((self.source / "pyproject.toml.bak").is_file())

    def test_build_update_end_to_end(self) -> None:
        config = self._config(
            new_version="0.1.2",
            notes=("Sửa lỗi bảng kê", "Bổ sung cập nhật phiên bản"),
        )

        result = UpdateBuilder().build(config)

        self.assertTrue(result.package_path.is_file())
        self.assertTrue(result.manifest_path.is_file())
        with zipfile.ZipFile(result.package_path) as archive:
            names = set(archive.namelist())
        self.assertIn("pyproject.toml", names)
        self.assertNotIn("data/DuLieuV3.db", names)
        manifest = json.loads(result.manifest_path.read_text(encoding="utf-8"))
        self.assertEqual(manifest["latest_version"], "0.1.2")
        self.assertEqual(manifest["package"], "AgribankV3_0.1.2.zip")

    def test_build_update_runs_with_log_callback(self) -> None:
        config = self._config(new_version="0.1.2", notes=("Log callback",))
        messages: list[str] = []
        progress_values: list[int] = []

        result = UpdateBuilder(messages.append, progress_values.append).build(config)

        self.assertTrue(result.package_path.is_file())
        self.assertTrue(any("Đang quét file source" in message for message in messages))
        self.assertTrue(any("Đang tạo zip" in message for message in messages))
        self.assertIn(100, progress_values)

    def test_new_version_equal_source_version_is_valid(self) -> None:
        self._create_manifest("0.1.0")
        self._create_source(self.source, version="0.1.1")
        config = self._config(new_version="0.1.1")

        result = UpdateBuilder().validate(config)

        self.assertEqual(result.current_version, "0.1.1")
        self.assertEqual(result.previous_release_version, "0.1.0")

    def test_new_version_must_be_greater_than_previous_release(self) -> None:
        self._create_manifest("0.1.0")
        self._create_source(self.source, version="0.1.1")

        result = UpdateBuilder().validate(self._config(new_version="0.1.1"))

        self.assertEqual(result.previous_release_version, "0.1.0")

    def test_new_version_equal_previous_release_requires_rebuild_confirmation(self) -> None:
        self._create_manifest("0.1.1")
        self._create_source(self.source, version="0.1.1")

        with self.assertRaisesRegex(Exception, "tạo lại cùng version"):
            UpdateBuilder().validate(self._config(new_version="0.1.1"))

        result = UpdateBuilder().validate(
            self._config(new_version="0.1.1", allow_rebuild_same_version=True)
        )
        self.assertEqual(result.previous_release_version, "0.1.1")

    def test_new_version_less_than_source_version_blocks(self) -> None:
        self._create_manifest("0.1.0")
        self._create_source(self.source, version="0.1.2")

        with self.assertRaisesRegex(Exception, "thấp hơn version trong source"):
            UpdateBuilder().validate(self._config(new_version="0.1.1"))

    def test_new_version_greater_than_source_with_auto_update(self) -> None:
        self._create_manifest("0.1.0")
        config = self._config(new_version="0.1.2", auto_update_source_version=True)

        result = UpdateBuilder().build(config)

        self.assertTrue(result.package_path.is_file())
        self.assertEqual(read_current_version(self.source), "0.1.2")

    def test_new_version_greater_than_source_without_auto_update_warns(self) -> None:
        self._create_manifest("0.1.0")

        result = UpdateBuilder().validate(self._config(new_version="0.1.2"))

        self.assertTrue(
            any("lớn hơn version trong source" in warning for warning in result.warnings)
        )

    def test_detect_previous_release_version_from_manifest(self) -> None:
        self._create_manifest("0.1.0")

        version = detect_previous_release_version(self.update, self.source, "0.1.1")

        self.assertEqual(version, "0.1.0")

    def _config(
        self,
        *,
        new_version: str,
        notes: tuple[str, ...] = ("Ghi chú cập nhật",),
        migrations: tuple[MigrationItem, ...] = (),
        auto_update_source_version: bool = False,
        allow_rebuild_same_version: bool = False,
        package_mode: str = "runtime",
    ) -> BuildUpdateConfig:
        return BuildUpdateConfig(
            source_path=self.source,
            update_path=self.update,
            new_version=new_version,
            release_date="2026-07-20",
            notes=notes,
            migrations=migrations,
            auto_update_source_version=auto_update_source_version,
            allow_rebuild_same_version=allow_rebuild_same_version,
            previous_release_version="0.1.1" if package_mode == "delta" else None,
            package_mode=package_mode,
        )

    def _create_manifest(self, version: str) -> None:
        self.update.mkdir(parents=True, exist_ok=True)
        manifest = {
            "latest_version": version,
            "package": f"AgribankV3_{version}.zip",
            "release_date": "2026-07-20",
            "required_app_restart": True,
            "notes": [],
            "database_migrations": [],
        }
        (self.update / "manifest.json").write_text(
            json.dumps(manifest, ensure_ascii=False),
            encoding="utf-8",
        )

    @staticmethod
    def _create_source(source: Path, *, version: str) -> None:
        package = source / "src" / "agribank_v3"
        package.mkdir(parents=True, exist_ok=True)
        (package / "__init__.py").write_text(
            f'"""Fake AgribankV3."""\n\n__version__ = "{version}"\n',
            encoding="utf-8",
        )
        (package / "__main__.py").write_text("print('hello')\n", encoding="utf-8")
        (source / "pyproject.toml").write_text(
            f'[project]\nname = "agribank-v3"\nversion = "{version}"\n',
            encoding="utf-8",
        )
        (source / "README.md").write_text("# Demo\n", encoding="utf-8")
        for folder in ("data", "logs", "backups", "temp", "KetQua"):
            (source / folder).mkdir(parents=True, exist_ok=True)
        (source / "data" / "DuLieuV3.db").write_text("user db", encoding="utf-8")
        (source / "data" / "quiz.db").write_text("quiz db", encoding="utf-8")
        (source / "logs" / "app.log").write_text("log", encoding="utf-8")
        (source / "backups" / "backup.zip").write_text("backup", encoding="utf-8")
        (source / "temp" / "scratch.tmp").write_text("temp", encoding="utf-8")
        (source / "KetQua" / "report.xlsx").write_text("result", encoding="utf-8")


class AutoUpdateBuilderTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temporary_directory = TemporaryDirectory()
        self.root = Path(self.temporary_directory.name)
        self.source = self.root / "source"
        self.update = self.root / "Update"
        self.snapshots = self.root / "snapshots"
        self.generated_migrations = self.root / "generated_migrations"
        self.baseline_db = self.root / "baseline.db"
        self.dev_db = self.source / "data" / "DuLieuV3.db"
        self._create_source(self.source, version="0.1.1")

    def tearDown(self) -> None:
        self.temporary_directory.cleanup()

    def test_auto_build_no_database_change(self) -> None:
        self._exec(self.baseline_db, "CREATE TABLE customers(id INTEGER PRIMARY KEY, name TEXT)")
        self._exec(self.dev_db, "CREATE TABLE customers(id INTEGER PRIMARY KEY, name TEXT)")
        save_schema_snapshot(self.baseline_db, "0.1.1", self.snapshots)

        result = auto_build_update(self._auto_config("0.1.2"))

        manifest = json.loads(result.build_result.manifest_path.read_text(encoding="utf-8"))
        self.assertEqual(manifest["database_migrations"], [])
        self.assertEqual(result.detection.status, "no_changes")

    def test_auto_build_010_to_011_when_source_is_011(self) -> None:
        self._set_source_version("0.1.1")
        self._exec(self.baseline_db, "CREATE TABLE customers(id INTEGER PRIMARY KEY, name TEXT)")
        self._exec(self.dev_db, "CREATE TABLE customers(id INTEGER PRIMARY KEY, name TEXT)")
        save_schema_snapshot(self.baseline_db, "0.1.0", self.snapshots)

        result = auto_build_update(self._auto_config("0.1.1"))

        manifest = json.loads(result.build_result.manifest_path.read_text(encoding="utf-8"))
        self.assertEqual(manifest["latest_version"], "0.1.1")
        self.assertEqual(result.build_result.warnings, ())

    def test_auto_build_detects_new_column(self) -> None:
        self._exec(self.baseline_db, "CREATE TABLE customers(id INTEGER PRIMARY KEY, name TEXT)")
        self._exec(
            self.dev_db,
            "CREATE TABLE customers(id INTEGER PRIMARY KEY, name TEXT, is_active INTEGER DEFAULT 1)",
        )
        save_schema_snapshot(self.baseline_db, "0.1.1", self.snapshots)

        result = auto_build_update(self._auto_config("0.1.2"))

        manifest = json.loads(result.build_result.manifest_path.read_text(encoding="utf-8"))
        self.assertEqual(manifest["database_migrations"][0]["version"], "0.1.2")
        self.assertEqual(manifest["database_migrations"][0]["file"], "")
        self.assertTrue(source_has_python_migration(self.source, "0.1.2"))
        self.assertTrue(result.generated_migration_path and result.generated_migration_path.is_file())

    def test_auto_build_detects_new_table(self) -> None:
        self._exec(self.baseline_db, "CREATE TABLE customers(id INTEGER PRIMARY KEY, name TEXT)")
        self._exec(
            self.dev_db,
            """
            CREATE TABLE customers(id INTEGER PRIMARY KEY, name TEXT);
            CREATE TABLE demo(id INTEGER PRIMARY KEY, title TEXT);
            """,
        )
        save_schema_snapshot(self.baseline_db, "0.1.1", self.snapshots)

        result = auto_build_update(self._auto_config("0.1.2"))

        self.assertEqual(result.detection.status, "safe_changes")
        self.assertTrue(source_has_python_migration(self.source, "0.1.2"))

    def test_auto_build_detects_new_app_preference(self) -> None:
        self._exec(
            self.baseline_db,
            "CREATE TABLE app_preferences(key TEXT PRIMARY KEY, value TEXT, updated_at TEXT)",
        )
        self._exec(
            self.dev_db,
            """
            CREATE TABLE app_preferences(key TEXT PRIMARY KEY, value TEXT, updated_at TEXT);
            INSERT INTO app_preferences(key, value, updated_at)
            VALUES ('update_path', 'X:/Update', '2026-07-20');
            """,
        )
        save_schema_snapshot(self.baseline_db, "0.1.1", self.snapshots)

        result = auto_build_update(self._auto_config("0.1.2"))

        migration_text = result.generated_migration_path.read_text(encoding="utf-8")
        self.assertIn("INSERT OR IGNORE INTO app_preferences", migration_text)

    def test_auto_build_blocks_dangerous_change(self) -> None:
        self._exec(
            self.baseline_db,
            "CREATE TABLE customers(id INTEGER PRIMARY KEY, name TEXT, old_column TEXT)",
        )
        self._exec(self.dev_db, "CREATE TABLE customers(id INTEGER PRIMARY KEY, name TEXT)")
        save_schema_snapshot(self.baseline_db, "0.1.1", self.snapshots)

        with self.assertRaisesRegex(Exception, "xử lý thủ công"):
            auto_build_update(self._auto_config("0.1.2"))

        self.assertFalse((self.update / "AgribankV3_0.1.2.zip").exists())

    def test_auto_build_saves_new_schema_snapshot(self) -> None:
        self._exec(self.baseline_db, "CREATE TABLE customers(id INTEGER PRIMARY KEY, name TEXT)")
        self._exec(self.dev_db, "CREATE TABLE customers(id INTEGER PRIMARY KEY, name TEXT)")
        save_schema_snapshot(self.baseline_db, "0.1.1", self.snapshots)

        result = auto_build_update(self._auto_config("0.1.2"))

        self.assertEqual(result.saved_snapshot_path, self.snapshots / "schema_0.1.2.json")
        self.assertTrue((self.snapshots / "schema_0.1.2.json").is_file())

    def test_auto_build_missing_snapshot_returns_action_required(self) -> None:
        self._exec(self.dev_db, "CREATE TABLE customers(id INTEGER PRIMARY KEY, name TEXT)")

        result = auto_detect_database_changes(
            self.source,
            self.update,
            "0.1.1",
            "0.1.2",
            snapshot_dir=self.snapshots,
        )

        self.assertEqual(result.status, "action_required")

    def test_auto_build_does_not_package_database(self) -> None:
        self._exec(self.baseline_db, "CREATE TABLE customers(id INTEGER PRIMARY KEY, name TEXT)")
        self._exec(self.dev_db, "CREATE TABLE customers(id INTEGER PRIMARY KEY, name TEXT)")
        save_schema_snapshot(self.baseline_db, "0.1.1", self.snapshots)

        result = auto_build_update(self._auto_config("0.1.2"))

        with zipfile.ZipFile(result.build_result.package_path) as archive:
            names = set(archive.namelist())
        self.assertFalse(any(name.endswith(".db") for name in names))
        self.assertNotIn("data/DuLieuV3.db", names)

    def test_cli_auto_detect_db(self) -> None:
        self._exec(self.baseline_db, "CREATE TABLE customers(id INTEGER PRIMARY KEY, name TEXT)")
        self._exec(self.dev_db, "CREATE TABLE customers(id INTEGER PRIMARY KEY, name TEXT)")
        save_schema_snapshot(self.baseline_db, "0.1.1", self.snapshots)

        exit_code = update_builder_main(
            [
                "--source",
                str(self.source),
                "--version",
                "0.1.2",
                "--update-path",
                str(self.update),
                "--notes",
                "Auto build",
                "--auto-detect-db",
                "--dev-db",
                str(self.dev_db),
                "--snapshot-dir",
                str(self.snapshots),
            ]
        )

        self.assertEqual(exit_code, 0)
        self.assertTrue((self.update / "manifest.json").is_file())

    def test_generated_python_migration_registered(self) -> None:
        self._exec(self.baseline_db, "CREATE TABLE customers(id INTEGER PRIMARY KEY, name TEXT)")
        self._exec(
            self.dev_db,
            "CREATE TABLE customers(id INTEGER PRIMARY KEY, name TEXT, is_active INTEGER DEFAULT 1)",
        )
        save_schema_snapshot(self.baseline_db, "0.1.1", self.snapshots)

        auto_build_update(self._auto_config("0.1.2"))

        text = (
            self.source / "src" / "agribank_v3" / "update" / "db_migrations.py"
        ).read_text(encoding="utf-8")
        self.assertIn("def migrate_0_1_2(conn):", text)
        self.assertIn('"0.1.2": migrate_0_1_2', text)

    def _auto_config(self, new_version: str) -> AutoBuildConfig:
        return AutoBuildConfig(
            source_path=self.source,
            update_path=self.update,
            new_version=new_version,
            release_date="2026-07-20",
            notes=("Auto build",),
            snapshot_dir=self.snapshots,
            generated_migration_dir=self.generated_migrations,
        )

    def _set_source_version(self, version: str) -> None:
        (self.source / "src" / "agribank_v3" / "__init__.py").write_text(
            f'"""Fake AgribankV3."""\n\n__version__ = "{version}"\n',
            encoding="utf-8",
        )
        (self.source / "pyproject.toml").write_text(
            f'[project]\nname = "agribank-v3"\nversion = "{version}"\n',
            encoding="utf-8",
        )

    @staticmethod
    def _create_source(source: Path, *, version: str) -> None:
        package = source / "src" / "agribank_v3"
        update_package = package / "update"
        update_package.mkdir(parents=True)
        (source / "data").mkdir(parents=True)
        (package / "__init__.py").write_text(
            f'"""Fake AgribankV3."""\n\n__version__ = "{version}"\n',
            encoding="utf-8",
        )
        (source / "pyproject.toml").write_text(
            f'[project]\nname = "agribank-v3"\nversion = "{version}"\n',
            encoding="utf-8",
        )
        (source / "README.md").write_text("# Demo\n", encoding="utf-8")
        (update_package / "db_migrations.py").write_text(
            """
from __future__ import annotations

import sqlite3
from typing import Callable

PythonMigration = Callable[[sqlite3.Connection], None]


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


def default_python_migrations() -> dict[str, PythonMigration]:
    return {
    }
""".lstrip(),
            encoding="utf-8",
        )

    @staticmethod
    def _exec(path: Path, script: str) -> None:
        path.parent.mkdir(parents=True, exist_ok=True)
        with closing(sqlite3.connect(path)) as connection:
            with connection:
                connection.executescript(script)


class UpdateBuilderWorkerTests(unittest.TestCase):
    def test_build_worker_emits_finished(self) -> None:
        results: list[object] = []
        worker = UpdateTaskWorker(lambda log, progress: "done")
        worker.finished.connect(results.append)

        worker.run()

        self.assertEqual(results, ["done"])

    def test_build_worker_emits_failed(self) -> None:
        errors: list[str] = []

        def task(log, progress):
            raise RuntimeError("boom")

        worker = UpdateTaskWorker(task)
        worker.failed.connect(errors.append)

        worker.run()

        self.assertEqual(errors, ["boom"])

    def test_worker_does_not_show_messagebox_directly(self) -> None:
        original_warning = QMessageBox.warning
        QMessageBox.warning = staticmethod(
            lambda *args, **kwargs: (_ for _ in ()).throw(
                AssertionError("worker must not call QMessageBox.warning")
            )
        )

        def task(log, progress):
            raise RuntimeError("boom")

        errors: list[str] = []
        worker = UpdateTaskWorker(task)
        worker.failed.connect(errors.append)
        try:
            worker.run()
        finally:
            QMessageBox.warning = original_warning

        self.assertEqual(errors, ["boom"])


class UpdateBuilderUiBasicValidationTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.app = QApplication.instance() or QApplication([])

    def setUp(self) -> None:
        self.temporary_directory = TemporaryDirectory()
        self.update_dir = Path(self.temporary_directory.name) / "Update"
        self.update_dir.mkdir(parents=True)
        self.window = UpdateBuilderWindow()
        self.window.source_edit.setText(str(Path.cwd()))
        self.window.new_version_edit.setText("9.9.9")
        self.window.update_path_edit.setText(str(self.update_dir))
        self.window.notes_edit.setPlainText("Ghi chú test")

    def tearDown(self) -> None:
        self.window.close()
        self.temporary_directory.cleanup()

    def test_basic_validation_release_notes_required(self) -> None:
        self.window.notes_edit.clear()

        error = self.window._validate_form_basic()

        self.assertEqual(error, "Release notes không được để trống.")

    def test_basic_validation_new_version_required(self) -> None:
        self.window.new_version_edit.clear()

        error = self.window._validate_form_basic()

        self.assertEqual(error, "Phiên bản mới không được để trống.")

    def test_basic_validation_source_required(self) -> None:
        self.window.source_edit.clear()

        error = self.window._validate_form_basic()

        self.assertEqual(error, "Thư mục source AgribankV3 không được để trống.")

    def test_build_button_does_not_start_worker_when_notes_empty(self) -> None:
        calls: list[str] = []
        original_warning = QMessageBox.warning
        QMessageBox.warning = staticmethod(lambda *args, **kwargs: QMessageBox.StandardButton.Ok)
        self.window._start_worker = lambda *args, **kwargs: calls.append("started")
        self.window.notes_edit.clear()
        try:
            self.window.build_update()
        finally:
            QMessageBox.warning = original_warning

        self.assertEqual(calls, [])
        self.assertIsNone(self.window._worker_thread)
        self.assertFalse(self.window._is_running)

    def test_no_continue_after_validation_error(self) -> None:
        calls: list[str] = []
        self.window._start_worker = lambda *args, **kwargs: calls.append("started")
        self.window.notes_edit.clear()

        self.window.build_update()

        log_text = self.window.log_edit.toPlainText()
        self.assertEqual(calls, [])
        self.assertIn("Lỗi kiểm tra dữ liệu: Release notes không được để trống.", log_text)
        self.assertNotIn("Đang kiểm tra dữ liệu trước khi tạo", log_text)
        self.assertNotIn("Đang tạo zip", log_text)

    def test_missing_update_drive_handled_gracefully(self) -> None:
        missing_drive = next(
            (drive for drive in "ZYXWVUTSRQPONMLKJIHGFEDCBA" if not Path(f"{drive}:\\").exists()),
            None,
        )
        if missing_drive is None:
            self.skipTest("Không tìm thấy ổ đĩa trống để test.")
        self.window.update_path_edit.setText(f"{missing_drive}:\\public\\AgribankV3\\Update")

        error = self.window._validate_form_basic()

        self.assertEqual(
            error,
            (
                f"Không tìm thấy ổ đĩa {missing_drive}:. "
                "Vui lòng kiểm tra kết nối mạng hoặc chọn thư mục Update khác."
            ),
        )
        self.assertNotIn("WinError", error)

    def test_missing_update_folder_can_prompt_create(self) -> None:
        target = Path(self.temporary_directory.name) / "new" / "Update"
        self.window.update_path_edit.setText(str(target))
        original_question = QMessageBox.question
        QMessageBox.question = staticmethod(
            lambda *args, **kwargs: QMessageBox.StandardButton.Yes
        )
        try:
            result = self.window._ensure_update_folder_ready("Tạo bản cập nhật")
        finally:
            QMessageBox.question = original_question

        self.assertTrue(result)
        self.assertTrue(target.is_dir())

    def test_worker_failed_reenables_ui(self) -> None:
        calls: list[str] = []
        self.window._set_busy(True, "Đang chạy")
        self.window._is_running = True
        self.window._worker_on_failure = calls.append

        self.window._worker_failed("boom")
        QApplication.processEvents()

        self.assertEqual(calls, ["boom"])
        self.assertFalse(self.window._is_running)
        self.assertIsNone(self.window._worker_thread)
        self.assertTrue(all(widget.isEnabled() for widget in self.window._busy_widgets))

    def test_non_ui_builder_modules_do_not_import_qmessagebox(self) -> None:
        for filename in ("update_builder_core.py", "db_schema_diff.py", "update_builder_app.py"):
            text = (BUILDER_DIR / filename).read_text(encoding="utf-8")
            self.assertNotIn("QMessageBox", text)

    def test_no_duplicate_signal_connections(self) -> None:
        calls: list[str] = []
        self.window._start_worker = lambda *args, **kwargs: calls.append("started")

        self.window.build_update()

        self.assertEqual(calls, ["started"])

    def test_validate_config_worker_finishes_without_crash(self) -> None:
        original_information = QMessageBox.information
        original_warning = QMessageBox.warning
        original_question = QMessageBox.question
        QMessageBox.information = staticmethod(
            lambda *args, **kwargs: QMessageBox.StandardButton.Ok
        )
        QMessageBox.warning = staticmethod(
            lambda *args, **kwargs: QMessageBox.StandardButton.Ok
        )
        QMessageBox.question = staticmethod(
            lambda *args, **kwargs: QMessageBox.StandardButton.Yes
        )
        try:
            self.window.validate_config()
            deadline = time.time() + 15
            while self.window._is_running and time.time() < deadline:
                QApplication.processEvents()
                time.sleep(0.01)
            for _ in range(10):
                QApplication.processEvents()
                time.sleep(0.01)
        finally:
            QMessageBox.information = original_information
            QMessageBox.warning = original_warning
            QMessageBox.question = original_question

        self.assertFalse(self.window._is_running)
        self.assertIsNone(self.window._worker_thread)
        self.assertEqual(self.window.status_label.text(), "Hoàn thành.")


if __name__ == "__main__":
    unittest.main()
