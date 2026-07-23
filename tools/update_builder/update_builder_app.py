from __future__ import annotations

import argparse
from pathlib import Path
import sys

from update_builder_core import (
    DEFAULT_UPDATE_PATH,
    AutoBuildConfig,
    BuildUpdateConfig,
    MigrationItem,
    UpdateBuilder,
    UpdateBuilderError,
    auto_build_update,
    config_from_json,
    normalize_package_mode,
    save_schema_snapshot,
)


def main(argv: list[str] | None = None) -> int:
    _configure_console_encoding()
    parser = argparse.ArgumentParser(description="AgribankV3 Update Builder")
    parser.add_argument("--config", help="Đường dẫn build_config.json")
    parser.add_argument("--source", help="Thư mục source AgribankV3")
    parser.add_argument("--version", help="Version mới, ví dụ 0.1.2")
    parser.add_argument("--update-path", default=DEFAULT_UPDATE_PATH)
    parser.add_argument("--release-date")
    parser.add_argument("--notes", action="append", default=[])
    parser.add_argument("--migration-version", action="append", default=[])
    parser.add_argument("--migration-file", action="append", default=[])
    parser.add_argument("--migration-description", action="append", default=[])
    parser.add_argument("--python-migration", action="append", default=[])
    parser.add_argument("--auto-update-source-version", action="store_true")
    parser.add_argument("--verbose", action="store_true")
    parser.add_argument("--previous-release-version")
    parser.add_argument("--allow-rebuild-same-version", action="store_true")
    parser.add_argument(
        "--package-mode",
        choices=("runtime", "source", "delta", "app", "exe", "frozen", "pyinstaller"),
        default="runtime",
        help="Kiểu gói cập nhật: runtime, source, delta hoặc app/exe.",
    )
    parser.add_argument("--auto-detect-db", action="store_true")
    parser.add_argument("--dev-db", help="Database dev hiện tại để auto-detect schema")
    parser.add_argument("--baseline-db", help="Database bản cũ dùng làm baseline")
    parser.add_argument("--create-baseline", action="store_true")
    parser.add_argument("--snapshot-dir", help="Thư mục lưu schema snapshot")
    parser.add_argument(
        "--code-only-if-missing-baseline",
        action="store_true",
        help="Cho phép tạo update chỉ đổi code nếu thiếu snapshot/baseline.",
    )
    args = parser.parse_args(argv)

    if not any(
        (
            args.config,
            args.source,
            args.version,
            args.notes,
            args.migration_version,
            args.auto_detect_db,
            args.create_baseline,
        )
    ):
        return run_ui()

    try:
        if args.create_baseline:
            if not args.baseline_db or not args.version:
                raise UpdateBuilderError("--create-baseline cần --baseline-db và --version.")
            snapshot = save_schema_snapshot(args.baseline_db, args.version, args.snapshot_dir)
            print(f"Đã tạo schema snapshot: {snapshot}")
            return 0
        if args.auto_detect_db:
            result = auto_build_update(_auto_config_from_args(args), print)
            print(f"Đã tạo package: {result.build_result.package_path}")
            print(f"Đã tạo manifest: {result.build_result.manifest_path}")
            if result.saved_snapshot_path:
                print(f"Schema snapshot: {result.saved_snapshot_path}")
            print(f"Log: {result.build_result.log_path}")
            return 0
        config = config_from_json(args.config) if args.config else _config_from_args(args)
        result = UpdateBuilder(print).build(config)
    except UpdateBuilderError as exc:
        print(f"Lỗi: {exc}", file=sys.stderr)
        return 1
    print(f"Đã tạo package: {result.package_path}")
    print(f"Đã tạo manifest: {result.manifest_path}")
    print(f"Log: {result.log_path}")
    return 0


def run_ui() -> int:
    from PySide6.QtWidgets import QApplication
    from update_builder_ui import UpdateBuilderWindow

    app = QApplication(sys.argv)
    window = UpdateBuilderWindow()
    window.show()
    return app.exec()


def _config_from_args(args: argparse.Namespace) -> BuildUpdateConfig:
    if not args.source or not args.version:
        raise UpdateBuilderError("CLI cần --source và --version.")
    migrations = []
    max_count = max(
        len(args.migration_version),
        len(args.migration_file),
        len(args.migration_description),
        len(args.python_migration),
        0,
    )
    python_versions = set(args.python_migration)
    for index in range(max_count):
        version = _item(args.migration_version, index)
        file_text = _item(args.migration_file, index)
        description = _item(args.migration_description, index)
        use_python = version in python_versions or (version and not file_text)
        migrations.append(
            MigrationItem(
                version=version,
                description=description,
                source_file=Path(file_text) if file_text and not use_python else None,
                use_python_migration=use_python,
            )
        )
    return BuildUpdateConfig(
        source_path=Path(args.source),
        update_path=Path(args.update_path),
        new_version=args.version,
        release_date=args.release_date or __import__("datetime").date.today().isoformat(),
        notes=tuple(args.notes),
        migrations=tuple(migrations),
        auto_update_source_version=args.auto_update_source_version,
        previous_release_version=args.previous_release_version,
        allow_rebuild_same_version=args.allow_rebuild_same_version,
        package_mode=normalize_package_mode(args.package_mode),
        verbose=args.verbose,
    )


def _auto_config_from_args(args: argparse.Namespace) -> AutoBuildConfig:
    if not args.source or not args.version:
        raise UpdateBuilderError("CLI auto-detect cần --source và --version.")
    return AutoBuildConfig(
        source_path=Path(args.source),
        update_path=Path(args.update_path),
        new_version=args.version,
        release_date=args.release_date or __import__("datetime").date.today().isoformat(),
        notes=tuple(args.notes),
        auto_update_source_version=args.auto_update_source_version,
        dev_db_path=Path(args.dev_db) if args.dev_db else None,
        baseline_db_path=Path(args.baseline_db) if args.baseline_db else None,
        snapshot_dir=Path(args.snapshot_dir) if args.snapshot_dir else None,
        code_only_if_missing_baseline=args.code_only_if_missing_baseline,
        previous_release_version=args.previous_release_version,
        allow_rebuild_same_version=args.allow_rebuild_same_version,
        package_mode=normalize_package_mode(args.package_mode),
        verbose=args.verbose,
    )


def _item(values: list[str], index: int) -> str:
    return values[index] if index < len(values) else ""


def _configure_console_encoding() -> None:
    for stream in (sys.stdout, sys.stderr):
        reconfigure = getattr(stream, "reconfigure", None)
        if callable(reconfigure):
            reconfigure(encoding="utf-8", errors="replace")


if __name__ == "__main__":
    raise SystemExit(main())
