from __future__ import annotations

from dataclasses import dataclass
import json
from pathlib import Path

from agribank_v3.update.db_migrations import MigrationSpec


class UpdateManifestError(RuntimeError):
    pass


@dataclass(frozen=True, slots=True)
class UpdateManifest:
    latest_version: str
    package: str
    package_type: str = "full"
    payload_layout: str = "auto"
    base_version: str = ""
    release_date: str = ""
    required_app_restart: bool = True
    notes: tuple[str, ...] = ()
    delete_files: tuple[Path, ...] = ()
    database_migrations: tuple[MigrationSpec, ...] = ()


def read_update_manifest(update_path: Path | str) -> UpdateManifest:
    root = Path(update_path)
    manifest_path = root / "manifest.json"
    if not manifest_path.is_file():
        raise UpdateManifestError(
            f"Không tìm thấy manifest.json trong thư mục cập nhật: {root}"
        )
    try:
        raw = json.loads(manifest_path.read_text(encoding="utf-8-sig"))
    except (OSError, json.JSONDecodeError) as exc:
        raise UpdateManifestError(f"Lỗi đọc thông tin cập nhật: {exc}") from exc
    if not isinstance(raw, dict):
        raise UpdateManifestError("manifest.json không đúng cấu trúc object.")
    latest_version = str(raw.get("latest_version") or "").strip()
    package = str(raw.get("package") or "").strip()
    if not latest_version:
        raise UpdateManifestError("manifest.json thiếu latest_version.")
    if not package:
        raise UpdateManifestError("manifest.json thiếu package.")
    notes_raw = raw.get("notes", ())
    notes = tuple(str(item) for item in notes_raw) if isinstance(notes_raw, list) else ()
    package_type = str(raw.get("package_type") or "full").strip().casefold()
    if package_type not in {"full", "delta"}:
        raise UpdateManifestError(f"package_type không hợp lệ: {package_type}")
    payload_layout = str(raw.get("payload_layout") or "auto").strip().casefold()
    payload_layout_aliases = {
        "runtime": "source",
        "source_root": "source",
        "app": "app_root",
        "exe": "app_root",
        "frozen": "app_root",
        "pyinstaller": "app_root",
    }
    payload_layout = payload_layout_aliases.get(payload_layout, payload_layout)
    if payload_layout not in {"auto", "source", "app_root"}:
        raise UpdateManifestError(f"payload_layout không hợp lệ: {payload_layout}")
    delete_files_raw = raw.get("delete_files", ()) or ()
    delete_files = (
        tuple(Path(str(item)) for item in delete_files_raw)
        if isinstance(delete_files_raw, list)
        else ()
    )
    migrations: list[MigrationSpec] = []
    for item in raw.get("database_migrations", ()) or ():
        if not isinstance(item, dict):
            continue
        migrations.append(
            MigrationSpec(
                version=str(item.get("version") or "").strip(),
                file=str(item.get("file") or "").strip(),
                description=str(item.get("description") or "").strip(),
            )
        )
    return UpdateManifest(
        latest_version=latest_version,
        package=package,
        package_type=package_type,
        payload_layout=payload_layout,
        base_version=str(raw.get("base_version") or "").strip(),
        release_date=str(raw.get("release_date") or "").strip(),
        required_app_restart=bool(raw.get("required_app_restart", True)),
        notes=notes,
        delete_files=delete_files,
        database_migrations=tuple(migrations),
    )
