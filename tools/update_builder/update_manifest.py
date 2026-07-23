from __future__ import annotations

from dataclasses import dataclass


@dataclass(frozen=True, slots=True)
class ManifestMigration:
    version: str
    file: str = ""
    description: str = ""


def manifest_dict(
    *,
    latest_version: str,
    package: str,
    release_date: str,
    required_app_restart: bool,
    notes: list[str],
    migrations: list[ManifestMigration],
    package_type: str = "full",
    payload_layout: str = "source",
    base_version: str = "",
    delete_files: list[str] | None = None,
) -> dict[str, object]:
    payload: dict[str, object] = {
        "latest_version": latest_version,
        "package": package,
        "package_type": package_type,
        "payload_layout": payload_layout,
        "release_date": release_date,
        "required_app_restart": required_app_restart,
        "notes": notes,
        "database_migrations": [
            {
                "version": migration.version,
                "file": migration.file,
                "description": migration.description,
            }
            for migration in migrations
        ],
    }
    if base_version:
        payload["base_version"] = base_version
    if delete_files:
        payload["delete_files"] = delete_files
    return payload
