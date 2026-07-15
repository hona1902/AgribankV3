from __future__ import annotations

from collections.abc import Callable, Iterable
from dataclasses import dataclass
from pathlib import Path

import pythoncom
import win32com.client


WORD_SUFFIXES = {".doc", ".docx", ".docm", ".rtf"}


class WordPrintError(RuntimeError):
    """Raised when Word documents cannot be printed."""


@dataclass(frozen=True, slots=True)
class WordPrintResult:
    folder_path: Path
    file_count: int
    printed_count: int
    failed: tuple[tuple[Path, str], ...] = ()


def list_word_files(folder_path: Path) -> tuple[Path, ...]:
    folder = Path(folder_path)
    if not folder.is_dir():
        raise WordPrintError(f"Không tìm thấy thư mục: {folder}")
    files = [
        path
        for path in folder.iterdir()
        if path.is_file()
        and path.suffix.casefold() in WORD_SUFFIXES
        and not path.name.startswith("~$")
    ]
    return tuple(sorted(files, key=lambda path: path.name.casefold()))


def print_word_files(
    source_paths: Iterable[Path],
    *,
    progress: Callable[[str], None] | None = None,
) -> WordPrintResult:
    sources = tuple(Path(path) for path in source_paths)
    if not sources:
        raise WordPrintError("Không có file Word để in.")
    folder_path = sources[0].parent
    pythoncom.CoInitialize()
    word = None
    printed_count = 0
    failed: list[tuple[Path, str]] = []
    try:
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0
        for index, source in enumerate(sources, start=1):
            if progress is not None:
                progress(f"Đang in {index}/{len(sources)}: {source.name}")
            if not source.is_file():
                failed.append((source, "Không tìm thấy file."))
                continue
            document = None
            try:
                document = word.Documents.Open(
                    str(source),
                    ConfirmConversions=False,
                    ReadOnly=True,
                    AddToRecentFiles=False,
                    Visible=False,
                )
                document.PrintOut(Background=False)
                printed_count += 1
            except Exception as exc:
                failed.append((source, str(exc)))
            finally:
                if document is not None:
                    try:
                        document.Close(False)
                    except Exception:
                        pass
    except Exception as exc:
        raise WordPrintError(f"Không mở được Microsoft Word để in: {exc}") from exc
    finally:
        if word is not None:
            try:
                word.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()
    return WordPrintResult(
        folder_path=folder_path,
        file_count=len(sources),
        printed_count=printed_count,
        failed=tuple(failed),
    )
