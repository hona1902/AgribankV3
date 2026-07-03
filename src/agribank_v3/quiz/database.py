from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
from contextlib import contextmanager
import hashlib
import json
import os
from pathlib import Path
import shutil
import sqlite3
from typing import Any, Iterator

import win32com.client
from pywintypes import com_error

from agribank_v3.quiz.models import Question, QuizResult
from agribank_v3.quiz.session import QuizSession


class QuizDatabaseError(RuntimeError):
    pass


@dataclass(frozen=True, slots=True)
class SyncResult:
    imported: bool
    question_count: int
    employee_count: int
    legacy_attempt_count: int
    source_hash: str


class QuizDatabase:
    SCHEMA_VERSION = 4
    ACCESS_PROVIDERS = ("Microsoft.ACE.OLEDB.12.0", "Microsoft.ACE.OLEDB.16.0")

    def __init__(
        self,
        sqlite_path: Path | None = None,
        access_path: Path | None = None,
    ) -> None:
        app_root = Path(__file__).resolve().parents[3]
        project_root = Path(__file__).resolve().parents[4]
        self.sqlite_path = sqlite_path or app_root / "data" / "agribank_v3.sqlite3"
        access_override = os.environ.get("AGRIBANKV3_ACCESS_DB")
        self.access_path = (
            Path(access_override)
            if access_override
            else access_path or project_root / "Data" / "AgribankMenuData.mdb"
        )
        self.sqlite_path.parent.mkdir(parents=True, exist_ok=True)
        self._initialize_schema()

    def connect(self) -> sqlite3.Connection:
        connection = sqlite3.connect(self.sqlite_path)
        connection.row_factory = sqlite3.Row
        connection.execute("PRAGMA foreign_keys = ON")
        connection.execute("PRAGMA journal_mode = WAL")
        return connection

    @contextmanager
    def _database(self) -> Iterator[sqlite3.Connection]:
        database = self.connect()
        try:
            with database:
                yield database
        finally:
            database.close()

    def _initialize_schema(self) -> None:
        with self._database() as database:
            current_version = int(
                database.execute("PRAGMA user_version").fetchone()[0]
            )
            database.executescript(
                """
                CREATE TABLE IF NOT EXISTS questions (
                    id INTEGER PRIMARY KEY,
                    legacy_id REAL,
                    question_number INTEGER NOT NULL,
                    question_text TEXT NOT NULL,
                    option_a TEXT,
                    option_b TEXT,
                    option_c TEXT,
                    option_d TEXT,
                    correct_answer TEXT NOT NULL,
                    topic_code TEXT NOT NULL DEFAULT '',
                    topic_name TEXT NOT NULL DEFAULT '',
                    source_reference TEXT NOT NULL DEFAULT '',
                    source_active INTEGER NOT NULL DEFAULT 1
                );
                CREATE INDEX IF NOT EXISTS idx_questions_topic
                    ON questions(topic_name, question_number);
                CREATE INDEX IF NOT EXISTS idx_questions_code
                    ON questions(topic_code);

                CREATE TABLE IF NOT EXISTS business_topics (
                    code TEXT PRIMARY KEY,
                    name TEXT NOT NULL,
                    active INTEGER NOT NULL DEFAULT 1,
                    sort_order INTEGER NOT NULL DEFAULT 0
                );

                CREATE TABLE IF NOT EXISTS question_topics (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    business_code TEXT NOT NULL,
                    name TEXT NOT NULL,
                    active INTEGER NOT NULL DEFAULT 1,
                    sort_order INTEGER NOT NULL DEFAULT 0,
                    UNIQUE (business_code, name),
                    FOREIGN KEY (business_code) REFERENCES business_topics(code)
                        ON UPDATE CASCADE
                );

                CREATE TABLE IF NOT EXISTS quiz_settings (
                    key TEXT PRIMARY KEY,
                    value TEXT NOT NULL
                );

                CREATE TABLE IF NOT EXISTS employees (
                    id INTEGER PRIMARY KEY,
                    full_name TEXT NOT NULL
                );

                CREATE TABLE IF NOT EXISTS quiz_attempts (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    legacy_access_id INTEGER,
                    employee_name TEXT NOT NULL DEFAULT '',
                    topic_name TEXT NOT NULL,
                    started_at TEXT NOT NULL,
                    finished_at TEXT NOT NULL,
                    duration_seconds INTEGER NOT NULL,
                    total_questions INTEGER NOT NULL,
                    answered_questions INTEGER NOT NULL,
                    correct_answers INTEGER NOT NULL,
                    incorrect_answers INTEGER NOT NULL,
                    unanswered_questions INTEGER NOT NULL,
                    percentage REAL NOT NULL,
                    rating TEXT NOT NULL
                );

                CREATE TABLE IF NOT EXISTS quiz_answers (
                    attempt_id INTEGER NOT NULL,
                    question_id INTEGER NOT NULL,
                    display_order INTEGER NOT NULL,
                    selected_answer TEXT NOT NULL DEFAULT '',
                    correct_answer TEXT NOT NULL,
                    is_correct INTEGER NOT NULL,
                    checked_before_finish INTEGER NOT NULL,
                    PRIMARY KEY (attempt_id, question_id),
                    FOREIGN KEY (attempt_id) REFERENCES quiz_attempts(id)
                        ON DELETE CASCADE,
                    FOREIGN KEY (question_id) REFERENCES questions(id)
                );

                CREATE TABLE IF NOT EXISTS sync_metadata (
                    key TEXT PRIMARY KEY,
                    value TEXT NOT NULL
                );
                """
            )
            columns = {
                str(row[1])
                for row in database.execute("PRAGMA table_info(questions)")
            }
            if "source_active" not in columns:
                database.execute(
                    "ALTER TABLE questions ADD COLUMN source_active "
                    "INTEGER NOT NULL DEFAULT 1"
                )
            if "subject_id" not in columns:
                database.execute(
                    "ALTER TABLE questions ADD COLUMN subject_id INTEGER"
                )
            if "locked_answers" not in columns:
                database.execute(
                    "ALTER TABLE questions ADD COLUMN locked_answers "
                    "TEXT NOT NULL DEFAULT ''"
                )
            attempt_columns = {
                str(row[1])
                for row in database.execute("PRAGMA table_info(quiz_attempts)")
            }
            if "legacy_access_id" not in attempt_columns:
                database.execute(
                    "ALTER TABLE quiz_attempts ADD COLUMN legacy_access_id INTEGER"
                )
            database.execute(
                "CREATE UNIQUE INDEX IF NOT EXISTS idx_attempts_legacy_access "
                "ON quiz_attempts(legacy_access_id) "
                "WHERE legacy_access_id IS NOT NULL"
            )
            database.execute(
                """
                INSERT OR IGNORE INTO business_topics (code, name, sort_order)
                SELECT topic_code, MIN(topic_name), MIN(id)
                FROM questions
                WHERE topic_code <> ''
                GROUP BY topic_code
                """
            )
            if current_version < 4:
                database.execute(
                    """
                    INSERT OR IGNORE INTO question_topics (
                        business_code, name, sort_order
                    )
                    SELECT topic_code, topic_name, MIN(id)
                    FROM questions
                    WHERE topic_code <> '' AND topic_name <> ''
                    GROUP BY topic_code, topic_name
                    """
                )
                database.execute(
                    """
                    UPDATE questions
                    SET subject_id = (
                        SELECT s.id
                        FROM question_topics s
                        WHERE s.business_code = questions.topic_code
                          AND s.name = questions.topic_name
                        LIMIT 1
                    )
                    WHERE subject_id IS NULL
                    """
                )
                database.execute(
                    """
                    UPDATE questions
                    SET topic_name = COALESCE((
                        SELECT t.name FROM business_topics t
                        WHERE t.code = questions.topic_code
                    ), topic_name)
                    """
                )
            database.execute(f"PRAGMA user_version = {self.SCHEMA_VERSION}")

    def sync_from_access(self, force: bool = False) -> SyncResult:
        if not self.access_path.is_file():
            question_count = self.question_count()
            if question_count:
                return SyncResult(
                    False,
                    question_count,
                    self.employee_count(),
                    self.legacy_attempt_count(),
                    "",
                )
            raise QuizDatabaseError(
                f"Không tìm thấy cơ sở dữ liệu Access:\n{self.access_path}"
            )

        source_hash = self._sha256(self.access_path)
        current_hash = self._metadata("access_sha256")
        if not force and current_hash == source_hash and self.question_count():
            return SyncResult(
                False,
                self.question_count(),
                self.employee_count(),
                self.legacy_attempt_count(),
                source_hash,
            )

        questions, employees, legacy_attempts = self._read_access()
        if not questions:
            raise QuizDatabaseError("Không đọc được câu hỏi nào từ Access.")

        if self.sqlite_path.exists() and self.question_count():
            backup_directory = self.sqlite_path.parent / "backups"
            backup_directory.mkdir(exist_ok=True)
            stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
            shutil.copy2(
                self.sqlite_path,
                backup_directory / f"agribank_v3-{stamp}.sqlite3",
            )

        with self._database() as database:
            database.execute("BEGIN IMMEDIATE")
            database.execute("UPDATE questions SET source_active = 0")
            database.execute("DELETE FROM employees")
            database.executemany(
                """
                INSERT INTO questions (
                    id, legacy_id, question_number, question_text,
                    option_a, option_b, option_c, option_d, correct_answer,
                    topic_code, topic_name, source_reference, source_active
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 1)
                ON CONFLICT(id) DO UPDATE SET
                    legacy_id = excluded.legacy_id,
                    question_number = excluded.question_number,
                    question_text = excluded.question_text,
                    option_a = excluded.option_a,
                    option_b = excluded.option_b,
                    option_c = excluded.option_c,
                    option_d = excluded.option_d,
                    correct_answer = excluded.correct_answer,
                    topic_code = excluded.topic_code,
                    topic_name = excluded.topic_name,
                    source_reference = excluded.source_reference,
                    source_active = 1
                """,
                questions,
            )
            database.executemany(
                "INSERT INTO employees (id, full_name) VALUES (?, ?)",
                employees,
            )
            database.executemany(
                """
                INSERT INTO quiz_attempts (
                    legacy_access_id, employee_name, topic_name,
                    started_at, finished_at, duration_seconds,
                    total_questions, answered_questions, correct_answers,
                    incorrect_answers, unanswered_questions, percentage, rating
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                ON CONFLICT DO UPDATE SET
                    employee_name = excluded.employee_name,
                    topic_name = excluded.topic_name,
                    started_at = excluded.started_at,
                    finished_at = excluded.finished_at,
                    duration_seconds = excluded.duration_seconds,
                    total_questions = excluded.total_questions,
                    answered_questions = excluded.answered_questions,
                    correct_answers = excluded.correct_answers,
                    incorrect_answers = excluded.incorrect_answers,
                    unanswered_questions = excluded.unanswered_questions,
                    percentage = excluded.percentage,
                    rating = excluded.rating
                """,
                legacy_attempts,
            )
            database.execute(
                """
                INSERT INTO sync_metadata (key, value) VALUES ('access_sha256', ?)
                ON CONFLICT(key) DO UPDATE SET value = excluded.value
                """,
                (source_hash,),
            )
            database.execute(
                """
                INSERT INTO sync_metadata (key, value) VALUES ('synced_at', ?)
                ON CONFLICT(key) DO UPDATE SET value = excluded.value
                """,
                (datetime.now().isoformat(timespec="seconds"),),
            )
            database.execute(
                """
                INSERT INTO business_topics (code, name, active, sort_order)
                SELECT topic_code, MIN(topic_name), 1, MIN(id)
                FROM questions
                WHERE topic_code <> ''
                GROUP BY topic_code
                ON CONFLICT(code) DO UPDATE SET active = 1
                """
            )
            database.execute(
                """
                INSERT OR IGNORE INTO question_topics (
                    business_code, name, active, sort_order
                )
                SELECT topic_code, topic_name, 1, MIN(id)
                FROM questions
                WHERE source_active = 1
                  AND topic_code <> '' AND topic_name <> ''
                GROUP BY topic_code, topic_name
                """
            )
            database.execute(
                """
                UPDATE questions
                SET subject_id = (
                    SELECT s.id FROM question_topics s
                    WHERE s.business_code = questions.topic_code
                      AND s.name = questions.topic_name
                    LIMIT 1
                )
                WHERE source_active = 1
                """
            )
            database.execute(
                """
                UPDATE questions
                SET topic_name = COALESCE((
                    SELECT t.name FROM business_topics t
                    WHERE t.code = questions.topic_code
                ), topic_name)
                WHERE source_active = 1
                """
            )
        return SyncResult(
            True,
            len(questions),
            len(employees),
            len(legacy_attempts),
            source_hash,
        )

    def _read_access(
        self,
    ) -> tuple[
        list[tuple[Any, ...]],
        list[tuple[int, str]],
        list[tuple[Any, ...]],
    ]:
        connection = None
        last_error: Exception | None = None
        for provider in self.ACCESS_PROVIDERS:
            try:
                connection = win32com.client.Dispatch("ADODB.Connection")
                connection.Open(
                    f"Provider={provider};Data Source={self.access_path};Mode=Read;"
                )
                break
            except com_error as exc:
                last_error = exc
                connection = None
        if connection is None:
            raise QuizDatabaseError(
                "Không mở được Access. Cần Microsoft Access Database Engine "
                "(ACE OLEDB 12.0 hoặc 16.0) cùng kiến trúc với Python."
            ) from last_error

        try:
            question_rows = self._ado_rows(
                connection,
                """
                SELECT ID2, ID, SoCauHoi, TenCauHoi,
                       A_DapAn, B_DapAn, C_DapAn, D_DapAn, DapAnDung,
                       ChuyenDe, ChuyenDe1, NguonThamKhao
                FROM DataKhaoSatKienThucNghiepVu
                ORDER BY ChuyenDe1, SoCauHoi, ID2
                """,
            )
            employee_rows = self._ado_rows(
                connection,
                "SELECT ID, HoVaTen FROM SysNhanVien ORDER BY HoVaTen",
            )
            attempt_rows = self._ado_rows(
                connection,
                """
                SELECT ID, Ngay, Gio, HoVaTen, LoaiOnTap, TenLoaiOnTap,
                       TimeStart, TimeFinish, ThoiGianDaTraLoi,
                       TongSoCauHoi, TongSoCauDaTraLoi,
                       TongSoCauTraLoiDung, TongSoCauTraLoiSai,
                       TrongSoCauChuaHoanThanh, TileHoanThanh, XepLoai
                FROM KetQuaOnTapKhaoSatKienThucNghiepVu
                ORDER BY ID
                """,
            )
        finally:
            try:
                connection.Close()
            except com_error:
                pass

        questions: list[tuple[Any, ...]] = []
        for row in question_rows:
            question_id = self._as_int(row["ID2"], len(questions) + 1)
            correct = self._text(row["DapAnDung"]).upper()
            if correct not in {"A", "B", "C", "D"}:
                continue
            questions.append(
                (
                    question_id,
                    row["ID"],
                    self._as_int(row["SoCauHoi"], question_id),
                    self._text(row["TenCauHoi"]),
                    self._text(row["A_DapAn"]),
                    self._text(row["B_DapAn"]),
                    self._text(row["C_DapAn"]),
                    self._text(row["D_DapAn"]),
                    correct,
                    self._text(row["ChuyenDe"]),
                    self._text(row["ChuyenDe1"]),
                    self._text(row["NguonThamKhao"]),
                )
            )
        employees = [
            (self._as_int(row["ID"], index), self._text(row["HoVaTen"]))
            for index, row in enumerate(employee_rows, start=1)
            if self._text(row["HoVaTen"])
        ]
        legacy_attempts = []
        for row in attempt_rows:
            started_at = self._combine_access_datetime(
                row["Ngay"],
                row["TimeStart"] or row["Gio"],
            )
            finished_at = self._combine_access_datetime(
                row["Ngay"],
                row["TimeFinish"] or row["Gio"],
            )
            duration = self._duration_seconds(row["ThoiGianDaTraLoi"])
            topic = " - ".join(
                part
                for part in (
                    self._text(row["LoaiOnTap"]),
                    self._text(row["TenLoaiOnTap"]),
                )
                if part
            )
            legacy_attempts.append(
                (
                    self._as_int(row["ID"], len(legacy_attempts) + 1),
                    self._text(row["HoVaTen"]),
                    topic,
                    started_at,
                    finished_at,
                    duration,
                    self._as_int(row["TongSoCauHoi"], 0),
                    self._as_int(row["TongSoCauDaTraLoi"], 0),
                    self._as_int(row["TongSoCauTraLoiDung"], 0),
                    self._as_int(row["TongSoCauTraLoiSai"], 0),
                    self._as_int(row["TrongSoCauChuaHoanThanh"], 0),
                    float(row["TileHoanThanh"] or 0),
                    self._text(row["XepLoai"]),
                )
            )
        return questions, employees, legacy_attempts

    @staticmethod
    def _ado_rows(connection: Any, sql: str) -> list[dict[str, Any]]:
        recordset = win32com.client.Dispatch("ADODB.Recordset")
        recordset.Open(sql, connection, 3, 1)
        try:
            names = [
                str(recordset.Fields(index).Name)
                for index in range(recordset.Fields.Count)
            ]
            if recordset.EOF:
                return []
            columns = recordset.GetRows()
            return [
                dict(zip(names, row, strict=True))
                for row in zip(*columns, strict=True)
            ]
        finally:
            recordset.Close()

    def topics(self) -> list[tuple[str, int]]:
        with self._database() as database:
            rows = database.execute(
                """
                SELECT s.name AS topic_name, COUNT(q.id) AS question_count
                FROM question_topics s
                JOIN questions q ON q.subject_id = s.id
                WHERE s.active = 1 AND q.source_active = 1
                GROUP BY s.id, s.name
                ORDER BY s.name COLLATE NOCASE
                """
            ).fetchall()
        return [(str(row["topic_name"]), int(row["question_count"])) for row in rows]

    def business_topics(self) -> list[tuple[str, str, int]]:
        with self._database() as database:
            rows = database.execute(
                """
                SELECT t.code, t.name, COUNT(q.id) AS question_count
                FROM business_topics t
                LEFT JOIN questions q
                  ON q.topic_code = t.code AND q.source_active = 1
                WHERE t.active = 1
                GROUP BY t.code, t.name, t.sort_order
                ORDER BY t.sort_order, t.name COLLATE NOCASE
                """
            ).fetchall()
        return [
            (str(row["code"]), str(row["name"]), int(row["question_count"]))
            for row in rows
        ]

    def save_business_topic(
        self,
        code: str,
        name: str,
        original_code: str | None = None,
    ) -> None:
        code = code.strip()
        name = name.strip()
        if not code or not name:
            raise QuizDatabaseError("Mã và tên nghiệp vụ không được để trống.")
        with self._database() as database:
            if original_code and original_code != code:
                if database.execute(
                    "SELECT 1 FROM business_topics WHERE code = ?",
                    (code,),
                ).fetchone():
                    raise QuizDatabaseError(f"Mã nghiệp vụ {code} đã tồn tại.")
                database.execute(
                    """
                    INSERT INTO business_topics (code, name, active, sort_order)
                    SELECT ?, ?, active, sort_order
                    FROM business_topics WHERE code = ?
                    """,
                    (code, name, original_code),
                )
                database.execute(
                    """
                    UPDATE question_topics
                    SET business_code = ?
                    WHERE business_code = ?
                    """,
                    (code, original_code),
                )
                database.execute(
                    """
                    UPDATE questions
                    SET topic_code = ?, topic_name = ?
                    WHERE topic_code = ?
                    """,
                    (code, name, original_code),
                )
                database.execute(
                    "DELETE FROM business_topics WHERE code = ?",
                    (original_code,),
                )
            else:
                database.execute(
                    """
                    INSERT INTO business_topics (code, name, active)
                    VALUES (?, ?, 1)
                    ON CONFLICT(code) DO UPDATE SET
                        name = excluded.name,
                        active = 1
                    """,
                    (code, name),
                )
                database.execute(
                    "UPDATE questions SET topic_name = ? WHERE topic_code = ?",
                    (name, code),
                )

    def delete_business_topic(self, code: str) -> None:
        with self._database() as database:
            database.execute(
                "UPDATE questions SET source_active = 0 WHERE topic_code = ?",
                (code,),
            )
            database.execute(
                "UPDATE business_topics SET active = 0 WHERE code = ?",
                (code,),
            )
            database.execute(
                "UPDATE question_topics SET active = 0 WHERE business_code = ?",
                (code,),
            )

    def question_topics(
        self,
        business_code: str | None = None,
    ) -> list[tuple[int, str, str, int]]:
        sql = """
            SELECT s.id, s.business_code, s.name, COUNT(q.id) AS question_count
            FROM question_topics s
            LEFT JOIN questions q
              ON q.subject_id = s.id AND q.source_active = 1
            WHERE s.active = 1
        """
        parameters: list[Any] = []
        if business_code:
            sql += " AND s.business_code = ?"
            parameters.append(business_code)
        sql += """
            GROUP BY s.id, s.business_code, s.name, s.sort_order
            ORDER BY s.sort_order, s.name COLLATE NOCASE
        """
        with self._database() as database:
            rows = database.execute(sql, parameters).fetchall()
        return [
            (
                int(row["id"]),
                str(row["business_code"]),
                str(row["name"]),
                int(row["question_count"]),
            )
            for row in rows
        ]

    def save_question_topic(
        self,
        business_code: str,
        name: str,
        subject_id: int | None = None,
    ) -> int:
        business_code = business_code.strip()
        name = name.strip()
        if not business_code or not name:
            raise QuizDatabaseError("Nghiệp vụ và tên chuyên đề là bắt buộc.")
        with self._database() as database:
            if subject_id is None:
                cursor = database.execute(
                    """
                    INSERT INTO question_topics (business_code, name, active)
                    VALUES (?, ?, 1)
                    ON CONFLICT(business_code, name)
                    DO UPDATE SET active = 1
                    RETURNING id
                    """,
                    (business_code, name),
                )
                return int(cursor.fetchone()[0])
            database.execute(
                """
                UPDATE question_topics
                SET business_code = ?, name = ?, active = 1
                WHERE id = ?
                """,
                (business_code, name, subject_id),
            )
            database.execute(
                """
                UPDATE questions
                SET topic_code = ?,
                    topic_name = (
                        SELECT name FROM business_topics WHERE code = ?
                    )
                WHERE subject_id = ?
                """,
                (business_code, business_code, subject_id),
            )
            return subject_id

    def delete_question_topic(self, subject_id: int) -> None:
        with self._database() as database:
            database.execute(
                "UPDATE questions SET source_active = 0 WHERE subject_id = ?",
                (subject_id,),
            )
            database.execute(
                "UPDATE question_topics SET active = 0 WHERE id = ?",
                (subject_id,),
            )

    def delete_questions_bulk(self, business_code: str | None = None) -> int:
        with self._database() as database:
            if business_code:
                cursor = database.execute(
                    """
                    UPDATE questions SET source_active = 0
                    WHERE source_active = 1 AND topic_code = ?
                    """,
                    (business_code,),
                )
            else:
                cursor = database.execute(
                    "UPDATE questions SET source_active = 0 WHERE source_active = 1"
                )
        return max(0, int(cursor.rowcount))

    def search_questions(
        self,
        search_text: str = "",
        topic_code: str | None = None,
        limit: int = 500,
    ) -> list[Question]:
        sql = """
            SELECT q.*, COALESCE(s.name, '') AS subject_name
            FROM questions q
            LEFT JOIN question_topics s ON s.id = q.subject_id
            WHERE q.source_active = 1
        """
        parameters: list[Any] = []
        if topic_code:
            sql += " AND q.topic_code = ?"
            parameters.append(topic_code)
        if search_text.strip():
            sql += (
                " AND (q.question_text LIKE ? OR q.option_a LIKE ? OR q.option_b LIKE ? "
                "OR q.option_c LIKE ? OR q.option_d LIKE ? OR q.source_reference LIKE ? "
                "OR s.name LIKE ?)"
            )
            pattern = f"%{search_text.strip()}%"
            parameters.extend([pattern] * 7)
        sql += " ORDER BY q.topic_name COLLATE NOCASE, s.name, q.question_number, q.id LIMIT ?"
        parameters.append(limit)
        with self._database() as database:
            rows = database.execute(sql, parameters).fetchall()
        return self._questions_from_rows(rows)

    def question_by_id(self, question_id: int) -> Question | None:
        with self._database() as database:
            row = database.execute(
                """
                SELECT q.*, COALESCE(s.name, '') AS subject_name
                FROM questions q
                LEFT JOIN question_topics s ON s.id = q.subject_id
                WHERE q.id = ?
                """,
                (question_id,),
            ).fetchone()
        return self._questions_from_rows([row])[0] if row else None

    def save_question(
        self,
        *,
        question_id: int | None,
        topic_code: str,
        text: str,
        options: dict[str, str],
        correct_answer: str,
        subject_id: int | None = None,
        locked_answers: str = "",
        source_reference: str = "",
        question_number: int | None = None,
    ) -> int:
        topic_code = topic_code.strip()
        text = text.strip()
        correct_answer = correct_answer.strip().upper()
        if not topic_code or not text:
            raise QuizDatabaseError("Nghiệp vụ và nội dung câu hỏi là bắt buộc.")
        if correct_answer not in {"A", "B", "C", "D"}:
            raise QuizDatabaseError("Đáp án đúng phải là A, B, C hoặc D.")
        normalized = {
            letter: str(options.get(letter, "")).strip()
            for letter in "ABCD"
        }
        if not normalized[correct_answer]:
            raise QuizDatabaseError("Đáp án đúng không được để trống.")
        locked_answers = "".join(
            letter for letter in "ABCD" if letter in locked_answers.upper()
        )
        with self._database() as database:
            topic = database.execute(
                "SELECT name FROM business_topics WHERE code = ? AND active = 1",
                (topic_code,),
            ).fetchone()
            if not topic:
                raise QuizDatabaseError("Nghiệp vụ đã chọn không tồn tại.")
            topic_name = str(topic["name"])
            if subject_id is not None and not database.execute(
                """
                SELECT 1 FROM question_topics
                WHERE id = ? AND business_code = ? AND active = 1
                """,
                (subject_id, topic_code),
            ).fetchone():
                raise QuizDatabaseError(
                    "Chuyên đề không thuộc nghiệp vụ đã chọn."
                )
            exists = (
                question_id is not None
                and database.execute(
                    "SELECT 1 FROM questions WHERE id = ?",
                    (question_id,),
                ).fetchone() is not None
            )
            if not exists:
                question_id = int(
                    database.execute(
                        "SELECT COALESCE(MAX(id), 0) + 1 FROM questions"
                    ).fetchone()[0]
                ) if question_id is None else int(question_id)
                if question_number is None:
                    question_number = int(
                        database.execute(
                            """
                            SELECT COALESCE(MAX(question_number), 0) + 1
                            FROM questions WHERE topic_code = ?
                            """,
                            (topic_code,),
                        ).fetchone()[0]
                    )
                database.execute(
                    """
                    INSERT INTO questions (
                        id, legacy_id, question_number, question_text,
                        option_a, option_b, option_c, option_d, correct_answer,
                        topic_code, topic_name, subject_id, locked_answers,
                        source_reference, source_active
                    ) VALUES (?, NULL, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 1)
                    """,
                    (
                        question_id, question_number, text,
                        normalized["A"], normalized["B"],
                        normalized["C"], normalized["D"],
                        correct_answer, topic_code, topic_name,
                        subject_id, locked_answers,
                        source_reference.strip(),
                    ),
                )
            else:
                database.execute(
                    """
                    UPDATE questions SET
                        question_text = ?, option_a = ?, option_b = ?,
                        option_c = ?, option_d = ?, correct_answer = ?,
                        topic_code = ?, topic_name = ?, subject_id = ?,
                        locked_answers = ?, source_reference = ?,
                        source_active = 1
                    WHERE id = ?
                    """,
                    (
                        text, normalized["A"], normalized["B"],
                        normalized["C"], normalized["D"], correct_answer,
                        topic_code, topic_name, subject_id, locked_answers,
                        source_reference.strip(),
                        question_id,
                    ),
                )
        return int(question_id)

    def delete_question(self, question_id: int) -> None:
        with self._database() as database:
            database.execute(
                "UPDATE questions SET source_active = 0 WHERE id = ?",
                (question_id,),
            )

    def load_random_exam_setting(self) -> dict[str, int]:
        with self._database() as database:
            row = database.execute(
                "SELECT value FROM quiz_settings WHERE key = 'random_exam_quotas'"
            ).fetchone()
        if not row:
            return {}
        try:
            return {
                str(code): max(0, int(count))
                for code, count in json.loads(str(row["value"])).items()
            }
        except (TypeError, ValueError, json.JSONDecodeError):
            return {}

    def save_random_exam_setting(self, quotas: dict[str, int]) -> None:
        value = json.dumps(quotas, ensure_ascii=False, sort_keys=True)
        with self._database() as database:
            database.execute(
                """
                INSERT INTO quiz_settings (key, value)
                VALUES ('random_exam_quotas', ?)
                ON CONFLICT(key) DO UPDATE SET value = excluded.value
                """,
                (value,),
            )

    def load_random_subject_setting(self) -> dict[int, int]:
        with self._database() as database:
            row = database.execute(
                """
                SELECT value FROM quiz_settings
                WHERE key = 'random_exam_subject_quotas'
                """
            ).fetchone()
        if not row:
            return {}
        try:
            return {
                int(subject_id): max(0, int(count))
                for subject_id, count in json.loads(str(row["value"])).items()
            }
        except (TypeError, ValueError, json.JSONDecodeError):
            return {}

    def save_random_subject_setting(self, quotas: dict[int, int]) -> None:
        value = json.dumps(quotas, ensure_ascii=False, sort_keys=True)
        with self._database() as database:
            database.execute(
                """
                INSERT INTO quiz_settings (key, value)
                VALUES ('random_exam_subject_quotas', ?)
                ON CONFLICT(key) DO UPDATE SET value = excluded.value
                """,
                (value,),
            )

    def topic_code_counts(self) -> dict[str, int]:
        with self._database() as database:
            rows = database.execute(
                """
                SELECT topic_code, COUNT(*) AS question_count
                FROM questions
                WHERE source_active = 1
                GROUP BY topic_code
                """
            ).fetchall()
        return {
            str(row["topic_code"]): int(row["question_count"])
            for row in rows
        }

    def employees(self) -> list[str]:
        with self._database() as database:
            rows = database.execute(
                "SELECT full_name FROM employees ORDER BY full_name COLLATE NOCASE"
            ).fetchall()
        return [str(row["full_name"]) for row in rows]

    def questions(
        self,
        topic_name: str | None,
        limit: int,
        randomize: bool = True,
    ) -> list[Question]:
        sql = """
            SELECT q.*, COALESCE(s.name, '') AS subject_name
            FROM questions q
            LEFT JOIN question_topics s ON s.id = q.subject_id
            WHERE q.source_active = 1
        """
        parameters: list[Any] = []
        if topic_name:
            sql += " AND (s.name = ? OR q.topic_name = ?)"
            parameters.append(topic_name)
            parameters.append(topic_name)
        sql += " ORDER BY "
        sql += "RANDOM()" if randomize else "q.question_number, q.id"
        if limit > 0:
            sql += " LIMIT ?"
            parameters.append(limit)
        with self._database() as database:
            rows = database.execute(sql, parameters).fetchall()
        return self._questions_from_rows(rows)

    def questions_by_quotas(self, quotas: dict[str, int]) -> list[Question]:
        selected: list[Question] = []
        with self._database() as database:
            for topic_code, limit in quotas.items():
                if limit <= 0:
                    continue
                rows = database.execute(
                    """
                    SELECT q.*, COALESCE(s.name, '') AS subject_name
                    FROM questions q
                    LEFT JOIN question_topics s ON s.id = q.subject_id
                    WHERE q.source_active = 1 AND q.topic_code = ?
                    ORDER BY RANDOM()
                    LIMIT ?
                    """,
                    (topic_code, limit),
                ).fetchall()
                selected.extend(self._questions_from_rows(rows))
        import random

        random.shuffle(selected)
        return selected

    def questions_by_subject_quotas(
        self,
        quotas: dict[int, int],
    ) -> list[Question]:
        selected: list[Question] = []
        with self._database() as database:
            for subject_id, limit in quotas.items():
                if limit <= 0:
                    continue
                rows = database.execute(
                    """
                    SELECT q.*, COALESCE(s.name, '') AS subject_name
                    FROM questions q
                    LEFT JOIN question_topics s ON s.id = q.subject_id
                    WHERE q.source_active = 1 AND q.subject_id = ?
                    ORDER BY RANDOM()
                    LIMIT ?
                    """,
                    (subject_id, limit),
                ).fetchall()
                selected.extend(self._questions_from_rows(rows))
        import random

        random.shuffle(selected)
        return selected

    @staticmethod
    def _questions_from_rows(rows: Any) -> list[Question]:
        questions = []
        for row in rows:
            options = {
                letter: str(row[f"option_{letter.lower()}"] or "")
                for letter in "ABCD"
                if str(row[f"option_{letter.lower()}"] or "").strip()
            }
            questions.append(
                Question(
                    id=int(row["id"]),
                    legacy_id=row["legacy_id"],
                    number=int(row["question_number"]),
                    text=str(row["question_text"]),
                    options=options,
                    correct_answer=str(row["correct_answer"]),
                    topic_code=str(row["topic_code"]),
                    topic_name=str(row["topic_name"]),
                    source_reference=str(row["source_reference"]),
                    subject_id=(
                        int(row["subject_id"])
                        if row["subject_id"] is not None
                        else None
                    ),
                    subject_name=(
                        str(row["subject_name"])
                        if "subject_name" in row.keys()
                        else ""
                    ),
                    locked_answers=str(row["locked_answers"] or ""),
                )
            )
        return questions

    def save_attempt(self, session: QuizSession, result: QuizResult) -> int:
        finished_at = session.finished_at or datetime.now()
        with self._database() as database:
            cursor = database.execute(
                """
                INSERT INTO quiz_attempts (
                    employee_name, topic_name, started_at, finished_at,
                    duration_seconds, total_questions, answered_questions,
                    correct_answers, incorrect_answers, unanswered_questions,
                    percentage, rating
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    session.employee_name,
                    session.topic_name,
                    session.started_at.isoformat(timespec="seconds"),
                    finished_at.isoformat(timespec="seconds"),
                    result.duration_seconds,
                    result.total,
                    result.answered,
                    result.correct,
                    result.incorrect,
                    result.unanswered,
                    result.percentage,
                    result.rating,
                ),
            )
            attempt_id = int(cursor.lastrowid)
            database.executemany(
                """
                INSERT INTO quiz_answers (
                    attempt_id, question_id, display_order, selected_answer,
                    correct_answer, is_correct, checked_before_finish
                ) VALUES (?, ?, ?, ?, ?, ?, ?)
                """,
                [
                    (
                        attempt_id,
                        question.id,
                        index,
                        session.answers.get(question.id, ""),
                        question.correct_answer,
                        int(
                            session.answers.get(question.id)
                            == question.correct_answer
                        ),
                        int(question.id in session.checked_questions),
                    )
                    for index, question in enumerate(session.questions, start=1)
                ],
            )
        return attempt_id

    def question_count(self) -> int:
        with self._database() as database:
            return int(
                database.execute(
                    "SELECT COUNT(*) FROM questions WHERE source_active = 1"
                ).fetchone()[0]
            )

    def employee_count(self) -> int:
        with self._database() as database:
            return int(database.execute("SELECT COUNT(*) FROM employees").fetchone()[0])

    def legacy_attempt_count(self) -> int:
        with self._database() as database:
            return int(
                database.execute(
                    "SELECT COUNT(*) FROM quiz_attempts "
                    "WHERE legacy_access_id IS NOT NULL"
                ).fetchone()[0]
            )

    def _metadata(self, key: str) -> str:
        with self._database() as database:
            row = database.execute(
                "SELECT value FROM sync_metadata WHERE key = ?",
                (key,),
            ).fetchone()
        return str(row["value"]) if row else ""

    @staticmethod
    def _sha256(path: Path) -> str:
        digest = hashlib.sha256()
        with path.open("rb") as source:
            for block in iter(lambda: source.read(1024 * 1024), b""):
                digest.update(block)
        return digest.hexdigest()

    @staticmethod
    def _text(value: Any) -> str:
        return "" if value is None else str(value).strip()

    @staticmethod
    def _as_int(value: Any, default: int) -> int:
        try:
            return int(float(value))
        except (TypeError, ValueError):
            return default

    @staticmethod
    def _combine_access_datetime(date_value: Any, time_value: Any) -> str:
        date_part = (
            date_value.date()
            if hasattr(date_value, "date")
            else datetime.now().date()
        )
        time_part = (
            time_value.time()
            if hasattr(time_value, "time")
            else datetime.min.time()
        )
        return datetime.combine(date_part, time_part).isoformat(timespec="seconds")

    @staticmethod
    def _duration_seconds(value: Any) -> int:
        try:
            hours, minutes, seconds = (int(part) for part in str(value).split(":"))
            return hours * 3600 + minutes * 60 + seconds
        except (TypeError, ValueError):
            return 0
