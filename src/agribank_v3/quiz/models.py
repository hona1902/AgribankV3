from __future__ import annotations

from dataclasses import dataclass


@dataclass(frozen=True, slots=True)
class Question:
    id: int
    legacy_id: float | None
    number: int
    text: str
    options: dict[str, str]
    correct_answer: str
    topic_code: str
    topic_name: str
    source_reference: str
    subject_id: int | None = None
    subject_name: str = ""
    locked_answers: str = ""


@dataclass(frozen=True, slots=True)
class QuizResult:
    total: int
    answered: int
    correct: int
    incorrect: int
    unanswered: int
    percentage: float
    rating: str
    duration_seconds: int
