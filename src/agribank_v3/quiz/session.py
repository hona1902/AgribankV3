from __future__ import annotations

from datetime import datetime

from agribank_v3.quiz.models import Question, QuizResult


class QuizSession:
    def __init__(
        self,
        questions: list[Question],
        employee_name: str,
        topic_name: str,
        time_limit_minutes: int = 0,
    ) -> None:
        if not questions:
            raise ValueError("Quiz requires at least one question.")
        self.questions = questions
        self.employee_name = employee_name.strip()
        self.topic_name = topic_name
        self.time_limit_minutes = max(0, time_limit_minutes)
        self.answers: dict[int, str] = {}
        self.checked_questions: set[int] = set()
        self.current_index = 0
        self.started_at = datetime.now()
        self.finished_at: datetime | None = None

    @property
    def current_question(self) -> Question:
        return self.questions[self.current_index]

    def answer(self, answer: str) -> None:
        normalized = answer.upper()
        if normalized not in self.current_question.options:
            raise ValueError(f"Invalid answer: {answer}")
        self.answers[self.current_question.id] = normalized

    def selected_answer(self, question: Question | None = None) -> str:
        target = question or self.current_question
        return self.answers.get(target.id, "")

    def check_current(self) -> bool | None:
        question = self.current_question
        selected = self.selected_answer(question)
        if not selected:
            return None
        self.checked_questions.add(question.id)
        return selected == question.correct_answer

    def finish(self) -> QuizResult:
        if self.finished_at is None:
            self.finished_at = datetime.now()
        correct = sum(
            self.answers.get(question.id) == question.correct_answer
            for question in self.questions
        )
        total = len(self.questions)
        answered = len(self.answers)
        percentage = round(correct * 100 / total, 2)
        return QuizResult(
            total=total,
            answered=answered,
            correct=correct,
            incorrect=answered - correct,
            unanswered=total - answered,
            percentage=percentage,
            rating=self.rating_for(percentage),
            duration_seconds=max(
                0,
                int((self.finished_at - self.started_at).total_seconds()),
            ),
        )

    @staticmethod
    def rating_for(percentage: float) -> str:
        if percentage < 50:
            return "Yếu"
        if percentage < 75:
            return "Trung bình"
        if percentage < 90:
            return "Khá"
        return "Xuất sắc"
