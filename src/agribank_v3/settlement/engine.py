from __future__ import annotations

from typing import Protocol

from agribank_v3.settlement.models import SettlementRequest, SettlementResult
from agribank_v3.settlement.registry import SETTLEMENT_SPECS


class SettlementError(RuntimeError):
    def __init__(self, message: str, code: str = "general") -> None:
        super().__init__(message)
        self.code = code


class SettlementProcessor(Protocol):
    def execute(self, request: SettlementRequest) -> SettlementResult: ...


class SettlementEngine:
    """Dispatch settlement requests to independently testable processors."""

    def __init__(self) -> None:
        self._processors: dict[str, SettlementProcessor] = {}

    def register(self, family: str, processor: SettlementProcessor) -> None:
        normalized = family.strip().casefold()
        if not normalized:
            raise ValueError("Processor family must not be empty.")
        self._processors[normalized] = processor

    def supports(self, spec_key: str) -> bool:
        spec = SETTLEMENT_SPECS.get(spec_key)
        return (
            spec is not None
            and spec.processor_family.casefold() in self._processors
        )

    def execute(self, request: SettlementRequest) -> SettlementResult:
        self._validate_request(request)
        family = request.spec.processor_family.casefold()
        processor = self._processors.get(family)
        if processor is None:
            raise SettlementError(
                f"Mẫu {request.spec.report_code}/QT chưa có processor Python.",
                code="processor_not_migrated",
            )
        return processor.execute(request)

    @staticmethod
    def _validate_request(request: SettlementRequest) -> None:
        registered = SETTLEMENT_SPECS.get(request.spec.key)
        if registered != request.spec:
            raise SettlementError(
                f"Mẫu quyết toán không hợp lệ: {request.spec.key}",
                code="unknown_spec",
            )
        if not request.profile.branch_code.strip():
            raise SettlementError(
                "Chưa có mã chi nhánh trong Cài đặt.",
                code="missing_branch_code",
            )
        if (
            request.spec.source_mode.value == "multiple_files"
            and not request.source_paths
        ):
            raise SettlementError(
                f"Mẫu {request.spec.report_code}/QT cần ít nhất một file nguồn.",
                code="missing_source_files",
            )
