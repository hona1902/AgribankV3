from __future__ import annotations

from collections import OrderedDict
from collections.abc import Callable, Hashable
from dataclasses import dataclass
import logging
from time import perf_counter
import threading
from typing import Any

from PySide6.QtCore import QObject

from agribank_v3.ui.workers import run_in_thread


LOGGER = logging.getLogger(__name__)
MISSING = object()


@dataclass(frozen=True, slots=True)
class QueryEnvelope:
    generation: int
    name: str
    payload: object
    elapsed_ms: float
    thread_id: int
    row_count: int


class LruQueryCache:
    def __init__(self, max_entries: int = 32) -> None:
        self.max_entries = max(1, int(max_entries or 32))
        self._items: OrderedDict[Hashable, object] = OrderedDict()

    def get(self, key: Hashable, default: object = MISSING) -> object:
        if key not in self._items:
            return default
        value = self._items.pop(key)
        self._items[key] = value
        return value

    def set(self, key: Hashable, value: object) -> None:
        if key in self._items:
            self._items.pop(key)
        self._items[key] = value
        while len(self._items) > self.max_entries:
            self._items.popitem(last=False)

    def clear(self) -> None:
        self._items.clear()

    def __len__(self) -> int:
        return len(self._items)


class AsyncQueryController(QObject):
    def __init__(self, parent: QObject | None = None, *, max_cache_entries: int = 32) -> None:
        super().__init__(parent)
        self.generation = 0
        self.latest_applied_generation = 0
        self.last_worker_thread_id = 0
        self.last_ui_thread_id = threading.get_ident()
        self.last_cache_hit = False
        self.stale_result_count = 0
        self.cache = LruQueryCache(max_cache_entries)
        self._threads = []
        self._loading_generation = 0

    def run(
        self,
        name: str,
        function: Callable[[], Any],
        on_success: Callable[[Any], None],
        on_error: Callable[[Exception], None] | None = None,
        *,
        cache_key: Hashable | None = None,
        use_cache: bool = True,
        state_callback: Callable[[str, str], None] | None = None,
        log_context: dict[str, object] | None = None,
    ) -> int:
        self.generation += 1
        generation = self.generation
        self.last_cache_hit = False
        if cache_key is not None and use_cache:
            cached = self.cache.get(cache_key, MISSING)
            if cached is not MISSING:
                self.last_cache_hit = True
                self.latest_applied_generation = generation
                self.last_ui_thread_id = threading.get_ident()
                LOGGER.debug("Customer query cache hit: %s", name)
                if state_callback is not None:
                    state_callback("ready", "")
                on_success(cached)
                return generation
        if state_callback is not None:
            state_callback("loading", "")
        self._loading_generation = generation
        context_text = _format_log_context(log_context)
        LOGGER.debug("Customer query requested: %s generation=%s%s", name, generation, context_text)

        def task() -> QueryEnvelope:
            started = perf_counter()
            LOGGER.debug(
                "Customer query worker start: %s generation=%s thread=%s%s",
                name,
                generation,
                threading.get_ident(),
                context_text,
            )
            payload = function()
            elapsed_ms = (perf_counter() - started) * 1000
            return QueryEnvelope(
                generation=generation,
                name=name,
                payload=payload,
                elapsed_ms=elapsed_ms,
                thread_id=threading.get_ident(),
                row_count=_row_count(payload),
            )

        def done(envelope: QueryEnvelope) -> None:
            self.last_ui_thread_id = threading.get_ident()
            if envelope.generation != self.generation:
                self.stale_result_count += 1
                LOGGER.debug(
                    "Customer query stale result skipped: %s generation=%s current_generation=%s",
                    envelope.name,
                    envelope.generation,
                    self.generation,
                )
                if self._loading_generation == envelope.generation:
                    self._loading_generation = 0
                    if state_callback is not None:
                        state_callback("ready", "")
                return
            self.latest_applied_generation = envelope.generation
            self.last_worker_thread_id = envelope.thread_id
            if cache_key is not None and use_cache:
                self.cache.set(cache_key, envelope.payload)
            if envelope.elapsed_ms > 2000:
                LOGGER.error(
                    "Customer query too slow: %s generation=%s %.1f ms rows=%s%s",
                    envelope.name,
                    envelope.generation,
                    envelope.elapsed_ms,
                    envelope.row_count,
                    context_text,
                )
            elif envelope.elapsed_ms > 500:
                LOGGER.warning(
                    "Customer query slow: %s generation=%s %.1f ms rows=%s%s",
                    envelope.name,
                    envelope.generation,
                    envelope.elapsed_ms,
                    envelope.row_count,
                    context_text,
                )
            else:
                LOGGER.debug(
                    "Customer query worker finish: %s generation=%s %.1f ms rows=%s%s",
                    envelope.name,
                    envelope.generation,
                    envelope.elapsed_ms,
                    envelope.row_count,
                    context_text,
                )
            if state_callback is not None:
                state_callback("ready", "")
            self._loading_generation = 0
            on_success(envelope.payload)

        def failed(exc: Exception) -> None:
            self.last_ui_thread_id = threading.get_ident()
            if generation != self.generation:
                self.stale_result_count += 1
                LOGGER.debug(
                    "Customer query stale error skipped: %s generation=%s current_generation=%s",
                    name,
                    generation,
                    self.generation,
                )
                if self._loading_generation == generation:
                    self._loading_generation = 0
                    if state_callback is not None:
                        state_callback("ready", "")
                return
            LOGGER.error("Customer query failed: %s", name, exc_info=(type(exc), exc, exc.__traceback__))
            if state_callback is not None:
                state_callback("error", str(exc))
            self._loading_generation = 0
            if on_error is not None:
                on_error(exc)

        thread = run_in_thread(self, task, done, failed)
        self._threads.append(thread)
        thread.finished.connect(lambda item=thread: self._forget_thread(item))
        return generation

    def invalidate_cache(self) -> None:
        self.cache.clear()

    def cancel_pending(self) -> None:
        self.generation += 1
        LOGGER.debug("Customer query generation cancelled: generation=%s", self.generation)

    def wait_for_idle(self, timeout_ms: int = 5000) -> None:
        for thread in list(self._threads):
            if thread.isRunning():
                thread.wait(max(0, int(timeout_ms or 0)))
            if not thread.isRunning():
                self._forget_thread(thread)

    def _forget_thread(self, thread) -> None:
        if thread in self._threads:
            self._threads.remove(thread)


def _row_count(payload: object) -> int:
    if hasattr(payload, "rows"):
        try:
            return len(getattr(payload, "rows"))
        except TypeError:
            return 0
    if isinstance(payload, dict):
        total = 0
        for value in payload.values():
            if isinstance(value, (list, tuple)):
                total += len(value)
        return total
    if isinstance(payload, (list, tuple)):
        return len(payload)
    return 0


def _format_log_context(context: dict[str, object] | None) -> str:
    if not context:
        return ""
    parts = []
    for key in sorted(context):
        value = context[key]
        if isinstance(value, (str, int, float, bool)) or value is None:
            parts.append(f"{key}={value}")
        elif isinstance(value, (list, tuple, set, frozenset)):
            parts.append(f"{key}={','.join(str(item) for item in value)}")
        else:
            parts.append(f"{key}={type(value).__name__}")
    return " " + " ".join(parts)
