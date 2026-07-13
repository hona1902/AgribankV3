from __future__ import annotations

from collections.abc import Callable
import inspect
from typing import Any

from PySide6.QtCore import QObject, QThread, Signal, Slot


class FunctionWorker(QObject):
    finished = Signal(object)
    failed = Signal(object)
    progress = Signal(str)

    def __init__(self, function: Callable[[], Any]) -> None:
        super().__init__()
        self._function = function

    @Slot()
    def run(self) -> None:
        try:
            if len(inspect.signature(self._function).parameters) == 1:
                self.finished.emit(self._function(self.progress.emit))
            else:
                self.finished.emit(self._function())
        except Exception as exc:  # pragma: no cover - delivered to UI thread.
            self.failed.emit(exc)


class CallbackBridge(QObject):
    def __init__(
        self,
        on_finished: Callable[[Any], None],
        on_failed: Callable[[Exception], None],
        on_progress: Callable[[str], None] | None,
    ) -> None:
        super().__init__()
        self._on_finished = on_finished
        self._on_failed = on_failed
        self._on_progress = on_progress

    @Slot(object)
    def finished(self, payload: object) -> None:
        self._on_finished(payload)

    @Slot(object)
    def failed(self, exc: object) -> None:
        if isinstance(exc, Exception):
            self._on_failed(exc)
        else:
            self._on_failed(RuntimeError(str(exc)))

    @Slot(str)
    def progress(self, message: str) -> None:
        if self._on_progress is not None:
            self._on_progress(message)


def run_in_thread(
    parent: QObject,
    function: Callable[[], Any],
    on_finished: Callable[[Any], None],
    on_failed: Callable[[Exception], None],
    on_progress: Callable[[str], None] | None = None,
) -> QThread:
    thread = QThread(parent)
    worker = FunctionWorker(function)
    bridge = CallbackBridge(on_finished, on_failed, on_progress)
    bridge.setParent(parent)
    thread._agribank_worker = worker  # Keep the Python wrapper alive while QThread runs.
    thread._agribank_bridge = bridge
    worker.moveToThread(thread)
    thread.started.connect(worker.run)
    worker.progress.connect(bridge.progress)
    worker.finished.connect(thread.quit)
    worker.failed.connect(thread.quit)
    worker.finished.connect(bridge.finished)
    worker.failed.connect(bridge.failed)
    worker.finished.connect(worker.deleteLater)
    worker.failed.connect(worker.deleteLater)
    thread.finished.connect(lambda: setattr(thread, "_agribank_worker", None))
    thread.finished.connect(lambda: setattr(thread, "_agribank_bridge", None))
    thread.finished.connect(thread.deleteLater)
    thread.start()
    return thread
