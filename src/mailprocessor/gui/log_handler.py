"""Log lines of a run, passed from the worker thread to the window."""

from __future__ import annotations

import logging
import queue

LOG_FORMAT = "%(asctime)s %(levelname)-7s %(message)s"


class QueueLogHandler(logging.Handler):
    """Forwards formatted log lines from the worker thread to the UI thread as ("log", text) items."""

    def __init__(self, sink: queue.Queue) -> None:
        super().__init__()
        self.sink = sink
        self.setFormatter(logging.Formatter(LOG_FORMAT, datefmt="%H:%M:%S"))

    def emit(self, record: logging.LogRecord) -> None:
        try:
            self.sink.put(("log", self.format(record)))
        except Exception:
            self.handleError(record)
