"""Log lines of a run reach the window through a queue."""

import logging
import queue

from mailprocessor.gui.log_handler import QueueLogHandler


def test_queue_log_handler_forwards_formatted_lines_respecting_level() -> None:
    sink: queue.Queue = queue.Queue()
    logger = logging.getLogger("mailprocessor.test_gui_handler")
    handler = QueueLogHandler(sink)
    logger.addHandler(handler)
    logger.setLevel(logging.INFO)
    try:
        logger.debug("hidden")
        logger.warning("Failed a.eml: Required fields missing: Name")
    finally:
        logger.removeHandler(handler)

    kind, line = sink.get_nowait()
    assert kind == "log"
    assert line.endswith("WARNING Failed a.eml: Required fields missing: Name")
    assert sink.empty()
