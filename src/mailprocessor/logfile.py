"""Local log file with rotation, so users can send it along with a problem report.

The file receives the same records as the console or GUI, so the logging rules apply unchanged:
no mail bodies, extracted values or credentials.
"""

from __future__ import annotations

import logging
from logging.handlers import RotatingFileHandler
from pathlib import Path

LOG_DIR_NAME = "logs"
LOG_FILE_NAME = "mailprocessor.log"
MAX_BYTES = 1_000_000
BACKUP_COUNT = 3  # mailprocessor.log.1 … .3; at most ~4 MB in total
FILE_FORMAT = "%(asctime)s %(levelname)-7s %(name)s: %(message)s"


class _AppLogFileHandler(RotatingFileHandler):
    """Marker type, so attaching again replaces the handler instead of adding a second one."""


def log_file_path(config_dir: Path) -> Path:
    """The log file sits next to the config file, like the other runtime folders (data, out)."""
    return config_dir / LOG_DIR_NAME / LOG_FILE_NAME


def attach_log_file(config_dir: Path) -> Path | None:
    """Write `mailprocessor` log records to the log file. Returns its path, or None if it cannot be written.

    A log file that cannot be created (e.g. a read-only folder) never stops the app.
    """
    package_logger = logging.getLogger("mailprocessor")
    for handler in list(package_logger.handlers):
        if isinstance(handler, _AppLogFileHandler):
            package_logger.removeHandler(handler)
            handler.close()
    path = log_file_path(config_dir)
    try:
        path.parent.mkdir(parents=True, exist_ok=True)
        handler = _AppLogFileHandler(path, maxBytes=MAX_BYTES, backupCount=BACKUP_COUNT, encoding="utf-8")
    except OSError:
        return None
    handler.setFormatter(logging.Formatter(FILE_FORMAT))
    package_logger.addHandler(handler)
    return path
