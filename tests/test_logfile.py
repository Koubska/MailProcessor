import logging
from pathlib import Path

from mailprocessor.logfile import (
    BACKUP_COUNT,
    MAX_BYTES,
    _AppLogFileHandler,
    attach_log_file,
    log_file_path,
)


def _file_handlers() -> list[_AppLogFileHandler]:
    return [h for h in logging.getLogger("mailprocessor").handlers if isinstance(h, _AppLogFileHandler)]


def test_attach_log_file_writes_records_next_to_the_config(tmp_path: Path) -> None:
    package_logger = logging.getLogger("mailprocessor")
    previous_level = package_logger.level
    package_logger.setLevel(logging.INFO)
    try:
        path = attach_log_file(tmp_path)
        logging.getLogger("mailprocessor.test").info("Processed a.eml")
    finally:
        package_logger.setLevel(previous_level)

    assert path == log_file_path(tmp_path) == tmp_path / "logs" / "mailprocessor.log"
    assert "INFO    mailprocessor.test: Processed a.eml" in path.read_text(encoding="utf-8")
    handler = _file_handlers()[0]
    assert (handler.maxBytes, handler.backupCount) == (MAX_BYTES, BACKUP_COUNT)


def test_attaching_again_replaces_the_handler(tmp_path: Path) -> None:
    attach_log_file(tmp_path / "first")
    attach_log_file(tmp_path / "second")

    handlers = _file_handlers()
    assert len(handlers) == 1
    assert Path(handlers[0].baseFilename) == tmp_path / "second" / "logs" / "mailprocessor.log"


def test_unwritable_log_location_does_not_fail(tmp_path: Path) -> None:
    blocker = tmp_path / "config_dir"
    blocker.mkdir()
    (blocker / "logs").write_text("a file where the folder should be", encoding="utf-8")

    assert attach_log_file(blocker) is None
    assert _file_handlers() == []
