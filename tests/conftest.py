import logging

import pytest

from mailprocessor.logfile import _AppLogFileHandler


@pytest.fixture(autouse=True)
def _detach_log_file_handler():
    """A log file attached by one test (CLI/GUI runs) must not keep receiving records from later tests."""
    yield
    package_logger = logging.getLogger("mailprocessor")
    for handler in list(package_logger.handlers):
        if isinstance(handler, _AppLogFileHandler):
            package_logger.removeHandler(handler)
            handler.close()
