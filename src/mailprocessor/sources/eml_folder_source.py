"""Local .eml folder source."""

from __future__ import annotations

import logging
from collections.abc import Callable, Iterator
from pathlib import Path

from mailprocessor.errors import MailFolderNotFoundError
from mailprocessor.models import MailReadError, NormalizedMail
from mailprocessor.sources.email_content import parse_message_bytes

logger = logging.getLogger(__name__)


def parse_eml_file(path: Path, source_location: str) -> NormalizedMail:
    return parse_message_bytes(path.read_bytes(), "eml", source_location, origin=path.name)


def matching_files(folder: Path, glob_pattern: str = "*.eml") -> list[Path]:
    """The files matching the pattern, ignoring case and sorted the same way on every platform.

    Windows matches "*.eml" case-insensitively anyway; without this, macOS and Linux would skip "Anfrage.EML".
    """
    paths = folder.glob(glob_pattern, case_sensitive=False)
    return sorted((path for path in paths if path.is_file()), key=lambda path: (path.name.casefold(), path.name))


def iter_eml_messages(
    folder: Path, glob_pattern: str = "*.eml", on_total: Callable[[int], None] | None = None
) -> Iterator[NormalizedMail | MailReadError]:
    """Yield every matching file; `on_total` receives the number of files before the first one is read."""
    if not folder.is_dir():
        raise MailFolderNotFoundError(f"EML folder does not exist: {folder}")
    source_location = str(folder.resolve())
    eml_paths = matching_files(folder, glob_pattern)
    logger.info("Found %d file(s) matching %s in %s", len(eml_paths), glob_pattern, source_location)
    if on_total is not None:
        on_total(len(eml_paths))
    for eml_path in eml_paths:
        logger.debug("Reading %s", eml_path.name)
        try:
            yield parse_eml_file(eml_path, source_location=source_location)
        except Exception as exc:  # one unreadable file must not abort the batch
            yield MailReadError(
                source_type="eml",
                source_location=source_location,
                message_identity=f"file:{eml_path.name}",
                reason=f"Could not read file ({type(exc).__name__}): {str(exc)[:200]}",
                origin=eml_path.name,
            )
