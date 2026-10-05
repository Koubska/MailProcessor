"""Fail a release whose tag does not match `version` in pyproject.toml.

Tag "3.1" matches version "3.1" or "3.1.0"; a leading "v" in the tag is ignored.
Usage: python scripts/check_version.py <tag>
"""

from __future__ import annotations

import sys
import tomllib
from pathlib import Path


def normalized(version: str) -> tuple[str, ...]:
    parts = version.strip().removeprefix("v").split(".")
    while len(parts) > 1 and parts[-1] == "0":
        parts.pop()
    return tuple(parts)


def main() -> int:
    tag = sys.argv[1]
    version = tomllib.loads(Path("pyproject.toml").read_text(encoding="utf-8"))["project"]["version"]
    if normalized(tag) != normalized(version):
        print(f"Tag {tag!r} does not match version {version!r} in pyproject.toml. Update the version, then tag again.")
        return 1
    print(f"Tag {tag!r} matches version {version!r}.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
