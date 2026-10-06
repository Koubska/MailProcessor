"""Static guard: mail sources must stay read-only. Fails if write/delete/mutating calls appear in sources/."""

import re
from pathlib import Path

SOURCES = Path(__file__).resolve().parents[1] / "src" / "mailprocessor" / "sources"

ALLOWED_CLIENT_CALLS = "login|select|response|uid|logout|examine|uid_search|uid_fetch|starttls|shutdown"

FORBIDDEN = [
    rf"\b_?client\.(?!(?:{ALLOWED_CLIENT_CALLS})\()\w+\(",  # any other IMAP command (STORE, COPY, EXPUNGE, ...)
    r"uid\(\s*[\"'](?!search|fetch)",  # raw UID commands other than SEARCH/FETCH
    r"readonly\s*=\s*False",
    r"\.(write_text|write_bytes|unlink|rmdir|rename|chmod|touch)\(",
    r"\b(os\.(remove|rename|replace|unlink)|shutil\.)",
    r"open\([^)]*[\"'][rb]*[wax+]",  # any write/append/update file mode
]


def test_sources_contain_no_mutating_operations() -> None:
    violations = []
    for path in sorted(SOURCES.glob("*.py")):
        for number, line in enumerate(path.read_text(encoding="utf-8").splitlines(), start=1):
            code = line.split("#", 1)[0]
            if any(re.search(pattern, code) for pattern in FORBIDDEN):
                violations.append(f"{path.name}:{number}: {line.strip()}")
    assert violations == []


def test_imap_commands_only_go_through_read_only_wrapper() -> None:
    source = (SOURCES / "imap_source.py").read_text(encoding="utf-8")
    raw_calls = re.findall(r"self\._client\.(\w+)\(", source)

    assert set(raw_calls) == {"login", "select", "response", "uid", "logout"}
    assert 'self._client.select(mailbox_argument(mailbox), readonly=True)' in source
    assert re.findall(r'self\._client\.uid\("(\w+)"', source) == ["search", "fetch"]
