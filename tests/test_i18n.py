import mailprocessor.i18n as i18n


def test_default_language_is_german_when_locale_unsupported(monkeypatch) -> None:
    monkeypatch.delenv("MAILPROCESSOR_LANG", raising=False)
    monkeypatch.setattr(i18n.locale, "getlocale", lambda: ("fr_FR", "UTF-8"))

    assert i18n.resolve_language() == "de"
    assert i18n.t("button.run") == "Jetzt übertragen"


def test_env_language_switches_to_english(monkeypatch) -> None:
    monkeypatch.setenv("MAILPROCESSOR_LANG", "en")
    monkeypatch.setattr(i18n.locale, "getlocale", lambda: ("de_DE", "UTF-8"))

    assert i18n.resolve_language() == "en"
    assert i18n.t("button.run") == "Transfer now"


def test_missing_german_key_falls_back_to_english(monkeypatch) -> None:
    monkeypatch.delenv("MAILPROCESSOR_LANG", raising=False)
    monkeypatch.setattr(i18n.locale, "getlocale", lambda: ("de_DE", "UTF-8"))

    assert i18n.t("test.only_en") == "English fallback"


# Dotted string literals that are file or host names, not text keys.
NOT_KEYS = (".toml", ".png", ".log", ".json", ".com")


def _keys_used_in_code() -> set[str]:
    """Every dotted lower-case string literal in the source is a text key, unless it is a file or host name."""
    import re
    from pathlib import Path

    used = set()
    for path in (Path(__file__).resolve().parents[1] / "src" / "mailprocessor").rglob("*.py"):
        if path.name == "i18n.py":
            continue
        for key in re.findall(r"[\"']((?:[a-z_]+\.)+[a-z_]+)[\"']", path.read_text(encoding="utf-8")):
            if not key.endswith(NOT_KEYS):
                used.add(key)
    return used


# Keys built at runtime, e.g. f"rule.type.{rule_type}".
DYNAMIC_PREFIXES = (
    "rule.type.",
    "rule.hint.",
    "rule.describe.",
    "view.field.",
    "preview.summary.",
    "error.form.",
    "card.",
)


def test_every_key_used_in_the_code_exists_in_both_languages() -> None:
    used = _keys_used_in_code()

    assert used, "the scan found no keys"
    assert sorted(key for key in used if key not in i18n.CATALOG["de"]) == []
    assert sorted(key for key in used if key not in i18n.CATALOG["en"]) == []


def test_catalogs_have_the_same_keys_and_no_unused_ones() -> None:
    german, english = set(i18n.CATALOG["de"]), set(i18n.CATALOG["en"]) - {"test.only_en"}
    used = _keys_used_in_code()

    assert german == english
    unused = sorted(key for key in german if key not in used and not key.startswith(DYNAMIC_PREFIXES))
    assert unused == []
