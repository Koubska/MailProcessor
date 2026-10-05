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
