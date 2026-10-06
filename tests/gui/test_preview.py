from importlib import resources
from pathlib import Path

from mailprocessor.config import DEFAULT_FILES, FieldRule, ParsingRules, Profile, load_parsing_rules
from mailprocessor.gui.preview import (
    RulePreview,
    SampleMail,
    best_profile,
    load_sample,
    preview_rule,
    profile_summary_text,
    result_text,
    sample_files,
    summary_text,
)

EXAMPLE = Path(__file__).resolve().parents[2] / "docs" / "example.eml"


def _default_rules() -> list[FieldRule]:
    with resources.as_file(DEFAULT_FILES / "parsing_rules.toml") as path:
        return load_parsing_rules(path).profiles[0].fields


def test_load_sample_reads_body_and_headers() -> None:
    sample = load_sample(EXAMPLE)

    assert sample.title == "example.eml"
    assert "Telefonnummer: 1234 567890" in sample.body
    # The example starts with "Von:"/"Betreff:", which a mail parser reads as headers.
    assert "Von: Max Mustermann" in sample.header_text


def test_preview_values_match_a_real_run() -> None:
    sample = load_sample(EXAMPLE)
    previews = [preview_rule(rule, sample) for rule in _default_rules()]

    assert [preview.value for preview in previews] == [
        "max.mustermann@mail.com",
        "Jan Mustermann",
        "Experimente",
        "Mo, 23.11.2026 (14:15-16:00 Uhr)",
        "1234 567890",
    ]


def test_preview_span_points_at_the_value_in_the_body() -> None:
    sample = load_sample(EXAMPLE)
    phone = preview_rule(FieldRule(column="T", type="label", label="Telefonnummer:"), sample)

    assert phone.span is not None
    assert sample.body[phone.span[0] : phone.span[1]] == "1234 567890"


def test_value_from_a_header_has_no_span() -> None:
    sample = SampleMail(title="x", body="Hallo", header_text="From: eltern@example.org")
    preview = preview_rule(FieldRule(column="M", type="email", label=["Von", "From"]), sample)

    assert preview.value == "eltern@example.org"
    assert preview.from_header


def test_not_found() -> None:
    preview = preview_rule(FieldRule(column="X", type="label", label="Gibt es nicht"), SampleMail("x", "Text"))

    assert not preview.found
    assert result_text(preview, "de") == "✗ nicht gefunden"
    assert result_text(None, "de") == "–"
    assert result_text(RulePreview("1234"), "de") == "✓ 1234"


def test_summary_explains_what_would_happen() -> None:
    rules = [
        FieldRule(column="Name", type="label", label="Name"),
        FieldRule(column="Telefon", type="label", label="Telefon"),
        FieldRule(column="Notiz", type="label", label="Notiz", required=False),
    ]
    found, missing = RulePreview("x"), RulePreview(None)

    assert summary_text(rules, [found, found, found], "fehler", "de") == (
        "✓ Alle 3 Felder gefunden – diese Mail würde in die Excel-Tabelle übernommen."
    )
    assert summary_text(rules, [found, found, missing], "fehler", "de") == (
        "✓ Alle Pflichtfelder gefunden (nicht gefunden: „Notiz“) – diese Mail würde übernommen."
    )
    assert summary_text(rules, [found, missing, missing], "fehler", "de") == (
        "✗ Pflichtfeld „Telefon“ nicht gefunden – diese Mail käme ins Blatt „fehler“."
    )
    assert summary_text(rules, [missing, missing, found], "errors", "en") == (
        "✗ Required fields “Name”, “Telefon” not found – this mail would go to the “errors” sheet."
    )


def test_sample_files_lists_matching_files_sorted(tmp_path: Path) -> None:
    for name in ("b.eml", "a.eml", "note.txt"):
        (tmp_path / name).write_text("x", encoding="utf-8")

    assert [path.name for path in sample_files(tmp_path)] == ["a.eml", "b.eml"]
    assert sample_files(tmp_path / "missing") == []


def test_best_profile_for_the_sample_mail() -> None:
    abmeldung = Profile(name="Abmeldung", fields=[FieldRule(column="Grund", type="label", label="Grund:")])
    rules = ParsingRules(profiles=[Profile(name="Anmeldung", fields=_default_rules()), abmeldung])
    sample = load_sample(EXAMPLE)

    best = best_profile(rules, sample)

    assert best.profile == "Anmeldung"
    assert profile_summary_text(best, "fehler", "de") == (
        "✓ Passt am besten zum Profil „Anmeldung“ (5 Felder gefunden) – diese Mail würde übernommen."
    )


def test_profile_summary_when_no_profile_fits() -> None:
    abmeldung = Profile(name="Abmeldung", fields=[FieldRule(column="Grund", type="label", label="Grund:")])
    rules = ParsingRules(profiles=[abmeldung, Profile(name="Leer", fields=[FieldRule(column="X", pattern="nie")])])

    best = best_profile(rules, SampleMail(title="t", body="Hallo"))

    assert profile_summary_text(best, "fehler", "de") == (
        "✗ Passt zu keinem Profil. Am ähnlichsten ist „Abmeldung“, dort fehlt „Grund“ – "
        "diese Mail käme ins Blatt „fehler“."
    )


def test_sample_files_ignore_case_of_the_extension(tmp_path: Path) -> None:
    for name in ("a.eml", "B.EML"):
        (tmp_path / name).write_text("x", encoding="utf-8")

    assert [path.name for path in sample_files(tmp_path)] == ["a.eml", "B.EML"]
