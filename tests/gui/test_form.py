"""The E-Mails and Einstellungen inputs: validation and conversion to a config."""

import pytest

from mailprocessor.config import MailFilter
from mailprocessor.gui.config_files import default_config as _default_config
from mailprocessor.gui.form import config_from_form, form_values


def test_form_values_roundtrip_to_the_same_config() -> None:
    config = _default_config()

    rebuilt, errors = config_from_form(form_values(config), "de")

    assert errors == {}
    assert rebuilt == config


def test_form_keeps_the_profile_sheets_choice() -> None:
    config = _default_config()
    config.app.profile_sheets = "per_profile"

    rebuilt, _errors = config_from_form(form_values(config), "de")

    assert rebuilt is not None and rebuilt.app.profile_sheets == "per_profile"


@pytest.mark.parametrize(
    ("changes", "key", "message"),
    [
        ({"eml_folder": " "}, "eml_folder", "Ordner auswählen"),
        ({"max_age_days": "drei"}, "max_age_days", "ganze Zahl"),
        ({"max_messages": "-1"}, "max_messages", "ganze Zahl"),
        ({"output_xlsx": "out/export.csv"}, "output_xlsx", ".xlsx"),
        ({"sheet_errors": "daten"}, "sheet_errors", "verschiedene Namen"),
        ({"sheet_errors": "Daten"}, "sheet_errors", "verschiedene Namen"),  # one sheet for Excel
        ({"source_type": "imap", "imap_host": ""}, "imap_host", "E-Mail-Server"),
        ({"source_type": "imap", "imap_port": "99999"}, "imap_port", "Portnummer"),
        # "²" and "³" (AltGr+2/3 on German keyboards) count as digits for str.isdigit, but int() rejects them.
        ({"max_age_days": "3²"}, "max_age_days", "ganze Zahl"),
        ({"max_messages": "¹"}, "max_messages", "ganze Zahl"),
        ({"source_type": "imap", "imap_port": "99³"}, "imap_port", "Portnummer"),
    ],
)
def test_invalid_inputs_are_reported_per_field(changes: dict, key: str, message: str) -> None:
    values = {**form_values(_default_config()), **changes}

    config, errors = config_from_form(values, "de")

    assert config is None
    assert message in errors[key]


def test_invalid_port_of_the_inactive_source_does_not_block_saving() -> None:
    values = {**form_values(_default_config()), "imap_port": "abc"}

    config, errors = config_from_form(values, "de")

    assert errors == {}
    assert config is not None and config.source.imap is not None and config.source.imap.port == 993


def test_filter_entries_are_typed_separated_by_semicolons() -> None:
    values = {
        **form_values(_default_config()),
        "filter_subject": "Kontaktformular;  Anmeldung ;",
        "filter_sender": "schule@example.com",
    }

    config, errors = config_from_form(values, "de")

    assert errors == {}
    assert config.filter == MailFilter(subject=["Kontaktformular", "Anmeldung"], sender=["schule@example.com"])
    assert form_values(config)["filter_subject"] == "Kontaktformular; Anmeldung"
    assert config_from_form(form_values(config), "de")[0] == config
