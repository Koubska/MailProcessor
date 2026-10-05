"""Lightweight i18n helpers with German default and English fallback."""

from __future__ import annotations

import locale
import os

SUPPORTED_LANGUAGES = ("de", "en")

CATALOG: dict[str, dict[str, str]] = {
    "de": {
        "app.title": "mailprocessor",
        "label.language": "Sprache",
        "lang.de": "Deutsch",
        "lang.en": "Englisch",
        "tab.config": "App-Konfiguration",
        "tab.fields": "Parsing-Felder",
        "label.config_file": "Konfigurationsdatei",
        "label.rules_file": "Regeldatei",
        "label.log_level": "Log-Level",
        "label.sqlite_path": "SQLite-Pfad",
        "label.output_xlsx": "Output-XLSX-Pfad",
        "label.sheet_data": "Daten-Sheet",
        "label.sheet_errors": "Fehler-Sheet",
        "label.max_age_days": "Maximales Alter (Tage)",
        "label.max_messages": "Max. neue Nachrichten pro Lauf (0 = alle)",
        "label.source_type": "Quelltyp",
        "label.field_column": "Spalte",
        "label.field_pattern": "Pattern",
        "label.field_required": "Pflichtfeld",
        "label.status.ready": "Bereit.",
        "label.status.saved": "config.toml und parsing_rules.toml gespeichert.",
        "label.status.save_failed": "Speichern fehlgeschlagen.",
        "label.status.reload": "Neu von Festplatte geladen.",
        "label.status.reload_failed": "Laden fehlgeschlagen.",
        "label.status.running": "Lauf läuft …",
        "label.status.run_finished": "Lauf abgeschlossen.",
        "label.status.run_failed": "Lauf fehlgeschlagen.",
        "label.status.field_updated": "Feld hinzugefügt/aktualisiert.",
        "label.status.field_removed": "Feld entfernt.",
        "label.status.field_reordered": "Feldreihenfolge aktualisiert.",
        "button.save": "Dateien speichern",
        "button.reload": "Neu laden",
        "button.run": "Ausführen",
        "button.field_add_update": "Feld hinzufügen/aktualisieren",
        "button.field_remove": "Entfernen",
        "button.field_up": "Nach oben",
        "button.field_down": "Nach unten",
        "source.details": "Quell-Details",
        "source.eml.folder": "E-Mail-Ordner",
        "button.browse": "Durchsuchen …",
        "dialog.choose_eml_folder": "Ordner mit den .eml-Dateien auswählen",
        "source.imap.host": "Host",
        "source.imap.port": "Port",
        "source.imap.username": "Benutzername",
        "source.imap.password": "Passwort",
        "source.imap.mailbox": "Mailbox",
        "source.imap.sender_filter": "Absender-Filter",
        "source.imap.use_ssl": "SSL verwenden",
        "bool.dry_run": "Dry Run",
        "view.field.column": "Spalte",
        "view.field.required": "Pflicht",
        "view.field.pattern": "Pattern",
        "view.field.required_yes": "ja",
        "view.field.required_no": "nein",
        "error.select_field_remove": "Bitte ein Feld zum Entfernen auswählen",
        "error.select_field_move": "Bitte ein Feld zum Verschieben auswählen",
    },
    "en": {
        "app.title": "mailprocessor",
        "label.language": "Language",
        "lang.de": "German",
        "lang.en": "English",
        "tab.config": "App config",
        "tab.fields": "Parsing fields",
        "label.config_file": "Config file",
        "label.rules_file": "Rules file",
        "label.log_level": "Log level",
        "label.sqlite_path": "SQLite path",
        "label.output_xlsx": "Output xlsx path",
        "label.sheet_data": "Data sheet",
        "label.sheet_errors": "Error sheet",
        "label.max_age_days": "Max age days",
        "label.max_messages": "Max new messages per run (0 = all)",
        "label.source_type": "Source type",
        "label.field_column": "Column",
        "label.field_pattern": "Pattern",
        "label.field_required": "Required",
        "label.status.ready": "Ready.",
        "label.status.saved": "Saved config.toml and parsing_rules.toml.",
        "label.status.save_failed": "Save failed.",
        "label.status.reload": "Reloaded from disk.",
        "label.status.reload_failed": "Reload failed.",
        "label.status.running": "Running …",
        "label.status.run_finished": "Run finished.",
        "label.status.run_failed": "Run failed.",
        "label.status.field_updated": "Field added/updated.",
        "label.status.field_removed": "Field removed.",
        "label.status.field_reordered": "Field order updated.",
        "button.save": "Save files",
        "button.reload": "Reload",
        "button.run": "Run",
        "button.field_add_update": "Add/Update field",
        "button.field_remove": "Remove",
        "button.field_up": "Move up",
        "button.field_down": "Move down",
        "source.details": "Source details",
        "source.eml.folder": "Email folder",
        "button.browse": "Browse …",
        "dialog.choose_eml_folder": "Choose the folder with the .eml files",
        "source.imap.host": "Host",
        "source.imap.port": "Port",
        "source.imap.username": "Username",
        "source.imap.password": "Password",
        "source.imap.mailbox": "Mailbox",
        "source.imap.sender_filter": "Sender filter",
        "source.imap.use_ssl": "Use SSL",
        "bool.dry_run": "Dry run",
        "view.field.column": "Column",
        "view.field.required": "Required",
        "view.field.pattern": "Pattern",
        "view.field.required_yes": "yes",
        "view.field.required_no": "no",
        "error.select_field_remove": "Select a field row to remove",
        "error.select_field_move": "Select a field row to move",
        "test.only_en": "English fallback",
    },
}


def resolve_language(explicit: str | None = None) -> str:
    for candidate in (explicit, os.getenv("MAILPROCESSOR_LANG")):
        if not candidate:
            continue
        code = candidate.strip().lower().split("_")[0].split("-")[0]
        if code in SUPPORTED_LANGUAGES:
            return code

    locale_value = locale.getlocale()[0] or ""
    locale_code = locale_value.lower().split("_")[0].split("-")[0]
    if locale_code in SUPPORTED_LANGUAGES:
        return locale_code

    return "de"


def t(key: str, language: str | None = None) -> str:
    lang = resolve_language(language)
    if key in CATALOG.get(lang, {}):
        return CATALOG[lang][key]
    if key in CATALOG["en"]:
        return CATALOG["en"][key]
    return key
