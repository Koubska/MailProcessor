"""Lightweight i18n helpers with German default and English fallback."""

from __future__ import annotations

import locale
import os

SUPPORTED_LANGUAGES = ("de", "en")

CATALOG: dict[str, dict[str, str]] = {
    "de": {
        "app.title": "Mail Processor",
        "label.language": "Sprache",
        "lang.de": "Deutsch",
        "lang.en": "Englisch",
        "tab.config": "App-Konfiguration",
        "tab.fields": "Parsing-Felder",
        "label.config_file": "Konfigurationsdatei",
        "label.rules_file": "Regeldatei",
        "label.log_level": "Log-Level",
        "label.sqlite_path": "SQLite-Pfad",
        "label.output_xlsx": "Excel-Datei",
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
        "view.field.column": "Spalte",
        "view.field.required": "Pflicht",
        "view.field.pattern": "Pattern",
        "view.field.required_yes": "ja",
        "view.field.required_no": "nein",
        "error.select_field_remove": "Bitte ein Feld zum Entfernen auswählen",
        "error.select_field_move": "Bitte ein Feld zum Verschieben auswählen",
        "button.open_excel": "Excel öffnen",
        "button.test_run": "Testlauf",
        "button.stop": "Stopp",
        "button.start_over": "Alles neu exportieren …",
        "button.open_log": "Protokoll öffnen",
        "label.status.progress": "Verarbeite E-Mail {current} von {total} …",
        "label.status.stopping": "Wird angehalten …",
        "summary.cancelled": "Angehalten – beim nächsten Ausführen geht es an dieser Stelle weiter.",
        "confirm.start_over": (
            "Alle E-Mails werden noch einmal in eine neue Excel-Datei übertragen.\n\n"
            "Die bisherige Excel-Datei bleibt als Sicherung erhalten (Name mit „_backup_“ und Datum). "
            "Die E-Mails selbst werden nicht verändert.\n\nFortfahren?"
        ),
        "info.backup_created": "Bisherige Excel-Datei gesichert als {file}",
        "info.no_log_yet": "Es gibt noch kein Protokoll ({file}). Es entsteht beim ersten Lauf.",
        "error.open_file": "Die Datei konnte nicht geöffnet werden: {file}\n\n{error}",
        "button.advanced_show": "▸ Erweiterte Einstellungen",
        "button.advanced_hide": "▾ Erweiterte Einstellungen",
        "dialog.choose_output_xlsx": "Excel-Datei für die Ergebnisse wählen",
        "info.no_excel_yet": "Es gibt noch keine Excel-Datei ({file}). Sie entsteht beim ersten Klick auf „Ausführen“.",
        "error.open_excel": "Die Excel-Datei konnte nicht geöffnet werden: {file}\n\n{error}",
        "summary.dry_run": (
            "Testlauf: {new} neue E-Mail(s) gefunden, davon {ok} fehlerfrei und {failed} mit Problemen. "
            "Es wurde nichts gespeichert."
        ),
        "summary.processed": "{count} neue Zeile(n) in {file} eingetragen.",
        "summary.failed": "{count} E-Mail(s) mit Problemen – Details im Blatt „{sheet}“.",
        "summary.nothing_new": "Keine neuen E-Mails gefunden.",
        "summary.skipped": "{count} bereits verarbeitete E-Mail(s) übersprungen.",
        "error.details": "Details",
        "error.workbook_locked": (
            "Die Excel-Datei ist gerade in einem anderen Programm geöffnet. "
            "Bitte Excel schließen und erneut auf „Ausführen“ klicken."
        ),
        "error.sheet_columns": (
            "Die Spalten in der Excel-Datei passen nicht mehr zu den Parsing-Feldern. "
            "Bitte im Reiter „App-Konfiguration“ einen eine neue Excel-Datei wählen "
            "oder die vorherigen Felder wiederherstellen."
        ),
        "error.mail_folder_missing": (
            "Der E-Mail-Ordner wurde nicht gefunden. "
            "Bitte im Reiter „App-Konfiguration“ über „Durchsuchen …“ einen vorhandenen Ordner auswählen."
        ),
        "error.imap_password_missing": "Bitte im Reiter „App-Konfiguration“ das IMAP-Passwort eingeben.",
        "error.imap_login": "Anmeldung am Mailserver fehlgeschlagen. Bitte Benutzername und Passwort prüfen.",
        "error.imap_tls": (
            "Keine sichere Verbindung zum Mailserver möglich. "
            "Bitte Host, Port und die Einstellung „SSL verwenden“ prüfen."
        ),
        "error.imap_connection": (
            "Der Mailserver ist nicht erreichbar. "
            "Bitte Internetverbindung sowie Host und Port prüfen."
        ),
        "error.unexpected": (
            "Unerwarteter Fehler ({name}). Für Details den Log-Level auf DEBUG stellen und erneut ausführen."
        ),
    },
    "en": {
        "app.title": "Mail Processor",
        "label.language": "Language",
        "lang.de": "German",
        "lang.en": "English",
        "tab.config": "App config",
        "tab.fields": "Parsing fields",
        "label.config_file": "Config file",
        "label.rules_file": "Rules file",
        "label.log_level": "Log level",
        "label.sqlite_path": "SQLite path",
        "label.output_xlsx": "Excel file",
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
        "view.field.column": "Column",
        "view.field.required": "Required",
        "view.field.pattern": "Pattern",
        "view.field.required_yes": "yes",
        "view.field.required_no": "no",
        "error.select_field_remove": "Select a field row to remove",
        "error.select_field_move": "Select a field row to move",
        "button.open_excel": "Open Excel",
        "button.test_run": "Test run",
        "button.stop": "Stop",
        "button.start_over": "Export everything again …",
        "button.open_log": "Open log",
        "label.status.progress": "Processing email {current} of {total} …",
        "label.status.stopping": "Stopping …",
        "summary.cancelled": "Stopped – the next run continues from here.",
        "confirm.start_over": (
            "All emails are transferred again into a new Excel file.\n\n"
            "The current Excel file is kept as a backup (name with “_backup_” and the date). "
            "The emails themselves are not changed.\n\nContinue?"
        ),
        "info.backup_created": "Previous Excel file backed up as {file}",
        "info.no_log_yet": "There is no log yet ({file}). It is created by the first run.",
        "error.open_file": "Could not open the file: {file}\n\n{error}",
        "button.advanced_show": "▸ Advanced settings",
        "button.advanced_hide": "▾ Advanced settings",
        "dialog.choose_output_xlsx": "Choose the Excel file for the results",
        "info.no_excel_yet": "There is no Excel file yet ({file}). It is created by the first click on “Run”.",
        "error.open_excel": "Could not open the Excel file: {file}\n\n{error}",
        "summary.dry_run": (
            "Test run: found {new} new email(s), {ok} without problems and {failed} with problems. "
            "Nothing was saved."
        ),
        "summary.processed": "Added {count} new row(s) to {file}.",
        "summary.failed": "{count} email(s) with problems – see the “{sheet}” sheet.",
        "summary.nothing_new": "No new emails found.",
        "summary.skipped": "Skipped {count} email(s) processed earlier.",
        "error.details": "Details",
        "error.workbook_locked": "The Excel file is open in another program. Close Excel and click “Run” again.",
        "error.sheet_columns": (
            "The columns in the Excel file no longer match the parsing fields. "
            "Choose a new Excel file in the “App config” tab or restore the previous fields."
        ),
        "error.mail_folder_missing": (
            "The email folder was not found. Use “Browse …” in the “App config” tab to choose an existing folder."
        ),
        "error.imap_password_missing": "Enter the IMAP password in the “App config” tab.",
        "error.imap_login": "The mail server rejected the login. Check username and password.",
        "error.imap_tls": "Could not connect securely to the mail server. Check host, port and “Use SSL”.",
        "error.imap_connection": "The mail server cannot be reached. Check the internet connection, host and port.",
        "error.unexpected": "Unexpected error ({name}). Set the log level to DEBUG and run again for details.",
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
