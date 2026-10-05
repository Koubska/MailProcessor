"""Lightweight i18n helpers with German default and English fallback."""

from __future__ import annotations

import locale
import os

SUPPORTED_LANGUAGES = ("de", "en")

CATALOG: dict[str, dict[str, str]] = {
    "de": {
        "app.title": "Mail Processor",
        "tab.fields": "Felder",
        "label.config_file": "Konfigurationsdatei",
        "label.rules_file": "Regeldatei",
        "label.log_level": "Log-Level",
        "label.sqlite_path": "SQLite-Pfad",
        "label.field_column": "Spalte",
        "label.field_pattern": "Regulärer Ausdruck",
        "label.field_type": "Art",
        "label.field_labels": "Bezeichnung",
        "label.field_labels_optional": "Bezeichnung (optional)",
        "label.field_start": "Anfang",
        "label.field_end": "Ende",
        "button.as_regex": "Als Regex bearbeiten",
        "rule.type.label": "Text nach Bezeichnung",
        "rule.type.next_line": "Wert in nächster Zeile",
        "rule.type.between": "Text zwischen zwei Stellen",
        "rule.type.email": "E-Mail-Adresse",
        "rule.type.regex": "Experte (Regex)",
        "rule.hint.label": (
            "Übernimmt den Rest der Zeile nach der Bezeichnung, z. B. „Telefonnummer:“. "
            "Groß-/Kleinschreibung, Leerzeichen und der Doppelpunkt sind egal. Mehrere Bezeichnungen mit ; trennen."
        ),
        "rule.hint.next_line": (
            "Die Bezeichnung steht allein in einer Zeile, der Wert in der nächsten. "
            "Mehrere Bezeichnungen mit ; trennen."
        ),
        "rule.hint.between": "Übernimmt den Text zwischen Anfang und Ende, auch über Zeilenumbrüche hinweg.",
        "rule.hint.email": (
            "Findet die E-Mail-Adresse in der Zeile mit der Bezeichnung, z. B. „Von“. "
            "Ohne Bezeichnung: die erste Adresse im Text. Mehrere Bezeichnungen mit ; trennen."
        ),
        "rule.hint.regex": (
            "Regulärer Ausdruck (Python). Die erste Klammergruppe ( ) ist der Wert. "
            "Wird nicht mehr in eine einfache Regel zurückverwandelt."
        ),
        "preview.not_found": "✗ nicht gefunden",
        "preview.summary.ok": (
            "✓ Alle {count} Felder gefunden – diese Mail würde in die Excel-Tabelle übernommen."
        ),
        "preview.summary.ok_optional_missing": (
            "✓ Alle Pflichtfelder gefunden (nicht gefunden: {columns}) – diese Mail würde übernommen."
        ),
        "preview.summary.missing_one": (
            "✗ Pflichtfeld {columns} nicht gefunden – diese Mail käme ins Blatt „{sheet}“."
        ),
        "preview.summary.missing_many": (
            "✗ Pflichtfelder {columns} nicht gefunden – diese Mail käme ins Blatt „{sheet}“."
        ),
        "preview.title": "Test mit Beispiel-Mail",
        "button.sample_load": "Mail laden …",
        "button.sample_paste": "Text einfügen",
        "button.field_new": "Neues Feld",
        "preview.position": "{name} ({index} von {total})",
        "preview.pasted": "Eingefügter Text",
        "preview.none": "Keine Beispiel-Mail geladen",
        "preview.no_sample_summary": (
            "Eine Beispiel-Mail laden oder den Text einer Mail einfügen, um die Regeln direkt zu testen."
        ),
        "preview.header_note": (
            "Die Kopfzeilen der Mail (z. B. From:) werden ebenfalls durchsucht. "
            "Änderungen am Text hier dienen nur dem Test und ändern keine Mail."
        ),
        "preview.editor.found": "In der Beispiel-Mail: ✓ {value}",
        "preview.editor.from_header": "In der Beispiel-Mail: ✓ {value} (aus den Kopfzeilen der Mail)",
        "preview.editor.not_found": (
            "In der Beispiel-Mail: ✗ nicht gefunden – Bezeichnung prüfen oder eine andere Art wählen."
        ),
        "preview.editor.no_sample": "Eine Beispiel-Mail laden, um das Ergebnis hier direkt zu sehen.",
        "preview.editor.waiting": "Eingaben ausfüllen – das Ergebnis in der Beispiel-Mail erscheint hier.",
        "dialog.choose_sample": "Beispiel-Mail auswählen",
        "dialog.eml_files": "E-Mails (.eml)",
        "dialog.all_files": "Alle Dateien",
        "error.clipboard_empty": "Die Zwischenablage enthält keinen Text. Bitte zuerst den Text einer Mail kopieren.",
        "error.sample_unreadable": "Die Mail konnte nicht gelesen werden: {file}\n\n{error}",
        "error.form.number": "Bitte eine ganze Zahl eingeben (0 = keine Begrenzung).",
        "error.form.folder_empty": "Bitte einen Ordner auswählen.",
        "error.form.host_empty": "Bitte den E-Mail-Server eingeben, z. B. imap.gmx.net.",
        "error.form.username_empty": "Bitte den Benutzernamen eingeben (meist die E-Mail-Adresse).",
        "error.form.port": "Bitte eine Portnummer zwischen 1 und 65535 eingeben.",
        "error.form.output_empty": "Bitte eine Excel-Datei wählen.",
        "error.form.output_suffix": "Die Excel-Datei muss auf .xlsx enden.",
        "error.form.sheet_empty": "Bitte einen Namen eingeben.",
        "error.form.sheets_equal": "Die beiden Blätter brauchen verschiedene Namen.",
        "error.form.sqlite_empty": "Bitte einen Dateipfad eingeben.",
        "card.mails.title": "E-Mails",
        "card.fields.title": "Felder",
        "card.excel.title": "Excel-Datei",
        "card.mails.folder_ready": "Ordner {folder} – {count} E-Mail(s)",
        "card.mails.folder_empty": "Ordner {folder} ist leer – bitte .eml-Dateien hineinlegen.",
        "card.mails.folder_missing": "Ordner {folder} nicht gefunden – bitte einen Ordner auswählen.",
        "card.mails.imap_incomplete": "Postfach noch nicht eingerichtet.",
        "card.mails.imap_password": "Postfach {user} – bitte das Passwort eingeben (wird nicht gespeichert).",
        "card.mails.imap_ready": "Postfach {user} auf {host}",
        "card.fields.none": "Noch keine Felder angelegt.",
        "card.fields.count": "{count} Feld(er)",
        "card.fields.ok": "{count} – in der Beispiel-Mail alles gefunden",
        "card.fields.missing": "{count} – in der Beispiel-Mail fehlt {columns}",
        "card.excel.new": "{file} – wird beim ersten Übertragen angelegt",
        "card.excel.rows": "{file} – {rows} Zeile(n)",
        "card.excel.problems": ", {count} E-Mail(s) mit Problemen",
        "tab.start": "Start",
        "tab.mails": "E-Mails",
        "tab.settings": "Einstellungen",
        "start.title": "E-Mails in die Excel-Tabelle übertragen",
        "start.intro": (
            "Neue E-Mails werden gelesen und als Zeilen in die Excel-Datei übernommen. "
            "Bereits übertragene E-Mails werden übersprungen, die E-Mails selbst nie verändert."
        ),
        "start.test_run_hint": "Der Testlauf zeigt, was passieren würde, ohne etwas zu speichern.",
        "button.change": "Ändern",
        "button.edit": "Bearbeiten",
        "button.details_show": "▸ Details anzeigen",
        "button.details_hide": "▾ Details ausblenden",
        "button.open_folder": "Ordner öffnen",
        "button.open_config_folder": "Ordner mit den Einstellungen öffnen",
        "button.test_connection": "Verbindung testen",
        "button.field_add": "Feld hinzufügen",
        "button.cancel": "Abbrechen",
        "status.saved": "✓ Alle Änderungen gespeichert",
        "status.not_saved": "Nicht gespeichert – bitte die rot markierten Eingaben prüfen.",
        "error.fix_inputs": "Bitte zuerst die rot markierten Eingaben korrigieren.",
        "error.load": "Die Einstellungen konnten nicht gelesen werden; es werden Standardwerte angezeigt.\n\n{error}",
        "confirm.close_running": (
            "Es läuft gerade eine Übertragung. Trotzdem beenden? Bisher Übertragenes bleibt "
            "erhalten."
        ),
        "confirm.close_invalid": "Einige Eingaben sind ungültig und wurden nicht gespeichert. Trotzdem beenden?",
        "confirm.remove_field": (
            "Das Feld „{column}“ löschen? Die Spalte wird bei neuen Excel-Dateien nicht mehr "
            "angelegt."
        ),
        "label.status.field_added": "Feld „{column}“ hinzugefügt.",
        "info.folder_missing": "Den Ordner gibt es nicht: {file}",
        "mails.title": "Woher kommen die E-Mails?",
        "mails.intro": "Die E-Mails werden nur gelesen – nichts wird verschoben, gelöscht oder als gelesen markiert.",
        "mails.eml.choice": "Aus einem Ordner mit gespeicherten E-Mails (.eml-Dateien)",
        "mails.eml.hint": (
            "Zum Beispiel E-Mails, die aus Outlook oder Thunderbird in einen Ordner gezogen oder "
            "als Datei gespeichert wurden."
        ),
        "mails.eml.count": "{count} E-Mail(s) in diesem Ordner",
        "mails.eml.missing": "Diesen Ordner gibt es noch nicht.",
        "mails.imap.choice": "Direkt aus dem Postfach (IMAP)",
        "mails.imap.hint": "Die Zugangsdaten stehen in der Hilfe Ihres E-Mail-Anbieters (Stichwort „IMAP“).",
        "mails.imap.host": "E-Mail-Server",
        "mails.imap.host_hint": "z. B. imap.gmx.net oder outlook.office365.com",
        "mails.imap.username": "Benutzername",
        "mails.imap.username_hint": "meist Ihre E-Mail-Adresse",
        "mails.imap.password": "Passwort",
        "mails.imap.password_hint": "wird nicht gespeichert – nach jedem Start neu eingeben",
        "mails.imap.mailbox": "Ordner im Postfach",
        "mails.imap.mailbox_hint": "INBOX = Posteingang",
        "mails.imap.sender_filter": "Nur E-Mails von",
        "mails.imap.sender_filter_hint": "optional, z. B. formular@schule.de",
        "mails.imap.ssl": "Verschlüsselte Verbindung (SSL)",
        "mails.imap.port": "Port",
        "mails.imap.testing": "Verbinde …",
        "mails.imap.connected": "✓ Verbindung klappt.",
        "mails.imap.connected_count": "✓ Verbindung klappt – {count} E-Mail(s) im Ordner „{mailbox}“.",
        "fields.title": "Was soll übernommen werden?",
        "fields.intro": (
            "Jedes Feld wird eine Spalte in der Excel-Tabelle. Rechts sehen Sie an einer Beispiel-Mail sofort, "
            "was gefunden wird. Änderungen werden automatisch gespeichert."
        ),
        "fields.editor_edit": "Ausgewähltes Feld bearbeiten",
        "fields.editor_new": "Neues Feld",
        "settings.title": "Einstellungen",
        "settings.excel": "Excel-Datei",
        "settings.excel_hint": "Neue Zeilen werden angehängt; die Datei wird nie ersetzt.",
        "settings.sheet_data": "Blatt für die Daten",
        "settings.sheet_errors": "Blatt für Probleme",
        "settings.language": "Sprache",
        "settings.limits": "Begrenzungen (0 = keine)",
        "settings.max_age_days": "Nur E-Mails der letzten … Tage",
        "settings.max_messages": "Höchstens … neue E-Mails pro Lauf",
        "settings.maintenance": "Wartung",
        "settings.start_over_hint": (
            "Überträgt alle E-Mails noch einmal in eine neue Excel-Datei. Die bisherige Datei bleibt als "
            "Sicherung erhalten."
        ),
        "settings.technical": "Technisches",
        "rule.or": " oder ",
        "rule.describe.label": "Zeile nach {labels}",
        "rule.describe.next_line": "Zeile unter {labels}",
        "rule.describe.between": "Zwischen {start} und {end}",
        "rule.describe.email": "E-Mail-Adresse in der Zeile {labels}",
        "rule.describe.email_anywhere": "Erste E-Mail-Adresse im Text",
        "error.rule.column": "Bitte einen Spaltennamen eingeben.",
        "error.rule.duplicate": "Die Spalte „{column}“ gibt es schon. Bitte einen anderen Namen wählen.",
        "error.rule.label": "Bitte eine Bezeichnung eingeben, z. B. „Telefonnummer:“.",
        "error.rule.between": "Bitte Anfang und Ende eingeben.",
        "error.rule.pattern": "Bitte einen regulären Ausdruck eingeben.",
        "label.field_required": "Pflichtfeld",
        "label.status.running": "Lauf läuft …",
        "button.run": "Jetzt übertragen",
        "button.field_remove": "Löschen",
        "button.field_up": "▲ Nach oben",
        "button.field_down": "▼ Nach unten",
        "button.browse": "Durchsuchen …",
        "dialog.choose_eml_folder": "Ordner mit den .eml-Dateien auswählen",
        "view.field.column": "Spalte",
        "view.field.required": "Pflicht",
        "view.field.type": "Art",
        "view.field.description": "Sucht nach",
        "view.field.result": "Ergebnis",
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
        "summary.cancelled": "Angehalten – bei der nächsten Übertragung geht es an dieser Stelle weiter.",
        "confirm.start_over": (
            "Alle E-Mails werden noch einmal in eine neue Excel-Datei übertragen.\n\n"
            "Die bisherige Excel-Datei bleibt als Sicherung erhalten (Name mit „_backup_“ und Datum). "
            "Die E-Mails selbst werden nicht verändert.\n\nFortfahren?"
        ),
        "info.backup_created": "Bisherige Excel-Datei gesichert als {file}",
        "info.no_log_yet": "Es gibt noch kein Protokoll ({file}). Es entsteht beim ersten Lauf.",
        "error.open_file": "Die Datei konnte nicht geöffnet werden: {file}\n\n{error}",
        "dialog.choose_output_xlsx": "Excel-Datei für die Ergebnisse wählen",
        "info.no_excel_yet": "Es gibt noch keine Excel-Datei ({file}). Sie entsteht bei der ersten Übertragung.",
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
            "Bitte Excel schließen und erneut auf „Jetzt übertragen“ klicken."
        ),
        "error.sheet_columns": (
            "Die Spalten in der Excel-Datei passen nicht mehr zu den Feldern. "
            "Bitte unter „Einstellungen“ eine neue Excel-Datei wählen "
            "oder die vorherigen Felder wiederherstellen."
        ),
        "error.mail_folder_missing": (
            "Der E-Mail-Ordner wurde nicht gefunden. "
            "Bitte im Reiter „E-Mails“ über „Durchsuchen …“ einen vorhandenen Ordner auswählen."
        ),
        "error.imap_password_missing": "Bitte im Reiter „E-Mails“ das Passwort für das Postfach eingeben.",
        "error.imap_login": "Anmeldung am Mailserver fehlgeschlagen. Bitte Benutzername und Passwort prüfen.",
        "error.imap_tls": (
            "Keine sichere Verbindung zum Mailserver möglich. "
            "Bitte E-Mail-Server, Port und „Verschlüsselte Verbindung (SSL)“ prüfen."
        ),
        "error.imap_connection": (
            "Der Mailserver ist nicht erreichbar. "
            "Bitte Internetverbindung, E-Mail-Server und Port prüfen."
        ),
        "error.unexpected": (
            "Unerwarteter Fehler ({name}). Für Details unter „Einstellungen“ den Log-Level auf DEBUG stellen "
            "und erneut übertragen."
        ),
    },
    "en": {
        "app.title": "Mail Processor",
        "tab.fields": "Fields",
        "label.config_file": "Config file",
        "label.rules_file": "Rules file",
        "label.log_level": "Log level",
        "label.sqlite_path": "SQLite path",
        "label.field_column": "Column",
        "label.field_pattern": "Regular expression",
        "label.field_type": "Type",
        "label.field_labels": "Label",
        "label.field_labels_optional": "Label (optional)",
        "label.field_start": "Start",
        "label.field_end": "End",
        "button.as_regex": "Edit as regex",
        "rule.type.label": "Text after label",
        "rule.type.next_line": "Value on next line",
        "rule.type.between": "Text between two phrases",
        "rule.type.email": "Email address",
        "rule.type.regex": "Expert (regex)",
        "rule.hint.label": (
            "Takes the rest of the line after the label, e.g. “Phone:”. "
            "Case, spaces and the colon do not matter. Separate several labels with ;."
        ),
        "rule.hint.next_line": (
            "The label stands alone on a line, the value on the next one. Separate several labels with ;."
        ),
        "rule.hint.between": "Takes the text between start and end, also across line breaks.",
        "rule.hint.email": (
            "Finds the email address in the line with the label, e.g. “From”. "
            "Without a label: the first address in the text. Separate several labels with ;."
        ),
        "rule.hint.regex": (
            "Regular expression (Python). The first group ( ) is the value. "
            "It is not turned back into a simple rule."
        ),
        "preview.not_found": "✗ not found",
        "preview.summary.ok": "✓ All {count} fields found – this mail would be added to the Excel file.",
        "preview.summary.ok_optional_missing": (
            "✓ All required fields found (not found: {columns}) – this mail would be added."
        ),
        "preview.summary.missing_one": (
            "✗ Required field {columns} not found – this mail would go to the “{sheet}” sheet."
        ),
        "preview.summary.missing_many": (
            "✗ Required fields {columns} not found – this mail would go to the “{sheet}” sheet."
        ),
        "preview.title": "Test with a sample mail",
        "button.sample_load": "Load mail …",
        "button.sample_paste": "Paste text",
        "button.field_new": "New field",
        "preview.position": "{name} ({index} of {total})",
        "preview.pasted": "Pasted text",
        "preview.none": "No sample mail loaded",
        "preview.no_sample_summary": "Load a sample mail or paste the text of a mail to test the rules right away.",
        "preview.header_note": (
            "The mail's header lines (e.g. From:) are searched too. "
            "Edits to the text here are only for testing and change no mail."
        ),
        "preview.editor.found": "In the sample mail: ✓ {value}",
        "preview.editor.from_header": "In the sample mail: ✓ {value} (from the mail's header lines)",
        "preview.editor.not_found": "In the sample mail: ✗ not found – check the label or choose another type.",
        "preview.editor.no_sample": "Load a sample mail to see the result here right away.",
        "preview.editor.waiting": "Fill in the inputs – the result in the sample mail appears here.",
        "dialog.choose_sample": "Choose a sample mail",
        "dialog.eml_files": "Emails (.eml)",
        "dialog.all_files": "All files",
        "error.clipboard_empty": "The clipboard contains no text. Copy the text of a mail first.",
        "error.sample_unreadable": "Could not read the mail: {file}\n\n{error}",
        "error.form.number": "Enter a whole number (0 = no limit).",
        "error.form.folder_empty": "Choose a folder.",
        "error.form.host_empty": "Enter the mail server, e.g. imap.gmx.net.",
        "error.form.username_empty": "Enter the username (usually the email address).",
        "error.form.port": "Enter a port number between 1 and 65535.",
        "error.form.output_empty": "Choose an Excel file.",
        "error.form.output_suffix": "The Excel file must end in .xlsx.",
        "error.form.sheet_empty": "Enter a name.",
        "error.form.sheets_equal": "The two sheets need different names.",
        "error.form.sqlite_empty": "Enter a file path.",
        "card.mails.title": "Emails",
        "card.fields.title": "Fields",
        "card.excel.title": "Excel file",
        "card.mails.folder_ready": "Folder {folder} – {count} email(s)",
        "card.mails.folder_empty": "Folder {folder} is empty – put .eml files into it.",
        "card.mails.folder_missing": "Folder {folder} not found – choose a folder.",
        "card.mails.imap_incomplete": "Mailbox not set up yet.",
        "card.mails.imap_password": "Mailbox {user} – enter the password (it is not saved).",
        "card.mails.imap_ready": "Mailbox {user} on {host}",
        "card.fields.none": "No fields yet.",
        "card.fields.count": "{count} field(s)",
        "card.fields.ok": "{count} – everything found in the sample mail",
        "card.fields.missing": "{count} – missing in the sample mail: {columns}",
        "card.excel.new": "{file} – created by the first transfer",
        "card.excel.rows": "{file} – {rows} row(s)",
        "card.excel.problems": ", {count} email(s) with problems",
        "tab.start": "Start",
        "tab.mails": "Emails",
        "tab.settings": "Settings",
        "start.title": "Transfer emails to the Excel file",
        "start.intro": (
            "New emails are read and added as rows to the Excel file. "
            "Emails transferred before are skipped; the emails themselves are never changed."
        ),
        "start.test_run_hint": "The test run shows what would happen without saving anything.",
        "button.change": "Change",
        "button.edit": "Edit",
        "button.details_show": "▸ Show details",
        "button.details_hide": "▾ Hide details",
        "button.open_folder": "Open folder",
        "button.open_config_folder": "Open the settings folder",
        "button.test_connection": "Test connection",
        "button.field_add": "Add field",
        "button.cancel": "Cancel",
        "status.saved": "✓ All changes saved",
        "status.not_saved": "Not saved – check the inputs marked in red.",
        "error.fix_inputs": "Correct the inputs marked in red first.",
        "error.load": "The settings could not be read; defaults are shown.\n\n{error}",
        "confirm.close_running": "A transfer is running. Quit anyway? What was transferred so far is kept.",
        "confirm.close_invalid": "Some inputs are invalid and were not saved. Quit anyway?",
        "confirm.remove_field": "Delete the field “{column}”? New Excel files will no longer get this column.",
        "label.status.field_added": "Field “{column}” added.",
        "info.folder_missing": "The folder does not exist: {file}",
        "mails.title": "Where do the emails come from?",
        "mails.intro": "Emails are only read – nothing is moved, deleted or marked as read.",
        "mails.eml.choice": "From a folder with saved emails (.eml files)",
        "mails.eml.hint": "For example emails dragged from Outlook or Thunderbird into a folder, or saved as files.",
        "mails.eml.count": "{count} email(s) in this folder",
        "mails.eml.missing": "This folder does not exist yet.",
        "mails.imap.choice": "Directly from the mailbox (IMAP)",
        "mails.imap.hint": "Your email provider's help pages list these details (search for “IMAP”).",
        "mails.imap.host": "Mail server",
        "mails.imap.host_hint": "e.g. imap.gmail.com or outlook.office365.com",
        "mails.imap.username": "Username",
        "mails.imap.username_hint": "usually your email address",
        "mails.imap.password": "Password",
        "mails.imap.password_hint": "not saved – enter it again after each start",
        "mails.imap.mailbox": "Mailbox folder",
        "mails.imap.mailbox_hint": "INBOX = inbox",
        "mails.imap.sender_filter": "Only emails from",
        "mails.imap.sender_filter_hint": "optional, e.g. form@school.org",
        "mails.imap.ssl": "Encrypted connection (SSL)",
        "mails.imap.port": "Port",
        "mails.imap.testing": "Connecting …",
        "mails.imap.connected": "✓ Connection works.",
        "mails.imap.connected_count": "✓ Connection works – {count} email(s) in the folder “{mailbox}”.",
        "fields.title": "What should be transferred?",
        "fields.intro": (
            "Each field becomes a column in the Excel file. On the right, a sample mail shows right away "
            "what is found. Changes are saved automatically."
        ),
        "fields.editor_edit": "Edit the selected field",
        "fields.editor_new": "New field",
        "settings.title": "Settings",
        "settings.excel": "Excel file",
        "settings.excel_hint": "New rows are appended; the file is never replaced.",
        "settings.sheet_data": "Sheet for the data",
        "settings.sheet_errors": "Sheet for problems",
        "settings.language": "Language",
        "settings.limits": "Limits (0 = none)",
        "settings.max_age_days": "Only emails of the last … days",
        "settings.max_messages": "At most … new emails per run",
        "settings.maintenance": "Maintenance",
        "settings.start_over_hint": (
            "Transfers all emails again into a new Excel file. The current file is kept as a backup."
        ),
        "settings.technical": "Technical",
        "rule.or": " or ",
        "rule.describe.label": "Line after {labels}",
        "rule.describe.next_line": "Line below {labels}",
        "rule.describe.between": "Between {start} and {end}",
        "rule.describe.email": "Email address in the line {labels}",
        "rule.describe.email_anywhere": "First email address in the text",
        "error.rule.column": "Enter a column name.",
        "error.rule.duplicate": "The column “{column}” already exists. Choose another name.",
        "error.rule.label": "Enter a label, e.g. “Phone:”.",
        "error.rule.between": "Enter a start and an end text.",
        "error.rule.pattern": "Enter a regular expression.",
        "label.field_required": "Required",
        "label.status.running": "Running …",
        "button.run": "Transfer now",
        "button.field_remove": "Delete",
        "button.field_up": "▲ Move up",
        "button.field_down": "▼ Move down",
        "button.browse": "Browse …",
        "dialog.choose_eml_folder": "Choose the folder with the .eml files",
        "view.field.column": "Column",
        "view.field.required": "Required",
        "view.field.type": "Type",
        "view.field.description": "Looks for",
        "view.field.result": "Result",
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
        "summary.cancelled": "Stopped – the next transfer continues from here.",
        "confirm.start_over": (
            "All emails are transferred again into a new Excel file.\n\n"
            "The current Excel file is kept as a backup (name with “_backup_” and the date). "
            "The emails themselves are not changed.\n\nContinue?"
        ),
        "info.backup_created": "Previous Excel file backed up as {file}",
        "info.no_log_yet": "There is no log yet ({file}). It is created by the first run.",
        "error.open_file": "Could not open the file: {file}\n\n{error}",
        "dialog.choose_output_xlsx": "Choose the Excel file for the results",
        "info.no_excel_yet": "There is no Excel file yet ({file}). It is created by the first transfer.",
        "summary.dry_run": (
            "Test run: found {new} new email(s), {ok} without problems and {failed} with problems. "
            "Nothing was saved."
        ),
        "summary.processed": "Added {count} new row(s) to {file}.",
        "summary.failed": "{count} email(s) with problems – see the “{sheet}” sheet.",
        "summary.nothing_new": "No new emails found.",
        "summary.skipped": "Skipped {count} email(s) processed earlier.",
        "error.details": "Details",
        "error.workbook_locked": (
            "The Excel file is open in another program. Close Excel and click “Transfer now” "
            "again."
        ),
        "error.sheet_columns": (
            "The columns in the Excel file no longer match the fields. "
            "Choose a new Excel file under “Settings” or restore the previous fields."
        ),
        "error.mail_folder_missing": (
            "The email folder was not found. Use “Browse …” in the “Emails” tab to choose an existing folder."
        ),
        "error.imap_password_missing": "Enter the mailbox password in the “Emails” tab.",
        "error.imap_login": "The mail server rejected the login. Check username and password.",
        "error.imap_tls": (
            "Could not connect securely to the mail server. Check server, port and “Encrypted connection "
            "(SSL)”."
        ),
        "error.imap_connection": (
            "The mail server cannot be reached. Check the internet connection, mail server and "
            "port."
        ),
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
