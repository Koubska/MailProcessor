# mailprocessor Requirements

## 1. Purpose

`mailprocessor` reads emails, extracts structured fields from their bodies using configurable regular expressions, and exports the results to an Excel workbook.

The typical input is a stream of form-generated emails (for example, a school's contact form for course registrations, see `docs/example.eml`) that arrive among many unrelated messages.

The emails are confidential. All processing happens locally on the user's computer.

---

# 2. Target environment

* Windows 10/11 (primary), macOS and Linux
* Standalone executable built with PyInstaller; end users do not need Python
* Python 3.12+ for development

Starting the executable without arguments (e.g. double-click) opens the GUI. Starting it with `--config` / `--rules` runs the command-line interface.

---

# 3. Confidentiality

This is a hard requirement.

Email content must never leave the local machine. The only permitted network connection is to the IMAP server the user configured, to read mail.

The application must not use:

* cloud APIs
* AI/LLM APIs (parsing is deterministic regex matching)
* external parsing, OCR or email-processing services
* telemetry, analytics or remote logging
* cloud storage or automatic uploads

Credentials:

* The application never writes the IMAP password to disk. It is taken from the `MAILPROCESSOR_IMAP_PASSWORD` environment variable, an interactive prompt (CLI) or the password field (GUI, kept in memory only).
* IMAP connections verify the server's TLS certificate. Plain connections must be upgraded via STARTTLS; the password is never sent unencrypted.

Attachments are out of scope and are never processed or saved.

---

# 4. Workflow

```text
Mail source (.eml folder or IMAP mailbox)
       ↓
Optional pre-filter (subject, sender, max age)
       ↓
Read message metadata and body locally
       ↓
Skip messages already processed (SQLite ledger)
       ↓
Extract configured fields
       ↓
Successful rows → data worksheet, failures → error worksheet
       ↓
Save workbook, then record processing state
```

A single run both processes and exports. There is no separate scan/export step.

The application never modifies source emails. It must not delete, move, flag, mark as read, or otherwise change messages or files in the source.

---

# 5. Mail sources

Every source converts messages into the application-owned model `NormalizedMail` as early as possible. No other layer may depend on source-specific objects (IMAP responses, file handles, and later Outlook COM objects).

A source yields one item per message: either a `NormalizedMail`, or a `MailReadError` if that single message cannot be read. A source raises only for failures that affect the whole run (unreachable server, failed login, missing folder).

### 5.1 `.eml` folder

* Reads all files matching a glob (default `*.eml`) in one folder, in sorted order. Files are only opened for reading.
* A missing folder is an error, not an empty result.
* The workbook and the ledger must not be located inside the mail folder.

### 5.2 IMAP

* Strictly read-only. Allowed commands: `STARTTLS`, `LOGIN`, `EXAMINE`, `UID SEARCH`, `UID FETCH ... (BODY.PEEK[])`, `LOGOUT`. No `STORE`, `COPY`, `MOVE`, `EXPUNGE`, `CLOSE`, append or delete.
* The server must confirm `[READ-ONLY]` for the mailbox; otherwise the run aborts before any message is fetched.
* Fetches use `BODY.PEEK[]` so the `\Seen` flag is never set, even on servers that ignore read-only mode.
* All server access goes through an allow-list wrapper (`ReadOnlyImapClient`) that exposes no mutating commands.
* The subject and sender filters (`FROM`, `SUBJECT`, ASCII entries only) and the age filter (`SINCE`) are passed to the server, so other mails are not downloaded.

### 5.3 Body extraction

* The first `text/plain` part is used; if there is none, text is extracted from the first `text/html` part.
* Unknown or invalid charsets fall back to UTF-8 instead of failing.
* The body is normalized (line endings, trailing whitespace) before parsing.

---

# 6. Filtering

Supported today:

* **Subject and sender** (`[filter]` in `config.toml`, GUI tab "E-Mails"): a mail is read only if its subject contains one of the `subject` entries and its From header (name or address) contains one of the `sender` entries; case and line breaks are ignored, an empty list matches every mail. Mails left out are not errors: they are counted as "filtered" in the run summary, not written to the ledger (so they are read once the filter changes), and their old rows on the error sheet are removed. Unreadable mails are always reported, because their subject and sender are unknown. The former IMAP-only `source.imap.sender_filter` is read as `filter.sender`.
* **Age**: `max_age_days` / `--max-age-days`. Messages without a valid `Date` header are still processed rather than dropped silently.
* **Volume**: `max_messages` limits how many *new* messages one run handles (already processed messages do not count).

The application should comfortably handle thousands of emails.

---

# 7. Stable message identification

Each message is identified by:

```text
(source_type, source_location, message_identity, content_hash)
```

* `message_identity` is the `Message-ID` header, or, if missing, a hash of date, sender, subject and body (same for all sources, independent of folder, mailbox or IMAP UID).
* `content_hash` is the SHA-256 of the normalized body.
* "Already processed" is decided by `message_identity` + `content_hash` only, so moving folders, moving the application, or renaming a mailbox never causes a second export.

Do not identify messages by sender + subject or sender + timestamp + subject alone.

---

# 8. Parsing

Extraction rules are configured in `parsing_rules.toml` as one or more **profiles**, each a list of fields, one per Excel column. A profile describes one kind of mail (e.g. a registration and a cancellation form). Users describe what to look for; the app generates the regular expression (`rule_patterns.py`):

| `type` | GUI name | Inputs | Value |
|---|---|---|---|
| `label` | Text nach Bezeichnung | `label` | rest of the line after the label |
| `next_line` | Wert in nächster Zeile | `label` | next non-empty line after a line holding only the label |
| `between` | Text zwischen zwei Stellen | `start`, `end` | shortest text between them, also across line breaks |
| `email` | E-Mail-Adresse | `label` (optional) | e-mail address in the label's line; without label the first address in the text |
| `regex` | Experte (Regex) | `pattern` | group 1 of a Python regular expression (whole match without groups) |

```toml
[[profiles]]
name = "Anmeldung"

[[profiles.fields]]
column = "Mail-Adresse"
type = "email"
label = ["Von", "From"]   # alternatives; a single label can be a plain string
required = true

[[profiles.fields]]
column = "Kurs"
type = "label"
label = "Angebot:"
required = true

[[profiles]]
name = "Abmeldung"

[[profiles.fields]]
column = "Grund"
type = "label"
label = "Grund:"
required = true
```

**Choosing the profile.** Every mail is parsed with every profile. Only a profile that finds all its required fields can succeed; of those, the one that finds the most fields (required and optional) wins, and a tie goes to the profile listed first. If no profile succeeds, the mail fails and is reported against the closest profile (fewest required fields missing, then most fields found). With a single profile, behaviour is unchanged.

* A file with only top-level `[[fields]]` (written before profiles existed) is one profile named `Standard`. Using both `[[fields]]` and `[[profiles]]` is an error. The GUI always writes `[[profiles]]`.
* Profile names are unique (ignoring case) and must be valid Excel sheet names (at most 31 characters, none of `[ ] : * ? / \`, no leading or trailing `'`), because they can become sheet names. Column names are unique within a profile; the same column may appear in several profiles.
* Every profile needs at least one field. The GUI keeps a new, still empty profile while editing but saves and runs only profiles with fields.
* In the GUI, the column name suggests the columns of the other profiles (choosing one for a new field copies that profile's rule) and states where the column ends up: shared with other profiles, only this profile's, or, as a warning, a near-duplicate of another profile's column (same name ignoring case, spaces and `-_.:/`) that would become a second column.

* Labels are matched tolerantly: case-insensitive, any spacing, optional colon, at the start of a line, and not as a word prefix ("Tag" does not match "Tagesordnung:"). `label` and `email` never take a value from the next line. Inputs are matched literally (escaped).
* A rule without `type` is a `regex` rule, so files from before the simple types keep working. The GUI writes only the inputs of each rule's type. "Als Regex bearbeiten" converts a simple rule into a `regex` rule; this is one-way.
* Rules are matched against the email text followed by the email's header lines (`Name: value`). Text comes first, so a label in the text wins over a header; headers are the fallback (e.g. the sender in `From:`). Whitespace in the value is collapsed.
* The default rules extract all fields from `docs/example.eml`; `tests/test_default_rules.py` guards this.
* Missing inputs for a type, invalid patterns, duplicate column names and invalid or duplicate profile names are rejected when the rules are loaded; the GUI explains them in plain language.
* A missing required field is a parsing error for that message. Missing values are never treated as valid.
* Parsing is deterministic and independent of the mail source and of Excel.

Prefer the simple types over `regex`, and labels over fragile positions.

---

# 9. Error handling

A single malformed or unreadable email must never abort the batch:

```text
email 1 → success
email 2 → success
email 3 → parse error   (reported in the error worksheet)
email 4 → success
```

Errors are visible to the user (error worksheet, run summary `seen/processed/skipped/failed`). Raw tracebacks are not shown in normal operation; details are available with `log_level = "DEBUG"`.

---

# 10. Processing state and idempotency

SQLite stores the processing history (`sqlite_path`).

* Repeated runs never create duplicate rows for successfully processed messages.
* Failed messages are retried automatically on every run.
* A message whose body changed is processed again (new `content_hash`).
* The workbook is saved first; the ledger is committed only afterwards. A crash, or a workbook locked by Excel, therefore never leaves the ledger ahead of the workbook, and the run can simply be repeated.
* `dry_run` / `--dry-run` parses and reports without writing the workbook or the ledger.
* A run can be stopped (GUI "Stopp"). It stops before the next message and saves everything handled until then, so the next run continues there.
* "Alles neu exportieren" (GUI, advanced settings) renames the current workbook to `<name>_backup_<date>_<time>.xlsx`, clears the ledger and starts a run, so every message is exported again into a fresh workbook. The workbook is moved first; if it is locked by Excel, nothing changes.

---

# 11. Excel output

A single workbook (`output_xlsx`) is appended to across runs. It is in German, like its sheet names.

* **Columns by name:** the app finds every column by its header, never by position. Missing columns are added at the end; nothing is moved or removed. Users can add their own columns (e.g. "Bestätigt") and notes; fields can be added, removed or reordered (a new field gets a new column at the end, a removed field's column stays and new rows leave it empty; the rule order only decides the column order of a new workbook). Only a sheet that has data but no header row stops the run with a clear message.
* **Data worksheet** (default `daten`): one column per configured field, then `Eingegangen am` (the mail's `Date` header in local time, empty if missing or unreadable), `Übertragen am` (time of the run), both as real Excel dates (`TT.MM.JJJJ hh:mm`), and `E-Mail-Inhalt` with the mail's full text (as extracted for parsing; truncated to Excel's limit of 32,767 characters per cell). These three names are reserved and cannot be used for a rule.
* **Several profiles** (`profile_sheets` in `[app]`): `"shared"` (default) writes all rows to the data worksheet; it then starts with a `Profil` column naming the profile each row was read with, followed by the columns of all profiles (each once). Cells of columns a profile does not have stay empty. `"per_profile"` writes each profile's rows to its own sheet, named like the profile, with only that profile's columns; `sheet_data` is then unused, and a profile named like the error sheet stops the run. Renaming a profile in this mode starts a new sheet. With a single profile and `"shared"`, there is no `Profil` column. `Profil` is reserved like the three names above. Older workbooks get the missing columns added; their old rows stay empty there.
* **Error worksheet** (default `fehler`): `E-Mail` (file name or IMAP uid), `Absender`, `Betreff`, `Eingegangen am`, `Fehlende Felder`, `Grund` (plain German, technical details in parentheses), `Geprüft am`, and the hidden `Kennung` that identifies the mail across runs. Exactly one row per currently failing message; the row is removed once the message succeeds. Error sheets of earlier versions (English, technical) are converted automatically.
* **Sheet names** (`sheet_data`, `sheet_errors`) follow the same rules as profile names, because Excel enforces them: at most 31 characters, none of `[ ] : * ? / \`, no leading or trailing spaces or apostrophes, and the two must differ ignoring case (Excel treats `Daten` and `daten` as one sheet).
* Control characters that Excel cannot store are removed from all cell values.
* Bold, frozen header row and autofilter.
* Column widths fitted to the content (between 10 and 50 characters; dates 17; `E-Mail-Inhalt` fixed at 80, not wrapped, so each mail stays one row high). Columns that already have a width, from an earlier run or set by the user, keep it.
* Values extracted from emails are always stored as text, never as formulas (email content is untrusted).
* The workbook is written atomically (temporary file, then replace). A workbook open in Excel is detected before processing starts.

---

# 12. Configuration

`config.toml` holds the app settings (`sqlite_path`, `output_xlsx`, sheet names, `profile_sheets`, `log_level`, `dry_run`, `max_messages`, `max_age_days`) and the source settings (`[source]`, `[source.eml]`, `[source.imap]`).

Relative paths are resolved against the directory containing `config.toml`, so the shipped bundle works regardless of the working directory.

The GUI edits `config.toml` and `parsing_rules.toml`, so normal users do not need to edit files by hand.

---

# 13. GUI

Built with `tkinter` (`gui_app.py`; testable logic in `gui.py`, `ui_model.py`, `preview.py`). Four tabs in the order of setup and use:

* **Start:** status cards for the e-mail source (folder and number of mails, or mailbox and whether the password is entered), the fields (and whether the sample mail is complete) and the Excel file (rows and problems), each with an action ("Ändern", "Bearbeiten", "Excel öffnen"). Large "Jetzt übertragen" button and "Testlauf" (`dry_run`, nothing is written; the `dry_run` value in `config.toml` is kept but ignored by the GUI). While running: progress "Verarbeite E-Mail n von N" and "Stopp". The result in plain language. Mails that failed (also in a test run) are listed with what was not found; "In Beispiel-Mail öffnen" (or a double-click) shows the mail in the Felder tab and selects the first missing field. The log only under "Details anzeigen".
* **E-Mails:** choice between a folder of `.eml` files (folder picker, "Ordner öffnen", number of mails; the file pattern is only configurable in `config.toml`) and the mailbox (IMAP) with explained inputs; the port follows "Verschlüsselte Verbindung (SSL)". "Verbindung testen" logs in and opens the mailbox read-only without reading any message (`imap_source.check_imap_connection`). The inactive choice is greyed out but kept.
* **Felder:** rule list and editor with the rule types from section 8, next to the live test with a sample mail (`preview.py`): the first mail of the configured folder is shown automatically (◀ ▶ steps through the folder); any `.eml` can be loaded, or mail text pasted. The list shows each rule's result (missing required fields in red), the editor shows the result of the rule being edited while typing and highlights the value in the mail text, and a summary says whether the mail would go to the data or the error sheet. Values are computed like in a run, including the header fallback. Sample mails are only read. Changes to the selected field apply as soon as they are valid; new fields are added with "Feld hinzufügen"; deleting asks first.
* **Einstellungen:** Excel file (file picker) and sheet names, language, limits, maintenance ("Alles neu exportieren", "Protokoll öffnen", settings folder) and technical settings (log level, SQLite path, file locations).
* Every change is saved automatically once valid (`ui_model.config_from_form` validates per input); invalid inputs are marked in red next to the field and not saved. Closing with invalid inputs or during a run asks first. The IMAP password is never written to disk.
* Plain-language error messages with a concrete next step; technical details stay visible.
* Hover hints on buttons and inputs; keyboard shortcuts Ctrl/⌘+Enter (Jetzt übertragen), Ctrl/⌘+Shift+Enter (Testlauf), Ctrl/⌘+E (Excel öffnen). Deleting a field needs no confirmation and can be undone. Window size, position and tab are remembered.
* Sharp on scaled Windows screens (125 %, 150 %): the app declares itself DPI aware and scales its pixel-based sizes with the screen.
* German by default, English available.

The pipeline and the connection test run on background threads so the window stays responsive.

---

# 14. Logging

Python's standard logging. Logs stay local: the console/stderr (CLI) or the output box (GUI), plus the log file `logs/mailprocessor.log` next to `config.toml` (rotated at 1 MB, 3 old files kept). The GUI opens it with "Protokoll öffnen" under the advanced settings. The file gets the same records as the console, at the configured `log_level`; if it cannot be created, the app runs without it.

Never log:

* email bodies
* attachments
* passwords or other credentials
* extracted values

---

# 15. Testing

Automated tests cover the business logic without requiring network access or a real mailbox. IMAP is tested with a fake client and, end-to-end, with the real `imaplib` against a scripted local IMAP server that asserts the mailbox is unchanged. A static test forbids mutating calls in `sources/`.

* parser: valid and invalid messages, optional groups, whitespace, line endings
* sources: `.eml` parsing, multipart and HTML bodies, unknown charsets, IMAP read-only behavior, filters, fetch and login failures
* ledger: processed/failed/retry, rollback, read-only mode
* Excel: headers, column order, formula-injection protection, header mismatch, error-row handling, locked workbook
* pipeline/CLI: idempotency, dry run, `max_messages`, error isolation, script entry point

Use only synthetic or anonymized emails. Never commit real email data.

---

# 16. Packaging and distribution

* `scripts/build_executable.py` builds a one-file PyInstaller executable and a ZIP containing the executable, `config.toml`, `parsing_rules.toml`, an empty `mails/` folder, `README.md`, `RUNNING.md` and `LICENSE` (MIT).
* GitHub Actions (`.github/workflows/ci.yml`) runs the tests on Linux, Windows and macOS for every push and pull request. For a pushed tag it then builds the Windows, macOS (Apple Silicon) and Linux ZIPs and creates a GitHub release with the three files attached. Releasing a version: set `version` in `pyproject.toml`, then `git tag -a X.Y -m "X.Y" && git push origin X.Y`.
* The executables are not code-signed; the README explains the SmartScreen/Gatekeeper warnings.

---

# 17. Not yet implemented

The following were part of the original Outlook-focused plan. They are not implemented; add them deliberately, following the rules above:

* **Outlook desktop source** (Windows, `pywin32`/COM, existing Outlook profile, no credentials, read-only, `EntryID` as identity) implemented as another source that yields `NormalizedMail` / `MailReadError`.
* Date-range and unread filters.
* Parser version tracking in the ledger, so that parser changes can deliberately trigger reprocessing.
* Date-typed columns.
* Remembering GUI settings per user (e.g. under `%LOCALAPPDATA%`).
