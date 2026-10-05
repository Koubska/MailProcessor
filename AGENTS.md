# mailprocessor — Coding Agent Instructions

## Project

This repository contains a small desktop application that reads emails from a local `.eml` folder or an IMAP mailbox, extracts structured data from their bodies using configurable regular expressions, and exports the results to Excel. It ships as a standalone executable for Windows (primary), macOS and Linux: started without arguments it opens a `tkinter` GUI, and with `--config`/`--rules` it runs as a CLI.

The application handles confidential email.

The primary goals are:

* local-only processing
* reliability and correctness of the exported data
* deterministic behavior
* simple operation
* minimal configuration
* maintainable code
* easy deployment

Read `docs/requirements.md` before implementing features. Section 17 lists planned but unimplemented features (including an Outlook source).

---

## Non-negotiable constraints

### Confidentiality

Email data must never leave the user's computer. The only permitted network connection is to the user-configured IMAP server.

Never introduce:

* cloud APIs
* AI/LLM APIs
* external parsing services
* telemetry
* analytics
* remote logging
* automatic uploading
* external OCR
* external email-processing services

Do not send email bodies, email metadata, extracted values, or attachments to external systems.

Do not use an LLM for email parsing.

Never write credentials to disk. The IMAP password comes from `MAILPROCESSOR_IMAP_PASSWORD`, an interactive prompt, or the GUI's in-memory password field.

Never weaken transport security: IMAP must verify TLS certificates, and plain connections must use STARTTLS.

Never log complete email bodies.

Never put real confidential emails into the repository or automated tests.

Only use synthetic/anonymized test data.

### Read-only sources

Sources are read-only. Do not modify, move, delete, flag, mark as read, or otherwise alter emails or source files unless the requirements explicitly request it.

The IMAP source may only use `STARTTLS`, `LOGIN`, `EXAMINE` (requiring the server's `[READ-ONLY]` confirmation), `UID SEARCH`, `UID FETCH` with `BODY.PEEK[]` (never `RFC822`/`BODY[]`, which set `\Seen`) and `LOGOUT`, all through `ReadOnlyImapClient`. Never add generic command access to that wrapper.

`tests/test_imap_readonly_integration.py` and `tests/test_sources_readonly_guard.py` enforce this. Do not weaken them.

### Local architecture

Expected local persistence: SQLite.

Expected local output: Excel `.xlsx`.

No server or database service is required.

---

## Technology

Current stack:

* Python 3.12+
* `tkinter` for the desktop UI
* `openpyxl` for Excel
* Python's built-in `sqlite3`, `imaplib`, `email`
* `pydantic` for config validation, `typer` for the CLI
* PyInstaller for packaging

Do not add dependencies without a concrete reason.

Before adding a dependency, check whether the standard library or an existing dependency can solve the problem adequately. (`pywin32` becomes justified only when an Outlook source is implemented.)

Do not introduce frameworks or infrastructure that are unnecessary for this application.

---

## Architecture

Keep these concerns separated:

```text
Source (.eml folder / IMAP)
   ↓
NormalizedMail | MailReadError
   ↓
Parser
   ↓
ParseResult
   ↓
SQLite ledger / Excel workbook
```

The parser must not know about sources or Excel.

The Excel writer must not know about sources.

The GUI must not contain parsing logic.

Source-specific objects (IMAP responses, file handles, future Outlook COM objects) must not leak out of `sources/`. Convert them into `NormalizedMail` as early as possible.

Prefer simple classes, dataclasses, and functions over elaborate abstractions.

---

## Parser

Email parsing is business-critical.

Parsing must be:

* deterministic
* testable
* independent of sources
* independent of Excel
* covered by automated tests

Extraction rules are user configuration (`parsing_rules.toml`), not code. Do not invent or hard-code business extraction rules.

When an email format is ambiguous, inspect the provided fixtures (`docs/example.eml`) and requirements before making assumptions.

If an ambiguity genuinely prevents correct implementation, ask for clarification.

Do not silently guess.

---

## Mail sources

A source yields one item per message: `NormalizedMail`, or `MailReadError` if that single message cannot be read.

Raise only for failures that affect the whole run (connection, login, missing folder), as `OSError`/`ValueError` with a message that is useful to the user and contains no credentials.

New sources (e.g. Outlook) must follow the same contract, be read-only, and use a stable message identifier (e.g. Outlook `EntryID`).

---

## Processing state

Use SQLite for local processing history.

Processing must be idempotent.

Running the application repeatedly must not create duplicate Excel rows for the same email.

Use a stable message identifier (`Message-ID`, `EntryID`, …) rather than sender/subject/date combinations.

Failed messages are retried on later runs.

Save the workbook first, then commit the ledger. Never let the ledger record a message whose row is not on disk.

`dry_run` must not write anything.

Parser version tracking is not implemented yet; add it deliberately when parser changes require reprocessing.

---

## GUI

Use `tkinter`.

Keep the UI simple and practical.

The user should be able to:

1. configure the source (`.eml` folder or IMAP)
2. configure filters and limits
3. edit the parsing fields
4. run processing
5. see processing statistics and errors

Long-running work must not block the tkinter UI thread. Run it on a background worker and pass results back through the event loop (`root.after`). Never touch tkinter objects from the worker thread.

All UI strings go through `i18n.t()`. German is the default, and every key must exist in both catalogs.

Do not create a complicated multi-page UI unless there is a clear usability reason.

---

## Excel

Use `openpyxl`.

Generated workbooks should be useful without manual cleanup.

Prefer:

* header row
* autofilter
* frozen header
* sensible column widths
* sensible date formatting
* separate error worksheet

Email content is untrusted: always store extracted values as text, never as formulas.

Do not expose internal implementation details such as database IDs unless useful to the user.

Do not overwrite existing files without explicit handling. Never append rows under a header that does not match the configured columns.

---

## Error handling

A single malformed email must never abort an entire batch.

Prefer:

```text
email 1 → success
email 2 → success
email 3 → parse error
email 4 → success
```

rather than failing the entire operation.

Errors must be visible to the user and useful for troubleshooting.

Do not display raw stack traces in the normal UI or CLI output.

Keep detailed technical diagnostics in local logs.

---

## Logging

Use Python's standard logging facilities.

Logs must remain local.

Never log:

* complete email bodies
* attachments
* credentials
* extracted values or other unnecessary confidential data

---

## Changes and pull requests

`main` is protected. Every change reaches it through a pull request; nobody pushes to `main` directly.

* Start each change on a new branch from the current `main` (`feature/…`, `fix/…`, `docs/…`, `chore/…`).
* Keep a pull request to one topic. Update tests and documentation in the same pull request.
* A pull request is merged only when CI is green (tests on Ubuntu, Windows and macOS, and lint). Merges are squash merges.
* Small Dependabot updates are merged automatically by CI; major updates need a review.
* Releases: tag the merged commit on `main` (e.g. `3.5`). The tag is the version; CI writes it into the build.

