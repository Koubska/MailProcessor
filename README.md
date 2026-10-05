# mailprocessor

Überträgt Angaben aus E-Mails (z. B. Kursanmeldungen über ein Kontaktformular) in eine Excel-Tabelle.
E-Mails werden nur gelesen, nie verändert. Alles läuft lokal, und jede E-Mail wird nur einmal übernommen.

## Installation

**Download:** ZIP-Datei für Ihr System von der
[Releases-Seite](https://github.com/Koubska/MailProcessor/releases/latest) herunterladen, entpacken und
`mailprocessor.exe` (Windows) bzw. `mailprocessor` doppelklicken. Das Programm ist nicht signiert:

- **Windows:** **Weitere Informationen → Trotzdem ausführen** wählen.
- **macOS** (Apple Silicon): Beim ersten Öffnen erscheint eine Warnung. Dann **Systemeinstellungen →
  Datenschutz & Sicherheit** öffnen und bei *mailprocessor* auf **Dennoch öffnen** klicken.

**Aus dem Quellcode** (im Ordner des Repositorys), mit [uv](https://docs.astral.sh/uv/):

```bash
uv run mailprocessor
```

oder mit Python 3.12 oder neuer:

```bash
python3 -m venv .venv && .venv/bin/pip install -q -e . && .venv/bin/mailprocessor          # macOS / Linux
py -3 -m venv .venv; .venv\Scripts\pip install -q -e .; .venv\Scripts\mailprocessor        # Windows
```

## Verwendung

1. Reiter **App-Konfiguration**: Quelle wählen. Entweder **eml** (gespeicherte `.eml`-Dateien im Ordner
   `mails`, oder mit **Durchsuchen …** einen anderen Ordner wählen) oder **imap** (Ihr Postfach; das
   Passwort wird nie gespeichert).
2. Reiter **Parsing-Felder**: eine Regel pro Excel-Spalte. Der Teil des Patterns in `( )` wird übernommen.
   Die Standardregeln passen zu [docs/example.eml](docs/example.eml).
3. Auf **Ausführen** klicken.

Die Ergebnisse landen in `out/mail_export.xlsx`: Das Blatt **daten** enthält eine Zeile pro E-Mail; die
letzte Spalte **E-Mail-Inhalt** enthält jeweils den vollständigen Text der E-Mail. Das Blatt **fehler** listet
E-Mails, die nicht verarbeitet werden konnten, mit Grund. Einfach erneut ausführen, wenn neue E-Mails
eingehen; nur neue werden hinzugefügt.

## Entwicklung

```bash
uv run pytest
```

Spezifikation: [docs/requirements.md](docs/requirements.md). Regeln für Mitwirkende: [AGENTS.md](AGENTS.md).
Lizenz: [MIT](LICENSE).
