# Mail Processor

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

Das Programm hat vier Reiter, in dieser Reihenfolge einzurichten:

1. **E-Mails:** wählen, woher die E-Mails kommen – aus einem **Ordner** mit gespeicherten `.eml`-Dateien
   oder **direkt aus dem Postfach** (IMAP; mit **Verbindung testen** prüfen, das Passwort wird nie gespeichert).
2. **Felder:** eine Zeile pro Excel-Spalte. Unter **Art** wählen, wie der Wert gefunden wird, z. B.
   **Text nach Bezeichnung** mit der Bezeichnung `Telefonnummer:`. Rechts zeigt eine Beispiel-Mail sofort,
   was gefunden wird. Die Standardregeln passen zu [docs/example.eml](docs/example.eml).
3. **Start:** zeigt, ob alles bereit ist. **Jetzt übertragen** klicken; **Testlauf** zeigt vorher, was
   passieren würde, ohne etwas zu speichern. **Excel öffnen** öffnet das Ergebnis. E-Mails mit Problemen
   erscheinen darunter; ein Doppelklick zeigt im Reiter **Felder**, welches Feld nicht gefunden wurde.
4. **Einstellungen:** Excel-Datei, Sprache und Weiteres – meist nicht nötig.

Alle Änderungen werden automatisch gespeichert.

Die Ergebnisse landen in `out/mail_export.xlsx`: Das Blatt **daten** enthält eine Zeile pro E-Mail; die
letzte Spalte **E-Mail-Inhalt** enthält jeweils den vollständigen Text der E-Mail. Das Blatt **fehler** listet
E-Mails, die nicht verarbeitet werden konnten, mit Grund. Einfach erneut übertragen, wenn neue E-Mails
eingehen; nur neue werden hinzugefügt.

Bei Problemen: unter **Einstellungen → Protokoll öffnen** steht, was passiert ist. Das Protokoll enthält
keine E-Mail-Inhalte und kann einer Fehlermeldung beigelegt werden.

## Entwicklung

```bash
uv run pytest
```

Spezifikation: [docs/requirements.md](docs/requirements.md). Regeln für Mitwirkende: [AGENTS.md](AGENTS.md).
Lizenz: [MIT](LICENSE).
