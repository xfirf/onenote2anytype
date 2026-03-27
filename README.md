# onenote2anytype

OneNote-DOCX -> Anytype-Import (Import-Typ `Anytype`).

Der aktuelle Fokus liegt auf `converter_anytype.py`.

## Was der Converter macht

- Liest `.docx` als einzelne Datei oder als kompletten Ordner.
- Erzeugt ein ZIP im Anytype-Exportformat mit:
  - `objects/*.pb.json`
  - `filesObjects/*.pb.json`
  - `files/*`
  - `relations/`, `types/`, `templates/` (aus Template-Export)
- Erkennt mehrere Journal-Eintraege in einem Sammel-DOCX:
  - jede Titelzeile `dd. Monat yyyy` startet eine neue Anytype-Seite
- Datumsregel:
  - nur die Titelzeile bestimmt `createdDate`
  - direkt folgende Wochentag-/Uhrzeit-Zeilen werden ignoriert
  - Uhrzeit ist immer `12:00` in der gewaehlten Zeitzone
  - wenn der Dateiname ein Jahr enthaelt (z. B. `2024.docx`) und ein Titel versehentlich ein anderes Jahr hat,
    wird fuer `createdDate` automatisch das Dokumentjahr verwendet
- Uebernimmt Bilder, Fett-Markierungen und einfache Aufzaehlungen.

## Voraussetzungen

- Python 3.11+
- Ein Anytype-Export als Template-ZIP, z. B.:
  - `Anytype-Template.zip`

Der Converter nutzt standardmaessig `Anytype-Template.zip` im aktuellen Ordner.

## Beispiele

Einzelne DOCX konvertieren:

```bash
python converter_anytype.py --input "F:\\dev\\onenote2anynote\\04.docx" --output "F:\\dev\\onenote2anynote\\anytype-native-import.zip"
```

Ordner mit vielen DOCX-Dateien konvertieren:

```bash
python converter_anytype.py --input "F:\\OneNoteExport\\2026" --output "F:\\dev\\onenote2anynote\\anytype-native-import-2026.zip"
```

Mit expliziter Zeitzone:

```bash
python converter_anytype.py --input "F:\\OneNoteExport\\2026" --output "F:\\dev\\onenote2anynote\\anytype-native-import-2026.zip" --timezone "Europe/Berlin"
```

Mit Handschrift-Review-Report:

```bash
python converter_anytype.py --input "F:\\dev\\onenote2anynote\\2025.docx" --output "F:\\dev\\onenote2anynote\\anytype-native-import-2025.zip" --ink-cluster-threshold 40 --manual-review-report "F:\\dev\\onenote2anynote\\manual-review-entries.md"
```

Wenn `--manual-review-report` nicht gesetzt ist, wird automatisch neben der Ausgabe-ZIP eine Datei `<output>-manual-review.md` erstellt (falls verdaechtige Eintraege erkannt wurden).

Dry-Run (nur Liste zum manuellen Nacharbeiten, ohne ZIP-Erzeugung):

```bash
python converter_anytype.py --input "F:\\dev\\onenote2anynote\\2025.docx" --dry-run --ink-cluster-threshold 40 --manual-review-report "F:\\dev\\onenote2anynote\\manual-review-entries.md"
```

Im Dry-Run ist kein Template noetig, weil nur analysiert wird.
Wenn keine verdaechtigen Eintraege gefunden werden, wird im Dry-Run keine Report-Datei angelegt.

Wenn dein Template an einem anderen Ort liegt:

```bash
python converter_anytype.py --input "F:\\OneNoteExport\\2026" --template-zip "F:\\Templates\\Anytype-Template.zip" --output "F:\\dev\\onenote2anynote\\anytype-native-import-2026.zip"
```

## Import in Anytype

1. In Anytype auf `Import` gehen.
2. Import-Typ `Anytype` waehlen.
3. Das erzeugte ZIP auswaehlen.

## Hinweise

- Das interne Anytype-Format ist nicht vollstaendig oeffentlich spezifiziert.
- Der Converter ist deshalb template-basiert (Reverse Engineering aus echtem Export).
- Wenn Titel-Datum und OneNote-Wochentag widersprechen, gilt immer der Titel.

Viel Spass beim Importieren 🚀
