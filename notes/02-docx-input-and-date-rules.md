# DOCX Input und Datumsregeln

## Quelle

- Datei: `04.docx`
- Enthält 2 Bilder:
  - `word/media/image1.jpeg`
  - `word/media/image2.jpeg`

## Titel- und Datumsregel (verbindlich)

- Created-Datum wird immer aus dem Notiztitel gelesen.
- Beispiele:
  - `04. Januar 2026`
  - `24. Juli 2025 - Urlaub xyz` -> Datumsteil ist `24. Juli 2025`
- Alles nach dem Datum wird ignoriert.

## Parserregel

- Regex am Titelanfang:
  - `^\s*(\d{1,2})\.\s*([A-Za-zÄÖÜäöüß]+)\s+(\d{4})`
- Monatsmapping Deutsch -> Zahl.
- Uhrzeit wird fix auf `12:00:00` gesetzt.
- Ausgabeformat für Timestamp: RFC3339 (`YYYY-MM-DDTHH:MM:SSZ`).

## Wichtig

- DOCX-Metadaten (`docProps/core.xml`) werden für createdDate nicht verwendet.
