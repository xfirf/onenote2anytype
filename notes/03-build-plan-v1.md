# Build-Plan V1 (OneNote DOCX -> Anytype ZIP)

## Ziel

OneNote-DOCX in ein Anytype-importierbares ZIP mit Markdown und Assets transformieren.

## Pipeline

1. DOCX öffnen (ZIP-basiert).
2. Titel aus erster sinnvollen Zeile lesen.
3. Created-Date aus Titel extrahieren (12:00:00 fix).
4. Textblöcke in Markdown umwandeln.
5. Bilder aus `word/media/*` extrahieren.
6. Bildreferenzen in Markdown setzen (relative Pfade).
7. Frontmatter schreiben (`title`, `source_created_date`, `source_format`).
8. Alle Notes und Assets als ZIP bündeln.

## V1-Outputstruktur

- `<note-slug>.md`
- `assets/<note-slug>-image-01.jpeg`
- `assets/<note-slug>-image-02.jpeg`

## V1-Akzeptanzkriterien

- Datum wird korrekt aus Titel geparst (inkl. Suffix-Fall).
- Bilder erscheinen nach Import in Anytype.
- UTF-8 Inhalte bleiben korrekt erhalten.
- Mehrere DOCX-Dateien werden batchweise verarbeitet.
