# onenote2anytype

Kleine erste Version eines Converters von OneNote-DOCX-Exporten zu einem Anytype-importierbaren ZIP (Markdown + Assets).

## Status

- V1 (MVP)
- Input: `.docx` (einzelne Datei oder ganzer Ordner)
- Output: `.zip` mit `.md` Dateien und `assets/` Bildern

## Features in V1

- Liest Texte und eingebettete Bilder aus DOCX.
- Nutzt die erste sinnvolle Textzeile als Titel.
- Schreibt Markdown-Dateinamen standardmaessig als echten Titel (z. B. `04. Januar 2026.md`).
- Extrahiert `source_created_date` immer aus dem Titelbeginn im Format `dd. Monat yyyy`.
  - Beispiel: `24. Juli 2025 - Urlaub xyz` -> Datum `2025-07-24T12:00:00Z`
- Setzt Uhrzeit immer auf `12:00:00` in der konfigurierten Zeitzone (Default `Europe/Berlin`).
- Uebernimmt einfache Formatierungen aus DOCX:
  - **fett**
  - Aufzaehlungen (Bullet/Nummerierung)
- Ignoriert OneNote-Header-Artefakte am Anfang (`Montag, ...` und `HH:MM`).

## Verwendung

```bash
python converter.py --input "F:\\path\\to\\docx-or-folder" --output "F:\\path\\to\\anytype-import.zip"
```

Wenn `--output` fehlt, wird `anytype-import.zip` im aktuellen Ordner erzeugt.

Optional kannst du die Zeitzone setzen:

```bash
python converter.py --input "F:\\path\\to\\docx-or-folder" --timezone "Europe/Berlin"
```

## Kompatibilitaetsoptionen

Falls ein Import fehlschlaegt, kannst du eine minimalere ZIP-Variante erzeugen:

```bash
python converter.py --input "F:\\path\\to\\docx-or-folder" --output "F:\\path\\to\\anytype-import-compat.zip" --no-frontmatter --zip-root vault
```

- `--no-frontmatter`: schreibt kein YAML-Frontmatter
- `--zip-root vault`: legt alle Dateien unter `vault/` in der ZIP ab

## Ausgabeformat

Die ZIP enthält z. B.:

```text
my-note.md
assets/my-note-image-01.jpeg
assets/my-note-image-02.jpeg
```

Markdown enthält Frontmatter:

```yaml
---
title: "04. Januar 2026"
date: "2026-01-04T12:00:00+01:00"
source_created_date: "2026-01-04T12:00:00+01:00"
source_created_unix: 1767524400
source_format: onenote-docx
---
```

## Hinweis

Diese V1 baut bewusst **Markdown+Assets** für den Anytype-Importpfad.
Das interne Any-Block-Exportformat (`*.pb.json`) wird in V1 nicht generiert.

## Anytype-native Exportstruktur (V2 Prototyp)

Wenn du im Anytype-Importer den Typ `Anytype` verwenden willst, nutze den nativen Converter:

```bash
python converter_anytype.py --input "F:\\path\\to\\docx-or-folder" --template-zip "F:\\path\\to\\Anytype-export.zip" --output "F:\\path\\to\\anytype-native-import.zip"
```

Dieser Modus:
- baut `objects/*.pb.json`, `filesObjects/*.pb.json`, `files/*`
- uebernimmt `relations/`, `types/`, `templates/` aus dem Template-Export
- setzt `details.createdDate` der Seite aus dem Titel-Datum (`12:00` in der gewaehlten Zeitzone)

Hinweis: Eine vollstaendige offizielle Spezifikation dieses Formats ist derzeit nicht oeffentlich dokumentiert. Der Prototyp basiert auf Reverse Engineering eines echten Anytype-Exports.
