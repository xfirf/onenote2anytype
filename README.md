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
- Erkennt auch alte Titel-Formate und normalisiert sie auf Deutsch:
  - `Month dd, yyyy`, `Month dd yyyy`, `... at hh:mmAM/PM`
  - Ausgabe-Titel immer als `dd. Monat yyyy` (plus optionale Erweiterung)
- Datumsregel:
  - wenn der Dateiname mit `YYYY-MM-DD_HH-mm` beginnt, ist das primaere Datum/Uhrzeit fuer `createdDate`
  - wenn kein passender Dateiname vorliegt, wird wie bisher aus Titel/Wochentag/Uhrzeit im Dokument aufgeloest
  - wenn Wochentag/Titel fehlen oder widersprechen, faellt der Converter auf `12:00` fuer den erkannten Tag zurueck
  - bei alten Sammel-DOCX ohne Dateiname-Praefix bleibt das bisherige Titel-Verhalten aktiv
- Uebernimmt Bilder, Fett-Markierungen und einfache Aufzaehlungen.
- Defekte/leer exportierte DOCX werden mit Warnung uebersprungen, damit der Rest weiter konvertiert.

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

## OneNote per Microsoft Graph exportieren (Seite fuer Seite)

Wenn OneNote-Sammel-Exporte kaputte DOCX erzeugen, kannst du Seiten einzeln per Graph ziehen:

1. Azure App als Public Client anlegen (delegated Permission: `Notes.Read`).
2. Abhaengigkeiten installieren:

```bash
pip install msal requests
```

3. Optional Pandoc installieren, wenn direkt `.docx` erzeugt werden soll.
4. Script starten (Beispiel fuer dein Notebook/Abschnitte):

```bash
python export_onenote_graph.py --client-id "<DEINE-APP-ID>" --notebook "Tagebücher" --sections "1981-2016" "2017-2018" --output "F:\OneNoteGraphExport" --convert-docx
```

Ergebnis:

- HTML je Seite unter `<output>/<Notebook>/<Section>/_html/`
- optional DOCX je Seite unter `<output>/<Notebook>/<Section>/`

Danach kannst du wie gewohnt den Anytype-Converter auf den Export-Ordner laufen lassen.

## OneNote ohne Azure/Graph (lokal per COM)

Wenn du keinen Azure/Graph-Setup willst, kannst du direkt die lokale OneNote-Desktop-App nutzen.
Das Script exportiert Seiten einzeln als DOCX ueber die OneNote COM-Schnittstelle.

Voraussetzungen:

- Windows mit installierter OneNote-Desktop-App (Microsoft 365 / OneNote 2016)
- Notebook und Abschnitte sind in OneNote geoeffnet/synchronisiert

Beispiel (dein Fall):

```powershell
powershell -ExecutionPolicy Bypass -File .\export_onenote_com.ps1 -Notebook "Tagebücher" -Sections "1981-2016","2017-2018" -Output "F:\OneNoteComExport"
```

Hinweis: Abschnittsnamen mit und ohne Leerzeichen um den Bindestrich werden akzeptiert
(`1981-2016` und `1981 - 2016` funktionieren beide).

Nutzliche Optionen:

- Nur pruefen/listen, ohne Export: `-ListOnly`
- Seitenlimit pro Abschnitt: `-LimitPages 20`
- Bei Fehlern sofort abbrechen (statt weiterlaufen): `-StopOnError`

Ergebnis:

- DOCX je Seite unter `<output>/<Notebook>/<Section>/`
- Vollreport mit allen Seiten unter `<output>/<Notebook>/_export-report.json`
- Nur Fehlseiten unter `<output>/<Notebook>/_export-failures.json` und `<output>/<Notebook>/_export-failures.csv`
- Kurzsummary unter `<output>/<Notebook>/_export-summary.txt`

Standardverhalten: Das Script loggt Fehler pro Seite und macht mit der naechsten Seite weiter.
