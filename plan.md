# Plan: 2019/2020 Titel-Normalisierung + Timestamp-Regeln

## Ziel
OneNote-DOCX Einträge konsistent in Anytype importieren, auch bei alten Titel-Formaten.
Ausgabe-Titel soll immer im Format:
`DD. Monat YYYY` (optional ` - <voller Originaltitel>`)

## Festgelegte Regeln

1. **Canonical Titel**
   - Standardausgabe: `DD. <deutscher Monat> YYYY`
   - Wenn Zusatz im Titel vorhanden: `DD. <deutscher Monat> YYYY - <Zusatz>`
   - Bei datumslosen Legacy-Titeln (z. B. `Israel 2019 - Tag 04`):
     - Datum aus der darunterliegenden Datum/Wochentag-Zeile nehmen
     - voller Originaltitel als Suffix anhängen
     - Beispiel: `13. März 2019 - Israel 2019 - Tag 04`

2. **Alte Titel-Formate erkennen**
   - `dd. Monat yyyy`
   - `Month dd, yyyy`
   - `Month dd yyyy`
   - optionales `at hh:mmAM/PM`
   - tolerante Month-Tippfehler (z. B. `Ocotober` -> `October`) über Alias/Fuzzy-Korrektur

3. **Timestamp-Regel**
   - Titel-Datum und Datum aus der darunterliegenden Header-Datumzeile vergleichen
   - Wenn **gleicher Kalendertag**: Uhrzeit aus Zeitzeile übernehmen
   - Sonst: Uhrzeit auf `12:00`
   - Zeitzone bleibt konfigurierbar (`--timezone`)

4. **Jahres-Tippfehlerregel bleibt**
   - Enthält Dateiname ein Jahr (z. B. `2024.docx`) und Titel hat anderes Jahr:
     - Jahr auf Dokumentjahr korrigieren (falls Datum valide)

## Robuste Section-Erkennung

Statt nur „Titel matcht Datum“:
- Eintragsstart über Header-Muster:
  - `Titelzeile`
  - `Wochentag+Datum` (de/en)
  - `Uhrzeit` (24h oder AM/PM)
- Fallback auf bisherigen Split, falls Muster nicht greift

## Umsetzungsplan (Code)

1. Parser-Layer erweitern
   - `parse_title_date(...)` für de/en + legacy Varianten
   - `parse_weekday_date(...)` für de/en
   - `parse_time_value(...)` inkl. AM/PM
   - Month-Alias/Fuzzy für bekannte Tippfehler

2. Titel normalisieren
   - `normalize_title(...)` erzeugt deutsches Canonical-Format
   - datumsloser Titel + Header-Datum => Canonical + ` - <voller Originaltitel>`

3. Timestamp-Resolver
   - `resolve_time_from_section(...)` gemäß Same-Day-Regel
   - Integration in `resolve_created_datetime(...)`

4. Section-Split verbessern
   - Header-Triplet-basierter Split
   - kompatibler Fallback

5. Type-/Import-Stabilität sicherstellen
   - bestehende Fixes beibehalten:
     - Template Type (`Tagebuch`)
     - source path mit `/`
     - multi-image support
     - dry-run report behavior

## Testmatrix

### A) 2019 Samples (`*2019.docx`)
- `Month dd, yyyy`
- `Month dd yyyy`
- `... at hh:mmAM/PM`
- `Israel 2019 - Tag xx`
- Tippfehlermonat (`Ocotober`)

### B) 2020 Sample (`2020.docx`)
- gemischte de/en Titel
- doppelte Jahresangabe (`04. Januar 2020 2020`)
- mismatched weekday-date vs title-date
- Zeitübernahme nur bei Same-Day

### C) Aktuelle Dateien
- `2025.docx` / `2026.docx`
- prüfen, dass bestehendes Verhalten nicht regressiert

## Dry-Run / Reporting

Optional erweitern:
- Report kann zusätzlich ausgeben:
  - original_title
  - normalized_title
  - timestamp_source (`header-time` / `fallback-12:00`)
  - year_corrected (bool)

## Offener Mini-Bugfix

Ein separater „kleiner Bug“ wird nach Context-Reset als erster Punkt umgesetzt:
- [ ] Bug reproduzieren
- [ ] Minimalfix
- [ ] Kurztest + Commit

## Reihenfolge nach Context-Reset

1. kleinen Bug fixen
2. Titel/Timestamp-Normalisierung für 2019/2020 implementieren
3. Tests über 2019/2020/2025/2026
4. Review mit dir
5. commit + push
