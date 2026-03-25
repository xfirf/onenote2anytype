# Anytype Export Format (Analyse)

## Quelle

- Datei: `Anytype.20260325.193715.65.zip`
- Typ: Any-Block Export (nicht Markdown-Import-ZIP)

## ZIP-Struktur

- `objects/*.pb.json` (Pages/Participant)
- `filesObjects/*.pb.json` (FileObject-Metadaten)
- `files/*` (binäre Dateien, z. B. Bilder)
- `relations/*.pb.json` (STRelation)
- `types/*.pb.json` (STType)
- `templates/*.pb.json`

## Wichtige technische Beobachtungen

- Seiteninhalt liegt in `snapshot.data.blocks`.
- Objekt-Metadaten liegen in `snapshot.data.details`.
- `details.createdDate` ist Unix-Timestamp in Sekunden.
- Bilder sind über `file.targetObjectId` -> `filesObjects/<id>.pb.json` -> `details.source` verknüpft.
- IDs sind interne `bafy...`-Objekt-IDs und stark vernetzt.

## Konsequenz für Projekt

- V1 setzt auf Markdown+Assets-ZIP (robuster, weniger fragil).
- Any-Block-Generator bleibt optionales späteres Advanced-Feature.
