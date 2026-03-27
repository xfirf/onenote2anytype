#!/usr/bin/env python3
from __future__ import annotations

import argparse
import posixpath
import re
import sys
import zipfile
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Iterable
from xml.etree import ElementTree as ET
from zoneinfo import ZoneInfo, ZoneInfoNotFoundError


W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PR_NS = "http://schemas.openxmlformats.org/package/2006/relationships"

MONTH_MAP = {
    "januar": 1,
    "februar": 2,
    "maerz": 3,
    "märz": 3,
    "april": 4,
    "mai": 5,
    "juni": 6,
    "juli": 7,
    "august": 8,
    "september": 9,
    "oktober": 10,
    "november": 11,
    "dezember": 12,
}

WEEKDAY_RE = re.compile(
    r"^(Montag|Dienstag|Mittwoch|Donnerstag|Freitag|Samstag|Sonntag),\s+\d{1,2}\.\s+[A-Za-zÄÖÜäöüß]+\s+\d{4}$",
    re.IGNORECASE,
)
WEEKDAY_EN_RE = re.compile(
    r"^(Monday|Tuesday|Wednesday|Thursday|Friday|Saturday|Sunday),\s+[A-Za-z]+\s+\d{1,2},\s*\d{4}$",
    re.IGNORECASE,
)
TIME_RE = re.compile(r"^\d{1,2}:\d{2}$")
TIME_AMPM_RE = re.compile(r"^\d{1,2}:\d{2}\s*[AP]M$", re.IGNORECASE)


@dataclass
class TextElement:
    plain_text: str
    markdown_text: str
    list_type: str | None = None
    list_level: int = 0


@dataclass
class ImageElement:
    source_path: str


Element = TextElement | ImageElement


def normalize_month_name(raw: str) -> str:
    value = raw.strip().lower()
    return (
        value.replace("ä", "ae")
        .replace("ö", "oe")
        .replace("ü", "ue")
        .replace("ß", "ss")
    )


def parse_created_datetime_from_title(title: str, timezone_name: str) -> datetime:
    match = re.match(r"^\s*(\d{1,2})\.\s*([A-Za-zÄÖÜäöüß]+)\s+(\d{4})", title)
    if not match:
        raise ValueError(
            f"Titel startet nicht mit Datumsmuster 'dd. Monat yyyy': {title!r}"
        )

    day = int(match.group(1))
    month_key_raw = match.group(2)
    year = int(match.group(3))

    month = MONTH_MAP.get(month_key_raw.lower())
    if month is None:
        month = MONTH_MAP.get(normalize_month_name(month_key_raw))
    if month is None:
        raise ValueError(f"Unbekannter Monatsname im Titel: {month_key_raw!r}")

    try:
        tz = ZoneInfo(timezone_name)
    except ZoneInfoNotFoundError as exc:
        raise ValueError(f"Unbekannte Zeitzone: {timezone_name!r}") from exc

    return datetime(year, month, day, 12, 0, 0, tzinfo=tz)


def slugify(value: str) -> str:
    value = normalize_month_name(value.lower())
    value = re.sub(r"[^a-z0-9]+", "-", value)
    value = value.strip("-")
    return value or "note"


def filename_from_title(title: str) -> str:
    """Create a human-readable filename from title.

    Keeps spaces/umlauts so importers can derive the object title from filename.
    """
    cleaned = title.strip()
    # Replace only characters invalid in common filesystems/zip tooling.
    cleaned = re.sub(r'[\\/:*?"<>|]+', "-", cleaned)
    cleaned = re.sub(r"\s+", " ", cleaned)
    cleaned = cleaned.strip(" .")
    return cleaned or "Notiz"


def discover_docx_files(input_path: Path) -> list[Path]:
    if input_path.is_file():
        if input_path.suffix.lower() != ".docx":
            raise ValueError(f"Input-Datei ist keine .docx: {input_path}")
        return [input_path]

    if not input_path.is_dir():
        raise ValueError(f"Input-Pfad existiert nicht: {input_path}")

    files = sorted(p for p in input_path.rglob("*.docx") if p.is_file())
    if not files:
        raise ValueError(f"Keine .docx Dateien gefunden in: {input_path}")
    return files


def parse_relationships(docx: zipfile.ZipFile) -> dict[str, str]:
    rel_xml = docx.read("word/_rels/document.xml.rels")
    root = ET.fromstring(rel_xml)
    rel_map: dict[str, str] = {}

    for rel in root.findall(f"{{{PR_NS}}}Relationship"):
        rel_id = rel.attrib.get("Id")
        target = rel.attrib.get("Target")
        rel_type = rel.attrib.get("Type", "")
        if rel_id and target and rel_type.endswith("/image"):
            rel_map[rel_id] = target

    return rel_map


def parse_numbering_map(docx: zipfile.ZipFile) -> dict[tuple[str, str], str]:
    try:
        numbering_xml = docx.read("word/numbering.xml")
    except KeyError:
        return {}

    root = ET.fromstring(numbering_xml)
    abstract_by_num: dict[str, str] = {}
    level_fmt_by_abstract: dict[tuple[str, str], str] = {}

    for num in root.findall(f"{{{W_NS}}}num"):
        num_id = num.attrib.get(f"{{{W_NS}}}numId")
        abstract = num.find(f"{{{W_NS}}}abstractNumId")
        if not num_id or abstract is None:
            continue
        abstract_id = abstract.attrib.get(f"{{{W_NS}}}val")
        if abstract_id:
            abstract_by_num[num_id] = abstract_id

    for abstract in root.findall(f"{{{W_NS}}}abstractNum"):
        abstract_id = abstract.attrib.get(f"{{{W_NS}}}abstractNumId")
        if not abstract_id:
            continue
        for lvl in abstract.findall(f"{{{W_NS}}}lvl"):
            ilvl = lvl.attrib.get(f"{{{W_NS}}}ilvl", "0")
            num_fmt = lvl.find(f"{{{W_NS}}}numFmt")
            if num_fmt is None:
                continue
            fmt = num_fmt.attrib.get(f"{{{W_NS}}}val")
            if fmt:
                level_fmt_by_abstract[(abstract_id, ilvl)] = fmt

    out: dict[tuple[str, str], str] = {}
    for num_id, abstract_id in abstract_by_num.items():
        for (abs_id, ilvl), fmt in level_fmt_by_abstract.items():
            if abs_id == abstract_id:
                out[(num_id, ilvl)] = fmt
    return out


def run_is_bold(run: ET.Element) -> bool:
    run_props = run.find(f"{{{W_NS}}}rPr")
    if run_props is None:
        return False
    b = run_props.find(f"{{{W_NS}}}b")
    if b is None:
        return False
    val = b.attrib.get(f"{{{W_NS}}}val", "1")
    return val != "0"


def markdown_from_runs(paragraph: ET.Element) -> tuple[str, str]:
    plain_parts: list[str] = []
    markdown_parts: list[str] = []

    for run in paragraph.findall(f"{{{W_NS}}}r"):
        run_text_parts = [
            (t.text or "") for t in run.findall(f"{{{W_NS}}}t") if t.text is not None
        ]
        if not run_text_parts:
            continue
        run_text = "".join(run_text_parts).replace("\xa0", " ")
        if not run_text:
            continue

        plain_parts.append(run_text)
        if run_is_bold(run):
            markdown_parts.append(f"**{run_text}**")
        else:
            markdown_parts.append(run_text)

    plain_text = "".join(plain_parts).strip()
    markdown_text = "".join(markdown_parts).strip()
    return plain_text, markdown_text


def parse_elements(docx: zipfile.ZipFile) -> list[Element]:
    rel_map = parse_relationships(docx)
    numbering_map = parse_numbering_map(docx)
    xml = docx.read("word/document.xml")
    root = ET.fromstring(xml)

    body = root.find(f"{{{W_NS}}}body")
    if body is None:
        return []

    elements: list[Element] = []
    for p in body.findall(f"{{{W_NS}}}p"):
        plain_text, markdown_text = markdown_from_runs(p)

        list_type: str | None = None
        list_level = 0
        ppr = p.find(f"{{{W_NS}}}pPr")
        if ppr is not None:
            num_pr = ppr.find(f"{{{W_NS}}}numPr")
            if num_pr is not None:
                num_id_el = num_pr.find(f"{{{W_NS}}}numId")
                ilvl_el = num_pr.find(f"{{{W_NS}}}ilvl")
                if num_id_el is not None:
                    num_id = num_id_el.attrib.get(f"{{{W_NS}}}val")
                    ilvl = ilvl_el.attrib.get(f"{{{W_NS}}}val", "0") if ilvl_el is not None else "0"
                    if num_id:
                        list_level = int(ilvl) if ilvl.isdigit() else 0
                        fmt = numbering_map.get((num_id, ilvl), "bullet")
                        if fmt in {"decimal", "upperRoman", "lowerRoman", "upperLetter", "lowerLetter"}:
                            list_type = "ordered"
                        else:
                            list_type = "bullet"

        if plain_text:
            elements.append(
                TextElement(
                    plain_text=plain_text,
                    markdown_text=markdown_text,
                    list_type=list_type,
                    list_level=list_level,
                )
            )

        for blip in p.findall(f".//{{{A_NS}}}blip"):
            rel_id = blip.attrib.get(f"{{{R_NS}}}embed")
            if not rel_id:
                continue

            rel_target = rel_map.get(rel_id)
            if not rel_target:
                continue

            normalized = posixpath.normpath(posixpath.join("word", rel_target))
            if normalized.startswith("../"):
                continue
            elements.append(ImageElement(source_path=normalized))

    return elements


def extract_title(elements: Iterable[Element]) -> str:
    for element in elements:
        if isinstance(element, TextElement) and element.plain_text.strip():
            return element.plain_text.strip()
    raise ValueError("Keine Titelzeile gefunden (keine Textabsätze in DOCX).")


def should_skip_header_artifact(text: str) -> bool:
    return bool(
        WEEKDAY_RE.match(text)
        or WEEKDAY_EN_RE.match(text)
        or TIME_RE.match(text)
        or TIME_AMPM_RE.match(text)
    )


def markdown_from_elements(
    title: str,
    created_iso: str,
    created_unix: int,
    elements: list[Element],
    image_name_map: dict[str, str],
    include_frontmatter: bool,
) -> str:
    if include_frontmatter:
        escaped_title = title.replace('"', '\\"')
        lines: list[str] = [
            "---",
            f'title: "{escaped_title}"',
            f'date: "{created_iso}"',
            f'created: "{created_iso}"',
            f'createdDate: "{created_iso}"',
            f'source_created_date: "{created_iso}"',
            f"source_created_unix: {created_unix}",
            "source_format: onenote-docx",
            "---",
            "",
        ]
    else:
        lines = [f"# {title}", ""]

    title_consumed = False
    content_started = False
    previous_was_list = False

    for element in elements:
        if isinstance(element, TextElement):
            plain_text = element.plain_text.strip()
            markdown_text = element.markdown_text.strip()
            if not plain_text:
                continue

            if not title_consumed and plain_text == title:
                title_consumed = True
                continue

            if not content_started and should_skip_header_artifact(plain_text):
                continue

            if element.list_type:
                if lines and lines[-1] != "" and not previous_was_list:
                    lines.append("")
                indent = "  " * element.list_level
                marker = "1." if element.list_type == "ordered" else "-"
                lines.append(f"{indent}{marker} {markdown_text}")
                previous_was_list = True
            else:
                if previous_was_list:
                    lines.append("")
                lines.append(markdown_text)
                lines.append("")
                previous_was_list = False
            content_started = True
            continue

        if isinstance(element, ImageElement):
            asset_name = image_name_map.get(element.source_path)
            if not asset_name:
                continue
            if previous_was_list:
                lines.append("")
                previous_was_list = False
            lines.append(f"![Bild](assets/{asset_name})")
            lines.append("")
            content_started = True

    # Prevent trailing blank lines from growing forever.
    while lines and lines[-1] == "":
        lines.pop()

    return "\n".join(lines) + "\n"


def build_zip_path(name: str, zip_root: str) -> str:
    root = zip_root.strip().strip("/")
    if not root:
        return name
    return f"{root}/{name}"


def convert_docx_files(
    docx_files: list[Path],
    output_zip: Path,
    include_frontmatter: bool,
    zip_root: str,
    timezone_name: str,
) -> None:
    output_zip.parent.mkdir(parents=True, exist_ok=True)

    used_markdown_names: set[str] = set()

    with zipfile.ZipFile(output_zip, mode="w", compression=zipfile.ZIP_DEFLATED) as out_zip:
        for docx_path in docx_files:
            with zipfile.ZipFile(docx_path, mode="r") as docx_zip:
                elements = parse_elements(docx_zip)
                if not elements:
                    raise ValueError(f"Keine Inhalte in DOCX gefunden: {docx_path}")

                title = extract_title(elements)
                created_dt = parse_created_datetime_from_title(title, timezone_name)
                created_iso = created_dt.isoformat()
                created_unix = int(created_dt.timestamp())

                base = filename_from_title(title)
                markdown_name = f"{base}.md"
                i = 2
                while markdown_name in used_markdown_names:
                    markdown_name = f"{base}-{i}.md"
                    i += 1
                used_markdown_names.add(markdown_name)

                asset_base = slugify(title)

                source_images: list[str] = []
                for element in elements:
                    if isinstance(element, ImageElement) and element.source_path not in source_images:
                        source_images.append(element.source_path)

                image_name_map: dict[str, str] = {}
                for idx, source in enumerate(source_images, start=1):
                    ext = Path(source).suffix.lower() or ".bin"
                    asset_name = f"{asset_base}-image-{idx:02d}{ext}"
                    image_name_map[source] = asset_name

                    try:
                        data = docx_zip.read(source)
                    except KeyError:
                        continue
                    out_zip.writestr(build_zip_path(f"assets/{asset_name}", zip_root), data)

                markdown = markdown_from_elements(
                    title,
                    created_iso,
                    created_unix,
                    elements,
                    image_name_map,
                    include_frontmatter=include_frontmatter,
                )
                out_zip.writestr(
                    build_zip_path(markdown_name, zip_root),
                    markdown.encode("utf-8"),
                )

                print(f"OK: {docx_path.name} -> {markdown_name}")


def parse_args(argv: list[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Konvertiert OneNote DOCX-Exporte in ein Anytype-kompatibles Markdown+Assets ZIP."
    )
    parser.add_argument(
        "--input",
        required=True,
        help="Pfad zu einer .docx Datei oder zu einem Ordner mit .docx Dateien",
    )
    parser.add_argument(
        "--output",
        default="anytype-import.zip",
        help="Pfad der erzeugten ZIP-Datei (Default: anytype-import.zip)",
    )
    parser.add_argument(
        "--no-frontmatter",
        action="store_true",
        help="Frontmatter nicht schreiben (Kompatibilitaetsmodus)",
    )
    parser.add_argument(
        "--zip-root",
        default="",
        help="Optionaler Root-Ordner innerhalb der ZIP (z. B. vault)",
    )
    parser.add_argument(
        "--timezone",
        default="Europe/Berlin",
        help="Zeitzone fuer den 12:00 Timestamp (Default: Europe/Berlin)",
    )
    return parser.parse_args(argv)


def main(argv: list[str]) -> int:
    try:
        args = parse_args(argv)
        input_path = Path(args.input).expanduser().resolve()
        output_zip = Path(args.output).expanduser().resolve()

        docx_files = discover_docx_files(input_path)
        convert_docx_files(
            docx_files,
            output_zip,
            include_frontmatter=not args.no_frontmatter,
            zip_root=args.zip_root,
            timezone_name=args.timezone,
        )
        print(f"Fertig: {output_zip}")
        return 0
    except Exception as exc:  # noqa: BLE001
        print(f"Fehler: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
