#!/usr/bin/env python3
from __future__ import annotations

import argparse
import copy
import hashlib
import json
import os
import re
import struct
import sys
import zipfile
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Iterable

from converter import (
    ImageElement,
    TextElement,
    discover_docx_files,
    extract_title,
    parse_created_datetime_from_title,
    parse_elements,
)


def varint_encode(value: int) -> bytes:
    out = bytearray()
    while True:
        to_write = value & 0x7F
        value >>= 7
        if value:
            out.append(to_write | 0x80)
        else:
            out.append(to_write)
            return bytes(out)


def make_bafy_id(seed: str) -> str:
    # CIDv1, dag-cbor codec (0x71), sha2-256 multihash.
    digest = hashlib.sha256(seed.encode("utf-8")).digest()
    cid_bytes = varint_encode(1) + varint_encode(0x71) + bytes([0x12, 0x20]) + digest
    alphabet = "abcdefghijklmnopqrstuvwxyz234567"

    bits = 0
    bits_left = 0
    encoded = []
    for byte in cid_bytes:
        bits = (bits << 8) | byte
        bits_left += 8
        while bits_left >= 5:
            idx = (bits >> (bits_left - 5)) & 31
            encoded.append(alphabet[idx])
            bits_left -= 5
    if bits_left:
        idx = (bits << (5 - bits_left)) & 31
        encoded.append(alphabet[idx])

    return "b" + "".join(encoded)


def make_block_id(seed: str) -> str:
    return hashlib.md5(seed.encode("utf-8")).hexdigest()[:24]


def slugify(value: str) -> str:
    value = value.lower().replace("ä", "ae").replace("ö", "oe").replace("ü", "ue").replace("ß", "ss")
    value = re.sub(r"[^a-z0-9]+", "-", value)
    value = value.strip("-")
    return value or "note"


def normalize_image_ext(ext: str) -> str:
    ext_l = ext.lower()
    if ext_l in {".jpeg", ".jpg"}:
        return ".jpg"
    return ext_l


def image_dimensions(data: bytes, mime: str) -> tuple[int, int]:
    if mime == "image/png" and len(data) >= 24 and data.startswith(b"\x89PNG\r\n\x1a\n"):
        width = int.from_bytes(data[16:20], "big")
        height = int.from_bytes(data[20:24], "big")
        return width, height

    if mime == "image/gif" and len(data) >= 10 and data[:3] == b"GIF":
        width = int.from_bytes(data[6:8], "little")
        height = int.from_bytes(data[8:10], "little")
        return width, height

    if mime == "image/webp" and len(data) >= 30 and data[:4] == b"RIFF" and data[8:12] == b"WEBP":
        chunk = data[12:16]
        if chunk == b"VP8X" and len(data) >= 30:
            width = 1 + int.from_bytes(data[24:27], "little")
            height = 1 + int.from_bytes(data[27:30], "little")
            return width, height

    if mime == "image/jpeg" and data.startswith(b"\xff\xd8"):
        i = 2
        while i + 9 < len(data):
            if data[i] != 0xFF:
                i += 1
                continue
            marker = data[i + 1]
            i += 2
            if marker in {0xD8, 0xD9}:
                continue
            if i + 2 > len(data):
                break
            seg_len = struct.unpack(">H", data[i : i + 2])[0]
            if seg_len < 2 or i + seg_len > len(data):
                break
            if marker in {0xC0, 0xC1, 0xC2, 0xC3, 0xC5, 0xC6, 0xC7, 0xC9, 0xCA, 0xCB, 0xCD, 0xCE, 0xCF}:
                if i + 7 < len(data):
                    height = struct.unpack(">H", data[i + 3 : i + 5])[0]
                    width = struct.unpack(">H", data[i + 5 : i + 7])[0]
                    return width, height
            i += seg_len

    return 0, 0


@dataclass
class TemplateData:
    page_proto: dict
    file_protos: list[dict]
    non_page_objects: dict[str, bytes]
    passthrough_entries: dict[str, bytes]


@dataclass
class EntrySection:
    title: str
    elements: list[Element]


@dataclass
class ManualReviewEntry:
    source_docx: str
    title: str
    reason: str
    total_images: int
    png_images: int
    tiny_png_images: int
    max_image_cluster: int
    short_text_lines: int
    long_text_lines: int


ENTRY_TITLE_RE = re.compile(
    r"^\s*(\d{1,2})\.\s*([A-Za-zÄÖÜäöüß]+)\s+(\d{4})(?:\b|\s*[-–—].*)",
    re.IGNORECASE,
)
WEEKDAY_RE = re.compile(
    r"^(Montag|Dienstag|Mittwoch|Donnerstag|Freitag|Samstag|Sonntag),\s+\d{1,2}\.\s+[A-Za-zÄÖÜäöüß]+\s+\d{4}$",
    re.IGNORECASE,
)
TIME_RE = re.compile(r"^\d{1,2}:\d{2}$")


def load_template(template_zip: Path) -> TemplateData:
    with zipfile.ZipFile(template_zip, "r") as z:
        names = z.namelist()

        page_proto = None
        file_protos: list[dict] = []
        non_page_objects: dict[str, bytes] = {}
        passthrough_entries: dict[str, bytes] = {}

        for name in names:
            data = z.read(name)
            if name.startswith("objects/") and name.endswith(".pb.json"):
                parsed = json.loads(data.decode("utf-8"))
                if parsed.get("sbType") == "Page" and page_proto is None:
                    page_proto = parsed
                elif parsed.get("sbType") != "Page":
                    non_page_objects[name] = data
                continue

            if name.startswith("filesObjects/") and name.endswith(".pb.json"):
                parsed = json.loads(data.decode("utf-8"))
                if parsed.get("sbType") == "FileObject":
                    file_protos.append(parsed)
                continue

            if (
                name.startswith("relations/")
                or name.startswith("types/")
                or name.startswith("templates/")
            ):
                passthrough_entries[name] = data

        if page_proto is None:
            raise ValueError("Kein Page-Prototyp in Template-ZIP gefunden")
        if not file_protos:
            raise ValueError("Kein FileObject-Prototyp in Template-ZIP gefunden")

        return TemplateData(
            page_proto=page_proto,
            file_protos=file_protos,
            non_page_objects=non_page_objects,
            passthrough_entries=passthrough_entries,
        )


def should_skip_header_artifact(text: str) -> bool:
    return bool(WEEKDAY_RE.match(text) or TIME_RE.match(text))


def split_into_sections(elements: list[Element]) -> list[EntrySection]:
    sections: list[EntrySection] = []
    current: EntrySection | None = None

    for element in elements:
        if isinstance(element, TextElement):
            text = element.plain_text.strip()
            if text and ENTRY_TITLE_RE.match(text):
                if current is not None:
                    sections.append(current)
                current = EntrySection(title=text, elements=[element])
                continue

        if current is not None:
            current.elements.append(element)

    if current is not None:
        sections.append(current)

    if sections:
        return sections

    fallback_title = extract_title(elements)
    return [EntrySection(title=fallback_title, elements=list(elements))]


def select_file_proto(template: TemplateData, mime: str) -> dict:
    mime_l = mime.lower()
    for proto in template.file_protos:
        details = proto.get("snapshot", {}).get("data", {}).get("details", {})
        if str(details.get("fileMimeType", "")).lower() == mime_l:
            return proto
    return template.file_protos[0]


def infer_doc_year_from_filename(docx_path: Path) -> int | None:
    match = re.search(r"\b(20\d{2})\b", docx_path.stem)
    if not match:
        return None
    return int(match.group(1))


def resolve_created_datetime(
    title: str,
    timezone_name: str,
    doc_year_override: int | None,
) -> tuple[datetime, bool]:
    created_dt = parse_created_datetime_from_title(title, timezone_name)
    if doc_year_override is None:
        return created_dt, False

    match = ENTRY_TITLE_RE.match(title)
    if not match:
        return created_dt, False

    parsed_year = int(match.group(3))
    if parsed_year == doc_year_override:
        return created_dt, False

    try:
        return created_dt.replace(year=doc_year_override), True
    except ValueError:
        return created_dt, False


def default_restrictions() -> dict:
    return {
        "read": False,
        "edit": False,
        "remove": False,
        "drag": False,
        "dropOn": False,
    }


def text_block(block_id: str, text: str, style: str, bold: bool) -> dict:
    marks = []
    if bold and text:
        marks = [{"range": {"from": 0, "to": len(text)}, "type": "Bold", "param": ""}]
    return {
        "id": block_id,
        "fields": None,
        "restrictions": default_restrictions(),
        "childrenIds": [],
        "backgroundColor": "",
        "align": "AlignLeft",
        "verticalAlign": "VerticalAlignTop",
        "text": {
            "text": text,
            "style": style,
            "marks": {"marks": marks},
            "checked": False,
            "color": "",
            "iconEmoji": "",
            "iconImage": "",
        },
    }


def file_embed_block(block_id: str, file_name: str, mime: str, size: int, target_id: str) -> dict:
    return {
        "id": block_id,
        "fields": None,
        "restrictions": default_restrictions(),
        "childrenIds": [],
        "backgroundColor": "",
        "align": "AlignLeft",
        "verticalAlign": "VerticalAlignTop",
        "file": {
            "hash": "",
            "name": file_name,
            "type": "Image",
            "mime": mime,
            "size": str(size),
            "addedAt": "0",
            "targetObjectId": target_id,
            "state": "Done",
            "style": "Embed",
        },
    }


def page_from_docx(
    docx_path: Path,
    template: TemplateData,
    timezone_name: str,
    seed_prefix: str,
    ink_cluster_threshold: int,
) -> tuple[list[tuple[dict, list[tuple[str, dict, bytes]]]], list[ManualReviewEntry], int]:
    with zipfile.ZipFile(docx_path, "r") as docx_zip:
        elements = parse_elements(docx_zip)
        if not elements:
            raise ValueError(f"Keine Inhalte in DOCX gefunden: {docx_path}")

        sections = split_into_sections(elements)
        out_pages: list[tuple[dict, list[tuple[str, dict, bytes]]]] = []
        review_entries: list[ManualReviewEntry] = []
        corrected_year_count = 0
        doc_year_override = infer_doc_year_from_filename(docx_path)

        for section_index, section in enumerate(sections, start=1):
            title = section.title
            created_dt, was_corrected = resolve_created_datetime(
                title=title,
                timezone_name=timezone_name,
                doc_year_override=doc_year_override,
            )
            if was_corrected:
                corrected_year_count += 1
            created_unix = int(created_dt.timestamp())
            page_id = make_bafy_id(
                f"{seed_prefix}|page|{docx_path.name}|section:{section_index}|{title}"
            )

            content_blocks: list[dict] = []
            file_entries: list[tuple[str, dict, bytes]] = []
            list_of_file_ids: list[str] = []

            title_consumed = False
            content_started = False
            image_counter = 0
            total_images = 0
            png_images = 0
            tiny_png_images = 0
            max_image_cluster = 0
            current_image_cluster = 0
            short_text_lines = 0
            long_text_lines = 0

            for idx, element in enumerate(section.elements):
                if isinstance(element, TextElement):
                    text = element.plain_text.strip()
                    if not text:
                        continue

                    if current_image_cluster > max_image_cluster:
                        max_image_cluster = current_image_cluster
                    current_image_cluster = 0

                    if not title_consumed and text == title:
                        title_consumed = True
                        continue

                    if not content_started and should_skip_header_artifact(text):
                        continue

                    lowered = text.lower()
                    if lowered.startswith("was war heute gut"):
                        pass
                    else:
                        if len(text) <= 2:
                            short_text_lines += 1
                        if len(text) >= 40:
                            long_text_lines += 1

                    style = "Marked" if element.list_type else "Paragraph"
                    bold = element.markdown_text == f"**{text}**"
                    block_id = make_block_id(f"{page_id}|text|{idx}|{text}")
                    content_blocks.append(text_block(block_id, text, style, bold))
                    content_started = True
                    continue

                if isinstance(element, ImageElement):
                    image_counter += 1
                    total_images += 1
                    current_image_cluster += 1
                    source_name = Path(element.source_path).name
                    ext = normalize_image_ext(Path(source_name).suffix or ".bin")
                    mime = {
                        ".jpg": "image/jpeg",
                        ".jpeg": "image/jpeg",
                        ".png": "image/png",
                        ".gif": "image/gif",
                        ".webp": "image/webp",
                    }.get(ext, "application/octet-stream")

                    try:
                        binary = docx_zip.read(element.source_path)
                    except KeyError:
                        continue

                    width_px, height_px = image_dimensions(binary, mime)
                    if mime == "image/png":
                        png_images += 1
                        if width_px <= 80 and height_px <= 80:
                            tiny_png_images += 1
                    asset_name = (
                        f"{slugify(docx_path.stem)}-s{section_index:03d}-i{image_counter:03d}{ext}"
                    )
                    file_id = make_bafy_id(
                        f"{seed_prefix}|file|{docx_path.name}|section:{section_index}|img:{image_counter}|{asset_name}"
                    )
                    block_id = make_block_id(f"{page_id}|image|{idx}|{asset_name}")

                    list_of_file_ids.append(file_id)
                    content_blocks.append(
                        file_embed_block(block_id, asset_name, mime, len(binary), file_id)
                    )

                    template_file_obj = copy.deepcopy(select_file_proto(template, mime))
                    file_obj = file_object_from_template(
                        template_obj=template_file_obj,
                        file_id=file_id,
                        page_id=page_id,
                        file_name=Path(asset_name).stem,
                        file_ext=ext.lstrip("."),
                        file_mime=mime,
                        file_size=len(binary),
                        file_source=asset_name,
                        width_px=width_px,
                        height_px=height_px,
                        created_unix=created_unix,
                    )
                    file_entries.append((asset_name, file_obj, binary))
                    content_started = True

            page_obj = page_object_from_proto(
                proto=template.page_proto,
                page_id=page_id,
                title=title,
                created_unix=created_unix,
                file_ids=list_of_file_ids,
                content_blocks=content_blocks,
            )
            out_pages.append((page_obj, file_entries))

            if current_image_cluster > max_image_cluster:
                max_image_cluster = current_image_cluster

            reason = ""
            if max_image_cluster >= ink_cluster_threshold:
                reason = f"ink-fragment cluster ({max_image_cluster} images in sequence)"
            elif png_images > 0 and short_text_lines >= 2 and long_text_lines == 0:
                reason = "possible handwriting-only note (PNG + very short text lines)"

            if reason:
                review_entries.append(
                    ManualReviewEntry(
                        source_docx=docx_path.name,
                        title=title,
                        reason=reason,
                        total_images=total_images,
                        png_images=png_images,
                        tiny_png_images=tiny_png_images,
                        max_image_cluster=max_image_cluster,
                        short_text_lines=short_text_lines,
                        long_text_lines=long_text_lines,
                    )
                )

        return out_pages, review_entries, corrected_year_count


def page_object_from_proto(
    proto: dict,
    page_id: str,
    title: str,
    created_unix: int,
    file_ids: list[str],
    content_blocks: list[dict],
) -> dict:
    obj = copy.deepcopy(proto)
    blocks = obj["snapshot"]["data"]["blocks"]

    header = next((b for b in blocks if b.get("id") == "header"), None)
    title_block = next((b for b in blocks if b.get("id") == "title"), None)
    featured = next((b for b in blocks if b.get("id") == "featuredRelations"), None)

    root_block = {
        "id": page_id,
        "fields": None,
        "restrictions": default_restrictions(),
        "childrenIds": ["header"] + [b["id"] for b in content_blocks],
        "backgroundColor": "",
        "align": "AlignLeft",
        "verticalAlign": "VerticalAlignTop",
        "smartblock": {},
    }

    new_blocks = [root_block]
    if header:
        new_blocks.append(header)
    if content_blocks:
        new_blocks.extend(content_blocks)
    if title_block:
        new_blocks.append(title_block)
    if featured:
        new_blocks.append(featured)
    obj["snapshot"]["data"]["blocks"] = new_blocks

    details = obj["snapshot"]["data"].get("details", {})
    details["id"] = page_id
    details["name"] = title
    details["createdDate"] = created_unix
    details["lastModifiedDate"] = created_unix
    details["addedDate"] = created_unix
    details["links"] = file_ids
    details["backlinks"] = []
    details["mentions"] = []
    details["snippet"] = ""
    details.pop("oldAnytypeID", None)
    details.pop("sourceFilePath", None)
    obj["snapshot"]["data"]["details"] = details

    return obj


def file_object_from_template(
    template_obj: dict,
    file_id: str,
    page_id: str,
    file_name: str,
    file_ext: str,
    file_mime: str,
    file_size: int,
    file_source: str,
    width_px: int,
    height_px: int,
    created_unix: int,
) -> dict:
    obj = copy.deepcopy(template_obj)
    blocks = obj["snapshot"]["data"]["blocks"]
    details = obj["snapshot"]["data"].get("details", {})

    old_root_id = blocks[0].get("id", "") if blocks else ""

    for block in blocks:
        if old_root_id and block.get("id") == old_root_id:
            block["id"] = file_id
        if "file" in block:
            block["file"]["name"] = file_name
            block["file"]["mime"] = file_mime
            block["file"]["size"] = str(file_size)
            block["file"]["targetObjectId"] = file_id

    details["id"] = file_id
    details["name"] = file_name
    details["iconImage"] = file_id
    details["fileExt"] = file_ext
    details["fileMimeType"] = file_mime
    details["sizeInBytes"] = file_size
    details["widthInPixels"] = width_px
    details["heightInPixels"] = height_px
    details["createdDate"] = created_unix
    details["lastModifiedDate"] = created_unix
    details["addedDate"] = created_unix
    details["source"] = f"files/{file_source}"
    details["fileId"] = ""
    details["fileSourceChecksum"] = ""
    details["fileVariantIds"] = []
    details["fileVariantPaths"] = []
    details["fileVariantKeys"] = []
    details["fileVariantChecksums"] = []
    details["fileVariantMills"] = []
    details["fileVariantWidths"] = []
    details["fileVariantOptions"] = []
    details["fileIndexingStatus"] = 0
    details["fileSyncStatus"] = 0
    details["fileBackupStatus"] = 0
    details["syncStatus"] = 0
    details["syncDate"] = 0
    details["syncError"] = 0
    details["backlinks"] = [page_id]
    details["links"] = []
    details["mentions"] = []
    details["snippet"] = ""
    details["importType"] = 3
    details["origin"] = 3
    details.pop("oldAnytypeID", None)
    details.pop("sourceFilePath", None)
    obj["snapshot"]["data"]["details"] = details

    return obj


def build_anytype_zip(
    docx_files: Iterable[Path],
    output_zip: Path,
    template_zip: Path,
    timezone_name: str,
    ink_cluster_threshold: int,
) -> tuple[list[ManualReviewEntry], int]:
    template = load_template(template_zip)
    output_zip.parent.mkdir(parents=True, exist_ok=True)
    all_review_entries: list[ManualReviewEntry] = []
    corrected_year_count = 0

    with zipfile.ZipFile(output_zip, "w", compression=zipfile.ZIP_DEFLATED) as out:
        # Keep baseline schema files.
        for name, data in template.passthrough_entries.items():
            out.writestr(name, data)

        # Keep participant/system objects from template.
        for name, data in template.non_page_objects.items():
            out.writestr(name, data)

        for idx, docx_path in enumerate(docx_files, start=1):
            seed_prefix = f"{output_zip.name}|{idx}|{docx_path.name}|{os.path.getsize(docx_path)}"
            pages, review_entries, corrected_count_for_docx = page_from_docx(
                docx_path=docx_path,
                template=template,
                timezone_name=timezone_name,
                seed_prefix=seed_prefix,
                ink_cluster_threshold=ink_cluster_threshold,
            )
            all_review_entries.extend(review_entries)
            corrected_year_count += corrected_count_for_docx

            for page_obj, file_entries in pages:
                page_id = page_obj["snapshot"]["data"]["details"]["id"]
                out.writestr(
                    f"objects/{page_id}.pb.json",
                    json.dumps(page_obj, ensure_ascii=False, separators=(",", ":")),
                )

                for asset_name, file_obj, binary in file_entries:
                    file_id = file_obj["snapshot"]["data"]["details"]["id"]
                    out.writestr(
                        f"filesObjects/{file_id}.pb.json",
                        json.dumps(file_obj, ensure_ascii=False, separators=(",", ":")),
                    )
                    out.writestr(f"files/{asset_name}", binary)

                print(f"OK: {docx_path.name} -> objects/{page_id}.pb.json")

    return all_review_entries, corrected_year_count


def collect_manual_review_entries(
    docx_files: Iterable[Path],
    ink_cluster_threshold: int,
) -> list[ManualReviewEntry]:
    entries: list[ManualReviewEntry] = []

    for docx_path in docx_files:
        with zipfile.ZipFile(docx_path, "r") as docx_zip:
            elements = parse_elements(docx_zip)
            if not elements:
                continue
            sections = split_into_sections(elements)

            for section in sections:
                total_images = 0
                png_images = 0
                tiny_png_images = 0
                max_image_cluster = 0
                current_image_cluster = 0
                short_text_lines = 0
                long_text_lines = 0

                title_consumed = False
                content_started = False

                for element in section.elements:
                    if isinstance(element, TextElement):
                        text = element.plain_text.strip()
                        if not text:
                            continue

                        if current_image_cluster > max_image_cluster:
                            max_image_cluster = current_image_cluster
                        current_image_cluster = 0

                        if not title_consumed and text == section.title:
                            title_consumed = True
                            continue

                        if not content_started and should_skip_header_artifact(text):
                            continue

                        lowered = text.lower()
                        if not lowered.startswith("was war heute gut"):
                            if len(text) <= 2:
                                short_text_lines += 1
                            if len(text) >= 40:
                                long_text_lines += 1

                        content_started = True
                        continue

                    if isinstance(element, ImageElement):
                        total_images += 1
                        current_image_cluster += 1

                        ext = normalize_image_ext(Path(element.source_path).suffix or ".bin")
                        mime = {
                            ".jpg": "image/jpeg",
                            ".jpeg": "image/jpeg",
                            ".png": "image/png",
                            ".gif": "image/gif",
                            ".webp": "image/webp",
                        }.get(ext, "application/octet-stream")

                        if mime == "image/png":
                            png_images += 1

                        try:
                            binary = docx_zip.read(element.source_path)
                        except KeyError:
                            continue

                        width_px, height_px = image_dimensions(binary, mime)
                        if mime == "image/png" and width_px <= 80 and height_px <= 80:
                            tiny_png_images += 1

                if current_image_cluster > max_image_cluster:
                    max_image_cluster = current_image_cluster

                reason = ""
                if max_image_cluster >= ink_cluster_threshold:
                    reason = f"ink-fragment cluster ({max_image_cluster} images in sequence)"
                elif png_images > 0 and short_text_lines >= 2 and long_text_lines == 0:
                    reason = "possible handwriting-only note (PNG + very short text lines)"

                if reason:
                    entries.append(
                        ManualReviewEntry(
                            source_docx=docx_path.name,
                            title=section.title,
                            reason=reason,
                            total_images=total_images,
                            png_images=png_images,
                            tiny_png_images=tiny_png_images,
                            max_image_cluster=max_image_cluster,
                            short_text_lines=short_text_lines,
                            long_text_lines=long_text_lines,
                        )
                    )

    return entries


def write_manual_review_report(report_path: Path, entries: list[ManualReviewEntry]) -> None:
    lines = ["# Manual Review: Handschrift-Verdaechtige Eintraege", ""]

    if not entries:
        lines.extend(
            [
                "Keine verdaechtigen Eintraege erkannt.",
                "",
            ]
        )
    else:
        lines.extend(
            [
                "Diese Eintraege enthalten wahrscheinlich OneNote-Handschriftfragmente und sollten manuell geprueft/exportiert werden.",
                "",
            ]
        )

        for entry in entries:
            lines.append(f"- Quelle: `{entry.source_docx}`")
            lines.append(f"- Titel: `{entry.title}`")
            lines.append(f"- Grund: {entry.reason}")
            lines.append(
                "- Metriken: "
                f"total_images={entry.total_images}, "
                f"png_images={entry.png_images}, "
                f"tiny_png_images={entry.tiny_png_images}, "
                f"max_image_cluster={entry.max_image_cluster}, "
                f"short_text_lines={entry.short_text_lines}, "
                f"long_text_lines={entry.long_text_lines}"
            )
            lines.append("")

    report_path.parent.mkdir(parents=True, exist_ok=True)
    report_path.write_text("\n".join(lines).rstrip() + "\n", encoding="utf-8")


def parse_args(argv: list[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Konvertiert OneNote DOCX in Anytype-Exportstruktur (objects/filesObjects/files)."
    )
    parser.add_argument("--input", required=True, help=".docx Datei oder Ordner mit .docx")
    parser.add_argument("--output", default="anytype-import.zip", help="Ziel-ZIP")
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Nur analysieren und Manual-Review-Report erzeugen (kein ZIP schreiben)",
    )
    parser.add_argument(
        "--template-zip",
        default="Anytype-Template.zip",
        help=(
            "Anytype Export-ZIP als Template (enthaelt relations/types/templates). "
            "Default: Anytype-Template.zip"
        ),
    )
    parser.add_argument(
        "--timezone",
        default="Europe/Berlin",
        help="Zeitzone fuer 12:00 Timestamp aus Titel",
    )
    parser.add_argument(
        "--ink-cluster-threshold",
        type=int,
        default=40,
        help="Schwellwert fuer Handschrift-Cluster (Default: 40)",
    )
    parser.add_argument(
        "--manual-review-report",
        default="",
        help="Optionaler Pfad fuer Handschrift-Review-Report (.md)",
    )
    return parser.parse_args(argv)


def main(argv: list[str]) -> int:
    try:
        args = parse_args(argv)
        input_path = Path(args.input).expanduser().resolve()
        output_zip = Path(args.output).expanduser().resolve()
        template_zip = Path(args.template_zip).expanduser().resolve()
        report_path = (
            Path(args.manual_review_report).expanduser().resolve()
            if args.manual_review_report
            else (
                Path("manual-review-dry-run.md").resolve()
                if args.dry_run
                else output_zip.with_name(f"{output_zip.stem}-manual-review.md")
            )
        )

        docx_files = discover_docx_files(input_path)
        if args.dry_run:
            review_entries = collect_manual_review_entries(
                docx_files=docx_files,
                ink_cluster_threshold=args.ink_cluster_threshold,
            )
            corrected_year_count = 0
        else:
            if not template_zip.exists():
                raise ValueError(
                    "Template-ZIP nicht gefunden: "
                    f"{template_zip}. Lege eine Datei wie 'Anytype-Template.zip' bereit "
                    "oder nutze --template-zip <pfad>."
                )
            review_entries, corrected_year_count = build_anytype_zip(
                docx_files=docx_files,
                output_zip=output_zip,
                template_zip=template_zip,
                timezone_name=args.timezone,
                ink_cluster_threshold=args.ink_cluster_threshold,
            )
        if args.dry_run:
            if review_entries:
                write_manual_review_report(report_path, review_entries)
                print(f"Manual-Review-Report: {report_path}")
            else:
                print("Dry-Run: keine verdaechtigen Eintraege gefunden (kein Report geschrieben).")
            print(f"Dry-Run fertig. Verdaechtige Eintraege: {len(review_entries)}")
        else:
            write_manual_review_report(report_path, review_entries)
            print(f"Manual-Review-Report: {report_path}")
            if corrected_year_count > 0:
                print(
                    "Korrigierte Jahres-Tippfehler aus DOCX-Dateiname: "
                    f"{corrected_year_count}"
                )
            print(f"Fertig: {output_zip}")
        return 0
    except Exception as exc:  # noqa: BLE001
        print(f"Fehler: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
