#!/usr/bin/env python3
from __future__ import annotations

import argparse
import copy
import hashlib
import json
import os
import re
import sys
import zipfile
from dataclasses import dataclass
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


def safe_file_name(value: str) -> str:
    cleaned = value.strip()
    cleaned = re.sub(r'[\\/:*?"<>|]+', "-", cleaned)
    cleaned = re.sub(r"\s+", " ", cleaned)
    return cleaned.strip(" .") or "Notiz"


def slugify(value: str) -> str:
    value = value.lower().replace("ä", "ae").replace("ö", "oe").replace("ü", "ue").replace("ß", "ss")
    value = re.sub(r"[^a-z0-9]+", "-", value)
    value = value.strip("-")
    return value or "note"


@dataclass
class TemplateData:
    page_proto: dict
    file_proto: dict
    non_page_objects: dict[str, bytes]
    passthrough_entries: dict[str, bytes]


def load_template(template_zip: Path) -> TemplateData:
    with zipfile.ZipFile(template_zip, "r") as z:
        names = z.namelist()

        page_proto = None
        file_proto = None
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
                if parsed.get("sbType") == "FileObject" and file_proto is None:
                    file_proto = parsed
                continue

            if (
                name.startswith("relations/")
                or name.startswith("types/")
                or name.startswith("templates/")
            ):
                passthrough_entries[name] = data

        if page_proto is None:
            raise ValueError("Kein Page-Prototyp in Template-ZIP gefunden")
        if file_proto is None:
            raise ValueError("Kein FileObject-Prototyp in Template-ZIP gefunden")

        return TemplateData(
            page_proto=page_proto,
            file_proto=file_proto,
            non_page_objects=non_page_objects,
            passthrough_entries=passthrough_entries,
        )


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
) -> tuple[dict, list[tuple[str, dict, bytes]]]:
    with zipfile.ZipFile(docx_path, "r") as docx_zip:
        elements = parse_elements(docx_zip)
        if not elements:
            raise ValueError(f"Keine Inhalte in DOCX gefunden: {docx_path}")

        title = extract_title(elements)
        created_dt = parse_created_datetime_from_title(title, timezone_name)
        created_unix = int(created_dt.timestamp())
        page_id = make_bafy_id(f"{seed_prefix}|page|{docx_path.name}|{title}")

        # Build content blocks and file objects in order.
        content_blocks: list[dict] = []
        file_entries: list[tuple[str, dict, bytes]] = []
        list_of_file_ids: list[str] = []
        title_consumed = False
        content_started = False
        image_counter = 0

        for idx, element in enumerate(elements):
            if isinstance(element, TextElement):
                text = element.plain_text.strip()
                if not text:
                    continue

                if not title_consumed and text == title:
                    title_consumed = True
                    continue

                # Skip OneNote header artifacts at beginning.
                if not content_started and (re.match(r"^(Montag|Dienstag|Mittwoch|Donnerstag|Freitag|Samstag|Sonntag),", text) or re.match(r"^\d{1,2}:\d{2}$", text)):
                    continue

                style = "Marked" if element.list_type else "Paragraph"
                bold = element.markdown_text == f"**{text}**"
                block_id = make_block_id(f"{page_id}|text|{idx}|{text}")
                content_blocks.append(text_block(block_id, text, style, bold))
                content_started = True
                continue

            if isinstance(element, ImageElement):
                image_counter += 1
                source_name = Path(element.source_path).name
                ext = Path(source_name).suffix.lower() or ".bin"
                mime = {
                    ".jpg": "image/jpeg",
                    ".jpeg": "image/jpeg",
                    ".png": "image/png",
                    ".gif": "image/gif",
                    ".webp": "image/webp",
                }.get(ext, "application/octet-stream")

                asset_name = f"{slugify(title)}-image-{image_counter:02d}{ext}"
                file_id = make_bafy_id(f"{seed_prefix}|file|{docx_path.name}|{asset_name}")
                block_id = make_block_id(f"{page_id}|image|{idx}|{asset_name}")

                try:
                    binary = docx_zip.read(element.source_path)
                except KeyError:
                    continue

                list_of_file_ids.append(file_id)
                content_blocks.append(file_embed_block(block_id, asset_name, mime, len(binary), file_id))
                file_obj = file_object_from_proto(
                    proto=template.file_proto,
                    file_id=file_id,
                    page_id=page_id,
                    file_name=Path(asset_name).stem,
                    file_ext=ext.lstrip("."),
                    file_mime=mime,
                    file_size=len(binary),
                    file_source=asset_name,
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
        return page_obj, file_entries


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


def file_object_from_proto(
    proto: dict,
    file_id: str,
    page_id: str,
    file_name: str,
    file_ext: str,
    file_mime: str,
    file_size: int,
    file_source: str,
    created_unix: int,
) -> dict:
    obj = copy.deepcopy(proto)
    blocks = obj["snapshot"]["data"]["blocks"]
    old_root_id = blocks[0]["id"]

    for block in blocks:
        if block.get("id") == old_root_id:
            block["id"] = file_id
        if "file" in block:
            block["file"]["name"] = file_name
            block["file"]["mime"] = file_mime
            block["file"]["size"] = str(file_size)
            block["file"]["targetObjectId"] = file_id

    details = obj["snapshot"]["data"].get("details", {})
    details["id"] = file_id
    details["name"] = file_name
    details["fileExt"] = file_ext
    details["fileMimeType"] = file_mime
    details["sizeInBytes"] = file_size
    details["addedDate"] = created_unix
    details["source"] = f"files\\{file_source}"
    details["backlinks"] = [page_id]
    details["importType"] = 3
    details["origin"] = 3
    details.pop("oldAnytypeID", None)
    obj["snapshot"]["data"]["details"] = details

    return obj


def build_anytype_zip(
    docx_files: Iterable[Path],
    output_zip: Path,
    template_zip: Path,
    timezone_name: str,
) -> None:
    template = load_template(template_zip)
    output_zip.parent.mkdir(parents=True, exist_ok=True)

    with zipfile.ZipFile(output_zip, "w", compression=zipfile.ZIP_DEFLATED) as out:
        # Keep baseline schema files.
        for name, data in template.passthrough_entries.items():
            out.writestr(name, data)

        # Keep participant/system objects from template.
        for name, data in template.non_page_objects.items():
            out.writestr(name, data)

        for idx, docx_path in enumerate(docx_files, start=1):
            seed_prefix = f"{output_zip.name}|{idx}|{docx_path.name}|{os.path.getsize(docx_path)}"
            page_obj, file_entries = page_from_docx(
                docx_path=docx_path,
                template=template,
                timezone_name=timezone_name,
                seed_prefix=seed_prefix,
            )

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


def parse_args(argv: list[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Konvertiert OneNote DOCX in Anytype-Exportstruktur (objects/filesObjects/files)."
    )
    parser.add_argument("--input", required=True, help=".docx Datei oder Ordner mit .docx")
    parser.add_argument("--output", default="anytype-import.zip", help="Ziel-ZIP")
    parser.add_argument(
        "--template-zip",
        required=True,
        help="Anytype Export-ZIP als Template (enthaelt relations/types/templates)",
    )
    parser.add_argument(
        "--timezone",
        default="Europe/Berlin",
        help="Zeitzone fuer 12:00 Timestamp aus Titel",
    )
    return parser.parse_args(argv)


def main(argv: list[str]) -> int:
    try:
        args = parse_args(argv)
        input_path = Path(args.input).expanduser().resolve()
        output_zip = Path(args.output).expanduser().resolve()
        template_zip = Path(args.template_zip).expanduser().resolve()

        docx_files = discover_docx_files(input_path)
        build_anytype_zip(
            docx_files=docx_files,
            output_zip=output_zip,
            template_zip=template_zip,
            timezone_name=args.timezone,
        )
        print(f"Fertig: {output_zip}")
        return 0
    except Exception as exc:  # noqa: BLE001
        print(f"Fehler: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
