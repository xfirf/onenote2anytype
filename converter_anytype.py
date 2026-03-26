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


def make_content_cid(data: bytes, codec: int = 0x70) -> str:
    digest = hashlib.sha256(data).digest()
    cid_bytes = varint_encode(1) + varint_encode(codec) + bytes([0x12, 0x20]) + digest
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
        template_file_pool_by_mime: dict[str, list[dict]] = {}
        for fp in template.file_protos:
            det = fp.get("snapshot", {}).get("data", {}).get("details", {})
            mime_key = str(det.get("fileMimeType", "")).lower()
            template_file_pool_by_mime.setdefault(mime_key, []).append(fp)
        template_file_pool_any = list(template.file_protos)

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
                ext = normalize_image_ext(Path(source_name).suffix or ".bin")
                mime = {
                    ".jpg": "image/jpeg",
                    ".jpeg": "image/jpeg",
                    ".png": "image/png",
                    ".gif": "image/gif",
                    ".webp": "image/webp",
                }.get(ext, "application/octet-stream")

                asset_name = f"{slugify(title)}-image-{image_counter:02d}{ext}"

                try:
                    binary = docx_zip.read(element.source_path)
                except KeyError:
                    continue
                width_px, height_px = image_dimensions(binary, mime)

                candidates = template_file_pool_by_mime.get(mime.lower()) or template_file_pool_any
                if not candidates:
                    raise ValueError(
                        "Template-ZIP enthaelt nicht genug FileObjects fuer alle Bilder"
                    )
                selected_template_obj = candidates.pop(0)
                if selected_template_obj in template_file_pool_any:
                    template_file_pool_any.remove(selected_template_obj)
                for pool in template_file_pool_by_mime.values():
                    if selected_template_obj in pool:
                        pool.remove(selected_template_obj)

                template_file_obj = copy.deepcopy(selected_template_obj)

                det = template_file_obj["snapshot"]["data"]["details"]
                file_id = det.get("id")
                if not file_id:
                    raise ValueError("Template-FileObject ohne id gefunden")

                source_path = str(det.get("source", ""))
                source_name = Path(source_path.replace("\\", "/")).name
                if not source_name:
                    source_name = asset_name

                block_id = make_block_id(f"{page_id}|image|{idx}|{source_name}")

                list_of_file_ids.append(file_id)
                content_blocks.append(file_embed_block(block_id, source_name, mime, len(binary), file_id))
                file_obj = file_object_from_template(
                    template_obj=template_file_obj,
                    page_id=page_id,
                    file_size=len(binary),
                    width_px=width_px,
                    height_px=height_px,
                    created_unix=created_unix,
                )
                file_entries.append((source_name, file_obj, binary))
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


def file_object_from_template(
    template_obj: dict,
    page_id: str,
    file_size: int,
    width_px: int,
    height_px: int,
    created_unix: int,
) -> dict:
    obj = copy.deepcopy(template_obj)
    blocks = obj["snapshot"]["data"]["blocks"]
    details = obj["snapshot"]["data"].get("details", {})
    file_id = details.get("id", "")

    for block in blocks:
        if "file" in block:
            block["file"]["size"] = str(file_size)
            block["file"]["targetObjectId"] = file_id

    details["sizeInBytes"] = file_size
    details["widthInPixels"] = width_px
    details["heightInPixels"] = height_px
    details["lastModifiedDate"] = created_unix
    details["addedDate"] = created_unix
    source_val = str(details.get("source", ""))
    if source_val:
        details["source"] = source_val.replace("\\", "/")
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
