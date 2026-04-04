#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
import mimetypes
import re
import subprocess
import sys
import time
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from urllib.parse import urlparse

import requests

try:
    import msal
except ImportError:  # pragma: no cover - runtime dependency guidance
    msal = None


GRAPH_ROOT = "https://graph.microsoft.com/v1.0"
DEFAULT_SCOPES = ["Notes.Read", "offline_access"]
URL_ATTR_RE = re.compile(
    r'(?P<attr>src|href|data-fullres-src|data-src|data-render-src)="(?P<url>https://[^"]+)"',
    re.IGNORECASE,
)


@dataclass
class SectionInfo:
    id: str
    display_name: str
    parent_path: str


def normalize_name(value: str) -> str:
    return value.strip().casefold()


def safe_name(value: str, fallback: str = "untitled") -> str:
    cleaned = re.sub(r'[\\/:*?"<>|]+', "-", value.strip())
    cleaned = re.sub(r"\s+", " ", cleaned).strip(" .")
    return cleaned or fallback


def save_token_cache(cache_path: Path, cache: "msal.SerializableTokenCache") -> None:
    if cache.has_state_changed:
        cache_path.parent.mkdir(parents=True, exist_ok=True)
        cache_path.write_text(cache.serialize(), encoding="utf-8")


def acquire_access_token(
    client_id: str,
    tenant: str,
    scopes: list[str],
    cache_path: Path,
) -> str:
    if msal is None:
        raise RuntimeError(
            "Paket 'msal' fehlt. Installiere mit: pip install msal requests"
        )

    cache = msal.SerializableTokenCache()
    if cache_path.exists():
        cache.deserialize(cache_path.read_text(encoding="utf-8"))

    authority = f"https://login.microsoftonline.com/{tenant}"
    app = msal.PublicClientApplication(
        client_id=client_id,
        authority=authority,
        token_cache=cache,
    )

    accounts = app.get_accounts()
    token_result = None
    if accounts:
        token_result = app.acquire_token_silent(scopes=scopes, account=accounts[0])

    if not token_result:
        flow = app.initiate_device_flow(scopes=scopes)
        if "user_code" not in flow:
            raise RuntimeError(f"Device-Flow konnte nicht gestartet werden: {flow}")

        print("Bitte im Browser anmelden:")
        print(flow["verification_uri"])
        print(f"Code: {flow['user_code']}")
        token_result = app.acquire_token_by_device_flow(flow)

    save_token_cache(cache_path, cache)

    access_token = token_result.get("access_token") if token_result else None
    if not access_token:
        error = (token_result or {}).get("error_description") or str(token_result)
        raise RuntimeError(f"Token konnte nicht abgerufen werden: {error}")
    return access_token


class GraphClient:
    def __init__(self, access_token: str, timeout: int = 60):
        self.session = requests.Session()
        self.session.headers.update(
            {
                "Authorization": f"Bearer {access_token}",
                "Accept": "application/json",
                "User-Agent": "onenote2anytype-export/1.0",
            }
        )
        self.timeout = timeout

    def request(self, method: str, url: str, **kwargs) -> requests.Response:
        full_url = url if url.startswith("http") else f"{GRAPH_ROOT}{url}"

        max_attempts = 6
        for attempt in range(1, max_attempts + 1):
            response = self.session.request(
                method=method,
                url=full_url,
                timeout=self.timeout,
                **kwargs,
            )

            if response.status_code in {429, 500, 502, 503, 504} and attempt < max_attempts:
                retry_after = response.headers.get("Retry-After")
                wait = float(retry_after) if retry_after and retry_after.isdigit() else min(2 ** (attempt - 1), 30)
                print(f"Warnung: {response.status_code} bei {full_url}, retry in {wait:.1f}s")
                time.sleep(wait)
                continue

            if response.status_code >= 400:
                message = response.text
                raise RuntimeError(
                    f"Graph API Fehler {response.status_code} fuer {full_url}: {message[:500]}"
                )
            return response

        raise RuntimeError(f"Graph API dauerhaft fehlgeschlagen: {full_url}")

    def paged_get(self, url: str) -> list[dict]:
        items: list[dict] = []
        next_url = url
        while next_url:
            response = self.request("GET", next_url)
            payload = response.json()
            items.extend(payload.get("value", []))
            next_url = payload.get("@odata.nextLink")
        return items

    def get_page_content_html(self, page_id: str) -> str:
        response = self.request("GET", f"/me/onenote/pages/{page_id}/content?includeIDs=true")
        content_type = response.headers.get("Content-Type", "")
        if "text/html" not in content_type and "application/xhtml+xml" not in content_type:
            raise RuntimeError(
                f"Unerwarteter Content-Type fuer Seite {page_id}: {content_type}"
            )
        return response.text


def find_notebook(client: GraphClient, notebook_name: str) -> dict:
    notebooks = client.paged_get("/me/onenote/notebooks")
    target = normalize_name(notebook_name)
    for notebook in notebooks:
        if normalize_name(notebook.get("displayName", "")) == target:
            return notebook

    available = ", ".join(sorted(n.get("displayName", "") for n in notebooks))
    raise RuntimeError(
        f"Notebook '{notebook_name}' nicht gefunden. Verfuegbar: {available}"
    )


def collect_sections_recursive(client: GraphClient, group_id: str, prefix: str) -> list[SectionInfo]:
    sections: list[SectionInfo] = []

    direct_sections = client.paged_get(f"/me/onenote/sectionGroups/{group_id}/sections")
    for section in direct_sections:
        sections.append(
            SectionInfo(
                id=section["id"],
                display_name=section.get("displayName", ""),
                parent_path=prefix,
            )
        )

    nested_groups = client.paged_get(f"/me/onenote/sectionGroups/{group_id}/sectionGroups")
    for nested in nested_groups:
        nested_name = nested.get("displayName", "")
        nested_prefix = f"{prefix}/{nested_name}" if prefix else nested_name
        sections.extend(
            collect_sections_recursive(
                client=client,
                group_id=nested["id"],
                prefix=nested_prefix,
            )
        )
    return sections


def list_sections_for_notebook(client: GraphClient, notebook_id: str) -> list[SectionInfo]:
    sections: list[SectionInfo] = []

    top_sections = client.paged_get(f"/me/onenote/notebooks/{notebook_id}/sections")
    for section in top_sections:
        sections.append(
            SectionInfo(
                id=section["id"],
                display_name=section.get("displayName", ""),
                parent_path="",
            )
        )

    groups = client.paged_get(f"/me/onenote/notebooks/{notebook_id}/sectionGroups")
    for group in groups:
        group_name = group.get("displayName", "")
        sections.extend(
            collect_sections_recursive(
                client=client,
                group_id=group["id"],
                prefix=group_name,
            )
        )

    return sections


def resolve_target_sections(all_sections: list[SectionInfo], wanted_names: list[str]) -> list[SectionInfo]:
    by_name: dict[str, list[SectionInfo]] = {}
    for section in all_sections:
        by_name.setdefault(normalize_name(section.display_name), []).append(section)

    resolved: list[SectionInfo] = []
    missing: list[str] = []

    for wanted in wanted_names:
        hits = by_name.get(normalize_name(wanted), [])
        if not hits:
            missing.append(wanted)
            continue
        if len(hits) > 1:
            paths = [f"{h.parent_path}/{h.display_name}".strip("/") for h in hits]
            raise RuntimeError(
                f"Abschnitt '{wanted}' ist mehrfach vorhanden: {paths}. Bitte eindeutiger waehlen."
            )
        resolved.append(hits[0])

    if missing:
        available = sorted(f"{s.parent_path}/{s.display_name}".strip("/") for s in all_sections)
        raise RuntimeError(
            "Abschnitte nicht gefunden: "
            f"{missing}. Verfuegbar: {available}"
        )

    return resolved


def should_download_asset(url: str) -> bool:
    parsed = urlparse(url)
    host = parsed.netloc.lower()
    path = parsed.path.lower()

    if "graph.microsoft.com" in host and "/onenote/resources/" in path:
        return True
    if "onenote.com" in host and "resources" in path:
        return True
    return False


def infer_extension_from_response(url: str, response: requests.Response) -> str:
    parsed = urlparse(url)
    path_suffix = Path(parsed.path).suffix
    if path_suffix:
        return path_suffix

    mime = response.headers.get("Content-Type", "").split(";")[0].strip().lower()
    if mime:
        ext = mimetypes.guess_extension(mime) or ""
        if ext == ".jpe":
            return ".jpg"
        if ext:
            return ext
    return ".bin"


def rewrite_html_with_local_assets(
    html: str,
    html_path: Path,
    client: GraphClient,
    download_assets: bool,
) -> tuple[str, int]:
    if not download_assets:
        return html, 0

    assets_dir = html_path.with_name(f"{html_path.stem}_assets")
    assets_dir.mkdir(parents=True, exist_ok=True)

    url_to_local: dict[str, str] = {}
    downloaded = 0

    def replacement(match: re.Match[str]) -> str:
        nonlocal downloaded

        attr = match.group("attr")
        url = match.group("url")
        if not should_download_asset(url):
            return match.group(0)

        local_rel = url_to_local.get(url)
        if local_rel is None:
            response = client.request("GET", url, headers={"Accept": "*/*"})
            ext = infer_extension_from_response(url, response)
            local_name = f"asset-{len(url_to_local) + 1:03d}{ext}"
            local_path = assets_dir / local_name
            local_path.write_bytes(response.content)
            local_rel = f"{assets_dir.name}/{local_name}"
            url_to_local[url] = local_rel
            downloaded += 1

        return f'{attr}="{local_rel}"'

    rewritten = URL_ATTR_RE.sub(replacement, html)
    if downloaded == 0:
        try:
            assets_dir.rmdir()
        except OSError:
            pass
    return rewritten, downloaded


def parse_dt(value: str) -> datetime:
    value_norm = value.replace("Z", "+00:00")
    return datetime.fromisoformat(value_norm)


def ensure_unique_stem(stem: str, used: set[str]) -> str:
    if stem not in used:
        used.add(stem)
        return stem

    counter = 2
    while True:
        candidate = f"{stem}-{counter}"
        if candidate not in used:
            used.add(candidate)
            return candidate
        counter += 1


def convert_html_to_docx(html_path: Path, docx_path: Path, pandoc_bin: str) -> None:
    command = [
        pandoc_bin,
        str(html_path),
        "-f",
        "html",
        "-t",
        "docx",
        "--resource-path",
        str(html_path.parent),
        "-o",
        str(docx_path),
    ]
    result = subprocess.run(command, capture_output=True, text=True)
    if result.returncode != 0:
        raise RuntimeError(
            f"Pandoc Fehler fuer {html_path.name}: {result.stderr.strip() or result.stdout.strip()}"
        )


def export_section_pages(
    client: GraphClient,
    section: SectionInfo,
    out_root: Path,
    convert_docx: bool,
    pandoc_bin: str,
    limit_pages: int,
    download_assets: bool,
) -> tuple[int, int, int]:
    section_dir_name = safe_name(section.display_name, fallback="section")
    section_dir = out_root / section_dir_name
    html_dir = section_dir / "_html"
    section_dir.mkdir(parents=True, exist_ok=True)
    html_dir.mkdir(parents=True, exist_ok=True)

    pages = client.paged_get(f"/me/onenote/sections/{section.id}/pages")
    pages.sort(key=lambda p: parse_dt(p.get("createdDateTime") or p.get("lastModifiedDateTime")))
    if limit_pages > 0:
        pages = pages[:limit_pages]

    used_stems: set[str] = set()
    total_assets = 0
    total_docx = 0

    for index, page in enumerate(pages, start=1):
        page_id = page["id"]
        title = page.get("title") or f"page-{index:03d}"
        created = page.get("createdDateTime") or page.get("lastModifiedDateTime") or ""
        created_dt = parse_dt(created)
        prefix = created_dt.strftime("%Y-%m-%d_%H-%M")
        stem_raw = safe_name(f"{prefix} {title}", fallback=f"page-{index:03d}")
        stem = ensure_unique_stem(stem_raw, used_stems)

        print(f"  -> Seite {index}/{len(pages)}: {title}")
        html = client.get_page_content_html(page_id)

        html_path = html_dir / f"{stem}.html"
        html_rewritten, downloaded_assets = rewrite_html_with_local_assets(
            html=html,
            html_path=html_path,
            client=client,
            download_assets=download_assets,
        )
        total_assets += downloaded_assets
        html_path.write_text(html_rewritten, encoding="utf-8")

        meta = {
            "id": page_id,
            "title": title,
            "createdDateTime": page.get("createdDateTime"),
            "lastModifiedDateTime": page.get("lastModifiedDateTime"),
            "links": page.get("links", {}),
        }
        (html_dir / f"{stem}.meta.json").write_text(
            json.dumps(meta, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )

        if convert_docx:
            docx_path = section_dir / f"{stem}.docx"
            convert_html_to_docx(html_path=html_path, docx_path=docx_path, pandoc_bin=pandoc_bin)
            total_docx += 1

    return len(pages), total_docx, total_assets


def parse_args(argv: list[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Exportiert OneNote-Seiten per Microsoft Graph als HTML und optional DOCX."
    )
    parser.add_argument("--client-id", required=True, help="Azure App Client ID (Public Client)")
    parser.add_argument("--tenant", default="common", help="Tenant-ID oder 'common' (Default)")
    parser.add_argument("--notebook", required=True, help="Notebook-Name, z. B. Tagebücher")
    parser.add_argument(
        "--sections",
        nargs="+",
        required=True,
        help="Abschnittsnamen, z. B. 1981-2016 2017-2018",
    )
    parser.add_argument("--output", required=True, help="Ausgabeordner")
    parser.add_argument(
        "--scopes",
        nargs="+",
        default=DEFAULT_SCOPES,
        help="OAuth Scopes (Default: Notes.Read offline_access)",
    )
    parser.add_argument(
        "--token-cache",
        default=".graph_token_cache.json",
        help="Pfad fuer lokalen Token-Cache",
    )
    parser.add_argument(
        "--convert-docx",
        action="store_true",
        help="Konvertiert exportierte HTML-Seiten direkt nach DOCX (Pandoc noetig)",
    )
    parser.add_argument(
        "--pandoc-bin",
        default="pandoc",
        help="Pandoc-Binary (Default: pandoc)",
    )
    parser.add_argument(
        "--limit-pages",
        type=int,
        default=0,
        help="Optionales Limit pro Abschnitt (0 = kein Limit)",
    )
    parser.add_argument(
        "--no-download-assets",
        action="store_true",
        help="Assets nicht lokal herunterladen (HTML verweist auf Remote-URLs)",
    )
    return parser.parse_args(argv)


def main(argv: list[str]) -> int:
    try:
        args = parse_args(argv)
        out_root = Path(args.output).expanduser().resolve()
        cache_path = Path(args.token_cache).expanduser().resolve()

        access_token = acquire_access_token(
            client_id=args.client_id,
            tenant=args.tenant,
            scopes=args.scopes,
            cache_path=cache_path,
        )
        client = GraphClient(access_token=access_token)

        notebook = find_notebook(client, args.notebook)
        notebook_name = notebook.get("displayName", "")
        notebook_dir = out_root / safe_name(notebook_name, fallback="notebook")
        notebook_dir.mkdir(parents=True, exist_ok=True)

        all_sections = list_sections_for_notebook(client, notebook["id"])
        target_sections = resolve_target_sections(all_sections, args.sections)

        print(f"Notebook: {notebook_name}")
        total_pages = 0
        total_docx = 0
        total_assets = 0

        for section in target_sections:
            label = f"{section.parent_path}/{section.display_name}".strip("/")
            print(f"\nAbschnitt: {label}")
            page_count, docx_count, assets_count = export_section_pages(
                client=client,
                section=section,
                out_root=notebook_dir,
                convert_docx=args.convert_docx,
                pandoc_bin=args.pandoc_bin,
                limit_pages=args.limit_pages,
                download_assets=not args.no_download_assets,
            )
            total_pages += page_count
            total_docx += docx_count
            total_assets += assets_count
            print(
                f"  Fertig: pages={page_count}, docx={docx_count}, assets={assets_count}"
            )

        print("\nExport abgeschlossen 🎉")
        print(f"Ausgabe: {notebook_dir}")
        print(f"Seiten gesamt: {total_pages}")
        print(f"DOCX gesamt: {total_docx}")
        print(f"Assets gesamt: {total_assets}")
        return 0
    except Exception as exc:  # noqa: BLE001
        print(f"Fehler: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
