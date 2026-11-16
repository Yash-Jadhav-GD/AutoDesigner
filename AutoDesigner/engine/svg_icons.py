"""
Keyword -> SVG icon matching for icon_badge rows (AUTOFORMAT_SPECIFICATION.md section 6).

Matching is offline: keywords are scored against the local icon database
(../Autoformat_v1/temp_icons.json). The chosen SVG is read from .cache/icons/, and downloaded
there once from the icon CDN if missing (set AUTODESIGNER_OFFLINE=1 to never download).
Icons are embedded the way common.inject_svg_assets embeds the branded ones: the builder writes
an svgBlip <p:pic> with a relationship id, and embed_svg_icons adds the relationships, media
and content type to the saved package.
"""

import io
import json
import os
import re
import urllib.request
import zipfile
from functools import lru_cache
from pathlib import Path
from typing import Dict, List, Optional, Tuple

BASE_DIR = Path(__file__).resolve().parents[1]
ICONS_DB_PATH = BASE_DIR.parent / "Autoformat_v1" / "temp_icons.json"
ICONS_BASE_URL = "https://yash-jadhav-gd.github.io/Icon_finder/assets/icons/"
CACHE_DIR = BASE_DIR / ".cache" / "icons"

STOP_WORDS = {
    "a", "an", "the", "and", "or", "but", "is", "are", "was", "were",
    "in", "on", "at", "to", "for", "with", "about", "against", "between",
    "into", "through", "during", "before", "after", "above", "below",
    "from", "up", "down", "of", "off", "over", "under", "again", "further",
    "then", "once", "here", "there", "when", "where", "why", "how", "all",
    "any", "both", "each", "few", "more", "most", "other", "some", "such",
    "no", "nor", "not", "only", "own", "same", "so", "than", "too", "very",
}
WORD = re.compile(r"\b[a-zA-Z]{3,}\b")

# (relationship id, media file name, svg bytes)
IconRef = Tuple[str, str, bytes]


@lru_cache(maxsize=1)
def load_database() -> tuple:
    """Icon entries {id, file, label, tags, ...}; empty when the database is missing."""
    if not ICONS_DB_PATH.exists():
        return ()
    raw = ICONS_DB_PATH.read_bytes()
    if raw.startswith((b"\xff\xfe", b"\xfe\xff")):
        text = raw.decode("utf-16")
    else:
        text = raw.decode("utf-8-sig")
    return tuple(json.loads(text))


def keywords(text: str) -> List[str]:
    words = [w.lower() for w in WORD.findall(text)]
    return [w for w in words if w not in STOP_WORDS] or ["circle"]


def best_icon(text: str) -> Optional[dict]:
    """Highest-scoring icon: label contains word +5, exact tag +2, else partial tag +1.
    Ties keep the first entry, so all-zero scores fall back to the first icon."""
    db = load_database()
    words = keywords(text)
    best, best_score = None, -1
    for icon in db:
        label = (icon.get("label") or "").lower()
        tags = [t.lower() for t in icon.get("tags") or []]
        score = 0
        for w in words:
            if w in label:
                score += 5
            if w in tags:
                score += 2
            elif any(w in t for t in tags):
                score += 1
        if score > best_score:
            best, best_score = icon, score
    return best


def fetch_svg(icon: dict) -> Optional[bytes]:
    """SVG bytes for an icon from the local cache, downloading it once if allowed."""
    name = icon.get("file")
    if not name or "/" in name or "\\" in name:
        return None
    path = CACHE_DIR / name
    if path.exists():
        return path.read_bytes()
    if os.environ.get("AUTODESIGNER_OFFLINE"):
        return None
    try:
        with urllib.request.urlopen(ICONS_BASE_URL + name, timeout=10) as r:
            data = r.read()
    except Exception:
        return None
    if b"<svg" not in data[:2000]:
        return None
    CACHE_DIR.mkdir(parents=True, exist_ok=True)
    path.write_bytes(data)
    return data


def icon_for(text: str) -> Optional[Tuple[dict, bytes]]:
    icon = best_icon(text)
    svg = fetch_svg(icon) if icon else None
    return (icon, svg) if svg else None


def media_name(icon: dict) -> str:
    return f"icon_{re.sub(r'[^A-Za-z0-9]', '', str(icon.get('id') or icon['file']))}.svg"


def pic_xml(shape_id: int, name: str, rel_id: str, left: int, top: int, width: int, height: int) -> str:
    """svgBlip picture element (same structure as the branded builders' icons). Sizes in EMU."""
    return f'''<p:pic xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
      <p:nvPicPr>
        <p:cNvPr id="{shape_id}" name="{name}"/>
        <p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr>
        <p:nvPr/>
      </p:nvPicPr>
      <p:blipFill>
        <a:blip>
          <a:extLst>
            <a:ext uri="{{96DAC541-7B7A-43D3-8B79-37D633B846F1}}">
              <asvg:svgBlip xmlns:asvg="http://schemas.microsoft.com/office/drawing/2016/SVG/main" r:embed="{rel_id}"/>
            </a:ext>
          </a:extLst>
        </a:blip>
        <a:stretch><a:fillRect/></a:stretch>
      </p:blipFill>
      <p:spPr>
        <a:xfrm><a:off x="{int(left)}" y="{int(top)}"/><a:ext cx="{int(width)}" cy="{int(height)}"/></a:xfrm>
        <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
      </p:spPr>
    </p:pic>'''


def embed_svg_icons(raw_pptx_bytes: bytes, slide_icons: Dict[int, List[IconRef]]) -> bytes:
    """Adds the icon SVGs to ppt/media/ and their relationships to the slides that reference them.
    `slide_icons` maps slide number -> [(relationship id, media file name, svg bytes)]."""
    rels_targets = {f"ppt/slides/_rels/slide{n}.xml.rels": refs for n, refs in slide_icons.items()}
    media = {f"ppt/media/{name}": svg for refs in slide_icons.values() for _, name, svg in refs}

    out_buf = io.BytesIO()
    with zipfile.ZipFile(io.BytesIO(raw_pptx_bytes), "r") as zin, zipfile.ZipFile(out_buf, "w") as zout:
        existing = set(zin.namelist())
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename in rels_targets:
                new = "".join(
                    f'<Relationship Id="{rel_id}" '
                    f'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" '
                    f'Target="../media/{name}"/>'
                    for rel_id, name, _ in rels_targets[item.filename]
                    if f'Id="{rel_id}"'.encode() not in data
                )
                data = data.replace(b"</Relationships>", new.encode() + b"</Relationships>")
            elif item.filename == "[Content_Types].xml" and b'Extension="svg"' not in data:
                data = data.replace(b"</Types>", b'<Default Extension="svg" ContentType="image/svg+xml"/></Types>')
            zout.writestr(item, data)
        for path, svg in media.items():
            if path not in existing:
                zout.writestr(path, svg, compress_type=zipfile.ZIP_DEFLATED)
    return out_buf.getvalue()
