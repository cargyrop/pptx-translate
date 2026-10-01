"""Extraction logic: walk a .pptx package and pull out translation units.

The public entry point is :func:`extract_to_json`. The walker helpers
(:func:`discover_parts`, :func:`iter_paragraph_units`) are shared with the
patch module so that unit IDs line up exactly between extract and patch.
"""

from __future__ import annotations

import json
import re
import zipfile
from dataclasses import dataclass, field
from typing import Iterator

from lxml import etree

from . import ooxml_utils as ox
from .ooxml_utils import NS, q

# ---------------------------------------------------------------------------
# Part discovery
# ---------------------------------------------------------------------------
# Each category: (regex over the zip path, part key prefix, human label prefix).
# Order here defines the order of units in the output file.
_PART_CATEGORIES = [
    (re.compile(r"^ppt/slides/slide(\d+)\.xml$"), "slide", "slide", "Slide"),
    (re.compile(r"^ppt/notesSlides/notesSlide(\d+)\.xml$"), "notes", "notes",
     "Notes for slide"),
    (re.compile(r"^ppt/slideLayouts/slideLayout(\d+)\.xml$"), "layout", "layout",
     "Slide Layout"),
    (re.compile(r"^ppt/slideMasters/slideMaster(\d+)\.xml$"), "master", "master",
     "Slide Master"),
    (re.compile(r"^ppt/notesMasters/notesMaster(\d+)\.xml$"), "notesmaster",
     "notesmaster", "Notes Master"),
    (re.compile(r"^ppt/handoutMasters/handoutMaster(\d+)\.xml$"), "handout",
     "handout", "Handout Master"),
]

# Non-visual property containers that hold a <p:cNvPr> (shape id + name).
_NV_LOCALNAMES = {
    "nvSpPr", "nvPicPr", "nvGraphicFramePr", "nvCxnSpPr", "nvGrpSpPr",
}


@dataclass
class PartInfo:
    """A single parsed XML part of the package that may contain text."""

    path: str                 # zip entry name, e.g. ppt/slides/slide1.xml
    part: str                 # category key: slide, notes, layout, ...
    index: int                # numeric index parsed from the file name
    label: str                # human readable label, e.g. "Slide 1"
    root: etree._Element      # parsed XML root
    raw: bytes                # original bytes (for faithful re-serialisation)
    changed: bool = field(default=False)


def discover_parts(zf: zipfile.ZipFile) -> list[PartInfo]:
    """Find and parse every text-bearing part, in a deterministic order."""
    found: list[tuple[int, int, PartInfo]] = []
    names = set(zf.namelist())
    for cat_order, (pattern, part_key, _key_prefix, label_prefix) in enumerate(
        _PART_CATEGORIES
    ):
        for name in names:
            m = pattern.match(name)
            if not m:
                continue
            idx = int(m.group(1))
            raw = zf.read(name)
            root = ox.parse_xml(raw)
            if part_key == "notes":
                label = f"Notes (notesSlide {idx})"
            else:
                label = f"{label_prefix} {idx}"
            found.append(
                (cat_order, idx, PartInfo(name, part_key, idx, label, root, raw))
            )
    found.sort(key=lambda t: (t[0], t[1]))
    return [p for _, _, p in found]


def _part_key(part: PartInfo) -> str:
    """Stable id prefix for a part, e.g. 'slide1', 'layout2', 'notes3'."""
    return f"{part.part}{part.index}"


# ---------------------------------------------------------------------------
# Shape metadata
# ---------------------------------------------------------------------------
def _local(tag) -> str:
    return etree.QName(tag).localname if isinstance(tag, str) else ""


def _shape_context(txbody: etree._Element) -> dict:
    """Resolve the nearest enclosing shape and describe it."""
    info = {
        "shape_id": None,
        "shape_name": None,
        "shape_type": "shape",
        "role": None,
    }
    for ancestor in txbody.iterancestors():
        nv = None
        for child in ancestor:
            if etree.QName(child).localname in _NV_LOCALNAMES:
                nv = child
                break
        if nv is None:
            continue

        cnvpr = nv.find(q("p", "cNvPr"))
        if cnvpr is not None:
            sid = cnvpr.get("id")
            info["shape_id"] = int(sid) if sid and sid.isdigit() else sid
            info["shape_name"] = cnvpr.get("name")

        anc_local = etree.QName(ancestor).localname
        type_map = {
            "sp": "shape",
            "pic": "picture",
            "graphicFrame": "table",
            "cxnSp": "connector",
            "grpSp": "group",
        }
        info["shape_type"] = type_map.get(anc_local, anc_local)

        # Placeholder detection (title / body / subtitle / etc.)
        ph = nv.find(f"{q('p', 'nvPr')}/{q('p', 'ph')}")
        if ph is not None:
            info["shape_type"] = "placeholder"
            info["role"] = ph.get("type") or "body"
        break
    return info


# ---------------------------------------------------------------------------
# The core walker
# ---------------------------------------------------------------------------
@dataclass
class ParagraphUnit:
    """One paragraph-level translation unit, with a live XML reference."""

    unit_id: str
    source: str
    context: dict
    paragraph: etree._Element
    part: PartInfo


# All text bodies, regardless of namespace prefix (p:txBody / a:txBody).
_TXBODY_XPATH = etree.XPath(".//*[local-name()='txBody']")


def iter_paragraph_units(part: PartInfo) -> Iterator[ParagraphUnit]:
    """Yield every paragraph-level unit inside a single part.

    Text bodies are visited in document order; paragraphs keep their true child
    index so patch can locate them again. Empty paragraphs are skipped for
    output but still consume an index, keeping IDs stable.
    """
    key = _part_key(part)
    for tb_idx, txbody in enumerate(_TXBODY_XPATH(part.root)):
        shape_ctx = _shape_context(txbody)
        paragraphs = txbody.findall(q("a", "p"))
        for p_idx, paragraph in enumerate(paragraphs):
            source = ox.paragraph_text(paragraph)
            if source.strip() == "":
                continue
            unit_id = f"{key}_tb{tb_idx}_p{p_idx}"
            context = {
                "part": part.part,
                "part_index": part.index,
                "part_label": part.label,
                "part_path": part.path,
                "shape_id": shape_ctx["shape_id"],
                "shape_name": shape_ctx["shape_name"],
                "shape_type": shape_ctx["shape_type"],
                "paragraph_index": p_idx,
                "character_limit_hint": len(source),
                "role": shape_ctx["role"],
            }
            yield ParagraphUnit(unit_id, source, context, paragraph, part)


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------
def build_units(pptx_path: str) -> list[dict]:
    """Return a list of JSON-serialisable translation units for a .pptx."""
    units: list[dict] = []
    with zipfile.ZipFile(pptx_path) as zf:
        for part in discover_parts(zf):
            for u in iter_paragraph_units(part):
                units.append(
                    {
                        "id": u.unit_id,
                        "source": u.source,
                        "target": "",
                        "source_hash": ox.sha256_text(u.source),
                        "context": u.context,
                        "status": "new",
                    }
                )
    return units


def extract_to_json(pptx_path: str, output_path: str) -> dict:
    """Extract translation units from ``pptx_path`` and write JSON to disk.

    Returns a small summary dict for the CLI to print.
    """
    units = build_units(pptx_path)
    document = {
        "schema": "pptx_translator/1.0",
        "source_file": pptx_path.replace("\\", "/").split("/")[-1],
        "unit_count": len(units),
        "instructions": (
            "Fill the 'target' field of each unit with the translated text. "
            "Keep the 'id' and 'source_hash' fields unchanged. Use \\n for line "
            "breaks. Leave 'target' empty to skip a unit."
        ),
        "units": units,
    }
    with open(output_path, "w", encoding="utf-8") as fh:
        json.dump(document, fh, ensure_ascii=False, indent=2)

    parts_seen = {}
    for u in units:
        parts_seen[u["context"]["part"]] = parts_seen.get(u["context"]["part"], 0) + 1
    return {"unit_count": len(units), "by_part": parts_seen, "output": output_path}
