"""Patching logic: apply a translated JSON back into a .pptx package.

Only the DrawingML text nodes that have a non-empty, hash-matching translation
are modified. Every other byte in the ZIP package is copied through untouched,
so images, layouts, theme, metadata, and even unmodified XML parts stay exactly
as they were in the original file.
"""

from __future__ import annotations

import copy
import json
import os
import shutil
import zipfile

from lxml import etree

from . import ooxml_utils as ox
from .extract import discover_parts, iter_paragraph_units
from .ooxml_utils import A_BR, A_R, A_RPR, A_T, q

XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"


class PatchError(Exception):
    """Raised for user-facing patch problems (bad JSON, etc.)."""


# ---------------------------------------------------------------------------
# JSON loading / validation
# ---------------------------------------------------------------------------
def load_translation(json_path: str) -> list[dict]:
    """Load and sanity-check a translation JSON file."""
    try:
        with open(json_path, "r", encoding="utf-8") as fh:
            data = json.load(fh)
    except json.JSONDecodeError as exc:
        raise PatchError(
            f"The translation file '{json_path}' is not valid JSON "
            f"(line {exc.lineno}, column {exc.colno}). Did you edit it by hand "
            f"and break a quote or comma?"
        ) from exc

    if isinstance(data, dict) and "units" in data:
        units = data["units"]
    elif isinstance(data, list):
        units = data
    else:
        raise PatchError(
            "The translation file does not look like a file produced by "
            "'extract'. Expected an object with a 'units' list."
        )

    if not isinstance(units, list):
        raise PatchError("The 'units' field in the translation file must be a list.")

    for i, u in enumerate(units):
        if not isinstance(u, dict) or "id" not in u:
            raise PatchError(
                f"Unit #{i} in the translation file is missing an 'id'. The file "
                f"may be corrupted."
            )
    return units


# ---------------------------------------------------------------------------
# Paragraph text replacement
# ---------------------------------------------------------------------------
def _make_run(rpr_template: etree._Element | None, text: str) -> etree._Element:
    """Build a new <a:r> run with optional cloned run-properties and text."""
    run = etree.SubElement(etree.Element("tmp"), A_R)
    run.getparent().remove(run)  # detach from throwaway parent
    if rpr_template is not None:
        run.append(copy.deepcopy(rpr_template))
    t = etree.SubElement(run, A_T)
    t.text = text
    if text != text.strip() or "\t" in text:
        t.set(XML_SPACE, "preserve")
    return run


def set_paragraph_text(paragraph: etree._Element, text: str) -> bool:
    """Replace a paragraph's visible text with ``text``.

    Policy:
      * The first run's formatting (<a:rPr>) is inherited for the new text.
      * All existing runs (<a:r>) and line breaks (<a:br>) are removed and
        rebuilt from ``text`` (``\n`` becomes an <a:br>), so the translator has
        full control of line breaking.
      * Other elements (fields <a:fld>, paragraph properties <a:pPr>, end-run
        properties <a:endParaRPr>) are preserved in place.

    Returns True if the paragraph had at least one run to patch.
    """
    runs = [c for c in paragraph if c.tag == A_R]
    if not runs:
        return False

    # Capture the first run's formatting template and its position.
    first_run = runs[0]
    rpr_template = first_run.find(A_RPR)
    insert_index = list(paragraph).index(first_run)

    # Remove all existing runs and line breaks.
    for child in list(paragraph):
        if child.tag in (A_R, A_BR):
            paragraph.remove(child)

    # Build the new run / break sequence from the translated text.
    lines = text.split("\n")
    new_nodes: list[etree._Element] = []
    for li, line in enumerate(lines):
        if li > 0:
            new_nodes.append(etree.Element(A_BR))
        new_nodes.append(_make_run(rpr_template, line))

    for offset, node in enumerate(new_nodes):
        paragraph.insert(insert_index + offset, node)
    return True


# ---------------------------------------------------------------------------
# Main patch routine
# ---------------------------------------------------------------------------
def patch_pptx(
    pptx_path: str,
    json_path: str,
    output_path: str,
    report_path: str | None = None,
    force: bool = False,
) -> dict:
    """Apply translations and write a new .pptx. Returns the QA report dict."""
    if not os.path.isfile(pptx_path):
        raise PatchError(f"Input presentation not found: {pptx_path}")
    if not os.path.isfile(json_path):
        raise PatchError(f"Translation file not found: {json_path}")
    if os.path.abspath(pptx_path) == os.path.abspath(output_path):
        raise PatchError(
            "The output file must be different from the input file. The original "
            "is never overwritten."
        )

    units = load_translation(json_path)
    # Index translations by id (last one wins on duplicates).
    translations = {u["id"]: u for u in units}

    report = {
        "input": os.path.basename(pptx_path),
        "translation": os.path.basename(json_path),
        "output": os.path.basename(output_path),
        "forced": force,
        "summary": {
            "total_units_in_file": len(units),
            "translated": 0,
            "empty": 0,
            "stale": 0,
            "missing": 0,
            "no_runs": 0,
            "unchanged_in_pptx": 0,
        },
        "issues": [],
    }

    with zipfile.ZipFile(pptx_path) as zf:
        parts = discover_parts(zf)

        # Build a lookup of every paragraph unit present in the ORIGINAL file.
        live_units: dict[str, object] = {}
        for part in parts:
            for pu in iter_paragraph_units(part):
                live_units[pu.unit_id] = pu

        applied_ids: set[str] = set()

        for uid, tu in translations.items():
            target = (tu.get("target") or "")
            if target.strip() == "":
                report["summary"]["empty"] += 1
                continue

            pu = live_units.get(uid)
            if pu is None:
                report["summary"]["missing"] += 1
                report["issues"].append(
                    {"id": uid, "issue": "missing",
                     "detail": "This id was not found in the presentation. The "
                               "slide may have been deleted or edited."}
                )
                continue

            # Stale check: has the source changed since extraction?
            expected_hash = tu.get("source_hash")
            actual_hash = ox.sha256_text(pu.source)
            if expected_hash and expected_hash != actual_hash and not force:
                report["summary"]["stale"] += 1
                report["issues"].append(
                    {"id": uid, "issue": "stale",
                     "detail": "The original text changed since extraction. "
                               "Re-run extract, or use --force to patch anyway.",
                     "source_now": pu.source}
                )
                continue

            ok = set_paragraph_text(pu.paragraph, target)
            if not ok:
                report["summary"]["no_runs"] += 1
                report["issues"].append(
                    {"id": uid, "issue": "no_runs",
                     "detail": "Paragraph had no text run to patch (field-only?)."}
                )
                continue

            pu.part.changed = True
            applied_ids.add(uid)
            report["summary"]["translated"] += 1

        # Units present in the file but not translated at all.
        report["summary"]["unchanged_in_pptx"] = len(
            [uid for uid in live_units if uid not in translations]
        )

        # Validate changed parts re-serialise as well-formed XML.
        changed_bytes: dict[str, bytes] = {}
        for part in parts:
            if not part.changed:
                continue
            data = ox.serialize_xml(part.root, part.raw)
            try:
                etree.fromstring(data)  # well-formed check
            except etree.XMLSyntaxError as exc:  # pragma: no cover - safety net
                raise PatchError(
                    f"Patching produced invalid XML in {part.path}: {exc}"
                ) from exc
            changed_bytes[part.path] = data

        _write_package(zf, output_path, changed_bytes)

    if report_path:
        with open(report_path, "w", encoding="utf-8") as fh:
            json.dump(report, fh, ensure_ascii=False, indent=2)
        report["report_file"] = report_path

    return report


def _write_package(
    src_zip: zipfile.ZipFile, output_path: str, changed: dict[str, bytes]
) -> None:
    """Copy the original ZIP to ``output_path``, swapping in changed parts.

    Entries are written in the original order with their original metadata, so
    unchanged parts are byte-identical to the source package.
    """
    tmp_path = output_path + ".tmp"
    with zipfile.ZipFile(tmp_path, "w") as out:
        for item in src_zip.infolist():
            data = changed.get(item.filename)
            if data is None:
                data = src_zip.read(item.filename)
            # Preserve compression type and external attributes per entry.
            new_info = zipfile.ZipInfo(item.filename, date_time=item.date_time)
            new_info.compress_type = item.compress_type
            new_info.external_attr = item.external_attr
            new_info.internal_attr = item.internal_attr
            new_info.create_system = item.create_system
            new_info.flag_bits = item.flag_bits
            out.writestr(new_info, data)
    shutil.move(tmp_path, output_path)
