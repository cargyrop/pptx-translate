"""End-to-end test: create a .pptx, extract, translate, patch, verify.

python-pptx is used ONLY to synthesise a realistic test deck. The tool itself
(extract / patch) never uses python-pptx for reading or writing the package.

Run with:  python -m pytest tests/ -v
       or:  python tests/test_roundtrip.py
"""

from __future__ import annotations

import hashlib
import json
import os
import shutil
import tempfile
import zipfile

from pptx import Presentation
from pptx.util import Inches, Pt

from pptx_translator.extract import extract_to_json
from pptx_translator.patch import patch_pptx


def make_test_pptx(path: str) -> None:
    """Build a deck with a title slide, bullets, a table, notes, and a textbox."""
    prs = Presentation()

    # Slide 1 - title layout.
    s1 = prs.slides.add_slide(prs.slide_layouts[0])
    s1.shapes.title.text = "Quarterly Report"
    s1.placeholders[1].text = "Prepared by the Finance Team"

    # Slide 2 - bullets + a free text box + notes.
    s2 = prs.slides.add_slide(prs.slide_layouts[1])
    s2.shapes.title.text = "Key Highlights"
    body = s2.placeholders[1].text_frame
    body.text = "Revenue increased"
    body.add_paragraph().text = "Costs decreased"
    body.add_paragraph().text = "Profit doubled"

    tb = s2.shapes.add_textbox(Inches(1), Inches(5), Inches(4), Inches(1))
    tb.text_frame.text = "Confidential draft"

    notes = s2.notes_slide
    notes.notes_text_frame.text = "Remember to thank the team"

    # Slide 3 - a table.
    s3 = prs.slides.add_slide(prs.slide_layouts[5])
    s3.shapes.title.text = "Numbers"
    rows, cols = 2, 2
    table = s3.shapes.add_table(
        rows, cols, Inches(1), Inches(2), Inches(6), Inches(2)
    ).table
    table.cell(0, 0).text = "Metric"
    table.cell(0, 1).text = "Value"
    table.cell(1, 0).text = "Users"
    table.cell(1, 1).text = "One million"

    prs.save(path)


# A tiny fake "translator": uppercases + tags each string, deterministically.
def fake_translate(text: str) -> str:
    return "XX " + text.upper()


def file_sha256(path: str) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as fh:
        for chunk in iter(lambda: fh.read(8192), b""):
            h.update(chunk)
    return h.hexdigest()


def run() -> None:
    workdir = tempfile.mkdtemp(prefix="pptx_translator_test_")
    try:
        src = os.path.join(workdir, "deck.pptx")
        src_backup = os.path.join(workdir, "deck_backup.pptx")
        json_path = os.path.join(workdir, "deck.json")
        out = os.path.join(workdir, "deck_translated.pptx")
        report = os.path.join(workdir, "qa.json")

        # 1. Create the test deck and keep a byte backup to prove immutability.
        make_test_pptx(src)
        shutil.copy(src, src_backup)
        original_hash = file_sha256(src)

        # 2. Extract.
        summary = extract_to_json(src, json_path)
        print(f"  extract: {summary['unit_count']} units, by part {summary['by_part']}")
        assert summary["unit_count"] > 0, "no units extracted"

        with open(json_path, encoding="utf-8") as fh:
            doc = json.load(fh)
        units = doc["units"]

        # Expected source strings must all be present.
        sources = {u["source"] for u in units}
        for expected in [
            "Quarterly Report", "Prepared by the Finance Team",
            "Key Highlights", "Revenue increased", "Costs decreased",
            "Profit doubled", "Confidential draft",
            "Remember to thank the team", "Metric", "Value", "Users",
            "One million",
        ]:
            assert expected in sources, f"missing source text: {expected!r}"
        print("  all expected source strings were extracted (slides, textbox, "
              "table, notes)")

        # Verify each unit has required fields.
        for u in units:
            assert u["id"] and "source_hash" in u and "context" in u
            assert u["source_hash"] == hashlib.sha256(
                u["source"].encode()).hexdigest()
            assert u["status"] == "new" and u["target"] == ""

        # 3. Simulate translation by filling targets.
        for u in units:
            u["target"] = fake_translate(u["source"])
        with open(json_path, "w", encoding="utf-8") as fh:
            json.dump(doc, fh, ensure_ascii=False, indent=2)

        # 4. Patch.
        rep = patch_pptx(src, json_path, out, report_path=report)
        print(f"  patch summary: {rep['summary']}")
        assert rep["summary"]["translated"] == len(units), "not all units patched"
        assert rep["summary"]["missing"] == 0
        assert rep["summary"]["stale"] == 0

        # 5. Original file must be byte-for-byte unchanged.
        assert file_sha256(src) == original_hash, "ORIGINAL FILE WAS MODIFIED!"
        assert file_sha256(src_backup) == original_hash
        print("  original .pptx is byte-for-byte unchanged")

        # 6. Output must contain translated text and NOT the original text.
        out_prs = Presentation(out)
        found = set()
        for slide in out_prs.slides:
            for shape in _iter_shapes(slide.shapes):
                if shape.has_text_frame:
                    for para in shape.text_frame.paragraphs:
                        found.add("".join(r.text for r in para.runs))
                if shape.has_table:
                    for row in shape.table.rows:
                        for cell in row.cells:
                            found.add(cell.text)
            if slide.has_notes_slide:
                found.add(slide.notes_slide.notes_text_frame.text)

        assert "XX QUARTERLY REPORT" in found, f"translated title missing: {found}"
        assert "XX ONE MILLION" in found, "translated table cell missing"
        assert "XX REMEMBER TO THANK THE TEAM" in found, "translated notes missing"
        assert "Quarterly Report" not in found, "original text still present!"
        print("  translated .pptx contains translations and not the originals")

        # 7. Output must be a valid, openable ZIP/pptx.
        assert zipfile.is_zipfile(out)
        print("  output .pptx is a valid package")

        # 8. Stale detection: tamper with a source hash, patch again w/o --force.
        doc2 = json.loads(open(json_path, encoding="utf-8").read())
        doc2["units"][0]["source_hash"] = "deadbeef"
        stale_json = os.path.join(workdir, "stale.json")
        with open(stale_json, "w", encoding="utf-8") as fh:
            json.dump(doc2, fh)
        out2 = os.path.join(workdir, "deck_stale.pptx")
        rep2 = patch_pptx(src, stale_json, out2)
        assert rep2["summary"]["stale"] == 1, "stale unit not detected"
        print("  stale translation correctly detected and skipped")

        # 9. --force applies stale anyway.
        out3 = os.path.join(workdir, "deck_force.pptx")
        rep3 = patch_pptx(src, stale_json, out3, force=True)
        assert rep3["summary"]["stale"] == 0 and rep3["summary"]["translated"] == len(
            doc2["units"]
        )
        print("  --force applies stale translations as expected")

        print("\n  ALL TESTS PASSED\n")
    finally:
        shutil.rmtree(workdir, ignore_errors=True)


def _iter_shapes(shapes):
    """Recurse into group shapes."""
    for shape in shapes:
        yield shape
        if shape.shape_type == 6:  # GROUP
            yield from _iter_shapes(shape.shapes)


def test_roundtrip():
    run()


if __name__ == "__main__":
    run()
