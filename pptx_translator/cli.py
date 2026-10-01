"""Command line interface for the pptx_translator tool.

Two sub-commands:

    extract   Pull all translatable text from a .pptx into a JSON file.
    patch     Apply a translated JSON back into a new .pptx + QA report.

Designed to give friendly, plain-English errors instead of Python tracebacks.
"""

from __future__ import annotations

import argparse
import os
import sys

from .extract import extract_to_json
from .patch import PatchError, patch_pptx

PROG = "pptx_translator"


def _err(message: str) -> int:
    """Print a clean error message (no traceback) and return exit code 1."""
    sys.stderr.write(f"\n  Error: {message}\n\n")
    return 1


def _check_input(path: str) -> str | None:
    if not os.path.exists(path):
        return f"File not found: {path}"
    if not path.lower().endswith(".pptx"):
        return (f"'{path}' does not look like a .pptx file. Please point me at a "
                f"PowerPoint (.pptx) file.")
    return None


def cmd_extract(args: argparse.Namespace) -> int:
    problem = _check_input(args.input)
    if problem:
        return _err(problem)

    output = args.output or (os.path.splitext(args.input)[0] + ".translation.json")
    try:
        summary = extract_to_json(args.input, output)
    except Exception as exc:  # noqa: BLE001 - surface clean message
        return _err(f"Could not read the presentation: {exc}")

    print(f"\n  Extracted {summary['unit_count']} translation units.")
    if summary["by_part"]:
        print("  Breakdown by part:")
        for part, count in sorted(summary["by_part"].items()):
            print(f"    - {part:<12} {count}")
    print(f"\n  Translation file written to:\n    {output}\n")
    print("  Next: open that JSON file, fill in each 'target' field with the")
    print("  translation, then run the 'patch' command.\n")
    return 0


def cmd_patch(args: argparse.Namespace) -> int:
    problem = _check_input(args.input)
    if problem:
        return _err(problem)

    output = args.output or (os.path.splitext(args.input)[0] + ".translated.pptx")
    report = args.report or (os.path.splitext(output)[0] + ".qa.json")
    try:
        result = patch_pptx(
            args.input, args.translation, output,
            report_path=report, force=args.force,
        )
    except PatchError as exc:
        return _err(str(exc))
    except Exception as exc:  # noqa: BLE001
        return _err(f"Unexpected problem while patching: {exc}")

    s = result["summary"]
    print(f"\n  Translated presentation written to:\n    {output}\n")
    print("  Summary:")
    print(f"    - translated units : {s['translated']}")
    print(f"    - empty (skipped)  : {s['empty']}")
    print(f"    - stale (skipped)  : {s['stale']}")
    print(f"    - missing ids      : {s['missing']}")
    if s["no_runs"]:
        print(f"    - no text run      : {s['no_runs']}")
    print(f"    - left untranslated: {s['unchanged_in_pptx']}")
    print(f"\n  QA report written to:\n    {report}\n")
    if s["stale"] and not args.force:
        print("  Note: some units were skipped as 'stale' (the slide text changed")
        print("  since extraction). Re-run 'extract', or add --force to apply them.\n")
    return 0


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog=PROG,
        description="Translate PowerPoint (.pptx) files while preserving 100%% of "
                    "the original formatting, images and layout.",
    )
    sub = parser.add_subparsers(dest="command", required=True)

    pe = sub.add_parser(
        "extract", help="Extract all text from a .pptx into a JSON file.")
    pe.add_argument("input", help="Path to the source .pptx file.")
    pe.add_argument("-o", "--output", help="Path for the JSON file to write.")
    pe.set_defaults(func=cmd_extract)

    pp = sub.add_parser(
        "patch", help="Apply a translated JSON back into a new .pptx.")
    pp.add_argument("input", help="Path to the ORIGINAL .pptx file.")
    pp.add_argument("translation", help="Path to the translated JSON file.")
    pp.add_argument("-o", "--output", help="Path for the new translated .pptx.")
    pp.add_argument("--report", help="Path for the JSON QA report.")
    pp.add_argument(
        "--force", action="store_true",
        help="Apply translations even when the source text changed since "
             "extraction (hash mismatch).")
    pp.set_defaults(func=cmd_patch)
    return parser


def main(argv: list[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)
    return args.func(args)


if __name__ == "__main__":  # pragma: no cover
    sys.exit(main())
