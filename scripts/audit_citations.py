#!/usr/bin/env python3
"""Audit manual and EndNote citations in a WPS/Word DOCX manuscript."""

from __future__ import annotations

import argparse
import json
import re
import sys
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

from inspect_endnote_fields import W, inspect_docx

NUMERIC_CITATION = re.compile(r"\[\s*(\d+(?:\s*[-–—,，;；]\s*\d+)*)\s*\]")
YEAR = re.compile(r"\b(?:19|20)\d{2}\b")
REFERENCE_HEADING = re.compile(r"^\s*(?:参考文献|references|bibliography)\s*[:：]?\s*$", re.I)
NUMBERED_REFERENCE = re.compile(r"^\s*\d+\s*[.．、\])]\s*")


def _paragraph_parts(paragraph: ET.Element) -> tuple[str, str]:
    visible: list[str] = []
    outside_endnote: list[str] = []
    field_stack: list[dict[str, object]] = []

    for node in paragraph.iter():
        if node.tag == f"{W}fldChar":
            kind = node.attrib.get(f"{W}fldCharType")
            if kind == "begin":
                field_stack.append({"instruction": "", "result": False})
            elif kind == "separate" and field_stack:
                field_stack[-1]["result"] = True
            elif kind == "end" and field_stack:
                field_stack.pop()
        elif node.tag == f"{W}instrText" and field_stack:
            field_stack[-1]["instruction"] = str(field_stack[-1]["instruction"]) + (node.text or "")
        elif node.tag == f"{W}t":
            text = node.text or ""
            visible.append(text)
            inside_endnote_cite = any(
                "ADDIN EN.CITE" in str(field["instruction"]) for field in field_stack
            )
            if not inside_endnote_cite:
                outside_endnote.append(text)
    return "".join(visible), "".join(outside_endnote)


def _expand_numbers(group: str) -> set[int]:
    numbers: set[int] = set()
    for part in re.split(r"\s*[,，;；]\s*", group):
        bounds = re.split(r"\s*[-–—]\s*", part)
        if len(bounds) == 2 and all(item.isdigit() for item in bounds):
            start, end = map(int, bounds)
            if 0 < start <= end and end - start <= 1000:
                numbers.update(range(start, end + 1))
        elif part.strip().isdigit():
            numbers.add(int(part.strip()))
    return numbers


def _is_reference_like(text: str) -> bool:
    return len(text) >= 30 and bool(YEAR.search(text)) and bool(NUMBERED_REFERENCE.match(text))


def audit_docx(path: Path) -> dict:
    field_report = inspect_docx(path)
    with zipfile.ZipFile(path) as archive:
        root = ET.fromstring(archive.read("word/document.xml"))

    paragraphs = list(root.iter(f"{W}p"))
    visible_texts: list[str] = []
    outside_texts: list[str] = []
    instructions: list[str] = []
    for paragraph in paragraphs:
        visible, outside = _paragraph_parts(paragraph)
        visible_texts.append(visible.strip())
        outside_texts.append(outside.strip())
        instructions.append(
            "".join(node.text or "" for node in paragraph.iter(f"{W}instrText"))
        )

    manual_numbers: set[int] = set()
    manual_occurrences = 0
    for text in outside_texts:
        for match in NUMERIC_CITATION.finditer(text):
            manual_occurrences += 1
            manual_numbers.update(_expand_numbers(match.group(1)))

    reflist_indexes = [
        index for index, instruction in enumerate(instructions) if "ADDIN EN.REFLIST" in instruction
    ]
    heading_indexes = [
        index for index, text in enumerate(visible_texts) if REFERENCE_HEADING.match(text)
    ]
    manual_reference_candidates: list[dict[str, object]] = []
    if heading_indexes:
        heading = heading_indexes[-1]
        stop = reflist_indexes[0] if reflist_indexes else len(paragraphs)
        for index in range(heading + 1, max(heading + 1, stop)):
            text = visible_texts[index]
            if _is_reference_like(text):
                manual_reference_candidates.append(
                    {"paragraph_index": index, "text": text}
                )
    else:
        # WPS documents may visually show a references heading that is not exposed
        # as ordinary paragraph text. Walk backward from EN.REFLIST when it exists,
        # otherwise from the document end, and collect the contiguous numbered,
        # year-bearing reference block.
        found_reference = False
        anchor = reflist_indexes[0] if reflist_indexes else len(paragraphs)
        for index in range(anchor - 1, -1, -1):
            text = visible_texts[index]
            if not text:
                if found_reference:
                    continue
                continue
            if _is_reference_like(text):
                found_reference = True
                manual_reference_candidates.append(
                    {"paragraph_index": index, "text": text}
                )
                continue
            if found_reference:
                break
        manual_reference_candidates.reverse()

    has_cites = field_report["endnote_citation_field_count"] > 0
    has_reflist = field_report["endnote_bibliography_field_count"] > 0
    has_manual = bool(manual_numbers or manual_reference_candidates)
    if has_cites and has_reflist and has_manual:
        status = "partial_conversion"
    elif has_cites and has_reflist:
        status = "dynamic_endnote"
    elif has_cites:
        status = "dynamic_citations_without_bibliography"
    elif has_manual:
        status = "manual_only"
    else:
        status = "no_citations_detected"

    warnings: list[str] = []
    if has_reflist and manual_reference_candidates:
        warnings.append("Manual reference entries remain before an EndNote EN.REFLIST field.")
    if field_report["endnote_bibliography_field_count"] > 1:
        warnings.append("Multiple EndNote EN.REFLIST fields were found.")
    if (has_cites or has_reflist) and not field_report["field_markers_balanced"]:
        warnings.append("Complex field marker counts are unbalanced.")
    if has_cites and manual_numbers:
        warnings.append("Dynamic EndNote citations and manual numeric citations coexist.")

    return {
        "path": str(path.resolve()),
        "status": status,
        "endnote_citation_field_count": field_report["endnote_citation_field_count"],
        "endnote_embedded_record_count": field_report["endnote_embedded_record_count"],
        "endnote_bibliography_field_count": field_report["endnote_bibliography_field_count"],
        "field_markers_balanced": field_report["field_markers_balanced"],
        "manual_numeric_citation_occurrences": manual_occurrences,
        "manual_numeric_citation_numbers": sorted(manual_numbers),
        "manual_reference_candidate_count": len(manual_reference_candidates),
        "manual_reference_candidates": manual_reference_candidates,
        "warnings": warnings,
    }


def _print_human(report: dict) -> None:
    print(f"Document: {report['path']}")
    print(f"Status: {report['status']}")
    print(f"EN.CITE fields: {report['endnote_citation_field_count']}")
    print(f"Embedded records: {report['endnote_embedded_record_count']}")
    print(f"EN.REFLIST fields: {report['endnote_bibliography_field_count']}")
    print(f"Field markers balanced: {report['field_markers_balanced']}")
    print(
        "Manual numeric citations: "
        f"{report['manual_numeric_citation_occurrences']} occurrences; "
        f"numbers={report['manual_numeric_citation_numbers']}"
    )
    print(f"Manual reference candidates: {report['manual_reference_candidate_count']}")
    if report["warnings"]:
        print("Warnings:")
        for warning in report["warnings"]:
            print(f"  - {warning}")


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("docx", type=Path)
    parser.add_argument("--json", action="store_true", dest="as_json")
    args = parser.parse_args()
    try:
        report = audit_docx(args.docx)
    except (FileNotFoundError, ValueError, zipfile.BadZipFile, ET.ParseError) as exc:
        print(f"error: {exc}", file=sys.stderr)
        return 2
    if args.as_json:
        print(json.dumps(report, ensure_ascii=False, indent=2))
    else:
        _print_human(report)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
