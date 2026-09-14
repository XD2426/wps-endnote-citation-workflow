#!/usr/bin/env python3
"""Inspect EndNote CWYW fields in a DOCX using only the Python standard library."""

from __future__ import annotations

import argparse
import json
import sys
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W = f"{{{W_NS}}}"


def _paragraph_text(paragraph: ET.Element) -> str:
    return "".join(node.text or "" for node in paragraph.iter(f"{W}t"))


def _parse_endnote_payload(instruction: str) -> list[dict[str, str | None]]:
    start = instruction.find("<EndNote>")
    if start < 0:
        return []
    try:
        root = ET.fromstring(instruction[start:].strip())
    except ET.ParseError:
        return []

    records: list[dict[str, str | None]] = []
    for cite in root.findall(".//Cite"):
        record = cite.find("record")
        records.append(
            {
                "author": cite.findtext("Author"),
                "year": cite.findtext("Year"),
                "record_number": cite.findtext("RecNum"),
                "display_text": cite.findtext("DisplayText"),
                "title": record.findtext("./titles/title") if record is not None else None,
                "journal": record.findtext("./titles/secondary-title") if record is not None else None,
            }
        )
    return records


def _paragraph_fields(paragraph: ET.Element) -> list[dict[str, str]]:
    """Return complex fields in document order without merging adjacent fields."""
    stack: list[dict[str, object]] = []
    fields: list[dict[str, str]] = []
    for node in paragraph.iter():
        if node.tag == f"{W}fldChar":
            kind = node.attrib.get(f"{W}fldCharType")
            if kind == "begin":
                stack.append({"instruction": [], "result": [], "in_result": False})
            elif kind == "separate" and stack:
                stack[-1]["in_result"] = True
            elif kind == "end" and stack:
                field = stack.pop()
                fields.append(
                    {
                        "instruction": "".join(field["instruction"]),
                        "result_text": "".join(field["result"]),
                    }
                )
        elif node.tag == f"{W}instrText" and stack:
            stack[-1]["instruction"].append(node.text or "")
        elif node.tag == f"{W}t" and stack and stack[-1]["in_result"]:
            stack[-1]["result"].append(node.text or "")

    # EN.REFLIST can begin in one paragraph and end after several paragraphs.
    # Keep an instruction-bearing open field so its signature is still detected.
    for field in stack:
        instruction = "".join(field["instruction"])
        if instruction:
            fields.append(
                {
                    "instruction": instruction,
                    "result_text": "".join(field["result"]),
                }
            )
    return fields


def inspect_docx(path: Path) -> dict:
    if not path.is_file():
        raise FileNotFoundError(path)
    if path.suffix.lower() != ".docx":
        raise ValueError("Input must be a .docx file")

    with zipfile.ZipFile(path) as archive:
        try:
            xml_bytes = archive.read("word/document.xml")
        except KeyError as exc:
            raise ValueError("DOCX is missing word/document.xml") from exc

    root = ET.fromstring(xml_bytes)
    field_markers = {"begin": 0, "separate": 0, "end": 0, "other": 0}
    for node in root.iter(f"{W}fldChar"):
        kind = node.attrib.get(f"{W}fldCharType", "other")
        field_markers[kind if kind in field_markers else "other"] += 1

    citations: list[dict] = []
    bibliographies: list[dict] = []
    paragraphs = list(root.iter(f"{W}p"))
    for index, paragraph in enumerate(paragraphs):
        text = _paragraph_text(paragraph)
        for field in _paragraph_fields(paragraph):
            instruction = field["instruction"]
            if "ADDIN EN.CITE" in instruction:
                records = _parse_endnote_payload(instruction)
                citations.append(
                    {
                        "paragraph_index": index,
                        "paragraph_text": text,
                        "field_result_text": field["result_text"],
                        "instruction_length": len(instruction),
                        "records": records,
                        "payload_parsed": bool(records),
                    }
                )
            if "ADDIN EN.REFLIST" in instruction:
                bibliographies.append(
                    {
                        "paragraph_index": index,
                        "paragraph_text": text,
                        "field_result_text": field["result_text"],
                    }
                )

    balanced = (
        field_markers["begin"]
        == field_markers["separate"]
        == field_markers["end"]
    )
    return {
        "path": str(path.resolve()),
        "field_markers": field_markers,
        "field_markers_balanced": balanced,
        "endnote_citation_field_count": len(citations),
        "endnote_embedded_record_count": sum(
            len(item["records"]) for item in citations
        ),
        "endnote_bibliography_field_count": len(bibliographies),
        "citations": citations,
        "bibliographies": bibliographies,
    }


def _print_human(report: dict) -> None:
    print(f"Document: {report['path']}")
    print(
        "Field markers: "
        f"begin={report['field_markers']['begin']}, "
        f"separate={report['field_markers']['separate']}, "
        f"end={report['field_markers']['end']}, "
        f"balanced={report['field_markers_balanced']}"
    )
    print(f"EN.CITE fields: {report['endnote_citation_field_count']}")
    print(f"Embedded records: {report['endnote_embedded_record_count']}")
    print(f"EN.REFLIST fields: {report['endnote_bibliography_field_count']}")
    for number, field in enumerate(report["citations"], 1):
        print(f"\nCitation field {number} at paragraph {field['paragraph_index']}:")
        for record in field["records"]:
            print(
                "  - "
                f"{record.get('display_text') or '?'} "
                f"{record.get('author') or '?'} ({record.get('year') or '?'}) "
                f"record={record.get('record_number') or '?'}"
            )
            if record.get("title"):
                print(f"    {record['title']}")


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("docx", type=Path)
    parser.add_argument("--json", action="store_true", dest="as_json")
    args = parser.parse_args()
    try:
        report = inspect_docx(args.docx)
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
