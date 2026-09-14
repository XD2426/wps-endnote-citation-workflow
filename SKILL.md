---
name: wps-endnote-citation-workflow
description: Inspect, audit, and safely migrate EndNote Cite While You Write citations in WPS-authored DOCX manuscripts. Use when Codex must verify whether citations are real EN.CITE fields, detect EN.REFLIST bibliographies, distinguish manual numeric citations from dynamic citations, find duplicate manual and EndNote reference lists, validate a partial EndNote conversion, or diagnose WPS and EndNote workflow failures without modifying the original manuscript.
---

# WPS EndNote Citation Workflow

## Purpose

Verify EndNote fields structurally before drawing conclusions from WPS appearance or installation folders. Preserve the source manuscript, identify partial conversions, and separate document defects from Windows UI-control failures.

## Workflow

1. Work read-only unless the user explicitly requests edits. For edits, create a clearly named test copy first.
2. Run `scripts/inspect_endnote_fields.py` on the DOCX. Treat `ADDIN EN.CITE` as a dynamic citation and `ADDIN EN.REFLIST` as the EndNote-generated bibliography.
3. Run `scripts/audit_citations.py` to detect manual numeric citations, manual reference entries, duplicate bibliographies, and partial conversion.
4. Inspect the relevant rendered pages when layout matters. Prefer the documents skill renderer. If LibreOffice is unavailable and WPS COM is already installed, use WPS only to export a read-only temporary PDF; never overwrite the manuscript during verification.
5. Report field evidence, document status, and the exact next action. Do not infer plugin absence merely because no WPS-specific DLL or template is visible on disk.

## Conversion Rules

- Preserve the manual bibliography until every intended in-text citation has an `EN.CITE` field.
- Do not delete a manual list when the audit status is `partial_conversion`.
- After all citations are dynamic, remove the manual bibliography and keep one `EN.REFLIST` field.
- Re-run both scripts after every conversion batch and confirm that manual citation numbers and duplicate-list warnings are gone.
- Keep citation style changes inside EndNote or WPS CWYW. Do not imitate an EndNote field by inserting plain text.

## Windows Control Fallback

- Use the computer-use skill for WPS or EndNote UI actions.
- If screenshot capture fails with `SetIsBorderRequired failed: unsupported interface (0x80004002)`, treat this as a capture-layer problem, not an EndNote compatibility result.
- An accessibility-only snapshot may support read-only inspection, but do not click unnamed ribbon controls or use stale coordinates.
- Continue structural DOCX verification with the bundled scripts. Ask the user to perform the exact WPS click when safe UI targeting is unavailable.

## Commands

```powershell
python scripts/inspect_endnote_fields.py manuscript.docx
python scripts/inspect_endnote_fields.py manuscript.docx --json
python scripts/audit_citations.py manuscript.docx
python scripts/audit_citations.py manuscript.docx --json
```

Read `references/endnote-field-signatures.md` only when interpreting unusual field layouts or grouped citations.
