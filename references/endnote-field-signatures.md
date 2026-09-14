# EndNote Field Signatures

## Dynamic citation

An EndNote Cite While You Write citation normally uses a complex Word field whose instruction begins with:

`ADDIN EN.CITE <EndNote>...`

The embedded XML commonly contains `Author`, `Year`, `RecNum`, `DisplayText`, and one or more `record` elements. The visible citation is the field result between the `separate` and `end` field markers.

## Dynamic bibliography

The EndNote-generated bibliography uses:

`ADDIN EN.REFLIST`

Its visible references are cached field results. If a manual reference list remains immediately before this field, the document is usually only partially converted.

## Interpretation

- `EN.CITE` present and `EN.REFLIST` absent: citations exist but the bibliography may not have been updated.
- `EN.CITE` and `EN.REFLIST` present, with manual numeric citations or manual reference entries: partial conversion.
- Multiple `EN.REFLIST` fields: duplicate EndNote bibliographies or damaged field layout.
- Visible citation text without `EN.CITE`: plain text, not an updateable EndNote citation.
- Unbalanced `begin`, `separate`, and `end` marker counts: investigate possible field corruption before editing.

Grouped citations can contain multiple `Cite` elements inside one `EN.CITE` instruction. Count both citation fields and embedded records.
