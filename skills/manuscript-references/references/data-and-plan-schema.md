# Data and plan schema

Read this file when creating or reusing the structured data that drives the Word and Excel outputs.

## Stable identifiers

Each unique work has a stable `Ref_ID`:

- use `REF-` plus the first ten uppercase hexadecimal characters of SHA-1 over the normalized verified DOI when present;
- otherwise hash the normalized verified title;
- reuse an existing workbook's `Ref_ID` for the same work.

Each occurrence has an `Occurrence_ID`, normally `CIT-0001`, `CIT-0002`, and so on in document order.

## Master JSON

The workbook builder accepts:

```json
{
  "references": [
    {
      "ref_id": "REF-A1B2C3D4E5",
      "type": "journal_article",
      "original_reference": "Complete text found after the sentence",
      "original_title": "Title as supplied",
      "verified_title": "Authoritative title",
      "authors": "Smith, Jane Q.; Jones, Alex",
      "year": "2025",
      "source_title": "Journal Name",
      "volume": "12",
      "issue": "3",
      "pages_or_article_number": "101-119",
      "doi": "10.1234/example",
      "url": "https://doi.org/10.1234/example",
      "verification_source": "https://publisher.example/item",
      "status": "verified",
      "manual_review": "pending",
      "notes": ""
    }
  ],
  "citation_map": [
    {
      "occurrence_id": "CIT-0001",
      "location_id": "body.p000012",
      "section": "Introduction",
      "context": "Sentence supported by the source.",
      "original_text": "Complete bibliographic block to replace",
      "ref_ids": ["REF-A1B2C3D4E5"],
      "formatted_citation": "(Smith & Jones, 2025)",
      "status": "resolved",
      "notes": ""
    }
  ],
  "format_rules": {
    "requested_style": "Verbatim user request or journal name",
    "citation_system": "author-year",
    "cluster_order": "alphabetical",
    "bibliography_order": "alphabetical",
    "et_al_rule": "three or more authors",
    "bibliography_heading": "References",
    "bibliography_page_break": true,
    "bibliography_hanging_indent_in": 0.5,
    "bibliography_line_spacing": 1.0,
    "assumptions": ""
  },
  "issues": [
    {
      "issue_id": "ISS-0001",
      "location_id": "body.p000018",
      "original_text": "Unresolved source block",
      "issue_type": "unresolved",
      "description": "No authoritative title match",
      "required_action": "Manual confirmation",
      "status": "open"
    }
  ]
}
```

Valid reference statuses are `verified`, `user_approved`, `ambiguous`, and `unresolved`. Valid occurrence statuses are `resolved`, `ambiguous`, and `unresolved`.

## Citation plan JSON

The Word editor accepts:

```json
{
  "replacements": [
    {
      "occurrence_id": "CIT-0001",
      "location_id": "body.p000012",
      "original_text": "Exact full-reference text in the paragraph",
      "replacement": "(Smith & Jones, 2025)",
      "ref_ids": ["REF-A1B2C3D4E5"],
      "status": "resolved"
    }
  ],
  "bibliography": {
    "heading": "References",
    "page_break": true,
    "hanging_indent_in": 0.5,
    "line_spacing": 1.0,
    "space_after_pt": 0,
    "entries": [
      {
        "ref_id": "REF-A1B2C3D4E5",
        "text": "Smith, J. Q., & Jones, A. (2025). ..."
      }
    ]
  }
}
```

Only `resolved` replacements are applied. `original_text` must match exactly within `location_id`; otherwise the editor reports failure and leaves the paragraph unchanged.

For one cited work, the editor links the entire `replacement` to the bibliography entry identified by the single value in `ref_ids`. For a cluster, segment the citation so each work has its own target:

```json
{
  "occurrence_id": "CIT-0002",
  "location_id": "body.p000019",
  "original_text": "Two full references to replace",
  "ref_ids": ["REF-A1B2C3D4E5", "REF-F6E7D8C9B0"],
  "replacement_runs": [
    {"text": "("},
    {"text": "Smith, 2025", "ref_id": "REF-A1B2C3D4E5"},
    {"text": "; "},
    {"text": "Jones, 2026", "ref_id": "REF-F6E7D8C9B0"},
    {"text": ")"}
  ],
  "status": "resolved"
}
```

Each bibliography entry must include the matching `ref_id`. The editor creates a stable Word bookmark and internal hyperlink without imposing blue or underlined hyperlink styling.

## Workbook sheets

### References

One row per unique work. Preserve the supplied full reference alongside verified fields so a reviewer can compare them.

### Citation_Map

One row per occurrence. Store `Ref_IDs` as semicolon-separated values in occurrence order.

### Format_Rules

Store the user's request verbatim and every applied rule or assumption. This sheet lets a later run distinguish a metadata correction from a style-only change.

### Issues

Store unresolved, ambiguous, extraction, replacement, and validation issues. Never hide a failed replacement by omitting it from this sheet.
