---
name: manuscript-references
description: Convert full bibliographic entries placed after manuscript sentences into formatted in-text citations and a final bibliography, while creating a reusable Excel master reference file. Use for DOCX manuscripts whose references need extraction, title-based verification, deduplication, citation replacement, bibliography generation, auditing, or later reformatting without EndNote. Do not use to discover new literature unless the user separately requests it.
---

# Manuscript References

Produce a publication-ready manuscript and a reusable reference workbook in one pass. The workbook, not the rendered bibliography, is the durable source for later checking and style changes. In the revised Word file, each resolved in-text citation must be an internal hyperlink to its corresponding bibliography entry.

## Default contract

Normally require only:

1. the manuscript `.docx`, including every full reference currently placed after its supporting sentence; and
2. the target journal, style guide, or citation and bibliography examples.

Use existing sources only. Do not add literature unless explicitly asked. Preserve the original file and write a new manuscript.

## Evidence and deletion rules

Treat the work title as the only trusted search anchor unless the user explicitly identifies another authoritative field. Supplied author, year, source, pages, DOI, and URL are clues until verified.

- Replace a full-reference block only after it maps to a verified or user-approved record and the exact source text is found at the intended occurrence.
- Keep ambiguous or unresolved blocks unchanged in the manuscript and record them in the `Issues` sheet.
- Never remove prose merely because it resembles a reference.
- Deduplicate by verified DOI, then verified normalized title. Preserve distinct versions, editions, corrections, translations, and conference/journal forms unless identity is established.

## Fast mode (default)

Optimize for low user effort, elapsed time, and token use:

- Trust only an exact-title or exact-DOI authoritative match. One authoritative metadata source is sufficient when its identity and core fields are internally consistent.
- Search and resolve unique titles in batches. Do not open multiple sources merely to reconfirm an already exact match.
- Perform deeper cross-checking only for ambiguous, conflicting, incomplete, retracted/corrected, or unresolved records.
- Mark exceptions for manual review instead of delaying the entire document. Use strict multi-source verification only when the user requests it.
- Reuse an existing master workbook without re-verifying unchanged `verified` or `user_approved` records.

## Token-efficient workflow

1. Run `scripts/docx_inventory.py` to extract only likely reference-bearing paragraphs plus local context. Do not paste the whole manuscript into chat when the structured inventory is sufficient.
2. Build one unique-title registry. Verify each unique work once, preferably in batches, then reuse that record for every occurrence.
3. Read [references/data-and-plan-schema.md](references/data-and-plan-schema.md). Assign deterministic `Ref_ID` values and prepare the master JSON and citation plan.
4. Parse the requested style once. Store the verbatim request and the applied rules in `Format_Rules`. Generate final in-text strings and bibliography entries from the verified registry.
5. Run `scripts/apply_citation_plan.py` to replace exact source blocks, create clickable internal citation links, and append bookmarked bibliography entries. Do not rewrite unaffected paragraphs.
6. Run `scripts/build_reference_workbook.py` to create the source workbook.
7. Run `scripts/validate_reference_package.py`, then render and inspect the final DOCX and visually check the workbook. Resolve any validation failures before delivery.

If the user supplies a prior workbook, enter reuse mode: treat user-approved fields as authoritative, verify only new, changed, ambiguous, or unresolved records, and preserve existing `Ref_ID` values.

## Citation formatting

Determine from the supplied requirements:

- author-year, numbered, note-based, or another citation system;
- parenthetical versus narrative forms;
- author truncation and `et al.` rules;
- citation-cluster ordering and delimiters;
- locator, prefix, and suffix handling;
- bibliography order, author-name rules, title capitalization, source styling, DOI/URL treatment, journal abbreviations, indentation, and spacing.

When requirements remain incomplete, prefer an identified journal/style standard. Ask only when the missing choice would materially change the result. Record any necessary assumption in `Format_Rules`.

## Word links

- Give every bibliography entry its stable `Ref_ID`; the editor converts it to a Word bookmark.
- For a single-work citation, `ref_ids` is sufficient and the whole citation links to that work.
- For a citation cluster, use `replacement_runs` and put `ref_id` on each author-year or number segment. Leave parentheses, separators, prefixes, suffixes, and locators unlinked unless they belong to one work.
- Keep links visually identical to the surrounding citation text; do not force blue or underlining.
- These are navigation links, not EndNote fields. If citations, records, or style rules change, rerun the skill from the workbook to synchronize the Word file.

## Deliverables

Return only:

- `<manuscript>_references_formatted.docx`;
- `<manuscript>_reference_master.xlsx`.

The workbook contains `References`, `Citation_Map`, `Format_Rules`, and `Issues`. Do not paste the full bibliography or workbook contents into the chat. Report only counts for unique works, replaced occurrences, unresolved issues, and validation status.

## Quality requirements

- Every replacement appears in `Citation_Map` and points to one or more `Ref_ID` values.
- Every resolved citation has a valid internal link for each cited `Ref_ID`, and every link target is a bookmarked bibliography entry.
- Every bibliography entry is cited unless the user requests otherwise.
- Every cited, resolved record appears once in the bibliography.
- Bibliography order matches the requested style.
- All unresolved source blocks remain present in the manuscript.
- The revised DOCX preserves non-reference text, figures, tables, equations, section order, and unrelated formatting.
- The workbook retains original text, verified metadata, evidence URL, status, all occurrence locations, the applied format rules, and notes.

For a small request or an already verified workbook, skip unnecessary research and produce proportionate outputs.
