# Manuscript References

`manuscript-references` turns full bibliographic entries placed after manuscript sentences into consistently formatted in-text citations and a final bibliography. It also creates a reusable Excel reference workbook for later checking, correction, and reformatting.

## Input

- A `.docx` manuscript containing prose and inline full-reference blocks.
- A journal name, citation style, formatting guide, or representative citation and bibliography examples.

## Output

- A revised `.docx` with resolved full-reference blocks replaced by formatted citations that can be clicked to jump to their matching bibliography entries.
- A structured `.xlsx` workbook containing verified metadata, citation locations, format rules, and unresolved issues.

The workbook is the reusable source file. For a later style change, provide the latest manuscript, the workbook, and the new formatting requirement; verified records do not need to be researched again unless they changed or remain unresolved. Word links are internal navigation links rather than EndNote fields, so rerun the skill after changing citation records or style rules.

The default fast mode resolves each unique title or DOI once from an authoritative source and reserves deeper checking for ambiguous or conflicting records.

## Install as a plugin

The repository root is an installable Codex plugin. Its manifest is:

```text
.codex-plugin/plugin.json
```

The plugin exposes the skill from:

```text
skills/manuscript-references
```

Install the repository as a plugin, then invoke the bundled skill with `$manuscript-references`.

## Safety

The workflow never deletes an unresolved source block. Only an exact, successfully mapped block is replaced. Ambiguous and unresolved items remain in the manuscript and are recorded in the workbook.

## License

MIT
