---
name: map-legal-provision-structure-skill
description: Build a traceable hierarchy and deterministic cross-reference graph from user-provided mainland-China normative legal texts. Use for laws, administrative regulations, judicial interpretations, departmental rules, or local normative documents supplied as folders, ZIPs, DOCX, PDF, TXT, Markdown, images, or scans when the user needs an offline interactive HTML structure/network graph, Chinese XLSX audit table, and scoped Mermaid diagrams. Do not search for missing laws, reconstruct amendments, determine validity, or infer unresolved references.
license: MIT
activation: /map-legal-provision-structure-skill
metadata:
  author: liusite
  version: 1.0.0
  created: 2026-08-27
  last_reviewed: 2026-08-27
provenance:
  maintainer: liusite
  version: 1.0.0
  created: 2026-08-27
  source_references:
    - user-provided mainland-China normative legal texts
---

# /map-legal-provision-structure-skill

Analyze only materials the user provides. Before a cloud Agent reads any material, obtain explicit informed confirmation and recommend redacted copies. Never use online conversion services or external legal databases.

## Workflow

1. Ask the user to select `structure` (hierarchy-first) or `network` (citation-first), supply a title, and label every submitted material/version. Do not infer version order or identity.
2. Run `prepare`. Plain TXT/Markdown is read directly. For other supported formats, the script detects MarkItDown and, if needed, installs it in a local managed virtual environment. It creates per-file Markdown and an index in a permission-restricted temporary workspace.
3. Read `legal-structure-index.json` before Markdown. Read only relevant chunks. The deterministic parser owns hierarchy and reference resolution; read `references/analysis-schema.md` only to record source identity corrections, numbering anomalies, and review items. Do not create semantic edges such as authorization, exception, definition, or legal basis.
4. Run `report` to build the offline HTML, Chinese XLSX, and scoped Mermaid files. Low-confidence OCR, ambiguous targets, external laws not supplied by the user, and malformed numbering must remain unresolved or pending review.

```bash
python3 scripts/run_pipeline.py prepare \
  --input FILE_OR_DIRECTORY... --query-json query.json \
  --workspace TEMP_DIR --processing-environment cloud --privacy-confirmed

python3 scripts/run_pipeline.py report \
  --bundle TEMP_DIR/legal-structure-bundle.json --analysis-json analysis.json \
  --output-dir OUTPUT_DIR --cleanup
```

## Boundaries

- `第X条` references resolve only when their target is unique. `本条`、`本款`、`前款`、`上一条` and `前条` resolve only when the parser can identify the required ancestor or sibling.
- A supplied law title may resolve an external reference only to one uniquely matched supplied document. Missing or conflicting material is recorded as `未解析引用`.
- Versions are independent document roots. This Skill does not compare amendments or determine an effective version.
- The output describes textual structure and references in the supplied sample. It is not a legal opinion, validity check, or complete legal database.

## Gotchas

- MarkItDown conversion quality and OCR reliability affect source locations; OCR-derived nodes and references require review.
- A full statute is intentionally split into scoped Mermaid files because a whole-statute Mermaid graph is usually unreadable.
