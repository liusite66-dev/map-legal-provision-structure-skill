# Regression Criteria

- The package has a valid Skill structure and deterministic pipeline syntax.
- A generated report must contain the HTML, XLSX, and scoped Mermaid artifact paths.
- A report must retain unresolved references instead of inventing targets.
- Cloud processing without explicit confirmation must terminate before input reading.

The synthetic cases below are input-only baselines. They cover hierarchy parsing, explicit and relative citation handling, and unresolved external references. The third is a holdout case.

```json
{
  "skill": "map-legal-provision-structure-skill",
  "criteria": [
    {"id": "pipeline-check", "text": "The package pipeline check exits successfully", "type": "command", "cmd": "python3 scripts/check_pipeline.py ."},
    {"id": "output-present", "text": "A reported artifact path exists", "type": "command", "cmd": "test -e {output}"},
    {"id": "no-guessed-target", "text": "Ambiguous or absent targets are retained as unresolved references", "type": "llm-judge"}
  ],
  "golden": [
    {"id": "hierarchy", "input": "golden/hierarchy/input.md", "expected": null, "expected_status": "pending-first-green", "split": "val", "compare": "none"},
    {"id": "cross-reference", "input": "golden/cross-reference/input.md", "expected": null, "expected_status": "pending-first-green", "split": "val", "compare": "none"},
    {"id": "unresolved-holdout", "input": "golden/unresolved-holdout/input.md", "expected": null, "expected_status": "pending-first-green", "split": "test", "compare": "none"}
  ],
  "judge": {"model": "claude-sonnet-4-20250514"}
}
```
