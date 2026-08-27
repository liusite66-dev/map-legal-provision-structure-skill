#!/usr/bin/env python3
import ast
import sys
from pathlib import Path


def main() -> int:
    root = Path(sys.argv[1] if len(sys.argv) > 1 else Path(__file__).resolve().parents[1])
    required = ["SKILL.md", "AGENTS.md", "agents/openai.yaml", "scripts/run_pipeline.py", "references/analysis-schema.md", "tests/test_pipeline.py"]
    missing = [item for item in required if not (root / item).is_file()]
    if missing:
        print("Missing: " + ", ".join(missing), file=sys.stderr)
        return 2
    for script in (root / "scripts").glob("*.py"):
        ast.parse(script.read_text(encoding="utf-8"))
    skill = (root / "SKILL.md").read_text(encoding="utf-8")
    if "TODO" in skill or "## Gotchas" not in skill:
        print("SKILL.md has TODOs or no Gotchas section", file=sys.stderr)
        return 2
    print("PIPELINE OK")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
