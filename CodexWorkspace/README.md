# Codex Workspace

This workspace collects the essentials from `Revit_2023_Development_Guidelines_with_LLM_Integration.md` so Codex sessions can start quickly on Revit 2023 automation tasks.

## Quick Start
- Target **CPython 3.9** with pyRevit 5.2 (`#! python3` header on every script).
- Assume Autodesk Revit 2023 APIs and avoid legacy `Document.New*` patterns.
- Keep transactions explicit (`Transaction.Start()` / `Commit()`), and roll back inside `except` blocks.
- Convert Python iterables to .NET collections (`List[ElementId]`) at API boundaries.

## Directory Tour
- `docs/checklists` &mdash; ready-to-use review and testing checklists.
- `prompts` &mdash; curated prompt snippets for generating, reviewing, and documenting code with Codex or GPT.
- `scripts` &mdash; pyRevit-ready templates that follow the guidelines.

## Suggested Workflow
1. Read the relevant checklist before authoring or reviewing a tool.
2. Start from the script template in `scripts/new_command_template.py`.
3. Use the prompt snippets to guide Codex conversations about implementation or documentation.
4. Link back to the full guidelines for edge cases or deeper API notes.

## References
- Primary source: `../Revit_2023_Development_Guidelines_with_LLM_Integration.md`.
- Autodesk Revit 2023 API docs (lookup specifics when prompted).

