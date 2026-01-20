# Codex Agent Guide (Repository Root)

This repo includes a curated knowledge base under `CodexWorkspace/` that Codex should use as the **default reference** for Revit 2023 + pyRevit work.

## Start Here
- Primary entrypoint: `CodexWorkspace/README.md`
- Docs index: `CodexWorkspace/docs/README.md`
- Checklists: `CodexWorkspace/docs/checklists/`
- Prompt snippets: `CodexWorkspace/prompts/`
- Script templates/harnesses: `CodexWorkspace/scripts/`

## How To Use `CodexWorkspace`
1. **Before coding**: open the relevant doc(s) from `CodexWorkspace/docs/` (especially the Revit 2023 guideline reference and the checklists).
2. **When creating a new pyRevit command**: start from `CodexWorkspace/scripts/new_command_template.py` and adapt it to the target tool folder.
3. **When reviewing or debugging**: use the checklists in `CodexWorkspace/docs/checklists/` and the prompt snippets in `CodexWorkspace/prompts/` to standardize analysis and outcomes.
4. **When unsure about patterns**: prefer the Revit 2023 guidance under `CodexWorkspace/docs/reference/` over ad-hoc API usage.

## Repo Layout Notes (pyRevit Extensions)
- Commands typically live under: `<Tab>.tab/<Panel>.panel/<Button>.pushbutton/script.py`
- Keep transactions explicit (`Start/Commit`, `RollBack` on exceptions).
- Convert Python iterables to .NET collections at Revit API boundaries when required (e.g., `List[ElementId]`).

