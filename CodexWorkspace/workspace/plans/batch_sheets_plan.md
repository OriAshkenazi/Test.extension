#! codex-env plan

# Batch Sheets Export Tool (Revit 2019 + Autodesk Docs / BIM 360)

## Executive Summary
Develop a pyRevit 5.3.1 (**IronPython 2.7**) toolset for **Revit 2019** models that are available via Autodesk Docs / BIM 360 and/or are accessible **locally** (e.g., via Desktop Connector) to export the most up-to-date “תוכניות לביצוע לפי דיסיפלינות” sheets for each Base Darom structure. For every structure, export from five discipline models (Architecture, Structure, HVAC, Plumbing, Electrical) and generate **one DWG + one PDF per sheet**. Because the provided “coding names” don’t reliably match sheet names/numbers, the default workflow exports **all sheets that match the BIM folder filter** and produces detailed reports for later filtering. DWG exports must be **unified per sheet (no external references/xrefs)** using a standardized DWG export setup. PDF output in Revit 2019 is implemented via **PrintManager (print-to-PDF)** with a required paper size of **900×1800 mm** (covers long sheets). The solution includes validation, **batch orchestration via `pyrevit run`**, resumable exports, and audit reports.

## How Users Run It (Interactive + Batch)
### One-time setup (per machine / per project)
1) **Install/attach the extension** on the export machine (same machine that has Revit 2019 + pyRevit installed).
2) **Make models available locally** from Autodesk Docs:
   - Ensure each target `.rvt` is fully downloaded/hydrated locally (Desktop Connector “online-only” files can fail to open in unattended runs).
   - Keep all models on the same Revit major version (2019) to avoid upgrade prompts during batch runs.
3) In Revit, run **Validate Setup** once per discipline template/version to confirm:
   - The required DWG export setup name exists and produces “no-xref” outputs.
   - PDF printing prerequisites are satisfied (PDF printer exists, silent print-to-file works, and the **900×1800 mm** paper size is available).
4) Configure `config.json` (sheet-folder filter, DWG setup name, PDF options, output root, resume/overwrite policy).

### Optional: Build a “Sheet List” to choose what’s important (recommended first step)
If the team can’t precisely define the target set up-front, start by generating an inventory of **all sheets that exist** across the provided model list, then mark which sheets should be exported.

1) Run **Build Sheet List** against the same model list used for batch export.
2) The tool produces a CSV/Excel-friendly file (one row per sheet per model) with columns like:
   - Model path, discipline, structure (inferred), sheet number, sheet name, and key classification parameters.
3) A coordinator marks “export = yes/no” (or adds priority/notes) in the sheet list file.
4) Export runs in “selection file” mode, exporting only the marked sheets (with reporting for “missing in model” cases).

### Interactive (inside Revit, single model)
1) Open a discipline model (local Autodesk Docs path is fine).
2) Click **Export Sheets**.
3) Pick output root and run mode (export all folder sheets vs filter-by-codes, resume vs overwrite).
4) Review the per-model report outputs and spot-check a few exported DWGs/PDFs.

### Batch (CLI runner over many local models)
This mode is intended for unattended exports across many `.rvt` files (multiple disciplines and/or multiple structures) using the built-in pyRevit runner.

1) Create a plain text file (one full model path per line), e.g. `models_2019.txt` pointing at local Autodesk Docs files.
2) Run the batch script with pyRevit CLI:
   - `pyrevit run "<path-to-runner-script.py>" --models="<path-to-models_2019.txt>" --revit=2019 --purge`
   - Use `--allowdialogs` only for troubleshooting; production runs should be dialog-free.
3) The script runs once per model, auto-detects discipline/structure (from filename/path and/or Project Information), exports matching sheets, and writes reports + checkpoints so reruns can resume safely.

### Outputs (what the user gets)
- Deterministic folder structure, e.g. `/<Structure>/<Discipline>/DWG/*.dwg` and `/<Structure>/<Discipline>/PDF/*.pdf`.
- A sheet inventory/selection file when using **Build Sheet List** (CSV intended for Excel marking).
- Reports alongside outputs (`report.model.*`, `report.structure.*`, `errors.log`) to support post-filtering and QA.
- **Live status files during long runs** (written incrementally as each sheet completes/fails):
  - `progress.json` (current model/sheet + percent complete)
  - `failed_sheets.csv` (append/update immediately on failure; not only at end)

## System Architecture with module breakdown
### Architectural principles
- **Revit-API-safe execution**: no long-running work inside transactions; explicit `Transaction` only when required.
- **Modern Revit API coding patterns** (compatible with 2019): explicit transactions, explicit failure handling, and **.NET collections at API boundaries** (`List[ElementId]`, `ViewSet`, `ICollection<ElementId>`).
- **Deterministic outputs**: stable naming, stable ordering, resumable runs.

### Component diagram (logical)
UI (pyRevit pushbuttons) / CLI (`pyrevit run`) → Batch Runner → (Sheet Query → DWG Exporter, PDF Printer) → Reporting

### pyRevit extension layout
- `BaseDarom.exporter.tab/Export Plans.panel/`
  - `Export Sheets` (main)
  - `Validate Setup` (pre-flight)
  - `Build Sheet List` (optional; generates sheet inventory for marking)
  - `Build Manifest` (optional; model list helper)
- `lib/`
  - `config.py`
  - `sheet_query.py`
  - `sheet_inventory.py`
  - `export_dwg.py`
  - `export_pdf.py`
  - `batch_runner.py`
  - `bim360_adapter.py`
  - `reporting.py`
  - `ui.py`
  - `runner_entry.py` (CLI-friendly entrypoint)
  - `utils/` (`path.py`, `revit.py`, `netcollections.py`)

### Module responsibilities and documentation requirements
1) `config.py`
- Loads/saves `config.json` with defaults + validation.
- Keys: sheet-folder filter, optional code-parameter name, DWG setup name, PDF printer + paper size, output templates, overwrite/resume mode.
- Docs: `docs/config.schema.md` with examples and upgrade notes.

2) `sheet_query.py`
- Collects `ViewSheet` targets:
  - Exclude placeholders (`ViewSheet.IsPlaceholder`).
  - Filter by configured parameter(s) that implement “BIM folder” (value: `"תוכניות לביצוע לפי דיסיפלינות"`).
  - Optional: filter by code list (CSV/Excel) and/or a configured “sheet code” parameter.
  - Optional: filter by **sheet selection file** generated by `sheet_inventory.py` (export=yes).
- Returns **ordered** list of sheets and a `.NET List[ElementId]` when required.
- Docs: `docs/sheet-selection.md` (how filtering works; how to configure parameter names).

2a) `sheet_inventory.py` (sheet list generator)
- Reads a model list (same input as `pyrevit run --models=...`) and exports a **flat sheet inventory** (CSV) containing sheet number/name + relevant parameters per model.
- Supports a coordinator “mark-up” workflow (columns like `export_pdf`, `export_dwg`, `priority`, `notes`).
- Docs: `docs/sheet-inventory.md` (how to generate; how to mark; how missing sheets are handled).

3) `export_dwg.py`
- Exports one DWG per sheet using `DWGExportOptions.GetPredefinedOptions(doc, setup_name)`.
- Enforces “no external references” by standardizing the Revit export setup and verifying outputs (no xref by-products).
- Docs: `docs/dwg-export-setup.md` (how to create/verify the required setup in Revit UI).

4) `export_pdf.py` (Revit 2019 constraint)
- Creates PDFs via `PrintManager` (print-to-file), one PDF per sheet.
- Requires a PDF printer configured for silent print-to-file and a paper size named/configured as **900×1800 mm**.
- Uses explicit transactions only if creating a temporary `ViewSheetSet` or persistent print setting; otherwise uses ephemeral `ViewSet` assignment.
- Docs: `docs/pdf-printing-prereqs.md` (supported printers, required paper size, troubleshooting).

5) `bim360_adapter.py`
- Extracts document identity and metadata (cloud/local, title, discipline inference, structure code inference from filename and/or Project Information parameters).
- Optional feature flag: open BIM 360 models from a manifest (GUID-based) if the Revit 2019 API supports it in the environment; otherwise operates on currently open docs or local model paths.
- Docs: `docs/bim360-manifest.md`.

6) `batch_runner.py`
- Orchestrates per-structure export across 5 models:
  - Mode A (default): export the **active document**.
  - Mode B: export a **set of currently open documents**.
  - Mode C (optional): export from `manifest.json` (structures → disciplines → model identifiers).
- Writes checkpoints after each sheet; supports resume (skip files that exist and match expected size > 0).
- **Silent per-sheet failure handling**: never block long runs with modal dialogs; catch exceptions per sheet and continue (failures are recorded).
- **Live progress + failure reporting**: update `progress.json` and `failed_sheets.csv` incrementally after each sheet attempt.
- Docs: `docs/batch-runner.md`.

7) `reporting.py`
- Emits:
  - `report.model.json` + `report.model.csv` (per model)
  - `report.structure.json` + `report.structure.csv` (per structure aggregation)
  - `errors.log` with stack traces
- Docs: `docs/reports.md`.

8) `ui.py`
- pyRevit wizard:
  - output root selection
  - structure code/name (or from manifest)
  - discipline selection
  - mode: “export all folder sheets” vs “filter by codes” vs “use selection file”
  - resume/overwrite toggle
- Clear progress bar per run (per model; per sheet count) + cancellation checkpoints between sheets.
- Docs: embedded help + `README.md` runbook.

## gpt-5.2-codex Integration Strategy with task categories
### Category 1: Scaffolding and conventions
Codex generates the pyRevit extension skeleton, shared lib imports, logging bootstrap, and standard command entrypoints.
- Example tasks Codex can implement:
  - “Create pyRevit pushbutton scripts with shared `lib` imports and structured logging.”
  - “Add a config loader with schema validation and default generation.”

### Category 2: Revit API building blocks
Codex implements small, testable API units with explicit transactions and .NET collection boundaries.
- Example tasks:
  - “Implement `collect_target_sheets(doc, cfg)` using `FilteredElementCollector` and parameter matching.”
  - “Implement `export_sheet_dwg(doc, sheet_id, out_dir, setup_name)` using predefined DWG export options.”
  - “Implement `print_sheet_pdf(doc, sheet, printer, paper, out_path)` using `PrintManager` and `ViewSet`.”

### Category 3: Batch orchestration and resiliency
Codex implements retry/resume logic, per-sheet try/except, and structured result objects.
- Example tasks:
  - “Implement `run_structure(structure_id, discipline_docs, cfg)` writing checkpoints after each sheet.”
  - “Implement resume: skip if PDF/DWG exists and file size > 0; otherwise re-export.”
  - “Write live status outputs: `progress.json` + `failed_sheets.csv` updated after each sheet.”

### Category 4: Diagnostics and validation
Codex builds a pre-flight validator to reduce field failures.
- Example tasks:
  - “Validate printer exists; list available paper sizes; fail with actionable message.”
  - “Validate DWG export setup exists; run a temp export and assert no xrefs were produced.”

### Category 5: Documentation and test generation
Codex generates docs and tests for non-Revit logic.
- Example tasks:
  - “Write `docs/pdf-printing-prereqs.md` and troubleshooting decision tree.”
  - “Generate unit tests for filename sanitization and config validation.”

## Quality Assurance Framework
### Automated (non-Revit) tests
- Unit tests for:
  - filename sanitization (Windows invalid chars, reserved names, max path)
  - config validation and defaulting
  - manifest parsing and structure/discipline resolution
  - sheet code matching (normalization, duplicates)
- Static checks:
  - import hygiene (no Revit API imports in pure utility modules)
  - logging coverage (every failure path writes a structured error)

### In-Revit validation and acceptance tests
- `Validate Setup` command must pass before production use:
  - DWG export setup exists and produces “no-xref” artifacts.
  - PDF printer and required paper size exist.
- Pilot acceptance on one structure (e.g., SK) across five models:
  - All sheets in the folder are exported (or clearly reported failed).
  - DWGs have no external references (artifact verification).
  - Output directory structure is deterministic:
    - `/<Structure>/<Discipline>/DWG/*.dwg`
    - `/<Structure>/<Discipline>/PDF/*.pdf`
  - Reports exist and match exported counts.

## Implementation Milestones with timeline
Week 1
- Scaffold extension + logging + config schema.
- Implement `Validate Setup` (printer, paper size, DWG setup presence).

Week 2
- Implement `sheet_query.py` (folder filter + placeholder exclusion + ordering).
- Implement output naming + path sanitization + collision handling.
- Implement `sheet_inventory.py` (generate sheet list CSV for coordinator marking).

Week 3
- Implement DWG export using predefined setup + no-xref verification.
- Add per-sheet result recording (JSON/CSV).

Week 4
- Implement PDF printing via `PrintManager` with 900×1800 mm.
- Add printer diagnostics and common failure remediation.

Week 5
- Implement batch runner (active doc + open docs + optional manifest).
- Add resume/retry logic, progress UI, and live failure log updates during long runs.

Week 6
- Hardening: performance, edge cases (empty folder, missing parameters, long names).
- Full pilot run + documentation finalization + handover checklist.

## Risk Assessment and Mitigation
1) PDF export automation limits (Revit 2019)
- Risk: print-to-PDF driver prompts dialogs or ignores paper size.
- Mitigation: require a silent PDF driver configuration; Validate Setup enumerates and confirms the paper size; provide a documented list of supported drivers/settings.

2) DWG external references
- Risk: xrefs appear if export setup differs across models.
- Mitigation: enforce one standardized setup name; Validate Setup performs a temp export and checks for xref by-products; fail fast with steps to fix the setup.

3) Autodesk Docs local file availability (batch runner)
- Risk: Desktop Connector paths point to files that are not hydrated locally, moved/renamed, or locked by another process/user.
- Mitigation: pre-flight checks existence + accessibility; manifest/model list is validated before launch; per-model failures are logged and don’t stop the batch.

4) Sheet coding vs actual Revit sheet metadata mismatch
- Risk: cannot safely export only a provided code list.
- Mitigation: default exports all sheets under the folder; optional filtering by a configured “sheet code” parameter; always produce reports to enable post-filtering.

5) Long runs and crash/restart scenarios
- Risk: large exports may be interrupted.
- Mitigation: per-sheet checkpointing, resume mode, and deterministic naming so reruns are safe and incremental.
