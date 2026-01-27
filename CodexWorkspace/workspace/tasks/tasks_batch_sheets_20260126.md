# Batch Sheets Export Tool — Implementation Tasks

Repository: https://github.com/OriAshkenazi/Test.extension.git  
Development Guide: `CodexWorkspace/workspace/plans/batch_sheets_plan.md`  
Target runtime: Revit 2019 + pyRevit 5.3.1 (IronPython 2.7)

## Reference Snapshot (from the plan)
- Export “most up-to-date” construction sheets per Base Darom structure across 5 disciplines (Arch/Str/HVAC/Plumbing/Electrical).
- Models may be on Autodesk Docs/BIM 360 and/or local (Desktop Connector). Batch runs must assume models are hydrated locally.
- DWG export: unified per sheet (no xrefs) using a standardized predefined DWG export setup.
- PDF export: Revit 2019 PrintManager print-to-PDF with required paper size 900×1800 mm.
- Default selection: export all sheets matching a BIM folder filter; always produce detailed reports.
- Optional selection workflow: Build Sheet List CSV -> coordinators mark sheets -> selection-file mode exports only marked sheets.
- Batch orchestration: `pyrevit run` across many models; resumable exports; per-sheet try/except; checkpointing; deterministic outputs.
- Live status files updated during runs: `progress.json`, `failed_sheets.csv`.
- Config-driven: `config.json` (filters, DWG setup name, PDF options, output root, resume/overwrite).

## Proposed Repo File Layout (pyRevit-friendly)
- UI entrypoints (pushbuttons)
  - `BaseDarom.exporter.tab/Export Plans.panel/Export Sheets.pushbutton/script.py`
  - `BaseDarom.exporter.tab/Export Plans.panel/Validate Setup.pushbutton/script.py`
  - `BaseDarom.exporter.tab/Export Plans.panel/Build Sheet List.pushbutton/script.py`
  - `BaseDarom.exporter.tab/Export Plans.panel/Build Manifest.pushbutton/script.py` (optional)
- Shared library (IronPython 2.7)
  - `lib/` (package)
  - `lib/utils/` (`path.py`, `csv_unicode.py`, `sorting.py`, `netcollections.py`, `revit.py`, `model_list.py`)
  - `lib/config.py`, `lib/sheet_query.py`, `lib/sheet_inventory.py`
  - `lib/export_dwg.py`, `lib/export_pdf.py`
  - `lib/reporting.py`, `lib/checkpoint.py`, `lib/batch_runner.py`, `lib/runner_entry.py`, `lib/ui.py`
  - `lib/manifest.py` (optional), `lib/validate.py`
- Project docs
  - `docs/config.schema.md`, `docs/sheet-selection.md`, `docs/sheet-inventory.md`
  - `docs/dwg-export-setup.md`, `docs/pdf-printing-prereqs.md`
  - `docs/batch-runner.md`, `docs/reports.md`, `docs/bim360-manifest.md`
- Workspace assets (examples)
  - `CodexWorkspace/workspace/batch_sheets/config.json`
  - `CodexWorkspace/workspace/batch_sheets/models_2019.txt`
  - `CodexWorkspace/workspace/batch_sheets/manifest.json` (optional)

---

Task [1]: Draft Development Plan (Project Docs) (Docs)

Parallel Development Note: Blocking prerequisite; it provides the canonical architecture/milestones referenced by all implementation tasks.

Context
`CodexWorkspace/workspace/development.md` is currently a placeholder; the authoritative feature plan is `CodexWorkspace/workspace/plans/batch_sheets_plan.md`.

Objective
Replace the placeholder development plan with an implementation-ready project development plan derived from the batch sheets plan.

Technical Requirements
1. Primary Component - Produce a complete development plan (goals, non-goals, user flows, module map, milestones, risks).
2. Integration Points - Map pushbuttons and the CLI runner (`pyrevit run`) to shared `lib/` modules.
3. Data/API Specifications - Summarize config schema, output structure, report schemas, checkpoint/resume rules, and live status files.
4. Error Handling - Document resiliency expectations (per-sheet try/except, checkpointing, deterministic reruns, dialog avoidance).

Files to Create/Modify
* `CodexWorkspace/workspace/development.md` - Replace placeholder with Batch Sheets Exporter project plan.

Implementation Specifications
- Include: Goals, Non-Goals, User Flows (Interactive + Batch), Module Responsibilities, Config Summary, Output Layout, Resume/Overwrite semantics, QA strategy, Pilot acceptance checklist.
- Keep language specific to Revit 2019 + pyRevit 5.3.1 (IronPython 2.7).

Success Criteria
* `CodexWorkspace/workspace/development.md` contains no placeholder text and is specific to this project.
* The document defines deterministic output paths and acceptance criteria from the plan.
* The document defines a dependency-aware milestone sequence aligned with the tasks in this file.
* All tests pass: N/A (documentation-only)

Documentation Updates Required
* Update `README.md` section “Codex Workspace” with a link to `CodexWorkspace/workspace/development.md`.
* Update `AGENTS.md` with any new conventions introduced by the Batch Sheets tool (paths, naming, runner modes).
* Create/update docstrings for all new functions: N/A (documentation-only)
* Update API documentation if applicable: N/A

Validation Commands
bash
rg -n "Replace this file" CodexWorkspace/workspace/development.md
rg -n "Revit 2019|pyRevit 5\\.3\\.1|IronPython 2\\.7|900" CodexWorkspace/workspace/development.md

Constraints
* Maximum implementation time: 3 hours
* Target code volume: ~100 lines
* Must maintain backward compatibility
* Follow existing code patterns in `CodexWorkspace/docs/agent/pyrevit/script-architecture.md`

---

Task [2]: Scaffold Batch Sheets UI Bundles + Shared `lib/` Package (Scaffolding)

Parallel Development Note: Can run concurrently with pure-doc tasks; scaffolding enables parallel implementation across modules without blocking on UI creation later.

Context
The plan specifies a new tab/panel with pushbuttons plus a shared `lib/` package. The repo currently does not contain this Batch Sheets tool layout.

Objective
Create the folder/file scaffolding for the Batch Sheets tools and their shared library modules.

Technical Requirements
1. Primary Component - Add tab/panel/pushbutton folder structure and thin `script.py` entrypoints that delegate to `lib/`.
2. Integration Points - Ensure imports resolve correctly under pyRevit and do not rely on Python 3-only stdlib.
3. Data/API Specifications - Create doc stubs under `docs/` matching the plan filenames.
4. Error Handling - Pushbuttons must guard against “not in Revit” execution and handle user cancel paths cleanly.

Files to Create/Modify
* `BaseDarom.exporter.tab/Export Plans.panel/_layout` - Order buttons: Validate Setup, Build Sheet List, Export Sheets, Build Manifest.
* `BaseDarom.exporter.tab/Export Plans.panel/Export Sheets.pushbutton/script.py` - Thin orchestrator calling `lib/ui.py`.
* `BaseDarom.exporter.tab/Export Plans.panel/Validate Setup.pushbutton/script.py` - Thin orchestrator calling `lib/validate.py` (or `lib/ui.py` render helper).
* `BaseDarom.exporter.tab/Export Plans.panel/Build Sheet List.pushbutton/script.py` - Thin orchestrator calling `lib/sheet_inventory.py`.
* `BaseDarom.exporter.tab/Export Plans.panel/Build Manifest.pushbutton/script.py` - Thin orchestrator calling `lib/manifest.py` (optional).
* `lib/__init__.py` - Package marker for shared modules.
* `lib/utils/__init__.py` - Package marker for utilities.
* `docs/config.schema.md` - Stub headings.
* `docs/sheet-selection.md` - Stub headings.
* `docs/sheet-inventory.md` - Stub headings.
* `docs/dwg-export-setup.md` - Stub headings.
* `docs/pdf-printing-prereqs.md` - Stub headings.
* `docs/batch-runner.md` - Stub headings.
* `docs/reports.md` - Stub headings.
* `docs/bim360-manifest.md` - Stub headings.
* `README.md` - Add “Batch Sheets Exporter” section listing tools and linking docs.

Implementation Specifications
- Keep each pushbutton `script.py` under ~40 LoC: resolve active doc, load config path (if needed), call a single library entry, handle exceptions by printing a friendly message.
- IronPython 2.7 constraints: no f-strings, no type annotations, avoid `pathlib`, avoid Python 3-only stdlib.
- Revit API transactions: do not start any transactions in scaffolding scripts.

Success Criteria
* Revit loads the extension and the new tab/panel renders without errors.
* Clicking each new button produces a controlled “not implemented yet” message (no unhandled exception).
* All required `docs/*.md` stub files exist.
* All tests pass: N/A (manual Revit validation)

Documentation Updates Required
* Update `README.md` with the new tool entrypoints and pointers to `docs/`.
* Update `AGENTS.md` with: “Batch Sheets tool lives under BaseDarom.exporter.tab” and “thin script.py, logic in lib/”.
* Create/update docstrings for all new functions (entrypoints, even if placeholders).
* Update API documentation if applicable: N/A

Validation Commands
bash
ls BaseDarom.exporter.tab
ls "BaseDarom.exporter.tab/Export Plans.panel"
ls lib
ls docs
Manual (Revit 2019): restart Revit; confirm tab/panel shows; click each button and confirm clean output.

Constraints
* Maximum implementation time: 3 hours
* Target code volume: ~100 lines
* Must maintain backward compatibility
* Follow existing code patterns in `Coordination.tab/3D Views.panel/Scope Boxs to 3D Views.pushbutton/script.py`

---

Task [3]: Add Pure-Python Test Harness for `lib/` Utilities (Testing)

Parallel Development Note: Enables parallel development by validating pure modules without requiring Revit; reduces rework on config/path/CSV logic.

Context
The plan calls for unit-like tests for non-Revit logic (config validation, path sanitization, parsing, deterministic ordering). The repo currently has no test harness.

Objective
Add a minimal `unittest` harness for pure-Python modules.

Technical Requirements
1. Primary Component - Add `tests/` with `unittest` discovery and at least one smoke test.
2. Integration Points - Document how tests import extension `lib/` (PYTHONPATH strategy).
3. Data/API Specifications - Standardize test naming (`test_*.py`) and scope (no Revit API imports).
4. Error Handling - Tests fail with actionable messages when import paths are misconfigured.

Files to Create/Modify
* `tests/README.md` - How to run tests (Windows) and scope boundaries.
* `tests/test_smoke_imports.py` - Smoke imports for `lib/config.py`, `lib/utils/*` (as they are added).
* `README.md` - Add “Testing” section with the canonical command.

Implementation Specifications
- Tests run under CPython 3 locally, but the code under test must remain IronPython 2.7 compatible (no Python 3-only constructs).
- Keep tests focused on pure helpers; Revit API validations remain manual.

Success Criteria
* `python -m unittest discover -s tests -p "test_*.py"` runs and passes.
* `tests/README.md` clearly explains how to run tests and what is excluded.
* All tests pass: python -m unittest discover -s tests -p "test_*.py"

Documentation Updates Required
* Update `README.md` section “Testing” with the command and scope.
* Update `AGENTS.md` with the project testing policy (pure-python only; Revit validations manual).
* Create/update docstrings for all new functions: N/A (test harness only)
* Update API documentation if applicable: N/A

Validation Commands
bash
python -m unittest discover -s tests -p "test_*.py"

Constraints
* Maximum implementation time: 3 hours
* Target code volume: ~100 lines
* Must maintain backward compatibility
* Follow existing code patterns in `CodexWorkspace/docs/agent/checklists/testing_strategy.md`

---

Task [4]: Specify `config.json` Schema + Example Assets (Docs)

Parallel Development Note: Can run concurrently with scaffolding and test harness; downstream coding tasks depend on stable config keys.

Context
The plan requires a `config.json` controlling sheet filters, DWG setup name, PDF printer/paper, output root, resume/overwrite, and report/status output paths.

Objective
Document the authoritative config schema and provide example config assets under `CodexWorkspace/workspace/`.

Technical Requirements
1. Primary Component - Document required/optional keys, defaults, allowed values, and schema version strategy.
2. Integration Points - Define config precedence (CLI `--config` override vs default search path).
3. Data/API Specifications - Exact fields for: selection mode, folder filter, codes mode, selection-file mode, DWG, PDF, resume policy, reporting/status.
4. Error Handling - Define invalid config handling (fail fast in Validate Setup; avoid silent coercion).

Files to Create/Modify
* `docs/config.schema.md` - Full schema doc with examples and upgrade notes.
* `CodexWorkspace/workspace/batch_sheets/config.json` - Example config (placeholders for printer/setup names).
* `CodexWorkspace/workspace/batch_sheets/models_2019.txt` - Example models list format (comments/blank lines).
* `README.md` - Add “Configuration” section linking to schema + example assets.

Implementation Specifications
- Required keys must cover: output root, selection mode, BIM folder filter param + value, DWG setup name, PDF printer + paper size name (900×1800 mm), resume/overwrite, report filenames, live status filenames, and checkpoint filenames.
- Example assets must not embed machine-specific absolute paths except clearly marked placeholders.

Success Criteria
* `docs/config.schema.md` documents all config keys required by the plan.
* The example `config.json` is valid JSON and matches the documented schema.
* The example `models_2019.txt` clearly documents the required input format.
* All tests pass: N/A (docs/assets only)

Documentation Updates Required
* Update `README.md` with config locations, precedence rules, and links to schema/examples.
* Update `AGENTS.md` with: “Config changes require updating docs + tests”.
* Create/update docstrings for all new functions: N/A (docs/assets only)
* Update API documentation if applicable: N/A

Validation Commands
bash
python -c "import json; json.load(open('CodexWorkspace/workspace/batch_sheets/config.json','r'))"

Constraints
* Maximum implementation time: 3 hours
* Target code volume: ~100 lines
* Must maintain backward compatibility
* Follow existing code patterns in `CodexWorkspace/workspace/plans/batch_sheets_plan.md`

---

Task [5]: Implement Config Loader + Validation (`lib/config.py`) (Core Library)

Parallel Development Note: Can run concurrently with utility modules; once stable, it unblocks Validate Setup, UI flows, and CLI runner.

Context
All tools must load config deterministically with clear validation errors and safe defaults (only where explicitly allowed by schema).

Objective
Implement a config loader/validator that returns a normalized config object used consistently across modules.

Technical Requirements
1. Primary Component - Implement config load, validate, default, and normalize functions in `lib/config.py`.
2. Integration Points - Used by pushbuttons, Validate Setup, batch runner, and CLI runner.
3. Data/API Specifications - Enforce enums, required keys, normalized paths, and schema version rules.
4. Error Handling - Raise a typed exception containing a user-facing error message; avoid raw tracebacks in UI.

Files to Create/Modify
* `lib/config.py` - Config loader/validator and exceptions.
* `tests/test_config_validation.py` - Unit tests for missing keys, invalid enums, defaults, and schema version.
* `docs/config.schema.md` - Update with normalization rules and error examples.

Implementation Specifications
- Normalization rules must be explicit (trim strings, normalize paths, validate non-empty names for printer/setup).
- Keep implementation IronPython 2.7 compatible (no f-strings, no type annotations, avoid Python 3-only libs).

Success Criteria
* Invalid config produces a single actionable error message listing failing keys and expected values.
* Valid config loads and returns a normalized dict used consistently by downstream modules.
* All tests pass: python -m unittest discover -s tests -p "test_*.py"

Documentation Updates Required
* Update `docs/config.schema.md` with validation and normalization rules.
* Update `README.md` with config precedence and “common config errors” section.
* Update `AGENTS.md` with config versioning and validation conventions.
* Create/update docstrings for config API functions and exception types.

Validation Commands
bash
python -m unittest discover -s tests -p "test_*.py"

Constraints
* Maximum implementation time: 3 hours
* Target code volume: ~100 lines
* Must maintain backward compatibility
* Follow existing code patterns in `CodexWorkspace/docs/agent/pyrevit/cpython-ironpython-development.md`

---

Task [6]: Implement Deterministic Path + Filename Utilities (`lib/utils/path.py`) (Core Utilities)

Parallel Development Note: Pure-python utility work; can run concurrently with Revit API tasks and is required for deterministic outputs and resume behavior.

Context
Batch exports require deterministic folder/file naming, safe Windows filenames, and atomic writes for status/checkpoints.

Objective
Implement shared filesystem helpers for naming, sanitization, directory creation, and atomic writes.

Technical Requirements
1. Primary Component - Provide filename sanitization and deterministic output path builders for sheets/models.
2. Integration Points - Used by exporters, reporting, checkpointing, and status writers.
3. Data/API Specifications - Define deterministic naming rules and collision strategy (stable suffixing).
4. Error Handling - Detect invalid/too-long names; shorten deterministically and log; avoid silent overwrites.

Files to Create/Modify
* `lib/utils/path.py` - Sanitization, naming, ensure-dir, atomic-write helpers.
* `tests/test_path_utils.py` - Unit tests for invalid chars, reserved names, trimming, collisions.
* `docs/batch-runner.md` - Document deterministic naming and collision policy.

Implementation Specifications
- Handle Windows invalid characters: `\\ / : * ? \" < > |` and reserved device names.
- Provide helper: “exists and size > min_bytes” for resume mode decisions.
- Atomic write strategy for JSON/status/checkpoint files (write temp, then replace).

Success Criteria
* Same inputs always produce the same output paths and filenames.
* Collision handling produces stable names without overwriting prior outputs unexpectedly.
* All tests pass: python -m unittest discover -s tests -p "test_*.py"

Documentation Updates Required
* Update `docs/batch-runner.md` with naming policy and collision behavior.
* Update `README.md` with output folder layout and naming rules.
* Update `AGENTS.md` with “All filesystem writes go through `lib/utils/path.py`”.
* Create/update docstrings for all helper functions.

Validation Commands
bash
python -m unittest discover -s tests -p "test_*.py"

Constraints
* Maximum implementation time: 3 hours
* Target code volume: ~100 lines
* Must maintain backward compatibility
* Follow existing code patterns in `Coordination.tab/3D Views.panel/Scope Boxs to 3D Views.pushbutton/script.py`

---

Task [7]: Implement Unicode-Safe CSV Helpers (`lib/utils/csv_unicode.py`) (Core Utilities)

Parallel Development Note: Pure-python; can run concurrently with inventory/reporting work and is required for Excel-friendly outputs.

Context
The plan requires CSV outputs (inventory, reports, failed_sheets) that are usable in Excel and preserve Hebrew/Unicode.

Objective
Implement robust Unicode-safe CSV read/write helpers under IronPython 2.7 constraints.

Technical Requirements
1. Primary Component - Provide CSV writer/reader utilities with UTF-8 (optionally BOM for Excel).
2. Integration Points - Used by inventory generation, report CSVs, selection-file parsing, and live failure logging.
3. Data/API Specifications - Standardize delimiter/quoting/newlines and header validation helpers.
4. Error Handling - Header mismatch fails fast (selection files); encoding errors logged with row/column context.

Files to Create/Modify
* `lib/utils/csv_unicode.py` - Unicode-safe CSV read/write and header validation.
* `tests/test_csv_unicode.py` - Unit tests for Hebrew strings, quoting, and header validation.
* `docs/sheet-inventory.md` - Document encoding expectations and required headers.

Implementation Specifications
- Provide strict header validation for coordinator selection files (exact required columns).
- Provide safe append/update strategy for `failed_sheets.csv` (deterministic behavior documented).

Success Criteria
* CSV outputs open in Excel with readable Unicode (Hebrew) content.
* Selection file header validation produces actionable errors.
* All tests pass: python -m unittest discover -s tests -p "test_*.py"

Documentation Updates Required
* Update `docs/sheet-inventory.md` with encoding rules and selection-file header requirements.
* Update `README.md` with coordinator workflow notes (encoding + headers).
* Update `AGENTS.md` with CSV conventions (UTF-8, optional BOM, strict headers).
* Create/update docstrings for CSV helpers.

Validation Commands
bash
python -m unittest discover -s tests -p "test_*.py"

Constraints
* Maximum implementation time: 3 hours
* Target code volume: ~100 lines
* Must maintain backward compatibility
* Follow existing code patterns in `CodexWorkspace/docs/agent/checklists/testing_strategy.md`

---

Task [8]: Implement Deterministic Sheet Ordering Helpers (`lib/utils/sorting.py`) (Core Utilities)

Parallel Development Note: Pure-python; can be implemented and tested without Revit and then reused by sheet query/reporting.

Context
Deterministic outputs require stable ordering of sheets across runs; sheet numbers often mix digits and letters.

Objective
Implement a stable, documented ordering rule for sheets and helper functions to produce sort keys.

Technical Requirements
1. Primary Component - Implement “natural-ish” ordering for sheet numbers and stable tie-break rules.
2. Integration Points - Used by `lib/sheet_query.py` and reporting to ensure deterministic output order.
3. Data/API Specifications - Order by: sheet number (natural), then sheet name, then element id (final tie-break).
4. Error Handling - Missing/empty sheet numbers handled deterministically (placed last with stable tie-breaks).

Files to Create/Modify
* `lib/utils/sorting.py` - Natural-ish sort key helpers.
* `tests/test_sorting.py` - Unit tests for representative sheet numbers and edge cases.
* `docs/sheet-selection.md` - Document ordering rules.

Implementation Specifications
- Keep helpers pure (no Revit API imports).
- Publish a single canonical helper used across modules to avoid divergence.

Success Criteria
* Sorting is deterministic for representative cases (e.g., A-2, A-10, A-2A, 01, 1, empty).
* Ordering rules are documented and match tests.
* All tests pass: python -m unittest discover -s tests -p "test_*.py"

Documentation Updates Required
* Update `docs/sheet-selection.md` with exact ordering rules.
* Update `README.md` with “Deterministic ordering” note.
* Update `AGENTS.md` with “All ordering uses `lib/utils/sorting.py`”.
* Create/update docstrings for sorting helpers.

Validation Commands
bash
python -m unittest discover -s tests -p "test_*.py"

Constraints
* Maximum implementation time: 3 hours
* Target code volume: ~100 lines
* Must maintain backward compatibility
* Follow existing code patterns in `CodexWorkspace/workspace/plans/batch_sheets_plan.md`

---

Task [9]: Implement Structured Logging + Error Capture (`lib/reporting.py`) (Diagnostics)

Parallel Development Note: Can run concurrently with exporters and runner work; it defines how failures are recorded without blocking long runs.

Context
The plan requires `errors.log` with stack traces and dialog-free batch execution; failures must be recorded and runs must continue.

Objective
Implement structured logging helpers that write `errors.log` and provide an exception capture utility used across modules.

Technical Requirements
1. Primary Component - Provide logger helpers and exception formatting/capture (stack traces to file).
2. Integration Points - Used by Validate Setup, exporters, batch runner, and CLI runner to record failures.
3. Data/API Specifications - Define log record fields: timestamp, level, model identifier, sheet identifier, message, exception/stack.
4. Error Handling - Logging must never crash the run; if log write fails, fall back to pyRevit output.

Files to Create/Modify
* `lib/reporting.py` - Logging helpers, exception capture, and base report schema definitions.
* `tests/test_error_formatting.py` - Unit tests for formatting and Unicode safety (pure python).
* `docs/reports.md` - Document `errors.log` format and location.

Implementation Specifications
- Unicode-safe writing for Hebrew; avoid smart quotes and platform-dependent encodings.
- Provide a single API used everywhere (avoid ad-hoc `print` for errors in batch mode).

Success Criteria
* Exceptions can be captured and written to `errors.log` with stack trace and context.
* Logging failures do not crash exports.
* All tests pass: python -m unittest discover -s tests -p "test_*.py"

Documentation Updates Required
* Update `docs/reports.md` with `errors.log` schema and troubleshooting guidance.
* Update `README.md` with log locations and “how to report errors”.
* Update `AGENTS.md` with “Do not show modal dialogs during batch; always log via reporting helpers”.
* Create/update docstrings for reporting/logging functions.

Validation Commands
bash
python -m unittest discover -s tests -p "test_*.py"

Constraints
* Maximum implementation time: 3 hours
* Target code volume: ~100 lines
* Must maintain backward compatibility
* Follow existing code patterns in `CodexWorkspace/workspace/plans/batch_sheets_plan.md`

---

Task [10]: Implement Live Status Writers (`progress.json`, `failed_sheets.csv`) (Resiliency)

Parallel Development Note: Can run concurrently with runner/exporter work; it is a self-contained IO component driven by per-sheet results.

Context
The plan requires live monitoring files during long runs: `progress.json` and `failed_sheets.csv`, updated after each sheet attempt.

Objective
Implement incremental writers for `progress.json` and `failed_sheets.csv` with atomic writes and deterministic schemas.

Technical Requirements
1. Primary Component - Write/update `progress.json` after each sheet with current model/sheet and percent complete.
2. Integration Points - Called by batch runner per sheet; CLI runner can expose file locations for monitoring.
3. Data/API Specifications - Define exact JSON keys and CSV columns (model path/id, structure, discipline, sheet number/name, export kind, error).
4. Error Handling - File lock/write failures must be non-fatal; record to `errors.log` and continue.

Files to Create/Modify
* `lib/reporting.py` - Add live status writer APIs (or create `lib/status.py` and re-export from reporting).
* `docs/reports.md` - Document schemas and update cadence.
* `tests/test_live_status_schema.py` - Unit tests for JSON schema and CSV header stability (pure python).

Implementation Specifications
- `progress.json` must always remain valid JSON (atomic replace).
- Define deterministic behavior for `failed_sheets.csv`: append-only vs dedup-by-key (choose one and document).

Success Criteria
* During a run, `progress.json` updates after each sheet and remains valid JSON.
* On any failure, `failed_sheets.csv` is updated immediately with a complete row.
* All tests pass: python -m unittest discover -s tests -p "test_*.py"

Documentation Updates Required
* Update `docs/reports.md` with `progress.json` and `failed_sheets.csv` schemas and examples.
* Update `README.md` with “Monitoring outputs” section pointing to these files.
* Update `AGENTS.md` with “Status writes are atomic and non-fatal on failure.”
* Create/update docstrings for status writer APIs.

Validation Commands
bash
python -m unittest discover -s tests -p "test_*.py"

Constraints
* Maximum implementation time: 3 hours
* Target code volume: ~100 lines
* Must maintain backward compatibility
* Follow existing code patterns in `CodexWorkspace/workspace/plans/batch_sheets_plan.md`

---

Task [11]: Implement Models List Parser (`models_2019.txt`) (Core Utilities)

Parallel Development Note: Pure-python parsing; can run concurrently with Revit API work and unblocks CLI batch runner input validation.

Context
Batch mode uses a text file with one local `.rvt` path per line. The plan requires validating that models are available locally (hydrated) and that the input list is well-formed before starting long runs.

Objective
Implement a deterministic parser for the models list file format used by `pyrevit run`.

Technical Requirements
1. Primary Component - Parse a text file into an ordered list of model paths with stable behavior.
2. Integration Points - Used by `lib/runner_entry.py` and inventory/export batch modes.
3. Data/API Specifications - Define accepted syntax: blank lines allowed; full-line comments supported; optional surrounding quotes stripped.
4. Error Handling - Missing file, empty list, and invalid lines must produce actionable errors (not silent drops).

Files to Create/Modify
* `lib/utils/model_list.py` - Parse models list file and return (paths, parse_report).
* `tests/test_model_list_parser.py` - Tests for comments, blank lines, ordering, quoted paths, and BOM handling.
* `docs/batch-runner.md` - Document models list format and examples.

Implementation Specifications
- Parsing rules (deterministic):
  - Strip UTF-8 BOM if present.
  - `line.strip()`; skip empty lines.
  - Treat lines starting with `#` as comments.
  - If a line begins and ends with `"`, strip the quotes.
  - Preserve remaining line content as the path (no inline comment parsing to avoid breaking legitimate paths).
- Return a `parse_report` dict with at least: `total_lines`, `paths_count`, `skipped_blank`, `skipped_comments`, `invalid_lines` (line numbers + content).

Success Criteria
* Parser returns an ordered list matching the file order (deterministic).
* Invalid inputs produce clear actionable errors or structured invalid line reports.
* All tests pass: python -m unittest discover -s tests -p "test_*.py"

Documentation Updates Required
* Update `docs/batch-runner.md` with a copy/paste example `models_2019.txt`.
* Update `README.md` section “Batch Mode” with reference to the models list file.
* Update `AGENTS.md` with “Validate batch inputs early; models list parsing is deterministic.”
* Create/update docstrings for parser functions and parse_report schema.

Validation Commands
bash
python -m unittest discover -s tests -p "test_*.py"

Constraints
* Maximum implementation time: 3 hours
* Target code volume: ~100 lines
* Must maintain backward compatibility
* Follow existing code patterns in `CodexWorkspace/workspace/plans/batch_sheets_plan.md`

---

Task [12]: Implement Local Revit Document Open/Close Helpers (Revit API) (Revit Integration)

Parallel Development Note: Can run concurrently with exporters; it is required for CLI batch mode and batch inventory mode.

Context
The plan requires batch orchestration across many models. For Autodesk Docs/Desktop Connector, models must be hydrated locally; opening can fail and must not block the entire run.

Objective
Implement safe helpers to open and close local `.rvt` files in a dialog-free batch context.

Technical Requirements
1. Primary Component - Open local documents by path and close without saving; provide a structured error contract.
2. Integration Points - Used by `lib/runner_entry.py` for inventory/export runs; uses `lib/reporting.py` for errors.
3. Data/API Specifications - Define `open_result` contract: success returns `(doc, meta)`; failure returns `(None, error_info)`.
4. Error Handling - Per-model open failures must be classified (missing file, access denied, version/upgrade prompt risk, file locked) and must not crash batch mode.

Files to Create/Modify
* `lib/utils/revit.py` - `open_document(app, path, cfg)` and `close_document(doc)` helpers + error classification.
* `docs/batch-runner.md` - Document hydration requirement and how batch handles open failures.

Implementation Specifications
- Default behavior: open in a way that avoids prompts as much as possible (within Revit API limits) and rely on `pyrevit run` dialog suppression for production runs.
- Close behavior: always attempt `doc.Close(False)` in `finally` blocks; never leave docs open on failure.
- Error classification: wrap exceptions into a dict with keys: `kind`, `message`, `path`, `exception_type`, `trace_id` (optional).

Success Criteria
* Opening a valid local model path succeeds and returns a `Document`.
* Opening a missing/inaccessible model returns a structured failure without crashing.
* Documents are closed reliably after processing.
* All tests pass: N/A (manual Revit validation)

Documentation Updates Required
* Update `docs/batch-runner.md` with a “Hydration + Open Failures” troubleshooting section.
* Update `README.md` with prerequisites: Revit 2019, pyRevit 5.3.1, hydrated local models.
* Update `AGENTS.md` with “Batch mode must always close docs; no modal dialogs in production runs.”
* Create/update docstrings for open/close helpers and error contract.

Validation Commands
bash
Manual (Revit 2019):
1) Prepare one valid local RVT path and one invalid path.
2) Run a small debug invocation of open/close helpers via a temporary dev entrypoint.
3) Confirm: valid opens/closes; invalid returns structured error; no dialogs block execution in production settings.

Constraints
* Maximum implementation time: 3 hours
* Target code volume: ~100 lines
* Must maintain backward compatibility
* Follow existing code patterns in `CodexWorkspace/workspace/plans/batch_sheets_plan.md`

---

Task [13]: Add .NET Collection Boundary Helpers (ElementId List, ViewSet) (Revit API) (Revit Integration)

Parallel Development Note: Small isolated helper module; can run concurrently with DWG/PDF implementation and prevents repetitive boundary code.

Context
The plan requires .NET collections at Revit API boundaries (e.g., `List[ElementId]`, `ViewSet`).

Objective
Implement explicit helpers to convert Python iterables into required .NET collection types safely.

Technical Requirements
1. Primary Component - Provide helpers for `List[ElementId]` and `ViewSet` creation.
2. Integration Points - Used by `lib/export_dwg.py` and `lib/export_pdf.py`.
3. Data/API Specifications - Validate input item types and raise clear errors on invalid inputs.
4. Error Handling - Never pass invalid collections into the Revit API; surface actionable messages.

Files to Create/Modify
* `lib/utils/netcollections.py` - Conversion helpers for `List[ElementId]` and `ViewSet`.
* `docs/batch-runner.md` - Document “convert at API boundaries” convention.

Implementation Specifications
- Keep functions small and explicit (no implicit magic conversions).
- Expected helpers:
  - `to_elementid_list(iterable_of_elementids)`
  - `to_viewset(iterable_of_views)`

Success Criteria
* DWG/PDF modules can obtain correct .NET collection types via these helpers (verified in Revit).
* Invalid input produces clear error messages before calling Revit API.
* All tests pass: N/A (manual Revit validation)

Documentation Updates Required
* Update `docs/batch-runner.md` with the .NET boundary conversion rule.
* Update `AGENTS.md` with “Convert Python iterables to .NET collections at API boundaries.”
* Create/update docstrings for conversion helpers.
* Update API documentation if applicable: N/A

Validation Commands
bash
Manual (Revit 2019):
1) Convert 2 sheet ElementIds into List[ElementId] and confirm type.
2) Convert 2 ViewSheets into a ViewSet and confirm printing/export APIs accept it.

Constraints
* Maximum implementation time: 3 hours
* Target code volume: ~100 lines
* Must maintain backward compatibility
* Follow existing code patterns in `Coordination.tab/3D Views.panel/Scope Boxs to 3D Views.pushbutton/script.py`

---

Task [14]: Implement Model Identity + Structure/Discipline Inference (Metadata) (Core Library)

Parallel Development Note: Pure-python inference logic and unit tests can be done in parallel with Revit exporter tasks; the contract unblocks output routing and reporting.

Context
Outputs must be organized by `/<Structure>/<Discipline>/...`. The plan calls for inferring structure and discipline from filename/path and/or Project Information.

Objective
Implement deterministic model identity extraction and inference for structure/discipline used for output routing and reporting.

Technical Requirements
1. Primary Component - Implement `ModelIdentity` extraction from local path and (optionally) Revit doc metadata.
2. Integration Points - Used by `lib/batch_runner.py`, `lib/runner_entry.py`, and report writers.
3. Data/API Specifications - Define identity fields: `model_path`, `model_title`, `structure_id`, `discipline_id`, and `inference_reason`.
4. Error Handling - If inference fails, return explicit `Unknown` values (never empty) and continue.

Files to Create/Modify
* `lib/bim360_adapter.py` - Model identity helpers and inference functions (local-path-first).
* `tests/test_model_inference.py` - Unit tests for discipline/structure inference from representative file paths.
* `docs/bim360-manifest.md` - Document inference rules and how to override via config/manifest.

Implementation Specifications
- Inference must be config-driven:
  - Map discipline keywords/patterns to discipline ids (ARCH/STR/HVAC/PLUMB/EL).
  - Map structure patterns (e.g., base darom codes) to structure ids.
- Provide `inference_reason` explaining what matched (for report transparency).

Success Criteria
* Given representative model paths, inference returns expected `structure_id` and `discipline_id`.
* Unknown inputs return deterministic `Unknown` values and a reason.
* All tests pass: python -m unittest discover -s tests -p "test_*.py"

Documentation Updates Required
* Update `docs/bim360-manifest.md` with inference rules and override strategy.
* Update `README.md` with output routing description (how structure/discipline are determined).
* Update `AGENTS.md` with “Inference is config-driven; do not hardcode project-specific codes.”
* Create/update docstrings for identity/inference functions and return schema.

Validation Commands
bash
python -m unittest discover -s tests -p "test_*.py"

Constraints
* Maximum implementation time: 3 hours
* Target code volume: ~100 lines
* Must maintain backward compatibility
* Follow existing code patterns in `CodexWorkspace/workspace/plans/batch_sheets_plan.md`

---

Task [15]: Implement Baseline Sheet Collection (Exclude Placeholders) (Revit API) (Revit Integration)

Parallel Development Note: Can run concurrently with inventory/export tasks; defines a single reusable selection foundation for all modes.

Context
The plan requires collecting `ViewSheet` targets, excluding placeholders, and returning a deterministic ordered list.

Objective
Implement baseline sheet collection that returns non-placeholder sheets with deterministic ordering and selection stats.

Technical Requirements
1. Primary Component - Collect `ViewSheet` via `FilteredElementCollector`, exclude placeholders (`IsPlaceholder`).
2. Integration Points - Used by inventory, export, and reporting workflows.
3. Data/API Specifications - Return list of `ViewSheet` plus selection stats; optionally produce `.NET List[ElementId]` using netcollections helpers.
4. Error Handling - No sheets found must be handled gracefully (empty list + reason).

Files to Create/Modify
* `lib/sheet_query.py` - Baseline collector and return contract.
* `docs/sheet-selection.md` - Document baseline selection and ordering.

Implementation Specifications
- Deterministic ordering uses `lib/utils/sorting.py` rules.
- Return a summary dict: `total`, `placeholders_excluded`, `selected`, `notes`.

Success Criteria
* Returns all non-placeholder sheets and excludes placeholder sheets.
* Ordering is stable across repeated runs (same model, same config).
* All tests pass: N/A (manual Revit validation)

Documentation Updates Required
* Update `docs/sheet-selection.md` with baseline rules and ordering.
* Update `README.md` with baseline selection behavior (placeholders excluded).
* Update `AGENTS.md` with “All selection logic is centralized in `lib/sheet_query.py`.”
* Create/update docstrings for sheet query functions and return schema.

Validation Commands
bash
Manual (Revit 2019):
1) Open a model with placeholder and real sheets.
2) Run a debug call to sheet_query baseline.
3) Confirm placeholders excluded and ordering stable.

Constraints
* Maximum implementation time: 3 hours
* Target code volume: ~100 lines
* Must maintain backward compatibility
* Follow existing code patterns in `CodexWorkspace/workspace/plans/batch_sheets_plan.md`

---

Task [16]: Implement BIM Folder Filter Sheet Selection (Revit API) (Revit Integration)

Parallel Development Note: Can run concurrently with inventory/exporter work; it only depends on baseline sheet collection and config keys.

Context
Default behavior must export all sheets matching a “BIM folder” filter (configured parameter name and required value).

Objective
Filter sheets by a configured parameter/value that represents the BIM folder, while preserving deterministic ordering.

Technical Requirements
1. Primary Component - Implement folder filter logic on `ViewSheet` using a configured parameter name and required string value.
2. Integration Points - Used by default export mode and inventory generation mode.
3. Data/API Specifications - Define normalization rules (trim strings; exact match after normalization); collect selection stats.
4. Error Handling - Missing parameter/value must not crash selection; report counts for missing/empty parameters.

Files to Create/Modify
* `lib/sheet_query.py` - Add folder filter selection mode and selection stats.
* `docs/sheet-selection.md` - Document folder filter behavior and troubleshooting.
* `docs/config.schema.md` - Ensure folder filter keys are documented.

Implementation Specifications
- Look up parameter by name via `LookupParameter(param_name)`; read string via `AsString()` when available.
- Normalize both values with trim; compare exact strings after normalization.
- Return selection stats: `matched`, `missing_param`, `empty_value`, `excluded`.

Success Criteria
* With correct config, only sheets in the configured BIM folder are selected.
* Missing/empty parameters are counted and reported; selection continues.
* All tests pass: N/A (manual Revit validation)

Documentation Updates Required
* Update `docs/sheet-selection.md` with the exact parameter/value matching rules and troubleshooting steps.
* Update `README.md` with “Default selection uses BIM folder filter” summary.
* Update `AGENTS.md` with “Folder filter strings are config-driven; never hardcode project strings.”
* Create/update docstrings for folder filter functions and stats.

Validation Commands
bash
Manual (Revit 2019):
1) Identify the BIM folder parameter name used on sheets in a test model.
2) Set config keys for parameter name and required value.
3) Run selection and confirm matched counts and sheet list correctness.

Constraints
* Maximum implementation time: 3 hours
* Target code volume: ~100 lines
* Must maintain backward compatibility
* Follow existing code patterns in `CodexWorkspace/workspace/plans/batch_sheets_plan.md`

---

Task [17]: Implement Sheet Inventory Schema + Writer (CSV) (Core Library)

Parallel Development Note: Mostly pure logic; can be developed in parallel with DWG/PDF exporters and enables early coordinator workflows.

Context
The plan requires a “Build Sheet List” output: a flat CSV inventory (one row per sheet per model) that coordinators can mark up to control exports.

Objective
Define the inventory CSV schema and implement a writer that produces inventory rows for a given model.

Technical Requirements
1. Primary Component - Implement stable inventory schema constants and CSV writer functions.
2. Integration Points - Used by interactive Build Sheet List and by CLI inventory mode (later).
3. Data/API Specifications - Define required columns and optional markup columns (`export_pdf`, `export_dwg`, `priority`, `notes`).
4. Error Handling - Missing parameters must yield empty strings; Unicode must be preserved (Hebrew-safe).

Files to Create/Modify
* `lib/sheet_inventory.py` - Inventory schema constants and `write_inventory_csv(...)`.
* `tests/test_inventory_schema.py` - Unit tests verifying header stability and column order (pure python).
* `docs/sheet-inventory.md` - Document schema, meanings, and coordinator usage.

Implementation Specifications
- Inventory columns must include at minimum:
  - `model_path`, `structure_id`, `discipline_id`, `sheet_number`, `sheet_name`
  - folder filter value (for QA), optional code parameter value (if configured)
  - markup columns: `export_dwg`, `export_pdf`, `priority`, `notes`
- Column order must be stable; changing it requires doc + test updates.

Success Criteria
* Inventory CSV header is stable and matches documentation.
* CSV can represent Hebrew strings without corruption.
* All tests pass: python -m unittest discover -s tests -p "test_*.py"

Documentation Updates Required
* Update `docs/sheet-inventory.md` with schema table and example rows.
* Update `README.md` with the coordinator workflow overview (Build -> mark -> export).
* Update `AGENTS.md` with “Inventory schema is stable; changes require docs + tests.”
* Create/update docstrings for schema constants and inventory writer functions.

Validation Commands
bash
python -m unittest discover -s tests -p "test_*.py"

Constraints
* Maximum implementation time: 3 hours
* Target code volume: ~100 lines
* Must maintain backward compatibility
* Follow existing code patterns in `CodexWorkspace/workspace/plans/batch_sheets_plan.md`

---

Task [18]: Implement Selection File Parser (Coordinator Markup CSV) (Core Library)

Parallel Development Note: Pure-python; can run concurrently with Revit API work and unblocks selection-file export mode.

Context
Optional workflow: coordinators mark `export_dwg`/`export_pdf` in the inventory CSV, then export runs in selection-file mode exporting only marked sheets.

Objective
Implement a strict parser for coordinator selection files that returns deterministic export instructions.

Technical Requirements
1. Primary Component - Parse the inventory/selection CSV and return a mapping of selected sheet numbers to export flags.
2. Integration Points - Used by selection-file sheet query mode and by the batch runner to decide DWG vs PDF exports per sheet.
3. Data/API Specifications - Strict required headers; accepted truthy values; deterministic duplicate handling.
4. Error Handling - Header mismatch fails fast with actionable message; invalid rows are reported deterministically.

Files to Create/Modify
* `lib/sheet_inventory.py` - `parse_selection_file(path)` and required header constants.
* `tests/test_selection_file_parsing.py` - Unit tests for headers, truthy values, duplicates, Unicode.
* `docs/sheet-inventory.md` - Document required headers and coordinator markup rules.

Implementation Specifications
- Accepted truthy values (case-insensitive): `yes/no`, `true/false`, `1/0`, empty treated as false.
- Duplicate handling must be deterministic and documented (recommended: last-wins for coordinator edits).
- Return structure (example):
  - `{"A-101": {"export_dwg": true, "export_pdf": false, "priority": "1", "notes": "..."}}`

Success Criteria
* Parser returns deterministic selection mapping and a structured list of invalid rows (if any).
* Header mismatch produces a single actionable error message.
* All tests pass: python -m unittest discover -s tests -p "test_*.py"

Documentation Updates Required
* Update `docs/sheet-inventory.md` with required headers and accepted truthy values.
* Update `README.md` with selection-file mode description and limitations.
* Update `AGENTS.md` with “Selection parsing is strict; changes require docs + tests.”
* Create/update docstrings for parser API and return schema.

Validation Commands
bash
python -m unittest discover -s tests -p "test_*.py"

Constraints
* Maximum implementation time: 3 hours
* Target code volume: ~100 lines
* Must maintain backward compatibility
* Follow existing code patterns in `CodexWorkspace/workspace/plans/batch_sheets_plan.md`

---

Task [19]: Integrate Selection-File Mode into Sheet Query (Revit API) (Revit Integration)

Parallel Development Note: Can run concurrently with exporter implementation; it depends on sheet query baseline + selection file parser only.

Context
Selection-file mode must export only marked sheets and report “missing in model” cases (CSV references not found in the current model).

Objective
Add selection-file mode filtering to sheet query and surface missing referenced sheets for reporting.

Technical Requirements
1. Primary Component - Filter candidate sheets by `sheet_number` based on selection mapping.
2. Integration Points - Used by batch runner (per model) and interactive export (active doc).
3. Data/API Specifications - Define matching rules (normalized exact match on sheet number) and missing reference outputs.
4. Error Handling - Missing sheet references do not crash; they are returned for reporting.

Files to Create/Modify
* `lib/sheet_query.py` - Add selection-file mode and return missing selection references.
* `docs/sheet-selection.md` - Document selection-file precedence and missing-sheet behavior.
* `docs/reports.md` - Document how missing selection references appear in reports.

Implementation Specifications
- Normalize sheet numbers (trim) consistently between selection file parser and Revit sheet numbers.
- Return structure should include:
  - `selected_sheets` (ordered list)
  - `missing_selection_rows` (sheet numbers referenced but not found)
  - selection stats (counts)

Success Criteria
* Only marked sheets are selected in selection-file mode.
* Missing selection sheet numbers are returned for reporting.
* All tests pass: N/A (manual Revit validation)

Documentation Updates Required
* Update `docs/sheet-selection.md` with precedence rules and missing reference reporting.
* Update `README.md` with selection-file mode usage steps.
* Update `AGENTS.md` with “Selection returns both selected and missing sets for reporting.”
* Create/update docstrings for selection-file mode query functions and return schema.

Validation Commands
bash
Manual (Revit 2019):
1) Create a selection CSV marking a subset of real sheets and at least one fake sheet number.
2) Run selection-file mode selection on the model.
3) Confirm only marked sheets selected and fake appears in missing list.

Constraints
* Maximum implementation time: 3 hours
* Target code volume: ~100 lines
* Must maintain backward compatibility
* Follow existing code patterns in `CodexWorkspace/workspace/plans/batch_sheets_plan.md`

---

Task [20]: Implement Optional Codes Mode (Code List + Sheet Parameter) (Revit Integration)

Parallel Development Note: Code list parser/tests can be built in parallel with Revit tasks; integration depends only on `sheet_query` and config.

Context
The plan describes an optional workflow to filter sheets by a configured “sheet code” parameter and a provided code list, but default remains BIM folder filter.

Objective
Add an explicit `selection_mode = "codes"` that filters sheets by code list and a configured sheet parameter.

Technical Requirements
1. Primary Component - Parse a code list file into a normalized set and filter sheets by configured code parameter.
2. Integration Points - Integrated into `lib/sheet_query.py` as an opt-in mode; uses config keys for parameter name and code list path.
3. Data/API Specifications - Define supported code list formats (one code per line, optional CSV single column) and normalization rules.
4. Error Handling - Missing code list or empty set fails fast with actionable error; missing sheet parameter counts as non-match and is reported.

Files to Create/Modify
* `lib/utils/code_list.py` - Code list parser (pure python).
* `lib/sheet_query.py` - Codes mode filtering and stats.
* `tests/test_code_list_parser.py` - Unit tests for parsing/normalization/dedup.
* `docs/sheet-selection.md` - Document codes mode behavior and configuration.

Implementation Specifications
- Normalization rules must be explicit and documented (trim; optional uppercase).
- Codes mode is opt-in only and must not change the default folder-filter behavior.

Success Criteria
* Codes mode selects only sheets whose code parameter matches the configured code set.
* Invalid/empty code list fails before export begins with actionable error.
* All tests pass: python -m unittest discover -s tests -p "test_*.py"

Documentation Updates Required
* Update `docs/sheet-selection.md` with codes mode format, config keys, and examples.
* Update `README.md` with “Optional codes mode” description and warnings.
* Update `AGENTS.md` with “Codes mode is opt-in; defaults must remain folder filter.”
* Create/update docstrings for code list parser and codes selection logic.

Validation Commands
bash
python -m unittest discover -s tests -p "test_*.py"
Manual (Revit 2019): configure code parameter name and provide code list; confirm selection matches expected sheets.

Constraints
* Maximum implementation time: 3 hours
* Target code volume: ~100 lines
* Must maintain backward compatibility
* Follow existing code patterns in `CodexWorkspace/workspace/plans/batch_sheets_plan.md`

---

Task [21]: Wire Interactive Build Sheet List Pushbutton (Active Document) (UI Integration)

Parallel Development Note: Can run concurrently with DWG/PDF exporter development; it only depends on sheet selection + inventory writer.

Context
The plan recommends starting with a sheet inventory export so coordinators can mark which sheets should be exported.

Objective
Implement the Build Sheet List pushbutton workflow to generate the inventory CSV for the active document.

Technical Requirements
1. Primary Component - Prompt for output CSV path (or output folder) and write inventory CSV for active doc.
2. Integration Points - Use `lib/config.py` (optional), `lib/sheet_query.py` (folder filter), and `lib/sheet_inventory.py`.
3. Data/API Specifications - Inventory output must match the schema in `docs/sheet-inventory.md` and include markup columns.
4. Error Handling - Cancel paths must exit cleanly; outside-Revit guard must provide a friendly message.

Files to Create/Modify
* `BaseDarom.exporter.tab/Export Plans.panel/Build Sheet List.pushbutton/script.py` - Orchestrator that calls the library workflow.
* `lib/ui.py` - Optional UI helpers (save-file prompt, consistent output messages).
* `docs/sheet-inventory.md` - Add interactive usage steps (button name, prompts, outputs).

Implementation Specifications
- Prefer a deterministic default filename (e.g., includes model title and date) but allow user override.
- Ensure the inventory includes `model_path`, inferred `structure_id`/`discipline_id`, and selection-relevant parameter values (folder/code).

Success Criteria
* In Revit 2019, running Build Sheet List produces a CSV at the chosen path.
* Cancelling the dialog exits cleanly without exceptions.
* CSV opens in Excel and matches documented headers.
* All tests pass: N/A (manual Revit validation)

Documentation Updates Required
* Update `docs/sheet-inventory.md` with interactive workflow steps and example output locations.
* Update `README.md` with the coordinator workflow (Build Sheet List -> mark -> export).
* Update `AGENTS.md` with UI orchestration rules (thin script.py; logic in lib/).
* Create/update docstrings for any new UI helper functions.

Validation Commands
bash
Manual (Revit 2019):
1) Open a discipline model with sheets.
2) Click Build Sheet List.
3) Choose output location; confirm CSV created and readable in Excel.

Constraints
* Maximum implementation time: 3 hours
* Target code volume: ~100 lines
* Must maintain backward compatibility
* Follow existing code patterns in `Coordination.tab/3D Views.panel/Scope Boxs to 3D Views.pushbutton/script.py`

---

Task [22]: Implement CLI Inventory Mode (`pyrevit run`) (CLI Integration)

Parallel Development Note: Can run concurrently with exporters and UI work; it primarily exercises model list parsing, doc open/close, and inventory generation.

Context
The plan supports batch runs via `pyrevit run` over many local models and recommends inventory generation against the same model list used for export.

Objective
Add a CLI mode that reads a models list, opens each model, generates inventory rows, and writes a combined inventory CSV.

Technical Requirements
1. Primary Component - Implement `--mode=inventory` in the CLI runner that generates a combined inventory CSV.
2. Integration Points - Uses `lib/utils/model_list.py`, `lib/utils/revit.py`, `lib/sheet_query.py`, `lib/sheet_inventory.py`, and `lib/reporting.py`.
3. Data/API Specifications - Define CLI args: `--models`, `--config`, optional `--out`; define deterministic output location when `--out` not provided.
4. Error Handling - Per-model failures are logged and recorded; batch continues to next model.

Files to Create/Modify
* `lib/runner_entry.py` - CLI arg parsing and inventory mode dispatch (create if missing).
* `docs/batch-runner.md` - Document inventory mode commands and outputs.
* `README.md` - Add “Batch Mode” section with inventory mode example.

Implementation Specifications
- Combined inventory CSV should include `model_path` per row so coordinators can filter by model if needed.
- Output defaults to a deterministic path under `output_root` (e.g., `inventory.csv`) unless overridden.

Success Criteria
* Running inventory mode processes all models in the list (skipping failed opens with logged errors).
* Combined CSV is produced and includes rows for each sheet per model.
* All tests pass: N/A (manual Revit validation)

Documentation Updates Required
* Update `docs/batch-runner.md` with inventory mode example commands and prerequisites.
* Update `README.md` with inventory mode usage and what the CSV is for.
* Update `AGENTS.md` with “CLI runner supports modes; keep arg parsing centralized.”
* Create/update docstrings for CLI entrypoint functions.

Validation Commands
bash
Manual (CLI; Revit 2019):
pyrevit run "<path-to-lib/runner_entry.py>" --revit=2019 --purge --models="<path-to-models_2019.txt>" --config="<path-to-config.json>" --mode=inventory

Constraints
* Maximum implementation time: 3 hours
* Target code volume: ~100 lines
* Must maintain backward compatibility
* Follow existing code patterns in `CodexWorkspace/workspace/plans/batch_sheets_plan.md`

---

Task [23]: Implement DWG Export Core (Predefined Setup, One DWG per Sheet) (DWG Export)

Parallel Development Note: Can run concurrently with PDF printing tasks; both are per-sheet operations called by the runner.

Context
DWG output must be unified per sheet (no xrefs) and must use a standardized predefined DWG export setup name.

Objective
Implement per-sheet DWG export using a predefined setup with deterministic naming and resume/overwrite support.

Technical Requirements
1. Primary Component - Export one `ViewSheet` to DWG via `DWGExportOptions.GetPredefinedOptions(doc, setup_name)`.
2. Integration Points - Called by batch runner; used by Validate Setup optional test export.
3. Data/API Specifications - Define result schema: `success`, `out_path`, `skipped`, `duration_ms`, `error_kind`, `error_message`.
4. Error Handling - Per-sheet export exceptions are caught and logged; run continues to next sheet.

Files to Create/Modify
* `lib/export_dwg.py` - `export_sheet_dwg(doc, sheet, cfg, out_dir)` implementation and result schema.
* `docs/dwg-export-setup.md` - Document required setup name and how to create/verify it in Revit UI.
* `docs/config.schema.md` - Ensure DWG config keys are documented (setup name, verify flag, min_size_bytes).

Implementation Specifications
- Resume behavior:
  - In resume mode, skip if expected output DWG exists and size > configured `min_size_bytes`.
  - In overwrite mode, export deterministically (document whether existing file is deleted first).
- Use `lib/utils/netcollections.py` to build the `.NET List[ElementId]` required by `Document.Export`.

Success Criteria
* Export produces exactly one DWG file for a sheet at the expected deterministic path.
* Resume mode correctly skips already-exported files (size threshold applied).
* Per-sheet failures are logged and do not stop subsequent sheets.
* All tests pass: N/A (manual Revit validation)

Documentation Updates Required
* Update `docs/dwg-export-setup.md` with setup creation steps and common pitfalls that cause xrefs.
* Update `docs/config.schema.md` with DWG keys and examples.
* Update `AGENTS.md` with “DWG export must use predefined setup; do not silently mutate options.”
* Create/update docstrings for DWG export functions and result schema.

Validation Commands
bash
Manual (Revit 2019):
1) Ensure the DWG export setup exists with the configured name.
2) Export 1 sheet and confirm the DWG exists and is non-empty.
3) Re-run in resume mode and confirm it skips exporting that sheet.

Constraints
* Maximum implementation time: 3 hours
* Target code volume: ~100 lines
* Must maintain backward compatibility
* Follow existing code patterns in `CodexWorkspace/workspace/plans/batch_sheets_plan.md`

---

Task [24]: Add DWG “No Xrefs” Verification Heuristic (DWG Validation)

Parallel Development Note: Can run concurrently with PDF work; it is an isolated post-export verification step controlled by config.

Context
The plan requires DWG exports to be unified per sheet with no xref by-products; Verify Setup should fail fast if setup produces xrefs.

Objective
Detect likely xref exports by inspecting DWG by-products and flag the export as failed when violations occur.

Technical Requirements
1. Primary Component - Implement pre/post export directory snapshot and validate only one new DWG is produced.
2. Integration Points - Used in `lib/export_dwg.py` and Validate Setup test export.
3. Data/API Specifications - Define `xref_violation` field in result schema and a clear failure message.
4. Error Handling - Verification failures are recorded and logged; run continues to next sheet.

Files to Create/Modify
* `lib/export_dwg.py` - Add verification routine and `xref_violation` field.
* `docs/dwg-export-setup.md` - Document how verification works and remediation steps.
* `docs/reports.md` - Document `xref_violation` in per-sheet results/reports.

Implementation Specifications
- Heuristic (deterministic):
  - Snapshot `.dwg` files in target folder before export.
  - Snapshot after export; compute new `.dwg` files created.
  - Pass if exactly one new DWG exists and matches expected output name; otherwise flag violation.
- Config flag: `dwg.verify_no_xrefs` (default true).

Success Criteria
* Known-bad setups producing multiple DWGs are flagged as `xref_violation`.
* Correct setup producing a single DWG passes verification.
* All tests pass: N/A (manual Revit validation)

Documentation Updates Required
* Update `docs/dwg-export-setup.md` with verification details and troubleshooting.
* Update `docs/reports.md` with how violations appear in reports.
* Update `AGENTS.md` with “Xref verification must be deterministic and config-controlled.”
* Create/update docstrings for verification helpers and result fields.

Validation Commands
bash
Manual (Revit 2019):
1) Run export with a known-bad setup that produces multiple DWGs.
2) Confirm xref violation is detected and recorded.
3) Run export with standardized setup and confirm it passes.

Constraints
* Maximum implementation time: 3 hours
* Target code volume: ~100 lines
* Must maintain backward compatibility
* Follow existing code patterns in `CodexWorkspace/workspace/plans/batch_sheets_plan.md`

---

Task [25]: Implement PDF Printer + Paper Size Discovery (PrintManager) (PDF Prereqs)

Parallel Development Note: Can run concurrently with DWG work; it is a prerequisite for Validate Setup and PDF printing.

Context
PDF output must be created via Revit 2019 PrintManager using a PDF printer configured for silent print-to-file and a required paper size 900×1800 mm.

Objective
Implement discovery and validation helpers for configured printer and paper size.

Technical Requirements
1. Primary Component - Enumerate installed printers and resolve configured printer name.
2. Integration Points - Used by Validate Setup and PDF printing core (`lib/export_pdf.py`).
3. Data/API Specifications - Resolve paper size name under the selected printer; require exact match by default.
4. Error Handling - Missing printer/paper size produces actionable error and blocks printing.

Files to Create/Modify
* `lib/export_pdf.py` - Discovery helpers (installed printers, available paper sizes, resolver).
* `docs/pdf-printing-prereqs.md` - Document required printer configuration and how to create the 900×1800 mm paper size.
* `docs/config.schema.md` - Ensure PDF keys are documented (printer_name, paper_size_name, min_size_bytes).

Implementation Specifications
- Prefer listing printers via .NET installed printers API; then validate that the configured printer exists.
- After selecting printer in PrintManager, enumerate paper sizes and validate required paper size exists (exact match).

Success Criteria
* Validate Setup can confirm configured printer exists and required paper size exists.
* Discovery does not trigger blocking dialogs.
* All tests pass: N/A (manual Revit validation)

Documentation Updates Required
* Update `docs/pdf-printing-prereqs.md` with prerequisites and troubleshooting checklist.
* Update `docs/config.schema.md` with PDF config keys and examples.
* Update `AGENTS.md` with “Always validate printer + paper size before printing.”
* Create/update docstrings for discovery/resolver helpers.

Validation Commands
bash
Manual (Revit 2019):
1) Configure a PDF printer and ensure paper size 900×1800 mm exists.
2) Run discovery helpers (via Validate Setup once wired) and confirm printer + paper are found.

Constraints
* Maximum implementation time: 3 hours
* Target code volume: ~100 lines
* Must maintain backward compatibility
* Follow existing code patterns in `CodexWorkspace/workspace/plans/batch_sheets_plan.md`

---

Task [26]: Implement PDF Export Core (One PDF per Sheet via PrintManager) (PDF Export)

Parallel Development Note: Can run concurrently with DWG export; it is a per-sheet operation that the batch runner calls.

Context
Revit 2019 PDF output must use PrintManager print-to-file with required paper size 900×1800 mm and a silent PDF printer configuration.

Objective
Implement per-sheet PDF printing with deterministic output naming and resume/overwrite support.

Technical Requirements
1. Primary Component - Print a single `ViewSheet` to PDF using PrintManager and print-to-file.
2. Integration Points - Called by batch runner; used by Validate Setup optional test print.
3. Data/API Specifications - Define result schema: `success`, `out_path`, `skipped`, `duration_ms`, `error_kind`, `error_message`.
4. Error Handling - Per-sheet print failures are caught and logged; never block long runs with modal dialogs.

Files to Create/Modify
* `lib/export_pdf.py` - `print_sheet_pdf(doc, sheet, cfg, out_path)` implementation and result schema.
* `docs/pdf-printing-prereqs.md` - Document printer settings required for silent print-to-file.
* `docs/reports.md` - Document PDF result fields and common failure causes.

Implementation Specifications
- Resume behavior:
  - In resume mode, skip if expected output PDF exists and size > configured `pdf.min_size_bytes`.
  - In overwrite mode, re-print deterministically (document how existing files are handled).
- Use explicit Revit API transactions only if required to set up a temporary `ViewSheetSet`; if a transaction is used, it must be explicit and rolled back on exceptions.
- Use `lib/utils/netcollections.py` for any required ViewSet conversions.

Success Criteria
* Printing a sheet produces a PDF at the expected deterministic path and size threshold is met.
* Resume mode skips previously printed PDFs deterministically.
* Per-sheet failures are logged and do not stop remaining sheets.
* All tests pass: N/A (manual Revit validation)

Documentation Updates Required
* Update `docs/pdf-printing-prereqs.md` with supported printers, paper size requirements, and troubleshooting decision tree.
* Update `docs/reports.md` with PDF error fields and common remediation steps.
* Update `AGENTS.md` with “PDF printing must validate prerequisites first; keep batch mode dialog-free.”
* Create/update docstrings for PDF printing functions and result schema.

Validation Commands
bash
Manual (Revit 2019):
1) Ensure the configured PDF printer is installed and supports silent print-to-file.
2) Ensure paper size 900×1800 mm exists and is selectable.
3) Print 1 sheet; confirm output file exists and opens.

Constraints
* Maximum implementation time: 3 hours
* Target code volume: ~100 lines
* Must maintain backward compatibility
* Follow existing code patterns in `CodexWorkspace/workspace/plans/batch_sheets_plan.md`

---

Task [27]: Implement Validate Setup Core (DWG + PDF + Output Root) (Diagnostics)

Parallel Development Note: Can run concurrently with runner/UI wiring once DWG/PDF prerequisite helpers exist; it reduces field failures early.

Context
The plan requires a Validate Setup command to confirm DWG export setup exists (and is no-xref) and PDF printing prerequisites are satisfied before production.

Objective
Implement a validator that checks environment/config prerequisites and produces an actionable validation report.

Technical Requirements
1. Primary Component - Validate config, output root writability, DWG setup presence, PDF printer presence, paper size presence.
2. Integration Points - Used by Validate Setup pushbutton and optionally by CLI runner as a pre-flight gate.
3. Data/API Specifications - Define a structured validation report (items with pass/fail, message, remediation).
4. Error Handling - Fail fast with actionable messages; never leave temp artifacts; explicit RollBack on any transaction used.

Files to Create/Modify
* `lib/validate.py` - `validate_setup(doc, cfg)` implementation and report schema.
* `docs/dwg-export-setup.md` - Document what Validate Setup checks for DWG.
* `docs/pdf-printing-prereqs.md` - Document what Validate Setup checks for PDF.
* `docs/batch-runner.md` - Document Validate Setup as a required pre-flight step.

Implementation Specifications
- Checks:
  - Config loads and validates (via `lib/config.py`).
  - Output root exists/creatable and is writable (create temp file and delete).
  - DWG predefined setup exists (`GetPredefinedOptions` returns non-null).
  - Printer exists and required paper size exists.
- Optional (config-flagged) test operations:
  - DWG: export 1 small sheet to temp folder and run no-xref verification heuristic.
  - PDF: print 1 sheet to temp PDF and validate file creation and size threshold.

Success Criteria
* Validate Setup report clearly indicates pass/fail for each prerequisite with remediation guidance.
* Optional tests (if enabled) run without blocking dialogs in production configuration.
* All tests pass: N/A (manual Revit validation)

Documentation Updates Required
* Update `docs/dwg-export-setup.md` and `docs/pdf-printing-prereqs.md` with Validate Setup checklist and remediation.
* Update `README.md` with “Run Validate Setup before batch runs” requirement.
* Update `AGENTS.md` with “Validate Setup is a gate; do not run long batches without it.”
* Create/update docstrings for validator functions and report schema.

Validation Commands
bash
Manual (Revit 2019):
1) Open a discipline model.
2) Run Validate Setup (once wired).
3) Confirm it detects missing printer/paper/setup and provides remediation steps.

Constraints
* Maximum implementation time: 3 hours
* Target code volume: ~100 lines
* Must maintain backward compatibility
* Follow existing code patterns in `CodexWorkspace/workspace/plans/batch_sheets_plan.md`

---

Task [28]: Wire Validate Setup Pushbutton + Report Rendering (UI Integration)

Parallel Development Note: Can run concurrently with Export Sheets UI; it depends only on config loader and validate core.

Context
Users must run Validate Setup inside Revit and receive a clear, readable pass/fail report.

Objective
Implement the Validate Setup pushbutton orchestration and present a structured report in pyRevit output.

Technical Requirements
1. Primary Component - Load config, run `validate_setup`, and render report (with links to docs).
2. Integration Points - Uses `lib/config.py`, `lib/validate.py`, and `lib/reporting.py` for error logging.
3. Data/API Specifications - Render stable report format: list of checks, pass/fail, remediation.
4. Error Handling - Any exception is logged; user sees a friendly summary message (no raw traceback).

Files to Create/Modify
* `BaseDarom.exporter.tab/Export Plans.panel/Validate Setup.pushbutton/script.py` - Thin orchestrator.
* `lib/ui.py` - Optional `render_validation_report(report)` helper reused by CLI output (if desired).
* `README.md` - Add Validate Setup runbook steps.

Implementation Specifications
- Keep `script.py` under ~60 LoC: resolve active doc, load config path, call validate, render.
- Avoid modal dialogs except for a single final summary if desired; batch mode must remain dialog-free.

Success Criteria
* Validate Setup button produces a readable checklist with pass/fail statuses.
* Missing prerequisites yield actionable remediation guidance and log locations.
* All tests pass: N/A (manual Revit validation)

Documentation Updates Required
* Update `README.md` with Validate Setup instructions and prerequisites.
* Update `AGENTS.md` with “UI scripts are thin; rendering helpers live in `lib/ui.py`.”
* Create/update docstrings for any UI rendering helpers.
* Update API documentation if applicable: N/A

Validation Commands
bash
Manual (Revit 2019):
1) Open a model.
2) Click Validate Setup.
3) Confirm report renders and failures are actionable.

Constraints
* Maximum implementation time: 3 hours
* Target code volume: ~100 lines
* Must maintain backward compatibility
* Follow existing code patterns in `Coordination.tab/3D Views.panel/Scope Boxs to 3D Views.pushbutton/script.py`

---

Task [29]: Implement Report Writers (`report.model.*`, `report.structure.*`) (Reporting)

Parallel Development Note: Can run concurrently with batch runner development; schema and writers are largely pure-python and can be unit-tested.

Context
The plan requires detailed reports: per-model JSON/CSV and per-structure aggregation JSON/CSV, plus `errors.log`.

Objective
Define report schemas and implement writers that produce deterministic, reconciliation-friendly outputs.

Technical Requirements
1. Primary Component - Implement per-model report writers (`report.model.json` and `report.model.csv`) and structure aggregation writers.
2. Integration Points - Called by batch runner after each sheet attempt; CLI runner writes structure-level aggregates after processing models.
3. Data/API Specifications - Define stable CSV headers and JSON schema fields (model identity, sheet identity, statuses, output paths, durations, errors).
4. Error Handling - Report write failures are logged and do not crash exports; JSON writes must be atomic.

Files to Create/Modify
* `lib/reporting.py` - Report schema definitions and writers (model + structure).
* `tests/test_reporting_schema.py` - Unit tests for CSV headers, JSON keys, and deterministic ordering (pure python).
* `docs/reports.md` - Document report file schemas and locations.

Implementation Specifications
- Per-sheet row fields must include at minimum:
  - model: `model_path`, `structure_id`, `discipline_id`
  - sheet: `sheet_number`, `sheet_name`, (optionally) sheet ElementId int
  - export results: dwg/pdf success, output path, skipped, duration, error fields, xref violation flag
- Structure report aggregates totals per discipline and overall.

Success Criteria
* After a run, report files exist and reconcile with produced outputs.
* CSV headers are stable and documented; JSON files are valid even after partial runs (atomic writes).
* All tests pass: python -m unittest discover -s tests -p "test_*.py"

Documentation Updates Required
* Update `docs/reports.md` with full schemas and “how to interpret” guidance.
* Update `README.md` with report locations and what to check after runs.
* Update `AGENTS.md` with “Report schemas are stable; changes require docs + tests.”
* Create/update docstrings for report writer APIs and schema constants.

Validation Commands
bash
python -m unittest discover -s tests -p "test_*.py"

Constraints
* Maximum implementation time: 3 hours
* Target code volume: ~100 lines
* Must maintain backward compatibility
* Follow existing code patterns in `CodexWorkspace/workspace/plans/batch_sheets_plan.md`

---

Task [30]: Implement Checkpointing + Resume/Overwrite Policy (Resiliency)

Parallel Development Note: Can run concurrently with batch runner core loop; once implemented, it enables resumable runs and deterministic reruns.

Context
The plan requires resumable exports, per-sheet try/except, deterministic outputs, and checkpoint files written after each sheet attempt.

Objective
Implement checkpoint files and resume/overwrite decision logic used by the batch runner.

Technical Requirements
1. Primary Component - Read/write checkpoint state after each sheet attempt with atomic writes.
2. Integration Points - Used by `lib/batch_runner.py` to skip work in resume mode and to record partial progress.
3. Data/API Specifications - Define checkpoint schema: model identity, sheet number/id, per-kind status (dwg/pdf), output paths, timestamps, last error.
4. Error Handling - Corrupted checkpoint files must be handled deterministically (fail fast or backup+restart per config).

Files to Create/Modify
* `lib/checkpoint.py` - Checkpoint schema, read/write helpers, atomic write usage.
* `docs/batch-runner.md` - Document checkpoint location/schema and resume/overwrite semantics.

Implementation Specifications
- Resume decision must check both:
  - checkpoint says success AND output file exists and size > threshold
  - if either is missing, re-export (deterministic)
- Always record failures too (for audit), not only successes.

Success Criteria
* Interrupting a run and re-running in resume mode skips completed sheet exports deterministically.
* Overwrite mode re-exports deterministically and updates checkpoints accordingly.
* All tests pass: python -m unittest discover -s tests -p "test_*.py" (pure checkpoint IO tests) and N/A (manual Revit validation for full behavior)

Documentation Updates Required
* Update `docs/batch-runner.md` with checkpoint schema and resume/overwrite rules.
* Update `README.md` with resume/overwrite behavior summary.
* Update `AGENTS.md` with “Checkpoint written after every attempt; changes require docs + tests.”
* Create/update docstrings for checkpoint IO functions and schema.

Validation Commands
bash
python -m unittest discover -s tests -p "test_*.py"
Manual (Revit 2019): run export, stop mid-way, rerun resume and confirm skips; rerun overwrite and confirm re-exports.

Constraints
* Maximum implementation time: 3 hours
* Target code volume: ~100 lines
* Must maintain backward compatibility
* Follow existing code patterns in `CodexWorkspace/workspace/plans/batch_sheets_plan.md`

---

Task [31]: Implement Single-Document Export Orchestration (Per-Sheet Loop) (Batch Runner)

Parallel Development Note: Can run concurrently with CLI/UI wiring once exporters, reporting, status, and checkpoint APIs exist.

Context
The plan requires per-sheet try/except, deterministic ordering, resumable behavior, and live status updates during long runs.

Objective
Implement the core export loop for one already-open Revit document: select sheets, export DWG/PDF per sheet, and record results.

Technical Requirements
1. Primary Component - Implement `run_model(doc, cfg, model_identity)` that iterates selected sheets and calls exporters.
2. Integration Points - Uses `lib/sheet_query.py`, `lib/export_dwg.py`, `lib/export_pdf.py`, `lib/reporting.py`, `lib/checkpoint.py`.
3. Data/API Specifications - Define per-sheet result object passed to reporting/status/checkpoint writers; ensure deterministic ordering.
4. Error Handling - Per-sheet try/except continues; any transaction used must be explicit and rolled back on exceptions.

Files to Create/Modify
* `lib/batch_runner.py` - Single-document export orchestration and summary return schema.
* `docs/batch-runner.md` - Document per-sheet failure handling and continuation semantics.
* `docs/reports.md` - Document how runner populates report/status fields.

Implementation Specifications
- Steps per document (deterministic):
  - Build `model_identity` (structure/discipline) and output root.
  - Select sheets per config mode (folder/selection-file/codes) and order deterministically.
  - For each sheet:
    - Update `progress.json` (current sheet and percent).
    - Check resume policy (checkpoint + filesystem size thresholds).
    - Export DWG (if enabled) and PDF (if enabled) with per-sheet try/except.
    - Update `failed_sheets.csv` immediately on failure(s).
    - Update per-model report and checkpoint after each sheet attempt.
- Provide a cancellation hook (checked between sheets) so UI can stop safely.

Success Criteria
* For an open model, exports selected sheets and produces outputs + reports + status + checkpoints.
* One sheet failure does not stop remaining sheets.
* Resume behavior is deterministic with checkpoint + file-size thresholds.
* All tests pass: N/A (manual Revit validation)

Documentation Updates Required
* Update `docs/batch-runner.md` with runner semantics and per-sheet error handling.
* Update `README.md` with “What happens during a run” summary (status files, reports, checkpoints).
* Update `AGENTS.md` with “Runner is the single source of truth for export loops; exporters are per-sheet.”
* Create/update docstrings for runner APIs and result schemas.

Validation Commands
bash
Manual (Revit 2019):
1) Open a model and run Export Sheets (once wired).
2) Confirm progress.json updates each sheet and failures are recorded without stopping the run.

Constraints
* Maximum implementation time: 3 hours
* Target code volume: ~100 lines
* Must maintain backward compatibility
* Follow existing code patterns in `CodexWorkspace/workspace/plans/batch_sheets_plan.md`

---

Task [32]: Implement CLI Batch Export Runner (`pyrevit run`) (Batch Runner)

Parallel Development Note: Can run concurrently with interactive UI wizard; both share `lib/batch_runner.py` and differ only in document acquisition and input parsing.

Context
The plan requires batch orchestration via `pyrevit run` across many models with resumable behavior, per-sheet try/except, and dialog-free operation.

Objective
Implement `lib/runner_entry.py` export mode that opens models, runs export orchestration, and closes docs reliably.

Technical Requirements
1. Primary Component - Implement `--mode=export` with args: `--models`, `--config`, optional overrides (resume/overwrite), optional output root override.
2. Integration Points - Uses `lib/utils/model_list.py`, `lib/utils/revit.py`, `lib/bim360_adapter.py`, `lib/batch_runner.py`, `lib/reporting.py`.
3. Data/API Specifications - Define CLI exit code semantics (documented) and deterministic output folder layout per model.
4. Error Handling - Per-model failures are logged and do not stop the batch; docs always closed in `finally`.

Files to Create/Modify
* `lib/runner_entry.py` - CLI arg parsing + dispatch for export/inventory (extend inventory mode from Task 22).
* `docs/batch-runner.md` - Document export mode commands, flags, and troubleshooting (`--allowdialogs` guidance).
* `README.md` - Add batch export command examples and prerequisites.

Implementation Specifications
- Batch behavior:
  - Validate config once at start; optionally run Validate Setup gate per model (config flag).
  - Iterate models list; open doc; run `batch_runner.run_model`; close doc.
  - Write/update structure-level reports after processing models for the same structure (or at end of run).
- Ensure outputs are deterministic and resumable; avoid timestamps in filenames unless explicitly configured.

Success Criteria
* Running export mode over a models list processes each model and produces expected outputs + reports.
* A failing model does not stop the batch; failure is captured in logs and summary.
* All tests pass: N/A (manual Revit validation)

Documentation Updates Required
* Update `docs/batch-runner.md` with export mode command examples and common failures (hydration, locked files, printers).
* Update `README.md` with copy/paste CLI commands and warnings (use dialog allowance for troubleshooting only).
* Update `AGENTS.md` with “Batch mode is dialog-free; per-model failures continue; always close docs.”
* Create/update docstrings for CLI entrypoints and argument parsing behavior.

Validation Commands
bash
Manual (CLI; Revit 2019):
pyrevit run "<path-to-lib/runner_entry.py>" --revit=2019 --purge --models="<path-to-models_2019.txt>" --config="<path-to-config.json>" --mode=export

Constraints
* Maximum implementation time: 3 hours
* Target code volume: ~100 lines
* Must maintain backward compatibility
* Follow existing code patterns in `CodexWorkspace/workspace/plans/batch_sheets_plan.md`

---

Task [33]: Implement Interactive Export Sheets Wizard (Active Document) (UI Integration)

Parallel Development Note: Can run concurrently with CLI runner once core export orchestration exists; UI focuses on input collection and progress presentation.

Context
The plan specifies an interactive workflow: pick output root, pick selection mode (folder/codes/selection file), choose resume/overwrite, run export, show progress, allow cancellation between sheets.

Objective
Implement the Export Sheets pushbutton workflow for active document using a guided UI and progress reporting.

Technical Requirements
1. Primary Component - Build a wizard to collect output root and run mode, then run export on the active doc with a progress bar.
2. Integration Points - Uses `lib/config.py` for defaults, `lib/batch_runner.py` for execution, `lib/reporting.py` for status files.
3. Data/API Specifications - Define UI override precedence vs config (documented).
4. Error Handling - Cancels exit cleanly; per-sheet failures summarized at end (no modal per-sheet alerts).

Files to Create/Modify
* `BaseDarom.exporter.tab/Export Plans.panel/Export Sheets.pushbutton/script.py` - Thin orchestrator calling `lib/ui.py`.
* `lib/ui.py` - Wizard UI and progress rendering; cancellation hook integration.
* `README.md` - Add interactive runbook steps.

Implementation Specifications
- UI steps:
  - Select output root folder.
  - Choose selection mode: folder (default), selection file, or codes.
  - Choose resume vs overwrite.
  - Optional: choose selection file path / codes list path (if mode requires).
  - Run export; show progress; allow cancel between sheets.
- End summary:
  - Output locations, report paths, failures count, pointer to `failed_sheets.csv`.

Success Criteria
* Interactive export runs for an open model and produces outputs + reports.
* Progress updates during run and cancel stops safely between sheets.
* All tests pass: N/A (manual Revit validation)

Documentation Updates Required
* Update `README.md` with interactive Export Sheets workflow steps and expected outputs.
* Update `AGENTS.md` with “UI scripts orchestrate only; logic in `lib/batch_runner.py`.”
* Create/update docstrings for UI helper functions and wizard flow.
* Update API documentation if applicable: N/A

Validation Commands
bash
Manual (Revit 2019):
1) Open a model with multiple sheets.
2) Click Export Sheets and run folder mode.
3) Confirm progress updates and outputs created.
4) Cancel mid-run and confirm clean stop + reports/checkpoints remain valid.

Constraints
* Maximum implementation time: 3 hours
* Target code volume: ~100 lines
* Must maintain backward compatibility
* Follow existing code patterns in `Coordination.tab/3D Views.panel/Scope Boxs to 3D Views.pushbutton/script.py`

---

Task [34]: Implement Optional Manifest Schema + Build Manifest Tool (Optional Workflow) (CLI/UI Integration)

Parallel Development Note: Mostly pure JSON parsing/writing; can run concurrently with export runner work as an optional enhancement.

Context
The plan includes an optional `manifest.json` mapping structures -> disciplines -> model identifiers to help manage multi-model runs.

Objective
Define a versioned manifest schema (local-path-first) and implement tools to build/consume it.

Technical Requirements
1. Primary Component - Implement manifest schema, reader, writer, and builder from a models list.
2. Integration Points - CLI runner optionally accepts `--manifest` as an alternative input to `--models`.
3. Data/API Specifications - Manifest must group by structure/discipline and preserve deterministic ordering; schema version is explicit.
4. Error Handling - Unknown inference grouped under `Unknown`; schema validation errors are actionable.

Files to Create/Modify
* `lib/manifest.py` - Manifest schema and builder/reader helpers.
* `BaseDarom.exporter.tab/Export Plans.panel/Build Manifest.pushbutton/script.py` - Thin orchestrator to build manifest from a models list (optional).
* `docs/bim360-manifest.md` - Document manifest schema and usage (local-path-first).
* `CodexWorkspace/workspace/batch_sheets/manifest.json` - Example manifest.

Implementation Specifications
- Manifest fields should include: structure id, discipline id, model path, and optional notes fields.
- Keep BIM 360 GUID-based opening out of scope unless explicitly enabled by feature flag and verified in the environment.

Success Criteria
* Manifest builder produces a valid manifest JSON grouping models by structure/discipline.
* CLI runner can read manifest and process the referenced models deterministically.
* All tests pass: python -m unittest discover -s tests -p "test_*.py" (manifest parsing/building tests)

Documentation Updates Required
* Update `docs/bim360-manifest.md` with schema, examples, and CLI usage.
* Update `README.md` with optional manifest workflow description.
* Update `AGENTS.md` with “Manifest schema is versioned; changes require docs + tests.”
* Create/update docstrings for manifest helpers and schema.

Validation Commands
bash
python -m unittest discover -s tests -p "test_*.py"
Manual (Revit 2019): (optional) click Build Manifest, generate manifest, then run CLI with --manifest.

Constraints
* Maximum implementation time: 3 hours
* Target code volume: ~100 lines
* Must maintain backward compatibility
* Follow existing code patterns in `CodexWorkspace/workspace/plans/batch_sheets_plan.md`

---

Task [35]: Finalize Documentation Pack + Runbook (Docs)

Parallel Development Note: Can be updated continuously in parallel with implementation; final pass depends on stable filenames/flags/schemas.

Context
The plan requires docs for config, selection logic, inventory workflow, DWG setup, PDF prerequisites, batch runner usage, and report interpretation.

Objective
Fill in the `docs/*.md` stubs and update `README.md` so operators can run the tool end-to-end.

Technical Requirements
1. Primary Component - Finalize all docs referenced by the plan with step-by-step instructions and troubleshooting.
2. Integration Points - Docs must reference actual implemented file paths, tool names, CLI flags, and output schemas.
3. Data/API Specifications - Include copy/paste config examples, schema tables, and sample command lines.
4. Error Handling - Document common failures (hydration, printer/paper size, xref violations, locked files, interruptions) and remediation.

Files to Create/Modify
* `docs/config.schema.md` - Final schema + examples.
* `docs/sheet-selection.md` - Folder/selection-file/codes modes + ordering rules.
* `docs/sheet-inventory.md` - Inventory schema + coordinator markup workflow.
* `docs/dwg-export-setup.md` - Setup creation + no-xref verification guidance.
* `docs/pdf-printing-prereqs.md` - Printer setup + 900×1800 mm paper guidance + troubleshooting.
* `docs/batch-runner.md` - CLI usage, resume/overwrite, checkpoints, hydration warnings.
* `docs/reports.md` - Report schemas and how to interpret outputs.
* `README.md` - Complete runbook summary and links to docs.

Implementation Specifications
- Each doc should include: Purpose, Prereqs, Steps, Outputs, Troubleshooting, and where logs live.
- Ensure documentation matches IronPython 2.7 constraints and Revit 2019 behavior (no Revit 2023-only APIs in examples).

Success Criteria
* A new operator can follow README + docs to run Validate Setup, generate inventory, and run export (interactive or batch).
* Docs match implemented flags/paths/schemas exactly.
* All tests pass: N/A (documentation-only)

Documentation Updates Required
* Update `README.md` with “Quick Start” + “Batch Mode” + “Troubleshooting” sections referencing the docs pack.
* Update `AGENTS.md` with any new patterns introduced during implementation (schemas, result objects, runner modes).
* Create/update docstrings for all new functions (completion gate referenced by docs).
* Update API documentation if applicable: N/A

Validation Commands
bash
Light checks:
ls docs
rg -n "docs/" README.md

Constraints
* Maximum implementation time: 3 hours
* Target code volume: ~100 lines
* Must maintain backward compatibility
* Follow existing code patterns in `CodexWorkspace/workspace/plans/batch_sheets_plan.md`

---

Task [36]: Execute Pilot Acceptance Run + Record Results (QA/Hardening)

Parallel Development Note: Final integration step; depends on core exporters, runner, reporting, and Validate Setup being implemented.

Context
The plan defines acceptance on one structure across five discipline models: exports complete (or failures are reported), DWGs have no xrefs, output structure is deterministic, and reports reconcile counts.

Objective
Run a pilot on one Base Darom structure across the five discipline models and record outcomes and remediation steps.

Technical Requirements
1. Primary Component - Execute the full workflow (Validate Setup -> inventory -> export in folder mode and selection-file mode -> resume test).
2. Integration Points - Validate end-to-end behavior: UI + CLI + reporting + live status + checkpoints.
3. Data/API Specifications - Verify output structure `/<Structure>/<Discipline>/DWG` and `/<Structure>/<Discipline>/PDF` plus reports/status/checkpoints.
4. Error Handling - Confirm per-sheet failure isolation, deterministic reruns, and resumable behavior after interruptions.

Files to Create/Modify
* `CodexWorkspace/workspace/reviews/pilot_batch_sheets_YYYYMMDD.md` - Pilot checklist, results, and issues found.
* `CodexWorkspace/workspace/backlog.md` - Capture non-blocking follow-ups discovered during pilot.
* `docs/batch-runner.md` - Update with real-world troubleshooting findings from pilot.

Implementation Specifications
- Pilot checklist must include:
  - Validate Setup passes (DWG setup + PDF printer + 900×1800 paper size).
  - Export in folder mode for one structure across 5 disciplines.
  - Verify DWG no-xref heuristic behavior.
  - Verify `progress.json` and `failed_sheets.csv` update live.
  - Interrupt and re-run in resume mode; confirm deterministic skip behavior.
  - Verify reports reconcile with actual outputs.

Success Criteria
* Pilot run produces expected outputs for one structure across five disciplines or records failures with actionable reasons.
* Output folder structure is deterministic and matches documentation.
* Reports reconcile counts and provide sufficient detail for coordinator filtering.
* All tests pass: N/A (manual Revit validation)

Documentation Updates Required
* Update `README.md` with the pilot acceptance checklist and expected outputs.
* Update `AGENTS.md` with any new hardening conventions introduced during pilot fixes.
* Create/update docstrings for any new functions added during hardening.
* Update API documentation if applicable: N/A

Validation Commands
bash
Manual (Revit 2019 + CLI):
1) Run Validate Setup for each discipline model template/version.
2) Run `pyrevit run` export mode for one structure across 5 models.
3) Verify outputs, reports, status updates, and resume behavior after interruption.

Constraints
* Maximum implementation time: 3 hours
* Target code volume: ~100 lines
* Must maintain backward compatibility
* Follow existing code patterns in `CodexWorkspace/docs/agent/checklists/testing_strategy.md`

---
