# Environment Setup Snapshot

- Use **pyRevit 5.2** with the **CPython 3.9** engine (`#! python3` header in every script).
- Target **Autodesk Revit 2023** APIs; remove or avoid 2019 compatibility shims.
- Keep a consistent CPython engine across the team to avoid behavioral drift.
- Treat ElementIds as opaque tokens (Revit 2023 moves toward 64-bit ids).
- Favor modern creation APIs (e.g., `Floor.Create`, `ViewSheet.Duplicate`) over removed `Document.New*` calls.

