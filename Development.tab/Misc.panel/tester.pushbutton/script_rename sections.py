#! python3
# Replace substrings in view names across the model.

from Autodesk.Revit.DB import (
    FilteredElementCollector, View, ViewSchedule, Transaction,
    ElementId, ViewType
)
import re

# --- configure ---
OLD = "Sec. "          # text to find
NEW = "S"          # replacement
CASE_SENSITIVE = False
INCLUDE_DEPENDENTS = False # skip dependent sections by default
# ------------------

doc = __revit__.ActiveUIDocument.Document

def is_dependent(v):
    return v.GetPrimaryViewId() != ElementId.InvalidElementId

# All non-template views for global name-uniqueness checks
all_views = [v for v in FilteredElementCollector(doc).OfClass(View) if not v.IsTemplate]
existing_names = {v.Name for v in all_views}

# Targets: SECTION views only
targets = [
    v for v in all_views
    if v.ViewType == ViewType.Section
    and not isinstance(v, ViewSchedule)
    and (INCLUDE_DEPENDENTS or not is_dependent(v))
]

pattern = re.compile(re.escape(OLD), 0 if CASE_SENSITIVE else re.IGNORECASE)

renamed, skipped_conflict, skipped_nohit = [], [], []

t = Transaction(doc, "Replace in SECTION view names")
t.Start()
try:
    for v in targets:
        if not pattern.search(v.Name):
            skipped_nohit.append(v.Name)
            continue
        new_name = pattern.sub(NEW, v.Name)
        if new_name != v.Name and new_name in existing_names:
            skipped_conflict.append((v.Name, new_name))
            continue
        try:
            v.Name = new_name
            existing_names.add(new_name)
            renamed.append((v.Id.IntegerValue, v.Name))
        except Exception as e:
            skipped_conflict.append((v.Name, "ERROR: {}".format(e)))
    t.Commit()
except:
    t.RollBack()
    raise

print("Renamed:", len(renamed))
print("Skipped (no match):", len(skipped_nohit))
print("Skipped (conflict/error):", len(skipped_conflict))
