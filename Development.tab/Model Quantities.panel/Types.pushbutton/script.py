# -*- coding: utf-8 -*-
"""
pyRevit IronPython 2.7 script:
Exports all element TYPES in the Revit model to a CSV with columns:
A: Category
B: Family
C: Type Name
D: Count of Instances
E..: Type Parameters

Usage:
 1) Click this button to run.
 2) Choose a folder to save the CSV.
 3) Open the CSV in Excel.

Author: ChatGPT, for Revit 2019 + IronPython 2.7
"""

import csv
import sys
import os

from pyrevit import revit, forms
from Autodesk.Revit.DB import (
    FilteredElementCollector,
    StorageType,
    ElementId
)

doc = revit.doc

def get_param_value(param):
    """Safely retrieve parameter values as strings for CSV export."""
    if not param or not param.HasValue:
        return ""
    st = param.StorageType
    if st == StorageType.Double:
        return str(param.AsDouble())
    elif st == StorageType.Integer:
        return str(param.AsInteger())
    elif st == StorageType.ElementId:
        eid = param.AsElementId()
        if eid != ElementId.InvalidElementId:
            linked_elem = doc.GetElement(eid)
            if linked_elem:
                return linked_elem.Name
        return ""
    elif st == StorageType.String:
        return param.AsString() or ""
    return ""

# 1. Collect all types and all instances in the model
all_types = FilteredElementCollector(doc).WhereElementIsElementType().ToElements()
all_instances = FilteredElementCollector(doc).WhereElementIsNotElementType().ToElements()

# 2. Build a map: TypeId -> count of instances
instance_count_by_type_id = {}
for inst in all_instances:
    t_id = inst.GetTypeId()
    if t_id not in instance_count_by_type_id:
        instance_count_by_type_id[t_id] = 0
    instance_count_by_type_id[t_id] += 1

# 3. Gather parameter data. We'll keep track of all param names for columns.
all_param_names = set()
all_type_data = []  # (Category, Family, Type Name, Count, {paramName -> paramValue})

for t in all_types:
    cat = t.Category
    if not cat:
        # skip items with no valid category
        continue
    cat_name = cat.Name
    # Some element types are FamilySymbols; we'll try to get FamilyName
    fam_name = ""
    try:
        fam_name = t.FamilyName
    except:
        pass

    type_name = t.Name
    type_id = t.Id

    # Count of instances
    count_instances = instance_count_by_type_id.get(type_id, 0)

    # Gather the type's parameters
    param_map = {}
    for p in t.GetOrderedParameters():
        p_name = p.Definition.Name
        p_val = get_param_value(p)
        param_map[p_name] = p_val
        all_param_names.add(p_name)

    all_type_data.append((cat_name, fam_name, type_name, count_instances, param_map))

# Sort the parameter names for consistent columns
param_name_list = sorted(list(all_param_names))

# Columns:
# A = Category, B = Family, C = Type Name, D = Count, E.. = parameters
header = ["Category", "Family", "Type", "Count"] + param_name_list

# 4. Prompt user for a folder to save the file
save_folder = forms.pick_folder()
if not save_folder:
    sys.exit()  # user canceled

csv_path = os.path.join(save_folder, "Export_Types_and_Parameters.csv")

# 5. Write the CSV
with open(csv_path, 'wb') as f:
    writer = csv.writer(f)
    writer.writerow(header)
    
    for (cat_name, fam_name, type_name, count_instances, param_map) in all_type_data:
        row = [cat_name, fam_name, type_name, count_instances]
        for pname in param_name_list:
            row.append(param_map.get(pname, ""))
        writer.writerow(row)

forms.alert(
    "Export Complete!\nSaved: {}".format(os.path.basename(csv_path)),
    exitscript=True
)
