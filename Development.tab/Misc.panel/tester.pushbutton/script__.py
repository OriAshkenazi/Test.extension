#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
Wall and Floor Copier from Linked Model
Prompts the user to select a linked model from a list, then copies all wall and floor elements
from that link into the active document, placing them in the same location and assigning them to the 'Finish' workset.
"""

__title__ = "Copy Walls and Floors from Link"
__author__ = "Ori Ashkenazi"
__doc__ = "Prompts for a link, then copies walls and floors from that link into the active model's 'Finish' workset."


# ────────────────────────────────────────────────
# IMPORTS
# ────────────────────────────────────────────────
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import TaskDialog
from Autodesk.Revit.DB.Structure import StructuralType
from System.Collections.Generic import List

from pyrevit import script
from pyrevit.forms import SelectFromList

# ────────────────────────────────────────────────
# SETUP
# ────────────────────────────────────────────────
uidoc = __revit__.ActiveUIDocument
doc = uidoc.Document
app = doc.Application
output = script.get_output()
logger = script.get_logger()

# ────────────────────────────────────────────────
# SELECT LINK FROM LIST
# ────────────────────────────────────────────────
link_instances = []
collector = FilteredElementCollector(doc).OfClass(RevitLinkInstance)
for inst in collector:
    link_doc = inst.GetLinkDocument()
    if link_doc:
        link_instances.append(inst)

if not link_instances:
    TaskDialog.Show("Error", "No loaded Revit links found in the project.")
    script.exit()

# Sort by name and present UI
link_name_to_instance = {}
link_names = []
for inst in link_instances:
    name = inst.Name
    link_name_to_instance[name] = inst
    link_names.append(name)

link_names.sort()

selected_name = SelectFromList.show(link_names, title="Select Linked Model", multiselect=False)

if not selected_name:
    script.exit()

link_instance = link_name_to_instance[selected_name]
link_doc = link_instance.GetLinkDocument()
transform = link_instance.GetTransform()

# ────────────────────────────────────────────────
# GET WALLS AND FLOORS
# ────────────────────────────────────────────────
walls = FilteredElementCollector(link_doc).OfClass(Wall).ToElementIds()
floors = FilteredElementCollector(link_doc).OfClass(Floor).ToElementIds()
elements_to_copy = List[ElementId]()
for eid in walls:
    elements_to_copy.Add(eid)
for eid in floors:
    elements_to_copy.Add(eid)

if elements_to_copy.Count == 0:
    TaskDialog.Show("Info", "No walls or floors found in selected link.")
    script.exit()

# ────────────────────────────────────────────────
# GET FINISH WORKSET ID
# ────────────────────────────────────────────────
def get_workset_id_by_name(name):
    collector = FilteredWorksetCollector(doc).OfKind(WorksetKind.UserWorkset)
    for ws in collector:
        if ws.Name == name:
            return ws.Id
    raise Exception("Workset '{}' not found.".format(name))

try:
    finish_workset_id = get_workset_id_by_name("Finish")
except Exception as e:
    TaskDialog.Show("Error", str(e))
    script.exit()

# ────────────────────────────────────────────────
# COPY ELEMENTS
# ────────────────────────────────────────────────
with Transaction(doc, "Copy Walls and Floors from Link") as t:
    t.Start()

    copy_options = CopyPasteOptions()
    new_element_ids = ElementTransformUtils.CopyElements(
        link_doc,
        elements_to_copy,
        doc,
        transform,
        copy_options
    )

    for eid in new_element_ids:
        elem = doc.GetElement(eid)
        param = elem.get_Parameter(BuiltInParameter.ELEM_PARTITION_PARAM)
        if param and not param.IsReadOnly:
            param.Set(finish_workset_id.IntegerValue)

    t.Commit()

# ────────────────────────────────────────────────
# REPORT
# ────────────────────────────────────────────────
output.print_md("### ✅ Copied {} elements".format(len(new_element_ids)))
output.print_md("From link: **{}**".format(selected_name))
output.print_md("Assigned all to workset: **Finish**")
