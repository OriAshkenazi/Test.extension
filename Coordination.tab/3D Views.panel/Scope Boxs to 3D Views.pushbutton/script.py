# -*- coding: utf-8 -*-
# pyRevit: Create 3D views per selected Scope Box and apply template

from Autodesk.Revit.DB import (
    FilteredElementCollector, BuiltInCategory, BuiltInParameter,
    Transaction, TransactionGroup,
    ViewFamilyType, ViewFamily, View3D, View, ViewType,
    BoundingBoxXYZ, XYZ
)
from System.Collections.Generic import List

try:
    from pyrevit import forms
    uiapp = __revit__
    uidoc = uiapp.ActiveUIDocument
    doc = uidoc.Document
except Exception:
    raise Exception("Run from pyRevit inside Revit.")

# ---------- helpers ----------
def all_scope_boxes(d):
    return list(
        FilteredElementCollector(d)
        .OfCategory(BuiltInCategory.OST_VolumeOfInterest)
        .WhereElementIsNotElementType()
    )

def sb_name(e):
    try:
        return e.Name
    except:
        p = e.LookupParameter("Name")
        return p.AsString() if p else "<unnamed>"

def sanitize_name(s):
    bad = u'\\/:*?"<>|'
    for ch in bad:
        s = s.replace(ch, "_")
    return s

def existing_view_names(d):
    return set(v.Name for v in FilteredElementCollector(d).OfClass(View))

def unique_name(base, name_set):
    name = base
    i = 1
    while name in name_set:
        name = u"{} ({})".format(base, i)
        i += 1
    name_set.add(name)
    return name

def get_3d_vft(d):
    for vft in FilteredElementCollector(d).OfClass(ViewFamilyType):
        if vft.ViewFamily == ViewFamily.ThreeDimensional:
            return vft
    return None

def find_template_by_name(d, name):
    # must be a 3D view template
    for v in FilteredElementCollector(d).OfClass(View):
        if not v.IsTemplate: continue
        if v.Name == name:
            # ensure template originates from a 3D view
            try:
                if v.ViewType == ViewType.ThreeD:
                    return v
            except:
                return v  # some builds lack ViewType on template; trust name
    return None

def scope_box_parameter(view):
    """Return the Parameter controlling the view's scope box, if any."""
    scope_param = None
    bip = getattr(BuiltInParameter, "VIEWER_VOLUME_OF_INTEREST", None)
    if bip is not None:
        scope_param = view.get_Parameter(bip)
    if not scope_param:
        scope_param = view.LookupParameter("Scope Box")
    return scope_param

# ---------- collect selection ----------
pre_ids = list(uidoc.Selection.GetElementIds())
pre_sb = []
for eid in pre_ids:
    el = doc.GetElement(eid)
    if el and el.Category and el.Category.Id.IntegerValue == int(BuiltInCategory.OST_VolumeOfInterest):
        pre_sb.append(el)

candidates = pre_sb if pre_sb else all_scope_boxes(doc)
if not candidates:
    forms.alert("No scope boxes in this model.", exitscript=True)

class SBWrap(object):
    def __init__(self, e): self.e=e; self.n=sb_name(e)
    def __str__(self): return self.n

picked = forms.SelectFromList.show([SBWrap(e) for e in candidates],
                                   title="Select scope boxes",
                                   multiselect=True)
if not picked:
    forms.alert("Nothing selected.", exitscript=True)

sboxes = [w.e for w in picked]

# ---------- prerequisites ----------
vft3d = get_3d_vft(doc)
if not vft3d:
    forms.alert("No 3D ViewFamilyType found.", exitscript=True)

tmpl_name = "SHP_VDC_BOX"
tmpl = None
# Prefer exact 3D template match
for v in FilteredElementCollector(doc).OfClass(View):
    if v.IsTemplate and v.Name == tmpl_name:
        tmpl = v
        break
if not tmpl:
    forms.alert("View template '{}' not found. Create it first.".format(tmpl_name), exitscript=True)

nameset = existing_view_names(doc)

# ---------- create views ----------
tg = TransactionGroup(doc, "3D Views per Scope Box")
tg.Start()
t = Transaction(doc, "Create 3D views and apply template")
t.Start()

created, failed = [], []
for sb in sboxes:
    try:
        base = u"01_{}".format(sanitize_name(sb_name(sb)))
        view_name = unique_name(base, nameset)

        v = View3D.CreateIsometric(doc, vft3d.Id)

        # set section box from scope box extents
        bbox = sb.get_BoundingBox(None)  # BoundingBoxXYZ with Transform
        if isinstance(bbox, BoundingBoxXYZ):
            v.IsSectionBoxActive = True
            v.SetSectionBox(bbox)

        # bind the view to the scope box so later edits follow it
        scope_param = scope_box_parameter(v)
        if scope_param and not scope_param.IsReadOnly:
            scope_param.Set(sb.Id)
        else:
            raise Exception("Scope Box parameter unavailable or locked by template.")

        # name then apply template
        v.Name = view_name
        v.ViewTemplateId = tmpl.Id

        created.append(view_name)
    except Exception as ex:
        failed.append("{}: {}".format(sb_name(sb), str(ex)))

t.Commit()
tg.Assimilate()

# ---------- report ----------
summary = []
summary.append("Created: {}".format(len(created)))
if created:
    summary.append("Views: {}".format(", ".join(created)))
if failed:
    summary.append("Failed: {}".format("; ".join(failed)))
forms.alert("\n".join(summary))
