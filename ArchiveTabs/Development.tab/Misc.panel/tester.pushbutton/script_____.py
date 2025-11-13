# -*- coding: utf-8 -*-
"""Copy section boxes from the current model to another open model."""

from Autodesk.Revit.DB import (
    BoundingBoxXYZ,
    FilteredElementCollector,
    Transaction,
    Transform,
    View3D,
    XYZ,
)

try:
    from pyrevit import forms
except Exception:
    raise Exception("pyRevit is required. Run this from pyRevit inside Revit.")

uiapp = __revit__  # type: ignore # noqa
uidoc = uiapp.ActiveUIDocument
if not uidoc:
    raise Exception("No active document found.")

src_doc = uidoc.Document
app = uiapp.Application


# -------- helpers --------
def get_3d_views(doc):
    return sorted(
        [v for v in FilteredElementCollector(doc).OfClass(View3D) if not v.IsTemplate],
        key=lambda v: v.Name.lower(),
    )


def clone_transform(src_transform):
    new_t = Transform.Identity
    new_t.BasisX = XYZ(
        src_transform.BasisX.X, src_transform.BasisX.Y, src_transform.BasisX.Z
    )
    new_t.BasisY = XYZ(
        src_transform.BasisY.X, src_transform.BasisY.Y, src_transform.BasisY.Z
    )
    new_t.BasisZ = XYZ(
        src_transform.BasisZ.X, src_transform.BasisZ.Y, src_transform.BasisZ.Z
    )
    new_t.Origin = XYZ(
        src_transform.Origin.X, src_transform.Origin.Y, src_transform.Origin.Z
    )
    return new_t


def clone_box(src_box):
    if not src_box:
        return None
    new_box = BoundingBoxXYZ()
    new_box.Enabled = src_box.Enabled
    new_box.Min = XYZ(src_box.Min.X, src_box.Min.Y, src_box.Min.Z)
    new_box.Max = XYZ(src_box.Max.X, src_box.Max.Y, src_box.Max.Z)
    new_box.Transform = clone_transform(src_box.Transform)
    return new_box


class ViewChoice(object):
    def __init__(self, view):
        self.view = view
        kind = "Perspective" if view.IsPerspective else "Isometric"
        section_state = (
            "Section Box ON" if getattr(view, "IsSectionBoxActive", False) else "OFF"
        )
        self.label = u"{0} [{1} | {2}]".format(view.Name, kind, section_state)

    def __str__(self):
        return self.label


class DocChoice(object):
    def __init__(self, doc):
        self.doc = doc
        path = doc.PathName or "Unsaved"
        self.label = u"{0} ({1})".format(doc.Title, path)

    def __str__(self):
        return self.label


# -------- source selection --------
source_views = get_3d_views(src_doc)
if not source_views:
    forms.alert("No 3D views found in the active model.", exitscript=True)

picked_views = forms.SelectFromList.show(
    [ViewChoice(v) for v in source_views],
    title="Select 3D views to copy section boxes from",
    multiselect=True,
)
if not picked_views:
    forms.alert("No views selected.", exitscript=True)
selected_views = [wrap.view for wrap in picked_views]


# -------- target document --------
other_docs = [DocChoice(doc) for doc in app.Documents if doc != src_doc]
if not other_docs:
    forms.alert("Open the target model before running this tool.", exitscript=True)

target_choice = forms.SelectFromList.show(
    other_docs,
    title="Select TARGET model",
    multiselect=False,
)
if not target_choice:
    forms.alert("No target model selected.", exitscript=True)
target_doc = target_choice.doc


# -------- gather data --------
source_payload = []
for view in selected_views:
    try:
        section_box = view.GetSectionBox()
    except Exception:
        section_box = None
    payload = {
        "name": view.Name,
        "perspective": view.IsPerspective,
        "section_on": getattr(view, "IsSectionBoxActive", False),
        "box": clone_box(section_box),
    }
    source_payload.append(payload)

target_view_map = dict((v.Name, v) for v in get_3d_views(target_doc))


# -------- copy section boxes --------
tx_name = u"Copy Section Boxes from {0}".format(src_doc.Title)
tx = Transaction(target_doc, tx_name)
tx.Start()

copied = []
skipped = []

for data in source_payload:
    t_view = target_view_map.get(data["name"])
    if not t_view:
        skipped.append(u"{0}: no matching view".format(data["name"]))
        continue
    if t_view.IsPerspective != data["perspective"]:
        skipped.append(u"{0}: type mismatch".format(data["name"]))
        continue

    try:
        t_view.IsSectionBoxActive = data["section_on"]
        if data["box"]:
            t_view.SetSectionBox(clone_box(data["box"]))
        copied.append(t_view.Name)
    except Exception as ex:
        skipped.append(u"{0}: {1}".format(t_view.Name, ex))

tx.Commit()


# -------- report --------
lines = [
    u"Source model: {0}".format(src_doc.Title),
    u"Target model: {0}".format(target_doc.Title),
    u"Copied section boxes: {0}".format(len(copied)),
]
if copied:
    lines.append(u"Views updated: {0}".format(", ".join(copied)))
if skipped:
    lines.append(u"Skipped: {0}".format("; ".join(skipped)))

forms.alert("\n".join(lines))
