# -*- coding: utf-8 -*-
# pyRevit: Copy 3D view "pan" (orientation) from one view to many

from Autodesk.Revit.DB import (
    FilteredElementCollector, Transaction, View3D
)

try:
    from pyrevit import forms
    uiapp = __revit__
    uidoc = uiapp.ActiveUIDocument
    doc = uidoc.Document
except Exception:
    raise Exception("Run from pyRevit inside Revit.")

# -------- helpers --------
def all_3d_views(d):
    return [v for v in FilteredElementCollector(d).OfClass(View3D) if not v.IsTemplate]

class VWrap(object):
    def __init__(self, v):
        self.v = v
        self.n = v.Name
        self.k = "Perspective" if v.IsPerspective else "Isometric"
    def __str__(self):
        return u"{}  [{}]".format(self.n, self.k)

views = all_3d_views(doc)
if not views:
    forms.alert("No 3D views found.", exitscript=True)

# -------- pick source --------
src_pick = forms.SelectFromList.show(
    [VWrap(v) for v in views],
    title="Select SOURCE 3D view",
    multiselect=False
)
if not src_pick:
    forms.alert("No source selected.", exitscript=True)
src = src_pick.v
src_is_persp = src.IsPerspective
src_orient = src.GetOrientation()

# -------- pick targets --------
targets_pick = forms.SelectFromList.show(
    [VWrap(v) for v in views if v.Id != src.Id],
    title="Select TARGET 3D views",
    multiselect=True
)
if not targets_pick:
    forms.alert("No targets selected.", exitscript=True)
targets = [w.v for w in targets_pick]

# -------- apply orientation --------
t = Transaction(doc, "Copy 3D view pan/orientation")
t.Start()

copied, skipped = [], []
for v in targets:
    try:
        # skip type mismatch
        if v.IsPerspective != src_is_persp:
            skipped.append(u"{}: type mismatch".format(v.Name))
            continue

        # handle locked views
        was_locked = False
        try:
            was_locked = bool(getattr(v, "IsLocked", False))
        except:
            was_locked = False

        if was_locked:
            try: v.Unlock()
            except: pass

        v.SetOrientation(src_orient)

        # re-lock if needed
        if was_locked:
            locked_ok = False
            try:
                v.Lock()
                locked_ok = True
            except:
                pass
            if not locked_ok:
                try:
                    # some builds expose SaveOrientationAndLock()
                    v.SaveOrientationAndLock()
                except:
                    pass

        copied.append(v.Name)
    except Exception as ex:
        skipped.append(u"{}: {}".format(v.Name, ex))

t.Commit()

# -------- report --------
lines = []
lines.append(u"Source: {} [{}]".format(src.Name, "Perspective" if src_is_persp else "Isometric"))
lines.append(u"Copied: {}".format(len(copied)))
if copied:
    lines.append(u"Views: {}".format(", ".join(copied)))
if skipped:
    lines.append(u"Skipped: {}".format("; ".join(skipped)))
forms.alert("\n".join(lines))
