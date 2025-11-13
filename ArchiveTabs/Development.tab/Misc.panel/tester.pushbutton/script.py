# -*- coding: utf-8 -*-
# pyRevit: Crop previously created 3D views back to their scope boxes

from Autodesk.Revit.DB import (
    FilteredElementCollector, BuiltInParameter,
    Transaction,
    View3D, ViewType,
    BoundingBoxXYZ, XYZ, Transform
)

try:
    from pyrevit import forms
    uiapp = __revit__
    uidoc = uiapp.ActiveUIDocument
    doc = uidoc.Document
except Exception:
    raise Exception("Run from pyRevit inside Revit.")


# ---------- helpers ----------
def candidate_views(d):
    """Return all non-template 3D views."""
    return [
        v for v in FilteredElementCollector(d).OfClass(View3D)
        if isinstance(v, View3D) and not v.IsTemplate and v.ViewType == ViewType.ThreeD
    ]


def selected_views(d):
    views = []
    for eid in uidoc.Selection.GetElementIds():
        v = d.GetElement(eid)
        if isinstance(v, View3D) and not v.IsTemplate:
            views.append(v)
    return views


def scope_box_parameter(view):
    """Return the parameter controlling the view's scope box, if any."""
    scope_param = None
    bip = getattr(BuiltInParameter, "VIEWER_VOLUME_OF_INTEREST", None)
    if bip is not None:
        try:
            scope_param = view.get_Parameter(bip)
        except AttributeError:
            scope_param = None
    if not scope_param:
        scope_param = view.LookupParameter("Scope Box")
    return scope_param


def get_scope_box(view):
    param = scope_box_parameter(view)
    if not param:
        return None
    sb_id = param.AsElementId()
    if sb_id and sb_id.IntegerValue > 0:
        return doc.GetElement(sb_id)
    return None


def bbox_corners(bbox):
    if not bbox:
        return []
    xs = (bbox.Min.X, bbox.Max.X)
    ys = (bbox.Min.Y, bbox.Max.Y)
    zs = (bbox.Min.Z, bbox.Max.Z)
    tf = getattr(bbox, "Transform", None)
    corners = []
    for x in xs:
        for y in ys:
            for z in zs:
                pt = XYZ(x, y, z)
                if tf:
                    pt = tf.OfPoint(pt)
                corners.append(pt)
    return corners


def bbox_center_world(bbox):
    corners = bbox_corners(bbox)
    if not corners:
        return None
    total = XYZ(0, 0, 0)
    for corner in corners:
        total = total.Add(corner)
    return total.Divide(len(corners))


def transform_bbox_to_view(bbox, view_tf):
    """Project the bbox into the local coordinates of view_tf."""
    if not bbox or not view_tf:
        return None
    corners = bbox_corners(bbox)
    if not corners:
        return None
    inv = view_tf.Inverse
    min_x = float("inf")
    min_y = float("inf")
    min_z = float("inf")
    max_x = float("-inf")
    max_y = float("-inf")
    max_z = float("-inf")
    for corner in corners:
        local = inv.OfPoint(corner)
        min_x = min(min_x, local.X)
        min_y = min(min_y, local.Y)
        min_z = min(min_z, local.Z)
        max_x = max(max_x, local.X)
        max_y = max(max_y, local.Y)
        max_z = max(max_z, local.Z)
    crop = BoundingBoxXYZ()
    crop.Transform = view_tf
    crop.Min = XYZ(min_x, min_y, min_z)
    crop.Max = XYZ(max_x, max_y, max_z)
    return crop


def view_crop_transform(view, scope_bbox=None, section_box=None):
    """Determine the transform that aligns with the current view crop."""
    try:
        crop = view.CropBox
        if crop and crop.Transform:
            return crop.Transform
    except Exception:
        pass

    origin = None
    if section_box and section_box.Transform:
        origin = section_box.Transform.Origin

    if not origin and scope_bbox:
        origin = bbox_center_world(scope_bbox)

    orient = None
    try:
        orient = view.GetOrientation()
    except Exception:
        orient = None

    tf = Transform.Identity
    if orient:
        forward = orient.ForwardDirection.Normalize()
        up = orient.UpDirection.Normalize()
        right = up.CrossProduct(forward).Normalize()
        tf.BasisX = right
        tf.BasisY = up
        tf.BasisZ = forward

    if origin:
        tf.Origin = origin
    return tf


class ViewWrap(object):
    def __init__(self, view):
        self.view = view
        self.name = view.Name

    def __str__(self):
        return self.name


# ---------- collect targets ----------
prepicked = selected_views(doc)
views_to_process = prepicked if prepicked else candidate_views(doc)
if not views_to_process:
    forms.alert("No 3D views available.", exitscript=True)

pick = forms.SelectFromList.show(
    [ViewWrap(v) for v in views_to_process],
    title="Select 3D views to crop",
    multiselect=True
)
if not pick:
    forms.alert("Nothing selected.", exitscript=True)

views = [wrap.view for wrap in pick]


# ---------- crop views ----------
t = Transaction(doc, "Apply scope-based crop to 3D views")
t.Start()
cropped, failed = [], []
for view in views:
    try:
        scope = get_scope_box(view)

        bbox = None
        if scope:
            bbox = scope.get_BoundingBox(None)
        if not bbox and isinstance(view, View3D):
            try:
                bbox = view.GetSectionBox()
            except Exception:
                bbox = None

        if not bbox:
            raise Exception("No scope or section box available.")

        section_box = None
        try:
            section_box = view.GetSectionBox()
        except Exception:
            section_box = None

        crop_tf = view_crop_transform(view, scope_bbox=bbox, section_box=section_box)
        crop_bbox = transform_bbox_to_view(bbox, crop_tf)
        if not crop_bbox:
            raise Exception("Bounding box could not be transformed into view space.")

        view.CropBox = crop_bbox
        view.CropBoxActive = True
        view.CropBoxVisible = False
        cropped.append(view.Name)
    except Exception as ex:
        failed.append("{}: {}".format(view.Name, ex))

t.Commit()


# ---------- report ----------
lines = []
lines.append("Cropped views: {}".format(len(cropped)))
if cropped:
    lines.append(", ".join(cropped))
if failed:
    lines.append("Failed: {}".format("; ".join(failed)))
forms.alert("\n".join(lines))
