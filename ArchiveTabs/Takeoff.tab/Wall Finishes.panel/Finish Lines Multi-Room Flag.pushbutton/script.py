#! python
# Finish Lines Multi-Room Flag: detect detail lines crossing multiple rooms and export CSV.

import clr
clr.AddReference('RevitAPI')
clr.AddReference('System')

from Autodesk.Revit.DB import (
    BuiltInCategory,
    BuiltInParameter,
    CurveElement,
    DetailCurve,
    ElementId,
    FilteredElementCollector,
    IntersectionResultArray,
    SetComparisonResult,
    SpatialElementBoundaryLocation,
    SpatialElementBoundaryOptions,
    UnitTypeId,
    UnitUtils,
    ViewPlan,
    ViewType,
    XYZ,
)
from pyrevit import forms
from System.IO import StreamWriter
from System.Text import UTF8Encoding

doc = __revit__.ActiveUIDocument.Document  # type: ignore[name-defined]


class ViewWrap(object):
    def __init__(self, view, label):
        self.view = view
        self.label = label

    def __str__(self):
        return self.label


class RoomInfo(object):
    def __init__(self, room, name, number, height_internal, level_elev, base_offset, boundary_curves):
        self.room = room
        self.room_id = room.Id
        self.name = name
        self.number = number
        self.height_internal = height_internal
        self.level_elev = level_elev
        self.base_offset = base_offset
        self.boundary_curves = boundary_curves

    def inside_z(self):
        if self.height_internal > 0:
            return self.level_elev + self.base_offset + (self.height_internal * 0.5)
        return self.level_elev + self.base_offset


def is_floor_plan(view):
    return isinstance(view, ViewPlan) and view.ViewType == ViewType.FloorPlan and not view.IsTemplate


def pick_floor_plan(doc):
    active_view = doc.ActiveView
    plans = [
        v for v in FilteredElementCollector(doc).OfClass(ViewPlan)
        if is_floor_plan(v)
    ]
    if not plans:
        forms.alert("No floor plan views found.", exitscript=True)

    wrapped = []
    if is_floor_plan(active_view):
        wrapped.append(ViewWrap(active_view, "{} (active)".format(active_view.Name)))
        for v in sorted(plans, key=lambda x: x.Name):
            if v.Id != active_view.Id:
                wrapped.append(ViewWrap(v, v.Name))
    else:
        for v in sorted(plans, key=lambda x: x.Name):
            wrapped.append(ViewWrap(v, v.Name))

    picked = forms.SelectFromList.show(
        wrapped,
        title="Select floor plan view",
        multiselect=False,
    )
    if not picked:
        forms.alert("No view selected.", exitscript=True)
    return picked.view


def collect_detail_lines(doc, view):
    return list(
        el for el in (
            FilteredElementCollector(doc, view.Id)
            .OfClass(CurveElement)
            .WhereElementIsNotElementType()
        )
        if isinstance(el, DetailCurve)
    )


def get_bip(name):
    try:
        return getattr(BuiltInParameter, name)
    except Exception:
        return None


def get_double_param(element, bip, default=None):
    if bip is None:
        return default
    param = element.get_Parameter(bip)
    if param and param.HasValue:
        try:
            return param.AsDouble()
        except Exception:
            return default
    return default


def get_double_param_by_names(element, names, default=None):
    for name in names:
        if not name:
            continue
        param = element.LookupParameter(name)
        if param and param.HasValue:
            try:
                return param.AsDouble()
            except Exception:
                continue
    return default


def get_room_height_internal(room):
    base_offset = get_double_param(
        room,
        get_bip("ROOM_BASE_OFFSET") or get_bip("ROOM_LOWER_OFFSET"),
        None,
    )
    if base_offset is None:
        base_offset = get_double_param_by_names(room, ["Base Offset", "Lower Offset"], 0.0)

    upper_offset = get_double_param(
        room,
        get_bip("ROOM_UPPER_OFFSET") or get_bip("ROOM_LIMIT_OFFSET"),
        None,
    )
    if upper_offset is None:
        upper_offset = get_double_param_by_names(room, ["Limit Offset", "Upper Offset"], None)

    height = None
    if upper_offset is not None:
        height = upper_offset - base_offset
    if height is None or height <= 0:
        height = get_double_param(room, get_bip("ROOM_HEIGHT"), None)
    if height is None or height <= 0:
        try:
            height = room.UnboundedHeight
        except Exception:
            height = 0.0
    return height, base_offset


def get_room_name(room):
    bip = get_bip("ROOM_NAME")
    name_param = room.get_Parameter(bip) if bip else None
    if name_param and name_param.HasValue:
        return name_param.AsString()
    name_param = room.LookupParameter("Name") or room.LookupParameter("Room Name")
    return name_param.AsString() if name_param else ""


def get_room_number(room):
    bip = get_bip("ROOM_NUMBER")
    number_param = room.get_Parameter(bip) if bip else None
    if number_param and number_param.HasValue:
        return number_param.AsString()
    number_param = room.LookupParameter("Number") or room.LookupParameter("Room Number")
    return number_param.AsString() if number_param else ""


def get_room_boundaries(room):
    options = SpatialElementBoundaryOptions()
    options.SpatialElementBoundaryLocation = SpatialElementBoundaryLocation.Finish
    curves = []
    try:
        segments = room.GetBoundarySegments(options)
    except Exception:
        return curves
    if not segments:
        return curves
    for loop in segments:
        for seg in loop:
            curve = seg.GetCurve()
            if curve:
                curves.append(curve)
    return curves


def collect_rooms_on_level(doc, level_id):
    rooms = []
    for room in (
        FilteredElementCollector(doc)
        .OfCategory(BuiltInCategory.OST_Rooms)
        .WhereElementIsNotElementType()
    ):
        try:
            if room.LevelId != level_id:
                continue
            if hasattr(room, "Area") and room.Area <= 0:
                continue
            rooms.append(room)
        except Exception:
            continue
    return rooms


def curve_from_line(line):
    try:
        curve = line.GeometryCurve
        if curve:
            return curve
    except Exception:
        pass
    try:
        return line.Curve
    except Exception:
        return None


def point_in_room(room, x, y, z):
    try:
        return room.IsPointInRoom(XYZ(x, y, z))
    except Exception:
        return False


def normalized_param(curve, param):
    start_param = curve.GetEndParameter(0)
    end_param = curve.GetEndParameter(1)
    if end_param == start_param:
        return 0.0
    return (param - start_param) / (end_param - start_param)


def is_intersection(comp):
    return comp in (
        SetComparisonResult.Overlap,
        SetComparisonResult.Subset,
        SetComparisonResult.Superset,
        SetComparisonResult.Equal,
    )


def intersection_params(line_curve, room_infos):
    params = set()
    for info in room_infos:
        for boundary in info.boundary_curves:
            try:
                results_ref = clr.Reference[IntersectionResultArray]()
                comp = line_curve.Intersect(boundary, results_ref)
                if not is_intersection(comp):
                    continue
                results = results_ref.Value
                if not results:
                    continue
                for result in results:
                    pt = result.XYZPoint
                    if not pt:
                        continue
                    proj = line_curve.Project(pt)
                    if not proj:
                        continue
                    param = normalized_param(line_curve, proj.Parameter)
                    if -1e-6 <= param <= 1.0 + 1e-6:
                        if param < 0.0:
                            param = 0.0
                        if param > 1.0:
                            param = 1.0
                        params.add(param)
            except Exception:
                continue
    return params


def find_room_for_point(room_infos, x, y):
    for info in room_infos:
        if point_in_room(info.room, x, y, info.inside_z()):
            return info
    return None


def to_cm(value_internal):
    return UnitUtils.ConvertFromInternalUnits(value_internal, UnitTypeId.Centimeters)


def format_num(value):
    if value is None:
        return ""
    try:
        return "{0:.2f}".format(value)
    except Exception:
        return str(value)


def csv_escape(value):
    if value is None:
        return ""
    text = str(value)
    if any(ch in text for ch in [",", "\"", "\n", "\r"]):
        text = text.replace("\"", "\"\"")
        return "\"{}\"".format(text)
    return text


def write_csv(path, headers, rows):
    encoding = UTF8Encoding(True)
    writer = StreamWriter(path, False, encoding)
    try:
        writer.WriteLine(",".join(headers))
        for row in rows:
            writer.WriteLine(",".join(csv_escape(val) for val in row))
    finally:
        writer.Close()


def main():
    view = pick_floor_plan(doc)
    view_level = getattr(view, "GenLevel", None)
    if not view_level and hasattr(view, "LevelId") and view.LevelId != ElementId.InvalidElementId:
        view_level = doc.GetElement(view.LevelId)
    if not view_level:
        forms.alert("Selected view has no associated level.", exitscript=True)

    lines = collect_detail_lines(doc, view)
    if not lines:
        forms.alert("No detail lines found in the selected view.", exitscript=True)

    rooms = collect_rooms_on_level(doc, view_level.Id)
    if not rooms:
        forms.alert("No rooms found on the view's level.", exitscript=True)

    room_infos = []
    for room in rooms:
        height_internal, base_offset = get_room_height_internal(room)
        level = doc.GetElement(room.LevelId)
        level_elev = level.Elevation if level else 0.0
        boundaries = get_room_boundaries(room)
        if not boundaries:
            continue
        room_infos.append(
            RoomInfo(
                room,
                get_room_name(room),
                get_room_number(room),
                height_internal,
                level_elev,
                base_offset,
                boundaries,
            )
        )

    if not room_infos:
        forms.alert("Rooms on this level are missing boundaries.", exitscript=True)

    rows = []
    tol = 1e-6

    for line in lines:
        curve = curve_from_line(line)
        if not curve:
            continue

        params = intersection_params(curve, room_infos)
        params.add(0.0)
        params.add(1.0)

        sorted_params = sorted(params)
        unique_params = []
        for param in sorted_params:
            if not unique_params or abs(param - unique_params[-1]) > tol:
                unique_params.append(param)

        found = {}
        if len(unique_params) >= 2:
            for idx in range(len(unique_params) - 1):
                a = unique_params[idx]
                b = unique_params[idx + 1]
                if (b - a) <= tol:
                    continue
                mid = (a + b) * 0.5
                try:
                    pt = curve.Evaluate(mid, True)
                except Exception:
                    continue
                info = find_room_for_point(room_infos, pt.X, pt.Y)
                if info:
                    found[info.room_id.IntegerValue] = info

        if not found:
            for t in [0.0, 0.5, 1.0]:
                try:
                    pt = curve.Evaluate(t, True)
                except Exception:
                    continue
                info = find_room_for_point(room_infos, pt.X, pt.Y)
                if info:
                    found[info.room_id.IntegerValue] = info

        rooms_found = sorted(found.values(), key=lambda r: (r.number or "", r.name or ""))
        room_count = len(rooms_found)
        room_ids = ";".join(str(r.room_id.IntegerValue) for r in rooms_found)
        room_numbers = ";".join((r.number or "") for r in rooms_found)
        room_names = ";".join((r.name or "") for r in rooms_found)
        multi_flag = "Yes" if room_count > 1 else "No"

        line_length_cm = to_cm(curve.Length)
        line_style = getattr(line, "LineStyle", None)
        line_type_name = ""
        if line_style:
            line_type_name = getattr(line_style, "Name", "") or ""
        if not line_type_name:
            try:
                line_type_name = doc.GetElement(line.GetTypeId()).Name
            except Exception:
                line_type_name = ""

        rows.append(
            [
                line.Id.ToString(),
                line_type_name,
                format_num(line_length_cm),
                str(room_count),
                room_ids,
                room_numbers,
                room_names,
                multi_flag,
            ]
        )

    if not rows:
        forms.alert("No valid detail lines were processed.", exitscript=True)

    default_name = "finish_lines_multi_room_flag.csv"
    save_path = forms.save_file(file_ext="csv", default_name=default_name)
    if not save_path:
        forms.alert("Export canceled.", exitscript=True)

    headers = [
        "LineId",
        "LineType",
        "LineLengthCm",
        "RoomCount",
        "RoomIds",
        "RoomNumbers",
        "RoomNames",
        "MultiRoomFlag",
    ]
    write_csv(save_path, headers, rows)

    forms.alert(
        "Exported {} lines to:\n{}".format(len(rows), save_path),
        title="Finish Lines Multi-Room Flag",
    )


if __name__ == "__main__":
    main()
