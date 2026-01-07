#! python
# Finish Lines Rooms Unified: multi-view, per-room segments, single-CSV export.

import clr
clr.AddReference('RevitAPI')
clr.AddReference('System')

import os
import datetime

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


def sanitize_filename(name):
    bad = '<>:"/\\|?*'
    cleaned = name
    for ch in bad:
        cleaned = cleaned.replace(ch, "_")
    cleaned = cleaned.strip() or "View"
    return cleaned


def unique_filename(base, existing):
    name = base
    idx = 2
    key = name.lower()
    while key in existing:
        name = "{} ({})".format(base, idx)
        key = name.lower()
        idx += 1
    existing.add(key)
    return name


def pick_floor_plans(doc):
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
        title="Select floor plan views",
        multiselect=True,
    )
    if not picked:
        forms.alert("No views selected.", exitscript=True)
    return [item.view for item in picked]


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


def sample_points_xy(curve, params):
    points = []
    for t in params:
        try:
            pt = curve.Evaluate(t, True)
            points.append((pt.X, pt.Y))
        except Exception:
            continue
    return points


def distance_point_to_curve(point, curve):
    try:
        result = curve.Project(point)
        if result:
            return point.DistanceTo(result.XYZPoint)
    except Exception:
        return None
    return None


def min_distance_to_boundaries(sample_xy, room_info):
    if not room_info.boundary_curves:
        return None
    min_dist = None
    for curve in room_info.boundary_curves:
        for x, y in sample_xy:
            pt = XYZ(x, y, room_info.level_elev)
            dist = distance_point_to_curve(pt, curve)
            if dist is None:
                continue
            if min_dist is None or dist < min_dist:
                min_dist = dist
    return min_dist


def nearest_room_by_distance(room_infos, sample_xy):
    best_info = None
    best_dist = None
    for info in room_infos:
        dist = min_distance_to_boundaries(sample_xy, info)
        if dist is None:
            continue
        if best_dist is None or dist < best_dist:
            best_dist = dist
            best_info = info
    return best_info, best_dist


def segment_length_internal(curve, start_norm, end_norm):
    if end_norm <= start_norm:
        return 0.0
    start_param = curve.GetEndParameter(0)
    end_param = curve.GetEndParameter(1)
    actual_start = start_param + (end_param - start_param) * start_norm
    actual_end = start_param + (end_param - start_param) * end_norm
    if actual_start == actual_end:
        return 0.0
    try:
        seg = curve.Clone()
        seg.MakeBound(min(actual_start, actual_end), max(actual_start, actual_end))
        return seg.Length
    except Exception:
        return curve.Length * (end_norm - start_norm)


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


def sort_rows(rows):
    def sort_key(row):
        line_type = (row[0] or "").lower()
        floor_name = (row[1] or "").lower()
        room_name = (row[2] or "").lower()
        try:
            wall_area = float(row[3])
        except Exception:
            wall_area = 0.0
        return (line_type, floor_name, room_name, wall_area)
    return sorted(rows, key=sort_key)


def build_rows_for_view(view):
    view_level = getattr(view, "GenLevel", None)
    if not view_level and hasattr(view, "LevelId") and view.LevelId != ElementId.InvalidElementId:
        view_level = doc.GetElement(view.LevelId)
    if not view_level:
        return None, "View '{}' has no associated level.".format(view.Name)

    lines = collect_detail_lines(doc, view)
    if not lines:
        return [], "No detail lines found in view '{}'.".format(view.Name)

    rooms = collect_rooms_on_level(doc, view_level.Id)
    if not rooms:
        return [], "No rooms found on level for view '{}'.".format(view.Name)

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
        return [], "Rooms in view '{}' are missing boundaries.".format(view.Name)

    room_info_by_id = {info.room_id.IntegerValue: info for info in room_infos}
    rows = []
    tol = 1e-6
    distance_params = [0.0, 0.25, 0.5, 0.75, 1.0]
    match_id_counter = 1
    floor_name = view.Name or ""

    for line in lines:
        curve = curve_from_line(line)
        if not curve:
            continue

        original_length_internal = curve.Length
        original_length_cm = to_cm(original_length_internal)

        params = intersection_params(curve, room_infos)
        params.add(0.0)
        params.add(1.0)
        sorted_params = sorted(params)
        unique_params = []
        for param in sorted_params:
            if not unique_params or abs(param - unique_params[-1]) > tol:
                unique_params.append(param)

        room_lengths = {}
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
                if not info:
                    continue
                seg_len = segment_length_internal(curve, a, b)
                if seg_len <= tol:
                    continue
                rid = info.room_id.IntegerValue
                room_lengths[rid] = room_lengths.get(rid, 0.0) + seg_len

        distance_internal = 0.0 if room_lengths else None
        if not room_lengths:
            distance_xy = sample_points_xy(curve, distance_params)
            nearest_room, distance_internal = nearest_room_by_distance(room_infos, distance_xy) if distance_xy else (None, None)
            if nearest_room:
                room_lengths[nearest_room.room_id.IntegerValue] = original_length_internal

        if not room_lengths:
            continue

        line_style = getattr(line, "LineStyle", None)
        line_type_name = ""
        if line_style:
            line_type_name = getattr(line_style, "Name", "") or ""
        if not line_type_name:
            try:
                line_type_name = doc.GetElement(line.GetTypeId()).Name
            except Exception:
                line_type_name = ""

        room_ids = sorted(room_lengths.keys(), key=lambda rid: (
            room_info_by_id[rid].number or "",
            room_info_by_id[rid].name or "",
        ))

        match_id = 0
        if len(room_ids) > 1:
            match_id = match_id_counter
            match_id_counter += 1

        for rid in room_ids:
            info = room_info_by_id[rid]
            segment_length_cm = to_cm(room_lengths[rid])
            room_height_cm = to_cm(info.height_internal)
            wall_area_m2 = (segment_length_cm * room_height_cm) / 10000.0
            rows.append(
                [
                    line_type_name,
                    floor_name,
                    info.name or "",
                    format_num(wall_area_m2),
                    line.Id.ToString(),
                    format_num(segment_length_cm),
                    format_num(original_length_cm),
                    str(info.room_id.IntegerValue),
                    info.number or "",
                    format_num(room_height_cm),
                    str(match_id),
                    format_num(to_cm(distance_internal)) if distance_internal is not None else "",
                ]
            )

    return rows, None


def main():
    views = pick_floor_plans(doc)
    headers = [
        "LineType",
        "FloorName",
        "RoomName",
        "WallAreaM^2",
        "LineId",
        "LineLengthCm",
        "OriginalLineLengthCm",
        "RoomId",
        "RoomNumber",
        "RoomHeightCm",
        "MultiRoomFlag",
        "NearestBoundaryDistanceCm",
    ]

    output_dir = forms.pick_folder(title="Select output folder for CSV exports")
    if not output_dir:
        forms.alert("Export canceled.", exitscript=True)

    messages = []
    all_rows = []

    for view in views:
        rows, warning = build_rows_for_view(view)
        if warning:
            messages.append(warning)
        if rows is None:
            continue
        all_rows.extend(rows)

    if not all_rows:
        forms.alert("No data to export.\n{}".format("\n".join(messages)), exitscript=True)

    model_name = sanitize_filename(doc.Title or "Model")
    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    base_name = "{}_{}".format(model_name, timestamp)
    file_name = sanitize_filename(base_name)
    path = os.path.join(output_dir, "{}.csv".format(file_name))
    write_csv(path, headers, sort_rows(all_rows))

    summary = ["Exported 1 CSV file to:", path]
    if messages:
        summary.append("")
        summary.append("Notes:")
        summary.extend(messages)
    forms.alert("\n".join(summary), title="Finish Lines Rooms Unified")


if __name__ == "__main__":
    main()
