#! python
# Align+Move Wall Opening command implemented with IronPython-safe pyRevit APIs.

from collections import namedtuple

from Autodesk.Revit.DB import (  # type: ignore import
    BuiltInCategory,
    ElementId,
    ElementTransformUtils,
    LabelUtils,
    Line,
    LocationCurve,
    LocationPoint,
    RevitLinkInstance,
    SpecTypeId,
    Transaction,
    UnitTypeId,
    UnitUtils,
    XYZ,
)
from Autodesk.Revit.Exceptions import OperationCanceledException
from Autodesk.Revit.UI import Selection
from pyrevit import forms

uidoc = __revit__.ActiveUIDocument  # type: ignore[name-defined]
doc = uidoc.Document

TRANSACTION_NAME = "Align+Move Wall Opening"
SELECTION_PROMPT = "Select wall openings to align"
REFERENCE_PROMPT = "Pick reference line or edge in the active view"

ALLOWED_CATEGORY_IDS = set(
    [
        int(BuiltInCategory.OST_GenericModel),
    ]
)

EPSILON = 1e-9

TransformPlan = namedtuple(
    "TransformPlan", ["element", "align_vector", "offset_vector", "total_vector"]
)


class OpeningSelectionFilter(Selection.ISelectionFilter):
    """Allow only elements that match supported wall opening categories."""

    def AllowElement(self, element):
        return is_supported_opening(element)

    def AllowReference(self, reference, point):
        return False


class ReferenceEdgeSelectionFilter(Selection.ISelectionFilter):
    """Allow edges that yield curve geometry in the active view."""

    def AllowElement(self, element):
        return True

    def AllowReference(self, reference, point):
        return curve_from_reference(reference) is not None


class ReferenceElementSelectionFilter(Selection.ISelectionFilter):
    """Allow model/detail lines that provide a curve location."""

    def AllowElement(self, element):
        location = getattr(element, "Location", None)
        return isinstance(location, LocationCurve)

    def AllowReference(self, reference, point):
        return False


class ReferenceLinkedElementSelectionFilter(Selection.ISelectionFilter):
    """Allow selection of linked elements; validation happens after pick."""

    def AllowElement(self, element):
        return True

    def AllowReference(self, reference, point):
        return True


def normalize(vector):
    """Return a unit vector or None when magnitude is ~0."""
    if not vector:
        return None
    length = vector.GetLength()
    if length < EPSILON:
        return None
    return XYZ(vector.X / length, vector.Y / length, vector.Z / length)


def is_supported_opening(element):
    """Return True when the element belongs to an opening-related category."""
    category = element.Category
    if category is None:
        return False
    return category.Id.IntegerValue in ALLOWED_CATEGORY_IDS


def describe_element(element):
    """Return a concise label for UI messaging."""
    name = getattr(element, "Name", element.__class__.__name__)
    return "[{0}] {1}".format(element.Id.IntegerValue, name)


def validate_candidates(elements):
    """Split input elements into validated openings and skip messages."""
    valid = []
    skipped = []
    for element in elements:
        if element is None:
            continue
        if not is_supported_opening(element):
            skipped.append("{0} skipped: unsupported category.".format(describe_element(element)))
            continue
        if getattr(element, "Pinned", False):
            skipped.append("{0} skipped: element is pinned.".format(describe_element(element)))
            continue
        design_option = getattr(element, "DesignOption", None)
        if design_option is not None and design_option.Id != ElementId.InvalidElementId:
            skipped.append(
                "{0} skipped: element belongs to design option '{1}'.".format(
                    describe_element(element),
                    getattr(design_option, "Name", design_option.Id.IntegerValue),
                )
            )
            continue
        valid.append(element)
    return valid, skipped


def collect_preselected_elements():
    """Return elements from the active selection set, if any."""
    ids = list(uidoc.Selection.GetElementIds())
    elements = []
    for element_id in ids:
        element = doc.GetElement(element_id)
        if element:
            elements.append(element)
    return elements


def prompt_for_elements():
    """Prompt the user to select openings when preselection is empty or invalid."""
    try:
        references = uidoc.Selection.PickObjects(
            Selection.ObjectType.Element,
            OpeningSelectionFilter(),
            SELECTION_PROMPT,
        )
    except OperationCanceledException:
        return [], [], True

    picked_elements = []
    for reference in references:
        element = doc.GetElement(reference.ElementId)
        if element:
            picked_elements.append(element)

    valid, skipped = validate_candidates(picked_elements)
    return valid, skipped, False


def gather_openings():
    """Gather and validate openings via preselection first, then prompt."""
    preselected = collect_preselected_elements()
    valid_openings, skipped = validate_candidates(preselected)

    if valid_openings:
        return valid_openings, skipped

    prompted_valid, prompted_skipped, canceled = prompt_for_elements()
    if canceled:
        return valid_openings, skipped

    valid_openings.extend(prompted_valid)
    skipped.extend(prompted_skipped)
    return valid_openings, skipped


def show_selection_summary(valid_openings, skipped_messages):
    """Display a summary toast showing ready and skipped elements."""
    lines = ["{0} opening(s) ready for Align+Move.".format(len(valid_openings))]
    if skipped_messages:
        lines.append("")
        lines.append("Skipped:")
        lines.extend(skipped_messages)
    forms.alert("\n".join(lines))


def request_user_settings():
    """Collect offset and batch intent from the user."""
    offset_text = forms.ask_for_string(
        default="0.0",
        prompt="Enter signed offset (project units)",
        title="Align+Move Offset",
    )
    if offset_text is None:
        return None, None
    offset_text = offset_text.strip()
    try:
        offset_value = float(offset_text)
    except Exception:
        forms.alert("Offset must be a numeric value.")
        return None, None

    reuse_prompt = "Apply the same reference and offset to every selected opening?"
    reuse_choice = forms.alert(reuse_prompt, yes=True, no=True, cancel=True)
    if reuse_choice is None:
        return None, None

    return float(offset_value), bool(reuse_choice)


def pick_reference_curve():
    """Obtain a line-like curve from the user's reference selection."""
    edge_filter = ReferenceEdgeSelectionFilter()
    while True:
        try:
            reference = uidoc.Selection.PickObject(
                Selection.ObjectType.Edge,
                edge_filter,
                REFERENCE_PROMPT,
            )
        except OperationCanceledException:
            break
        except Exception:
            break
        curve = curve_from_reference(reference)
        if curve:
            return curve
        forms.alert("Reference must be a straight edge. Press ESC to pick a line instead.")

    element_filter = ReferenceElementSelectionFilter()
    while True:
        try:
            reference = uidoc.Selection.PickObject(
                Selection.ObjectType.Element,
                element_filter,
                REFERENCE_PROMPT,
            )
        except OperationCanceledException:
            break
        element = doc.GetElement(reference.ElementId)
        if element is None:
            continue
        curve = curve_from_element(element)
        if curve:
            return curve
        forms.alert("Selected element is not a straight line. Press ESC to pick from a linked model.")

    linked_filter = ReferenceLinkedElementSelectionFilter()
    while True:
        try:
            reference = uidoc.Selection.PickObject(
                Selection.ObjectType.LinkedElement,
                linked_filter,
                REFERENCE_PROMPT,
            )
        except OperationCanceledException:
            return None
        curve = curve_from_linked_reference(reference)
        if curve:
            return curve
        forms.alert("Linked element must be a straight line. Try again or press ESC to cancel.")


def curve_from_reference(reference):
    """Return the curve corresponding to an edge reference."""
    element = doc.GetElement(reference.ElementId)
    if element is None:
        return None
    geometry = element.GetGeometryObjectFromReference(reference)
    if geometry is None:
        return None
    if isinstance(element, RevitLinkInstance) and hasattr(geometry, "CreateTransformed"):
        geometry = geometry.CreateTransformed(element.GetTotalTransform())
    if isinstance(geometry, Line):
        return geometry
    return None


def curve_from_element(element, transform=None):
    """Return a curve from a line-like element."""
    location = getattr(element, "Location", None)
    if not isinstance(location, LocationCurve):
        return None
    curve = location.Curve
    if isinstance(curve, Line):
        line = curve
    else:
        start = curve.GetEndPoint(0)
        end = curve.GetEndPoint(1)
        if start.IsAlmostEqualTo(end):
            return None
        line = Line.CreateBound(start, end)
    if transform is not None and hasattr(line, "CreateTransformed"):
        return line.CreateTransformed(transform)
    return line


def curve_from_linked_reference(reference):
    """Return a curve from a linked element selection."""
    linked_id = getattr(reference, "LinkedElementId", ElementId.InvalidElementId)
    if linked_id == ElementId.InvalidElementId:
        return None
    link_instance = doc.GetElement(reference.ElementId)
    if link_instance is None or not isinstance(link_instance, RevitLinkInstance):
        return None
    link_doc = link_instance.GetLinkDocument()
    if link_doc is None:
        return None
    linked_element = link_doc.GetElement(linked_id)
    if linked_element is None:
        return None
    transform = link_instance.GetTotalTransform()
    return curve_from_element(linked_element, transform)


def get_view_context(view):
    """Return origin and orthonormal basis vectors for the active view."""
    return {
        "origin": view.Origin,
        "normal": normalize(view.ViewDirection),
        "up": normalize(view.UpDirection),
        "right": normalize(view.RightDirection),
    }


def project_point_to_plane(point, origin, normal):
    """Project a 3D point onto the plane defined by origin and normal."""
    if normal is None:
        return point
    vector = point - origin
    distance = vector.DotProduct(normal)
    return point - normal.Multiply(distance)


def project_point_to_line(point, line):
    """Project a point onto a (flattened) reference line."""
    start = line.GetEndPoint(0)
    direction = normalize(line.Direction)
    if direction is None:
        return start
    delta = point - start
    distance_along = delta.DotProduct(direction)
    return start + direction.Multiply(distance_along)


def ensure_line_in_view_plane(curve, view_context):
    """Flatten the picked curve into the active view plane and return a line."""
    if not isinstance(curve, Line):
        return None
    origin = view_context["origin"]
    normal = view_context["normal"]
    start = project_point_to_plane(curve.GetEndPoint(0), origin, normal)
    end = project_point_to_plane(curve.GetEndPoint(1), origin, normal)
    if start.IsAlmostEqualTo(end):
        return None
    short_tol = getattr(doc.Application, "ShortCurveTolerance", 1e-6)
    if (end - start).GetLength() < short_tol:
        return None
    return Line.CreateBound(start, end)


def get_length_unit_id():
    """Return the project display unit id for lengths."""
    units = doc.GetUnits()
    options = units.GetFormatOptions(SpecTypeId.Length)
    if options:
        return options.GetUnitTypeId()
    return UnitTypeId.Feet


def convert_offset_to_internal(offset_value, unit_id):
    """Convert a project unit offset to internal feet."""
    return UnitUtils.ConvertToInternalUnits(offset_value, unit_id)


def format_length(value, unit_id, signed=False):
    """Return a formatted length string using project units."""
    display_value = UnitUtils.ConvertFromInternalUnits(value, unit_id)
    label = LabelUtils.GetLabelForUnit(unit_id)
    if signed:
        return "{0:+.3f} {1}".format(display_value, label)
    return "{0:.3f} {1}".format(display_value, label)


def get_primary_point(element):
    """Determine an element point suitable for align/move."""
    location = getattr(element, "Location", None)
    if isinstance(location, LocationPoint):
        return location.Point
    if isinstance(location, LocationCurve):
        try:
            return location.Curve.Evaluate(0.5, True)
        except Exception:
            pass
    bbox = element.get_BoundingBox(doc.ActiveView) or element.get_BoundingBox(None)
    if bbox:
        return XYZ(
            (bbox.Min.X + bbox.Max.X) * 0.5,
            (bbox.Min.Y + bbox.Max.Y) * 0.5,
            (bbox.Min.Z + bbox.Max.Z) * 0.5,
        )
    return None


def compute_offset_direction(axis_mode, reference_line, view_context):
    """Return the unit vector for the perpendicular offset."""
    ref_direction = normalize(reference_line.Direction)
    if ref_direction is None:
        return None
    if axis_mode == "Vertical":
        return view_context.get("up")
    view_normal = view_context.get("normal")
    if view_normal is None:
        return None
    perpendicular = view_normal.CrossProduct(ref_direction)
    return normalize(perpendicular)


def determine_axis_mode(reference_line, view_context):
    """Infer axis mode based on the picked reference orientation in view."""
    ref_direction = normalize(reference_line.Direction)
    if ref_direction is None:
        return "Horizontal"
    up = view_context.get("up")
    right = view_context.get("right")
    vertical_alignment = abs(ref_direction.DotProduct(up)) if up else 0.0
    horizontal_alignment = abs(ref_direction.DotProduct(right)) if right else 0.0
    if vertical_alignment >= horizontal_alignment:
        return "Vertical"
    return "Horizontal"


def plan_transforms(openings, reference_line, offset_vector, view_context):
    """Pre-compute align and move vectors for each opening."""
    origin = view_context["origin"]
    normal = view_context["normal"]
    plans = []
    skipped = []
    for element in openings:
        primary_point = get_primary_point(element)
        if primary_point is None:
            skipped.append("{0} skipped: unable to determine reference point.".format(describe_element(element)))
            continue
        projected_point = project_point_to_plane(primary_point, origin, normal)
        target_point = project_point_to_line(projected_point, reference_line)
        align_vector = target_point - projected_point
        total_vector = align_vector + offset_vector
        plans.append(TransformPlan(element, align_vector, offset_vector, total_vector))
    return plans, skipped


def build_preview(plans, axis_mode, offset_internal, unit_id):
    """Generate preview text describing upcoming operations."""
    lines = [
        "Align+Move Preview",
        "Openings ready: {0}".format(len(plans)),
        "Axis priority: {0}".format(axis_mode),
        "Offset: {0}".format(format_length(offset_internal, unit_id, signed=True)),
        "",
    ]
    for plan in plans[:5]:
        movement = plan.total_vector.GetLength()
        lines.append(
            "- {0}: move {1}".format(
                describe_element(plan.element),
                format_length(movement, unit_id),
            )
        )
    if len(plans) > 5:
        lines.append("... {0} more opening(s)".format(len(plans) - 5))
    lines.append("")
    lines.append("Proceed with Align+Move transaction?")
    return "\n".join(lines)


def apply_transforms(plans, unit_id):
    """Execute the prepared transforms inside a single transaction."""
    if not plans:
        return [], []

    transaction = Transaction(doc, TRANSACTION_NAME)
    transaction.Start()
    applied = []
    failures = []
    try:
        for plan in plans:
            element = plan.element
            vector = plan.total_vector
            try:
                if vector.GetLength() < EPSILON:
                    applied.append("{0}: already aligned.".format(describe_element(element)))
                    continue
                ElementTransformUtils.MoveElement(doc, element.Id, vector)
                applied.append(
                    "{0}: moved {1}".format(
                        describe_element(element),
                        format_length(vector.GetLength(), unit_id),
                    )
                )
            except Exception as exc:
                failures.append("{0}: failed ({1})".format(describe_element(element), exc))
        if applied:
            transaction.Commit()
        else:
            transaction.RollBack()
    except Exception:
        transaction.RollBack()
        raise
    return applied, failures


def summarize_results(applied, failures, skipped):
    """Show a transaction summary dialog."""
    lines = []
    lines.append("{0} opening(s) updated.".format(len(applied)))
    if applied:
        lines.append("")
        lines.append("Updated:")
        lines.extend(applied[:5])
        if len(applied) > 5:
            lines.append("... {0} more update(s)".format(len(applied) - 5))
    if failures:
        lines.append("")
        lines.append("Failed:")
        lines.extend(failures)
    if skipped:
        lines.append("")
        lines.append("Skipped:")
        lines.extend(skipped)
    forms.alert("\n".join(lines))


def main():
    openings, skipped = gather_openings()
    if not openings:
        message = ["No modifiable wall openings were selected."]
        if skipped:
            message.append("")
            message.append("Skipped:")
            message.extend(skipped)
        forms.alert("\n".join(message), exitscript=True)
        return

    show_selection_summary(openings, skipped)

    offset_value, batch_apply = request_user_settings()
    if offset_value is None:
        forms.alert("Command canceled at input stage.", exitscript=True)
        return
    if not batch_apply:
        forms.alert("Batch toggle disabled, no changes applied.", exitscript=True)
        return

    curve = pick_reference_curve()
    if curve is None:
        forms.alert("Command canceled before reference selection.", exitscript=True)
        return

    view_context = get_view_context(doc.ActiveView)
    reference_line = ensure_line_in_view_plane(curve, view_context)
    if reference_line is None:
        forms.alert("Reference must be a straight line visible in the active view.", exitscript=True)
        return

    axis_mode = determine_axis_mode(reference_line, view_context)
    unit_id = get_length_unit_id()
    offset_internal = convert_offset_to_internal(offset_value, unit_id)
    offset_direction = compute_offset_direction(axis_mode, reference_line, view_context)
    if offset_direction is None:
        forms.alert("Unable to derive offset direction for the selected axis mode.", exitscript=True)
        return
    offset_vector = offset_direction.Multiply(offset_internal)

    plans, geom_skips = plan_transforms(openings, reference_line, offset_vector, view_context)
    skipped.extend(geom_skips)
    if not plans:
        message = ["No openings could be prepared for Align+Move."]
        if skipped:
            message.append("")
            message.append("Skipped:")
            message.extend(skipped)
        forms.alert("\n".join(message), exitscript=True)
        return

    preview = build_preview(plans, axis_mode, offset_internal, unit_id)
    proceed = forms.alert(preview, yes=True, no=True, cancel=False)
    if not proceed:
        forms.alert("Align+Move canceled before applying transforms.", exitscript=True)
        return

    applied, failures = apply_transforms(plans, unit_id)
    summarize_results(applied, failures, skipped)


if __name__ == "__main__":
    main()
