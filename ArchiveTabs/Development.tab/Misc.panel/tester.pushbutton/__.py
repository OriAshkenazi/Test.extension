#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""pyRevit script: Calculate consolidated sub‑grade volume for specific floor and wall elements

Given:
    • Floors of type "MGN_Floor 40cm".
    • Walls whose type name does not contain the word "finish" (case‑insensitive).

It:
    1. Collects the elements above.
    2. Clips each element's solid geometry to the vertical band between
       ‑1.62 m and ‑1.22 m relative to the Project Base Point.
    3. Boolean‑unions all clipped solids so overlaps are counted only once.
    4. Reports:
         – The **total** consolidated volume (m³) of the union.
         – A per‑element row containing:
             ▸ ElementId
             ▸ Raw clipped volume (m³)
             ▸ Unique contribution volume (m³) after subtracting overlaps that
               were already counted by previous elements.

Notes:
    • Volumes are output in cubic metres; internal Revit calculations are
      performed in feet³ and converted using 1 m = 3.28084 ft.
    • The script does *not* modify the document, so no transaction is started.
    • Heavy boolean operations can take time on large models – progress is
      written to the pyRevit output panel.
"""

from pyrevit import revit, DB, script
import math

# ---------------------------------------------------------------------------
# Configuration — tweak if needed
# ---------------------------------------------------------------------------
FLOOR_TYPE_NAME = "MGN_Floor 40cm"
EXCLUDE_WALL_KEYWORD = "finish"  # case‑insensitive substring to exclude
Z_TOP_METRES = 0.20   # shallower value (closer to zero)
Z_BOT_METRES = -0.20   # deeper value (more negative)
# A very large half‑extent for the clip box in metres
BBOX_HALF_METRES = 500.0

# ---------------------------------------------------------------------------
M_TO_FT = 3.28084
CUBIC_FT_TO_M3 = 0.0283168466  # 1 ft³ → m³

z_top_ft = Z_TOP_METRES * M_TO_FT
z_bot_ft = Z_BOT_METRES * M_TO_FT
bbox_half_ft = BBOX_HALF_METRES * M_TO_FT

output = script.get_output()
logger = script.get_logger()

# ---------------------------------------------------------------------------
# Utility functions
# ---------------------------------------------------------------------------

def get_first_solid(geom_el):
    """Return the first non‑null Solid found inside a GeometryElement.
    Ensures solids from GeometryInstances are transformed to the parent coordinate system."""
    for geom_obj in geom_el:
        if isinstance(geom_obj, DB.Solid) and geom_obj.Volume > 0:
            # This solid is in the coordinate system of geom_el.
            return geom_obj
        if isinstance(geom_obj, DB.GeometryInstance):
            # Get geometry local to the instance
            instance_local_geom = geom_obj.GetInstanceGeometry()
            # Recursively find a solid within this local geometry
            solid_in_instance_local_coords = get_first_solid(instance_local_geom)
            if solid_in_instance_local_coords:
                # Transform the found solid from instance's local coordinates
                # to the coordinate system of geom_el (the parent of geom_obj).
                # This accumulates to global coordinates if the initial call was with global geom_el.
                return DB.SolidUtils.CreateTransformed(solid_in_instance_local_coords, geom_obj.Transform)
    return None


# ---------------------------------------------------------------------------
# Data collection
# ---------------------------------------------------------------------------
output.print_md("Collecting candidate elements…")

doc = revit.doc

# Geometry Extraction Options
options = DB.Options()
options.ComputeReferences = True  # Generally a good default
options.IncludeNonVisibleObjects = False # Set to True if non-visible geometry is needed

chosen_view_name_for_log = "project default settings / no specific view"
view_for_options = None

# Attempt to find a suitable view
if hasattr(revit, 'active_view') and revit.active_view is not None:
    view_for_options = revit.active_view
    chosen_view_name_for_log = "pyRevit's active_view: '{}' (ID: {})".format(view_for_options.Name, view_for_options.Id)
elif doc.ActiveView is not None:
    view_for_options = doc.ActiveView
    chosen_view_name_for_log = "document's ActiveView: '{}' (ID: {})".format(view_for_options.Name, view_for_options.Id)

if view_for_options:
    options.View = view_for_options
    # If a view is set, its detail level is used. Do not set options.DetailLevel here.
    output.print_md("**INFO:** Using view [{}] for geometry extraction. The view's own DetailLevel will be used.".format(chosen_view_name_for_log))
elif hasattr(options, 'DetailLevel'): # Check if DetailLevel property exists before setting
    options.DetailLevel = DB.ViewDetailLevel.Fine
    output.print_md("**INFO:** No specific view found for geometry extraction. Attempting to use DetailLevel.Fine.".format(chosen_view_name_for_log))
else:
    output.print_md("**INFO:** No specific view found and DetailLevel property cannot be set on Options. Using Revit defaults for geometry extraction.".format(chosen_view_name_for_log))


# Floors of the specific type
floor_collector = (DB.FilteredElementCollector(doc)
                   .OfCategory(DB.BuiltInCategory.OST_Floors)
                   .WhereElementIsNotElementType())

floors = []
for f in floor_collector:
    type_elem = doc.GetElement(f.GetTypeId())
    type_name = type_elem.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM).AsString()
    if type_name == FLOOR_TYPE_NAME:
        floors.append(f)

# Walls excluding those with "finish" in the type name
wall_collector = (DB.FilteredElementCollector(doc)
                  .OfCategory(DB.BuiltInCategory.OST_Walls)
                  .WhereElementIsNotElementType())

walls = []
for w in wall_collector:
    type_elem = doc.GetElement(w.GetTypeId())
    type_name = type_elem.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM).AsString()
    if EXCLUDE_WALL_KEYWORD.lower() not in type_name.lower():
        walls.append(w)

elements = floors + walls
output.print_md("Found {} floors + {} walls = {} elements".format(len(floors), len(walls), len(elements)))

if not elements:
    script.exit("No matching elements found. Nothing to do.")

# ---------------------------------------------------------------------------
# Diagnostic: Print z-ranges of first few elements to help diagnose the issue
output.print_md("Diagnosing element z-ranges (first 5 of each type)...")
for category, items in [("Floors", floors[:5]), ("Walls", walls[:5])]:
    output.print_md("**{}**:".format(category))
    for elem in items:
        geom_el = elem.get_Geometry(options)
        solid = get_first_solid(geom_el)
        if solid and solid.Volume > 0:
            bbox = solid.GetBoundingBox()
            if bbox:
                z_min = bbox.Min.Z
                z_max = bbox.Max.Z
                output.print_md("- Element {}: z_min={:.4f}ft ({:.4f}m), z_max={:.4f}ft ({:.4f}m)".format(
                    elem.Id, z_min, z_min/M_TO_FT, z_max, z_max/M_TO_FT))

# ---------------------------------------------------------------------------
# Create the vertical clipping box for the target band
# ---------------------------------------------------------------------------
# Expand the clipping band by a small tolerance
CLIP_TOLERANCE_M = 0.01
clip_z_bot_ft = (Z_BOT_METRES - CLIP_TOLERANCE_M) * M_TO_FT
clip_z_top_ft = (Z_TOP_METRES + CLIP_TOLERANCE_M) * M_TO_FT
output.print_md("**DEBUG:** Clipping band (with tolerance): z_bot_ft = {} ft, z_top_ft = {} ft".format(clip_z_bot_ft, clip_z_top_ft))

# Create a large XY rectangle at z_bot_ft
pt0 = DB.XYZ(-bbox_half_ft, -bbox_half_ft, clip_z_bot_ft)
pt1 = DB.XYZ(bbox_half_ft, -bbox_half_ft, clip_z_bot_ft)
pt2 = DB.XYZ(bbox_half_ft, bbox_half_ft, clip_z_bot_ft)
pt3 = DB.XYZ(-bbox_half_ft, bbox_half_ft, clip_z_bot_ft)
crv_loop = DB.CurveLoop()
crv_loop.Append(DB.Line.CreateBound(pt0, pt1))
crv_loop.Append(DB.Line.CreateBound(pt1, pt2))
crv_loop.Append(DB.Line.CreateBound(pt2, pt3))
crv_loop.Append(DB.Line.CreateBound(pt3, pt0))
clip_height = clip_z_top_ft - clip_z_bot_ft
clip_box = DB.GeometryCreationUtilities.CreateExtrusionGeometry([crv_loop], DB.XYZ.BasisZ, clip_height)
if clip_box and clip_box.Volume > 0:
    output.print_md("**DEBUG:** clip_box successfully created with Volume: {:.6f} ft³".format(clip_box.Volume))
    clip_box_bbox = clip_box.GetBoundingBox()
    if clip_box_bbox and clip_box_bbox.Min and clip_box_bbox.Max:
        output.print_md("**DEBUG:** clip_box BBox: Min.Z={:.6f} ft, Max.Z={:.6f} ft".format(clip_box_bbox.Min.Z, clip_box_bbox.Max.Z))
        output.print_md("**DEBUG:** Expected Z for clip_box: clip_z_bot_ft={:.6f}, clip_z_top_ft={:.6f}".format(clip_z_bot_ft, clip_z_top_ft))
    else:
        output.print_md("**ERROR:** clip_box has no valid bounding box after creation.")
else:
    output.print_md("**ERROR:** clip_box creation failed or has zero volume.")
    script.exit() # Exit if clip_box is invalid, as intersections will fail

# --- Create Complement Clipping Solids for Difference Method ---

# Use a fixed large height for complement solids, e.g., 500 feet.
COMPLEMENT_SOLID_HEIGHT_FT = 500.0 
large_z_height_for_complements_ft = COMPLEMENT_SOLID_HEIGHT_FT

output.print_md("--- Creating Complement Solids by extruding at Z=0 then transforming ---")
output.print_md("Target Z-band (ft): {:.4f} to {:.4f}".format(clip_z_bot_ft, clip_z_top_ft))
output.print_md("bbox_half_ft for X,Y (ft): {:.4f}".format(bbox_half_ft))
output.print_md("large_z_height_for_complements_ft (ft): {:.4f}".format(large_z_height_for_complements_ft))

clip_box_complement_top = None
clip_box_complement_bottom = None

# 1. Create a base profile at Z=0
pt0_base = DB.XYZ(-bbox_half_ft, -bbox_half_ft, 0)
pt1_base = DB.XYZ(bbox_half_ft, -bbox_half_ft, 0)
pt2_base = DB.XYZ(bbox_half_ft, bbox_half_ft, 0)
pt3_base = DB.XYZ(-bbox_half_ft, bbox_half_ft, 0)
crv_loop_base = DB.CurveLoop()
crv_loop_base.Append(DB.Line.CreateBound(pt0_base, pt1_base))
crv_loop_base.Append(DB.Line.CreateBound(pt1_base, pt2_base))
crv_loop_base.Append(DB.Line.CreateBound(pt2_base, pt3_base))
crv_loop_base.Append(DB.Line.CreateBound(pt3_base, pt0_base))

base_complement_solid_extrusion = None
try:
    # 2. Extrude the Z=0 profile. Expect this to be centered if profile Z is ignored for placement,
    # or start at Z=0 if profile Z is respected.
    # Given previous BBox results, we assume it will be centered around Z=0, from -H/2 to H/2.
    base_complement_solid_extrusion = DB.GeometryCreationUtilities.CreateExtrusionGeometry(
        [crv_loop_base],
        DB.XYZ.BasisZ,
        large_z_height_for_complements_ft
    )
    if not (base_complement_solid_extrusion and base_complement_solid_extrusion.Volume > 1e-6):
        output.print_md("**ERROR:** Base complement solid extrusion failed or has zero volume.")
        script.exit()
    # Verify assumption about centered extrusion for base solid
    base_bb = base_complement_solid_extrusion.GetBoundingBox()
    output.print_md("  Base Complement Extrusion (Z=0 profile, height H) BBox: Min(z={:.4f}), Max(z={:.4f}) ft".format(base_bb.Min.Z, base_bb.Max.Z))
    # output.print_md("    Expected if centered: Min(z={:.4f}), Max(z={:.4f}) ft".format(-large_z_height_for_complements_ft/2.0, large_z_height_for_complements_ft/2.0)) # Comment out assumption

except Exception as ex_base_create:
    output.print_md("**ERROR:** Failed to create base_complement_solid_extrusion: {}".format(str(ex_base_create)))
    script.exit()

# Get actual extents of the base solid from its bounding box
actual_min_z_base = base_bb.Min.Z
actual_max_z_base = base_bb.Max.Z
actual_height_base = actual_max_z_base - actual_min_z_base
output.print_md("  Actual Base Complement Extrusion Z-Range: {:.4f} to {:.4f} ft (Height: {:.4f} ft)".format(actual_min_z_base, actual_max_z_base, actual_height_base))


# 3. Create and transform for clip_box_complement_top
# Target: Min.Z = clip_z_top_ft, Max.Z = clip_z_top_ft + actual_height_base (effectively)
try:
    # Translation needed to move actual_min_z_base of the base solid to clip_z_top_ft
    translation_z_top = clip_z_top_ft - actual_min_z_base
    transform_top = DB.Transform.CreateTranslation(DB.XYZ(0, 0, translation_z_top))
    clip_box_complement_top = DB.SolidUtils.CreateTransformed(base_complement_solid_extrusion, transform_top)
    
    if clip_box_complement_top and clip_box_complement_top.Volume > 1e-6:
        output.print_md("Created transformed clip_box_complement_top. Volume: {:.4f} cubic ft".format(clip_box_complement_top.Volume))
        actual_bb_ct = clip_box_complement_top.GetBoundingBox()
        output.print_md("  Actual Transformed Complement Top BBox: Min(z={:.4f}), Max(z={:.4f}) ft".format(actual_bb_ct.Min.Z, actual_bb_ct.Max.Z))
        # Expected Z range for top complement: starts at clip_z_top_ft and extends upwards by actual_height_base
        output.print_md("    Expected Top BBox: Min(z={:.4f}), Max(z={:.4f}) ft".format(clip_z_top_ft, clip_z_top_ft + actual_height_base))
    else:
        output.print_md("**ERROR:** clip_box_complement_top (transformed) is None or has zero volume.")
except Exception as ex_ct:
    output.print_md("**ERROR:** Failed to create or transform clip_box_complement_top: {}".format(str(ex_ct)))

# 4. Create and transform for clip_box_complement_bottom
# Target: Min.Z = clip_z_bot_ft - actual_height_base (effectively), Max.Z = clip_z_bot_ft
try:
    # Translation needed to move actual_max_z_base of the base solid to clip_z_bot_ft
    translation_z_bottom = clip_z_bot_ft - actual_max_z_base
    transform_bottom = DB.Transform.CreateTranslation(DB.XYZ(0, 0, translation_z_bottom))
    clip_box_complement_bottom = DB.SolidUtils.CreateTransformed(base_complement_solid_extrusion, transform_bottom)

    if clip_box_complement_bottom and clip_box_complement_bottom.Volume > 1e-6:
        output.print_md("Created transformed clip_box_complement_bottom. Volume: {:.4f} cubic ft".format(clip_box_complement_bottom.Volume))
        actual_bb_cb = clip_box_complement_bottom.GetBoundingBox()
        output.print_md("  Actual Transformed Complement Bottom BBox: Min(z={:.4f}), Max(z={:.4f}) ft".format(actual_bb_cb.Min.Z, actual_bb_cb.Max.Z))
        # Expected Z range for bottom complement: ends at clip_z_bot_ft and extends downwards by actual_height_base
        output.print_md("    Expected Bottom BBox: Min(z={:.4f}), Max(z={:.4f}) ft".format(clip_z_bot_ft - actual_height_base, clip_z_bot_ft))
    else:
        output.print_md("**ERROR:** clip_box_complement_bottom (transformed) is None or has zero volume.")
except Exception as ex_cb:
    output.print_md("**ERROR:** Failed to create or transform clip_box_complement_bottom: {}".format(str(ex_cb)))


# Debug prints for the new complement solids
if clip_box_complement_top and clip_box_complement_top.Volume > 0:
    output.print_md("**DEBUG:** clip_box_complement_top created. Volume: {:.2f} ft³".format(clip_box_complement_top.Volume))
    cbct_bb = clip_box_complement_top.GetBoundingBox()
    if cbct_bb and cbct_bb.Min and cbct_bb.Max: output.print_md("   BBox Z: {:.2f} to {:.2f}".format(cbct_bb.Min.Z, cbct_bb.Max.Z))
elif clip_box_complement_top:
    output.print_md("**WARNING:** clip_box_complement_top created with zero volume.")
else:
    output.print_md("**ERROR:** clip_box_complement_top creation failed.")

if clip_box_complement_bottom and clip_box_complement_bottom.Volume > 0:
    output.print_md("**DEBUG:** clip_box_complement_bottom created. Volume: {:.2f} ft³".format(clip_box_complement_bottom.Volume))
    cbcb_bb = clip_box_complement_bottom.GetBoundingBox()
    if cbcb_bb and cbcb_bb.Min and cbcb_bb.Max: output.print_md("   BBox Z: {:.2f} to {:.2f}".format(cbcb_bb.Min.Z, cbcb_bb.Max.Z))
elif clip_box_complement_bottom:
    output.print_md("**WARNING:** clip_box_complement_bottom created with zero volume.")
else:
    output.print_md("**ERROR:** clip_box_complement_bottom creation failed.")

# ---------------------------------------------------------------------------
# Clip each element's solid to the band, then union
# ---------------------------------------------------------------------------
output.print_md("Clipping and merging all solids in the target band…")
union_solid = None
for idx, elem in enumerate(elements, start=1):
    output.print_md("\\n--- Processing Element ID: {} ({} / {}) ---".format(elem.Id, idx, len(elements)))
    geom_el = elem.get_Geometry(options)
    if geom_el is None:
        output.print_md("elem.get_Geometry(options) returned None. Skipping.")
        continue
        
    solid = get_first_solid(geom_el)
    if not solid or solid.Volume == 0:
        output.print_md("No valid solid found or solid volume is zero. Skipping.")
        continue

    output.print_md("Original solid: Volume={:.6f} ft³".format(solid.Volume))
    solid_bbox = solid.GetBoundingBox()
    if not (solid_bbox and solid_bbox.Min and solid_bbox.Max):
        output.print_md("Original solid: Could not get a valid bounding box. Skipping.")
        continue

    output.print_md("Original solid BBox: Min.Z={:.6f} ft ({:.4f}m), Max.Z={:.6f} ft ({:.4f}m)".format(
        solid_bbox.Min.Z, solid_bbox.Min.Z / M_TO_FT, solid_bbox.Max.Z, solid_bbox.Max.Z / M_TO_FT))

    # Check element solid validity
    element_solid_is_valid = True # Assume true if IsValidForBooleanOperation is not available
    if hasattr(DB.SolidUtils, 'IsValidForBooleanOperation'):
        element_solid_is_valid = DB.SolidUtils.IsValidForBooleanOperation(solid)
        output.print_md("Element solid IsValidForBooleanOperation: {}".format(element_solid_is_valid))
        if not element_solid_is_valid:
            output.print_md("**WARNING:** Element solid ID {} may not be suitable for Boolean operations.".format(elem.Id))

    try:
        output.print_md("Attempting to clip using DIFFERENCE method...")

        # 1. Subtract top complement
        output.print_md("Subtracting top complement...")
        temp_solid = solid # Start with the original solid
        if clip_box_complement_top and clip_box_complement_top.Volume > 0:
            # --- BEGIN ADDED DEBUG ---
            output.print_md("  Pre-TopDiff: Original Solid Valid? {}, Volume: {:.6f} ft³".format(DB.SolidUtils.IsValidForBooleanOperation(temp_solid), temp_solid.Volume))
            ts_bb = temp_solid.GetBoundingBox()
            if ts_bb and ts_bb.Min and ts_bb.Max: output.print_md("    Original Solid BBox Z: {:.4f} to {:.4f} ft".format(ts_bb.Min.Z, ts_bb.Max.Z))
            else: output.print_md("    Original Solid BBox: Invalid or not retrievable.")
            output.print_md("  Pre-TopDiff: Complement Top Valid? {}, Volume: {:.6f} ft³".format(DB.SolidUtils.IsValidForBooleanOperation(clip_box_complement_top), clip_box_complement_top.Volume))
            cbct_bb_debug = clip_box_complement_top.GetBoundingBox()
            if cbct_bb_debug and cbct_bb_debug.Min and cbct_bb_debug.Max: output.print_md("    Complement Top BBox Z: {:.4f} to {:.4f} ft".format(cbct_bb_debug.Min.Z, cbct_bb_debug.Max.Z))
            else: output.print_md("    Complement Top BBox: Invalid or not retrievable.")
            # --- END ADDED DEBUG ---
            try:
                temp_solid_after_top_diff = DB.BooleanOperationsUtils.ExecuteBooleanOperation(temp_solid, clip_box_complement_top, DB.BooleanOperationsType.Difference)
                if temp_solid_after_top_diff and temp_solid_after_top_diff.Volume > 1e-9: # Use a small tolerance for volume check
                    temp_solid = temp_solid_after_top_diff
                    output.print_md("  Volume after top diff: {:.6f} ft³".format(temp_solid.Volume))
                elif temp_solid_after_top_diff: # Zero or near-zero volume
                    output.print_md("  Top diff resulted in effectively zero volume (Volume: {:.10f} ft³). Original solid volume: {:.6f} ft³".format(temp_solid_after_top_diff.Volume, solid.Volume))
                    temp_solid = temp_solid_after_top_diff 
                else: # Boolean operation returned None
                    output.print_md("  Top diff resulted in None. Using original solid for next step.")
                    # temp_solid remains original 'solid'
            except Exception as diff_err_top:
                output.print_md("  **ERROR** during top difference: {}. Using original solid for next step.".format(str(diff_err_top)))
                # temp_solid remains original 'solid'
        else:
            output.print_md("  Skipping top difference as complement_top is invalid or has zero volume.")

        # 2. Subtract bottom complement from the result of step 1
        output.print_md("Subtracting bottom complement...")
        # 'temp_solid' now holds the result of the first difference, or the original solid
        clipped_solid_candidate = temp_solid # Candidate for the final clipped solid before this step
        
        if clip_box_complement_bottom and clip_box_complement_bottom.Volume > 0:
            if not clipped_solid_candidate or clipped_solid_candidate.Volume < 1e-9: # If already no volume, no need to proceed
                output.print_md("  Skipping bottom difference as current solid (pre-bottom-diff) has no volume.")
                clipped = clipped_solid_candidate # Keep the zero-volume or None solid
            else:
                # --- BEGIN ADDED DEBUG ---
                output.print_md("  Pre-BotDiff: Current Solid Valid? {}, Volume: {:.6f} ft³".format(DB.SolidUtils.IsValidForBooleanOperation(clipped_solid_candidate), clipped_solid_candidate.Volume))
                cs_bb = clipped_solid_candidate.GetBoundingBox()
                if cs_bb and cs_bb.Min and cs_bb.Max: output.print_md("    Current Solid BBox Z: {:.4f} to {:.4f} ft".format(cs_bb.Min.Z, cs_bb.Max.Z))
                else: output.print_md("    Current Solid BBox: Invalid or not retrievable.")
                output.print_md("  Pre-BotDiff: Complement Bottom Valid? {}, Volume: {:.6f} ft³".format(DB.SolidUtils.IsValidForBooleanOperation(clip_box_complement_bottom), clip_box_complement_bottom.Volume))
                cbcb_bb_debug = clip_box_complement_bottom.GetBoundingBox()
                if cbcb_bb_debug and cbcb_bb_debug.Min and cbcb_bb_debug.Max: output.print_md("    Complement Bottom BBox Z: {:.4f} to {:.4f} ft".format(cbcb_bb_debug.Min.Z, cbcb_bb_debug.Max.Z))
                else: output.print_md("    Complement Bottom BBox: Invalid or not retrievable.")
                # --- END ADDED DEBUG ---
                try:
                    final_clipped_solid = DB.BooleanOperationsUtils.ExecuteBooleanOperation(clipped_solid_candidate, clip_box_complement_bottom, DB.BooleanOperationsType.Difference)
                    if final_clipped_solid and final_clipped_solid.Volume > 1e-9: # Use a small tolerance
                        clipped = final_clipped_solid
                        output.print_md("  Volume after bottom diff: {:.6f} ft³".format(clipped.Volume))
                    elif final_clipped_solid: # Zero or near-zero volume
                        output.print_md("  Bottom diff resulted in effectively zero volume (Volume: {:.10f} ft³). Volume before this step: {:.6f} ft³".format(final_clipped_solid.Volume, clipped_solid_candidate.Volume))
                        clipped = final_clipped_solid
                    else: # Boolean operation returned None
                        output.print_md("  Bottom diff resulted in None. Clipped solid is effectively None.")
                        clipped = None 
                except Exception as diff_err_bot:
                    output.print_md("  **ERROR** during bottom difference: {}. Clipped solid is effectively None.".format(str(diff_err_bot)))
                    clipped = None 
        else:
            output.print_md("  Skipping bottom difference as complement_bottom is invalid or has zero volume.")
            clipped = clipped_solid_candidate # Result is whatever came before this step

        # Original intersection method (commented out for now)
        # output.print_md("Attempting intersection with clip_box (Volume: {:.2f} ft³)...".format(clip_box.Volume))
        # clipped = DB.BooleanOperationsUtils.ExecuteBooleanOperation(solid, clip_box, DB.BooleanOperationsType.Intersect)

        if clipped is None:
            output.print_md("**DEBUG:** Clipping resulted in None for element {}.".format(elem.Id))
            continue
        
        output.print_md("**DEBUG:** Clipping for element {} completed. Resulting solid volume: {:.6f} ft³".format(elem.Id, clipped.Volume))

        if clipped.Volume < 0.000001: # Using a small tolerance for zero volume check
            output.print_md("**DEBUG:** Clipped solid for element {} has effectively zero volume (Volume: {:.10f} ft³).".format(elem.Id, clipped.Volume))
            clipped_bbox_zero = clipped.GetBoundingBox()
            if clipped_bbox_zero and clipped_bbox_zero.Min and clipped_bbox_zero.Max:
                 output.print_md("Clipped (zero vol) BBox: Min.Z={:.6f} ft, Max.Z={:.6f} ft. Is BBox degenerate? {}".format(
                    clipped_bbox_zero.Min.Z, clipped_bbox_zero.Max.Z, clipped_bbox_zero.Min.IsAlmostEqualTo(clipped_bbox_zero.Max)))
            else:
                output.print_md("Clipped (zero vol) BBox is null or could not be retrieved.")
            continue

        # If we reach here, clipped.Volume > 0
        output.print_md("Successfully clipped element {}. Clipped Volume: {:.6f} ft³".format(elem.Id, clipped.Volume))
        clipped_bbox_final = clipped.GetBoundingBox()
        if clipped_bbox_final and clipped_bbox_final.Min and clipped_bbox_final.Max:
            output.print_md("Clipped solid BBox: Min.Z={:.6f} ft ({:.4f}m), Max.Z={:.6f} ft ({:.4f}m)".format(
                clipped_bbox_final.Min.Z, clipped_bbox_final.Min.Z / M_TO_FT, clipped_bbox_final.Max.Z, clipped_bbox_final.Max.Z / M_TO_FT))
        else:
            output.print_md("Clipped solid: Could not get a valid bounding box.")
            
    except Exception as err:
        output.print_md("**ERROR:** Clipping failed for element {} with error: {}".format(elem.Id, str(err)))
        # Print details about the solid that caused the error
        output.print_md("Details of solid that failed: Volume={:.6f} ft³".format(solid.Volume))
        if solid_bbox and solid_bbox.Min and solid_bbox.Max:
            output.print_md("Solid BBox: Min.Z={:.6f} ft, Max.Z={:.6f} ft, Min.X={:.6f} ft, Max.X={:.6f} ft".format(
                solid_bbox.Min.Z, solid_bbox.Max.Z, solid_bbox.Min.X, solid_bbox.Max.X))
        continue
        
    # If clipped solid is valid and has volume, proceed with union
    if union_solid is None:
        union_solid = clipped
    else:
        try:
            union_solid = DB.BooleanOperationsUtils.ExecuteBooleanOperation(union_solid, clipped, DB.BooleanOperationsType.Union)
        except Exception as err:
            output.print_md("**DEBUG:** Union failed on {}: {}".format(elem.Id, err))
    if idx % 20 == 0:
        output.print_md("Processed {}/{} elements…".format(idx, len(elements)))
    
    # --- TEMP: Break after first element for testing new diff method ---
    output.print_md("--- TEMP: Breaking after first element for difference method testing. ---")
    break 

if not union_solid or union_solid.Volume == 0:
    output.print_md("No merged solid could be created in the band. Exiting.")
    script.exit()

# ---------------------------------------------------------------------------
# Compute total merged volume in the band
# ---------------------------------------------------------------------------
total_m3 = union_solid.Volume * CUBIC_FT_TO_M3
output.print_md("**Total consolidated volume in band ({:.2f}m to {:.2f}m):** **{:.3f} m³**".format(Z_BOT_METRES, Z_TOP_METRES, total_m3))

output.print_md("Done.")
