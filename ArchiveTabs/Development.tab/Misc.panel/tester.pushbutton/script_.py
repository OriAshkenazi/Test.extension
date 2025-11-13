#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
Concrete Volume Calculator for Specific Elevation Slice - Enhanced with Debugging
Calculates the total volume of concrete (floors and structural walls) 
within a specific elevation band, handling overlaps via boolean union.
"""

__title__ = "Concrete Slice Volume (Debug)"
__author__ = "Your Name"
__doc__ = "Calculate concrete volume in elevation band with debugging"

# Revit API imports
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *
from Autodesk.Revit.Exceptions import *
from Autodesk.Revit.DB import SolidUtils

# System imports
import clr
import sys
from System.Collections.Generic import List

# Add references for geometry operations
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')

# pyRevit imports
from pyrevit import script, forms, DB, revit

# Get current document
doc = revit.doc
uidoc = revit.uidoc
app = doc.Application

# Output handler
output = script.get_output()

# Configuration options
DEBUG_MODE = True  # Set to True for detailed debugging output
SKIP_INTRA_ELEMENT_UNION = True  # Set to True to skip combining solids within elements (recommended if union errors occur)
AUTO_DETECT_FLOOR_ELEVATION = True  # Set to True to automatically find floor elevations
TRANSFORM_GEOMETRY = True  # Set to True to transform geometry from local to world coordinates

# Default elevation range (in meters, will be converted to feet)
DEFAULT_MIN_ELEVATION = -1.62  # meters
DEFAULT_MAX_ELEVATION = -1.22  # meters

def debug_print(message, indent=0):
    """Print debug message if debug mode is enabled"""
    if DEBUG_MODE:
        prefix = "  " * indent + "DEBUG: "
        output.print_md(prefix + message)

def meters_to_feet(meters):
    """Convert meters to feet"""
    return meters * 3.28084

def feet_to_meters(feet):
    """Convert feet to meters"""
    return feet / 3.28084

def get_element_name(element):
    """Safely get element name or type name"""
    try:
        element_type = doc.GetElement(element.GetTypeId())
        if element_type:
            return element_type.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString()
    except:
        pass
    
    try:
        name_param = element.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS)
        if name_param:
            return name_param.AsString()
    except:
        pass
    
    return "Unknown"

def get_element_bounds(element):
    """Get element bounding box and return min/max Z in both feet and meters"""
    try:
        bbox = element.get_BoundingBox(None)
        if bbox:
            return {
                'min_z_ft': bbox.Min.Z,
                'max_z_ft': bbox.Max.Z,
                'min_z_m': feet_to_meters(bbox.Min.Z),
                'max_z_m': feet_to_meters(bbox.Max.Z)
            }
    except:
        pass
    return None

def collect_relevant_elements():
    """Collect floors and walls based on criteria with debugging"""
    floors = []
    walls = []
    floor_elevations = []
    
    debug_print("Starting element collection...")
    
    # Collect floors with type "MGN_Floor 40 cm"
    floor_collector = FilteredElementCollector(doc).OfClass(Floor).WhereElementIsNotElementType()
    all_floor_count = floor_collector.GetElementCount()
    debug_print("Total floors in model: {}".format(all_floor_count))
    
    for floor in floor_collector:
        try:
            floor_type = doc.GetElement(floor.GetTypeId())
            if floor_type:
                type_name = floor_type.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString()
                
                # Debug: Show all floor types found
                if DEBUG_MODE and type_name:
                    bounds = get_element_bounds(floor)
                    if bounds:
                        debug_print("Floor Type: '{}', ID: {}, Z-range: {:.2f}m to {:.2f}m ({:.2f}ft to {:.2f}ft)".format(
                            type_name, floor.Id.IntegerValue,
                            bounds['min_z_m'], bounds['max_z_m'],
                            bounds['min_z_ft'], bounds['max_z_ft']
                        ), 1)
                
                # Check for floor type with or without space before "cm"
                if type_name == "MGN_Floor 40 cm" or type_name == "MGN_Floor 40cm":
                    floors.append(floor)
                    bounds = get_element_bounds(floor)
                    if bounds:
                        floor_elevations.extend([bounds['min_z_m'], bounds['max_z_m']])
                        debug_print("Found target floor: ID {}, Elevation: {:.2f}m to {:.2f}m".format(
                            floor.Id.IntegerValue, bounds['min_z_m'], bounds['max_z_m']
                        ), 1)
        except Exception as e:
            debug_print("Error processing floor: {}".format(str(e)), 1)
    
    # Collect walls that don't have "finish" in their type name
    wall_collector = FilteredElementCollector(doc).OfClass(Wall).WhereElementIsNotElementType()
    all_wall_count = wall_collector.GetElementCount()
    debug_print("Total walls in model: {}".format(all_wall_count))
    
    structural_wall_count = 0
    for wall in wall_collector:
        try:
            wall_type = doc.GetElement(wall.GetTypeId())
            if wall_type:
                type_name = wall_type.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString()
                if type_name and "finish" not in type_name.lower():
                    walls.append(wall)
                    structural_wall_count += 1
                    
                    if DEBUG_MODE and structural_wall_count <= 5:  # Show first 5 walls
                        bounds = get_element_bounds(wall)
                        if bounds:
                            debug_print("Structural wall: '{}', ID: {}, Z-range: {:.2f}m to {:.2f}m".format(
                                type_name, wall.Id.IntegerValue,
                                bounds['min_z_m'], bounds['max_z_m']
                            ), 1)
        except Exception as e:
            debug_print("Error processing wall: {}".format(str(e)), 1)
    
    debug_print("Collection complete: {} floors, {} structural walls".format(len(floors), len(walls)))
    
    return floors, walls, floor_elevations

def get_element_geometry(element):
    """Get solid geometry from element (basic method without transformation)"""
    options = Options()
    options.ComputeReferences = True
    options.IncludeNonVisibleObjects = True
    options.DetailLevel = ViewDetailLevel.Fine
    
    geom_elem = element.get_Geometry(options)
    solids = []
    
    if geom_elem:
        for geom_obj in geom_elem:
            if isinstance(geom_obj, Solid) and geom_obj.Volume > 0:
                solids.append(geom_obj)
            elif isinstance(geom_obj, GeometryInstance):
                instance_geom = geom_obj.GetInstanceGeometry()
                if instance_geom:
                    for inst_obj in instance_geom:
                        if isinstance(inst_obj, Solid) and inst_obj.Volume > 0:
                            solids.append(inst_obj)
    
    return solids

def get_element_geometry_with_transform(element):
    """Get solid geometry from element with proper transformation"""
    element_id = element.Id.IntegerValue
    debug_print("Getting geometry for element {}...".format(element_id), 2)
    
    # Get element's actual bounding box first
    elem_bbox = element.get_BoundingBox(None)
    if elem_bbox:
        debug_print("Element bbox: Z = {:.2f}ft to {:.2f}ft".format(elem_bbox.Min.Z, elem_bbox.Max.Z), 3)
    
    options = Options()
    options.ComputeReferences = True
    options.IncludeNonVisibleObjects = True
    options.DetailLevel = ViewDetailLevel.Fine
    
    # Try to get geometry with transformation
    try:
        # For some elements, we need to get the transformed geometry
        transform = Transform.Identity
        
        # Check if element has a location
        location = element.Location
        if location:
            if hasattr(location, 'Point'):
                # Point-based location (like some families)
                point = location.Point
                transform = Transform.CreateTranslation(point)
                debug_print("Using point-based transform: Z = {:.2f}ft".format(point.Z), 3)
            elif hasattr(location, 'Curve'):
                # Curve-based location (like some walls)
                curve = location.Curve
                if curve:
                    start_point = curve.GetEndPoint(0)
                    transform = Transform.CreateTranslation(XYZ(0, 0, start_point.Z))
                    debug_print("Using curve-based transform: Z = {:.2f}ft".format(start_point.Z), 3)
        
        # Get geometry
        geom_elem = element.get_Geometry(options)
        solids = []
        
        if geom_elem:
            for geom_obj in geom_elem:
                if isinstance(geom_obj, Solid) and geom_obj.Volume > 0:
                    # Check if solid needs transformation
                    bbox = geom_obj.GetBoundingBox()
                    if bbox and elem_bbox and abs(bbox.Min.Z - elem_bbox.Min.Z) > 1.0:
                        # Solid is not at element location, transform it
                        z_offset = elem_bbox.Min.Z - bbox.Min.Z
                        solid_transform = Transform.CreateTranslation(XYZ(0, 0, z_offset))
                        transformed = SolidUtils.CreateTransformed(geom_obj, solid_transform)
                        solids.append(transformed)
                        debug_print("Transformed solid by Z-offset: {:.2f}ft".format(z_offset), 3)
                    else:
                        solids.append(geom_obj)
                    debug_print("Found solid with volume: {:.3f} ft³".format(geom_obj.Volume), 3)
                    
                elif isinstance(geom_obj, GeometryInstance):
                    instance_transform = geom_obj.Transform
                    instance_geom = geom_obj.GetInstanceGeometry()
                    
                    if instance_geom:
                        for inst_obj in instance_geom:
                            if isinstance(inst_obj, Solid) and inst_obj.Volume > 0:
                                # Apply instance transform
                                transformed = SolidUtils.CreateTransformed(inst_obj, instance_transform)
                                
                                # Check if additional transform needed
                                bbox = transformed.GetBoundingBox()
                                if bbox and elem_bbox and abs(bbox.Min.Z - elem_bbox.Min.Z) > 1.0:
                                    z_offset = elem_bbox.Min.Z - bbox.Min.Z
                                    additional_transform = Transform.CreateTranslation(XYZ(0, 0, z_offset))
                                    transformed = SolidUtils.CreateTransformed(transformed, additional_transform)
                                    debug_print("Applied additional Z-offset: {:.2f}ft".format(z_offset), 3)
                                
                                solids.append(transformed)
                                debug_print("Found instance solid with volume: {:.3f} ft³".format(transformed.Volume), 3)
        
        debug_print("Total solids found: {}".format(len(solids)), 2)
        return solids
        
    except Exception as e:
        debug_print("Error in geometry extraction: {}".format(str(e)), 2)
        # Fallback to simple method
        return get_element_geometry(element)

def combine_element_solids(solids, element_id):
    """Combine multiple solids from same element into one"""
    if not solids:
        return None
    
    # Filter out zero-volume solids
    valid_solids = [s for s in solids if s.Volume > 0.001]  # Small tolerance for numerical errors
    
    if not valid_solids:
        debug_print("No valid solids with volume > 0.001 ft³", 3)
        return None
    
    if len(valid_solids) == 1:
        return valid_solids[0]
    
    if SKIP_INTRA_ELEMENT_UNION:
        debug_print("Skipping intra-element union (returning largest solid)", 3)
        return max(valid_solids, key=lambda s: s.Volume)
    
    debug_print("Attempting to union {} valid solids for element {}".format(len(valid_solids), element_id), 3)
    
    try:
        result = valid_solids[0]
        for i in range(1, len(valid_solids)):
            debug_print("Union operation {}/{}...".format(i, len(valid_solids)-1), 4)
            result = BooleanOperationsUtils.ExecuteBooleanOperation(
                result, valid_solids[i], BooleanOperationsType.Union)
        
        debug_print("Successfully combined solids. Final volume: {:.3f} ft³".format(result.Volume), 3)
        return result
        
    except Exception as e:
        debug_print("Failed to combine solids: {}".format(str(e)), 3)
        debug_print("Returning largest solid as fallback", 3)
        # Return the actual largest solid, not just any object
        largest_solid = max(valid_solids, key=lambda s: s.Volume)
        debug_print("Largest solid volume: {:.3f} ft³".format(largest_solid.Volume), 3)
        return largest_solid

def slice_solid_by_elevation(solid, z_bottom, z_top, element_id):
    """Slice a solid by elevation range with debugging"""
    try:
        # Verify we have a valid solid
        if not solid or not hasattr(solid, 'Volume'):
            debug_print("Invalid solid object for element {}".format(element_id), 3)
            return None
            
        if solid.Volume <= 0.001:
            debug_print("Solid has zero or negligible volume for element {}".format(element_id), 3)
            return None
        
        # Get bounding box with error handling
        bbox = None
        try:
            bbox = solid.get_BoundingBox()
        except:
            try:
                # Alternative method if get_BoundingBox doesn't work
                debug_print("Using alternative bounding box method", 3)
                bbox = solid.GetBoundingBox()
            except:
                debug_print("Using computed bounding box method", 3)
                # Compute bounding box from vertices
                bbox = solid.ComputeBoundingBox()
            
        if not bbox:
            debug_print("Could not get bounding box for element {}".format(element_id), 3)
            return None
            
        solid_min_z = bbox.Min.Z
        solid_max_z = bbox.Max.Z
        
        debug_print("Slicing solid for element {}: Z-range {:.2f}ft to {:.2f}ft ({:.2f}m to {:.2f}m)".format(
            element_id, solid_min_z, solid_max_z, 
            feet_to_meters(solid_min_z), feet_to_meters(solid_max_z)
        ), 3)
        debug_print("Target slice: {:.2f}ft to {:.2f}ft ({:.2f}m to {:.2f}m)".format(
            z_bottom, z_top, feet_to_meters(z_bottom), feet_to_meters(z_top)
        ), 3)
        
        # Check if solid intersects with slice
        if solid_max_z < z_bottom or solid_min_z > z_top:
            debug_print("Solid is OUT OF RANGE", 3)
            debug_print("Difference: solid top is {:.2f}ft below slice bottom".format(z_bottom - solid_max_z), 3)
            return None
        
        debug_print("Solid INTERSECTS with slice", 3)
        
        # Create cutting planes
        bottom_plane = Plane.CreateByNormalAndOrigin(XYZ.BasisZ, XYZ(0, 0, z_bottom))
        top_plane = Plane.CreateByNormalAndOrigin(XYZ.BasisZ.Negate(), XYZ(0, 0, z_top))
        
        # Cut with bottom plane (keep above)
        result = solid
        if solid_min_z < z_bottom:
            debug_print("Cutting with bottom plane...", 4)
            try:
                result = BooleanOperationsUtils.CutWithHalfSpace(solid, bottom_plane)
                if result:
                    debug_print("Bottom cut successful. Volume: {:.3f} ft³".format(result.Volume), 4)
            except Exception as e:
                debug_print("Bottom cut failed: {}".format(str(e)), 4)
                return None
        
        # Cut with top plane (keep below)
        if result and result.Volume > 0 and solid_max_z > z_top:
            debug_print("Cutting with top plane...", 4)
            try:
                result = BooleanOperationsUtils.CutWithHalfSpace(result, top_plane)
                if result:
                    debug_print("Top cut successful. Volume: {:.3f} ft³".format(result.Volume), 4)
            except Exception as e:
                debug_print("Top cut failed: {}".format(str(e)), 4)
                return None
        
        if result and result.Volume > 0:
            debug_print("Slicing complete. Final volume: {:.3f} ft³ ({:.3f} m³)".format(
                result.Volume, result.Volume / 35.3147
            ), 3)
            return result
            
    except Exception as e:
        debug_print("Error during slicing: {}".format(str(e)), 3)
        import traceback
        debug_print("Traceback: {}".format(traceback.format_exc()), 3)
    
    return None

def calculate_slice_volumes():
    """Main function to calculate volumes in elevation slice"""
    
    # Collect elements first to potentially auto-detect elevation
    floors, walls, floor_elevations = collect_relevant_elements()
    
    # Determine elevation range
    if AUTO_DETECT_FLOOR_ELEVATION and floor_elevations:
        # Use the average floor elevation as center of slice
        avg_floor_elevation = sum(floor_elevations) / len(floor_elevations)
        min_elevation = avg_floor_elevation - 0.2  # 40cm slice centered on floor
        max_elevation = avg_floor_elevation + 0.2
        
        output.print_md("### Auto-detected floor elevation: {:.2f}m".format(avg_floor_elevation))
        output.print_md("### Using adjusted elevation range: {:.2f}m to {:.2f}m".format(min_elevation, max_elevation))
    else:
        min_elevation = DEFAULT_MIN_ELEVATION
        max_elevation = DEFAULT_MAX_ELEVATION
        if not floor_elevations:
            output.print_md("### No target floors found for auto-detection")
        output.print_md("### Using default elevation range: {:.2f}m to {:.2f}m".format(min_elevation, max_elevation))
    
    # Convert to feet
    z_bottom = meters_to_feet(min_elevation)
    z_top = meters_to_feet(max_elevation)
    
    debug_print("Elevation range in feet: {:.2f}ft to {:.2f}ft".format(z_bottom, z_top))
    
    output.print_md("## Concrete Volume Analysis")
    output.print_md("---")
    
    output.print_md("### Elements Found:")
    output.print_md("- **Floors (MGN_Floor 40 cm):** {}".format(len(floors)))
    output.print_md("- **Structural Walls:** {}".format(len(walls)))
    output.print_md("---")
    
    if not floors:
        output.print_md("**⚠️ No floors found with type 'MGN_Floor 40 cm' or 'MGN_Floor 40cm'**")
        output.print_md("Please check:")
        output.print_md("1. Floor type name spelling (with or without space before 'cm')")
        output.print_md("2. If floors exist with this exact type name")
        output.print_md("3. The debug output above shows all floor types found")
        output.print_md("---")
    
    if not floors and not walls:
        output.print_md("**No relevant elements found!**")
        output.print_md("\n### Troubleshooting:")
        output.print_md("1. Check if floor type name is exactly 'MGN_Floor 40 cm'")
        output.print_md("2. Verify elements exist in the model")
        output.print_md("3. Enable DEBUG_MODE for detailed information")
        return
    
    # Debug: Check for coordinate system issues
    if floors and DEBUG_MODE:
        test_floor = floors[0]
        test_bbox = test_floor.get_BoundingBox(None)
        test_solids = get_element_geometry(test_floor)
        if test_solids and test_bbox:
            solid_bbox = test_solids[0].GetBoundingBox()
            if solid_bbox:
                z_diff = abs(test_bbox.Min.Z - solid_bbox.Min.Z)
                if z_diff > 10:  # More than 10 feet difference
                    output.print_md("### ⚠️ COORDINATE SYSTEM MISMATCH DETECTED!")
                    output.print_md("Element is at {:.1f}m but geometry is at {:.1f}m".format(
                        feet_to_meters(test_bbox.Min.Z), feet_to_meters(solid_bbox.Min.Z)
                    ))
                    output.print_md("**Solution: Ensure TRANSFORM_GEOMETRY = True**")
                    output.print_md("---")
    
    # Process elements and collect sliced solids
    element_data = []
    all_sliced_solids = []
    
    debug_print("\nProcessing {} floors...".format(len(floors)))
    
    # Process floors
    for i, floor in enumerate(floors):
        element_id = floor.Id.IntegerValue
        debug_print("\nProcessing floor {}/{} (ID: {})".format(i+1, len(floors), element_id), 1)
        
        # Get geometry with or without transformation based on configuration
        if TRANSFORM_GEOMETRY:
            solids = get_element_geometry_with_transform(floor)
        else:
            solids = get_element_geometry(floor)
        if not solids:
            debug_print("No solids found for floor {}".format(element_id), 2)
            continue
        
        # Combine solids if multiple
        combined_solid = combine_element_solids(solids, element_id)
        if not combined_solid:
            debug_print("Failed to get combined solid for floor {}".format(element_id), 2)
            continue
        
        # Slice the solid
        sliced = slice_solid_by_elevation(combined_solid, z_bottom, z_top, element_id)
        if sliced and sliced.Volume > 0:
            volume_m3 = sliced.Volume / 35.3147
            element_data.append({
                'element': floor,
                'type': 'Floor',
                'name': get_element_name(floor),
                'id': floor.Id,
                'original_volume': volume_m3,
                'solid': sliced
            })
            all_sliced_solids.append(sliced)
            debug_print("Floor {} added to results. Volume: {:.3f} m³".format(element_id, volume_m3), 2)
        else:
            debug_print("Floor {} has no volume in slice".format(element_id), 2)
    
    debug_print("\nProcessing {} walls...".format(len(walls)))
    
    # Process walls (similar logic)
    walls_processed = 0
    for wall in walls:
        element_id = wall.Id.IntegerValue
        walls_processed += 1
        
        if DEBUG_MODE and walls_processed <= 10:  # Detailed debug for first 10 walls
            debug_print("\nProcessing wall {}/{} (ID: {})".format(walls_processed, len(walls), element_id), 1)
        
        # Get geometry with or without transformation based on configuration
        if TRANSFORM_GEOMETRY:
            solids = get_element_geometry_with_transform(wall)
        else:
            solids = get_element_geometry(wall)
        if not solids:
            continue
        
        combined_solid = combine_element_solids(solids, element_id)
        if not combined_solid:
            continue
        
        sliced = slice_solid_by_elevation(combined_solid, z_bottom, z_top, element_id)
        if sliced and sliced.Volume > 0:
            volume_m3 = sliced.Volume / 35.3147
            element_data.append({
                'element': wall,
                'type': 'Wall',
                'name': get_element_name(wall),
                'id': wall.Id,
                'original_volume': volume_m3,
                'solid': sliced
            })
            all_sliced_solids.append(sliced)
            if DEBUG_MODE and walls_processed <= 10:
                debug_print("Wall {} added to results. Volume: {:.3f} m³".format(element_id, volume_m3), 2)
    
    debug_print("\nTotal elements with volume in slice: {}".format(len(element_data)))
    
    if not all_sliced_solids:
        output.print_md("**No geometry found in the specified elevation range!**")
        output.print_md("\n### Diagnostic Summary:")
        
        if not floors:
            output.print_md("❌ **No target floors found** - Check floor type name")
            output.print_md("   Looking for: 'MGN_Floor 40 cm' or 'MGN_Floor 40cm'")
            output.print_md("   Found types: Check debug output above")
        else:
            output.print_md("✓ Found {} target floors".format(len(floors)))
            
        output.print_md("\n📏 **Elevation Range Used:** {:.2f}m to {:.2f}m ({:.2f}ft to {:.2f}ft)".format(
            min_elevation, max_elevation, z_bottom, z_top
        ))
        
        output.print_md("\n### Recommended Actions:")
        output.print_md("1. **If geometry appears at wrong elevation (near 0,0,0):**")
        output.print_md("   - Ensure TRANSFORM_GEOMETRY = True (currently: {})".format(TRANSFORM_GEOMETRY))
        output.print_md("   - This handles local vs world coordinate issues")
        output.print_md("2. **If floor type name is different:**")
        output.print_md("   - Update line ~158 to match your exact floor type name")
        output.print_md("3. **If union errors occur:**")
        output.print_md("   - Already set: SKIP_INTRA_ELEMENT_UNION = True")
        output.print_md("4. **If elevation is wrong:**")
        output.print_md("   - Set AUTO_DETECT_FLOOR_ELEVATION = False")
        output.print_md("   - Adjust DEFAULT_MIN_ELEVATION and DEFAULT_MAX_ELEVATION")
        output.print_md("5. **Review debug output** for specific element elevations")
        return
    
    # Perform boolean union to eliminate overlaps
    output.print_md("### Calculating Union Volume...")
    
    try:
        # Start with first solid
        union_solid = all_sliced_solids[0]
        failed_unions = 0
        
        # Union with remaining solids
        for i in range(1, len(all_sliced_solids)):
            try:
                debug_print("Union operation {}/{}...".format(i, len(all_sliced_solids)-1), 1)
                union_solid = BooleanOperationsUtils.ExecuteBooleanOperation(
                    union_solid, 
                    all_sliced_solids[i], 
                    BooleanOperationsType.Union
                )
            except Exception as e:
                failed_unions += 1
                debug_print("Union failed for solid {}: {}".format(i, str(e)), 1)
        
        if failed_unions > 0:
            debug_print("Warning: {} union operations failed".format(failed_unions))
        
        # Calculate final volume
        total_volume_m3 = union_solid.Volume / 35.3147
        union_successful = True
        
    except Exception as e:
        output.print_md("**Error during boolean union: {}**".format(str(e)))
        # Fallback: sum all volumes (with overlap)
        total_volume_m3 = sum(data['original_volume'] for data in element_data)
        union_successful = False
        output.print_md("*Using sum of individual volumes (may include overlaps)*")
    
    # Output results
    output.print_md("---")
    output.print_md("## Results")
    output.print_md("### **Total Net Volume: {:.2f} m³**".format(total_volume_m3))
    if not union_successful:
        output.print_md("*Note: This may include overlaps due to union failure*")
    output.print_md("---")
    
    # Detailed breakdown
    output.print_md("### Element Contributions:")
    
    # Sort by volume
    element_data.sort(key=lambda x: x['original_volume'], reverse=True)
    
    # Create table
    headers = ["Element Type", "Element ID", "Type Name", "Volume (m³)"]
    rows = []
    
    for data in element_data:
        rows.append([
            data['type'],
            str(data['id'].IntegerValue),
            data['name'],
            "{:.3f}".format(data['original_volume'])
        ])
    
    output.print_table(headers, rows)
    
    # Summary statistics
    output.print_md("---")
    output.print_md("### Summary:")
    output.print_md("- **Total Elements in Slice:** {}".format(len(element_data)))
    output.print_md("- **Floor Elements:** {}".format(len([d for d in element_data if d['type'] == 'Floor'])))
    output.print_md("- **Wall Elements:** {}".format(len([d for d in element_data if d['type'] == 'Wall'])))
    output.print_md("- **Sum of Individual Volumes:** {:.2f} m³".format(sum(data['original_volume'] for data in element_data)))
    output.print_md("- **Net Volume (after union):** {:.2f} m³".format(total_volume_m3))
    if union_successful:
        overlap = sum(data['original_volume'] for data in element_data) - total_volume_m3
        output.print_md("- **Overlap Volume:** {:.2f} m³".format(overlap))
    
    # Debug summary
    if DEBUG_MODE:
        output.print_md("\n### Debug Summary:")
        output.print_md("- Elevation Range: {:.2f}m to {:.2f}m ({:.2f}ft to {:.2f}ft)".format(
            min_elevation, max_elevation, z_bottom, z_top
        ))
        output.print_md("- Auto-detect elevation: {}".format("Enabled" if AUTO_DETECT_FLOOR_ELEVATION else "Disabled"))
        output.print_md("- Skip intra-element union: {}".format("Yes" if SKIP_INTRA_ELEMENT_UNION else "No"))

# Main execution
if __name__ == '__main__':
    try:
        # Start transaction for any potential modifications
        t = Transaction(doc, "Calculate Concrete Slice Volume (Debug)")
        t.Start()
        
        output.print_md("# Concrete Slice Volume Calculator (Enhanced)")
        output.print_md("### Configuration:")
        output.print_md("- Debug Mode: **{}**".format("ON" if DEBUG_MODE else "OFF"))
        output.print_md("- Skip Intra-Element Union: **{}**".format("YES" if SKIP_INTRA_ELEMENT_UNION else "NO"))
        output.print_md("- Auto-Detect Floor Elevation: **{}**".format("YES" if AUTO_DETECT_FLOOR_ELEVATION else "NO"))
        output.print_md("- Transform Geometry: **{}**".format("YES" if TRANSFORM_GEOMETRY else "NO"))
        output.print_md("---")
        
        calculate_slice_volumes()
        
        t.Commit()
        
    except Exception as e:
        if t.HasStarted():
            t.RollBack()
        
        output.print_md("## Critical Error")
        output.print_md("An error occurred: {}".format(str(e)))
        
        import traceback
        output.print_md("```")
        output.print_md(traceback.format_exc())
        output.print_md("```")