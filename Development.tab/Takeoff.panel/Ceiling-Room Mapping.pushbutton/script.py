#! python3
"""
Revit Ceiling-Room Mapping Script
--------------------------------
This script maps ceilings to rooms in a Revit model based on their geometric relationships.
It uses sophisticated intersection checking including direct 3D intersection, XY projection,
and area calculations.

Key Features:
- Handles both direct 3D intersections and XY plane projections
- Computes intersection areas for better accuracy
- Considers complex shape geometries
- Uses convex hull for ceiling projections
- Provides fallback mechanisms for geometric operations
"""

from Autodesk.Revit.DB import *
from System.Collections.Generic import List
from scipy.spatial import ConvexHull
from shapely.geometry import Polygon
from shapely.ops import unary_union
from shapely.validation import make_valid
from typing import Optional, Tuple
import logging
import traceback

# ========================
# CONSTANTS AND GLOBALS
# ========================

FEET_TO_INTERNAL = 0.0833333333  # Conversion factor for feet to internal units
DEFAULT_TOLERANCE = 1e-2         # Default tolerance for geometric comparisons
VOLUME_THRESHOLD = 1e-6         # Minimum volume for solid intersection validation
NORMAL_TOLERANCE = 0.001        # Tolerance for face normal calculations

# ========================
# GEOMETRY UTILITIES
# ========================

def get_solids(element):
    """
    Extract valid solids from a Revit element's geometry.
    
    Args:
        element: Revit element to extract geometry from
        
    Returns:
        list: List of valid Solid objects with positive volume
    """
    geometry = element.get_Geometry(Options())
    solids = []
    if geometry:
        for geom_obj in geometry:
            if isinstance(geom_obj, Solid) and geom_obj.Volume > 0:
                solids.append(geom_obj)
    return solids

def get_element_bounding_box(element_id):
    """
    Get the bounding box of an element.
    
    Args:
        element_id (ElementId): The ID of the element
        
    Returns:
        BoundingBoxXYZ: The element's bounding box
    """
    element = doc.GetElement(element_id)
    return element.get_Geometry(Options()).GetBoundingBox()

def project_to_xy_plane(solid, is_ceiling=False):
    """
    Project the solid geometry to the XY plane, creating a single polygon for ceilings.
    
    Args:
        solid (Solid): The solid geometry
        is_ceiling (bool): Whether the solid represents a ceiling
        
    Returns:
        Polygon or list[Polygon]: The projected polygon(s) on the XY plane
    """
    vertices = []
    for face in solid.Faces:
        try:
            face_normal = face.ComputeNormal(UV(0.5, 0.5))
            if abs(face_normal.Z) > NORMAL_TOLERANCE:
                edge_loops = face.GetEdgesAsCurveLoops()
                for loop in edge_loops:
                    for curve in loop:
                        for point in curve.Tessellate():
                            vertices.append((point.X, point.Y))
        except Exception as e:
            logging.exception(f"Error processing face in project_to_xy_plane: {e}")
            continue
    
    if not vertices:
        return None

    try:
        if is_ceiling:
            # Create single polygon using convex hull for ceilings
            hull = ConvexHull(vertices)
            hull_points = [vertices[i] for i in hull.vertices]
            polygon = Polygon(hull_points)
            return polygon if polygon.is_valid else make_valid(polygon)
        else:
            # Create multiple polygons if necessary for rooms
            polygon = Polygon(vertices)
            return [polygon if polygon.is_valid else make_valid(polygon)]
    except Exception as e:
        logging.exception(f"Error creating polygon: {e}")
        return None

# ========================
# INTERSECTION CHECKS
# ========================

def check_bounding_box_intersection(bb1, bb2):
    """
    Check if two bounding boxes intersect.
    
    Args:
        bb1 (BoundingBoxXYZ): First bounding box
        bb2 (BoundingBoxXYZ): Second bounding box
        
    Returns:
        bool: True if the bounding boxes intersect
    """
    return (bb1.Min.X <= bb2.Max.X and bb1.Max.X >= bb2.Min.X and
            bb1.Min.Y <= bb2.Max.Y and bb1.Max.Y >= bb2.Min.Y and
            bb1.Min.Z <= bb2.Max.Z and bb1.Max.Z >= bb2.Min.Z)

def check_direct_intersection(room_id, ceiling_id, prefix="", timestamp=""):
    """
    Check if there's a direct intersection between room and ceiling geometries.
    
    Args:
        room_id (ElementId): Room element ID
        ceiling_id (ElementId): Ceiling element ID
        prefix (str): Optional prefix for debugging
        timestamp (str): Optional timestamp for debugging
        
    Returns:
        bool: True if there's a direct intersection
    """
    room = doc.GetElement(room_id)
    ceiling = doc.GetElement(ceiling_id)
    room_geom = get_solids(room)
    ceiling_geom = get_solids(ceiling)

    for room_solid in room_geom:
        for ceiling_solid in ceiling_geom:
            try:
                intersection_solid = BooleanOperationsUtils.ExecuteBooleanOperation(
                    room_solid, ceiling_solid, BooleanOperationsType.Intersect)
                if intersection_solid.Volume > VOLUME_THRESHOLD:
                    return True
            except InvalidOperationException:
                # Fallback to bounding box check
                room_bb = room_solid.GetBoundingBox()
                ceiling_bb = ceiling_solid.GetBoundingBox()
                if check_bounding_box_intersection(room_bb, ceiling_bb):
                    return True
    return False

def project_and_check_xy_intersection(room_id, ceiling_id):
    """
    Project geometries to XY plane and check for intersection.
    
    Args:
        room_id (ElementId): Room element ID
        ceiling_id (ElementId): Ceiling element ID
        
    Returns:
        bool: True if there's an intersection in the XY projection
    """
    room = doc.GetElement(room_id)
    ceiling = doc.GetElement(ceiling_id)
    room_geom = get_solids(room)
    ceiling_geom = get_solids(ceiling)
    
    if not room_geom or not ceiling_geom:
        return False
    
    room_polygons = []
    ceiling_polygons = []
    
    for solid in room_geom:
        proj = project_to_xy_plane(solid, is_ceiling=False)
        if proj:
            room_polygons.extend(proj)
            
    for solid in ceiling_geom:
        proj = project_to_xy_plane(solid, is_ceiling=True)
        if proj:
            ceiling_polygons.append(proj)
    
    room_polygons = [p for p in room_polygons if p is not None and p.is_valid]
    ceiling_polygons = [p for p in ceiling_polygons if p is not None and p.is_valid]

    if not room_polygons or not ceiling_polygons:
        return False
    
    try:
        room_union = unary_union(room_polygons)
        ceiling_union = unary_union(ceiling_polygons)
        return room_union.intersects(ceiling_union)
    except Exception as e:
        logging.exception(f"Error in project_and_check_xy_intersection: {e}")
        # Fallback to bounding box check
        room_bb = get_element_bounding_box(room_id)
        ceiling_bb = get_element_bounding_box(ceiling_id)
        return check_bounding_box_intersection(room_bb, ceiling_bb)

def delta_ceiling_above_room(room_id, ceiling_id):
    """
    Calculate the vertical distance between a room and ceiling.
    
    Args:
        room_id (ElementId): Room element ID
        ceiling_id (ElementId): Ceiling element ID
        
    Returns:
        float: Distance between the ceiling and room
    """
    room_bb = get_element_bounding_box(room_id)
    ceiling_bb = get_element_bounding_box(ceiling_id)
    return ceiling_bb.Min.Z - room_bb.Max.Z

def check_intersections(room_id, ceiling_id, prefix="", timestamp=""):
    """
    Comprehensive intersection check between room and ceiling.
    
    Args:
        room_id (ElementId): Room element ID
        ceiling_id (ElementId): Ceiling element ID
        prefix (str): Optional prefix for debugging
        timestamp (str): Optional timestamp for debugging
        
    Returns:
        tuple or None: (direct_intersection, xy_intersection, intersection_area_3d,
                       intersection_area_xy, is_complex_shape, distance)
    """
    try:
        direct_intersection = check_direct_intersection(room_id, ceiling_id, prefix, timestamp)
        xy_intersection = project_and_check_xy_intersection(room_id, ceiling_id) if not direct_intersection else True
        distance = delta_ceiling_above_room(room_id, ceiling_id)
        return direct_intersection, xy_intersection, distance
    except Exception as e:
        logging.exception(f"Error in intersection checks: {e}")
        return None

# ========================
# MAIN EXECUTION
# ========================

def main():
    """Main execution function."""
    doc = __revit__.ActiveUIDocument.Document
    
    # Collect all rooms and ceilings
    rooms = FilteredElementCollector(doc).OfCategory(
        BuiltInCategory.OST_Rooms
    ).WhereElementIsNotElementType().ToElements()
    
    ceilings = FilteredElementCollector(doc).OfCategory(
        BuiltInCategory.OST_Ceilings
    ).WhereElementIsNotElementType().ToElements()
    
    t = Transaction(doc, "Assign VDC Room Number to Ceilings")
    t.Start()
    
    try:
        for ceiling in ceilings:
            mapped_rooms = set()
            
            for room in rooms:
                intersection_result = check_intersections(
                    room.Id, ceiling.Id
                )
                
                if intersection_result is None:
                    continue
                    
                direct_intersection, xy_intersection, distance = intersection_result
                
                if direct_intersection or (xy_intersection and distance > 0):
                    room_number = room.LookupParameter("Number").AsString()
                    mapped_rooms.add(room_number)
            
            # Assign room numbers to ceiling
            room_number_param = ceiling.LookupParameter("VDC Room Number")
            if room_number_param and not room_number_param.IsReadOnly:
                value_to_set = ",".join(sorted(mapped_rooms)) if mapped_rooms else "None"
                room_number_param.Set(value_to_set)
        
        t.Commit()
        print("Room numbers assigned to ceilings successfully.")
    except Exception as e:
        t.RollBack()
        print(f"Error: {str(e)}")

if __name__ == "__main__":
    main()
