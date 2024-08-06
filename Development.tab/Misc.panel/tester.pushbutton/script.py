#! python3

import clr
from functools import lru_cache
import pandas as pd
import datetime
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from shapely.geometry import Polygon, MultiPolygon
from shapely.ops import unary_union
from shapely.validation import make_valid

clr.AddReference('RevitAPI')
clr.AddReference('System')
from Autodesk.Revit.DB import *
from Autodesk.Revit.Exceptions import InvalidOperationException

# Get the current document
doc = __revit__.ActiveUIDocument.Document

debug_messages = []

@lru_cache(maxsize=None)
def get_element_geometry(element_id):
    """
    Memoized function to get the geometry of a Revit element.
    
    Args:
        element_id (ElementId): The ID of the Revit element.
    
    Returns:
        tuple or None: A tuple of Solid objects representing the element's geometry, or None if no valid geometry is found.
    """
    element = doc.GetElement(element_id)
    try:
        geom = element.get_Geometry(Options())
        if geom:
            solids = [solid for solid in geom if isinstance(solid, Solid) and solid.Volume > 0]
            return tuple(solids) if solids else None
        return None
    except Exception as e:
        debug_messages.append(f"Error in get_element_geometry: {e}")
        return None

@lru_cache(maxsize=None)
def get_element_bounding_box(element_id):
    """
    Memoized function to get the bounding box of an element.
    
    Args:
        element_id (ElementId): The ID of the Revit element.
    
    Returns:
        BoundingBoxXYZ: The bounding box of the element.
    """
    element = doc.GetElement(element_id)
    return element.get_BoundingBox(None)

def check_bounding_box_intersection(bb1, bb2):
    """
    Check if two bounding boxes intersect.
    
    Args:
        bb1 (BoundingBoxXYZ): First bounding box
        bb2 (BoundingBoxXYZ): Second bounding box
    
    Returns:
        bool: True if the bounding boxes intersect, False otherwise.
    """
    return (bb1.Min.X <= bb2.Max.X and bb1.Max.X >= bb2.Min.X and
            bb1.Min.Y <= bb2.Max.Y and bb1.Max.Y >= bb2.Min.Y and
            bb1.Min.Z <= bb2.Max.Z and bb1.Max.Z >= bb2.Min.Z)

def check_direct_intersection(room_id, ceiling_id):
    """
    Check if there's a direct intersection between room and ceiling geometries.
    
    Args:
        room_id (ElementId): The ID of the room element.
        ceiling_id (ElementId): The ID of the ceiling element.
    
    Returns:
        bool: True if there's a direct intersection, False otherwise.
    """
    room_geom = get_element_geometry(room_id)
    ceiling_geom = get_element_geometry(ceiling_id)
    for room_solid in room_geom:
        for ceiling_solid in ceiling_geom:
            try:
                intersection_solid = BooleanOperationsUtils.ExecuteBooleanOperation(
                    room_solid, ceiling_solid, BooleanOperationsType.Intersect)
                print(f"area: {intersection_solid.SurfaceArea}")
                print(f"volume: {intersection_solid.Volume}")
                if intersection_solid.Volume > 0:
                    return True
            except InvalidOperationException:
                # If Boolean operation fails, fall back to bounding box check
                print("bounding box check")
                room_bb = room_solid.GetBoundingBox()
                ceiling_bb = ceiling_solid.GetBoundingBox()
                if check_bounding_box_intersection(room_bb, ceiling_bb):
                    return True
    return False

def project_to_xy_plane(solid):
    """
    Project the entire solid geometry to the XY plane.
    
    Args:
        solid (Solid): The solid geometry.
    
    Returns:
        Polygon or None: The projected polygon on the XY plane, or None if no valid polygon is found.
    """
    vertices = []
    for face in solid.Faces:
        edge_loops = face.GetEdgesAsCurveLoops()
        for loop in edge_loops:
            for curve in loop:
                for point in curve.Tessellate():
                    vertices.append((point.X, point.Y))
    
    if vertices:
        polygon = Polygon(vertices)
        if not polygon.is_valid:
            polygon = make_valid(polygon)
        return polygon
    return None

def project_and_check_xy_intersection(room_id, ceiling_id):
    """
    Project geometries to XY plane and check for intersection.
    
    Args:
        room_id (ElementId): The ID of the room element.
        ceiling_id (ElementId): The ID of the ceiling element.
    
    Returns:
        bool: True if there's an intersection in the XY projection, False otherwise.
    """
    room_geom = get_element_geometry(room_id)
    ceiling_geom = get_element_geometry(ceiling_id)
    
    room_polygons = [project_to_xy_plane(solid) for solid in room_geom if solid]
    ceiling_polygons = [project_to_xy_plane(solid) for solid in ceiling_geom if solid]
    
    room_polygons = [p for p in room_polygons if p is not None]
    ceiling_polygons = [p for p in ceiling_polygons if p is not None]
    
    if not room_polygons or not ceiling_polygons:
        return False
    
    room_union = unary_union(room_polygons)
    ceiling_union = unary_union(ceiling_polygons)
    
    return room_union.intersects(ceiling_union)

@lru_cache(maxsize=None)
def calculate_intersection_areas(geom1_id, geom2_id):
    geom1 = get_element_geometry(geom1_id)
    geom2 = get_element_geometry(geom2_id)
    intersection_area_3d = 0
    intersection_area_xy = 0

    try:
        # Check for direct 3D intersection
        for solid1 in geom1:
            for solid2 in geom2:
                try:
                    intersection_solid = BooleanOperationsUtils.ExecuteBooleanOperation(
                        solid1, solid2, BooleanOperationsType.Intersect)
                    print(f"3D intersection volume: {intersection_solid.Volume}")
                    if intersection_solid.Volume > 0:
                        largest_face_area = max(face.Area for face in intersection_solid.Faces)
                        intersection_area_3d = max(intersection_area_3d, largest_face_area * 0.092903)
                        print(f"3D intersection area: {intersection_area_3d}")
                except InvalidOperationException:
                    print("Invalid operation in 3D intersection")

        # Calculate XY projection intersection
        polygons1 = [project_to_xy_plane(solid) for solid in geom1 if solid]
        polygons2 = [project_to_xy_plane(solid) for solid in geom2 if solid]
        
        polygons1 = [p for p in polygons1 if p is not None]
        polygons2 = [p for p in polygons2 if p is not None]
        
        print(f"Number of polygons1: {len(polygons1)}")
        print(f"Number of polygons2: {len(polygons2)}")
        
        if polygons1 and polygons2:
            union1 = unary_union(polygons1)
            union2 = unary_union(polygons2)
            
            print(f"Union1 bounds: {union1.bounds}")
            print(f"Union2 bounds: {union2.bounds}")
            print(f"Intersects: {union1.intersects(union2)}")
            
            if union1.intersects(union2):
                intersection = union1.intersection(union2)
                intersection_area_xy = intersection.area * 0.092903
                print(f"XY intersection area: {intersection_area_xy}")

        # Bounding box check
        if intersection_area_xy == 0:
            bb1 = get_element_bounding_box(geom1_id)
            bb2 = get_element_bounding_box(geom2_id)
            if check_bounding_box_intersection(bb1, bb2):
                x_overlap = min(bb1.Max.X, bb2.Max.X) - max(bb1.Min.X, bb2.Min.X)
                y_overlap = min(bb1.Max.Y, bb2.Max.Y) - max(bb1.Min.Y, bb2.Min.Y)
                intersection_area_xy = x_overlap * y_overlap * 0.092903
                print(f"Bounding box intersection area: {intersection_area_xy}")

    except Exception as e:
        print(f"Error in calculate_intersection_areas: {e}")

    return intersection_area_3d, intersection_area_xy

from shapely.geometry import Polygon, MultiPolygon

@lru_cache(maxsize=None)
def calculate_intersection_areas_new(geom1_id, geom2_id):
    geom1 = get_element_geometry(geom1_id)
    geom2 = get_element_geometry(geom2_id)
    intersection_area_3d = 0
    intersection_area_xy = 0

    try:
        # Check for direct 3D intersection
        for solid1 in geom1:
            for solid2 in geom2:
                try:
                    intersection_solid = BooleanOperationsUtils.ExecuteBooleanOperation(
                        solid1, solid2, BooleanOperationsType.Intersect)
                    if intersection_solid.Volume > 0:
                        # Project the intersection solid to XY plane and use envelope
                        intersection_polygon = project_to_xy_plane(intersection_solid)
                        if intersection_polygon:
                            intersection_envelope = intersection_polygon.envelope
                            intersection_area_3d = max(intersection_area_3d, intersection_envelope.area * 0.092903)
                except InvalidOperationException:
                    pass  # If Boolean operation fails, continue to next check

        # Calculate XY projection intersection
        polygons1 = [project_to_xy_plane(solid) for solid in geom1 if solid]
        polygons2 = [project_to_xy_plane(solid) for solid in geom2 if solid]
        
        polygons1 = [p for p in polygons1 if p is not None]
        polygons2 = [p for p in polygons2 if p is not None]
        
        if polygons1 and polygons2:
            # Create a single polygon for each geometry, filling in any holes
            union1 = unary_union(polygons1).envelope
            union2 = unary_union(polygons2).envelope
            
            if union1.intersects(union2):
                intersection = union1.intersection(union2)
                if isinstance(intersection, (Polygon, MultiPolygon)):
                    intersection_area_xy = intersection.area * 0.092903
                else:
                    # Handle cases where the intersection might be a line or point
                    intersection_area_xy = 0

        # Bounding box check (as a fallback)
        if intersection_area_xy == 0:
            bb1 = get_element_bounding_box(geom1_id)
            bb2 = get_element_bounding_box(geom2_id)
            if check_bounding_box_intersection(bb1, bb2):
                x_overlap = min(bb1.Max.X, bb2.Max.X) - max(bb1.Min.X, bb2.Min.X)
                y_overlap = min(bb1.Max.Y, bb2.Max.Y) - max(bb1.Min.Y, bb2.Min.Y)
                intersection_area_xy = x_overlap * y_overlap * 0.092903

    except Exception as e:
        debug_messages.append(f"Error in calculate_intersection_areas: {e}")

    return intersection_area_3d, intersection_area_xy

room_id = ElementId(int("8087401"))
ceiling_id = ElementId(int("11327031"))

room_id_new = ElementId(int("9497352"))
ceiling_id_new = ElementId(int("10481073"))

print(project_and_check_xy_intersection(room_id, ceiling_id))
print(calculate_intersection_areas(room_id, ceiling_id))

print(project_and_check_xy_intersection(room_id_new, ceiling_id_new))
print(calculate_intersection_areas_new(room_id_new, ceiling_id_new))