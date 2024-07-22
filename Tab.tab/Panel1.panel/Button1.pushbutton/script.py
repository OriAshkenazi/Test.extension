#! python3

import clr
clr.AddReference('RevitAPI')
clr.AddReference('System')
from Autodesk.Revit.DB import *
import pandas as pd

# Get the current document
doc = __revit__.ActiveUIDocument.Document

# Initialize debug messages
debug_messages = []

def get_element_geometry(element):
    """
    Get the geometry of a Revit element.
    """
    try:
        geom = element.get_Geometry(Options())
        if geom:
            solids = [solid for solid in geom if isinstance(solid, Solid) and solid.Volume > 0]
            return solids if solids else None
        return None
    except Exception as e:
        debug_messages.append(f"Error in get_element_geometry: {e}")
        return None

def project_to_xy_plane(solid):
    """
    Project solid geometry to the XY plane.
    """
    return [face for face in solid.Faces if abs(face.ComputeNormal(UV(0.5, 0.5)).Z) > 0.99]

def get_intersecting_area(solid1, solid2):
    """
    Calculate the intersecting area between two solids.
    """
    try:
        intersection = BooleanOperationsUtils.ExecuteBooleanOperation(solid1, solid2, BooleanOperationsType.Intersect)
        if intersection and isinstance(intersection, Solid):
            return sum(face.Area for face in project_to_xy_plane(intersection))
    except Exception as e:
        debug_messages.append(f"Error in get_intersecting_area: {e}")
    return 0

def calculate_intersection_percentage(geom1, geom2):
    """
    Calculate the intersection percentage between two geometries.
    """
    area1 = sum(face.Area for solid in geom1 for face in project_to_xy_plane(solid))
    area2 = sum(face.Area for solid in geom2 for face in project_to_xy_plane(solid))
    intersection_area = get_intersecting_area(geom1[0], geom2[0]) if geom1 and geom2 else 0
    percentage1 = (intersection_area / area1) * 100 if area1 > 0 else 0
    percentage2 = (intersection_area / area2) * 100 if area2 > 0 else 0
    return percentage1, percentage2, intersection_area

def find_ceiling_room_relationships(room_elements, ceiling_elements):
    """
    Find the relationship between ceilings and rooms in terms of intersection percentage.
    """
    relationships = []

    for room in room_elements:
        room_id = room.Id.IntegerValue
        room_geom = get_element_geometry(room)
        if not room_geom:
            debug_messages.append(f"No geometry found for room ID: {room_id}")
            continue

        room_level = room.LevelId

        for ceiling in ceiling_elements:
            ceiling_id = ceiling.Id.IntegerValue
            ceiling_geom = get_element_geometry(ceiling)
            if not ceiling_geom:
                debug_messages.append(f"No geometry found for ceiling ID: {ceiling_id}")
                continue

            if ceiling.LevelId != room_level:
                debug_messages.append(f"Ceiling ID {ceiling_id} is not on the same level as Room ID {room_id}")
                continue

            intersection_percentage_room, intersection_percentage_ceiling, intersection_area = calculate_intersection_percentage(room_geom, ceiling_geom)
            if intersection_percentage_room > 0:
                relationships.append({
                    'Room_ID': room_id,
                    'Ceiling_ID': ceiling_id,
                    'Intersection_Percentage_Room': intersection_percentage_room,
                    'Intersection_Percentage_Ceiling': intersection_percentage_ceiling,
                    'Intersection_Area': intersection_area
                })

    return pd.DataFrame(relationships)

# Main script execution
try:
    # Inputs
    room_elements = FilteredElementCollector(doc).OfClass(SpatialElement).ToElements()
    ceiling_elements = FilteredElementCollector(doc).OfClass(Ceiling).ToElements()

    # Find the relationships between rooms and ceilings
    df_relationships = find_ceiling_room_relationships(room_elements, ceiling_elements)

    # Output the dataframe
    print(df_relationships)

except Exception as e:
    debug_messages.append(f"Error in main execution: {e}")
    print(debug_messages)
