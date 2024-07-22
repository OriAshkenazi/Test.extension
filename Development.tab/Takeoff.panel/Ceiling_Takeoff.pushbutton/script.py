#! python3

import clr
clr.AddReference('RevitAPI')
clr.AddReference('System')
from Autodesk.Revit.DB import *
import pandas as pd
import datetime
from shapely.geometry import Polygon, MultiPolygon
from shapely.ops import unary_union
from shapely.validation import make_valid

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
    Project the lower horizontal face of solid geometry to the XY plane.
    """
    for face in solid.Faces:
        normal = face.ComputeNormal(UV(0.5, 0.5))
        if abs(normal.Z) > 0.99 and normal.Z < 0:  # Check if the face is horizontal and facing down
            edge_loops = face.GetEdgesAsCurveLoops()
            for loop in edge_loops:
                vertices = []
                for curve in loop:
                    for point in curve.Tessellate():
                        vertices.append((point.X, point.Y))
                polygon = Polygon(vertices)
                if not polygon.is_valid:
                    polygon = make_valid(polygon)
                return polygon
    return None

def calculate_intersection_area(geom1, geom2):
    """
    Calculate the intersecting area between two geometries in sqm.
    """
    intersection_area = 0

    try:
        polygons1 = []
        for solid in geom1:
            polygon = project_to_xy_plane(solid)
            if polygon:
                polygons1.append(polygon)
        
        polygons2 = []
        for solid in geom2:
            polygon = project_to_xy_plane(solid)
            if polygon:
                polygons2.append(polygon)

        if polygons1 and polygons2:
            union1 = unary_union(polygons1)
            union2 = unary_union(polygons2)
            intersection = union1.intersection(union2)
            if isinstance(intersection, (Polygon, MultiPolygon)):
                if isinstance(intersection, Polygon):
                    intersection = MultiPolygon([intersection])
                intersection_area = sum(p.area for p in intersection) * 0.092903  # Convert from square feet to square meters
            elif isinstance(intersection, MultiPolygon):
                intersection_area = sum(p.area for p in intersection) * 0.092903  # Convert from square feet to square meters
    except Exception as e:
        debug_messages.append(f"Error in calculate_intersection_area: {e}")
        intersection_area = 0

    return intersection_area

def get_room_details(room):
    """
    Get details of a room.
    """
    room_id = room.Id.IntegerValue
    room_name = room.get_Parameter(BuiltInParameter.ROOM_NAME).AsString()
    room_number = room.get_Parameter(BuiltInParameter.ROOM_NUMBER).AsString()
    room_level = doc.GetElement(room.LevelId).Name
    room_building_param = room.LookupParameter("בניין")
    room_building = room_building_param.AsString() if room_building_param else None
    return room_id, room_name, room_number, room_level, room_building

def get_ceiling_details(ceiling):
    """
    Get details of a ceiling.
    """
    ceiling_id = ceiling.Id.IntegerValue
    ceiling_type_element = doc.GetElement(ceiling.GetTypeId())
    ceiling_type = ceiling_type_element.FamilyName
    
    ceiling_description_param = ceiling_type_element.LookupParameter("Description")
    ceiling_description = ceiling_description_param.AsString() if ceiling_description_param else None
    
    ceiling_area_param = ceiling.LookupParameter("Area")
    ceiling_area = ceiling_area_param.AsDouble() * 0.092903 if ceiling_area_param else None # Convert from square feet to square meters

    ceiling_level = doc.GetElement(ceiling.LevelId).Name if ceiling.LevelId else None

    
    return ceiling_id, ceiling_type, ceiling_description, ceiling_area, ceiling_level

def find_ceiling_room_relationships(room_elements, ceiling_elements):
    """
    Find the relationship between ceilings and rooms in terms of intersection area.
    """
    relationships = []

    for ceiling in ceiling_elements:
        ceiling_id, ceiling_type, ceiling_description, ceiling_area, ceiling_level = get_ceiling_details(ceiling)
        ceiling_geom = get_element_geometry(ceiling)
        if not ceiling_geom:
            debug_messages.append(f"No geometry found for ceiling ID: {ceiling_id}")
            continue
        
        for room in room_elements:
            room_id, room_name, room_number, room_level, room_building = get_room_details(room)
            room_geom = get_element_geometry(room)
            if not room_geom:
                debug_messages.append(f"No geometry found for room ID: {room_id}")
                continue

            if ceiling.LevelId != room.LevelId:
                continue

            intersection_area = calculate_intersection_area(room_geom, ceiling_geom)
            if intersection_area > 0:
                relationships.append({
                    'Ceiling_ID': ceiling_id,
                    'Ceiling_Type': ceiling_type,
                    'Ceiling_Description': ceiling_description,
                    'Ceiling_Area_sqm': ceiling_area,
                    'Ceiling_Level': ceiling_level,
                    'Room_ID': room_id,
                    'Room_Name': room_name,
                    'Room_Number': room_number,
                    'Room_Level': room_level,
                    'Room_Building': room_building,
                    'Intersection_Area_sqm': intersection_area
                })

        if ceiling_id not in [r['Ceiling_ID'] for r in relationships]:
            relationships.append({
                'Ceiling_ID': ceiling_id,
                'Ceiling_Type': ceiling_type,
                'Ceiling_Description': ceiling_description,
                'Ceiling_Area_sqm': ceiling_area,
                'Ceiling_Level': ceiling_level,
                'Room_ID': None,
                'Room_Name': None,
                'Room_Number': None,
                'Room_Level': None,
                'Room_Building': None,
                'Intersection_Area_sqm': 0
            })

    return pd.DataFrame(relationships)

# Main script execution

# Collect room and ceiling elements
room_elements = FilteredElementCollector(doc).OfClass(SpatialElement).OfCategory(BuiltInCategory.OST_Rooms).ToElements()
ceiling_elements = FilteredElementCollector(doc).OfClass(Ceiling).ToElements()

# Find the relationships between ceilings and rooms
df_relationships = find_ceiling_room_relationships(room_elements, ceiling_elements)

# Output the dataframe with timestamp
timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
output_file_path = f"C:\\Users\\oriashkenazi\\Exports\\ceiling_room_relationships_{timestamp}.xlsx"
df_relationships.to_excel(output_file_path, index=False)
print(f"Schedule saved to {output_file_path}")
print(debug_messages)
