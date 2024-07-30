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
    
    Args:
        element (Element): The Revit element.
    
    Returns:
        List[Solid] or None: The list of solids representing the geometry of the element, or None if no geometry is found.
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
    
    Args:
        solid (Solid): The solid geometry.
    
    Returns:
        Polygon or None: The projected polygon on the XY plane, or None if no valid polygon is found.
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
    
    Args:
        geom1 (List[Solid]): The list of solids representing the first geometry.
        geom2 (List[Solid]): The list of solids representing the second geometry.
    
    Returns:
        float: The intersecting area in square meters.
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
            if isinstance(intersection, Polygon):
                intersection_area = intersection.area * 0.092903  # Convert from square feet to square meters
            elif isinstance(intersection, MultiPolygon):
                intersection_area = sum(p.area for p in intersection) * 0.092903  # Convert from square feet to square meters
    except Exception as e:
        debug_messages.append(f"Error in calculate_intersection_area: {e}")
        intersection_area = 0

    return intersection_area

def get_room_details(room):
    """
    Get details of a room.
    
    Args:
        room (SpatialElement): The room element.
    
    Returns:
        Tuple[int, str, str, str, str]: The room details (room_id, room_name, room_number, room_level, room_building).
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
    
    Args:
        ceiling (Ceiling): The ceiling element.
    
    Returns:
        Tuple[int, str, str, float, str]: The ceiling details (ceiling_id, ceiling_type, ceiling_description, ceiling_area, ceiling_level).
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
    
    Args:
        room_elements (list): List of room elements.
        ceiling_elements (list): List of ceiling elements.
    
    Returns:
        pandas.DataFrame: DataFrame containing the relationship between ceilings and rooms.
    """
    relationships = []

    for ceiling in ceiling_elements:
        ceiling_id, ceiling_type, ceiling_description, ceiling_area, ceiling_level = get_ceiling_details(ceiling)
        ceiling_geom = get_element_geometry(ceiling)
        if not ceiling_geom:
            debug_messages.append(f"No geometry found for ceiling ID: {ceiling_id}")
            continue
        
        ceiling_has_intersection = False
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
                ceiling_has_intersection = True

        if not ceiling_has_intersection:
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

def pivot_data(df):
    """
    Pivot the data based on intersecting and non-intersecting ceilings.

    Args:
        df (pandas.DataFrame): The input DataFrame containing ceiling data.

    Returns:
        tuple: A tuple containing two DataFrames. The first DataFrame contains the pivoted data
               for intersecting ceilings, and the second DataFrame contains the non-intersecting ceilings.
    """
    # Separate intersecting and non-intersecting ceilings
    intersecting_df = df[df['Intersection_Area_sqm'] > 0]
    non_intersecting_df = df[df['Intersection_Area_sqm'] == 0]

    # Pivot intersecting data
    pivot_df = intersecting_df.pivot_table(index=['Room_Building', 'Room_Level', 'Room_Name', 'Room_Number', 'Room_ID'],
                                           values=['Ceiling_ID', 'Ceiling_Type', 'Ceiling_Description', 'Ceiling_Area_sqm', 'Intersection_Area_sqm'],
                                           aggfunc='first').reset_index()

    # Sort pivoted data
    pivot_df.sort_values(by=['Room_Building', 'Room_Level', 'Room_Name'], inplace=True)

    return pivot_df, non_intersecting_df

# Main script execution

# Collect room and ceiling elements
room_elements = FilteredElementCollector(doc).OfClass(SpatialElement).OfCategory(BuiltInCategory.OST_Rooms).ToElements()
ceiling_elements = FilteredElementCollector(doc).OfClass(Ceiling).ToElements()

# Find the relationships between ceilings and rooms
df_relationships = find_ceiling_room_relationships(room_elements, ceiling_elements)

# Normalize and clean data
df_relationships.fillna('', inplace=True)

# Pivot and sort data
pivot_df, non_intersecting_df = pivot_data(df_relationships)

# Ensure non-intersecting DataFrame has correct columns
non_intersecting_df = non_intersecting_df[['Ceiling_ID', 'Ceiling_Type', 'Ceiling_Description', 'Ceiling_Area_sqm', 'Ceiling_Level', 'Room_ID', 'Room_Name', 'Room_Number', 'Room_Level', 'Room_Building', 'Intersection_Area_sqm']]

# Output the dataframe with timestamp and formatted Excel
timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
output_file_path = f"C:\\Mac\\Home\\Documents\\Shapir\\Exports\\ceiling_room_relationships_{timestamp}.xlsx"

# Export to Excel with formatting
with pd.ExcelWriter(output_file_path, engine='xlsxwriter') as writer:
    pivot_df.to_excel(writer, sheet_name='Ceiling-Room Relationships', index=False)
    non_intersecting_df.to_excel(writer, sheet_name='Non-Intersecting Ceilings', index=False)
    
    workbook = writer.book
    pivot_worksheet = writer.sheets['Ceiling-Room Relationships']
    non_intersecting_worksheet = writer.sheets['Non-Intersecting Ceilings']
    
    # Apply header format
    header_format = workbook.add_format({
        'bold': True,
        'text_wrap': True,
        'valign': 'top',
        'fg_color': '#D7E4BC',
        'border': 1
    })
    
    for worksheet in [pivot_worksheet, non_intersecting_worksheet]:
        for col_num, value in enumerate(pivot_df.columns.values):
            worksheet.write(0, col_num, value, header_format)
            worksheet.set_column(col_num, col_num, 20)  # Set column width

print(f"Schedule saved to {output_file_path}")
print(debug_messages)
