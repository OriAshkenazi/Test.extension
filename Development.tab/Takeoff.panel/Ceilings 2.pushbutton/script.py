#! python3

import clr
from functools import lru_cache
clr.AddReference('RevitAPI')
clr.AddReference('System')
from Autodesk.Revit.DB import *
from Autodesk.Revit.DB import BooleanOperationsUtils, BooleanOperationsType
from Autodesk.Revit.Exceptions import InvalidOperationException
import pandas as pd
import datetime
from shapely.geometry import Polygon, MultiPolygon
from shapely.ops import unary_union
from shapely.validation import make_valid

# Get the current document
doc = __revit__.ActiveUIDocument.Document

# Initialize debug messages
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

@lru_cache(maxsize=None)
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

@lru_cache(maxsize=None)
def calculate_intersection_area(geom1_id, geom2_id):
    """
    Calculate the intersecting area between two geometries in sqm.
    
    Args:
        geom1_id (ElementId): The ID of the first Revit element.
        geom2_id (ElementId): The ID of the second Revit element.
    
    Returns:
        float: The intersecting area in square meters.
    """
    geom1 = get_element_geometry(geom1_id)
    geom2 = get_element_geometry(geom2_id)

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
                intersection_area = sum(p.area for p in intersection.geoms) * 0.092903  # Handle MultiPolygon
            else:
                intersection_area = 0  # Handle other geometry types or empty intersections
    except Exception as e:
        debug_messages.append(f"Error in calculate_intersection_area: {e}")
        intersection_area = 0

    return intersection_area

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
                if intersection_solid.Volume > 0:
                    return True
            except InvalidOperationException:
                # If Boolean operation fails, fall back to bounding box check
                room_bb = room_solid.GetBoundingBox()
                ceiling_bb = ceiling_solid.GetBoundingBox()
                if check_bounding_box_intersection(room_bb, ceiling_bb):
                    return True
    return False

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
    
    room_union = unary_union(room_polygons)
    ceiling_union = unary_union(ceiling_polygons)
    
    return room_union.intersects(ceiling_union)

def is_ceiling_above_room(room_id, ceiling_id):
    """
    Check if the ceiling is above the room.
    
    Args:
        room_id (ElementId): The ID of the room element.
        ceiling_id (ElementId): The ID of the ceiling element.
    
    Returns:
        bool: True if the ceiling is above the room, False otherwise.
    """
    room_bb = get_element_bounding_box(room_id)
    ceiling_bb = get_element_bounding_box(ceiling_id)
    return ceiling_bb.Min.Z >= room_bb.Max.Z

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
    Find the relationship between ceilings and rooms based on the two vectors.
    
    Args:
        room_elements (list): List of room elements.
        ceiling_elements (list): List of ceiling elements.
    
    Returns:
        tuple: Two pandas DataFrames - one for related ceilings and rooms, one for unrelated ceilings.
    """
    relationships = []
    unrelated_ceilings = []
    room_ids = [room.Id for room in room_elements]

    for ceiling in ceiling_elements:
        ceiling_id = ceiling.Id
        ceiling_details = get_ceiling_details(ceiling)
        ceiling_bb = get_element_bounding_box(ceiling_id)
        
        if not get_element_geometry(ceiling_id):
            debug_messages.append(f"No geometry found for ceiling ID: {ceiling_id.IntegerValue}")
            unrelated_ceilings.append(ceiling_details)
            continue
        
        matched_rooms = []
        
        for room_id in room_ids:
            room = doc.GetElement(room_id)
            room_details = get_room_details(room)
            room_bb = get_element_bounding_box(room_id)
            
            if not get_element_geometry(room_id):
                debug_messages.append(f"No geometry found for room ID: {room_id.IntegerValue}")
                continue
            
            # Check for direct intersection
            try:
                direct_intersection = check_direct_intersection(room_id, ceiling_id)
            except Exception as e:
                debug_messages.append(f"Error in direct intersection check: {e}")
                direct_intersection = check_bounding_box_intersection(room_id, ceiling_id)
            
            # Check for XY projection intersection
            xy_intersection = project_and_check_xy_intersection(room_id, ceiling_id)
            
            if direct_intersection or xy_intersection:
                intersection_area = calculate_intersection_area(room_id, ceiling_id)
                matched_rooms.append({
                    'Ceiling_ID': ceiling_details[0],
                    'Ceiling_Type': ceiling_details[1],
                    'Ceiling_Description': ceiling_details[2],
                    'Ceiling_Area_sqm': ceiling_details[3],
                    'Ceiling_Level': ceiling_details[4],
                    'Room_ID': room_details[0],
                    'Room_Name': room_details[1],
                    'Room_Number': room_details[2],
                    'Room_Level': room_details[3],
                    'Room_Building': room_details[4],
                    'Intersection_Area_sqm': intersection_area,
                    'Direct_Intersection': direct_intersection,
                    'XY_Projection_Intersection': xy_intersection
                })
        
        if matched_rooms:
            relationships.extend(matched_rooms)
        else:
            unrelated_ceilings.append(ceiling_details)

    df_relationships = pd.DataFrame(relationships)
    df_unrelated = pd.DataFrame(unrelated_ceilings, columns=['Ceiling_ID', 'Ceiling_Type', 'Ceiling_Description', 'Ceiling_Area_sqm', 'Ceiling_Level'])
    
    # Sort the relationships DataFrame
    df_relationships = df_relationships.sort_values(by=['Room_Building', 'Room_Level', 'Room_Number'])
    
    # Sort the unrelated ceilings DataFrame
    df_unrelated = df_unrelated.sort_values(by=['Ceiling_Level', 'Ceiling_Type'])

    return df_relationships, df_unrelated

def find_rooms_without_ceilings(df_relationships, room_elements):
    """
    Identify rooms that don't have associated ceilings.
    
    Args:
        df_relationships (pandas.DataFrame): DataFrame containing ceiling-room relationships.
        room_elements (list): List of all room elements.
    
    Returns:
        pandas.DataFrame: DataFrame containing rooms without ceilings.
    """
    # Get all room IDs from the relationships DataFrame
    rooms_with_ceilings = set(df_relationships['Room_ID'])
    
    # Identify rooms without ceilings
    rooms_without_ceilings = []
    for room in room_elements:
        room_id = room.Id.IntegerValue
        if room_id not in rooms_with_ceilings:
            room_details = get_room_details(room)
            rooms_without_ceilings.append({
                'Room_ID': room_details[0],
                'Room_Name': room_details[1],
                'Room_Number': room_details[2],
                'Room_Level': room_details[3],
                'Room_Building': room_details[4]
            })
    
    df_rooms_without_ceilings = pd.DataFrame(rooms_without_ceilings)
    
    # Sort the DataFrame
    df_rooms_without_ceilings = df_rooms_without_ceilings.sort_values(by=['Room_Building', 'Room_Level', 'Room_Number'])
    
    return df_rooms_without_ceilings

def pivot_data(df_relationships):
    """
    Pivot the relationships data around rooms.
    
    Args:
        df_relationships (pandas.DataFrame): DataFrame containing ceiling-room relationships.
    
    Returns:
        pandas.DataFrame: Pivoted DataFrame with rooms as index and ceiling information as columns.
    """
    pivoted = df_relationships.pivot_table(
        index=['Room_Building', 'Room_Level', 'Room_Number', 'Room_Name', 'Room_ID'],
        values=['Ceiling_ID', 'Ceiling_Type', 'Ceiling_Description', 'Ceiling_Area_sqm', 'Intersection_Area_sqm'],
        aggfunc='first'
    ).reset_index()
    
    # Sort the pivoted DataFrame
    pivoted = pivoted.sort_values(by=['Room_Building', 'Room_Level', 'Room_Number'])
    
    return pivoted

def main():
    """
    Main function to execute the script.
    """
    try:
        # Collect room and ceiling elements
        room_elements = FilteredElementCollector(doc).OfClass(SpatialElement).OfCategory(BuiltInCategory.OST_Rooms).ToElements()
        ceiling_elements = FilteredElementCollector(doc).OfClass(Ceiling).ToElements()

        # Find the relationships between ceilings and rooms
        df_relationships, df_unrelated = find_ceiling_room_relationships(room_elements, ceiling_elements)

        # Pivot the relationships data
        df_pivoted = pivot_data(df_relationships)

        # Find rooms without ceilings
        df_rooms_without_ceilings = find_rooms_without_ceilings(df_relationships, room_elements)

        # Output the dataframe with timestamp and formatted Excel
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        output_file_path = f"C:\\Mac\\Home\\Documents\\Shapir\\Exports\\ceiling_room_relationships_{timestamp}.xlsx"

        # Export to Excel with formatting
        with pd.ExcelWriter(output_file_path, engine='xlsxwriter') as writer:
            df_pivoted.to_excel(writer, sheet_name='Ceiling-Room Relationships', index=False)
            df_unrelated.to_excel(writer, sheet_name='Unrelated Ceilings', index=False)
            
            workbook = writer.book
            pivot_worksheet = writer.sheets['Ceiling-Room Relationships']
            unrelated_worksheet = writer.sheets['Unrelated Ceilings']
            
            # Apply header format
            header_format = workbook.add_format({
                'bold': True,
                'text_wrap': True,
                'valign': 'top',
                'fg_color': '#D7E4BC',
                'border': 1
            })
            
            for worksheet, df in [(pivot_worksheet, df_pivoted), (unrelated_worksheet, df_unrelated)]:
                for col_num, value in enumerate(df.columns):
                    worksheet.write(0, col_num, value, header_format)
                    worksheet.set_column(col_num, col_num, 20)  # Set column width

        print(f"Schedule saved to {output_file_path}")
    
    except Exception as e:
        debug_messages.append(f"Error in main function: {e}")
    
    print("Debug messages:")
    for msg in debug_messages:
        print(msg)

# Call the main function
if __name__ == "__main__":
    main()