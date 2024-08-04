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
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment

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
    geom1 = get_element_geometry(geom1_id)
    geom2 = get_element_geometry(geom2_id)

    intersection_area = 0

    try:
        # Check for direct 3D intersection
        for solid1 in geom1:
            for solid2 in geom2:
                try:
                    intersection_solid = BooleanOperationsUtils.ExecuteBooleanOperation(
                        solid1, solid2, BooleanOperationsType.Intersect)
                    if intersection_solid.Volume > 0:
                        # Calculate the area of the intersection solid's largest face
                        largest_face_area = max(face.Area for face in intersection_solid.Faces)
                        intersection_area = max(intersection_area, largest_face_area * 0.092903)  # Convert to sqm
                except InvalidOperationException:
                    pass  # If Boolean operation fails, continue to next check

        # If no direct intersection, check XY projection
        if intersection_area == 0:
            polygons1 = [project_to_xy_plane(solid) for solid in geom1 if solid]
            polygons2 = [project_to_xy_plane(solid) for solid in geom2 if solid]
            
            if polygons1 and polygons2:
                union1 = unary_union(polygons1)
                union2 = unary_union(polygons2)
                intersection = union1.intersection(union2)
                if isinstance(intersection, (Polygon, MultiPolygon)):
                    intersection_area = intersection.area * 0.092903  # Convert to sqm

        # If still no intersection, check bounding box intersection
        if intersection_area == 0:
            bb1 = get_element_bounding_box(geom1_id)
            bb2 = get_element_bounding_box(geom2_id)
            if check_bounding_box_intersection(bb1, bb2):
                # Calculate the area of overlap in XY plane
                x_overlap = min(bb1.Max.X, bb2.Max.X) - max(bb1.Min.X, bb2.Min.X)
                y_overlap = min(bb1.Max.Y, bb2.Max.Y) - max(bb1.Min.Y, bb2.Min.Y)
                intersection_area = x_overlap * y_overlap * 0.092903  # Convert to sqm

    except Exception as e:
        debug_messages.append(f"Error in calculate_intersection_area: {e}")

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

def custom_sort_key(value):
    """
    Custom sorting function to handle numeric strings as numbers and None values.
    """
    if value is None:
        return float('inf')  # This will put None values at the end of the sort order
    try:
        return float(value)
    except ValueError:
        return value

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
    no_geometry_rooms = []

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
                if room_id not in no_geometry_rooms:
                    no_geometry_rooms.append(room_id)
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

    for room_id in no_geometry_rooms:
        debug_messages.append(f"No geometry found for room ID: {room_id}")
    
    df_relationships = pd.DataFrame(relationships)
    df_unrelated = pd.DataFrame(unrelated_ceilings, columns=['Ceiling_ID', 'Ceiling_Type', 'Ceiling_Description', 'Ceiling_Area_sqm', 'Ceiling_Level'])
    
    # Sort the relationships DataFrame
    df_relationships = df_relationships.sort_values(
        by=['Room_Building', 'Room_Level', 'Room_Number', 'Ceiling_Description'],
        key=lambda x: x.map(custom_sort_key)
    )
    
    # Sort the unrelated ceilings DataFrame
    df_unrelated = df_unrelated.sort_values(
        by=['Ceiling_Level', 'Ceiling_Description'],
        key=lambda x: x.map(custom_sort_key)
    )

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
    df_rooms_without_ceilings = df_rooms_without_ceilings.sort_values(
        by=['Room_Building', 'Room_Level', 'Room_Number'],
        key=lambda x: x.map(custom_sort_key)
    )

    return df_rooms_without_ceilings

def pivot_data(df_relationships):
    """
    Group the relationships data around rooms, preserving all ceiling information in separate rows.
    
    Args:
        df_relationships (pandas.DataFrame): DataFrame containing ceiling-room relationships.
    
    Returns:
        pandas.DataFrame: Grouped DataFrame with rooms and all associated ceiling information.
    """
    # Sort the DataFrame
    df_sorted = df_relationships.sort_values(
        by=['Room_Building', 'Room_Level', 'Room_Number', 'Ceiling_ID'],
        key=lambda x: x.map(custom_sort_key)
    )
    
    # Add a column for the number of ceilings per room
    ceilings_per_room = df_sorted.groupby(['Room_ID'])['Ceiling_ID'].transform('count')
    df_sorted['Ceilings_in_Room'] = ceilings_per_room
    
    # Add helper new columns
    df_sorted['Room'] = df_sorted['Room_Number'] + ' - ' + df_sorted['Room_Name']
    df_sorted['Has_Gypsum_Ceiling'] = df_sorted.groupby('Room_ID').apply(
        lambda x: 1 if (x['Ceiling_Description'] == 'תקרת גבס').any() else 0
    ).reset_index(level=0, drop=True)
    
    # Adjust Intersection_Area_sqm
    df_sorted['Intersection_Area_sqm'] = df_sorted['Intersection_Area_sqm'].apply(lambda x: 0 if x < 1 else x)

    # Reorder columns for better readability
    columns_order = ['Room_Building', 'Room_Level', 'Room_Number', 'Room_Name', 'Room', 'Room_ID',
                     'Ceilings_in_Room', 'Has_Gypsum_Ceiling', 'Ceiling_ID', 'Ceiling_Type', 'Ceiling_Description',
                     'Ceiling_Area_sqm', 'Ceiling_Level', 'Intersection_Area_sqm',
                     'Direct_Intersection', 'XY_Projection_Intersection']
    
    df_grouped = df_sorted[columns_order]
    
    return df_grouped

def adjust_gypsum_ceiling_relationships_pivot(wb, source_sheet_name):
    ws_pivot = wb["Gypsum Ceiling Relationships"]
    pivot_table = ws_pivot._pivots[0]  # Get the existing pivot table
    
    # Update pivot table source range
    ws_data = wb[source_sheet_name]
    data_range = f"{source_sheet_name}!A1:{get_column_letter(ws_data.max_column)}{ws_data.max_row}"
    pivot_table.cache.cacheSource.worksheetSource.ref = data_range
    
    # Set fields
    pivot_table.rows = ["Room_Building", "Room"]
    pivot_table.columns = ["Ceiling_Description"]
    pivot_table.values = [("Intersection_Area_sqm", "Sum")]
    pivot_table.filters = ["Has_Gypsum_Ceiling"]
    
    # Set filter to "1" if possible
    try:
        pivot_table.filters[0].values = [True]
    except:
        debug_messages.append("Unable to set filter programmatically. Please set it manually in Excel.")

def adjust_building_ceiling_type_pivot(wb, source_sheet_name):
    ws_pivot = wb["Building-Ceiling Type Pivot"]
    pivot_table = ws_pivot._pivots[0]  # Get the existing pivot table
    
    # Update pivot table source range
    ws_data = wb[source_sheet_name]
    data_range = f"{source_sheet_name}!A1:{get_column_letter(ws_data.max_column)}{ws_data.max_row}"
    pivot_table.cache.cacheSource.worksheetSource.ref = data_range
    
    # Set fields
    pivot_table.rows = ["Room_Building", "Ceiling_Description", "Room", "Room_ID", "Ceiling_ID"]
    pivot_table.values = [
        ("Intersection_Area_sqm", "Sum"),
        ("Ceiling_ID", "Count")
    ]

def adjust_column_widths(ws):
    for column in ws.columns:
        max_length = 0
        column_letter = column[0].column_letter
        for cell in column:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(cell.value)
            except:
                pass
        adjusted_width = (max_length + 2) * 1.2
        ws.column_dimensions[column_letter].width = adjusted_width

def apply_header_format(ws):
    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.fill = PatternFill(start_color="D7E4BC", end_color="D7E4BC", fill_type="solid")
        cell.border = Border(left=Side(style='thin'), 
                             right=Side(style='thin'), 
                             top=Side(style='thin'), 
                             bottom=Side(style='thin'))
        cell.alignment = Alignment(wrap_text=True, vertical='top')

def main():
    """
    Main function to execute the script.
    """
    try:
        # Collect room and ceiling elements
        room_elements = FilteredElementCollector(doc).OfClass(SpatialElement).OfCategory(BuiltInCategory.OST_Rooms).ToElements()
        ceiling_elements = FilteredElementCollector(doc).OfClass(Ceiling).ToElements()

        # Find the relationships between ceilings and rooms
        try:
            df_relationships, df_unrelated = find_ceiling_room_relationships(room_elements, ceiling_elements)
        except Exception as e:
            debug_messages.append(f"Error in find_ceiling_room_relationships: {e}")
            raise

        # Pivot the relationships data
        try:
            df_grouped = pivot_data(df_relationships)
        except Exception as e:
            debug_messages.append(f"Error in pivot_data: {e}")
            raise

        # Find rooms without ceilings
        try:
            df_rooms_without_ceilings = find_rooms_without_ceilings(df_relationships, room_elements)
        except Exception as e:
            debug_messages.append(f"Error in find_rooms_without_ceilings: {e}")
            raise
        
        # Load the template workbook
        template_path = "C:\\Mac\\Home\\Documents\\Shapir\\Exports\\ceiling_room_relationships_template.xlsx"
        wb = load_workbook(template_path)

        # Export to Excel with formatting
        with pd.ExcelWriter(template_path, engine='openpyxl') as writer:
            writer.book = wb
            writer.sheets = dict((ws.title, ws) for ws in wb.worksheets)
            
            df_relationships.to_excel(writer, sheet_name='Ceiling-Room Relationships', index=False)
            df_unrelated.to_excel(writer, sheet_name='Unrelated Ceilings', index=False)
            df_rooms_without_ceilings.to_excel(writer, sheet_name='Rooms Without Ceilings', index=False)
            
            # Apply header format
            apply_header_format(wb['Ceiling-Room Relationships'])
            apply_header_format(wb['Unrelated Ceilings'])
            apply_header_format(wb['Rooms Without Ceilings'])

        # Adjust pivot tables
        adjust_gypsum_ceiling_relationships_pivot(wb, 'Ceiling-Room Relationships')
        adjust_building_ceiling_type_pivot(wb, 'Ceiling-Room Relationships')

        # Adjust column widths
        for sheet in wb.sheetnames:
            adjust_column_widths(wb[sheet])
        
        # Save the workbook
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        output_file_path = f"C:\\Mac\\Home\\Documents\\Shapir\\Exports\\ceiling_room_relationships_{timestamp}.xlsx"
        try:
            wb.save(output_file_path)
            print(f"Schedule with adjusted pivot tables saved to {output_file_path}")
        except Exception as e:
            debug_messages.append(f"Error in saving workbook: {e}")

    except Exception as e:
        debug_messages.append(f"Error in main function: {e}")
    
    print("Debug messages:")
    for msg in debug_messages:
        print(msg)

# Call the main function
if __name__ == "__main__":
    main()