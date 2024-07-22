#! python3

import clr
clr.AddReference('RevitAPI')
clr.AddReference('System')

from Autodesk.Revit.DB import *

def get_ceiling_details(ceiling):
    """
    Get details of a ceiling.
    """
    ceiling_id = ceiling.Id.IntegerValue
    ceiling_type_element = doc.GetElement(ceiling.GetTypeId())
    ceiling_type = ceiling_type_element.FamilyName
    
    ceiling_description_param = ceiling_type_element.LookupParameter("Description")
    ceiling_description = ceiling_description_param.AsString() if ceiling_description_param else None
    
    return ceiling_id, ceiling_type, ceiling_description

def get_room_details(room):
    """
    Get details of a room.
    """
    room_id = room.Id.IntegerValue
    room_name = room.get_Parameter(BuiltInParameter.ROOM_NAME).AsString()
    room_level = doc.GetElement(room.LevelId).Name
    room_building_param = room.LookupParameter("בניין")
    room_building = room_building_param.AsString() if room_building_param else None
    return room_id, room_name, room_level, room_building

# Example usage:
doc = __revit__.ActiveUIDocument.Document
# ceilings = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Ceilings).WhereElementIsNotElementType().ToElements()

# for ceiling in ceilings:
#     ceiling_id, ceiling_type, ceiling_description = get_ceiling_details(ceiling)
#     print(f"Ceiling ID: {ceiling_id}, Type: {ceiling_type}, Description: {ceiling_description}")
#     continue

# Collect all rooms in the model
rooms = FilteredElementCollector(doc).OfClass(SpatialElement).OfCategory(BuiltInCategory.OST_Rooms).ToElements()

# Print details of each room
for room in rooms:
    room_id, room_name, room_level, room_building = get_room_details(room)
    print(f"Room ID: {room_id}, Name: {room_name}, Level: {room_level}, Room Building: {room_building}\n")


