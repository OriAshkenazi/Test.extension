#! python3

import clr
clr.AddReference('RevitAPI')
clr.AddReference('RevitServices')
from RevitServices.Persistence import DocumentManager
from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory, CeilingType

# Get the current document
doc = __revit__.Instance.CurrentDBDocument

# Collect all ceiling elements in the document
ceilings = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Ceilings).WhereElementIsNotElementType().ToElements()

# Initialize an empty list to store ceiling information
ceiling_info = []

# Iterate through each ceiling element
for ceiling in ceilings:
    # Get the ceiling type
    ceiling_type = doc.GetElement(ceiling.GetTypeId())
    
    # Get the ceiling type name
    type_name = ceiling_type.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString()
    
    # Get the height offset from level
    height_offset_param = ceiling.get_Parameter(BuiltInParameter.CEILING_HEIGHTABOVELEVEL_PARAM)
    height_offset = height_offset_param.AsDouble() if height_offset_param else 0.0
    
    # Append the ceiling information to the list
    ceiling_info.append((type_name, height_offset))

# Print the ceiling information
for info in ceiling_info:
    print(f"Ceiling Type: {info[0]}, Height Offset from Level: {info[1]:.2f} feet")
