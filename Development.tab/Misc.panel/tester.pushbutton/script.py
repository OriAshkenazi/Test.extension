#! python3

import clr
clr.AddReference('RevitAPI')
clr.AddReference('System')
from Autodesk.Revit.DB import *

# Get the current document
doc = __revit__.ActiveUIDocument.Document

# Collect all ceiling elements in the document
rooms = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Rooms).WhereElementIsNotElementType().ToElements()

debug_messages = []

# Iterate through each ceiling element
for param in rooms[0].Parameters:
        debug_messages.append(f"{param.Definition.Name}")

print(f"Room parameters:")
for msg in debug_messages:
    print(msg)
