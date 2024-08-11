#! python3

import clr
clr.AddReference('RevitAPI')
from Autodesk.Revit.DB import *

doc = __revit__.ActiveUIDocument.Document

def print_level_names_and_ids(doc):
    # Collect all levels in the document
    levels = FilteredElementCollector(doc).OfClass(Level).ToElements()
    
    # Sort levels by elevation
    levels = sorted(levels, key=lambda l: l.Elevation)
    
    # Print tuples of level names and IDs
    print("Level Name and ID tuples:")
    for level in levels:
        name = level.Name
        id = level.Id.IntegerValue
        print(f"('{name}', {id})")


print_level_names_and_ids(__revit__.ActiveUIDocument.Document)

print(doc.PathName)