#! python3

import clr
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *

uidoc = __revit__.ActiveUIDocument
doc = uidoc.Document

# Get the current selection
selection = uidoc.Selection
selected_element = selection.GetElementIds()

if len(selected_element) == 0:
    print("No elements selected. Please select an element and run the script again.")
elif len(selected_element) > 1:
    print("Multiple elements selected. Please select only one element and run the script again.")
else:
    # Get the single selected element
    element_id = list(selected_element)[0]
    element = doc.GetElement(element_id)
    
    if isinstance(element, RevitLinkInstance):
        print("A Revit link was selected. Please select an element within a linked model, not the link itself.")
    else:
        # Check if the element is from a linked model
        if element.Document.Title != doc.Title:
            # Element is from a linked model
            linked_element = element
            unique_id = linked_element.UniqueId
            print(f"UniqueId of the selected element (from linked model '{linked_element.Document.Title}'): {unique_id}")
        else:
            # Element is from the active document
            unique_id = element.UniqueId
            print(f"UniqueId of the selected element: {unique_id}")
        
        # Copy to clipboard (requires pyrevit)
        from pyrevit import script
        script.get_output().set_width(800)
        output = script.get_output()
        output.clipboard_copy(unique_id)
        print("UniqueId has been copied to clipboard.")