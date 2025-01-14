#! python3
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter
from System.Collections.Generic import List
from pyrevit import forms

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

def main():
    """
    Creates multiple copies of a selected Revit element.
    Compatible with Revit 2019 and later.
    """
    try:
        # Prompt the user to select an element
        selection = uidoc.Selection
        try:
            ref_picked_obj = selection.PickObject(ObjectType.Element, "Select an element to duplicate")
        except:
            forms.alert("Selection cancelled by user.", exitscript=True)
            return

        picked_element = doc.GetElement(ref_picked_obj.ElementId)
        if not picked_element:
            forms.alert("Invalid element selected.", exitscript=True)
            return

        # Ask for number of copies
        try:
            num_copies = forms.ask_for_string(
                prompt="Enter the total number of elements (including the original):",
                title="Number of Copies",
                default="2"
            )
            if not num_copies or not num_copies.isdigit() or int(num_copies) <= 1:
                forms.alert("Please enter a valid number greater than 1.", exitscript=True)
                return
            
            num_copies = int(num_copies)
        except:
            forms.alert("Operation cancelled by user.", exitscript=True)
            return

        # Start transaction
        t = Transaction(doc, "Duplicate Selected Element")
        t.Start()
        
        try:
            # Create copies
            for _ in range(num_copies - 1):
                ElementTransformUtils.CopyElement(doc, picked_element.Id, XYZ(0, 0, 0))
            t.Commit()
            forms.alert(f"{num_copies} total elements created successfully.", exitscript=False)
        except Exception as e:
            t.RollBack()
            forms.alert(f"Failed to create copies: {str(e)}", exitscript=True)

    except Exception as e:
        forms.alert(f"An error occurred: {str(e)}", exitscript=True)

if __name__ == '__main__':
    main()
