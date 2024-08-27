#! python3

import clr
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *
import sys
import traceback        
from pyrevit import script
from System.Windows import Clipboard
from System.Collections.Generic import List

# Function to log messages
def log(message):
    print(message)

try:
    doc = __revit__.ActiveUIDocument.Document
    uidoc = __revit__.ActiveUIDocument

    log("Script started: Select or Highlight Item by UniqueId")

    # Get UniqueId from clipboard
    unique_id = Clipboard.GetText()

    if not unique_id:
        log("No UniqueId found in clipboard. Please copy a UniqueId and run the script again.")
    else:
        log(f"Searching for element with UniqueId: {unique_id}")

        # Function to search for element in a document
        def find_element_by_unique_id(document, uid):
            return document.GetElement(uid)

        # Try to find the element in the active document
        element = find_element_by_unique_id(doc, unique_id)
        
        if element:
            # Create a List[ElementId] for selection
            ids_to_select = List[ElementId]()
            ids_to_select.Add(element.Id)
            uidoc.Selection.SetElementIds(ids_to_select)
            log(f"Element with UniqueId {unique_id} has been selected in the active document.")
        else:
            log("Element not found in active document. Searching in linked documents...")
            # Search in linked documents
            collector = FilteredElementCollector(doc).OfClass(RevitLinkInstance)
            found = False
            for link_instance in collector:
                link_document = link_instance.GetLinkDocument()
                if link_document:
                    log(f"Searching in linked document: {link_document.Title}")
                    element = find_element_by_unique_id(link_document, unique_id)
                    if element:
                        # Highlight the element in the linked model
                        uidoc.ShowElements(element.Id)
                        uidoc.RefreshActiveView()
                        log(f"Element with UniqueId {unique_id} has been highlighted in the linked model: {link_document.Title}")
                        found = True
                        break
            
            if not found:
                log(f"No element found with UniqueId {unique_id} in active or linked documents.")

    log("Script finished.")

except Exception as e:
    log(f"An error occurred: {str(e)}")
    import traceback
    log(f"Stack trace:\n{traceback.format_exc()}")  # noqa: E501