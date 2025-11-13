#! python3

import clr
import traceback
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *
from System.Windows import Clipboard

def log(message):
    print(message)

def get_element_info(element, doc):
    info = f"Element Type: {type(element)}\n"
    info += f"Category: {element.Category.Name if element.Category else 'No Category'}\n"
    info += f"ID: {element.Id}\n"
    
    try:
        info += f"Name: {element.Name}\n"
    except:
        info += "Name: Not available\n"
    
    try:
        info += f"UniqueId: {element.UniqueId}\n"
    except:
        info += "UniqueId: Not available\n"
    
    info += f"Document: {element.Document.Title}\n"
    
    return info

def get_unique_id(doc, uidoc):
    log("Getting current selection...")
    selection = uidoc.Selection.GetElementIds()
    
    log(f"Number of selected elements: {len(selection)}")
    
    if not selection:
        log("No elements selected. Please select an element and run the script again.")
        return None

    if len(selection) > 1:
        log("Multiple elements selected. Please select only one element and run the script again.")
        return None

    log("One element selected. Processing...")
    element_id = list(selection)[0]
    element = doc.GetElement(element_id)

    if element is None:
        # The element might be from a linked model
        log("Element not found in the main document. Checking linked models...")
        element = uidoc.Document.GetElement(element_id)  # This should work for linked elements
    
    if element is None:
        log("Failed to retrieve the selected element. The selection might not be a valid Revit element.")
        return None

    log("Detailed Element Information:")
    log(get_element_info(element, doc))

    if isinstance(element, RevitLinkInstance):
        log("A Revit link instance was selected. Please select an element within the linked model, not the link itself.")
        return None

    try:
        unique_id = element.UniqueId
        if element.Document.Title != doc.Title:
            linked_doc_path = element.Document.PathName
            log(f"UniqueId of the selected element (from linked model '{element.Document.Title}'): {unique_id}")
            log(f"Linked document path: {linked_doc_path}")
            unique_id = f"{linked_doc_path}|{unique_id}"
            log(f"Combined UniqueId: {unique_id}")
        else:
            log(f"UniqueId of the selected element: {unique_id}")
        return unique_id
    except Exception as e:
        log(f"Error retrieving UniqueId: {str(e)}")
        return None

try:
    doc = __revit__.ActiveUIDocument.Document
    uidoc = __revit__.ActiveUIDocument

    log("Script started: Get UniqueId of Selected Element")

    unique_id = get_unique_id(doc, uidoc)
    
    if unique_id:
        Clipboard.SetText(unique_id)
        log("UniqueId copied to clipboard.")
    else:
        log("Failed to retrieve UniqueId.")

    log("Script finished.")

except Exception as e:
    log(f"An error occurred: {str(e)}")
    log(f"Stack trace:\n{traceback.format_exc()}")
