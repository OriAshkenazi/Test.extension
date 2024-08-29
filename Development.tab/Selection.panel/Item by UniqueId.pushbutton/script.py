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

def zoom_to_element_in_current_view(doc, uidoc, element, transform):
    """Zoom to the linked element in the current view."""
    current_view = uidoc.ActiveView
    if not isinstance(current_view, View3D):
        log("The current view is not a 3D view. Please open a 3D view and try again.")
        return False

    bbox = element.get_BoundingBox(None)
    if bbox:
        # Transform the bounding box to the local model coordinates
        if transform:
            min_point = transform.OfPoint(bbox.Min)
            max_point = transform.OfPoint(bbox.Max)
            new_bbox = BoundingBoxXYZ()
            new_bbox.Min = min_point
            new_bbox.Max = max_point
        else:
            new_bbox = bbox
        
        # Add some padding to the bounding box
        padding = 5  # feet
        new_bbox.Min = XYZ(new_bbox.Min.X - padding, new_bbox.Min.Y - padding, new_bbox.Min.Z - padding)
        new_bbox.Max = XYZ(new_bbox.Max.X + padding, new_bbox.Max.Y + padding, new_bbox.Max.Z + padding)
        
        # Start a transaction
        trans = Transaction(doc, "Zoom to Linked Element")
        try:
            trans.Start()
            
            # Set the section box for the current view
            current_view.SetSectionBox(new_bbox)
            
            trans.Commit()
            
            # Use the ZoomToFit method to ensure the element is visible
            uidoc.GetOpenUIViews()[0].ZoomToFit()
            
            log("Zoomed to the linked element in the current view.")
            return True
        except Exception as e:
            log(f"An error occurred while zooming to the element: {str(e)}")
            trans.RollBack()
            return False
        finally:
            if trans.HasStarted() and not trans.HasEnded():
                trans.RollBack()
    else:
        log("Couldn't create a bounding box for the element.")
        return False

try:
    doc = __revit__.ActiveUIDocument.Document
    uidoc = __revit__.ActiveUIDocument

    log("Script started: Zoom to Linked Element by UniqueId")

    # Get UniqueId from clipboard
    unique_id = Clipboard.GetText().strip()

    if not unique_id:
        log("No UniqueId found in clipboard. Please copy a UniqueId and run the script again.")
    else:
        log(f"Searching for element with UniqueId: {unique_id}")

        # Function to search for element in a document
        def find_element_by_unique_id(doc, uid):
            return doc.GetElement(uid)

        # Try to find the element in the active document
        element = find_element_by_unique_id(doc, unique_id)
        
        if element:
            log("Element found in active document. This script is intended for linked elements.")
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
                        # Get the transform of the link instance
                        transform = link_instance.GetTotalTransform()
                        
                        # Zoom to the linked element in the current view
                        success = zoom_to_element_in_current_view(doc, uidoc, element, transform)
                        
                        if success:
                            log(f"Zoomed to element with UniqueId {unique_id} from the linked model: {link_document.Title}")
                        else:
                            log(f"Failed to zoom to element with UniqueId {unique_id} from the linked model: {link_document.Title}")
                        found = True
                        break
                
            if not found:
                log(f"No element found with UniqueId {unique_id} in active or linked documents.")

    log("Script finished.")

except Exception as e:
    log(f"An error occurred: {str(e)}")
    log(f"Stack trace:\n{traceback.format_exc()}")