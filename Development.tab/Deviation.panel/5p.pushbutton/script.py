#! python3

import clr
import math
import pandas as pd
import numpy as np
from Autodesk.Revit.DB import FilteredElementCollector, ElementId, XYZ, BuiltInCategory, RevitLinkInstance, FamilyInstance, FamilySymbol, Transaction
from Autodesk.Revit.UI import TaskDialog
from System.Collections.Generic import List as ClrList
from datetime import datetime
import os

clr.AddReference('RevitAPI')
clr.AddReference('RevitServices')
from RevitServices.Persistence import DocumentManager
from RevitServices.Transactions import TransactionManager

# Get the current document and UIDocument
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

# Get all elements in the model
collector = FilteredElementCollector(doc).WhereElementIsNotElementType().ToElements()

# Function to get the center point of an element
def get_element_center(element):
    bbox = element.get_BoundingBox(None)
    if bbox:
        center = XYZ((bbox.Min.X + bbox.Max.X) / 2.0,
                     (bbox.Min.Y + bbox.Max.Y) / 2.0,
                     (bbox.Min.Z + bbox.Max.Z) / 2.0)
        return center
    return None

# Function to get additional details of an element
def get_element_details(element):
    element_details = {
        'ElementId': element.Id.IntegerValue,
        'Category': element.Category.Name if element.Category else 'N/A',
        'Family': 'N/A',
        'Type': 'N/A',
        'LinkedModel': 'N/A'
    }
    if isinstance(element, FamilyInstance):
        try:
            element_details['Family'] = element.Symbol.Family.Name if element.Symbol else 'N/A'
            element_details['Type'] = element.Symbol.Name if element.Symbol else 'N/A'
        except:
            element_details['Family'] = 'N/A'
            element_details['Type'] = 'N/A'
    if isinstance(element, RevitLinkInstance):
        linked_doc = element.GetLinkDocument()
        if linked_doc:
            element_details['LinkedModel'] = linked_doc.Title
    return element_details

# Create a list to store element data
element_data = []

# Get coordinates of all elements
for element in collector:
    center = get_element_center(element)
    if center:
        details = get_element_details(element)
        details.update({'X': center.X, 'Y': center.Y, 'Z': center.Z})
        element_data.append(details)

# Convert the list to a pandas DataFrame
df = pd.DataFrame(element_data)

# Calculate the center of mass using NumPy
coordinates = df[['X', 'Y', 'Z']].to_numpy()
center_of_mass = np.mean(coordinates, axis=0)

# Calculate the distance of each element from the center of mass using NumPy
df['Distance'] = np.linalg.norm(coordinates - center_of_mass, axis=1)

# Sort the DataFrame by distance
df = df.sort_values(by='Distance').reset_index(drop=True)

# Calculate the first derivative of the distances
df['Derivative'] = df['Distance'].diff().abs()

# Find the maximum derivative
max_derivative = df['Derivative'].max()

# Identify elements until the maximum derivative
drastic_change_index = df['Derivative'].idxmax()

# Select elements up to the point of drastic change
selected_elements = df.iloc[:drastic_change_index + 1]

# # Prepare the file path and name
# timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
# file_path = f"\\Mac\\Home\\Documents\\Shapir\\Exports\\Deviation\\FarElements_{timestamp}.xlsx"

# # Ensure the directory exists
# os.makedirs(os.path.dirname(file_path), exist_ok=True)

# Export the list to Excel
# selected_elements.to_excel(file_path, index=False)

# # Output detailed information about far elements
# output = selected_elements.to_string(columns=['ElementId', 'Category', 'Family', 'Type', 'LinkedModel', 'Distance', 'X', 'Y', 'Z', 'Derivative'], index=False)
# TaskDialog.Show("Far Elements", f"Elements far from the center of mass:\n{output}\n\nExported to {file_path}")

# Get the ElementIds of the selected elements
far_element_ids = [ElementId(int(eid)) for eid in selected_elements['ElementId'].tolist()]
print(far_element_ids)

# Function to hide very far elements in the current view
def hide_very_far_elements(doc, view, element_ids):
    if element_ids:
        # Filter out elements that cannot be hidden or are already hidden
        element_ids_to_hide = ClrList[ElementId]()
        for eid in element_ids:
            element = doc.GetElement(ElementId(eid))
            if element and element.CanBeHidden(view) and not element.IsHidden(view):
                element_ids_to_hide.Add(ElementId(eid))
        
        if element_ids_to_hide.Count > 0:
            t = Transaction(doc)
            t.Start("Hide Very Far Elements")
            try:
                view.HideElements(element_ids_to_hide)
                t.Commit()
            except Exception as e:
                t.RollBack()
                TaskDialog.Show("Error", str(e))
        else:
            TaskDialog.Show("Info", "No elements to hide.")
    else:
        TaskDialog.Show("Info", "No elements to hide.")

# Hide the very far elements in the current view
hide_very_far_elements(doc, uidoc.ActiveView, far_element_ids)

# If you want to return the element IDs instead of showing a TaskDialog
# You can use the following line to return them as a list
# far_elements_ids = selected_elements['ElementId'].tolist()
