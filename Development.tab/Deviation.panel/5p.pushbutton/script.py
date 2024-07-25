#! python3

import clr
import math
import pandas as pd
import numpy as np
from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory, ElementId, XYZ, Transaction
from Autodesk.Revit.UI import TaskDialog
from System.Collections.Generic import List
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

# Create a list to store element data
element_data = []

# Get coordinates of all elements
for element in collector:
    center = get_element_center(element)
    if center:
        element_data.append({'ElementId': element.Id.IntegerValue, 'X': center.X, 'Y': center.Y, 'Z': center.Z})

# Convert the list to a pandas DataFrame
df = pd.DataFrame(element_data)

# Calculate the center of mass using NumPy
coordinates = df[['X', 'Y', 'Z']].to_numpy()
center_of_mass = np.mean(coordinates, axis=0)

# Calculate the distance of each element from the center of mass using NumPy
df['Distance'] = np.linalg.norm(coordinates - center_of_mass, axis=1)

# Sort the DataFrame by distance
df = df.sort_values(by='Distance').reset_index(drop=True)

# Identify elements until a drastic change in distance
drastic_change_index = None
for i in range(1, len(df)):
    if df['Distance'].iloc[i] > 1.5 * df['Distance'].iloc[i - 1]:  # This is an arbitrary threshold for drastic change
        drastic_change_index = i
        break

if drastic_change_index is None:
    drastic_change_index = len(df)

selected_elements = df.iloc[:drastic_change_index]

# Prepare the file path and name
timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
file_path = f"\\Mac\\Home\\Documents\\Shapir\\Exports\\Deviation\\FarElements_{timestamp}.xlsx"

# Ensure the directory exists
os.makedirs(os.path.dirname(file_path), exist_ok=True)

# Export the list to Excel
selected_elements.to_excel(file_path, index=False)

# Output detailed information about far elements
output = selected_elements.to_string(columns=['ElementId', 'Distance', 'X', 'Y', 'Z'], index=False)
TaskDialog.Show("Far Elements", f"Elements far from the center of mass:\n{output}\n\nExported to {file_path}")

# Optionally, highlight far elements in Revit
far_element_ids = [ElementId(int(eid)) for eid in selected_elements['ElementId']]
uidoc.Selection.SetElementIds(List[ElementId](far_element_ids))

# If you want to return the element IDs instead of showing a TaskDialog
# You can use the following line to return them as a list
# far_elements_ids = selected_elements['ElementId'].tolist()
