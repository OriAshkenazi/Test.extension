#! python3

import clr
import math
import pandas as pd
import numpy as np
from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory, ElementId, XYZ
from Autodesk.Revit.UI import TaskDialog
from System.Collections.Generic import List

clr.AddReference('RevitAPI')
clr.AddReference('RevitServices')
from RevitServices.Persistence import DocumentManager
from RevitServices.Transactions import TransactionManager

# Initialize the document
doc = DocumentManager.Instance.CurrentDBDocument

# Get all elements in the model
collector = FilteredElementCollector(doc).WhereElementIsNotElementType().ToElements()

# Function to get the center point of an element
def get_element_center(element):
    bbox = element.get_BoundingBox(None)
    if bbox:
        center = (bbox.Min + bbox.Max) / 2.0
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

# Identify elements whose distance is in the top 5%
threshold = df['Distance'].quantile(0.95)
far_elements = df[df['Distance'] > threshold]

# Output the element IDs
output = "\n".join(map(str, far_elements['ElementId'].tolist()))
TaskDialog.Show("Far Elements", f"Element IDs of distinctly far objects:\n{output}")

# If you want to return the element IDs instead of showing a TaskDialog
# You can use the following line to return them as a list
# far_elements_ids = far_elements['ElementId'].tolist()
