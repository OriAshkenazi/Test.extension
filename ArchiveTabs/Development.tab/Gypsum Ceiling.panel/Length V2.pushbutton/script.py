#! python3
# -*- coding: utf-8 -*-

import clr
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')

from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *
from System.Collections.Generic import List

# Get the current Revit application and document
app = __revit__.Application
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

def feet_to_cm(value):
    """Convert feet to centimeters"""
    return value * 30.48

def get_edge_info(curve):
    """Get formatted edge information"""
    start = curve.GetEndPoint(0)
    end = curve.GetEndPoint(1)
    length = curve.Length
    dx = abs(end.X - start.X)
    dy = abs(end.Y - start.Y)
    is_vertical = dx < dy
    x_avg = (start.X + end.X) / 2
    return {
        'curve': curve,
        'length': length,
        'length_cm': feet_to_cm(length),
        'start': start,
        'end': end,
        'z': start.Z,
        'is_vertical': is_vertical,
        'x_avg': x_avg,
        'start_x': start.X,
        'end_x': end.X
    }

def process_geometry(geometry):
    """Process geometry to extract horizontal edges"""
    edges = []
    for geom_obj in geometry:
        if isinstance(geom_obj, Solid):
            for edge in geom_obj.Edges:
                curve = edge.AsCurve()
                if isinstance(curve, Line):
                    start_point = curve.GetEndPoint(0)
                    end_point = curve.GetEndPoint(1)

                    # Only process horizontal edges (same Z coordinate)
                    if abs(start_point.Z - end_point.Z) < 0.001:
                        edges.append(get_edge_info(curve))
        elif isinstance(geom_obj, GeometryInstance):
            edges.extend(process_geometry(geom_obj.GetInstanceGeometry()))
    return edges

# Get the selected element and its geometry
selection = uidoc.Selection
element_id_list = List[ElementId]()
for id in selection.GetElementIds():
    element_id_list.Add(id)

ceiling = doc.GetElement(element_id_list[0])
opt = Options()
opt.ComputeReferences = True
opt.DetailLevel = ViewDetailLevel.Fine
opt.IncludeNonVisibleObjects = True
geom_element = ceiling.get_Geometry(opt)

# Collect all edges
edges = process_geometry(geom_element)

# Group edges by Z elevation
edges_by_z = {}
for edge in edges:
    z = round(edge['z'], 3)
    if z not in edges_by_z:
        edges_by_z[z] = []
    edges_by_z[z].append(edge)

# Determine the best Z elevation
best_z = max(edges_by_z.keys(), key=lambda z: len(edges_by_z[z]))
selected_edges = edges_by_z[best_z]

# Classify edges into vertical and horizontal
vertical_edges = [e for e in selected_edges if e['is_vertical']]
horizontal_edges = [e for e in selected_edges if not e['is_vertical']]

# Find the longest horizontal edges for perimeter calculation
horizontal_edges.sort(key=lambda x: x['length'], reverse=True)
outer_width = feet_to_cm(horizontal_edges[0]['length'])

# Find the longest vertical edges for perimeter calculation
vertical_edges.sort(key=lambda x: x['length'], reverse=True)
inner_height = feet_to_cm(vertical_edges[0]['length'])

# Calculate total skirting perimeter
skirting_perimeter = 2 * (outer_width + inner_height)

# Update the parameter
try:
    t = Transaction(doc, "Update Ceiling Skirting Length")
    t.Start()
    param = ceiling.LookupParameter("VDC Skirting Length")
    if param:
        param.Set(skirting_perimeter / 30.48)  # Convert cm back to feet
        t.Commit()
        print(f"Parameter updated successfully: {skirting_perimeter:.2f} cm")
    else:
        t.RollBack()
        print("Error: Parameter 'VDC Skirting Length' not found.")
except Exception as e:
    if t.HasStarted():
        t.RollBack()
    print(f"Error updating parameter: {str(e)}")
