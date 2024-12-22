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

def process_solid(solid):
    edges = []
    if not solid.Faces:
        return edges
        
    for edge in solid.Edges:
        curve = edge.AsCurve()
        if isinstance(curve, Line):
            start_point = curve.GetEndPoint(0)
            end_point = curve.GetEndPoint(1)
            
            # Only process horizontal edges (same Z coordinate)
            if abs(start_point.Z - end_point.Z) < 0.001:
                edges.append(get_edge_info(curve))
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
edges = []
for geom_instance in geom_element:
    if isinstance(geom_instance, Solid):
        edges.extend(process_solid(geom_instance))
    elif isinstance(geom_instance, GeometryInstance):
        geom_object = geom_instance.GetInstanceGeometry()
        for obj in geom_object:
            if isinstance(obj, Solid):
                edges.extend(process_solid(obj))

print("\nSTEP 1: Find best Z elevation")
# Group edges by Z elevation
edges_by_z = {}
for edge in edges:
    z = round(edge['z'], 3)
    if z not in edges_by_z:
        edges_by_z[z] = []
    edges_by_z[z].append(edge)

# Find best Z elevation (closest to 2+2 edges)
best_z = None
best_score = float('inf')
for z, z_edges in edges_by_z.items():
    vertical_edges = [e for e in z_edges if e['is_vertical']]
    horizontal_edges = [e for e in z_edges if not e['is_vertical']]
    score = abs(len(vertical_edges) - 2) + abs(len(horizontal_edges) - 2)
    print(f"Z={feet_to_cm(z):.2f} cm: {len(vertical_edges)} vertical, {len(horizontal_edges)} horizontal edges")
    if score < best_score:
        best_score = score
        best_z = z

print(f"\nSelected Z elevation: {feet_to_cm(best_z):.2f} cm")
selected_edges = edges_by_z[best_z]

print("\nSTEP 2: Classify Edges")
# Split into vertical and horizontal edges
vertical_edges = [e for e in selected_edges if e['is_vertical']]
horizontal_edges = [e for e in selected_edges if not e['is_vertical']]

# Find outer frame X coordinates
all_x_coords = []
for edge in selected_edges:
    all_x_coords.extend([edge['start_x'], edge['end_x']])
outer_left_x = min(all_x_coords)
outer_right_x = max(all_x_coords)
tolerance = 0.1  # feet

# Classify vertical edges as outer or inner
outer_vertical_edges = []
inner_vertical_edges = []
for edge in vertical_edges:
    # If edge is at the leftmost or rightmost X coordinate, it's outer frame
    if (abs(edge['start_x'] - outer_left_x) < tolerance or 
        abs(edge['start_x'] - outer_right_x) < tolerance):
        outer_vertical_edges.append(edge)
    else:
        inner_vertical_edges.append(edge)

print("\nOuter vertical edges:")
for edge in outer_vertical_edges:
    print(f"Length: {feet_to_cm(edge['length']):.2f} cm, X: {feet_to_cm(edge['start_x']):.2f} cm")

print("\nInner vertical edges:")
for edge in inner_vertical_edges:
    print(f"Length: {feet_to_cm(edge['length']):.2f} cm, X: {feet_to_cm(edge['start_x']):.2f} cm")

print("\nHorizontal edges:")
for edge in horizontal_edges:
    print(f"Length: {feet_to_cm(edge['length']):.2f} cm")

print("\nSTEP 3: Select edges for measurement")
# Get outer width (longest horizontal)
horizontal_edges.sort(key=lambda x: x['length'], reverse=True)
outer_width = feet_to_cm(horizontal_edges[0]['length'])

# Get inner height (longest of the inner vertical edges)
inner_vertical_edges.sort(key=lambda x: x['length'], reverse=True)
inner_height = feet_to_cm(inner_vertical_edges[0]['length'])

print(f"\nFinal measurements:")
print(f"Outer width (longest horizontal): {outer_width:.2f} cm")
print(f"Inner height (longest inner vertical): {inner_height:.2f} cm")

# Calculate total perimeter
total_perimeter = 2 * (outer_width + inner_height)
print(f"Total perimeter: {total_perimeter:.2f} cm")

# Update the parameter
try:
    t = Transaction(doc, "Update Ceiling Perimeter Length")
    t.Start()
    param = ceiling.LookupParameter("VDC Skirting Length")
    if param:
        param.Set(total_perimeter / 30.48)
        t.Commit()
        print(f"\nParameter updated successfully")
    else:
        t.RollBack()
        print("Error: Parameter 'VDC Skirting Length' not found.")
except Exception as e:
    if t.HasStarted():
        t.RollBack()
    print(f"Error updating parameter: {str(e)}")