# -*- coding: utf-8 -*-
import sys
import clr

from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *

clr.AddReference("System.Windows.Forms")
from System.Windows.Forms import *

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

def get_linked_models():
    """Retrieve all linked models in the project."""
    collector = FilteredElementCollector(doc).OfClass(RevitLinkInstance)
    links = list(collector)
    return links

def select_origin_link(links):
    """Present available links in a Windows Form dialog."""
    if not links:
        TaskDialog.Show("Error", "No linked models found. Please load a link and try again.")
        return None
        
    form = Form()
    form.Text = "Select Origin Link"
    form.Width = 400
    form.Height = 150
    form.FormBorderStyle = FormBorderStyle.FixedDialog
    form.StartPosition = FormStartPosition.CenterScreen
    
    # Create and configure ComboBox
    combo = ComboBox()
    combo.Dock = DockStyle.Top
    combo.Margin = Padding(10)
    combo.Width = 360
    combo.DropDownStyle = ComboBoxStyle.DropDownList
    
    # Add links to ComboBox
    urs_index = -1
    for i, link in enumerate(links):
        link_doc = link.GetLinkDocument()
        link_name = link_doc.Title if link_doc else "Unloaded Link"
        combo.Items.Add(link_name)
        if "URS" in link_name.upper() and urs_index == -1:
            urs_index = i
    
    # Select URS link if found
    if urs_index != -1:
        combo.SelectedIndex = urs_index
    elif combo.Items.Count > 0:
        combo.SelectedIndex = 0
    
    # Create and configure Button
    button = Button()
    button.Text = "OK"
    button.DialogResult = DialogResult.OK
    button.Dock = DockStyle.Bottom
    button.Margin = Padding(10)
    
    # Add controls to form
    form.Controls.Add(combo)
    form.Controls.Add(button)
    form.AcceptButton = button
    
    if form.ShowDialog() == DialogResult.OK and combo.SelectedIndex != -1:
        return links[combo.SelectedIndex]
    return None

def get_grid_parameters(grid):
    """Extract relevant parameters from a grid."""
    params = {
        'element': grid,  # Store the element itself
        'name': grid.Name,
        'start_point': grid.Curve.GetEndPoint(0),
        'end_point': grid.Curve.GetEndPoint(1),
        'curve_type': type(grid.Curve).__name__
    }
    
    # Get all parameters
    for param in grid.Parameters:
        if param.HasValue:
            params[param.Definition.Name] = param.AsString()
    
    return params

def collect_grids(doc_or_link):
    """Collect all grids from specified document."""
    if isinstance(doc_or_link, RevitLinkInstance):
        doc_to_use = doc_or_link.GetLinkDocument()
    else:
        doc_to_use = doc_or_link
    
    collector = FilteredElementCollector(doc_to_use).OfClass(Grid)
    grids = {}
    for grid in collector:
        params = get_grid_parameters(grid)
        grids[params['name']] = params  # Store parameters directly
    
    return grids

def compare_points(p1, p2):
    """Compare two points for exact equality."""
    return (p1.X == p2.X and 
            p1.Y == p2.Y and 
            p1.Z == p2.Z)

def compare_grids(current_grid, origin_grid):
    """Compare only grid geometry between models."""
    discrepancies = []
    
    # Compare start and end points
    if not compare_points(current_grid['start_point'], origin_grid['start_point']):
        discrepancies.append("Start point mismatch")
    if not compare_points(current_grid['end_point'], origin_grid['end_point']):
        discrepancies.append("End point mismatch")
        
    return discrepancies

def highlight_grid(grid_data, color=None):
    """Apply graphic override to grid in current view."""
    try:
        if color is None:
            color = Color(255, 0, 0)  # Red
        
        override = OverrideGraphicSettings()
        override.SetProjectionLineColor(color)
        override.SetProjectionLineWeight(6)
        
        current_view = doc.ActiveView
        
        # Access the element directly as it's now at the top level
        grid_element = grid_data['element']
        
        t = Transaction(doc, "Highlight problematic grid")
        t.Start()
        try:
            current_view.SetElementOverrides(grid_element.Id, override)
            t.Commit()
        except Exception as e:
            t.RollBack()
            raise
    except Exception as e:
        print("Error type:", type(e))
        print("Grid data keys:", grid_data.keys())
        raise

def main():
    # Get linked models
    links = get_linked_models()
    if not links:
        TaskDialog.Show("Error", "No linked models found. Please load a link and try again.")
        return
    
    # Select origin link
    origin_link = select_origin_link(links)
    if not origin_link:
        TaskDialog.Show("Info", "No link selected. Operation cancelled.")
        return
    
    # Collect grids from both models
    origin_grids = collect_grids(origin_link)
    current_grids = collect_grids(doc)
    
    if not origin_grids or not current_grids:
        TaskDialog.Show("Error", "No grids found in one or both models.")
        return
    
    # Compare grids and highlight problems
    problematic_grids = []
    message_lines = []
    
    for name, current_grid in current_grids.items():
        if name not in origin_grids:
            message_lines.append("Grid %s not found in origin model" % name)
            problematic_grids.append((current_grid, ["Missing in origin model"]))
            continue
        
        discrepancies = compare_grids(current_grid, origin_grids[name])
        if discrepancies:
            message_lines.append("Discrepancies found in grid %s:" % name)
            for disc in discrepancies:
                message_lines.append("  - %s" % disc)
            problematic_grids.append((current_grid, discrepancies))
    
    # Check for levels in origin model that don't exist in current model
    for name in origin_grids:
        if name not in current_grids:
            message_lines.append("Grid %s from origin model missing in current model" % name)


    # Highlight problematic grids
    if problematic_grids:
        for grid, _ in problematic_grids:
            highlight_grid(grid)
        
        summary = "Found %d problematic grids.\n\n%s" % (
            len(problematic_grids),
            "\n".join(message_lines)
        )
        TaskDialog.Show("Results", summary)
    else:
        TaskDialog.Show("Results", "No problems found. All grids match the origin model.")

if __name__ == '__main__':
    main()