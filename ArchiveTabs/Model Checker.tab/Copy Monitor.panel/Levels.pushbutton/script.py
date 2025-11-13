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

def get_level_parameters(level):
    """Extract relevant parameters from a level."""
    params = {
        'element': level,  # Store the element itself
        'name': level.Name,
        'elevation': level.Elevation
    }
    
    # Get all parameters
    for param in level.Parameters:
        if param.HasValue:
            params[param.Definition.Name] = param.AsString()
    
    return params

def collect_levels(doc_or_link):
    """Collect all levels from specified document."""
    if isinstance(doc_or_link, RevitLinkInstance):
        doc_to_use = doc_or_link.GetLinkDocument()
    else:
        doc_to_use = doc_or_link
    
    collector = FilteredElementCollector(doc_to_use).OfClass(Level)
    levels = {}
    for level in collector:
        params = get_level_parameters(level)
        levels[params['name']] = params  # Store parameters directly
    
    return levels

def compare_levels(current_level, origin_level):
    """Compare level elevations between models."""
    discrepancies = []
    
    # Compare elevations with a small tolerance (1mm = 0.001m)
    tolerance = 0.001
    if abs(current_level['elevation'] - origin_level['elevation']) > tolerance:
        discrepancies.append("Elevation mismatch: Current = %.3fm, Origin = %.3fm" % (
            current_level['elevation'],
            origin_level['elevation']
        ))
        
    return discrepancies

def highlight_level(level_data):
    """Apply graphic override to level in current view."""
    try:
        # Create color using Revit's Color structure
        color = Color(255, 0, 0)  # Red
        
        override = OverrideGraphicSettings()
        override.SetProjectionLineColor(color)
        override.SetProjectionLineWeight(6)
        
        current_view = doc.ActiveView
        level_element = level_data['element']
        
        t = Transaction(doc, "Highlight problematic level")
        t.Start()
        try:
            current_view.SetElementOverrides(level_element.Id, override)
            t.Commit()
        except Exception as e:
            t.RollBack()
            raise
    except Exception as e:
        print("Error type:", type(e))
        print("Level data keys:", level_data.keys())
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
    
    # Collect levels from both models
    origin_levels = collect_levels(origin_link)
    current_levels = collect_levels(doc)
    
    if not origin_levels or not current_levels:
        TaskDialog.Show("Error", "No levels found in one or both models.")
        return
    
    # Compare levels and highlight problems
    problematic_levels = []
    message_lines = []
    
    for name, current_level in current_levels.items():
        if name not in origin_levels:
            message_lines.append("Level %s not found in origin model" % name)
            problematic_levels.append((current_level, ["Missing in origin model"]))
            continue
        
        discrepancies = compare_levels(current_level, origin_levels[name])
        if discrepancies:
            message_lines.append("Discrepancies found in level %s:" % name)
            for disc in discrepancies:
                message_lines.append("  - %s" % disc)
            problematic_levels.append((current_level, discrepancies))
    
    # Check for levels in origin model that don't exist in current model
    for name in origin_levels:
        if name not in current_levels:
            message_lines.append("Level %s from origin model missing in current model" % name)
    
    # Highlight problematic levels
    if problematic_levels:
        for level, _ in problematic_levels:
            highlight_level(level)
        
        summary = "Found %d problematic levels.\n\n%s" % (
            len(problematic_levels),
            "\n".join(message_lines)
        )
        TaskDialog.Show("Results", summary)
    else:
        TaskDialog.Show("Results", "No problems found. All levels match the origin model.")

if __name__ == '__main__':
    main()