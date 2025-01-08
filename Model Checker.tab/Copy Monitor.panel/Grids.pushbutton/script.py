#! python3

# -*- coding: utf-8 -*-
__title__ = "Highlight Grids"
__doc__ = """Highlight Grids that are not copy-monitored or have moved from their linked model positions."""

from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *
from System.Collections.Generic import List
from pyrevit import revit

# Initialize document and active view
doc = __revit__.ActiveUIDocument.Document
active_view = doc.ActiveView

# Helper function to get linked model grids
def get_linked_grids(linked_doc):
    grids = FilteredElementCollector(linked_doc).OfClass(Grid).ToElements()
    return {grid.UniqueId: grid for grid in grids}

# Helper function to compare grid positions
def grids_are_different(grid_a, grid_b):
    # Compare grid curves
    return not grid_a.Curve.IsAlmostEqualTo(grid_b.Curve)

# Main function
def highlight_non_copy_monitored_grids():
    linked_docs = [link_instance.GetLinkDocument() for link_instance in FilteredElementCollector(doc).OfClass(RevitLinkInstance)]
    if not linked_docs:
        return

    # Collect all grids in the active document
    grids_in_doc = FilteredElementCollector(doc).OfClass(Grid).ToElements()

    # Override graphic settings for problematic grids
    override_settings = OverrideGraphicSettings()
    override_settings.SetProjectionLineColor(Color(255, 0, 0))

    # Initialize transaction
    t = Transaction(doc, "Highlight Grids")
    t.Start()

    for grid in grids_in_doc:
        grid_problematic = False  # Default to not problematic

        for linked_doc in linked_docs:
            if not linked_doc:
                continue

            linked_grids = get_linked_grids(linked_doc)

            if grid.UniqueId in linked_grids:
                linked_grid = linked_grids[grid.UniqueId]
                if grids_are_different(grid, linked_grid):
                    grid_problematic = True
                    break
            else:
                grid_problematic = True

        if grid_problematic:
            active_view.SetElementOverrides(grid.Id, override_settings)

    t.Commit()

# Run the function
highlight_non_copy_monitored_grids()
