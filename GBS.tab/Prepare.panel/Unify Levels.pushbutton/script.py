# -*- coding: utf-8 -*-
"""Unifies level parameters for all applicable elements in the Revit model."""

__title__ = "Unify\nLevels"
__author__ = "Copilot"
__doc__ = "Unifies level parameters for all applicable elements based on predefined rules."

import clr
import os
import sys
import time
from collections import defaultdict

# Add references for Revit API
clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
from Autodesk.Revit.DB import *

# Add references for IronPython Windows Forms
clr.AddReference("System.Windows.Forms")
clr.AddReference("System.Drawing")
from System.Windows.Forms import Application, Form, Label, Button, MessageBox, CheckBox
from System.Drawing import Point, Size

# Add references for pyRevit
from pyrevit import revit, DB, UI, forms, script

# Get the current document and active view
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

# Set up logging
logger = script.get_logger()
output = script.get_output()

# Constants
DEBUG_MODE = True  # Set to True to enable debug output
TEST_MODE = True   # Set to True to only process a limited number of items
TEST_ITEM_COUNT = 10  # Number of items to process in test mode
OUTPUT_TO_EXCEL = True  # Set to True to export results to Excel

# Define Revit built-in parameter names for level parameters
LEVEL_PARAM_NAMES = {
    "LEVEL_PARAM": BuiltInParameter.LEVEL_PARAM,  # "Level" parameter
    "SCHEDULE_LEVEL_PARAM": BuiltInParameter.SCHEDULE_LEVEL_PARAM,  # "Reference Level" parameter  
    "ASSOCIATED_LEVEL_PARAM": BuiltInParameter.ASSOCIATED_LEVEL_PARAM,  # "Associated Level" parameter
    "FAMILY_LEVEL_PARAM": BuiltInParameter.FAMILY_LEVEL_PARAM,  # "Base Level" parameter
    "INSTANCE_SCHEDULE_ONLY_LEVEL_PARAM": BuiltInParameter.INSTANCE_SCHEDULE_ONLY_LEVEL_PARAM,  # "Base Constraint" parameter
}

# Pre-defined level parameters (static data from the table headers)
LEVEL_PARAMETERS = [
    "Level ElementId Instance Constraints",
    "Reference Level ElementId Instance Constraints",
    "Associated Level ElementId Instance Constraints", 
    "Base Level ElementId Instance Constraints", 
    "Base Constraint ElementId Instance Constraints"
]

# Mapping between the table headers and Revit built-in parameters
PARAM_MAPPING = {
    "Level ElementId Instance Constraints": LEVEL_PARAM_NAMES["LEVEL_PARAM"],
    "Reference Level ElementId Instance Constraints": LEVEL_PARAM_NAMES["SCHEDULE_LEVEL_PARAM"],
    "Associated Level ElementId Instance Constraints": LEVEL_PARAM_NAMES["ASSOCIATED_LEVEL_PARAM"],
    "Base Level ElementId Instance Constraints": LEVEL_PARAM_NAMES["FAMILY_LEVEL_PARAM"],
    "Base Constraint ElementId Instance Constraints": LEVEL_PARAM_NAMES["INSTANCE_SCHEDULE_ONLY_LEVEL_PARAM"]
}

# Try to import Excel libraries for IronPython
try:
    # For IronPython
    clr.AddReference("Microsoft.Office.Interop.Excel")
    from Microsoft.Office.Interop.Excel import ApplicationClass as ExcelApp
    from Microsoft.Office.Interop.Excel import XlFileFormat, Constants
    EXCEL_AVAILABLE = True
    PY_VERSION = "IronPython"
except:
    EXCEL_AVAILABLE = False
    PY_VERSION = "CPython"
    if OUTPUT_TO_EXCEL:
        forms.alert("Excel output requires IronPython. Falling back to on-screen output only.", title="Excel Not Available")

def debug_print(message):
    """Prints debug messages if DEBUG_MODE is True."""
    if DEBUG_MODE:
        logger.debug(message)
        # Using string concatenation instead of f-strings which are not supported in Python 2.7
        output.print_md("**DEBUG:** " + str(message))

def read_excel_data():
    """Reads the Excel file containing level parameter configuration for categories."""
    debug_print("Starting to read Excel data...")
    excel_path = r"C:\Users\orias\OneDrive - shapir.co.il\Documents\--- SBS ---\552  –  Generi  –   ST –  R23\552  –  Generi  – SC  –  R23.xlsm"
    sheet_name = "Sheet1"
    
    if not EXCEL_AVAILABLE:
        forms.alert("Excel interaction is not available. Please ensure Excel libraries are properly installed.", title="Excel Not Available")
        script.exit()
    
    try:
        # Start Excel application
        excel = ExcelApp()
        excel.Visible = False
        debug_print("Excel application started")
        
        # Open workbook
        workbook = excel.Workbooks.Open(excel_path)
        worksheet = workbook.Worksheets[sheet_name]
        debug_print("Opened workbook: " + excel_path + ", sheet: " + sheet_name)
        
        # Read header row (parameter names and special columns)
        level_param_names = []
        special_columns = {}
        
        # Assuming headers are in cells C1:J1
        header_range = worksheet.Range["C1:J1"]
        for i, cell in enumerate(header_range):
            value = cell.Value2
            if value:
                col_letter = chr(ord('C') + i)
                if value in ["Ignore", "Is Level", "Is Dual"]:
                    special_columns[value] = col_letter
                else:
                    level_param_names.append(value)
        
        debug_print("Found {0} level parameters: {1}".format(len(level_param_names), ', '.join(level_param_names)))
        debug_print("Special columns: " + str(special_columns))
        
        # Read category names (B2:B82)
        category_data = {}
        for row in range(2, 83):  # B2:B82
            category_name = worksheet.Range["B{0}".format(row)].Value2
            if not category_name:
                continue
                
            category_info = {
                "level_params": {},
                "ignore": worksheet.Range["{0}{1}".format(special_columns['Ignore'], row)].Value2 == True,
                "has_level": worksheet.Range["{0}{1}".format(special_columns['Is Level'], row)].Value2 == True,
                "is_dual": worksheet.Range["{0}{1}".format(special_columns['Is Dual'], row)].Value2 == True
            }
            
            # Read parameter flags for this category
            for i, param_name in enumerate(level_param_names):
                col_letter = chr(ord('C') + i)
                value = worksheet.Range["{0}{1}".format(col_letter, row)].Value2
                category_info["level_params"][param_name] = value == True
            
            category_data[category_name] = category_info
            debug_print("Loaded config for category: {0}, ignore={1}, has_level={2}, is_dual={3}".format(
                category_name, 
                category_info['ignore'], 
                category_info['has_level'], 
                category_info['is_dual']
            ))
        
        # Clean up
        workbook.Close(False)
        excel.Quit()
        debug_print("Excel data read completed. Found {0} categories.".format(len(category_data)))
        
        return category_data, level_param_names
    
    except Exception as e:
        forms.alert("Error reading Excel file: " + str(e), title="Excel Error")
        debug_print("Excel read error: " + str(e))
        script.exit()

def get_all_elements_by_category():
    """Gets all elements in the model, grouped by category name."""
    debug_print("Collecting all elements by category...")
    elements_by_category = defaultdict(list)
    
    start_time = time.time()
    
    # Get all elements in the model
    collector = FilteredElementCollector(doc).WhereElementIsNotElementType()
    
    count = 0
    for element in collector:
        try:
            # Skip non-categorized elements
            if not element.Category:
                continue
                
            category_name = element.Category.Name
            elements_by_category[category_name].append(element)
            count += 1
            
            if count % 1000 == 0:
                debug_print("Processed {0} elements...".format(count))
        except Exception as e:
            debug_print("Error processing element: " + str(e))
            continue
    
    end_time = time.time()
    debug_print("Collected {0} elements in {1} categories in {2:.2f} seconds".format(
        count, len(elements_by_category), end_time - start_time))
    
    # If in test mode, limit the number of elements per category
    if TEST_MODE:
        debug_print("TEST MODE: Limiting to {0} elements per category".format(TEST_ITEM_COUNT))
        for category_name in elements_by_category:
            if len(elements_by_category[category_name]) > TEST_ITEM_COUNT:
                elements_by_category[category_name] = elements_by_category[category_name][:TEST_ITEM_COUNT]
    
    return elements_by_category

def get_level_from_z_coordinate(z_coord, levels):
    """Determines which level an element belongs to based on its Z coordinate."""
    levels_sorted = sorted(levels, key=lambda l: l.Elevation)
    
    for i in range(len(levels_sorted) - 1):
        lower_level = levels_sorted[i]
        upper_level = levels_sorted[i + 1]
        
        if lower_level.Elevation <= z_coord < upper_level.Elevation:
            return lower_level
            
    # Handle extreme cases
    if z_coord < levels_sorted[0].Elevation:
        debug_print("Z-coordinate {0} is below the lowest level {1}, assigning to the lowest level".format(
            z_coord, levels_sorted[0].Name))
        return levels_sorted[0]  # First level
    else:
        debug_print("Z-coordinate {0} is above the highest level {1}, assigning to the highest level".format(
            z_coord, levels_sorted[-1].Name))
        return levels_sorted[-1]  # Last level

def get_element_levels(element, category_info, level_param_names, all_levels):
    """Gets all level information for an element based on its parameters and category configuration."""
    element_levels = []
    level_params_found = []
    
    element_id = element.Id.IntegerValue
    debug_print("Processing element ID: {0}".format(element_id))
    
    # Check if category has level parameters defined
    if category_info["has_level"]:
        debug_print("Category has level parameters defined in the configuration")
        
        # Get the active level parameters for this category
        active_params = [param for param in level_param_names if category_info["level_params"].get(param, False)]
        debug_print("Active level parameters for this category: {0}".format(active_params))
        
        # Check the defined level parameters for this category
        for param_name in active_params:
            # Try first with LookupParameter using the string name
            param = element.LookupParameter(param_name)
            
            # If that fails, try with the built-in parameter if it's in our mapping
            if not param and param_name in PARAM_MAPPING:
                try:
                    param = element.get_Parameter(PARAM_MAPPING[param_name])
                    debug_print("Using built-in parameter for {0}".format(param_name))
                except Exception as e:
                    debug_print("Error getting built-in parameter: {0}".format(str(e)))
                    param = None
            
            if param and param.HasValue and param.StorageType == StorageType.ElementId:
                level_id = param.AsElementId()
                if level_id != ElementId.InvalidElementId:
                    level = doc.GetElement(level_id)
                    if level and isinstance(level, Level):
                        element_levels.append(level)
                        level_params_found.append(param_name)
                        debug_print("Element {0} has level parameter {1} with value {2}".format(
                            element_id, param_name, level.Name))
    else:
        debug_print("Category does not have predefined level parameters, checking all possible parameters")
        
        # Try both string parameter names and built-in parameters
        for param_name in LEVEL_PARAMETERS:
            # Try with string parameter name
            param = element.LookupParameter(param_name)
            
            # If that fails, try with built-in parameter
            if not param and param_name in PARAM_MAPPING:
                try:
                    param = element.get_Parameter(PARAM_MAPPING[param_name])
                except:
                    param = None
            
            if param and param.HasValue and param.StorageType == StorageType.ElementId:
                level_id = param.AsElementId()
                if level_id != ElementId.InvalidElementId:
                    level = doc.GetElement(level_id)
                    if level and isinstance(level, Level):
                        element_levels.append(level)
                        level_params_found.append(param_name)
                        debug_print("Element {0} has generic level parameter {1} with value {2}".format(
                            element_id, param_name, level.Name))
            
        # Try directly with built-in parameters as a fallback
        if not element_levels:
            for param_name, bip in LEVEL_PARAM_NAMES.items():
                try:
                    param = element.get_Parameter(bip)
                    if param and param.HasValue and param.StorageType == StorageType.ElementId:
                        level_id = param.AsElementId()
                        if level_id != ElementId.InvalidElementId:
                            level = doc.GetElement(level_id)
                            if level and isinstance(level, Level):
                                element_levels.append(level)
                                # Use the original string name for reporting
                                for orig_name, mapped_bip in PARAM_MAPPING.items():
                                    if mapped_bip == bip:
                                        level_params_found.append(orig_name)
                                        break
                                else:
                                    level_params_found.append(param_name)
                                debug_print("Element {0} has built-in level parameter {1} with value {2}".format(
                                    element_id, param_name, level.Name))
                except:
                    continue
    
    # If no level parameters found, try to determine by Z coordinate
    if not element_levels:
        debug_print("Element {0} has no level parameters, checking Z coordinate".format(element_id))
        try:
            # Get the bounding box
            bbox = element.get_BoundingBox(None)
            if bbox:
                # Use the center Z of the bounding box
                z_coord = (bbox.Min.Z + bbox.Max.Z) / 2.0
                debug_print("Element {0} Z-coordinate: {1}".format(element_id, z_coord))
                
                # If it spans multiple levels
                levels_spanned = []
                sorted_levels = sorted(all_levels, key=lambda l: l.Elevation)
                
                for i, level in enumerate(sorted_levels):
                    if i < len(sorted_levels) - 1:
                        next_level = sorted_levels[i + 1]
                        if level.Elevation <= bbox.Min.Z and bbox.Max.Z >= next_level.Elevation:
                            # Element spans multiple levels
                            levels_spanned.append(level)
                            debug_print("Element {0} spans from level {1} to {2}".format(
                                element_id, level.Name, next_level.Name))
                
                if levels_spanned:
                    debug_print("Element {0} spans multiple levels".format(element_id))
                    return levels_spanned, ["Z-Coordinate (Multi-Level)"]
                else:
                    level = get_level_from_z_coordinate(z_coord, all_levels)
                    debug_print("Element {0} assigned to level {1} by Z-coordinate".format(
                        element_id, level.Name))
                    return [level], ["Z-Coordinate"]
            else:
                debug_print("Element {0} has no bounding box".format(element_id))
        except Exception as e:
            debug_print("Error getting Z-coordinate for element {0}: {1}".format(element_id, str(e)))
    
    return element_levels, level_params_found

def process_elements(category_data, level_param_names):
    """Processes all elements according to the level unification rules."""
    debug_print("Starting to process elements...")
    start_time = time.time()
    
    # Get all elements by category
    elements_by_category = get_all_elements_by_category()
    
    # Get all levels in the project
    all_levels = FilteredElementCollector(doc).OfClass(Level).ToElements()
    debug_print("Found {0} levels in the project".format(len(all_levels)))
    
    results = {
        "ignored": [],
        "single_param": [],
        "multi_param_agree": [],
        "multi_param_disagree": [],
        "z_coord": [],
        "no_level": []
    }
    
    # Results by category for Excel output
    category_results = defaultdict(list)
    
    # Process each category
    for category_name, elements in elements_by_category.items():
        debug_print("Processing category: {0} with {1} elements".format(category_name, len(elements)))
        
        if category_name not in category_data:
            debug_print("Category {0} not found in Excel configuration, skipping".format(category_name))
            continue
            
        category_info = category_data[category_name]
        
        # Skip ignored categories
        if category_info["ignore"]:
            debug_print("Category {0} is marked as ignored, skipping".format(category_name))
            results["ignored"].append((category_name, len(elements)))
            continue
            
        # Process each element in the category
        for element in elements:
            element_id = element.Id.IntegerValue
            element_levels, level_params_found = get_element_levels(element, category_info, level_param_names, all_levels)
            
            # No level parameters found
            if not element_levels:
                debug_print("Element {0} has no level information".format(element_id))
                results["no_level"].append((category_name, element.Id))
                category_results[category_name].append({"id": element_id, "status": "No Level", "level": "None", "params": "None"})
                continue
                
            # One level parameter found
            if len(element_levels) == 1:
                level_name = element_levels[0].Name
                param_name = level_params_found[0] if level_params_found else "Unknown"
                
                # Check if this was determined by Z-coordinate
                if param_name == "Z-Coordinate":
                    results["z_coord"].append((category_name, element.Id))
                else:
                    results["single_param"].append((category_name, element.Id, level_name, param_name))
                
                debug_print("Element {0} has single level parameter: {1} from {2}".format(
                    element_id, level_name, param_name))
                category_results[category_name].append({"id": element_id, "status": "Single Parameter", "level": level_name, "params": param_name})
                continue
                
            # Multiple level parameters found - check if they agree
            level_ids = [level.Id for level in element_levels]
            if len(set(level_ids)) == 1:
                # All level parameters agree
                level_name = element_levels[0].Name
                debug_print("Element {0} has multiple agreeing level parameters: {1}".format(
                    element_id, level_name))
                results["multi_param_agree"].append((category_name, element.Id, level_name, ", ".join(level_params_found)))
                category_results[category_name].append({"id": element_id, "status": "Multiple Parameters - Agree", "level": level_name, "params": ", ".join(level_params_found)})
            else:
                # Level parameters disagree
                level_names = ", ".join([level.Name for level in element_levels])
                debug_print("Element {0} has multiple disagreeing level parameters: {1}".format(
                    element_id, level_names))
                results["multi_param_disagree"].append((category_name, element.Id, level_names, ", ".join(level_params_found)))
                category_results[category_name].append({"id": element_id, "status": "Multiple Parameters - Disagree", "level": level_names, "params": ", ".join(level_params_found)})
    
    end_time = time.time()
    debug_print("Finished processing elements in {0:.2f} seconds".format(end_time - start_time))
    
    return results, category_results

def export_to_excel(category_results):
    """Exports the results to an Excel file with a sheet for each category."""
    if not EXCEL_AVAILABLE:
        debug_print("Excel is not available for export")
        return
    
    debug_print("Exporting results to Excel...")
    
    # Create a unique filename based on the current timestamp
    timestamp = time.strftime("%Y%m%d-%H%M%S")
    user_desktop = os.path.join(os.path.expanduser("~"), "Desktop")
    excel_path = os.path.join(user_desktop, "RevitLevelUnification_{0}.xlsx".format(timestamp))
    
    try:
        # Start Excel application
        excel = ExcelApp()
        excel.Visible = False
        excel.DisplayAlerts = False
        
        # Create a new workbook
        workbook = excel.Workbooks.Add()
        
        # Remove default sheets
        for i in range(workbook.Sheets.Count, 0, -1):
            workbook.Sheets.Item[i].Delete()
        
        # Create a summary sheet
        summary_sheet = workbook.Sheets.Add()
        summary_sheet.Name = "Summary"
        
        # Add headers
        summary_sheet.Cells[1, 1].Value = "Category"
        summary_sheet.Cells[1, 2].Value = "Total Elements"
        summary_sheet.Cells[1, 3].Value = "With Level"
        summary_sheet.Cells[1, 4].Value = "Without Level"
        summary_sheet.Cells[1, 5].Value = "Single Parameter"
        summary_sheet.Cells[1, 6].Value = "Multiple Agreeing"
        summary_sheet.Cells[1, 7].Value = "Multiple Disagreeing"
        
        # Format the headers
        header_range = summary_sheet.Range["A1:G1"]
        header_range.Font.Bold = True
        
        # Add data for each category
        row = 2
        for category_name, elements in category_results.items():
            debug_print("Adding Excel sheet for category: {0}".format(category_name))
            
            # Create a sheet for this category
            cat_sheet = workbook.Sheets.Add()
            cat_sheet.Name = category_name[:31]  # Excel sheet names are limited to 31 characters
            
            # Add headers to category sheet
            cat_sheet.Cells[1, 1].Value = "Element ID"
            cat_sheet.Cells[1, 2].Value = "Status"
            cat_sheet.Cells[1, 3].Value = "Level"
            cat_sheet.Cells[1, 4].Value = "Parameters"
            
            # Format the headers
            cat_header_range = cat_sheet.Range["A1:D1"]
            cat_header_range.Font.Bold = True
            
            # Count different types
            single_param_count = 0
            multi_agree_count = 0
            multi_disagree_count = 0
            no_level_count = 0
            
            # Add data for each element
            for i, element_data in enumerate(elements):
                cat_sheet.Cells[i+2, 1].Value = element_data["id"]
                cat_sheet.Cells[i+2, 2].Value = element_data["status"]
                cat_sheet.Cells[i+2, 3].Value = element_data["level"]
                cat_sheet.Cells[i+2, 4].Value = element_data["params"]
                
                # Update counts
                if element_data["status"] == "Single Parameter":
                    single_param_count += 1
                elif element_data["status"] == "Multiple Parameters - Agree":
                    multi_agree_count += 1
                elif element_data["status"] == "Multiple Parameters - Disagree":
                    multi_disagree_count += 1
                elif element_data["status"] == "No Level":
                    no_level_count += 1
            
            # Autofit columns
            cat_sheet.Columns.AutoFit()
            
            # Add summary row
            summary_sheet.Cells[row, 1].Value = category_name
            summary_sheet.Cells[row, 2].Value = len(elements)
            summary_sheet.Cells[row, 3].Value = len(elements) - no_level_count
            summary_sheet.Cells[row, 4].Value = no_level_count
            summary_sheet.Cells[row, 5].Value = single_param_count
            summary_sheet.Cells[row, 6].Value = multi_agree_count
            summary_sheet.Cells[row, 7].Value = multi_disagree_count
            
            row += 1
        
        # Autofit summary columns
        summary_sheet.Columns.AutoFit()
        
        # Save the workbook
        debug_print("Saving Excel file to: {0}".format(excel_path))
        workbook.SaveAs(excel_path)
        
        # Clean up
        workbook.Close(True)
        excel.Quit()
        
        # Show success message
        debug_print("Excel export completed successfully")
        forms.alert("Results exported to Excel file: {0}".format(excel_path), title="Export Successful")
        
    except Exception as e:
        debug_print("Error exporting to Excel: {0}".format(str(e)))
        forms.alert("Error exporting to Excel: {0}".format(str(e)), title="Excel Export Error")

def main():
    """Main function to run the level unification process."""
    debug_print("Starting Level Unification Process")
    if TEST_MODE:
        debug_print("*** TEST MODE ENABLED - Processing only {0} elements per category ***".format(TEST_ITEM_COUNT))
    
    output.print_md("# Unify Levels")
    output.print_md("Reading level parameter configuration from Excel...")
    
    # Read data from Excel
    category_data, level_param_names = read_excel_data()
    
    output.print_md("Loaded configuration for {0} categories with {1} level parameters.".format(
        len(category_data), len(level_param_names)))
    
    # Process elements
    with forms.ProgressBar(title="Processing Elements") as pb:
        results, category_results = process_elements(category_data, level_param_names)
    
    # Export to Excel if enabled
    if OUTPUT_TO_EXCEL and EXCEL_AVAILABLE:
        export_to_excel(category_results)
    
    # Report results on screen
    output.print_md("## Results Summary")
    
    output.print_md("### Ignored Categories")
    output.print_table(
        [[category, count] for category, count in results["ignored"]],
        columns=["Category", "Element Count"]
    )
    
    output.print_md("### Elements with Single Level Parameter")
    output.print_table(
        [[category, element_id.IntegerValue, level_name, param_name] 
         for category, element_id, level_name, param_name in results["single_param"][:100]],  # Limit to 100 rows
        columns=["Category", "Element ID", "Level", "Parameter"]
    )
    
    output.print_md("### Elements with Multiple Agreeing Level Parameters")
    output.print_table(
        [[category, element_id.IntegerValue, level_name, params] 
         for category, element_id, level_name, params in results["multi_param_agree"][:100]],
        columns=["Category", "Element ID", "Level", "Parameters"]
    )
    
    output.print_md("### Elements with Disagreeing Level Parameters")
    output.print_table(
        [[category, element_id.IntegerValue, level_names, params] 
         for category, element_id, level_names, params in results["multi_param_disagree"][:100]],
        columns=["Category", "Element ID", "Levels", "Parameters"]
    )
    
    output.print_md("### Elements Assigned by Z-Coordinate")
    output.print_table(
        [[category, element_id.IntegerValue] 
         for category, element_id in results["z_coord"][:100]],
        columns=["Category", "Element ID"]
    )
    
    output.print_md("### Elements with No Level Assigned")
    output.print_table(
        [[category, element_id.IntegerValue] 
         for category, element_id in results["no_level"][:100]],
        columns=["Category", "Element ID"]
    )
    
    debug_print("Level Unification Process Completed")

if __name__ == "__main__":
    main()