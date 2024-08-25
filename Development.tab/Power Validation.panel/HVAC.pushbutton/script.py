#! python3

import clr
import os
import sys
import time
import traceback
import io
import math
from collections import defaultdict
from typing import Tuple, Dict, List, Callable
from functools import wraps, lru_cache

clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *
from System.Windows.Forms import FolderBrowserDialog, DialogResult
from System.Collections.Generic import List

import System
import datetime
import openpyxl
from openpyxl.styles import Font, Alignment
from tqdm import tqdm

# Get the Revit application and document
uidoc = __revit__.ActiveUIDocument
doc = uidoc.Document

lru_cached_functions = []

def tracked_lru_cache(*args, **kwargs):
    """
    Decorator to track functions using LRU cache.
    
    Args:
        *args: Positional arguments for lru_cache.
        **kwargs: Keyword arguments for lru_cache.
    
    Returns:
        Decorator function to wrap the target function.
    """
    def decorator(func):
        cached_func = lru_cache(*args, **kwargs)(func)
        lru_cached_functions.append(cached_func)
        return cached_func
    return decorator

def clear_all_lru_caches():
    """
    Clear all LRU caches for tracked functions.
    """
    for func in lru_cached_functions:
        func.cache_clear()

def get_linked_documents():
    """
    Retrieve linked documents in the current project.
    
    Returns:
        List[Document]: List of linked Revit documents.
    """
    linked_docs = []
    link_instances = FilteredElementCollector(doc).OfClass(RevitLinkInstance)

    for link_instance in link_instances:
        linked_doc = link_instance.GetLinkDocument()
        if linked_doc is not None:
            linked_docs.append(linked_doc)
    
    print(f"Found {len(linked_docs)} linked documents.")
    return linked_docs

def get_3d_view_family_type():
    view_family_types = FilteredElementCollector(doc).OfClass(ViewFamilyType).ToElements()
    for vft in view_family_types:
        if vft.ViewFamily == ViewFamily.ThreeDimensional:
            return vft.Id
    return None

def get_elements_from_linked_docs(category):
    elements = []
    links = FilteredElementCollector(doc).OfClass(RevitLinkInstance).ToElements()
    for link in links:
        link_doc = link.GetLinkDocument()
        if link_doc:
            elements.extend(FilteredElementCollector(link_doc).OfCategory(category).WhereElementIsNotElementType().ToElements())
    return elements

def create_parameter_filter(doc, name, category_id, parameter_id, rule_string):
    filter_rules = FilterStringRule(parameter_id, FilterStringEquals(), rule_string, False)
    param_filter = ElementParameterFilter(filter_rules)
    
    return ParameterFilterElement.Create(doc, name, [category_id], param_filter)

def create_3d_view_with_elements(power_devices, outlets, view_name="Power Devices and Outlets"):
    t = Transaction(doc, "Create 3D View")
    try:
        t.Start()
        
        view_family_type_id = get_3d_view_family_type()
        if view_family_type_id is None:
            raise Exception("No 3D view family type found in the document.")
        
        new_3d_view = View3D.CreateIsometric(doc, view_family_type_id)
        new_3d_view.Name = view_name
        new_3d_view.DetailLevel = ViewDetailLevel.Fine
        new_3d_view.DisplayStyle = DisplayStyle.Realistic
        
        # Get floors and walls from linked documents
        floors = get_elements_from_linked_docs(BuiltInCategory.OST_Floors)
        walls = get_elements_from_linked_docs(BuiltInCategory.OST_Walls)
        
        print(f"Number of power devices: {len(power_devices)}")
        print(f"Number of outlets: {len(outlets)}")
        print(f"Number of floors from linked docs: {len(floors)}")
        print(f"Number of walls from linked docs: {len(walls)}")
        
        # Set visibility for power devices and outlets
        all_visible_elements = power_devices + outlets + floors + walls
        
        # Hide all elements except the ones we want to show
        all_elements = FilteredElementCollector(doc).WhereElementIsNotElementType().ToElements()
        elements_to_hide = [e.Id for e in all_elements if e.Id not in [elem.Id for elem in all_visible_elements]]
        new_3d_view.HideElements(List[ElementId](elements_to_hide))
        
        # Set transparency for floors and walls
        override_settings = OverrideGraphicSettings()
        override_settings.SetSurfaceTransparency(90)
        
        for element in floors + walls:
            new_3d_view.SetElementOverrides(element.Id, override_settings)
        
        t.Commit()
        return new_3d_view
    except Exception as e:
        t.RollBack()
        print(f"Error creating 3D view: {str(e)}")
        return None
    finally:
        if t.HasStarted() and not t.HasEnded():
            t.RollBack()

def get_elements_by_type_ids_from_linked_docs(type_ids, linked_docs):
    """
    Fetch elements in the linked documents that match the specified type IDs.
    
    Args:
        type_ids (List[ElementId]): List of element type IDs to collect.
        linked_docs (List[Document]): List of linked Revit documents.
    
    Returns:
        List[Element]: List of elements that match the provided type IDs.
    """
    elements = []
    for linked_doc in linked_docs:
        collector = FilteredElementCollector(linked_doc).OfClass(FamilyInstance)
        for element in collector:
            if element.GetTypeId().IntegerValue in type_ids:
                elements.append(element)
    print(f"Found {len(elements)} elements in linked documents.")
    return elements

@tracked_lru_cache(maxsize=None)
def calculate_distance(source_point, target_point):
    """
    Calculate the distance between two points and convert it to meters.
    
    Args:
        source_point (XYZ): The source point.
        target_point (XYZ): The target point.
    
    Returns:
        float: The distance between the source and target points in meters.
    """
    distance_feet = source_point.DistanceTo(target_point)
    return distance_feet * 0.3048  # Convert feet to meters

def calculate_nearest_distance(source_elements, target_elements):
    """
    Calculate the nearest distance between each source element and target elements using caching.
    
    Args:
        source_elements (List[Element]): List of source elements.
        target_elements (List[Element]): List of target elements.
    
    Returns:
        List[Dict]: List of dictionaries containing source, nearest target, and distance.
    """
    results = []
    total_comparisons = len(source_elements) * len(target_elements)
    
    # Pre-calculate all target points
    target_points = [target.Location.Point for target in target_elements]
    
    with tqdm(total=total_comparisons, desc="Calculating distances") as pbar:
        for source in source_elements:
            source_point = source.Location.Point
            min_distance = float('inf')
            nearest_target = None
            
            for i, target_point in enumerate(target_points):
                distance = calculate_distance(source_point, target_point)
                
                if distance < min_distance:
                    min_distance = distance
                    nearest_target = target_elements[i]
                
                pbar.update(1)
            
            results.append({
                'source': source,
                'nearest_target': nearest_target,
                'distance': min_distance
            })
    
    return results

def get_elements_by_category_and_description(category, search_strings, linked_docs):
    """
    Fetch elements of a specific category from linked documents that match any of the search strings.
    
    Args:
        category (BuiltInCategory): The category of elements to collect.
        search_strings (List[str]): List of strings to search for in element parameters.
        linked_docs (List[Document]): List of linked Revit documents.
    
    Returns:
        List[Element]: List of elements that match the provided category and search strings.
    """
    elements = []
    for linked_doc in linked_docs:
        collector = FilteredElementCollector(linked_doc).OfCategory(category).WhereElementIsNotElementType()
        for element in collector:
            element_type = linked_doc.GetElement(element.GetTypeId())
            if element_type:
                for param in element_type.Parameters:
                    if param.HasValue and param.StorageType == StorageType.String:
                        value = param.AsString().lower()
                        if any(search.lower() in value for search in search_strings):
                            elements.append(element)
                            break
    print(f"Found {len(elements)} elements of category {category} matching search strings in linked documents")  # Debugging
    return elements

def print_element_info(elements, category_name):
    """
    Print detailed information about elements.
    
    Args:
        elements (List[Element]): List of elements to print information about.
        category_name (str): Name of the category for logging purposes.
    """
    print(f"\nDetailed information for {category_name}:")
    for i, element in enumerate(elements[:5], 1):  # Print info for first 5 elements
        print(f"  Element {i}:")
        print(f"    ID: {element.Id.IntegerValue}")
        print(f"    Category: {safe_get_property(element, 'Category')}")
        element_type = element.Document.GetElement(element.GetTypeId())
        print(f"    Family: {safe_get_property(element_type, 'Family')}")
        print(f"    Type: {safe_get_property(element_type, 'Name')}")
        for param_name in ["Family and Type Name", "Type Comments", "Description", "DESCRIPTION", "DESCRIPTION HEB", "Legend Description"]:
            param = element_type.LookupParameter(param_name)
            if param and param.HasValue:
                print(f"    {param_name}: {param.AsString()}")
        
        # Print all parameters and their values
        print("    All Parameters:")
        for param in element.Parameters:
            if param.HasValue:
                if param.StorageType == StorageType.String:
                    print(f"      {param.Definition.Name}: {param.AsString()}")
                elif param.StorageType == StorageType.Double:
                    print(f"      {param.Definition.Name}: {param.AsDouble()}")
                elif param.StorageType == StorageType.Integer:
                    print(f"      {param.Definition.Name}: {param.AsInteger()}")
                else:
                    print(f"      {param.Definition.Name}: <Unsupported StorageType>")
        
    print(f"  ... and {len(elements) - 5} more elements") if len(elements) > 5 else None

def create_excel_report(comparison_data, output_path):
    """
    Create an Excel report from the comparison data.

    Args:
        comparison_data (List[Dict]): List of dictionaries containing source, nearest target, and distance.
        output_path (str): The file path to save the Excel report.
    """
    def get_sort_key(item):
        source = item['source']
        return (
            safe_get_property(source, 'Category'),
            safe_get_property(source, 'Family'),
            safe_get_property(source, 'Name')
        )

    sorted_data = sorted(comparison_data, key=get_sort_key)
    
    wb = openpyxl.Workbook()
    ws_details = wb.active
    ws_details.title = "Type Comparison Details"
    
    headers = ["Source ID", "Source Category", "Source Family", "Source Type", "Source Mark",
               "Nearest Target ID", "Target Category", "Target Family", "Target Type", "Target Mark",
               "Distance (m)"]
    
    for col, header in enumerate(headers, start=1):
        cell = ws_details.cell(row=1, column=col, value=header)
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal='center')
    
    for row, data in enumerate(sorted_data, start=2):
        source = data['source']
        target = data['nearest_target']
        
        ws_details.cell(row=row, column=1, value=source.Id.IntegerValue)
        ws_details.cell(row=row, column=2, value=safe_get_property(source, 'Category'))
        ws_details.cell(row=row, column=3, value=safe_get_property(source, 'Family'))
        ws_details.cell(row=row, column=4, value=safe_get_property(source, 'Name'))
        ws_details.cell(row=row, column=5, value=get_parameter_value(source, 'Mark'))
        
        ws_details.cell(row=row, column=6, value=target.Id.IntegerValue)
        ws_details.cell(row=row, column=7, value=safe_get_property(target, 'Category'))
        ws_details.cell(row=row, column=8, value=safe_get_property(target, 'Family'))
        ws_details.cell(row=row, column=9, value=safe_get_property(target, 'Name'))
        ws_details.cell(row=row, column=10, value=get_parameter_value(target, 'Mark'))
        
        ws_details.cell(row=row, column=11, value=round(data['distance'], 2))  # Round to 2 decimal places

    # Adjust column widths for better readability
    for column in ws_details.columns:
        max_length = max(len(str(cell.value)) for cell in column)
        column_letter = column[0].column_letter
        ws_details.column_dimensions[column_letter].width = max_length

    ws_details.auto_filter.ref = ws_details.dimensions

    # Freeze the top row
    ws_details.freeze_panes = ws_details['A2']

    wb.save(output_path)
    print(f"Excel report saved to {output_path}")

def get_folder_path(prompt):
    try:
        from System.Windows.Forms import FolderBrowserDialog, DialogResult
        folder_dialog = FolderBrowserDialog()
        folder_dialog.Description = prompt
        if folder_dialog.ShowDialog() == DialogResult.OK:
            return folder_dialog.SelectedPath
    except:
        pass

    try:
        folder_path = input(f"{prompt} (enter full path): ")
        if os.path.isdir(folder_path):
            return folder_path
        else:
            print("Invalid directory. Please try again.")
    except:
        pass

    try:
        path = TaskDialog.Show("Folder Path Input",
                               prompt,
                               TaskDialogCommonButtons.Ok | TaskDialogCommonButtons.Cancel,
                               TaskDialogResult.Ok)
        if path == TaskDialogResult.Ok:
            folder_path = TaskDialog.Show("Folder Path Input", "Enter the full folder path:")
            if os.path.isdir(folder_path):
                return folder_path
            else:
                TaskDialog.Show("Error", "Invalid directory. Please try again.")
    except:
        pass

    TaskDialog.Show("Error", "Unable to get folder path input.")
    return None

def validate_folder_path(folder_path):
    if not folder_path:
        return False
    if not os.path.exists(folder_path):
        TaskDialog.Show("Error", f"The specified folder does not exist: {folder_path}")
        return False
    return True

def safe_get_property(element, property_name):
    """
    Safely get a property value from an element or its type.
    
    Args:
        element (Element): The Revit element.
        property_name (str): The name of the property to retrieve.
    
    Returns:
        str: The property value if accessible, or an error message if not.
    """
    try:
        if property_name == 'Category':
            return element.Category.Name if element.Category else "<Category not found>"
        elif property_name == 'Family':
            return element.Symbol.Family.Name if element.Symbol and element.Symbol.Family else "<Family not found>"
        elif hasattr(element, property_name):
            return str(getattr(element, property_name))
        elif hasattr(element.Symbol, property_name):
            return str(getattr(element.Symbol, property_name))
        else:
            return f"<{property_name} not found>"
    except Exception as e:
        return f"<Error accessing {property_name}: {str(e)}>"

def get_parameter_value(element, param_name):
    """
    Get the value of a parameter from an element.
    
    Args:
        element (Element): The Revit element.
        param_name (str): The name of the parameter.
    
    Returns:
        str: The parameter value if found, or None if not found.
    """
    param = element.LookupParameter(param_name)
    if param and param.HasValue:
        if param.StorageType == StorageType.String:
            return param.AsString()
        elif param.StorageType == StorageType.Double:
            return str(param.AsDouble())
        elif param.StorageType == StorageType.Integer:
            return str(param.AsInteger())
    return None
    
def main():
    folder_path = get_folder_path("Select a folder to save the Excel report")
    if not validate_folder_path(folder_path):
        return

    linked_docs = get_linked_documents()
    if not linked_docs:
        TaskDialog.Show("Error", "No linked documents found.")
        return

    # power_device_type_ids = get_mechanical_equipment_type_ids_by_name("AES")
    # power_device_type_ids.extend(get_mechanical_equipment_type_ids_by_name("מעבים"))
    # outlet_type_ids = get_electrical_fixtures_type_ids_by_description("Socket")
    # outlet_type_ids.extend(get_electrical_fixtures_type_ids_by_description("בית תקע"))

    # print(f"Total power device type IDs: {len(power_device_type_ids)}")  # Debugging
    # print(f"Total outlet type IDs: {len(outlet_type_ids)}")  # Debugging

    # power_devices = get_elements_by_type_ids_from_linked_docs(power_device_type_ids, linked_docs)
    # outlets = get_elements_by_type_ids_from_linked_docs(outlet_type_ids, linked_docs)

    # print(f"Total power devices found: {len(power_devices)}")  # Debugging
    # print(f"Total outlets found: {len(outlets)}")  # Debugging

    # Get power devices and outlets from linked documents
    power_devices = get_elements_by_category_and_description(
        BuiltInCategory.OST_MechanicalEquipment,
        ["AES", "מעבים"],
        linked_docs
    )
    outlets = get_elements_by_category_and_description(
        BuiltInCategory.OST_ElectricalFixtures,
        ["Socket", "בית תקע"],
        linked_docs
    )

    # print_element_info(power_devices, "Power Devices")
    # print_element_info(outlets, "Outlets")

    results = calculate_nearest_distance(power_devices, outlets)

    print(f"Total results: {len(results)}")  # Debugging

    output_path = os.path.join(folder_path, "Comparison_Report.xlsx")
    create_excel_report(results, output_path)

    new_view = create_3d_view_with_elements(power_devices, outlets)
    
    if new_view:
        # Set the active view to the new 3D view
        uidoc.ActiveView = new_view
        TaskDialog.Show("Success", "3D view created successfully.")
    else:     
        TaskDialog.Show("Error", "Failed to create 3D view.")


if __name__ == '__main__':
    main()
