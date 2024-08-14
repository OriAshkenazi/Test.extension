#! python3

import clr
import os
import sys
import io
import time
import traceback
from collections import defaultdict
from functools import wraps, lru_cache

clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory, BuiltInParameter, ElementId, Wall, Floor, Ceiling, FamilyInstance, Area, LocationCurve, RoofBase
from Autodesk.Revit.DB import UnitUtils, DisplayUnitType
from Autodesk.Revit.UI import *

import System
import datetime

# Get the Revit application and document
uidoc = __revit__.ActiveUIDocument
doc = uidoc.Document

# List to keep track of all LRU cached functions
lru_cached_functions = []

def tracked_lru_cache(*args, **kwargs):
    def decorator(func):
        cached_func = lru_cache(*args, **kwargs)(func)
        lru_cached_functions.append(cached_func)
        return cached_func
    return decorator

def clear_all_lru_caches():
    for func in lru_cached_functions:
        func.cache_clear()

def progress_tracker(total_elements):
    '''
    Decorator function that tracks the progress of a function that processes elements.
    The function must accept an `element_counter` argument that yields elements.
    '''
    def decorator(func):
        @wraps(func)
        def wrapper(*args, **kwargs):
            start_time = time.time()
            last_update_time = start_time
            
            def update_progress(i, force_print=False):
                nonlocal last_update_time
                
                current_time = time.time()
                total_elapsed_time = current_time - start_time
                progress = i / total_elements * 100
                
                if force_print or i % 400 == 0 or i == total_elements:
                    print(f"Progress: {progress:.2f}% ({i}/{total_elements} elements) | "
                          f"Time elapsed: {total_elapsed_time:.2f}s")
                
                last_update_time = current_time

            # Create a generator that yields after each element is processed
            def element_counter():
                elements = FilteredElementCollector(args[0]).WhereElementIsNotElementType().ToElements()
                for i, element in enumerate(elements, 1):
                    yield element
                    update_progress(i, force_print=(i % 400 == 0 or i == total_elements))

            # Pass the generator to the wrapped function
            return func(*args, element_counter=element_counter(), **kwargs)
        return wrapper
    return decorator

def get_family_and_type_names(element, doc):
    '''
    Get the family name, type name, and type ID of an element.
    
    Args:
        element (Element): The element to get the names for.
        doc (Document): The Revit document.
    Returns:
        tuple: A tuple containing the family name, type name, type ID, and additional info.
    '''
    try:
        family_name = "Unknown Family"
        type_name = "Unknown Type"
        type_id = "Unknown Type ID"
        additional_info = "No additional info"

        element_type = doc.GetElement(element.GetTypeId())
        
        if element_type is not None:
            type_name_param = element_type.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
            if type_name_param and type_name_param.HasValue:
                type_name = type_name_param.AsString()
            
            type_id = str(element_type.Id.IntegerValue)

            if isinstance(element, FamilyInstance):
                family = element.Symbol.Family
                if family:
                    family_name = family.Name
            elif hasattr(element_type, 'FamilyName'):
                family_name = element_type.FamilyName
        
        if family_name == "Unknown Family" and element.Category:
            family_name = element.Category.Name

        additional_info = get_additional_type_info(element_type)

        return family_name, type_name, type_id, additional_info

    except Exception as e:
        return "Error Family", "Error Type", "Error Type ID", f"Error getting info: {str(e)}"

@tracked_lru_cache(maxsize=None)
def get_additional_type_info(element_type):
    '''
    Get additional information about the element type.

    Args:
        element_type (Element): The element type to get additional info for.
    Returns:
        str: A string containing additional information about the element type.
    '''
    if element_type is None:
        return "No additional info"

    info_params = [
        BuiltInParameter.ALL_MODEL_TYPE_MARK,
        BuiltInParameter.ALL_MODEL_DESCRIPTION,
        BuiltInParameter.ALL_MODEL_MANUFACTURER,
        BuiltInParameter.ALL_MODEL_MODEL
    ]

    additional_info = []
    for param in info_params:
        param_value = element_type.get_Parameter(param)
        if param_value and param_value.HasValue:
            additional_info.append(f"{param_value.Definition.Name}: {param_value.AsString()}")

    return " | ".join(additional_info) if additional_info else "No additional info"

def get_all_parameters(element):
    '''
    Get all parameters of an element as a list of strings.

    Args:
        element (Element): The element to get the parameters for.
    Returns:
        list: A list of strings containing the parameter names and values.
    '''
    params = []
    try:
        for param in element.Parameters:
            try:
                if param.HasValue:
                    value = param.AsValueString() or param.AsString() or str(param.AsDouble())
                else:
                    value = "No Value"
                params.append(f"{param.Definition.Name}: {value}")
            except Exception as e:
                params.append(f"{param.Definition.Name}: Error - {str(e)}")
    except Exception as e:
        params.append(f"Error getting parameters: {str(e)}")
    return params

def calculate_element_metrics(element, doc):
    '''
    Calculate the area, volume, and length of an element.

    Args:
        element (Element): The element to calculate the metrics for.
        doc (Document): The Revit document.
    Returns:
        dict: A dictionary containing the calculated metrics.
        list: A list of strings containing debug information.
    '''
    metrics = {'area': 0.0, 'length': 0.0, 'volume': 0.0}
    debug_info = [f"Element ID: {element.Id.IntegerValue}"]
   
    try:
        category = element.Category
        if category:
            debug_info.append(f"Category: {category.Name}")
            category_id = category.Id.IntegerValue
        else:
            debug_info.append("Category: None")
            category_id = -1

        # Try to get general parameters first
        area_param = element.get_Parameter(BuiltInParameter.HOST_AREA_COMPUTED)
        volume_param = element.get_Parameter(BuiltInParameter.HOST_VOLUME_COMPUTED)
        
        if area_param and area_param.HasValue:
            metrics['area'] = area_param.AsDouble()
        if volume_param and volume_param.HasValue:
            metrics['volume'] = volume_param.AsDouble()

        # Specific element handling
        if isinstance(element, Wall):
            location = element.Location
            if isinstance(location, LocationCurve):
                metrics['length'] = location.Curve.Length
        elif isinstance(element, (Floor, Ceiling, RoofBase)):
            pass  # Already handled by general parameters
        elif isinstance(element, FamilyInstance):
            bbox = element.get_BoundingBox(None)
            if bbox:
                metrics['length'] = max(bbox.Max.X - bbox.Min.X, bbox.Max.Y - bbox.Min.Y, bbox.Max.Z - bbox.Min.Z)
        elif category_id == int(BuiltInCategory.OST_Rooms):
            room_area_param = element.get_Parameter(BuiltInParameter.ROOM_AREA)
            room_volume_param = element.get_Parameter(BuiltInParameter.ROOM_VOLUME)
            if room_area_param and room_area_param.HasValue:
                metrics['area'] = room_area_param.AsDouble()
            if room_volume_param and room_volume_param.HasValue:
                metrics['volume'] = room_volume_param.AsDouble()
        elif category_id == int(BuiltInCategory.OST_Stairs):
            stair_length_param = element.get_Parameter(BuiltInParameter.STAIRS_ACTUAL_RUN_LENGTH)
            if stair_length_param and stair_length_param.HasValue:
                metrics['length'] = stair_length_param.AsDouble()
        elif category_id == int(BuiltInCategory.OST_Railings):
            location = element.Location
            if isinstance(location, LocationCurve):
                metrics['length'] = location.Curve.Length

        # Convert units (assuming input is in imperial units)
        metrics['area'] = UnitUtils.ConvertFromInternalUnits(metrics['area'], DisplayUnitType.DUT_SQUARE_METERS)
        metrics['volume'] = UnitUtils.ConvertFromInternalUnits(metrics['volume'], DisplayUnitType.DUT_CUBIC_METERS)
        metrics['length'] = UnitUtils.ConvertFromInternalUnits(metrics['length'], DisplayUnitType.DUT_METERS)

        debug_info.extend(get_all_parameters(element))

        for metric, value in metrics.items():
            debug_info.append(f"{metric.capitalize()}: {value:.2f}")

    except Exception as e:
        debug_info.append(f"Error calculating metrics: {str(e)}")
        debug_info.append(traceback.format_exc())

    return metrics, debug_info

@progress_tracker(total_elements=FilteredElementCollector(doc).WhereElementIsNotElementType().GetElementCount())
def get_type_metrics(doc, element_counter=None):
    '''
    Get the metrics for each unique element type in the document.

    Args:
        doc (Document): The Revit document.
        element_counter (generator): A generator that yields elements.
    Returns:
        dict: A dictionary containing the metrics for each unique element type.
        list: A list of strings containing errors encountered during processing.
        int: The total number of elements processed
    '''
    metrics = defaultdict(lambda: {'count': 0, 'area': 0.0, 'volume': 0.0, 'length': 0.0, 'type_id': '', 'additional_info': '', 'debug_info': []})
    
    errors = []
    processed_elements = 0
    
    for element in (element_counter or FilteredElementCollector(doc).WhereElementIsNotElementType().ToElements()):
        try:
            family_name, type_name, type_id, additional_info = get_family_and_type_names(element, doc)
            key = (family_name, type_name)
            
            metrics[key]['count'] += 1
            metrics[key]['type_id'] = type_id
            metrics[key]['additional_info'] = additional_info
            
            element_metrics, debug_info = calculate_element_metrics(element, doc)
            
            for metric, value in element_metrics.items():
                if value is not None:
                    metrics[key][metric] += value
            
            if not metrics[key]['debug_info']:
                metrics[key]['debug_info'] = debug_info
            
            processed_elements += 1
        
        except Exception as e:
            error_msg = f"Error processing element {element.Id}: {str(e)}\n{traceback.format_exc()}"
            if error_msg not in errors:
                errors.append(error_msg)
    
    return metrics, errors, processed_elements

def save_metrics_to_file(metrics, errors, processed_elements, output_path):
    '''
    Save the metrics to a text file.

    Args:
        metrics (dict): A dictionary containing the metrics for each unique element type.
        errors (list): A list of strings containing errors encountered during processing.
        processed_elements (int): The total number of elements processed.
        output_path (str): The path to save the output file to.
    '''
    with io.open(output_path, 'w', encoding='utf-8') as f:
        f.write(f"Processed {processed_elements} elements\n\n")
        f.write("Metrics by Type:\n")
        for key, value in metrics.items():
            try:
                f.write(f"\nFamily: {key[0]}, Type: {key[1]}\n")
                f.write(f"Type ID: {value['type_id']}\n")
                f.write(f"Additional Info: {value['additional_info']}\n")
                f.write(f"Count: {value['count']}\n")
                f.write(f"Total Area: {value['area']:.2f} sq m\n")
                f.write(f"Total Volume: {value['volume']:.2f} cu m\n")
                f.write(f"Total Length: {value['length']:.2f} m\n")
                f.write(f"Average Area: {value['area'] / value['count']:.2f} sq m\n")
                f.write(f"Average Volume: {value['volume'] / value['count']:.2f} cu m\n")
                f.write(f"Average Length: {value['length'] / value['count']:.2f} m\n")
                f.write("Debug Info:\n")
                for debug_line in value['debug_info']:
                    f.write(f"  {debug_line}\n")
            except UnicodeEncodeError as e:
                f.write(f"Error writing data for {key}: {str(e)}\n")
        
        if errors:
            f.write("\nErrors:\n")
            for error in errors:
                try:
                    f.write(f"{error}\n")
                except UnicodeEncodeError as e:
                    f.write(f"Error writing error message: {str(e)}\n")

def summarize_metrics(metrics, errors, processed_elements):
    '''
    Generate a summary of the metrics.

    Args:
        metrics (dict): A dictionary containing the metrics for each unique element type.
        errors (list): A list of strings containing errors encountered during processing.
        processed_elements (int): The total number of elements processed.
    Returns:
        str: A string containing the summary of the metrics.
    '''
    summary = f"Processed {processed_elements} elements\n"
    summary += f"Total unique element types: {len(metrics)}\n"
    summary += f"Total errors encountered: {len(errors)}\n\n"

    # Focus on specific element types
    focus_types = ['Wall', 'Floor', 'Ceiling', 'Roof', 'Room', 'Stair', 'Railing']
    focus_metrics = {t: [] for t in focus_types}

    for (family, type_name), data in metrics.items():
        for focus_type in focus_types:
            if focus_type.lower() in family.lower() and "tag" not in family.lower() and "mark" not in family.lower() and "plan" not in family.lower() and "sketch" not in family.lower():
                focus_metrics[focus_type].append((family, type_name, data))

    for focus_type in focus_types:
        summary += f"{focus_type}:\n"
        if focus_metrics[focus_type]:
            for family, type_name, data in focus_metrics[focus_type]:
                summary += f"  {family} - {type_name}:\n"
                summary += f"    Count: {data['count']}\n"
                summary += f"    Total Area: {data['area']:.2f} sq m\n"
                summary += f"    Total Volume: {data['volume']:.2f} cu m\n"
                summary += f"    Total Length: {data['length']:.2f} m\n"
                summary += f"    Average Area: {data['area'] / data['count']:.2f} sq m\n"
                summary += f"    Average Volume: {data['volume'] / data['count']:.2f} cu m\n"
                summary += f"    Average Length: {data['length'] / data['count']:.2f} m\n"
                summary += f"    Additional Info: {data['additional_info']}\n\n"
        else:
            summary += "  No elements of this type processed\n\n"

    # Add information about elements with non-zero metrics
    summary += "Other elements with non-zero metrics:\n"
    for (family, type_name), data in metrics.items():
        if not any(focus_type.lower() in family.lower() for focus_type in focus_types):
            if data['area'] > 0 or data['volume'] > 0 or data['length'] > 0:
                summary += f"  {family} - {type_name}:\n"
                summary += f"    Count: {data['count']}\n"
                summary += f"    Total Area: {data['area']:.2f} sq m\n"
                summary += f"    Total Volume: {data['volume']:.2f} cu m\n"
                summary += f"    Total Length: {data['length']:.2f} m\n"
                summary += f"    Average Area: {data['area'] / data['count']:.2f} sq m\n"
                summary += f"    Average Volume: {data['volume'] / data['count']:.2f} cu m\n"
                summary += f"    Average Length: {data['length'] / data['count']:.2f} m\n"
                summary += f"    Additional Info: {data['additional_info']}\n\n"

    if errors:
        summary += "First 5 errors:\n"
        for error in errors[:5]:
            summary += f"{error}\n\n"

    return summary

def main():
    clear_all_lru_caches()
    try:
        file_name = doc.PathName.split("/")[-1].split(".")[0].split("_")
        prefix = f"{file_name[0]}_{file_name[1]}"

        # Get output folder path
        output_folder_path = get_folder_path("Select the output folder location")
        if not validate_folder_path(output_folder_path):
            return

        # Generate a timestamp for the file name
        timestamp = datetime.datetime.now().strftime("%Y%m%d%H%M%S")
        
        # Create a new file name with the timestamp
        output_file_name = f"{prefix}_ElementAnalysis_{timestamp}.txt"
        
        # Create the full file path for the output file
        output_path = os.path.join(output_folder_path, output_file_name)

        print("Starting element analysis...")
        metrics, errors, processed_elements = get_type_metrics(doc)

        print("Saving detailed results to file...")
        save_metrics_to_file(metrics, errors, processed_elements, output_path)

        # Create a summary file name with the timestamp
        summary_file_name = f"{prefix}_SummaryElementAnalysis_{timestamp}.txt"

        # Create the full file path for the summary file
        summary_path = os.path.join(output_folder_path, summary_file_name)

        print("Generating summary...")
        summary = summarize_metrics(metrics, errors, processed_elements)

        print("Saving summary to file...")
        with io.open(summary_path, 'w', encoding='utf-8') as f:
            f.write(summary)

        TaskDialog.Show("Analysis Complete", f"The analysis results have been saved to {output_path}\nA summary has been saved to {summary_path}")
        clear_all_lru_caches()

    except Exception as e:
        print(f"Error in main function: {str(e)}")
        print(traceback.format_exc())
        TaskDialog.Show("Error", f"An unexpected error occurred: {str(e)}\n\nPlease check the output window for more details.")
        clear_all_lru_caches()

def get_folder_path(prompt):
    try:
        from System.Windows.Forms import FolderBrowserDialog, DialogResult
        folder_dialog = FolderBrowserDialog()
        folder_dialog.Description = prompt
        if folder_dialog.ShowDialog() == DialogResult.OK:
            return folder_dialog.SelectedPath
    except:
        return TaskDialog.Show("Folder Path Input", prompt)

def validate_folder_path(folder_path):
    if not folder_path:
        return False
    if not os.path.exists(folder_path):
        TaskDialog.Show("Error", f"The specified folder does not exist: {folder_path}")
        return False
    return True

# Call the main function
if __name__ == "__main__":
    main()