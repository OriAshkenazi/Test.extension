import clr
import os
import sys
import time
import traceback
from collections import defaultdict
from functools import wraps

clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory, BuiltInParameter, ElementId, Wall, Floor, Ceiling, FamilyInstance, Area, LocationCurve, RoofBase
from Autodesk.Revit.UI import *

import System
import datetime

# Get the Revit application and document
uidoc = __revit__.ActiveUIDocument
doc = uidoc.Document

def progress_tracker(total_elements):
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

            # Create a generator that yields after every element
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

def get_additional_type_info(element_type):
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

def calculate_element_metrics(element, doc):
    metrics = {'area': 0.0, 'length': 0.0, 'volume': 0.0}
    debug_info = []
   
    try:
        category = element.Category
        category_id = category.Id.IntegerValue if category else -1

        def get_param_value(param, default=0.0):
            if isinstance(param, BuiltInParameter):
                param = element.get_Parameter(param)
            value = param.AsDouble() if param and param.HasValue else default
            debug_info.append(f"{param.Definition.Name if param else 'Unknown'}: {value}")
            return value

        # Specific element handling
        if isinstance(element, Wall):
            metrics['length'] = element.WallLength
            metrics['area'] = element.WallArea
            metrics['volume'] = get_param_value(BuiltInParameter.HOST_VOLUME_COMPUTED)
            debug_info.append(f"Wall Length: {metrics['length']}")
            debug_info.append(f"Wall Area: {metrics['area']}")
            if element.WallType.Kind == WallKind.Curtain:
                curtain_grid = element.CurtainGrid
                if curtain_grid:
                    grid_length = sum(grid_line.FullCurve.Length for grid_line in curtain_grid.GetUGridLines() + curtain_grid.GetVGridLines())
                    metrics['length'] += grid_length
                    debug_info.append(f"Curtain Wall Grid Length: {grid_length}")
        elif isinstance(element, Floor) or isinstance(element, Ceiling) or isinstance(element, RoofBase):
            metrics['area'] = get_param_value(BuiltInParameter.HOST_AREA_COMPUTED)
            metrics['volume'] = get_param_value(BuiltInParameter.HOST_VOLUME_COMPUTED)
        elif isinstance(element, FamilyInstance):
            metrics['area'] = get_param_value(BuiltInParameter.HOST_AREA_COMPUTED)
            metrics['volume'] = get_param_value(BuiltInParameter.HOST_VOLUME_COMPUTED)
            for param in [BuiltInParameter.FAMILY_HEIGHT_PARAM, BuiltInParameter.FAMILY_WIDTH_PARAM, BuiltInParameter.FAMILY_DEPTH_PARAM]:
                value = get_param_value(param)
                if value > 0:
                    metrics['length'] = max(metrics['length'], value)
        elif isinstance(element, Area):
            metrics['area'] = element.Area
            debug_info.append(f"Area: {metrics['area']}")
        elif category_id == int(BuiltInCategory.OST_Rooms):
            metrics['area'] = get_param_value(BuiltInParameter.ROOM_AREA)
            metrics['volume'] = get_param_value(BuiltInParameter.ROOM_VOLUME)
        elif category_id == int(BuiltInCategory.OST_Stairs):
            metrics['length'] = get_param_value(BuiltInParameter.STAIRS_ACTUAL_RUN_WIDTH)
            metrics['area'] = get_param_value(BuiltInParameter.HOST_AREA_COMPUTED)
        elif category_id == int(BuiltInCategory.OST_Railings):
            metrics['length'] = get_param_value(BuiltInParameter.CURVE_ELEM_LENGTH)
        else:
            metrics['area'] = get_param_value(BuiltInParameter.HOST_AREA_COMPUTED)
            metrics['volume'] = get_param_value(BuiltInParameter.HOST_VOLUME_COMPUTED)
            
            if hasattr(element, 'Location') and isinstance(element.Location, LocationCurve):
                metrics['length'] = element.Location.Curve.Length
                debug_info.append(f"Location Curve Length: {metrics['length']}")
            else:
                metrics['length'] = get_param_value(BuiltInParameter.CURVE_ELEM_LENGTH)

        # Fallback to LookupParameter if metrics are still zero
        for metric in ['area', 'length', 'volume']:
            if metrics[metric] == 0:
                param = element.LookupParameter(metric.capitalize())
                if param:
                    metrics[metric] = param.AsDouble()
                    debug_info.append(f"Fallback {metric.capitalize()}: {metrics[metric]}")

        # Convert units (assuming input is in imperial units)
        metrics['area'] *= 0.092903  # sq ft to sq m
        metrics['volume'] *= 0.0283168  # cu ft to cu m
        metrics['length'] *= 0.3048  # ft to m

    except Exception as e:
        debug_info.append(f"Error calculating metrics: {str(e)}")

    return metrics, debug_info

@progress_tracker(total_elements=FilteredElementCollector(doc).WhereElementIsNotElementType().GetElementCount())
def get_type_metrics(doc, element_counter=None):
    metrics = defaultdict(lambda: {'count': 0, 'area': 0.0, 'volume': 0.0, 'length': 0.0, 'type_id': '', 'additional_info': '', 'debug_info': []})
    
    processed_types = set()
    errors = []
    processed_elements = 0
    
    for element in (element_counter or FilteredElementCollector(doc).WhereElementIsNotElementType().ToElements()):
        try:
            family_name, type_name, type_id, additional_info = get_family_and_type_names(element, doc)
            key = (family_name, type_name)
            
            if key not in processed_types:
                processed_types.add(key)
                metrics[key]['count'] = 1
                metrics[key]['type_id'] = type_id
                metrics[key]['additional_info'] = additional_info
                
                element_metrics, debug_info = calculate_element_metrics(element, doc)
                
                for metric, value in element_metrics.items():
                    if value is not None:
                        metrics[key][metric] = value
                
                metrics[key]['debug_info'] = debug_info
            
            processed_elements += 1
        
        except Exception as e:
            error_msg = f"Error processing element {element.Id}: {str(e)}"
            if error_msg not in errors:
                errors.append(error_msg)
    
    return metrics, errors, processed_elements

def save_metrics_to_file(metrics, errors, processed_elements, output_path):
    with open(output_path, 'w') as f:
        f.write(f"Processed {processed_elements} elements\n\n")
        f.write("Metrics by Type:\n")
        for key, value in metrics.items():
            f.write(f"\nFamily: {key[0]}, Type: {key[1]}\n")
            f.write(f"Type ID: {value['type_id']}\n")
            f.write(f"Additional Info: {value['additional_info']}\n")
            f.write(f"Count: {value['count']}\n")
            f.write(f"Area: {value['area']:.2f} sq m\n")
            f.write(f"Volume: {value['volume']:.2f} cu m\n")
            f.write(f"Length: {value['length']:.2f} m\n")
            f.write("Debug Info:\n")
            for debug_line in value['debug_info']:
                f.write(f"  {debug_line}\n")
        
        if errors:
            f.write("\nErrors:\n")
            for error in errors:
                f.write(f"{error}\n")

def main():
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
        
        # Create the full file path by joining the folder path and the new file name
        output_path = os.path.join(output_folder_path, output_file_name)

        print("Starting element analysis...")
        metrics, errors, processed_elements = get_type_metrics(doc)

        print("Saving results to file...")
        save_metrics_to_file(metrics, errors, processed_elements, output_path)
        TaskDialog.Show("Analysis Complete", f"The analysis results have been saved to {output_path}")

    except Exception as e:
        print(f"Error in main function: {str(e)}")
        print(traceback.format_exc())
        TaskDialog.Show("Error", f"An unexpected error occurred: {str(e)}\n\nPlease check the output window for more details.")

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