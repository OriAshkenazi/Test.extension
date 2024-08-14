#! python3

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
from System.Windows.Forms import FolderBrowserDialog, DialogResult
from System.Collections.Generic import List

import System
import openpyxl
from openpyxl.styles import Font, Alignment
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
            estimated_times = []
            use_estimated_time = False
            
            def update_progress(i, force_print=False):
                nonlocal last_update_time, use_estimated_time
                
                current_time = time.time()
                total_elapsed_time = current_time - start_time
                progress = i / total_elements * 100
                
                if i > 1:  # Skip estimation for the first element
                    estimated_time_remaining = (total_elapsed_time / i) * (total_elements - i)
                    estimated_times.append(estimated_time_remaining)
                    
                    # Check negative gradient over last 300 elements
                    if len(estimated_times) >= 300:
                        if all(estimated_times[-j] < estimated_times[-j-1] for j in range(1, 300)) and not use_estimated_time:
                            use_estimated_time = True
                            print("Switching to estimated time remaining...")
                        estimated_times.pop(0)  # Remove oldest estimate to maintain 300 elements
                    
                    if use_estimated_time:
                        time_remaining = estimated_time_remaining
                    else:
                        time_remaining = max(0, 85 - total_elapsed_time)  # Start with 85 seconds
                    
                    if force_print or i % 400 == 0 or i == total_elements:
                        print(f"Progress: {progress:.2f}% ({i}/{total_elements} elements) | "
                              f"Time elapsed: {total_elapsed_time:.2f}s | "
                              f"Time remaining: {time_remaining:.2f}s")
                elif force_print:
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
    '''
    Get the family and type names of an element, along with Type ID and additional type info.
    Args:
        element (Element): The element to get the names from.
        doc (Document): The Revit document.
    Returns:
        tuple: A tuple containing the family name, type name, type ID, and additional type info.
    '''
    try:
        family_name = "Unknown Family"
        type_name = "Unknown Type"
        type_id = "Unknown Type ID"
        additional_info = "No additional info"

        # Get element type
        element_type = doc.GetElement(element.GetTypeId())
        
        if element_type is not None:
            # Get type name
            type_name_param = element_type.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
            if type_name_param and type_name_param.HasValue:
                type_name = type_name_param.AsString()
            
            # Get type ID
            type_id = str(element_type.Id.IntegerValue)

            # Get family name
            if isinstance(element, FamilyInstance):
                family = element.Symbol.Family
                if family:
                    family_name = family.Name
            elif hasattr(element_type, 'FamilyName'):
                family_name = element_type.FamilyName
        
        # If family name is still unknown, use category name
        if family_name == "Unknown Family" and element.Category:
            family_name = element.Category.Name

        # Get additional type info
        additional_info = get_additional_type_info(element_type)

        return family_name, type_name, type_id, additional_info

    except Exception as e:
        return "Error Family", "Error Type", "Error Type ID", "Error getting info"

def get_additional_type_info(element_type):
    '''
    Get additional informative parameters for the element type.

    Args:
        element_type (ElementType): The element type to get information from.
    Returns:
        str: A string containing additional type information.
    '''
    if element_type is None:
        return "No additional info"

    info_params = [
        BuiltInParameter.ALL_MODEL_TYPE_MARK,
        BuiltInParameter.ALL_MODEL_DESCRIPTION,
        BuiltInParameter.ALL_MODEL_MANUFACTURER,
        BuiltInParameter.ALL_MODEL_MODEL
    ]

    for param in info_params:
        param_value = element_type.get_Parameter(param)
        if param_value and param_value.HasValue:
            return f"{param_value.Definition.Name}: {param_value.AsString()}"

    return "No additional info"

def safe_get_parameter_value(element, parameter_name):
    '''
    Safely get a parameter value from an element.

    Args:
        element (Element): The Revit element to get the parameter value from.
        parameter_name (str): The name of the parameter to get.
    Returns:
        str: The parameter value as a string, or "None" if not found.
    '''
    try:
        param = element.get_Parameter(parameter_name)
        if param and param.HasValue:
            return param.AsString() or param.AsValueString()
        return getattr(element, parameter_name, None)
    except:
        return None

def calculate_element_metrics(element, doc):
    """
    Calculate area, length, and volume for applicable elements using a robust approach compatible with Revit 2019.
   
    Args:
        element (Element): The Revit element to calculate metrics for.
        doc (Document): The current Revit document.
    Returns:
        tuple: A tuple containing the calculated metrics (area, length, volume) and any error messages.
    """
    metrics = {'area': 0.0, 'length': 0.0, 'volume': 0.0}
    errors = []
   
    try:
        category = element.Category
        category_id = category.Id.IntegerValue if category else -1

        def get_param_value(param, default=0.0):
            if isinstance(param, BuiltInParameter):
                param = element.get_Parameter(param)
            return param.AsDouble() if param and param.HasValue else default

        # Specific element handling
        if isinstance(element, Wall):
            metrics['length'] = element.WallLength
            metrics['area'] = element.WallArea
            metrics['volume'] = get_param_value(BuiltInParameter.HOST_VOLUME_COMPUTED)
            if element.WallType.Kind == WallKind.Curtain:
                curtain_grid = element.CurtainGrid
                if curtain_grid:
                    for grid_line in curtain_grid.GetUGridLines() + curtain_grid.GetVGridLines():
                        metrics['length'] += grid_line.FullCurve.Length
        elif isinstance(element, Floor) or isinstance(element, Ceiling) or isinstance(element, RoofBase):
            metrics['area'] = get_param_value(BuiltInParameter.HOST_AREA_COMPUTED)
            metrics['volume'] = get_param_value(BuiltInParameter.HOST_VOLUME_COMPUTED)
        elif isinstance(element, FamilyInstance):
            # This covers Doors, Windows, Furniture, Plumbing Fixtures, Mechanical Equipment, Electrical Fixtures
            metrics['area'] = get_param_value(BuiltInParameter.HOST_AREA_COMPUTED)
            metrics['volume'] = get_param_value(BuiltInParameter.HOST_VOLUME_COMPUTED)
            for param in [BuiltInParameter.FAMILY_HEIGHT_PARAM, BuiltInParameter.FAMILY_WIDTH_PARAM, BuiltInParameter.FAMILY_DEPTH_PARAM]:
                value = get_param_value(param)
                if value > 0:
                    metrics['length'] = max(metrics['length'], value)
        elif isinstance(element, Area):
            metrics['area'] = element.Area
        elif category_id == int(BuiltInCategory.OST_Rooms):
            metrics['area'] = get_param_value(BuiltInParameter.ROOM_AREA)
            metrics['volume'] = get_param_value(BuiltInParameter.ROOM_VOLUME)
        elif category_id == int(BuiltInCategory.OST_Stairs):
            metrics['length'] = get_param_value(BuiltInParameter.STAIRS_ACTUAL_RUN_WIDTH)
            metrics['area'] = get_param_value(BuiltInParameter.HOST_AREA_COMPUTED)
        elif category_id == int(BuiltInCategory.OST_Railings):
            metrics['length'] = get_param_value(BuiltInParameter.CURVE_ELEM_LENGTH)
        else:
            # Generic approach for other element types
            metrics['area'] = get_param_value(BuiltInParameter.HOST_AREA_COMPUTED)
            metrics['volume'] = get_param_value(BuiltInParameter.HOST_VOLUME_COMPUTED)
            
            if hasattr(element, 'Location') and isinstance(element.Location, LocationCurve):
                metrics['length'] = element.Location.Curve.Length
            else:
                metrics['length'] = get_param_value(BuiltInParameter.CURVE_ELEM_LENGTH)

        # Fallback to LookupParameter if metrics are still zero
        if metrics['area'] == 0:
            metrics['area'] = get_param_value(element.LookupParameter("Area"))
        if metrics['length'] == 0:
            metrics['length'] = get_param_value(element.LookupParameter("Length"))
        if metrics['volume'] == 0:
            metrics['volume'] = get_param_value(element.LookupParameter("Volume"))

        # Convert units (assuming input is in imperial units)
        metrics['area'] *= 0.092903  # sq ft to sq m
        metrics['volume'] *= 0.0283168  # cu ft to cu m
        metrics['length'] *= 0.3048  # ft to m

    except Exception as e:
        errors.append(f"Error calculating metrics: {str(e)}")

    return metrics, errors

@progress_tracker(total_elements=FilteredElementCollector(doc).WhereElementIsNotElementType().GetElementCount())
def get_type_metrics(doc, element_counter=None):
    metrics = defaultdict(lambda: {'count': 0, 'area': 0.0, 'volume': 0.0, 'length': 0.0, 'type_id': ''})
    
    errors = []
    processed_elements = 0
    for element in (element_counter or FilteredElementCollector(doc).WhereElementIsNotElementType().ToElements()):
        try:
            family_name, type_name, type_id, additional_info = get_family_and_type_names(element, doc)
            key = (family_name, type_name)
            
            metrics[key]['count'] += 1
            metrics[key]['type_id'] = type_id
            metrics[key]['additional_info'] = additional_info
            
            element_metrics, local_errors = calculate_element_metrics(element, doc)
            
            for metric, value in element_metrics.items():
                if value is not None:  # Include zero values
                    metrics[key][metric] += value            
            processed_elements += 1
        
        except Exception as e:
            error_msg = f"Error processing element {element.Id}: {str(e)}"
            if error_msg not in errors:
                errors.append(error_msg)
            for err in local_errors:
                if err not in errors:
                    errors.append(err)
    
    print(f"Processed {processed_elements} elements")
    print("Final metrics:")
    for key, value in metrics.items():
        print(f"{key}: {value}")
    
    return metrics, errors

def get_element_info(element, doc):
    '''
    Get the family name, type name, and type ID of an element.

    Args:
        element (Element): The element to get information from.
        doc (Document): The Revit document.
    Returns:
        tuple: A tuple containing the family name, type name, and type ID.
    '''
    try:
        element_type = doc.GetElement(element.GetTypeId())
        category = element.Category
        
        family_name = category.Name if category else "No Category"
        type_name = element_type.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_NAME).AsString() if element_type else "No Type"
        type_id = str(element_type.Id.IntegerValue) if element_type else "No Type ID"
        
        return family_name, type_name, type_id
    except:
        return "Unknown", "Unknown", "Unknown"

def get_element_metrics(element):
    '''
    Get the area, volume, and length of an element in metric.

    Args:
        element (Element): The element to get metrics from.
    Returns:
        dict: A dictionary containing the area, volume, and length of the element
    '''
    metrics = {'area': 0.0, 'volume': 0.0, 'length': 0.0}
    
    params = {
        'area': [BuiltInParameter.HOST_AREA_COMPUTED, BuiltInParameter.SURFACE_AREA],
        'volume': [BuiltInParameter.HOST_VOLUME_COMPUTED],
        'length': [BuiltInParameter.CURVE_ELEM_LENGTH]
    }
    
    for metric, param_list in params.items():
        for param in param_list:
            param_value = element.get_Parameter(param)
            if param_value and param_value.HasValue:
                metrics[metric] = param_value.AsDouble()
                break
    
    # Convert to metric
    metrics['area'] *= 0.092903  # sq ft to sq m
    metrics['volume'] *= 0.0283168  # cu ft to cu m
    metrics['length'] *= 0.3048  # ft to m
    
    return metrics

def compare_models(current_doc, old_doc_path):
    '''
    Compare the types between the current and old models.
    Args:
        current_doc (Document): The current Revit document.
        old_doc_path (str): The path to the old Revit model file.
    Returns:
        tuple: A tuple containing a list of dictionaries with comparison data and a list of errors.
    '''
    app = current_doc.Application
    all_errors = []

    print("Processing current model...")
    current_metrics, current_errors = get_type_metrics(current_doc)
    all_errors.extend(current_errors)

    print("\nProcessing old model...")
    try:
        old_doc = app.OpenDocumentFile(old_doc_path)
        old_metrics, old_errors = get_type_metrics(old_doc)
        all_errors.extend(old_errors)
        old_doc.Close(False)
    except Exception as e:
        all_errors.append(f"Error opening old document: {str(e)}")
        return [], all_errors

    # Add this check
    if current_metrics is None or old_metrics is None:
        all_errors.append("Error: Failed to process one or both models.")
        return [], all_errors

    all_types = set(current_metrics.keys()) | set(old_metrics.keys())

    comparison_data = []
    for key in all_types:
        if key is None:
            continue  # Skip None keys
        family_name, type_name = key
        current = current_metrics.get(key, {'count': 0, 'area': 0, 'volume': 0, 'length': 0})
        old = old_metrics.get(key, {'count': 0, 'area': 0, 'volume': 0, 'length': 0})

        comparison_data.append({
            'Family': family_name,
            'Type': type_name,
            'Type ID': current.get('type_id', 'N/A'),
            'Current Count': current['count'],
            'Old Count': old['count'],
            'Count Diff': current['count'] - old['count'],
            'Current Area': current['area'],
            'Old Area': old['area'],
            'Area Diff': current['area'] - old['area'],
            'Current Volume': current['volume'],
            'Old Volume': old['volume'],
            'Volume Diff': current['volume'] - old['volume'],
            'Current Length': current['length'],
            'Old Length': old['length'],
            'Length Diff': current['length'] - old['length']
        })

    return comparison_data, all_errors

def generate_impact_assessment(summary_metrics):
    '''
    Generate an impact assessment based on the summary metrics.

    Args:
        summary_metrics (dict): A dictionary containing the summary metrics.
    Returns:
        str: A text describing the potential impact on schedule and budget.
    '''
    impact_text = "Based on the changes observed:\n\n"
    
    if summary_metrics["Count"] > 100 or summary_metrics["Area (sqm)"] > 1000:
        impact_text += "- The project has grown significantly in size, which may require additional time and resources.\n"
    elif summary_metrics["Count"] < -50 or summary_metrics["Area (sqm)"] < -500:
        impact_text += "- The project has decreased in size, which might allow for schedule optimization.\n"
    
    if abs(summary_metrics["Volume (cu m)"]) > 100:
        impact_text += "- Substantial changes in volume may affect material quantities and costs.\n"
    
    if abs(summary_metrics["Length (m)"]) > 1000:
        impact_text += "- Significant changes in linear elements may impact installation time and costs.\n"
    
    impact_text += "\nRecommendation: Review the detailed changes and consult with the project team to assess the need for schedule and budget adjustments."
    
    return impact_text

def create_excel_report(comparison_data, output_path):
    '''
    Create an Excel report from the comparison data.

    Args:
        comparison_data (list): A list of dictionaries containing the comparison data.
        output_path (str): The path to save the Excel file.
    Returns:
        None
    '''
    wb = openpyxl.Workbook()
    ws_details = wb.active
    ws_details.title = "Type Comparison Details"
    ws_summary = wb.create_sheet("Summary")
    
    headers = [
        "Family", "Type", "Type ID", "Additional Info",
        "Current Count", "Old Count", "Count Diff",
        "Current Area (sqm)", "Old Area (sqm)", "Area Diff (sqm)",
        "Current Volume (cu m)", "Old Volume (cu m)", "Volume Diff (cu m)",
        "Current Length (m)", "Old Length (m)", "Length Diff (m)"
    ]
    
    sorted_comparison_data = sorted(comparison_data, key=lambda x: safe_sort_key((x.get('Family'), x.get('Type'))))

    # Populate details worksheet
    for col, header in enumerate(headers, start=1):
        cell = ws_details.cell(row=1, column=col, value=header)
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal='center')
    
    for row, data in enumerate(sorted_comparison_data, start=2):
        for col, key in enumerate(headers, start=1):
            value = data.get(key, "")
            cell = ws_details.cell(row=row, column=col, value=value)
            if isinstance(value, float):
                cell.number_format = '0.00'
    
    # Adjust column widths for details worksheet
    for column in ws_details.columns:
        max_length = 0
        column_letter = column[0].column_letter
        for cell in column:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(cell.value)
            except:
                pass
        adjusted_width = (max_length + 2) * 1.2
        ws_details.column_dimensions[column_letter].width = adjusted_width
    
    # Populate summary worksheet
    summary_headers = ["Metric", "Total Change", "Significant Changes"]
    for col, header in enumerate(summary_headers, start=1):
        cell = ws_summary.cell(row=1, column=col, value=header)
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal='center')

    summary_metrics = {
        "Count": sum(data['Count Diff'] for data in comparison_data),
        "Area (sqm)": sum(data.get('Area Diff (sqm)', 0) for data in comparison_data),
        "Volume (cu m)": sum(data.get('Volume Diff (cu m)', 0) for data in comparison_data),
        "Length (m)": sum(data.get('Length Diff (m)', 0) for data in comparison_data)
    }

    significant_changes = {
        "Count": [],
        "Area (sqm)": [],
        "Volume (cu m)": [],
        "Length (m)": []
    }

    for data in comparison_data:
        if abs(data.get('Count Diff', 0)) > 10:
            significant_changes["Count"].append(f"{data['Family']} - {data['Type']}: {data['Count Diff']}")
        if abs(data.get('Area Diff (sqm)', 0)) > 100:
            significant_changes["Area (sqm)"].append(f"{data['Family']} - {data['Type']}: {data.get('Area Diff (sqm)', 0):.2f}")
        if abs(data.get('Volume Diff (cu m)', 0)) > 10:
            significant_changes["Volume (cu m)"].append(f"{data['Family']} - {data['Type']}: {data.get('Volume Diff (cu m)', 0):.2f}")
        if abs(data.get('Length Diff (m)', 0)) > 100:
            significant_changes["Length (m)"].append(f"{data['Family']} - {data['Type']}: {data.get('Length Diff (m)', 0):.2f}")

    for row, (metric, value) in enumerate(summary_metrics.items(), start=2):
        ws_summary.cell(row=row, column=1, value=metric)
        ws_summary.cell(row=row, column=2, value=value)
        ws_summary.cell(row=row, column=3, value="\n".join(significant_changes[metric]))

    # Add potential impact on schedule and budget
    impact_row = len(summary_metrics) + 3
    ws_summary.cell(row=impact_row, column=1, value="Potential Impact on Schedule and Budget")
    ws_summary.cell(row=impact_row, column=1).font = Font(bold=True)
    
    impact_text = generate_impact_assessment(summary_metrics)
    ws_summary.cell(row=impact_row + 1, column=1, value=impact_text)

    # Adjust column widths for summary worksheet
    for column in ws_summary.columns:
        max_length = 0
        column_letter = column[0].column_letter
        for cell in column:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(cell.value)
            except:
                pass
        adjusted_width = (max_length + 2) * 1.2
        ws_summary.column_dimensions[column_letter].width = adjusted_width

    wb.save(output_path)

def get_file_path_flexible(prompt, file_extension):
    '''
    Get a file path from the user using various methods.

    Args:
        prompt (str): The prompt message.
        file_extension (str): The file extension to filter by.
    Returns:
        str: The file path selected by the user.
    '''
    # Try using Windows Forms
    try:
        from System.Windows.Forms import OpenFileDialog, DialogResult
        file_dialog = OpenFileDialog()
        file_dialog.Filter = f"{file_extension.upper()} Files (*.{file_extension})|*.{file_extension}"
        file_dialog.Title = prompt
        if file_dialog.ShowDialog() == DialogResult.OK:
            return file_dialog.FileName
    except:
        pass

    # Try using standard Python input
    try:
        return input(f"{prompt} (enter full path): ")
    except:
        pass

    # If all else fails, use TaskDialog for input
    try:
        path = TaskDialog.Show("File Path Input",
                               prompt,
                               TaskDialogCommonButtons.Ok | TaskDialogCommonButtons.Cancel,
                               TaskDialogResult.Ok)
        if path == TaskDialogResult.Ok:
            return TaskDialog.Show("File Path Input", "Enter the full file path:")
    except:
        pass

    # If we get here, all methods failed
    TaskDialog.Show("Error", "Unable to get file path input.")
    return None

def validate_file_path(file_path, should_exist=True):
    '''
    Validate the file path.

    Args:
        file_path (str): The file path to validate.
        should_exist (bool): Whether the file should exist.
    Returns:
        bool: True if the file path is valid, False otherwise.
    '''
    if not file_path:
        return False
    if should_exist and not os.path.exists(file_path):
        TaskDialog.Show("Error", f"The specified file does not exist: {file_path}")
        return False
    if not file_path.lower().endswith('.rvt') and not file_path.lower().endswith('.xlsx'):
        TaskDialog.Show("Error", f"Invalid file extension: {file_path}")
        return False
    return True

def get_folder_path(prompt):
    '''
    Get a folder path from the user using various methods.

    Args:
        prompt (str): The prompt message.
    Returns:
        str: The folder path selected by the user.
    '''
    try:
        from System.Windows.Forms import FolderBrowserDialog, DialogResult
        folder_dialog = FolderBrowserDialog()
        folder_dialog.Description = prompt
        if folder_dialog.ShowDialog() == DialogResult.OK:
            return folder_dialog.SelectedPath
    except:
        return TaskDialog.Show("Folder Path Input", prompt)

def validate_folder_path(folder_path):
    '''
    Validate the folder path.

    Args:
        folder_path (str): The folder path to validate.
    Returns:
        bool: True if the folder path is valid, False otherwise.
    '''
    if not folder_path:
        return False
    if not os.path.exists(folder_path):
        TaskDialog.Show("Error", f"The specified folder does not exist: {folder_path}")
        return False
    return True

def safe_sort_key(item):
    if item is None:
        return ("", "")  # Return empty tuple for None values
    if isinstance(item, tuple):
        return tuple(safe_sort_key(i) for i in item)
    return (str(item), "")

def main():
    try:
        file_name = doc.PathName.split("/")[-1].split(".")[0].split("_")
        prefix = f"{file_name[0]}_{file_name[1]}"

        # Get the old model path
        old_doc_path = get_file_path_flexible("Select the old Revit model file", "rvt")
        if not validate_file_path(old_doc_path):
            return

        # Get output folder path
        output_folder_path = get_folder_path("Select the output folder location")
        if not validate_folder_path(output_folder_path):
            return

        # Generate a timestamp for the file name
        timestamp = datetime.datetime.now().strftime("%Y%m%d%H%M%S")
        
        # Create a new file name with the timestamp
        output_file_name = f"{prefix}_Comparison_{timestamp}.xlsx"
        
        # Create the full file path by joining the folder path and the new file name
        output_path = os.path.join(output_folder_path, output_file_name)

        print("Starting model comparison...")
        comparison_data, errors = compare_models(doc, old_doc_path)

        if errors: # Check for errors during processing
            print("Errors occurred during processing: ")
            for error in errors:
                print(error)
            TaskDialog.Show("Error", f"Errors occurred during processing. Check the output window for details.")
            return

        if not comparison_data:
            TaskDialog.Show("Comparison Failed", "No comparison data was generated. Please check the errors.")
            return

        print("Creating Excel report...")
        create_excel_report(comparison_data, output_path)
        TaskDialog.Show("Comparison Complete", f"The comparison results have been saved to {output_path}")

    except Exception as e:
        print(f"Error in main function: {str(e)}")
        print(traceback.format_exc())
        TaskDialog.Show("Error", f"An unexpected error occurred: {str(e)}\n\nPlease check the output window for more details.")

# Call the main function
if __name__ == "__main__":
    main()