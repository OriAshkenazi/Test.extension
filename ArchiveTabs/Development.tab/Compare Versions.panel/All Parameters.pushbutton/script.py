#! python3

import os
import clr
import datetime
import time
clr.AddReference('RevitAPI')
clr.AddReference('RevitServices')

from Autodesk.Revit.DB import FilteredElementCollector, ElementCategoryFilter, ElementId, BuiltInCategory
from Autodesk.Revit.DB import Element, Document, Parameter, Category, StorageType
from typing import Dict, List

import pandas as pd

def get_folder_path(prompt):
    '''
    Get a folder path from the user using various methods.

    Args:
        prompt (str): The prompt message.
    Returns:
        str: The folder path selected by the user, or None if unsuccessful.
    '''
    # Try using Windows Forms
    try:
        from System.Windows.Forms import FolderBrowserDialog, DialogResult
        folder_dialog = FolderBrowserDialog()
        folder_dialog.Description = prompt
        if folder_dialog.ShowDialog() == DialogResult.OK:
            return folder_dialog.SelectedPath
    except:
        pass

    # Try using standard Python input
    try:
        folder_path = input(f"{prompt} (enter full path): ")
        if os.path.isdir(folder_path):
            return folder_path
        else:
            print("Invalid directory. Please try again.")
    except:
        pass

    # If all else fails, use TaskDialog for input
    try:
        from Autodesk.Revit.UI import TaskDialog, TaskDialogCommonButtons, TaskDialogResult
        path = TaskDialog.Show("Folder Path Input", prompt, TaskDialogCommonButtons.Ok | TaskDialogCommonButtons.Cancel)
        if path == TaskDialogResult.Ok:
            folder_path = TaskDialog.Show("Folder Path Input", "Enter the full folder path:")
            if os.path.isdir(folder_path):
                return folder_path
            else:
                TaskDialog.Show("Error", "Invalid directory. Please try again.")
    except:
        pass

    # If we get here, all methods failed
    TaskDialog.Show("Error", "Unable to get folder path input.")
    return None

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

def list_all_category_parameters(doc: Document) -> Dict[str, Dict[str, str]]:
    '''
    List all categories and their parameters in a Revit document.

    Args:
        doc (Document): The Revit document.

    Returns:
        Dict[str, Dict[str, str]]: A dictionary where the key is the category name, and the value is another dictionary
                                   containing parameter names and their types.
    '''
    global_start = time.time()
    
    category_parameters = {}

    def get_storage_type_name(storage_type):
        if storage_type == StorageType.Integer:
            return "Integer"
        elif storage_type == StorageType.Double:
            return "Double"
        elif storage_type == StorageType.String:
            return "String"
        elif storage_type == StorageType.ElementId:
            return "ElementId"
        else:
            return "Unknown"

    # Collect all categories in the document
    categories = doc.Settings.Categories

    print(f"Found {categories.Size} categories in the document")
    print("Starting to list all categories and their parameters...")
    for i, category in enumerate(categories, 1):
        start = time.time()
        if not category or not category.AllowsBoundParameters:
            continue

        category_name = category.Name
        category_parameters[category_name] = {}

        # Filter elements by category
        category_filter = ElementCategoryFilter(category.Id)
        elements = FilteredElementCollector(doc).WherePasses(category_filter).ToElements()

        # Iterate over the elements and retrieve parameters
        for element in elements:
            for param in element.Parameters:
                param_name = param.Definition.Name
                param_type = get_storage_type_name(param.StorageType)
                if param_name not in category_parameters[category_name]:
                    category_parameters[category_name][param_name] = param_type
        print(f"Found {len(category_parameters[category_name])} parameters in category {category_name}, \
                Time taken: {time.time() - start:.2f} seconds")
    print(f"\nDone listing all categories and their parameters")

    # Remove empty categories and list their amount
    empty_categories = [k for k, v in category_parameters.items() if not v]
    print(f"\nFound {len(empty_categories)} empty categories to be removed")
    category_parameters = {k: v for k, v in category_parameters.items() if v}

    print(f"\nElapsed time: {time.time() - global_start:.2f} seconds")
    return category_parameters

def save_category_parameters_to_excel(doc: Document, folder_path: str) -> None:
    '''
    Save all categories and their parameters to an Excel file.

    Args:
        doc (Document): The Revit document.
        folder_path (str): The folder path where the file will be saved.
    '''
    category_parameters = list_all_category_parameters(doc)
    file_path = os.path.join(folder_path, "category_parameters.xlsx")

    # Create a DataFrame to hold the data
    df = pd.DataFrame()

    for category_name, parameters in category_parameters.items():
        # skip empty categories
        if not parameters:
            continue

        # Create a new DataFrame for each category's parameters
        cat_data = {"Parameter Name": list(parameters.keys()), "Parameter Type": list(parameters.values())}
        cat_df = pd.DataFrame(cat_data)
        
        # Flatten the MultiIndex by adding category name as a prefix
        cat_df.columns = pd.MultiIndex.from_product([[category_name], cat_df.columns])
        
        df = pd.concat([df, cat_df], axis=1)

    # Flatten the columns for writing to Excel
    df.columns = ['_'.join(col).strip() for col in df.columns.values]

    # Save the DataFrame to an Excel file
    with pd.ExcelWriter(file_path) as writer:
        df.to_excel(writer, index=False)

    print(f"\nCategory parameters have been saved to {file_path}")

def main(doc: Document):
    '''
    Main function to execute the script.

    Args:
        doc (Document): The Revit document.
    '''
    folder_path = get_folder_path("Select a folder to save the category parameters file")
    
    if not validate_folder_path(folder_path):
        return
    
    save_category_parameters_to_excel(doc, folder_path)

# Usage example
if __name__ == "__main__":
    doc = __revit__.ActiveUIDocument.Document
    main(doc)
