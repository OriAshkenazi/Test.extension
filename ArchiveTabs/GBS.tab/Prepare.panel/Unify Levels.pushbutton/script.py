# -*- coding: utf-8 -*-
"""Unifies level parameters for all applicable elements in the Revit model."""

__title__ = "Unify\nLevels"
__author__ = "Copilot"
__doc__ = "Unifies level parameters for all applicable elements based on predefined rules."

import clr
import os
import sys
import time
import json
from collections import defaultdict

# Add references for Revit API
clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
from Autodesk.Revit.DB import *

# Add references for IronPython Windows Forms
clr.AddReference("System.Windows.Forms")
clr.AddReference("System.Drawing")
from System.Windows.Forms import Application, Form, Label, Button, MessageBox, CheckBox, ComboBox, CheckedListBox, TextBox
from System.Windows.Forms import DialogResult, FormBorderStyle, FormStartPosition, CheckState, GroupBox, Panel, AutoScaleMode
from System.Drawing import Point, Size, Color, Font, FontStyle, SolidBrush, Rectangle

# Add references for pyRevit
from pyrevit import revit, DB, UI, forms, script

# Get the current document and active view
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

# Set up logging
logger = script.get_logger()
output = script.get_output()

# Settings file path for persistent settings
SETTINGS_FILE = os.path.join(os.path.dirname(__file__), "settings.json")

# Default constants - these will be overridden by UI or saved settings
DEFAULT_SETTINGS = {
    "DEBUG_MODE": True,
    "TEST_MODE": True,
    "TEST_ITEM_COUNT": 10,
    "TEST_LINK_ITEM_COUNT": 2,
    "OUTPUT_TO_EXCEL": True,
    "INCLUDE_LINKED_MODELS": True,
    "EXCEL_CONFIG_PATH": r"C:\Users\orias\OneDrive - shapir.co.il\Documents\--- SBS ---\552  –  Generi  –   ST –  R23\552  –  Generi  – SC  –  R23.xlsm",
    "CATEGORIES_TO_CHECK": ["Assemblies", "Duct Linings", "Air Terminals", "Walls"]
}

# Load settings from file
def load_settings():
    """Loads settings from the settings file, or returns defaults if not found."""
    if os.path.exists(SETTINGS_FILE):
        try:
            with open(SETTINGS_FILE, 'r') as f:
                settings = json.load(f)
            return settings
        except Exception as e:
            logger.error("Error loading settings: " + str(e))
    return DEFAULT_SETTINGS

# Save settings to file
def save_settings(settings):
    """Saves the current settings to the settings file."""
    try:
        with open(SETTINGS_FILE, 'w') as f:
            json.dump(settings, f)
    except Exception as e:
        logger.error("Error saving settings: " + str(e))

# Global variables from settings
SETTINGS = load_settings()
DEBUG_MODE = SETTINGS.get("DEBUG_MODE", DEFAULT_SETTINGS["DEBUG_MODE"])
TEST_MODE = SETTINGS.get("TEST_MODE", DEFAULT_SETTINGS["TEST_MODE"])
TEST_ITEM_COUNT = SETTINGS.get("TEST_ITEM_COUNT", DEFAULT_SETTINGS["TEST_ITEM_COUNT"])
TEST_LINK_ITEM_COUNT = SETTINGS.get("TEST_LINK_ITEM_COUNT", DEFAULT_SETTINGS["TEST_LINK_ITEM_COUNT"])
OUTPUT_TO_EXCEL = SETTINGS.get("OUTPUT_TO_EXCEL", DEFAULT_SETTINGS["OUTPUT_TO_EXCEL"])
INCLUDE_LINKED_MODELS = SETTINGS.get("INCLUDE_LINKED_MODELS", DEFAULT_SETTINGS["INCLUDE_LINKED_MODELS"])
EXCEL_CONFIG_PATH = SETTINGS.get("EXCEL_CONFIG_PATH", DEFAULT_SETTINGS["EXCEL_CONFIG_PATH"])
CATEGORIES_TO_CHECK = tuple(SETTINGS.get("CATEGORIES_TO_CHECK", DEFAULT_SETTINGS["CATEGORIES_TO_CHECK"]))

# Settings UI form - adjusted layout and sizes for better visibility
class SettingsForm(Form):
    def __init__(self):
        super(SettingsForm, self).__init__()
        self.InitializeComponent()
        self.LoadSettings()
        
    def InitializeComponent(self):
        self.Text = "Unify Levels Settings"
        self.Width = 600  # Increased width
        self.Height = 640  # Increased height to fit the Excel path control
        self.FormBorderStyle = FormBorderStyle.FixedDialog
        self.StartPosition = FormStartPosition.CenterScreen
        self.AutoScaleMode = AutoScaleMode.Font
        
        # Create controls
        y_pos = 20
        
        # Debug mode
        self.cbDebugMode = CheckBox()
        self.cbDebugMode.Text = "Debug Mode (Enable detailed logging)"
        self.cbDebugMode.Location = Point(20, y_pos)
        self.cbDebugMode.Width = 400  # Increased width
        self.Controls.Add(self.cbDebugMode)
        y_pos += 30
        
        # Test mode
        self.cbTestMode = CheckBox()
        self.cbTestMode.Text = "Test Mode (Process limited number of elements)"
        self.cbTestMode.Location = Point(20, y_pos)
        self.cbTestMode.Width = 400  # Increased width
        self.cbTestMode.CheckedChanged += self.OnTestModeChanged
        self.Controls.Add(self.cbTestMode)
        y_pos += 30
        
        # Test item count
        self.lblTestItemCount = Label()
        self.lblTestItemCount.Text = "Elements per category to process:"
        self.lblTestItemCount.Location = Point(40, y_pos)
        self.lblTestItemCount.Width = 250  # Increased width
        self.Controls.Add(self.lblTestItemCount)
        
        self.txtTestItemCount = TextBox()
        self.txtTestItemCount.Location = Point(300, y_pos)  # Adjusted position
        self.txtTestItemCount.Width = 80  # Increased width
        self.Controls.Add(self.txtTestItemCount)
        y_pos += 30
        
        # Test link item count
        self.lblTestLinkItemCount = Label()
        self.lblTestLinkItemCount.Text = "Elements per linked model to process:"
        self.lblTestLinkItemCount.Location = Point(40, y_pos)
        self.lblTestLinkItemCount.Width = 250  # Increased width
        self.Controls.Add(self.lblTestLinkItemCount)
        
        self.txtTestLinkItemCount = TextBox()
        self.txtTestLinkItemCount.Location = Point(300, y_pos)  # Adjusted position
        self.txtTestLinkItemCount.Width = 80  # Increased width
        self.Controls.Add(self.txtTestLinkItemCount)
        y_pos += 40
        
        # Excel config path
        self.lblExcelPath = Label()
        self.lblExcelPath.Text = "Excel Configuration File Path:"
        self.lblExcelPath.Location = Point(20, y_pos)
        self.lblExcelPath.Width = 250
        self.Controls.Add(self.lblExcelPath)
        y_pos += 25
        
        self.txtExcelPath = TextBox()
        self.txtExcelPath.Location = Point(40, y_pos)
        self.txtExcelPath.Width = 500
        self.Controls.Add(self.txtExcelPath)
        y_pos += 40
        
        # Output to Excel
        self.cbOutputToExcel = CheckBox()
        self.cbOutputToExcel.Text = "Output results to Excel file"
        self.cbOutputToExcel.Location = Point(20, y_pos)
        self.cbOutputToExcel.Width = 400  # Increased width
        self.Controls.Add(self.cbOutputToExcel)
        y_pos += 30
        
        # Include linked models
        self.cbIncludeLinkedModels = CheckBox()
        self.cbIncludeLinkedModels.Text = "Include elements from linked models"
        self.cbIncludeLinkedModels.Location = Point(20, y_pos)
        self.cbIncludeLinkedModels.Width = 400  # Increased width
        self.Controls.Add(self.cbIncludeLinkedModels)
        y_pos += 40
        
        # Categories to check
        self.lblCategories = Label()
        self.lblCategories.Text = "Categories to check:"
        self.lblCategories.Location = Point(20, y_pos)
        self.lblCategories.Width = 250  # Increased width
        self.Controls.Add(self.lblCategories)
        y_pos += 25
        
        # Get all model categories from Revit
        all_categories = []
        for category in doc.Settings.Categories:
            if category.CategoryType == CategoryType.Model:
                all_categories.append(category.Name)
        all_categories.sort()
        
        # Create a checked list box with all categories
        self.clbCategories = CheckedListBox()
        self.clbCategories.Location = Point(20, y_pos)
        self.clbCategories.Size = Size(540, 220)  # Reduced height slightly to fit everything
        for cat in all_categories:
            self.clbCategories.Items.Add(cat)
        self.Controls.Add(self.clbCategories)
        y_pos += 230  # Adjusted for smaller list box
        
        # OK button
        self.btnOK = Button()
        self.btnOK.Text = "OK"
        self.btnOK.Location = Point(370, y_pos)  # Adjusted position
        self.btnOK.Width = 80  # Increased width
        self.btnOK.Click += self.OnOKButtonClick
        self.Controls.Add(self.btnOK)
        
        # Cancel button
        self.btnCancel = Button()
        self.btnCancel.Text = "Cancel"
        self.btnCancel.Location = Point(470, y_pos)  # Adjusted position
        self.btnCancel.Width = 80  # Increased width
        self.btnCancel.Click += self.OnCancelButtonClick
        self.Controls.Add(self.btnCancel)
        
    def LoadSettings(self):
        """Loads the current settings into the form."""
        self.cbDebugMode.Checked = DEBUG_MODE
        self.cbTestMode.Checked = TEST_MODE
        self.txtTestItemCount.Text = str(TEST_ITEM_COUNT)
        self.txtTestLinkItemCount.Text = str(TEST_LINK_ITEM_COUNT)
        self.cbOutputToExcel.Checked = OUTPUT_TO_EXCEL
        self.cbIncludeLinkedModels.Checked = INCLUDE_LINKED_MODELS
        self.txtExcelPath.Text = EXCEL_CONFIG_PATH
        
        # Enable/disable test count fields based on test mode
        self.txtTestItemCount.Enabled = TEST_MODE
        self.lblTestItemCount.Enabled = TEST_MODE
        self.txtTestLinkItemCount.Enabled = TEST_MODE
        self.lblTestLinkItemCount.Enabled = TEST_MODE
        
        # Check the appropriate categories
        for i in range(self.clbCategories.Items.Count):
            if self.clbCategories.Items[i] in CATEGORIES_TO_CHECK:
                self.clbCategories.SetItemChecked(i, True)
        
    def OnTestModeChanged(self, sender, args):
        """Enables or disables test count fields based on test mode."""
        self.txtTestItemCount.Enabled = self.cbTestMode.Checked
        self.lblTestItemCount.Enabled = self.cbTestMode.Checked
        self.txtTestLinkItemCount.Enabled = self.cbTestMode.Checked
        self.lblTestLinkItemCount.Enabled = self.cbTestMode.Checked
        
    def OnOKButtonClick(self, sender, args):
        """Saves the settings and closes the form."""
        try:
            # Validate numeric fields
            test_item_count = int(self.txtTestItemCount.Text)
            test_link_item_count = int(self.txtTestLinkItemCount.Text)
            
            # Validate Excel path
            excel_path = self.txtExcelPath.Text.strip()
            if not os.path.exists(excel_path):
                MessageBox.Show("The Excel configuration file path does not exist. Please enter a valid path.", "Validation Error")
                return
            
            # Get selected categories
            selected_categories = []
            for i in range(self.clbCategories.Items.Count):
                if self.clbCategories.GetItemChecked(i):
                    selected_categories.append(self.clbCategories.Items[i])
            
            if not selected_categories:
                MessageBox.Show("Please select at least one category to process.", "Validation Error")
                return
                
            # Save settings
            global DEBUG_MODE, TEST_MODE, TEST_ITEM_COUNT, TEST_LINK_ITEM_COUNT
            global OUTPUT_TO_EXCEL, INCLUDE_LINKED_MODELS, CATEGORIES_TO_CHECK, EXCEL_CONFIG_PATH
            
            DEBUG_MODE = self.cbDebugMode.Checked
            TEST_MODE = self.cbTestMode.Checked
            TEST_ITEM_COUNT = test_item_count
            TEST_LINK_ITEM_COUNT = test_link_item_count
            OUTPUT_TO_EXCEL = self.cbOutputToExcel.Checked
            INCLUDE_LINKED_MODELS = self.cbIncludeLinkedModels.Checked
            EXCEL_CONFIG_PATH = excel_path
            CATEGORIES_TO_CHECK = tuple(selected_categories)
            
            # Save to file
            settings = {
                "DEBUG_MODE": DEBUG_MODE,
                "TEST_MODE": TEST_MODE,
                "TEST_ITEM_COUNT": TEST_ITEM_COUNT,
                "TEST_LINK_ITEM_COUNT": TEST_LINK_ITEM_COUNT,
                "OUTPUT_TO_EXCEL": OUTPUT_TO_EXCEL,
                "INCLUDE_LINKED_MODELS": INCLUDE_LINKED_MODELS,
                "EXCEL_CONFIG_PATH": EXCEL_CONFIG_PATH,
                "CATEGORIES_TO_CHECK": list(CATEGORIES_TO_CHECK)
            }
            save_settings(settings)
            
            self.DialogResult = DialogResult.OK
            self.Close()
        except ValueError:
            MessageBox.Show("Please enter valid numbers for item counts.", "Validation Error")
            
    def OnCancelButtonClick(self, sender, args):
        """Closes the form without saving."""
        self.DialogResult = DialogResult.Cancel
        self.Close()

# Define Revit built-in parameter names for level parameters
LEVEL_PARAM_NAMES = {
    "LEVEL_PARAM": BuiltInParameter.LEVEL_PARAM,  # "Level" parameter
    "SCHEDULE_LEVEL_PARAM": BuiltInParameter.SCHEDULE_LEVEL_PARAM,  # "Reference Level" parameter  
    "FAMILY_LEVEL_PARAM": BuiltInParameter.FAMILY_LEVEL_PARAM,  # "Base Level" parameter
    "INSTANCE_SCHEDULE_ONLY_LEVEL_PARAM": BuiltInParameter.INSTANCE_SCHEDULE_ONLY_LEVEL_PARAM,  # "Base Constraint" parameter
}

# Pre-defined level parameters (static data from the table headers)
LEVEL_PARAMETERS = [
    "Level ElementId Instance Constraints",
    "Reference Level ElementId Instance Constraints",
    "Base Level ElementId Instance Constraints", 
    "Base Constraint ElementId Instance Constraints"
]

# Mapping between the table headers and Revit built-in parameters
PARAM_MAPPING = {
    "Level ElementId Instance Constraints": LEVEL_PARAM_NAMES["LEVEL_PARAM"],
    "Reference Level ElementId Instance Constraints": LEVEL_PARAM_NAMES["SCHEDULE_LEVEL_PARAM"],
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

# Updated to filter categories from the Excel file and load elements only for selected categories
# Improved parameter lookup logic to ensure correct handling based on table headers

# Filter categories from Excel based on CATEGORIES_TO_CHECK
def read_excel_data():
    """Reads the Excel file containing level parameter configuration for categories."""
    debug_print("Starting to read Excel data...")
    excel_path = EXCEL_CONFIG_PATH
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
            if not category_name or (CATEGORIES_TO_CHECK and category_name not in CATEGORIES_TO_CHECK):
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

# Load elements only for selected categories, including linked models if enabled
def get_all_elements_by_category():
    """Gets all elements in the model, grouped by category name."""
    debug_print("Collecting all elements by category...")
    elements_by_category = defaultdict(list)
    
    start_time = time.time()
    
    # Dictionary to track elements from linked models
    elements_from_links = {}
    linked_models = {}
    
    # Get all linked models if enabled
    if INCLUDE_LINKED_MODELS:
        debug_print("Collecting linked models...")
        linked_docs = {}
        
        # Get all RevitLinkInstance elements
        link_collector = FilteredElementCollector(doc).OfClass(RevitLinkInstance)
        links = list(link_collector)
        debug_print("Found {0} linked models".format(len(links)))
        
        # Get the link documents
        for link in links:
            if link.GetLinkDocument():
                link_doc = link.GetLinkDocument()
                link_name = link_doc.Title
                linked_docs[link_name] = link_doc
                linked_models[link_name] = link
                debug_print("Found linked model: {0}".format(link_name))
    
    # Get all elements in the main model
    collector = FilteredElementCollector(doc).WhereElementIsNotElementType()
    
    count = 0
    for element in collector:
        try:
            # Skip non-categorized elements
            if not element.Category:
                continue
                
            category_name = element.Category.Name
            if CATEGORIES_TO_CHECK and category_name not in CATEGORIES_TO_CHECK:
                continue
                
            # Add the element with origin information
            element_info = {
                "element": element,
                "origin": "Main Model",
                "host_name": doc.Title,
                "link_transform": None
            }
            
            elements_by_category[category_name].append(element_info)
            count += 1
            
            if count % 1000 == 0:
                debug_print("Processed {0} elements in main model...".format(count))
        except Exception as e:
            debug_print("Error processing element in main model: " + str(e))
            continue
    
    debug_print("Collected {0} elements in {1} categories from main model".format(
        count, len(elements_by_category)))
    
    # Process linked models if enabled
    if INCLUDE_LINKED_MODELS and linked_docs:
        link_count = 0
        
        for link_name, link_doc in linked_docs.items():
            debug_print("Processing linked model: {0}".format(link_name))
            link_transform = linked_models[link_name].GetTotalTransform()
            
            # Get all elements in the linked model
            link_collector = FilteredElementCollector(link_doc).WhereElementIsNotElementType()
            
            link_elements = 0
            for element in link_collector:
                try:
                    # Skip non-categorized elements
                    if not element.Category:
                        continue
                        
                    category_name = element.Category.Name
                    if CATEGORIES_TO_CHECK and category_name not in CATEGORIES_TO_CHECK:
                        continue
                    
                    # Check test mode limit for linked models
                    if TEST_MODE:
                        category_elements = [e for e in elements_by_category[category_name] 
                                           if e["origin"] == "Linked Model" and 
                                              e["host_name"] == link_name]
                        if len(category_elements) >= TEST_LINK_ITEM_COUNT:
                            continue
                    
                    # Add the element with origin information
                    element_info = {
                        "element": element,
                        "origin": "Linked Model",
                        "host_name": link_name,
                        "link_transform": link_transform
                    }
                    
                    elements_by_category[category_name].append(element_info)
                    link_elements += 1
                    link_count += 1
                    
                    if link_count % 1000 == 0:
                        debug_print("Processed {0} elements from linked models...".format(link_count))
                except Exception as e:
                    debug_print("Error processing element in linked model {0}: {1}".format(link_name, str(e)))
                    continue
            
            debug_print("Collected {0} elements from linked model: {1}".format(link_elements, link_name))
    
    end_time = time.time()
    total_count = count + (link_count if INCLUDE_LINKED_MODELS else 0)
    debug_print("Collected {0} elements in {1} categories in {2:.2f} seconds".format(
        total_count, len(elements_by_category), end_time - start_time))
    
    # Apply test mode limits for main model
    if TEST_MODE:
        for category_name in elements_by_category.keys():
            # Separate main model and linked model elements
            main_elements = [e for e in elements_by_category[category_name] if e["origin"] == "Main Model"]
            link_elements = [e for e in elements_by_category[category_name] if e["origin"] == "Linked Model"]
            
            # Limit main model elements
            if len(main_elements) > TEST_ITEM_COUNT:
                debug_print("Limiting category {0} from {1} to {2} elements (main model)".format(
                    category_name, len(main_elements), TEST_ITEM_COUNT))
                main_elements = main_elements[:TEST_ITEM_COUNT]
            
            # Keep link elements as they've already been limited during collection
            
            # Combine the lists
            elements_by_category[category_name] = main_elements + link_elements
    
    return elements_by_category

# Ensure level parameters are found based on table headers - Fixed linked file handling
def get_element_levels(element_info, category_info, level_param_names, all_levels):
    """Gets all level information for an element based on its parameters and category configuration."""
    element_levels = []
    level_params_found = []
    
    # Extract element and origin information
    element = element_info["element"]
    origin = element_info["origin"]
    host_name = element_info["host_name"]
    link_transform = element_info["link_transform"]
    
    element_id = element.Id.IntegerValue
    debug_print("Processing element ID: {0} from {1}: {2}".format(element_id, origin, host_name))
    
    # Get the document that contains the element - use the element's document directly
    # for linked elements rather than trying to open the file
    if origin == "Main Model":
        element_doc = doc
    else:
        # For linked elements, use the document the element already belongs to
        element_doc = element.Document
    
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
                    # Get the level from the document that contains the element
                    level = element_doc.GetElement(level_id)
                    if level and isinstance(level, Level):
                        # Transform levels from linked models to main model coordinates if needed
                        if origin == "Linked Model" and link_transform:
                            # We need to transform the level's location
                            # For now, we'll just note the level and its source
                            element_levels.append(level)
                            level_params_found.append(param_name + " (from " + host_name + ")")
                        else:
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
                # Transform bounding box if from linked model
                if origin == "Linked Model" and link_transform:
                    # Transform the bounding box coordinates to main model coordinates
                    min_point = link_transform.OfPoint(bbox.Min)
                    max_point = link_transform.OfPoint(bbox.Max)
                    z_min = min_point.Z
                    z_max = max_point.Z
                    z_coord = (z_min + z_max) / 2.0
                else:
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
    for category_name, element_infos in elements_by_category.items():
        # Skip categories not in CATEGORIES_TO_CHECK if it's defined
        if CATEGORIES_TO_CHECK and category_name not in CATEGORIES_TO_CHECK:
            debug_print("Category {0} is not in the list of categories to check, skipping".format(category_name))
            continue

        debug_print("Processing category: {0} with {1} elements".format(category_name, len(element_infos)))

        if category_name not in category_data:
            debug_print("Category {0} not found in Excel configuration, skipping".format(category_name))
            continue

        category_info = category_data[category_name]

        # Skip ignored categories
        if category_info["ignore"]:
            debug_print("Category {0} is marked as ignored, skipping".format(category_name))
            results["ignored"].append((category_name, len(element_infos)))
            continue

        # Process each element in the category
        for element_info in element_infos:
            element = element_info["element"]
            origin = element_info["origin"]
            host_name = element_info["host_name"]
            
            element_id = element.Id.IntegerValue
            element_levels, level_params_found = get_element_levels(element_info, category_info, level_param_names, all_levels)

            # Build location info for reporting
            location_info = origin + ": " + host_name

            # No level parameters found
            if not element_levels:
                debug_print("Element {0} has no level information".format(element_id))
                results["no_level"].append((category_name, element.Id, location_info))
                category_results[category_name].append({
                    "id": element_id, 
                    "status": "No Level", 
                    "level": "None", 
                    "params": "None",
                    "location": location_info
                })
                continue

            # One level parameter found
            if len(element_levels) == 1:
                level_name = element_levels[0].Name
                param_name = level_params_found[0] if level_params_found else "Unknown"

                # Check if this was determined by Z-coordinate
                if "Z-Coordinate" in param_name:
                    results["z_coord"].append((category_name, element.Id, location_info))
                else:
                    results["single_param"].append((category_name, element.Id, level_name, param_name, location_info))

                debug_print("Element {0} has single level parameter: {1} from {2}".format(
                    element_id, level_name, param_name))
                category_results[category_name].append({
                    "id": element_id, 
                    "status": "Single Parameter", 
                    "level": level_name, 
                    "params": param_name,
                    "location": location_info
                })
                continue

            # Multiple level parameters found - check if they agree
            level_names = [level.Name for level in element_levels]
            if len(set(level_names)) == 1:
                # All level parameters agree
                level_name = element_levels[0].Name
                debug_print("Element {0} has multiple agreeing level parameters: {1}".format(
                    element_id, level_name))
                results["multi_param_agree"].append((category_name, element.Id, level_name, ", ".join(level_params_found), location_info))
                category_results[category_name].append({
                    "id": element_id, 
                    "status": "Multiple Parameters - Agree", 
                    "level": level_name, 
                    "params": ", ".join(level_params_found),
                    "location": location_info
                })
            else:
                # Level parameters disagree
                level_names_str = ", ".join(level_names)
                debug_print("Element {0} has multiple disagreeing level parameters: {1}".format(
                    element_id, level_names_str))
                results["multi_param_disagree"].append((category_name, element.Id, level_names_str, ", ".join(level_params_found), location_info))
                category_results[category_name].append({
                    "id": element_id, 
                    "status": "Multiple Parameters - Disagree", 
                    "level": level_names_str, 
                    "params": ", ".join(level_params_found),
                    "location": location_info
                })

    end_time = time.time()
    debug_print("Finished processing elements in {0:.2f} seconds".format(end_time - start_time))

    return results, category_results

# Fix for Excel Export: Ensure at least one visible sheet remains
# Fix for ValueError: Add a check for empty data before calling print_table

# Updated export_to_excel to handle invalid sheet names and duplicates

def export_to_excel(category_results):
    """Exports the results to an Excel file with a sheet for each category."""
    if not EXCEL_AVAILABLE:
        debug_print("Excel is not available for export")
        return

    debug_print("Exporting results to Excel...")

    # Create a unique filename based on the current timestamp
    timestamp = time.strftime("%Y%m%d_%H%M%S")
    user_desktop = os.path.join(os.path.expanduser("~"), "Desktop")
    filename = "RevitLevels_{0}.xlsx".format(timestamp)
    excel_path = os.path.join(user_desktop, filename)
    
    # Get current directory as fallback
    script_dir = os.path.dirname(os.path.realpath(__file__))
    fallback_path = os.path.join(script_dir, filename)
    
    # Another fallback to temp directory
    temp_dir = os.environ.get('TEMP', os.path.join(os.path.expanduser("~"), "Temp"))
    temp_path = os.path.join(temp_dir, filename)
    
    debug_print("Will attempt to save to: {0}".format(excel_path))
    debug_print("Fallback path 1: {0}".format(fallback_path))
    debug_print("Fallback path 2: {0}".format(temp_path))

    excel = None
    workbook = None

    try:
        # Start Excel application
        excel = ExcelApp()
        excel.Visible = False
        excel.DisplayAlerts = False
        
        # Kill any existing Excel processes with same filename to prevent locks
        try:
            import subprocess
            subprocess.call("taskkill /f /im excel.exe", shell=True)
            debug_print("Terminated any existing Excel processes")
            time.sleep(1)  # Wait for Excel to fully close
            
            # Restart Excel
            excel = ExcelApp()
            excel.Visible = False
            excel.DisplayAlerts = False
        except:
            debug_print("Could not terminate Excel processes, continuing...")

        # Create a new workbook
        workbook = excel.Workbooks.Add()
        
        # Get initial sheet count
        sheet_count = workbook.Sheets.Count
        debug_print("Initial workbook has {0} sheets".format(sheet_count))
        
        # Create a summary sheet first
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

        # Track used sheet names to avoid duplicates
        used_sheet_names = {}  # Using dict instead of set for py2.7 compatibility

        # Add data for each category
        row = 2
        for category_name, elements in category_results.items():
            debug_print("Adding Excel sheet for category: {0}".format(category_name))

            # Ensure sheet name is valid and unique (max 31 chars for Excel)
            # Remove invalid characters for Excel sheet names
            invalid_chars = ['/', '\\', '?', '*', '[', ']', ':', '\'']
            valid_category_name = category_name
            for char in invalid_chars:
                valid_category_name = valid_category_name.replace(char, '_')
                
            sheet_name = valid_category_name[:31]  # Truncate to 31 characters
            original_sheet_name = sheet_name
            suffix = 1
            while sheet_name in used_sheet_names:
                sheet_name = "{0}_{1}".format(original_sheet_name[:28], suffix)  # Ensure uniqueness
                suffix += 1
            used_sheet_names[sheet_name] = True

            # Create a sheet for this category
            cat_sheet = workbook.Sheets.Add()
            cat_sheet.Name = sheet_name

            # Add headers to category sheet
            cat_sheet.Cells[1, 1].Value = "Element ID"
            cat_sheet.Cells[1, 2].Value = "Status"
            cat_sheet.Cells[1, 3].Value = "Level"
            cat_sheet.Cells[1, 4].Value = "Parameters"
            cat_sheet.Cells[1, 5].Value = "Location"

            # Format the headers
            cat_header_range = cat_sheet.Range["A1:E1"]
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
                cat_sheet.Cells[i+2, 5].Value = element_data["location"]

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
        
        # Now delete the default sheets only after creating all our sheets
        # Get updated sheet count and delete the default sheets
        # Start from the end to avoid index shifting issues
        for i in range(workbook.Sheets.Count, 0, -1):
            try:
                sheet = workbook.Sheets.Item[i]
                if sheet.Name != "Summary" and sheet.Name not in used_sheet_names:
                    debug_print("Deleting default sheet: {0}".format(sheet.Name))
                    sheet.Delete()
            except:
                debug_print("Error trying to delete sheet at index {0}".format(i))

        # Try different save approaches
        save_success = False
        save_path = ""
        error_msg = ""
        
        # Approach 1: Try main path with direct SaveAs
        try:
            debug_print("Trying to save to: {0}".format(excel_path))
            workbook.SaveAs(excel_path)
            save_success = True
            save_path = excel_path
            debug_print("Successfully saved to main path")
        except Exception as e:
            error_msg = str(e)
            debug_print("Error saving to main path: {0}".format(error_msg))
        
        # Approach 2: Try fallback path with direct SaveAs 
        if not save_success:
            try:
                debug_print("Trying to save to fallback path: {0}".format(fallback_path))
                workbook.SaveAs(fallback_path)
                save_success = True
                save_path = fallback_path
                debug_print("Successfully saved to fallback path")
            except Exception as e:
                error_msg = str(e)
                debug_print("Error saving to fallback path: {0}".format(error_msg))
        
        # Approach 3: Try temp directory
        if not save_success:
            try:
                debug_print("Trying to save to temp path: {0}".format(temp_path))
                workbook.SaveAs(temp_path)
                save_success = True
                save_path = temp_path
                debug_print("Successfully saved to temp path")
            except Exception as e:
                error_msg = str(e)
                debug_print("Error saving to temp path: {0}".format(error_msg))
        
        # Approach 4: Let Excel choose path with SaveAs dialog
        if not save_success:
            try:
                debug_print("Falling back to Excel's SaveAs dialog")
                excel.Visible = True
                excel.DisplayAlerts = True
                workbook.SaveAs()  # This will prompt user with SaveAs dialog
                save_success = True
                save_path = "User selected location"
                debug_print("Successfully saved with user dialog")
            except Exception as e:
                error_msg = str(e)
                debug_print("Error with SaveAs dialog: {0}".format(error_msg))

        # Clean up
        try:
            workbook.Close(save_success)
        except:
            debug_print("Error closing workbook")
            
        try:
            excel.Quit()
        except:
            debug_print("Error quitting Excel")

        # Show success message if save was successful
        if save_success:
            forms.alert("Results exported to Excel file: {0}".format(save_path), title="Export Successful")
        else:
            forms.alert("Failed to save Excel file. Error: {0}".format(error_msg), title="Excel Export Error")

    except Exception as e:
        debug_print("Error in Excel export process: {0}".format(str(e)))
        forms.alert("Error in Excel export process: {0}".format(str(e)), title="Excel Export Error")
    finally:
        # Ensure proper cleanup of COM objects
        try:
            if workbook:
                try:
                    workbook.Close(False)
                except:
                    pass
        except:
            pass
            
        try:
            if excel:
                try:
                    excel.Quit()
                except:
                    pass
        except:
            pass

# Fix for ValueError in print_table
def safe_print_table(data, columns):
    """Safely prints a table, ensuring data is not empty."""
    if not data:
        output.print_md("**No data available to display.**")
    else:
        output.print_table(data, columns=columns)

def main():
    """Main function to run the level unification process."""
    # Show settings form
    settings_form = SettingsForm()
    result = settings_form.ShowDialog()
    if result == DialogResult.Cancel:
        return
        
    debug_print("Starting Level Unification Process")
    if TEST_MODE:
        debug_print("*** TEST MODE ENABLED - Processing only {0} elements per category ***".format(TEST_ITEM_COUNT))
        debug_print("*** TEST MODE ENABLED - Processing only {0} elements per linked model ***".format(TEST_LINK_ITEM_COUNT))
        debug_print("*** TEST MODE ENABLED - Only processing categories: {0} ***".format(CATEGORIES_TO_CHECK))
    
    if INCLUDE_LINKED_MODELS:
        debug_print("*** LINKED MODELS ENABLED - Processing elements from linked models ***")
    
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
    safe_print_table(
        [[category, count] for category, count in results["ignored"]],
        columns=["Category", "Element Count"]
    )
    
    output.print_md("### Elements with Single Level Parameter")
    safe_print_table(
        [[category, element_id.IntegerValue, level_name, param_name, location] 
         for category, element_id, level_name, param_name, location in results["single_param"][:100]],  # Limit to 100 rows
        columns=["Category", "Element ID", "Level", "Parameter", "Location"]
    )
    
    output.print_md("### Elements with Multiple Agreeing Level Parameters")
    safe_print_table(
        [[category, element_id.IntegerValue, level_name, params, location] 
         for category, element_id, level_name, params, location in results["multi_param_agree"][:100]],
        columns=["Category", "Element ID", "Level", "Parameters", "Location"]
    )
    
    output.print_md("### Elements with Disagreeing Level Parameters")
    safe_print_table(
        [[category, element_id.IntegerValue, level_names, params, location] 
         for category, element_id, level_names, params, location in results["multi_param_disagree"][:100]],
        columns=["Category", "Element ID", "Levels", "Parameters", "Location"]
    )
    
    output.print_md("### Elements Assigned by Z-Coordinate")
    safe_print_table(
        [[category, element_id.IntegerValue, location] 
         for category, element_id, location in results["z_coord"][:100]],
        columns=["Category", "Element ID", "Location"]
    )
    
    output.print_md("### Elements with No Level Assigned")
    safe_print_table(
        [[category, element_id.IntegerValue, location] 
         for category, element_id, location in results["no_level"][:100]],
        columns=["Category", "Element ID", "Location"]
    )
    
    debug_print("Level Unification Process Completed")

if __name__ == "__main__":
    main()