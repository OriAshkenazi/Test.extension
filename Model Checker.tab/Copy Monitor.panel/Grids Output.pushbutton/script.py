# -*- coding: utf-8 -*-
import sys
import clr

# Revit references
clr.AddReference("RevitServices")
from RevitServices.Persistence import DocumentManager
from RevitServices.Transactions import TransactionManager

clr.AddReference("RevitAPI")
from Autodesk.Revit.DB import *

clr.AddReference("RevitAPIUI")
from Autodesk.Revit.UI import *

# For XML writing
import xml.etree.ElementTree as ET
import os

doc = __revit__.ActiveUIDocument.Document
if doc is None:
    TaskDialog.Show("Error", "No active document found.")
    sys.exit(1)

def get_element_parameters(element):
    """
    Retrieve all relevant parameters from an element, including geometry if needed.
    Returns a dictionary with parameter-name -> value.
    """
    data = {}
    
    # Example 1: Basic parameters
    data["ElementId"] = str(element.Id.IntegerValue)
    if element.Category:
        data["Category"] = element.Category.Name
    else:
        data["Category"] = "NoCategory"

    # Typically valid for grids, levels, etc.
    data["Name"] = element.Name  
    
    # Example 2: If we want geometry or bounding box
    # (Caution: for large models, retrieving geometry can be expensive)
    bbox = element.get_BoundingBox(doc.ActiveView)
    if bbox:
        min_pt = bbox.Min
        max_pt = bbox.Max
        data["BoundingBox_Min"] = "({0}, {1}, {2})".format(min_pt.X, min_pt.Y, min_pt.Z)
        data["BoundingBox_Max"] = "({0}, {1}, {2})".format(max_pt.X, max_pt.Y, max_pt.Z)
    
    # Example 3: All built-in or shared parameters
    for param in element.Parameters:
        if not param.HasValue:
            continue
        def_name = param.Definition.Name
        
        # Try AsString(), fallback to AsValueString() or numeric conversions
        val_str = None
        try:
            val_str = param.AsString()
        except:
            pass
        
        if not val_str:
            try:
                val_str = param.AsValueString()
            except:
                pass
        
        if not val_str:
            # Fallback to AsDouble() or AsInteger() depending on param type
            if param.StorageType == StorageType.Double:
                val_str = str(param.AsDouble())
            elif param.StorageType == StorageType.Integer:
                val_str = str(param.AsInteger())
            else:
                val_str = "N/A"
        
        data[def_name] = val_str
    
    return data


def collect_elements(document):
    """
    Example collector for all Grid elements in the current document.
    If you want *all* elements, adapt the FilteredElementCollector accordingly.
    """
    # Verify document parameter
    if document is None:
        raise ValueError("Document parameter cannot be null")
        
    # For just Grids:
    collector = FilteredElementCollector(document).OfClass(Grid)
    
    elements_data = []
    for elem in collector:
        elem_params = get_element_parameters(elem)
        elements_data.append(elem_params)
    return elements_data


def save_data_to_xml(elements_data, output_path):
    """
    Save the collected element data to an XML file.
    """
    # Create the XML root
    root = ET.Element("ModelParameters")
    
    for elem_data in elements_data:
        # Each element is a node in the XML
        element_node = ET.SubElement(root, "Element")
        
        # Write each parameter as a sub-node or attribute
        for param_name, param_value in elem_data.items():
            param_node = ET.SubElement(element_node, "Parameter")
            param_node.set("Name", param_name)
            param_node.text = param_value
    
    # Construct the XML tree
    tree = ET.ElementTree(root)
    
    # Write it out to disk
    tree.write(output_path, encoding="utf-8", xml_declaration=True)


def main():
    try:
        # 1. Collect the elements (this sample collects Grids)
        # Pass the document explicitly
        elements_data = collect_elements(doc)
        
        # 2. Specify an output path for the XML
        output_dir = os.path.join(os.environ['TEMP'], "RevitExport")
        # Create directory if it doesn't exist
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)
            
        output_xml = os.path.join(output_dir, "model_parameters.xml")
        
        # 3. Save the data
        save_data_to_xml(elements_data, output_xml)
        
        # 4. Show completion dialog
        print("Export Complete",
                        "Exported element parameters to:\n{0}".format(output_xml))
                        
    except Exception as e:
        TaskDialog.Show("Error", "An error occurred:\n{0}".format(str(e)))


if __name__ == "__main__":
    main()
