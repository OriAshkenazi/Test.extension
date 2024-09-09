#! python3

import clr
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')

from Autodesk.Revit.DB import *
from Autodesk.Revit.DB import IDuplicateTypeNamesHandler, DuplicateTypeAction
from Autodesk.Revit.UI import *
from Autodesk.Revit.Exceptions import ArgumentException
from System.Collections.Generic import List
from System.Windows.Forms import *
from System.Drawing import *
from System import Object, EventHandler, Array, Type, Func
from datetime import datetime
from tqdm import tqdm

import sys
import traceback
import logging
import os

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

# Define the log directory
log_dir = r"C:\Users\orias\Documents\exports\copy system\124"

# Ensure the directory exists
os.makedirs(log_dir, exist_ok=True)

# Set up logging
log_file = os.path.join(log_dir, f"revit_system_copy_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log")
logging.basicConfig(filename=log_file, level=logging.INFO,
                    format='%(asctime)s - %(levelname)s - %(message)s')

# Add a log message to confirm the log file location
logging.info(f"Log file created at: {log_file}")

def get_all_links():
    collector = FilteredElementCollector(doc).OfClass(RevitLinkInstance)
    return {link.Name: link for link in collector}

def get_linked_document(link_instance):
    return link_instance.GetLinkDocument()

def get_all_system_types(linked_doc):
    collector = FilteredElementCollector(linked_doc).WhereElementIsNotElementType()
    system_types = set()
    for elem in tqdm(collector, desc="Analyzing elements", unit="elem"):
        param = elem.get_Parameter(BuiltInParameter.RBS_SYSTEM_CLASSIFICATION_PARAM)
        if param and param.AsString():
            system_types.add(param.AsString())
    return sorted(list(system_types))

def get_element_ids_by_system_type(linked_doc, system_type):
    collector = FilteredElementCollector(linked_doc).WhereElementIsNotElementType()
    element_ids = [elem.Id for elem in tqdm(collector, desc=f"Finding elements for {system_type}", unit="elem") if has_matching_system_type(elem, system_type)]
    logging.info(f"Found {len(element_ids)} elements for System Type: {system_type}")
    if element_ids:
        logging.info(f"First few element IDs: {', '.join(str(id.IntegerValue) for id in element_ids[:5])}")
    return element_ids

def has_matching_system_type(elem, system_type):
    param = elem.get_Parameter(BuiltInParameter.RBS_SYSTEM_CLASSIFICATION_PARAM)
    if param is None:
        return False
    param_value = param.AsString()
    return param_value is not None and param_value == system_type

def create_or_clean_workset(doc, name):
    existing_workset = None
    for workset in FilteredWorksetCollector(doc).ToWorksets():
        if workset.Name == name:
            existing_workset = workset
            break
    
    if existing_workset:
        collector = FilteredElementCollector(doc)
        workset_filter = ElementWorksetFilter(existing_workset.Id)
        elements_in_workset = collector.WherePasses(workset_filter).ToElements()
        
        elemIds = [elem.Id for elem in elements_in_workset]
        if elemIds:
            doc.Delete(List[ElementId](elemIds))
        
        new_workset = existing_workset
    else:
        new_workset = Workset.Create(doc, name)
    
    return new_workset

def calculate_transformation(source_doc, dest_doc):
    link_instances = FilteredElementCollector(dest_doc).OfClass(RevitLinkInstance)
    source_link_instance = next((link for link in link_instances if link.GetLinkDocument().Title == source_doc.Title), None)

    if source_link_instance is None:
        raise Exception("Could not find the link instance in the destination document")

    return source_link_instance.GetTotalTransform()

def copy_system(source_doc, dest_doc, system_type, transform):
    '''
    Copy elements of a specific system type from source to destination document.
    '''
    element_ids = get_element_ids_by_system_type(source_doc, system_type)
    element_id_list = List[ElementId](element_ids)
    
    # Create a dictionary to store the mapping of old to new ids
    id_mapping = {}
    
    # Set up options for copy/paste
    options = CopyPasteOptions()
    options.SetDuplicateTypeNamesHandler(CopyUseDestination())
    
    # Perform the copy operation
    new_ids = ElementTransformUtils.CopyElements(source_doc, element_id_list, dest_doc, transform, options)
    
    # Create the id mapping
    for old_id, new_id in tqdm(zip(element_ids, new_ids), desc="Creating ID mapping", total=len(element_ids)):
        id_mapping[old_id] = new_id
    
    return id_mapping

def update_connections(dest_doc, id_mapping):
    '''
    Update connections and dependencies for copied elements.
    '''
    for old_id, new_id in tqdm(id_mapping.items(), desc="Updating connections"):
        new_elem = dest_doc.GetElement(new_id)
        if new_elem:
            # Update connections (this is a simplified example, you may need to add more specific logic)
            if hasattr(new_elem, "MEPModel") and new_elem.MEPModel:
                for connector in new_elem.MEPModel.ConnectorManager.Connectors:
                    if connector.IsConnected:
                        ref_connector = connector.GetConnectedConnector()
                        if ref_connector and ref_connector.Owner.Id in id_mapping:
                            connector.ConnectTo(ref_connector)

def copy_system_with_dependencies(source_doc, dest_doc, system_type, new_workset):
    '''
    Copy a system and update its dependencies.
    '''
    transform = calculate_transformation(source_doc, dest_doc)
    
    try:
        # Copy the system elements
        id_mapping = copy_system(source_doc, dest_doc, system_type, transform)
        
        # Update connections and dependencies
        update_connections(dest_doc, id_mapping)
        
        # Assign elements to the new workset
        for new_id in id_mapping.values():
            elem = dest_doc.GetElement(new_id)
            workset_param = elem.get_Parameter(BuiltInParameter.ELEM_PARTITION_PARAM)
            workset_param.Set(new_workset.Id.IntegerValue)
        
        logging.info(f"Successfully copied and updated system: {system_type}")
        return len(id_mapping)
    except Exception as e:
        logging.error(f"Error copying system {system_type}: {str(e)}")
        logging.error(f"Error type: {type(e).__name__}")
        logging.error("Traceback:", exc_info=True)
        return 0

def select_from_list(options, prompt):
    form = Form()
    form.Text = prompt
    form.Width = 400
    form.Height = 300

    listBox = ListBox()
    listBox.DataSource = options
    listBox.Dock = DockStyle.Fill

    button = Button()
    button.Text = "Select"
    button.Dock = DockStyle.Bottom

    form.Controls.Add(listBox)
    form.Controls.Add(button)

    selected_item = [None]

    def on_click(sender, args):
        if listBox.SelectedItem:
            selected_item[0] = listBox.SelectedItem
            form.Close()

    button.Click += on_click

    form.ShowDialog()
    return selected_item[0]

def select_multiple_from_list(options, prompt):
    form = Form()
    form.Text = prompt
    form.Width = 600
    form.Height = 450

    listBox = ListBox()
    listBox.DataSource = options
    listBox.SelectionMode = SelectionMode.MultiSimple
    listBox.Dock = DockStyle.Fill

    button = Button()
    button.Text = "Select"
    button.Dock = DockStyle.Bottom

    form.Controls.Add(listBox)
    form.Controls.Add(button)

    selected_items = []

    def on_click(sender, args):
        selected_items.extend(listBox.SelectedItems)
        form.Close()

    button.Click += on_click

    form.ShowDialog()
    return selected_items

def main():
    try:
        links = get_all_links()
        link_names = sorted(links.keys())
        
        selected_link_name = select_from_list(link_names, "Select a Revit Link")
        if selected_link_name is None:
            logging.warning("No link selected. Exiting.")
            return

        linked_doc = get_linked_document(links[selected_link_name])
        
        system_types = get_all_system_types(linked_doc)
        
        selected_system_types = select_multiple_from_list(system_types, "Select System Types")
        if not selected_system_types:
            logging.warning("No system types selected. Exiting.")
            return

        for system_type in tqdm(selected_system_types, desc="Processing system types"):
            logging.info(f"\nProcessing System Type: {system_type}")
            
            workset_name = f"VDC - {system_type}"
            
            t = Transaction(doc, f"Copy System: {system_type}")
            t.Start()
            
            try:
                new_workset = create_or_clean_workset(doc, workset_name)
                copied_count = copy_system_with_dependencies(linked_doc, doc, system_type, new_workset)
                
                if copied_count > 0:
                    logging.info(f"Successfully copied {copied_count} elements to workset '{workset_name}'")
                else:
                    logging.warning(f"No elements were copied for System Type: {system_type}")
                
                t.Commit()
            except Exception as e:
                t.RollBack()
                logging.error(f"Error processing {system_type}: {str(e)}")
                logging.error(f"Error type: {type(e).__name__}")
                logging.error("Traceback:", exc_info=True)
                logging.info("Transaction rolled back.")

        logging.info("\nAll operations completed.")
    except Exception as e:
        logging.error(f"\nAn error occurred: {str(e)}")
        logging.error(f"Error type: {type(e).__name__}")
        logging.error("Traceback:", exc_info=True)

if __name__ == '__main__':
    main()
