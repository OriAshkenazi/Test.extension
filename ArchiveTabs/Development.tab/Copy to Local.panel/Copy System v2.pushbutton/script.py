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
import time
import traceback
import logging
import os

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

def setup_logging():
    try:
        print("Starting logging setup...")
        log_dir = r"C:\Users\orias\Documents\exports\copy system\124"
        
        if not os.path.exists(log_dir):
            os.makedirs(log_dir, exist_ok=True)
        
        log_file = os.path.join(log_dir, f"revit_system_copy_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log")

        # Remove all existing handlers to reset the logger setup
        logger = logging.getLogger()
        if logger.hasHandlers():
            for handler in logger.handlers[:]:
                logger.removeHandler(handler)

        # Set up new handler
        logging.basicConfig(filename=log_file, level=logging.INFO,
                            format='%(levelname)s - %(message)s',
                            datefmt='%Y-%m-%d %H:%M:%S')
        
        logging.info("Log file created and writable")
        print(f"Log file created at: {log_file}")
        
        return True
    except Exception as e:
        print(f"An error occurred while setting up logging: {str(e)}")
        traceback.print_exc()
        return False

if not setup_logging():
    print("Failed to set up logging. The script will exit.")
    sys.exit()

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
            delete_list = List[ElementId]()
            for id in elemIds:
                delete_list.Add(id)
            doc.Delete(delete_list)
        
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

def get_element_ids_by_system_type(linked_doc, system_type):
    collector = FilteredElementCollector(linked_doc).WhereElementIsNotElementType()
    
    element_id_list = List[ElementId]()
    for elem in collector:
        param = elem.get_Parameter(BuiltInParameter.RBS_SYSTEM_CLASSIFICATION_PARAM)
        if param and param.AsString() == system_type:
            element_id_list.Add(elem.Id)
    
    logging.info(f"Found {element_id_list.Count} elements for System Type: {system_type}")
    if element_id_list.Count > 0:
        logging.info(f"First few element IDs: {', '.join(str(id.IntegerValue) for id in list(element_id_list)[:5])}")
    
    return element_id_list

def copy_elements(source_doc, element_ids, dest_doc, new_workset):
    options = CopyPasteOptions()
    new_ids = []
    
    try:
        transform = calculate_transformation(source_doc, dest_doc)
        new_element_ids = ElementTransformUtils.CopyElements(source_doc, element_ids, dest_doc, transform, options)
        
        for new_id in new_element_ids:
            elem = dest_doc.GetElement(new_id)
            if elem:
                try:
                    workset_param = elem.get_Parameter(BuiltInParameter.ELEM_PARTITION_PARAM)
                    if workset_param and not workset_param.IsReadOnly:
                        workset_param.Set(new_workset.Id.IntegerValue)
                    else:
                        logging.warning(f"Cannot set workset for element {new_id.IntegerValue}. Parameter is read-only or not found.")
                except InvalidOperationException as ioe:
                    logging.warning(f"Cannot set workset for element {new_id.IntegerValue}: {str(ioe)}")
            new_ids.append(new_id)
    except Exception as e:
        logging.warning(f"Failed to copy elements: {str(e)}")
    
    return new_ids

def batch_copy_elements(source_doc, element_ids, dest_doc, new_workset, initial_batch_size=10):
    all_new_ids = []
    failed_ids = []
    total_elements = len(element_ids)
    
    processed_count = 0
    start_time = time.time()

    def update_progress():
        elapsed_time = time.time() - start_time
        progress = processed_count / total_elements
        bar_length = 50
        filled_length = int(bar_length * progress)
        bar = '█' * filled_length + '-' * (bar_length - filled_length)
        
        sys.stdout.write(f'Progress: |{bar}| {processed_count}/{total_elements} '
                         f'Succeeded: {len(all_new_ids)}, Failed: {len(failed_ids)} '
                         f'Time: {elapsed_time:.2f}s\n')

    def copy_batch(batch, batch_size):
        nonlocal processed_count
        logging.info(f"Attempting to copy batch of {len(batch)} elements")
        batch_list = List[ElementId]()
        for id in batch:
            batch_list.Add(id)
        new_ids = copy_elements(source_doc, batch_list, dest_doc, new_workset)
        if new_ids:
            all_new_ids.extend(new_ids)
            processed_count += len(new_ids)
        else:
            if batch_size > 1:
                logging.warning(f"Batch copy failed. Retrying with smaller batches.")
                mid = len(batch) // 2
                copy_batch(batch[:mid], batch_size // 2)
                copy_batch(batch[mid:], batch_size // 2)
            elif batch:
                for element_id in batch:
                    logging.warning(f"Failed to copy element {element_id.IntegerValue}")
                    failed_ids.append(element_id)
                processed_count += len(batch)
            else:
                logging.warning("Empty batch encountered")
        
        update_progress()

    for i in range(0, len(element_ids), initial_batch_size):
        batch = element_ids[i:i+initial_batch_size]
        copy_batch(batch, initial_batch_size)

    sys.stdout.write('\n')  # Move to the next line after progress is complete
    logging.info(f"Successfully copied {len(all_new_ids)} elements")
    if failed_ids:
        logging.warning(f"Failed to copy {len(failed_ids)} elements: {', '.join(str(id.IntegerValue) for id in failed_ids[:5])}")
    
    return all_new_ids

def copy_system_with_dependencies(source_doc, dest_doc, system_type, new_workset):
    try:
        # Get element IDs by system type
        element_ids = get_element_ids_by_system_type(source_doc, system_type)
        
        if not element_ids:
            logging.warning(f"No elements found for System Type: {system_type}")
            return 0

        # Convert element_ids to a Python list
        element_ids_list = list(element_ids)

        # Copy the system elements
        new_ids = batch_copy_elements(source_doc, element_ids_list, dest_doc, new_workset)
        
        if not new_ids:
            logging.warning(f"No elements were copied for System Type: {system_type}")
            return 0

        # Update connections and dependencies
        update_connections(dest_doc, dict(zip(element_ids_list, new_ids)))
        
        logging.info(f"Successfully copied and updated system: {system_type}")
        return len(new_ids)
    except Exception as e:
        logging.error(f"Error copying system {system_type}: {str(e)}")
        logging.error(f"Error type: {type(e).__name__}")
        logging.error("Traceback:", exc_info=True)
        return 0

def update_connections(dest_doc, id_mapping):
    for old_id, new_id in id_mapping.items():
        new_elem = dest_doc.GetElement(new_id)
        if new_elem and hasattr(new_elem, "MEPModel") and new_elem.MEPModel:
            connectors = new_elem.MEPModel.ConnectorManager.Connectors if new_elem.MEPModel.ConnectorManager else []
            for connector in connectors:
                if connector.IsConnected:
                    ref_connectors = connector.AllRefs
                    for ref in ref_connectors:
                        if ref.Owner.Id in id_mapping:
                            try:
                                connector.ConnectTo(ref)
                            except Exception as e:
                                logging.warning(f"Failed to connect elements {new_id.IntegerValue} and {ref.Owner.Id.IntegerValue}: {str(e)}")

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
        logging.info("Starting main function")
        print("Starting main function")

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
            logging.info(f"Processing System Type: {system_type}")
            
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
                print(f"Error processing {system_type}: {str(e)}, check log file for details.")
                logging.error(f"Error type: {type(e).__name__}")
                logging.error("Traceback:", exc_info=True)
                logging.info("Transaction rolled back.")

        logging.info("All operations completed.")
        print("All operations completed.")
        logging.shutdown()
    except Exception as e:
        print(f"An error occurred in main: {str(e)}")
        print(f"Error type: {type(e).__name__}")
        print("Traceback:")
        traceback.print_exc()
        logging.error(f"An error occurred in main: {str(e)}", exc_info=True)
        logging.shutdown()

if __name__ == '__main__':
    main()
