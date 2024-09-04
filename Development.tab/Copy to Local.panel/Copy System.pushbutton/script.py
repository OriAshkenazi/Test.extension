#! python3

import clr
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')

from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *
from System.Collections.Generic import List
from System.Windows.Forms import *
from System.Drawing import *
from System import Object


import sys
import traceback

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument


# Define the IDuplicateTypeNamesHandler implementation
class SilentTypeConversionHandler(object):
    __interface__ = IDuplicateTypeNamesHandler

    def OnDuplicateTypeNamesFound(self, args):
        print("OnDuplicateTypeNamesFound called")
        print(f"Args type: {type(args)}")
        return DuplicateTypeAction.UseDestinationTypes


def get_all_links():
    collector = FilteredElementCollector(doc).OfClass(RevitLinkInstance)
    return {link.Name: link for link in collector}

def get_linked_document(link_instance):
    return link_instance.GetLinkDocument()

def get_all_system_types(linked_doc):
    collector = FilteredElementCollector(linked_doc).WhereElementIsNotElementType()
    system_types = set()
    for elem in collector:
        param = elem.get_Parameter(BuiltInParameter.RBS_SYSTEM_CLASSIFICATION_PARAM)
        if param and param.AsString():
            system_types.add(param.AsString())
    return sorted(list(system_types))

def get_element_ids_by_system_type(linked_doc, system_type):
    collector = FilteredElementCollector(linked_doc).WhereElementIsNotElementType()
    element_ids = [elem.Id for elem in collector if has_matching_system_type(elem, system_type)]
    print(f"Found {len(element_ids)} elements for System Type: {system_type}")
    if element_ids:
        print(f"First few element IDs: {', '.join(str(id.IntegerValue) for id in element_ids[:5])}")
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

def filter_copyable_elements(linked_doc, element_ids):
    copyable_ids = []
    non_copyable_reasons = {}
    for id in element_ids:
        element = linked_doc.GetElement(id)
        if element is None:
            non_copyable_reasons[id] = "Element not found"
        elif element.Category is None:
            non_copyable_reasons[id] = "No category"
        elif not element.Category.AllowsBoundParameters:
            non_copyable_reasons[id] = "Category doesn't allow bound parameters"
        else:
            copyable_ids.append(id)
    
    print(f"Filtered {len(copyable_ids)} copyable elements out of {len(element_ids)} total elements")
    if len(non_copyable_reasons) > 0:
        print(f"Examples of non-copyable elements:")
        for id, reason in list(non_copyable_reasons.items())[:5]:
            print(f"  ElementId {id}: {reason}")
    
    return copyable_ids

def calculate_transformation(source_doc, dest_doc):
    source_position = source_doc.ActiveProjectLocation.GetProjectPosition(XYZ.Zero)
    dest_position = dest_doc.ActiveProjectLocation.GetProjectPosition(XYZ.Zero)
    
    source_point = XYZ(source_position.EastWest, source_position.NorthSouth, source_position.Elevation)
    dest_point = XYZ(dest_position.EastWest, dest_position.NorthSouth, dest_position.Elevation)
    
    # Calculate the translation vector manually
    translation = XYZ(
        dest_point.X - source_point.X,
        dest_point.Y - source_point.Y,
        dest_point.Z - source_point.Z
    )
    
    return Transform.CreateTranslation(translation)

def copy_elements(source_doc, element_ids, dest_doc, new_workset):
    if not element_ids:
        print("No elements to copy. Skipping this batch.")
        return []

    print(f"copy_elements received {len(element_ids)} ElementIds")
    
    # Create a List[ElementId] manually
    element_id_list = List[ElementId]()
    for id in element_ids:
        element_id_list.Add(id)
    
    print(f"Created .NET List with {element_id_list.Count} ElementIds")
    
    if element_id_list.Count > 0:
        print(f"First few ElementIds: {', '.join(str(id.IntegerValue) for id in list(element_id_list)[:5])}")
    else:
        print("ElementId list is empty.")
        return []

    options = CopyPasteOptions()
    print(f"CopyPasteOptions created: {options}")
    
    # Remove the SetDuplicateTypeNamesHandler call for now
    # We'll handle duplicate types manually if needed

    # Calculate the correct transformation
    transform = calculate_transformation(source_doc, dest_doc)
    print(f"Calculated transformation: {transform}")

    try:
        print("Attempting to copy elements")
        new_ids = ElementTransformUtils.CopyElements(source_doc, element_id_list, dest_doc, transform, options)
        print(f"Successfully copied {len(new_ids)} elements.")
        
        for new_id in new_ids:
            elem = dest_doc.GetElement(new_id)
            workset_param = elem.get_Parameter(BuiltInParameter.ELEM_PARTITION_PARAM)
            workset_param.Set(new_workset.Id.IntegerValue)
        
        return new_ids
    except Exception as e:
        print(f"Error during CopyElements: {str(e)}")
        print(f"Error type: {type(e).__name__}")
        print("Traceback:")
        import traceback
        traceback.print_exc()
        return []

def batch_copy_elements(source_doc, element_ids, dest_doc, new_workset, initial_batch_size=10):
    all_new_ids = []
    failed_ids = []
    
    def copy_batch(batch, batch_size):
        print(f"Attempting to copy batch of {len(batch)} elements")
        new_ids = copy_elements(source_doc, batch, dest_doc, new_workset)
        if new_ids:
            all_new_ids.extend(new_ids)
        else:
            if batch_size > 1:
                print(f"Batch copy failed. Retrying with smaller batches.")
                mid = len(batch) // 2
                copy_batch(batch[:mid], batch_size // 2)
                copy_batch(batch[mid:], batch_size // 2)
            else:
                print(f"Failed to copy element {batch[0].IntegerValue}")
                failed_ids.append(batch[0])

    for i in range(0, len(element_ids), initial_batch_size):
        batch = element_ids[i:i+initial_batch_size]
        copy_batch(batch, initial_batch_size)

    print(f"Successfully copied {len(all_new_ids)} elements")
    if failed_ids:
        print(f"Failed to copy {len(failed_ids)} elements: {', '.join(str(id.IntegerValue) for id in failed_ids[:5])}")
    
    return all_new_ids

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
            print("No link selected. Exiting.")
            return

        linked_doc = get_linked_document(links[selected_link_name])
        
        system_types = get_all_system_types(linked_doc)
        
        selected_system_types = select_multiple_from_list(system_types, "Select System Types")
        if not selected_system_types:
            print("No system types selected. Exiting.")
            return

        for system_type in selected_system_types:
            print(f"\nProcessing System Type: {system_type}")
            
            element_ids = get_element_ids_by_system_type(linked_doc, system_type)
            
            if not element_ids:
                print(f"No elements found for System Type: {system_type}. Skipping.")
                continue
            
            workset_name = f"VDC - {system_type}"
            
            t = Transaction(doc, f"Copy Elements for {system_type}")
            t.Start()
            
            try:
                new_workset = create_or_clean_workset(doc, workset_name)
                copyable_ids = filter_copyable_elements(linked_doc, element_ids)
                if not copyable_ids:
                    print(f"No copyable elements found for System Type: {system_type}. Skipping.")
                    t.RollBack()
                    continue
                
                valid_ids = [id for id in copyable_ids if linked_doc.GetElement(id) is not None]
                print(f"Found {len(valid_ids)} valid ElementIds out of {len(copyable_ids)} copyable ElementIds")
                
                print(f"Passing {len(valid_ids)} valid ElementIds to batch_copy_elements")
                new_ids = batch_copy_elements(linked_doc, valid_ids, doc, new_workset)
                
                if new_ids:
                    print(f"Successfully copied {len(new_ids)} elements to workset '{workset_name}'")
                else:
                    print(f"No elements were copied for System Type: {system_type}")
                
                t.Commit()
            except Exception as e:
                t.RollBack()
                print(f"Error processing {system_type}: {str(e)}")
                print(f"Error type: {type(e).__name__}")
                print("Traceback:")
                import traceback
                traceback.print_exc()
                print("Transaction rolled back.")

        print("\nAll operations completed.")
    except Exception as e:
        print(f"\nAn error occurred: {str(e)}")
        print(f"Error type: {type(e).__name__}")
        print("Traceback:")
        import traceback
        traceback.print_exc()

if __name__ == '__main__':
    main()
