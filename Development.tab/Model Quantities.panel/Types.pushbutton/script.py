import clr
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
import csv
import sys
import os
from System.Text import Encoding
import codecs
import traceback

from pyrevit import revit, forms
from Autodesk.Revit.DB import *

doc = revit.doc

def get_param_value(param):
    """Get parameter value as string"""
    if not param or not param.HasValue:
        return ""
    try:
        st = param.StorageType
        if st == StorageType.Double:
            return str(param.AsDouble())
        elif st == StorageType.Integer:
            return str(param.AsInteger())
        elif st == StorageType.String:
            return param.AsString() or ""
        elif st == StorageType.ElementId:
            eid = param.AsElementId()
            if eid and eid != ElementId.InvalidElementId:
                linked_elem = doc.GetElement(eid)
                if linked_elem:
                    return linked_elem.Name or str(eid.IntegerValue)
            return str(eid.IntegerValue) if eid else ""
    except Exception as e:
        print("Error getting parameter value: {}".format(str(e)))
    return ""

def get_all_parameters(element):
    """Get all parameters from an element"""
    params = []
    
    # Get standard parameters
    try:
        for p in element.Parameters:
            if p and p.HasValue and p.Definition:
                param_name = p.Definition.Name
                param_value = get_param_value(p)
                if param_value:
                    params.append((param_name, param_value))
    except Exception as e:
        print("Error getting standard parameters: {}".format(str(e)))

    # Get built-in parameters by category
    try:
        # Common parameters
        common_params = [
            BuiltInParameter.ALL_MODEL_TYPE_NAME,
            BuiltInParameter.ALL_MODEL_TYPE_MARK,
            BuiltInParameter.ALL_MODEL_DESCRIPTION,
            BuiltInParameter.ALL_MODEL_MANUFACTURER,
            BuiltInParameter.ALL_MODEL_MODEL,
            BuiltInParameter.ALL_MODEL_URL,
            BuiltInParameter.ALL_MODEL_COST,
            BuiltInParameter.SYMBOL_NAME_PARAM,
            BuiltInParameter.SYMBOL_FAMILY_NAME_PARAM,
            BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS,
            BuiltInParameter.ALL_MODEL_TYPE_COMMENTS
        ]
        
        # Dimensional parameters
        dimension_params = [
            BuiltInParameter.SYMBOL_HEIGHT_PARAM,
            BuiltInParameter.SYMBOL_WIDTH_PARAM,
            BuiltInParameter.SYMBOL_DEPTH_PARAM,
            BuiltInParameter.CASEWORK_HEIGHT,
            BuiltInParameter.CASEWORK_WIDTH,
            BuiltInParameter.CASEWORK_DEPTH
        ]
        
        # Material parameters
        material_params = [
            BuiltInParameter.ALL_MODEL_MATERIAL_NAME,
            BuiltInParameter.ALL_MODEL_MATERIAL_ID_PARAM,
            BuiltInParameter.MATERIAL_ID_PARAM,
            BuiltInParameter.MATERIAL_ASSET_PARAM_NAME
        ]
        
        # Family parameters
        family_params = [
            BuiltInParameter.FAMILY_LEVEL_PARAM,
            BuiltInParameter.FAMILY_BASE_LEVEL_PARAM,
            BuiltInParameter.FAMILY_TOP_LEVEL_PARAM,
            BuiltInParameter.FAMILY_NAME_PARAM,
            BuiltInParameter.FAMILY_WORK_PLANE_PARAM
        ]
        
        # Component parameters
        component_params = [
            BuiltInParameter.COMPONENT_DETAILS,
            BuiltInParameter.COMPONENT_CLASSIFICATION_PARAM,
            BuiltInParameter.COMPONENT_CODE,
            BuiltInParameter.COMPONENT_ID
        ]

        # Combine all parameter groups
        all_builtin_params = (
            common_params + 
            dimension_params + 
            material_params + 
            family_params + 
            component_params
        )
        
        for bip in all_builtin_params:
            try:
                param = element.get_Parameter(bip)
                if param and param.HasValue:
                    param_name = param.Definition.Name
                    param_value = get_param_value(param)
                    if param_value:
                        params.append((param_name, param_value))
            except:
                continue
    except Exception as e:
        print("Error getting built-in parameters: {}".format(str(e)))
        
    return params

try:
    print("Script starting...")
    
    # Get all elements by category
    print("\nCollecting categories...")
    categories = doc.Settings.Categories
    print("Found {} categories".format(len(list(categories))))
    
    all_type_data = []
    processed_categories = 0
    
    for cat in categories:
        try:
            if not cat.AllowsBoundParameters:
                continue
                
            processed_categories += 1
            print("\nProcessing category {}: {}".format(processed_categories, cat.Name))
            
            collector = FilteredElementCollector(doc).OfCategoryId(cat.Id).WhereElementIsElementType()
            types = list(collector.ToElements())
            
            if not types:
                continue
                
            print("Found {} types in {}".format(len(types), cat.Name))
            
            # Get instance counts for this category
            instance_collector = FilteredElementCollector(doc).OfCategoryId(cat.Id).WhereElementIsNotElementType()
            instances = list(instance_collector.ToElements())
            print("Found {} instances in {}".format(len(instances), cat.Name))
            
            type_counts = {}
            for inst in instances:
                try:
                    type_id = inst.GetTypeId()
                    if type_id != ElementId.InvalidElementId:
                        type_counts[type_id] = type_counts.get(type_id, 0) + 1
                except Exception as e:
                    print("Error counting instance: {}".format(str(e)))
                    continue
            
            # Process types
            print("\n" + "="*50)
            print("Processing Types")
            print("="*50)

            for t in types:
                try:
                    type_id = t.Id
                    count = type_counts.get(type_id, 0)
                    
                    # Get names
                    type_name = t.Name
                    try:
                        fam_name = t.FamilyName
                    except:
                        fam_name = cat.Name
                        
                    if not type_name:
                        continue
                        
                    # Get parameters
                    params = get_all_parameters(t)
                    
                    if params:
                        all_type_data.append((cat.Name, fam_name, type_name, count, params))
                        print("\nType: {}".format(type_name))
                        print("Family: {}".format(fam_name))
                        print("Category: {}".format(cat.Name))
                        print("Instance Count: {}".format(count))
                        print("Parameters: {}".format(len(params)))
                        
                except Exception as e:
                    print("Error processing type: {}".format(str(e)))
                    continue
                    
        except Exception as e:
            print("Error processing category {}: {}".format(cat.Name, str(e)))
            print(traceback.format_exc())
            continue

    print("\nProcessed {} categories".format(processed_categories))
    print("Found {} types with parameters".format(len(all_type_data)))

    if not all_type_data:
        forms.alert("No types found with parameters", exitscript=True)
        sys.exit()

    # Get save location
    save_folder = forms.pick_folder()
    if not save_folder:
        sys.exit()

    # Get all parameter names
    param_names = set()
    for _, _, _, _, params in all_type_data:
        param_names.update(name for name, _ in params)
    param_names = sorted(param_names)

    print("\nFound {} unique parameters".format(len(param_names)))
    
    # Write CSV
    csv_path = os.path.join(save_folder, "Export_Types_and_Parameters.csv")
    print("\nWriting to {}".format(csv_path))
    
    with codecs.open(csv_path, 'wb', encoding='utf-8-sig') as f:
        writer = csv.writer(f)
        
        # Write header
        header = ['Category', 'Family', 'Type', 'Instance Count'] + list(param_names)
        writer.writerow([unicode(x) for x in header])
        
        # Write data
        rows_written = 0
        for cat_name, fam_name, type_name, count, params in all_type_data:
            param_dict = dict(params)
            row = [cat_name, fam_name, type_name, count]
            row.extend(param_dict.get(param_name, "") for param_name in param_names)
            writer.writerow([unicode(x) for x in row])
            rows_written += 1
            
        print("Wrote {} rows".format(rows_written))

    forms.alert(
        "Export Complete!\n" +
        "Processed {} categories\n".format(processed_categories) +
        "Found {} types\n".format(len(all_type_data)) +
        "Wrote {} rows\n".format(rows_written) +
        "Saved to: {}".format(os.path.basename(csv_path))
    )
    
except Exception as e:
    print("Error in main script:")
    print(str(e))
    print(traceback.format_exc())
    forms.alert("An error occurred: {}".format(str(e)), exitscript=True)
