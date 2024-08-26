#! python3

# import clr
# clr.AddReference('RevitAPI')
# from Autodesk.Revit.DB import *

# doc = __revit__.ActiveUIDocument.Document

# def print_level_names_and_ids(doc):
#     # Collect all levels in the document
#     levels = FilteredElementCollector(doc).OfClass(Level).ToElements()
    
#     # Sort levels by elevation
#     levels = sorted(levels, key=lambda l: l.Elevation)
    
#     # Print tuples of level names and IDs
#     print("Level Name and ID tuples:")
#     for level in levels:
#         name = level.Name
#         id = level.Id.IntegerValue
#         print(f"('{name}', {id})")


# print_level_names_and_ids(__revit__.ActiveUIDocument.Document)

# print(doc.PathName)

# import clr
# clr.AddReference('RevitAPI')
# from Autodesk.Revit.DB import *

# def safe_get_property(obj, prop_name):
#     try:
#         return getattr(obj, prop_name)
#     except Exception as e:
#         return f"Error accessing {prop_name}: {str(e)}"

# def debug_compound_ceiling(doc, element_id):
#     """
#     Debug function to investigate issues with compound ceiling elements.
    
#     Args:
#         doc (Document): The Revit document.
#         element_id (int): The ID of the compound ceiling element to debug.
#     """
#     try:
#         # Get the element
#         element = doc.GetElement(ElementId(element_id))
#         if not element:
#             print(f"No element found with ID {element_id}")
#             return

#         print(f"Debugging Compound Ceiling Element (ID: {element_id})")
#         print("-" * 50)

#         # Basic element info
#         print(f"Element Type: {type(element)}")
#         print(f"Category: {safe_get_property(element, 'Category').Name if safe_get_property(element, 'Category') else 'No Category'}")

#         # Check if it's a ceiling
#         print(f"Is Ceiling instance: {isinstance(element, Ceiling)}")

#         # Get element type
#         element_type = doc.GetElement(element.GetTypeId())
#         print(f"Element Type: {type(element_type)}")
#         print(f"Element Type ID: {element.GetTypeId()}")
#         print(f"Element Type Name: {safe_get_property(element_type, 'Name')}")

#         # Try to get family information
#         family = safe_get_property(element_type, 'Family')
#         print(f"Family: {family}")
#         if family:
#             print(f"Family Name: {safe_get_property(family, 'Name')}")

#         # Check for compound structure
#         if hasattr(element_type, 'GetCompoundStructure'):
#             try:
#                 compound_structure = element_type.GetCompoundStructure()
#                 if compound_structure:
#                     print("Compound Structure found:")
#                     layers = compound_structure.GetLayers()
#                     for i, layer in enumerate(layers):
#                         material = doc.GetElement(layer.MaterialId)
#                         material_name = safe_get_property(material, 'Name') if material else "No Material"
#                         print(f"  Layer {i+1}: Thickness = {layer.Width}, Material = {material_name}")
#                 else:
#                     print("No Compound Structure found")
#             except Exception as e:
#                 print(f"Error accessing Compound Structure: {str(e)}")
#         else:
#             print("GetCompoundStructure method not found")

#         # Try to get area
#         area_param = element.get_Parameter(BuiltInParameter.HOST_AREA_COMPUTED)
#         if area_param:
#             area = area_param.AsDouble() * 0.092903  # Convert to square meters
#             print(f"Area: {area:.2f} m²")
#         else:
#             print("No HOST_AREA_COMPUTED parameter found")
        
#         # Try to get volume
#         volume_param = element.get_Parameter(BuiltInParameter.HOST_VOLUME_COMPUTED)
#         if volume_param:
#             volume = volume_param.AsDouble() * 0.0283168  # Convert to cubic meters
#             print(f"Volume: {volume:.2f} m³")
#         else:
#             print("No HOST_VOLUME_COMPUTED parameter found")

#         # List all parameters
#         print("\nAll Parameters:")
#         for param in element.Parameters:
#             try:
#                 if param.HasValue:
#                     if param.StorageType == StorageType.Double:
#                         value = param.AsDouble()
#                     elif param.StorageType == StorageType.Integer:
#                         value = param.AsInteger()
#                     elif param.StorageType == StorageType.String:
#                         value = param.AsString()
#                     elif param.StorageType == StorageType.ElementId:
#                         value = param.AsElementId()
#                     else:
#                         value = "Unknown"
#                     print(f"  {param.Definition.Name}: {value}")
#             except Exception as e:
#                 print(f"  Error accessing parameter {param.Definition.Name}: {str(e)}")

#     except Exception as e:
#         print(f"Error in debug_compound_ceiling: {str(e)}")
#         import traceback
#         print(traceback.format_exc())

# # Usage
# doc = __revit__.ActiveUIDocument.Document  # Get the active document
# debug_compound_ceiling(doc, 10476805)  # Call the debug function with the specific element ID

import clr
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *

doc = __revit__.ActiveUIDocument.Document

def get_view_template_info(template_name):
    # Find the view template
    collector = FilteredElementCollector(doc).OfClass(View)
    template = next((v for v in collector if v.IsTemplate and v.Name == template_name), None)
    
    if not template:
        print(f"View template '{template_name}' not found.")
        return
    
    info = []
    info.append(f"View Template: {template_name}\n")
    
    # Basic properties
    info.append(f"Detail Level: {template.DetailLevel}")
    info.append(f"Visual Style: {template.DisplayStyle}")
    
    # Visibility/Graphics overrides
    visible_categories = []
    hidden_categories = []
    transparency_settings = []
    
    for category in doc.Settings.Categories:
        if category.CategoryType == CategoryType.Model:
            is_visible = not template.GetCategoryHidden(category.Id)
            if is_visible:
                visible_categories.append(category.Name)
            else:
                hidden_categories.append(category.Name)
            
            overrides = template.GetCategoryOverrides(category.Id)
            if overrides.Transparency > 0:
                transparency_settings.append(f"{category.Name}: {overrides.Transparency}%")
    
    info.append("\nVisible Categories:")
    info.extend(f"- {cat}" for cat in sorted(visible_categories))
    
    info.append("\nHidden Categories:")
    info.extend(f"- {cat}" for cat in sorted(hidden_categories))
    
    if transparency_settings:
        info.append("\nTransparency Settings:")
        info.extend(transparency_settings)
    
    # Filters
    filters = template.GetFilters()
    if filters:
        info.append("\nFilters:")
        for filter_id in filters:
            filter_element = doc.GetElement(filter_id)
            info.append(f"- {filter_element.Name}")
    
    # Other properties
    info.append("\nOther Properties:")
    for param in template.Parameters:
        if param.HasValue:
            if param.StorageType == StorageType.String:
                value = param.AsString()
            elif param.StorageType == StorageType.Double:
                value = param.AsDouble()
            elif param.StorageType == StorageType.Integer:
                value = param.AsInteger()
            else:
                value = "Unable to retrieve"
            info.append(f"- {param.Definition.Name}: {value}")
    
    return "\n".join(info)

# Usage
template_name = "VDC - HVAC ELEC verification"
template_info = get_view_template_info(template_name)

if template_info:
    print(template_info)
    
    # Optionally, save to a file
    with open("view_template_info.txt", "w") as f:
        f.write(template_info)
    print("\nInformation has been saved to 'view_template_info.txt'")
else:
    print("Failed to retrieve template information.")