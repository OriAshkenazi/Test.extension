#! python3

import clr
import os
import sys
import time
import traceback
import io
import math
from collections import defaultdict
from typing import Tuple, Dict, List, Callable
from functools import wraps, lru_cache

clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *
from System.Windows.Forms import FolderBrowserDialog, DialogResult
from System.Collections.Generic import List

import System
import datetime
import openpyxl
from openpyxl.styles import Font, Alignment
from tqdm import tqdm

# Get the Revit application and document
uidoc = __revit__.ActiveUIDocument
doc = uidoc.Document

def get_elements_by_type_ids(type_ids):
    elements = []
    for type_id in type_ids:
        collector = FilteredElementCollector(doc).OfClass(FamilyInstance)
        filter = ElementTypeFilter(ElementId(type_id))
        elements.extend(collector.WherePasses(filter).ToElements())
    return elements

def calculate_nearest_distance(source_elements, target_elements):
    results = []
    for source in source_elements:
        source_point = source.Location.Point
        min_distance = float('inf')
        nearest_target = None
        
        for target in target_elements:
            target_point = target.Location.Point
            distance = source_point.DistanceTo(target_point)
            if distance < min_distance:
                min_distance = distance
                nearest_target = target
        
        results.append({
            'source': source,
            'nearest_target': nearest_target,
            'distance': min_distance
        })
    
    return results

# Example usage
power_device_type_ids = [123, 456, 789]  # Replace with actual Type IDs for power-drawing devices
outlet_type_ids = [321, 654, 987]  # Replace with actual Type IDs for outlets

power_devices = get_elements_by_type_ids(power_device_type_ids)
outlets = get_elements_by_type_ids(outlet_type_ids)

results = calculate_nearest_distance(power_devices, outlets)

# Output results
for result in results:
    print(f"Device ID: {result['source'].Id}")
    print(f"Nearest Outlet ID: {result['nearest_target'].Id}")
    print(f"Distance: {result['distance']} feet")
    print("---")
