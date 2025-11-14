# -*- coding: utf-8 -*-
"""High-level API for managing pyArea extensible storage data

Provides convenient methods for working with AreaSchemes, Sheets, AreaPlans, and Areas.
"""

import sys
import os

# Add schemas folder to path
lib_path = os.path.dirname(__file__)
schemas_path = os.path.join(lib_path, "schemas")
if schemas_path not in sys.path:
    sys.path.insert(0, schemas_path)

from pyrevit import DB
from schemas import schema_manager, municipality_schemas
import System


# Helper function for Revit 2026 compatibility - ElementId constructor changed
def create_element_id(int_value):
    """Create ElementId from integer - compatible with Revit 2024, 2025 and 2026+"""
    # In Revit 2026, ElementId has multiple overloads, need to specify Int64 explicitly
    if isinstance(int_value, DB.ElementId):
        return int_value
    # Convert to System.Int64 to resolve overload ambiguity
    return DB.ElementId(System.Int64(int_value))


# ==================== AreaScheme Methods ====================

def get_municipality(area_scheme):
    """Get municipality from AreaScheme element.
    
    Args:
        area_scheme: AreaScheme element
        
    Returns:
        str: Municipality name or None
    """
    data = schema_manager.get_data(area_scheme)
    return data.get("Municipality")


def set_municipality(area_scheme, municipality):
    """Set municipality on AreaScheme element.
    
    Args:
        area_scheme: AreaScheme element
        municipality: Municipality name ("Common", "Jerusalem", "Tel-Aviv")
        
    Returns:
        bool: True if successful
    """
    if municipality not in municipality_schemas.MUNICIPALITIES:
        print("Invalid municipality: {}".format(municipality))
        return False
    
    data = {"Municipality": municipality}
    return schema_manager.set_data(area_scheme, data)


def get_variant(area_scheme):
    """Get variant from AreaScheme element.
    
    Args:
        area_scheme: AreaScheme element
        
    Returns:
        str: Variant name or "Default"
    """
    data = schema_manager.get_data(area_scheme)
    return data.get("Variant", "Default")


def set_variant(area_scheme, variant):
    """Set variant on AreaScheme element.
    
    Args:
        area_scheme: AreaScheme element
        variant: Variant name
        
    Returns:
        bool: True if successful
    """
    # Get existing data to preserve Municipality
    data = schema_manager.get_data(area_scheme) or {}
    data["Variant"] = variant
    return schema_manager.set_data(area_scheme, data)


def get_municipality_and_variant(area_scheme):
    """Get both municipality and variant from AreaScheme element.
    
    Args:
        area_scheme: AreaScheme element
        
    Returns:
        tuple: (municipality, variant) or (None, "Default")
    """
    municipality = get_municipality(area_scheme)
    variant = get_variant(area_scheme)
    return municipality, variant


# ==================== Sheet Methods ====================

def get_sheet_data(sheet):
    """Get all data from Sheet element.
    
    Args:
        sheet: Sheet element
        
    Returns:
        dict: Sheet data
    """
    return schema_manager.get_data(sheet)


def set_sheet_data(sheet, data_dict, municipality):
    """Set data on Sheet element with validation.
    
    Args:
        sheet: Sheet element
        data_dict: Data dictionary
        municipality: Municipality name
        
    Returns:
        tuple: (success, error_messages)
    """
    # Validate data
    is_valid, errors = municipality_schemas.validate_data("Sheet", data_dict, municipality)
    
    if not is_valid:
        return False, errors
    
    # Store data
    success = schema_manager.set_data(sheet, data_dict)
    return success, [] if success else ["Failed to store data"]


# ==================== AreaPlan (View) Methods ====================

def get_areaplan_data(view):
    """Get all data from AreaPlan view element.
    
    Args:
        view: AreaPlan view element
        
    Returns:
        dict: AreaPlan data
    """
    return schema_manager.get_data(view)


def set_areaplan_data(view, data_dict, municipality):
    """Set data on AreaPlan view element with validation.
    
    Args:
        view: AreaPlan view element
        data_dict: Data dictionary
        municipality: Municipality name
        
    Returns:
        tuple: (success, error_messages)
    """
    # Validate data
    is_valid, errors = municipality_schemas.validate_data("AreaPlan", data_dict, municipality)
    
    if not is_valid:
        return False, errors
    
    # Store data
    success = schema_manager.set_data(view, data_dict)
    return success, [] if success else ["Failed to store data"]


# ==================== Area Methods ====================

def get_area_data(area):
    """Get all data from Area element.
    
    Args:
        area: Area element
        
    Returns:
        dict: Area data
    """
    return schema_manager.get_data(area)


def set_area_data(area, data_dict, municipality):
    """Set data on Area element with validation.
    
    Args:
        area: Area element
        data_dict: Data dictionary
        municipality: Municipality name
        
    Returns:
        tuple: (success, error_messages)
    """
    # Validate data
    is_valid, errors = municipality_schemas.validate_data("Area", data_dict, municipality)
    
    if not is_valid:
        return False, errors
    
    # Store data
    success = schema_manager.set_data(area, data_dict)
    return success, [] if success else ["Failed to store data"]


# ==================== General Methods ====================

def has_data(element):
    """Check if element has pyArea data.
    
    Args:
        element: Revit element
        
    Returns:
        bool: True if element has data
    """
    return schema_manager.has_data(element)


def delete_data(element):
    """Delete pyArea data from element.
    
    Args:
        element: Revit element
        
    Returns:
        bool: True if successful
    """
    return schema_manager.delete_data(element)


def get_data(element):
    """Get raw data from element (no validation).
    
    Args:
        element: Revit element
        
    Returns:
        dict: Data dictionary
    """
    return schema_manager.get_data(element)


def set_data(element, data_dict):
    """Set raw data on element (no validation).
    
    Args:
        element: Revit element
        data_dict: Data dictionary
        
    Returns:
        bool: True if successful
    """
    return schema_manager.set_data(element, data_dict)


# ==================== Helper Methods ====================

def get_area_scheme_by_id(doc, element_id):
    """Get AreaScheme element by ElementId.
    
    Args:
        doc: Revit document
        element_id: ElementId (as int or string)
        
    Returns:
        AreaScheme element or None
    """
    try:
        if isinstance(element_id, str):
            element_id = int(element_id)
        
        elem_id = create_element_id(element_id)
        element = doc.GetElement(elem_id)
        
        if isinstance(element, DB.AreaScheme):
            return element
        
        return None
        
    except:
        return None


def get_municipality_from_sheet(doc, sheet):
    """Get municipality from sheet by finding its AreaScheme.
    
    Args:
        doc: Revit document
        sheet: Sheet element
        
    Returns:
        str: Municipality name or None
    """
    sheet_data = get_sheet_data(sheet)
    area_scheme_id = sheet_data.get("AreaSchemeId")
    
    if not area_scheme_id:
        return None
    
    area_scheme = get_area_scheme_by_id(doc, area_scheme_id)
    
    if not area_scheme:
        return None
    
    return get_municipality(area_scheme)


def get_municipality_from_view(doc, view):
    """Get municipality and variant from AreaPlan view by getting its AreaScheme.
    
    For AreaPlan views, the AreaScheme is directly accessible via view parameters.
    
    Args:
        doc: Revit document
        view: View element (AreaPlan)
        
    Returns:
        tuple: (municipality, variant) or (None, "Default")
    """
    try:
        # For AreaPlan views, get the AreaScheme property directly
        if hasattr(view, 'AreaScheme'):
            area_scheme = view.AreaScheme
            
            if area_scheme:
                return get_municipality_and_variant(area_scheme)
        
        # Fallback: try to find via sheet (for other view types)
        view_id = view.Id
        collector = DB.FilteredElementCollector(doc)
        sheets = collector.OfClass(DB.ViewSheet).ToElements()
        
        for sheet in sheets:
            view_ids = sheet.GetAllPlacedViews()
            if view_id in view_ids:
                municipality = get_municipality_from_sheet(doc, sheet)
                # For backward compatibility, also get variant if available
                sheet_data = get_sheet_data(sheet)
                area_scheme_id = sheet_data.get("AreaSchemeId")
                if area_scheme_id:
                    area_scheme = doc.GetElement(DB.ElementId(int(area_scheme_id)))
                    if area_scheme:
                        variant = get_variant(area_scheme)
                        return municipality, variant
                return municipality, "Default"
        
    except Exception as e:
        # Keep error reporting for actual errors
        print("ERROR getting municipality from view: {}".format(e))
        import traceback
        traceback.print_exc()
    
    return None, "Default"


# ==================== Preferences Methods ====================

def get_preferences():
    """
    Load preferences from ProjectInformation.
    Returns default preferences if not found.
    
    Returns:
        dict: Preferences dictionary
    """
    try:
        from pyrevit import revit
        from export_utils import get_default_preferences
        
        doc = revit.doc
        proj_info = doc.ProjectInformation
        data = schema_manager.get_data(proj_info)
        
        if data and "Preferences" in data:
            return data["Preferences"]
        
        return get_default_preferences()
    except Exception as e:
        print("Warning: Failed to load preferences: {}".format(str(e)))
        from export_utils import get_default_preferences
        return get_default_preferences()


def set_preferences(preferences_dict):
    """
    Save preferences to ProjectInformation.
    
    Args:
        preferences_dict: Preferences dictionary to save
    
    Returns:
        bool: True if successful
    """
    try:
        from pyrevit import revit
        
        doc = revit.doc
        proj_info = doc.ProjectInformation
        
        # Get existing data or create new
        data = schema_manager.get_data(proj_info)
        if not data:
            data = {}
        
        # Update Preferences key
        data["Preferences"] = preferences_dict
        
        # Save back
        return schema_manager.set_data(proj_info, data)
    except Exception as e:
        print("ERROR: Failed to save preferences: {}".format(str(e)))
        return False
