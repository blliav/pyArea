# -*- coding: utf-8 -*-
"""High-level API for managing pyArea extensible storage data

Provides convenient methods for working with AreaSchemes, Sheets, AreaPlans, and Areas.
"""

import sys
import os
import uuid

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


# ==================== Calculation Methods ====================

def generate_calculation_guid():
    """Generate a new unique GUID for a Calculation.
    
    Returns:
        str: UUID string (e.g., "a7b3c9d1-e5f2-4a8b-9c3d-1e2f3a4b5c6d")
    """
    return str(uuid.uuid4())


def get_all_calculations(area_scheme):
    """Get all Calculations from an AreaScheme element.
    
    Args:
        area_scheme: AreaScheme element
        
    Returns:
        dict: Dictionary of Calculations {calculation_guid: calculation_data}
    """
    data = schema_manager.get_data(area_scheme) or {}
    return data.get("Calculations", {})


def get_calculation(area_scheme, calculation_guid):
    """Get a specific Calculation by GUID from an AreaScheme.
    
    Args:
        area_scheme: AreaScheme element
        calculation_guid: Calculation GUID string
        
    Returns:
        dict: Calculation data or None if not found
    """
    calculations = get_all_calculations(area_scheme)
    return calculations.get(calculation_guid)


def set_calculation(area_scheme, calculation_guid, calculation_data, municipality):
    """Create or update a Calculation on an AreaScheme element.
    
    Args:
        area_scheme: AreaScheme element
        calculation_guid: Calculation GUID string
        calculation_data: Calculation data dictionary
        municipality: Municipality name
        
    Returns:
        tuple: (success, error_messages)
    """
    # Ensure Name is set
    if "Name" not in calculation_data:
        return False, ["Calculation must have a Name"]
    
    # Validate data
    is_valid, errors = municipality_schemas.validate_data("Calculation", calculation_data, municipality)
    
    if not is_valid:
        return False, errors
    
    # Get existing AreaScheme data
    data = schema_manager.get_data(area_scheme) or {}
    
    # Ensure Calculations dict exists
    if "Calculations" not in data:
        data["Calculations"] = {}
    
    # Store the calculation
    data["Calculations"][calculation_guid] = calculation_data
    
    # Save back to AreaScheme
    success = schema_manager.set_data(area_scheme, data)
    return success, [] if success else ["Failed to store Calculation"]


def delete_calculation(area_scheme, calculation_guid):
    """Delete a Calculation from an AreaScheme element.
    
    Args:
        area_scheme: AreaScheme element
        calculation_guid: Calculation GUID string
        
    Returns:
        bool: True if successful
    """
    data = schema_manager.get_data(area_scheme) or {}
    
    if "Calculations" not in data or calculation_guid not in data["Calculations"]:
        return False
    
    del data["Calculations"][calculation_guid]
    return schema_manager.set_data(area_scheme, data)


def get_calculation_from_sheet(doc, sheet):
    """Get Calculation data from a Sheet by resolving its CalculationGuid.
    
    Args:
        doc: Revit document
        sheet: Sheet element
        
    Returns:
        tuple: (area_scheme, calculation_data) or (None, None) if not found
    """
    try:
        # Get CalculationGuid from Sheet
        sheet_data = get_sheet_data(sheet)
        calculation_guid = sheet_data.get("CalculationGuid")
        
        if not calculation_guid:
            return None, None
        
        # Get first viewport to find AreaScheme
        view_ids = sheet.GetAllPlacedViews()
        if not view_ids or view_ids.Count == 0:
            return None, None
        
        # Get the view (should be AreaPlan)
        first_view_id = list(view_ids)[0]
        view = doc.GetElement(first_view_id)
        
        if not hasattr(view, 'AreaScheme'):
            return None, None
        
        area_scheme = view.AreaScheme
        if not area_scheme:
            return None, None
        
        # Get the Calculation from AreaScheme
        calculation_data = get_calculation(area_scheme, calculation_guid)
        
        return area_scheme, calculation_data
        
    except Exception as e:
        print("ERROR getting calculation from sheet: {}".format(e))
        return None, None


def resolve_field_value(field_name, element_data, calculation_data, municipality, element_type):
    """Resolve field value with inheritance support.
    
    Resolution order:
    1. Element's explicit value (if not None)
    2. Calculation's AreaPlanDefaults or AreaDefaults (if present and not None)
    3. Field schema default (from municipality_schemas)
    
    Args:
        field_name: Field name to resolve
        element_data: Element's data dictionary
        calculation_data: Calculation data dictionary
        municipality: Municipality name
        element_type: "AreaPlan" or "Area"
        
    Returns:
        Field value (resolved through inheritance chain) or None
    """
    # Step 1: Check element's explicit value
    if element_data and field_name in element_data and element_data[field_name] is not None:
        return element_data[field_name]
    
    # Step 2: Check Calculation defaults for this element type
    if calculation_data:
        # Determine which defaults field to use
        defaults_field = "AreaPlanDefaults" if element_type == "AreaPlan" else "AreaDefaults"
        
        if defaults_field in calculation_data:
            type_defaults = calculation_data[defaults_field]
            if field_name in type_defaults and type_defaults[field_name] is not None:
                return type_defaults[field_name]
    
    # Step 3: Get field schema default
    try:
        fields = municipality_schemas.get_fields_for_element_type(element_type, municipality)
        if field_name in fields:
            field_def = fields[field_name]
            if "default" in field_def:
                return field_def["default"]
    except:
        pass
    
    return None


# ==================== Sheet Methods ====================

def get_sheet_data(sheet):
    """Get all data from Sheet element.
    
    Args:
        sheet: Sheet element
        
    Returns:
        dict: Sheet data
    """
    return schema_manager.get_data(sheet)


def set_sheet_data(sheet, calculation_guid):
    """Set CalculationGuid on Sheet element.
    
    Note: Sheets only store a reference to their parent Calculation.
    The AreaScheme is determined by finding which AreaScheme contains this Calculation.
    This avoids data redundancy and prevents mismatch errors.
    
    This function MERGES the CalculationGuid into existing sheet data to preserve
    other optional fields like DWFx_UnderlayFilename.
    
    Args:
        sheet: Sheet element
        calculation_guid: Calculation GUID string
        
    Returns:
        bool: True if successful
    """
    # Get existing sheet data to preserve optional fields
    existing_data = schema_manager.get_data(sheet) or {}
    
    # Update CalculationGuid
    existing_data["CalculationGuid"] = calculation_guid
    
    # Clean up legacy v1.0 field if present
    existing_data.pop("AreaSchemeId", None)
    
    return schema_manager.set_data(sheet, existing_data)


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


def get_area_scheme_from_sheet(doc, sheet):
    """Get AreaScheme element from sheet by resolving its Calculation reference.
    
    Derives the AreaScheme by finding which AreaScheme contains the Calculation
    that this sheet references. This avoids storing redundant AreaSchemeId on sheets.
    
    Args:
        doc: Revit document
        sheet: Sheet element
        
    Returns:
        AreaScheme element or None
    """
    sheet_data = get_sheet_data(sheet)
    calculation_guid = sheet_data.get("CalculationGuid")
    
    if not calculation_guid:
        # Fallback: try legacy AreaSchemeId for backward compatibility
        area_scheme_id = sheet_data.get("AreaSchemeId")
        if area_scheme_id:
            return get_area_scheme_by_id(doc, area_scheme_id)
        return None
    
    # Search all AreaSchemes to find which one contains this Calculation
    collector = DB.FilteredElementCollector(doc)
    area_schemes = collector.OfClass(DB.AreaScheme).ToElements()
    
    for area_scheme in area_schemes:
        calculations = get_all_calculations(area_scheme)
        if calculation_guid in calculations:
            return area_scheme
    
    return None


def get_municipality_from_sheet(doc, sheet):
    """Get municipality from sheet by finding its AreaScheme.
    
    Resolves the AreaScheme via the sheet's Calculation reference.
    
    Args:
        doc: Revit document
        sheet: Sheet element
        
    Returns:
        str: Municipality name or None
    """
    area_scheme = get_area_scheme_from_sheet(doc, sheet)
    
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


# ==================== Schema Version Methods ====================

def get_schema_version(doc):
    """Get schema version from ProjectInformation.
    
    Args:
        doc: Revit document
        
    Returns:
        str: Schema version (e.g., "1.0", "2.0") or None if not set
    """
    try:
        proj_info = doc.ProjectInformation
        data = schema_manager.get_data(proj_info) or {}
        return data.get("SchemaVersion")
    except:
        return None


def set_schema_version(doc, version):
    """Set schema version on ProjectInformation.
    
    Args:
        doc: Revit document
        version: Version string (e.g., "2.0")
        
    Returns:
        bool: True if successful
    """
    try:
        proj_info = doc.ProjectInformation
        data = schema_manager.get_data(proj_info) or {}
        data["SchemaVersion"] = version
        return schema_manager.set_data(proj_info, data)
    except Exception as e:
        print("ERROR: Failed to set schema version: {}".format(str(e)))
        return False
