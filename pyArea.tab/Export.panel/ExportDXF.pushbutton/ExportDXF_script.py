#! python3
# -*- coding: utf-8 -*-
"""
Export Area Plans to DXF with municipality-specific formatting.
Uses JSON-based extensible storage schema.
"""

# ============================================================================
# SECTION 1: IMPORTS & SETUP
# ============================================================================

# Standard library imports
import sys
import os
import json
import math
import re

# Add local lib directory for ezdxf package
script_dir = os.path.dirname(__file__)
lib_dir = os.path.join(script_dir, 'lib')
if lib_dir not in sys.path:
    sys.path.insert(0, lib_dir)

# Add schemas directory for municipality_schemas
schemas_dir = os.path.join(script_dir, '..', '..', 'lib', 'schemas')
schemas_dir = os.path.abspath(schemas_dir)
if schemas_dir not in sys.path:
    sys.path.insert(0, schemas_dir)

# Add lib directory for export_utils
utils_lib_dir = os.path.join(script_dir, '..', '..', 'lib')
utils_lib_dir = os.path.abspath(utils_lib_dir)
if utils_lib_dir not in sys.path:
    sys.path.insert(0, utils_lib_dir)

# pyRevit imports (CPython compatible)
from pyrevit import revit, DB, UI, script

# External package (bundled in lib folder)
import ezdxf

# .NET interop
import clr
import System
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
clr.AddReference('System.Windows.Forms')
from Autodesk.Revit.DB.ExtensibleStorage import Schema as ESSchema
from System.Windows.Forms import MessageBox, MessageBoxButtons, MessageBoxIcon

# Import export utilities
import export_utils

# Current Revit document
doc = revit.doc


# ============================================================================
# SECTION 2: CONSTANTS & CONFIGURATION
# ============================================================================

# Import schema identification from schema_guids.py
from schema_guids import SCHEMA_GUID, SCHEMA_NAME, FIELD_NAME

# Coordinate conversion constants
FEET_TO_CM = 30.48          # Revit internal units (feet) to centimeters
FEET_TO_METERS = 0.3048     # Revit internal units (feet) to meters
DEFAULT_VIEW_SCALE = 100.0  # Default scale (1:100) if not found

# Import municipality-specific configuration
from municipality_schemas import (
    MUNICIPALITIES,
    DXF_CONFIG,
    SHEET_FIELDS,
    AREAPLAN_FIELDS,
    AREA_FIELDS
)


# ============================================================================
# SECTION 3: DATA EXTRACTION (JSON + Revit API)
# ============================================================================

def get_element_id_value(element_id):
    """Get integer value from ElementId - compatible with Revit 2024, 2025 and 2026+.
    
    Args:
        element_id: DB.ElementId
        
    Returns:
        int: Integer value of the ElementId
    """
    try:
        # Revit 2024-2025
        return element_id.IntegerValue
    except AttributeError:
        # Revit 2026+ - IntegerValue removed, use Value instead
        return int(element_id.Value)


def get_json_data(element):
    """Read JSON data from extensible storage (CPython compatible).
    
    Args:
        element: Revit element with extensible storage
        
    Returns:
        dict: Parsed JSON data, or empty dict if no data found
    """
    try:
        # Get schema by GUID
        schema_guid = System.Guid(SCHEMA_GUID)
        schema = ESSchema.Lookup(schema_guid)
        
        if not schema:
            return {}
        
        # Get entity from element
        entity = element.GetEntity(schema)
        if not entity.IsValid():
            return {}
        
        # Read JSON string from Data field
        json_string = entity.Get[str](FIELD_NAME)
        
        if not json_string:
            return {}
        
        # Parse JSON
        return json.loads(json_string)
        
    except Exception as e:
        print("Warning: Error reading JSON from element {}: {}".format(element.Id, e))
        return {}


def get_area_scheme_by_id(element_id):
    """Get AreaScheme element by ID.
    
    Args:
        element_id: ElementId as string or int
        
    Returns:
        DB.AreaScheme: AreaScheme element, or None if not found
    """
    try:
        # Convert to ElementId
        if isinstance(element_id, str):
            element_id = int(element_id)
        elem_id = DB.ElementId(element_id)
        
        # Get element
        element = doc.GetElement(elem_id)
        
        # Verify it's an AreaScheme
        if isinstance(element, DB.AreaScheme):
            return element
        
        return None
        
    except Exception as e:
        print("Warning: Error getting AreaScheme by ID {}: {}".format(element_id, e))
        return None


def load_preferences():
    """
    Load export preferences from ProjectInformation.
    Returns default preferences if not found.
    
    Returns:
        dict: Preferences dictionary
    """
    try:
        proj_info = doc.ProjectInformation
        data = get_json_data(proj_info)
        
        if data and "Preferences" in data:
            return data["Preferences"]
        
        # Return defaults if not found
        return export_utils.get_default_preferences()
    except Exception as e:
        print("Warning: Failed to load preferences, using defaults: {}".format(str(e)))
        return export_utils.get_default_preferences()


def get_municipality_from_areascheme(area_scheme):
    """Extract municipality from AreaScheme element.
    
    Args:
        area_scheme: DB.AreaScheme element
        
    Returns:
        str: Municipality name ("Common", "Jerusalem", "Tel-Aviv"), defaults to "Common"
    """
    if not area_scheme:
        return "Common"
    
    data = get_json_data(area_scheme)
    municipality = data.get("Municipality", "Common")
    
    # Validate municipality
    if municipality not in MUNICIPALITIES:
        print("Warning: Invalid municipality '{}', using 'Common'".format(municipality))
        return "Common"
    
    return municipality


def get_sheet_data_for_dxf(sheet_elem):
    """Extract sheet data for DXF export (now via Calculation).
    
    Args:
        sheet_elem: DB.ViewSheet element
        
    Returns:
        dict: Data including calculation_data, area_scheme, municipality, or None if error
    """
    try:
        # Get CalculationGuid from sheet
        sheet_data = get_json_data(sheet_elem)
        calculation_guid = sheet_data.get("CalculationGuid")
        
        # Fallback for legacy v1.0 data (AreaSchemeId)
        area_scheme_id = sheet_data.get("AreaSchemeId")
        
        if not calculation_guid and not area_scheme_id:
            print("Warning: Sheet {} has no CalculationGuid or AreaSchemeId".format(sheet_elem.Id))
            return None
        
        # Get AreaScheme from first viewport
        view_ids = sheet_elem.GetAllPlacedViews()
        if not view_ids or view_ids.Count == 0:
            print("Warning: Sheet {} has no viewports".format(sheet_elem.Id))
            return None
        
        first_view_id = list(view_ids)[0]
        view = doc.GetElement(first_view_id)
        
        if not hasattr(view, 'AreaScheme'):
            print("Warning: First view on sheet {} is not an AreaPlan".format(sheet_elem.Id))
            return None
        
        area_scheme = view.AreaScheme
        if not area_scheme:
            print("Warning: Could not get AreaScheme from view on sheet {}".format(sheet_elem.Id))
            return None
        
        # Get municipality
        municipality = get_municipality_from_areascheme(area_scheme)
        
        # Get Calculation data
        calculation_data = None
        if calculation_guid:
            # v2.0: Get Calculation from AreaScheme
            calculations = get_json_data(area_scheme) or {}
            all_calculations = calculations.get("Calculations", {})
            calculation_data = all_calculations.get(calculation_guid)
            
            if not calculation_data:
                print("Warning: Calculation {} not found on AreaScheme for sheet {}".format(
                    calculation_guid, sheet_elem.Id))
                return None
        else:
            # v1.0: Use sheet data directly as calculation data
            calculation_data = sheet_data
        
        # Get optional DWFx underlay filename from sheet
        dwfx_underlay = sheet_data.get("DWFx_UnderlayFilename")
        
        # Return combined data
        result = {
            "Municipality": municipality,
            "area_scheme": area_scheme,
            "calculation_data": calculation_data,
            "DWFx_UnderlayFilename": dwfx_underlay,
            "_element": sheet_elem
        }
        
        return result
        
    except Exception as e:
        print("Error getting sheet data: {}".format(e))
        import traceback
        traceback.print_exc()
        return None


def get_areaplan_data_for_dxf(areaplan_elem, calculation_data, municipality):
    """Extract areaplan (view) data for DXF export with inheritance.
    
    Args:
        areaplan_elem: DB.ViewPlan element (AreaPlan type)
        calculation_data: Calculation data dictionary (for inheritance)
        municipality: Municipality name
        
    Returns:
        dict: Resolved AreaPlan data with element reference
    """
    try:
        # Get JSON data from view (may have None values for inheritance)
        areaplan_raw = get_json_data(areaplan_elem)
        
        # Get fields for this municipality
        from municipality_schemas import get_fields_for_element_type
        areaplan_fields = get_fields_for_element_type("AreaPlan", municipality)
        
        # Resolve each field with inheritance
        areaplan_data = {}
        for field_name in areaplan_fields.keys():
            # Get element's explicit value
            element_value = areaplan_raw.get(field_name)
            
            # If not None, use it
            if element_value is not None:
                areaplan_data[field_name] = element_value
                continue
            
            # Try Calculation defaults (AreaPlanDefaults)
            if calculation_data and "AreaPlanDefaults" in calculation_data:
                default_value = calculation_data["AreaPlanDefaults"].get(field_name)
                if default_value is not None:
                    areaplan_data[field_name] = default_value
                    continue
            
            # Fall back to schema default
            field_def = areaplan_fields[field_name]
            if "default" in field_def:
                areaplan_data[field_name] = field_def["default"]
            else:
                areaplan_data[field_name] = None
        
        # Add element reference
        areaplan_data["_element"] = areaplan_elem
        
        return areaplan_data
        
    except Exception as e:
        print("Warning: Error getting areaplan data for view {}: {}".format(
            areaplan_elem.Id, e))
        import traceback
        traceback.print_exc()
        return {}


def get_area_data_for_dxf(area_elem, calculation_data, municipality):
    """Extract area data + parameters for DXF export with inheritance.
    
    Args:
        area_elem: DB.Area element
        calculation_data: Calculation data dictionary (for inheritance)
        municipality: Municipality name
        
    Returns:
        dict: Resolved Area data including Usage Type parameters
    """
    try:
        # Get JSON data from area (may have None values for inheritance)
        area_raw = get_json_data(area_elem)
        
        # Get fields for this municipality
        from municipality_schemas import get_fields_for_element_type
        area_fields = get_fields_for_element_type("Area", municipality)
        
        # Resolve each field with inheritance
        area_data = {}
        for field_name in area_fields.keys():
            # Get element's explicit value
            element_value = area_raw.get(field_name)
            
            # If not None, use it
            if element_value is not None:
                area_data[field_name] = element_value
                continue
            
            # Try Calculation defaults (AreaDefaults)
            if calculation_data and "AreaDefaults" in calculation_data:
                default_value = calculation_data["AreaDefaults"].get(field_name)
                if default_value is not None:
                    area_data[field_name] = default_value
                    continue
            
            # Fall back to schema default
            field_def = area_fields[field_name]
            if "default" in field_def:
                area_data[field_name] = field_def["default"]
            else:
                area_data[field_name] = None
        
        # Get shared parameters (NOT from JSON)
        usage_type = ""
        usage_type_prev = ""
        
        param = area_elem.LookupParameter("Usage Type")
        if param and param.HasValue:
            usage_type = param.AsString() or ""
        
        param = area_elem.LookupParameter("Usage Type Prev")
        if param and param.HasValue:
            usage_type_prev = param.AsString() or ""
        
        # Add parameters to data
        area_data["UsageType"] = usage_type
        area_data["UsageTypePrev"] = usage_type_prev
        
        # Add element reference
        area_data["_element"] = area_elem
        
        return area_data
        
    except Exception as e:
        print("Warning: Error getting area data for area {}: {}".format(
            area_elem.Id, e))
        import traceback
        traceback.print_exc()
        return {}


def get_shared_coordinates(point):
    """Convert a point from project coordinates to shared coordinates.
    
    Generic function that transforms any point to shared coordinate system.
    Returns all three coordinates (X, Y, Z) in meters.
    
    Args:
        point: DB.XYZ point in project coordinates
        
    Returns:
        tuple: (x_meters, y_meters, z_meters) or (None, None, None) if error
        
    Examples:
        # Internal origin
        x, y, z = get_shared_coordinates(DB.XYZ(0, 0, 0))
        
        # Project Base Point
        pbp = get_project_base_point()
        x, y, z = get_shared_coordinates(pbp.Position)
    """
    try:
        # Get active project location
        project_location = doc.ActiveProjectLocation
        if not project_location:
            print("Warning: No active project location found")
            return None, None, None
        
        # Get transformation from project to shared coordinates
        project_position = project_location.GetProjectPosition(point)
        
        # Convert from feet to meters
        x_meters = project_position.EastWest * FEET_TO_METERS
        y_meters = project_position.NorthSouth * FEET_TO_METERS
        z_meters = project_position.Elevation * FEET_TO_METERS
        
        return x_meters, y_meters, z_meters
        
    except Exception as e:
        print("Warning: Error converting point to shared coordinates: {}".format(e))
        return None, None, None


def get_project_base_point():
    """Get the Project Base Point element.
    
    Returns:
        DB.BasePoint: Project Base Point element, or None if not found
    """
    try:
        collector = DB.FilteredElementCollector(doc)
        base_points = collector.OfCategory(DB.BuiltInCategory.OST_ProjectBasePoint).ToElements()
        
        if not base_points or len(base_points) == 0:
            print("Warning: Project Base Point not found")
            return None
        
        return base_points[0]
        
    except Exception as e:
        print("Warning: Error getting Project Base Point: {}".format(e))
        return None


def format_meters(value_in_meters):
    """Format a meter value to string with 2 decimal places.
    
    Args:
        value_in_meters: Numeric value in meters
        
    Returns:
        str: Formatted string with 2 decimals, or empty string if None
    """
    if value_in_meters is None:
        return ""
    return "{:.2f}".format(value_in_meters)


# ============================================================================
# SECTION 4: COORDINATE & GEOMETRY UTILITIES
# ============================================================================

def calculate_realworld_scale_factor(view_scale):
    """Compute real-world scale factor: FEET_TO_CM * view_scale.
    
    Args:
        view_scale: View scale number (e.g., 100 for 1:100, 200 for 1:200)
        
    Returns:
        float: Scale factor to convert sheet feet to real-world cm
        
    Example:
        At 1:100 scale: 30.48 * 100 = 3048
        At 1:200 scale: 30.48 * 200 = 6096
    """
    return FEET_TO_CM * view_scale


def convert_point_to_realworld(xyz, scale_factor, offset_x, offset_y):
    """Convert single Revit XYZ point to DXF real-world coordinates.
    
    Args:
        xyz: DB.XYZ point in Revit sheet coordinates (feet)
        scale_factor: REALWORLD_SCALE_FACTOR (from calculate_realworld_scale_factor)
        offset_x: Horizontal sheet offset for multi-sheet layout (feet)
        offset_y: Vertical sheet offset (usually 0) (feet)
        
    Returns:
        tuple: (x, y) in DXF real-world cm coordinates
        
    Note:
        Offset is applied BEFORE scaling to move origin correctly.
        Formula: (point - offset) * scale
    """
    return ((xyz.X - offset_x) * scale_factor, (xyz.Y - offset_y) * scale_factor)


def transform_point_to_sheet(view_point, viewport):
    """Transform point from view coordinates to sheet coordinates.
    
    Uses Revit's transformation matrices to properly convert from view space
    to sheet space, accounting for viewport scale and position.
    
    Note: Only called for validated AreaPlan views with crop regions,
    so transformation matrices are always available.
    
    Args:
        view_point: DB.XYZ point in view coordinates
        viewport: DB.Viewport element
        
    Returns:
        DB.XYZ: Point in sheet coordinates
    """
    # Get view from viewport
    view = doc.GetElement(viewport.ViewId)
    
    # Get transformation chain: view → projection → sheet
    transform_w_boundary = view.GetModelToProjectionTransforms()[0]
    model_to_proj = transform_w_boundary.GetModelToProjectionTransform()
    proj_to_sheet = viewport.GetProjectionToSheetTransform()
    
    # Apply transformations
    proj_point = model_to_proj.OfPoint(view_point)
    sheet_point = proj_to_sheet.OfPoint(proj_point)
    
    return sheet_point


def calculate_arc_bulge(start_pt, end_pt, center_pt, mid_pt):
    """Calculate DXF bulge: tan(angle/4) with mid-point determining arc direction.
    
    The bulge is the tangent of 1/4 the included angle of the arc.
    This version uses the arc center and tests both directions to ensure correct orientation.
    
    Args:
        start_pt: Start point (DB.XYZ)
        end_pt: End point (DB.XYZ)
        center_pt: Arc center point (DB.XYZ)
        mid_pt: Mid point on arc (DB.XYZ)
        
    Returns:
        float: Bulge value for DXF polyline, or 0 if calculation fails
    """
    try:
        # Angles from center to start/end
        start_angle = math.atan2(start_pt.Y - center_pt.Y, start_pt.X - center_pt.X)
        end_angle = math.atan2(end_pt.Y - center_pt.Y, end_pt.X - center_pt.X)
        angle_diff = (end_angle - start_angle) % (2 * math.pi)
        
        # Radius
        radius = math.hypot(start_pt.X - center_pt.X, start_pt.Y - center_pt.Y)
        
        # Test both directions: which computed mid-point is closer to actual mid-point?
        test_ccw = start_angle + angle_diff / 2.0
        dist_ccw = math.hypot(
            mid_pt.X - (center_pt.X + radius * math.cos(test_ccw)),
            mid_pt.Y - (center_pt.Y + radius * math.sin(test_ccw))
        )
        
        test_cw = start_angle - (2 * math.pi - angle_diff) / 2.0
        dist_cw = math.hypot(
            mid_pt.X - (center_pt.X + radius * math.cos(test_cw)),
            mid_pt.Y - (center_pt.Y + radius * math.sin(test_cw))
        )
        
        # Use the direction that matches the actual arc
        included_angle = angle_diff if dist_ccw < dist_cw else -(2 * math.pi - angle_diff)
        return math.tan(included_angle / 4.0)
        
    except Exception as e:
        print("Warning: Error calculating arc bulge: {}".format(e))
        return 0.0


# ============================================================================
# SECTION 5: STRING FORMATTING (Municipality-specific)
# ============================================================================

def is_blank(value):
    """Check if a value should be considered blank/missing.
    
    Args:
        value: Value to check
        
    Returns:
        bool: True if value is None or empty string, False otherwise
    
    Note: 0 and False are NOT considered blank (important for int/bool fields)
    """
    return value is None or (isinstance(value, str) and value.strip() == "")


def with_defaults(data, schema_fields_for_muni):
    """Merge schema defaults into data dictionary for missing fields.
    
    Args:
        data: Data dictionary from JSON
        schema_fields_for_muni: Field schema dict for specific municipality
        
    Returns:
        dict: New dictionary with defaults applied
    """
    result = dict(data)
    for field_name, field_spec in schema_fields_for_muni.items():
        if is_blank(result.get(field_name)) and "default" in field_spec:
            result[field_name] = field_spec["default"]
    return result


def resolve_placeholder(placeholder_value, element):
    """Resolve a placeholder string to its actual value.
    
    Simple direct resolution - just pass the element and get the value.
    
    Args:
        placeholder_value: String that may be a placeholder (e.g., "<View Name>")
        element: Revit element to extract value from (View, Sheet, Area, etc.)
        
    Returns:
        str: Resolved value, or original value if not a placeholder
    """
    if not placeholder_value or not isinstance(placeholder_value, str):
        return placeholder_value or ""
    
    # Not a placeholder - return as-is
    if not (placeholder_value.startswith("<") and placeholder_value.endswith(">")):
        return placeholder_value
    
    # Resolve based on placeholder type
    try:
        if placeholder_value == "<View Name>":
            return element.Name if hasattr(element, 'Name') else ""
        
        elif placeholder_value == "<Level Name>":
            # Get associated level name
            if hasattr(element, 'GenLevel'):
                level = element.GenLevel
                if level:
                    return level.Name
            return ""
        
        elif placeholder_value == "<Title on Sheet>":
            param = element.LookupParameter("Title on Sheet")
            if param and param.HasValue:
                title_value = param.AsString()
                # Check if the value is not empty/blank
                if title_value and title_value.strip():
                    return title_value
            # Fallback to level name if no title on sheet or if empty
            if hasattr(element, 'GenLevel'):
                level = element.GenLevel
                if level:
                    return level.Name
            return ""
        
        elif placeholder_value == "<by Project Base Point>":
            # Get level elevation relative to Project Base Point and convert to meters
            if hasattr(element, 'GenLevel'):
                level = element.GenLevel
                if level:
                    elevation_feet = level.Elevation
                    elevation_meters = elevation_feet * FEET_TO_METERS
                    return format_meters(elevation_meters)
            return ""
        
        elif placeholder_value == "<by Shared Coordinates>":
            # Get level elevation in shared coordinate system
            if hasattr(element, 'GenLevel'):
                level = element.GenLevel
                if level:
                    # Create a point at the level's elevation (in project coordinates)
                    level_point = DB.XYZ(0, 0, level.Elevation)
                    # Transform to shared coordinates and get Z component
                    _, _, z_meters = get_shared_coordinates(level_point)
                    return format_meters(z_meters)
            return ""
        
        elif placeholder_value == "<by Level Above>":
            # TODO: Implement height from level above
            return ""
        
        # Project-level placeholders (from ProjectInformation)
        elif placeholder_value == "<Project Name>":
            proj_info = doc.ProjectInformation
            if proj_info:
                param = proj_info.LookupParameter("Project Name")
                if param and param.HasValue:
                    return param.AsString()
            return ""
        
        elif placeholder_value == "<Project Number>":
            proj_info = doc.ProjectInformation
            if proj_info:
                param = proj_info.LookupParameter("Project Number")
                if param and param.HasValue:
                    return param.AsString()
            return ""
        
        # Coordinate placeholders - Shared Coordinates
        elif placeholder_value == "<E/W@InternalOrigin>":
            x, _, _ = get_shared_coordinates(DB.XYZ(0, 0, 0))
            return format_meters(x)
        
        elif placeholder_value == "<N/S@InternalOrigin>":
            _, y, _ = get_shared_coordinates(DB.XYZ(0, 0, 0))
            return format_meters(y)
        
        elif placeholder_value == "<E/W@ProjectBasePoint>":
            pbp = get_project_base_point()
            if pbp and hasattr(pbp, 'Position'):
                x, _, _ = get_shared_coordinates(pbp.Position)
                return format_meters(x)
            return ""
        
        elif placeholder_value == "<N/S@ProjectBasePoint>":
            pbp = get_project_base_point()
            if pbp and hasattr(pbp, 'Position'):
                _, y, _ = get_shared_coordinates(pbp.Position)
                return format_meters(y)
            return ""
        
        elif placeholder_value == "<SharedElevation@ProjectBasePoint>":
            pbp = get_project_base_point()
            if pbp and hasattr(pbp, 'Position'):
                _, _, z = get_shared_coordinates(pbp.Position)
                return format_meters(z)
            return ""
        
        # Area-specific placeholders
        elif placeholder_value == "<AreaNumber>":
            # Get the area number from the Area element
            if isinstance(element, DB.Area) and hasattr(element, 'Number'):
                area_number = element.Number
                if area_number:
                    return str(area_number)
            return ""
        
    except Exception as e:
        print("  Warning: Error resolving placeholder '{}': {}".format(placeholder_value, e))
    
    # Unresolved placeholder
    return ""


def get_representedViews_data(view_elem_id, municipality, calculation_data):
    """Get floor data from a view in the RepresentedViews list.
    
    Extracts floor name and elevation from AreaPlan views stored in the RepresentedViews JSON field,
    using the SAME inheritance pipeline as regular AreaPlans:
    explicit value → Calculation AreaPlanDefaults → schema default.
    
    This ensures represented views respect Calculation defaults (e.g. "<Title on Sheet>")
    exactly like the main AreaPlan views shown in CalculationSetup.
    
    Args:
        view_elem_id: ElementId (as int or string) of the AreaPlan view from RepresentedViews list
        municipality: Municipality name to determine which fields to extract
        calculation_data: Calculation data dictionary (for inheritance)
        
    Returns:
        tuple: (floor_name, elevation_str) or (None, None) if view not found
    """
    try:
        # Convert to ElementId
        if isinstance(view_elem_id, str):
            elem_id = DB.ElementId(int(view_elem_id))
        elif isinstance(view_elem_id, int):
            elem_id = DB.ElementId(view_elem_id)
        else:
            elem_id = view_elem_id
        
        # Get the view element
        view = doc.GetElement(elem_id)
        
        # Verify it's an AreaPlan view
        if not (isinstance(view, DB.ViewPlan) and view.ViewType == DB.ViewType.AreaPlan):
            return None, None
        
        # Use the same inheritance logic as main AreaPlans to get field values
        # (explicit value → Calculation defaults → schema default)
        represented_data = get_areaplan_data_for_dxf(view, calculation_data, municipality)
        if not represented_data:
            return None, None

        # Extract floor name and elevation based on municipality
        if municipality == "Jerusalem":
            floor_name_field = represented_data.get("FLOOR_NAME", "")
            elevation_field = represented_data.get("FLOOR_ELEVATION", "")
        elif municipality == "Tel-Aviv":
            floor_name_field = represented_data.get("FLOOR", "")
            elevation_field = None  # Tel-Aviv doesn't use elevation in the template
        else:  # Common
            floor_name_field = represented_data.get("FLOOR", "")
            elevation_field = represented_data.get("LEVEL_ELEVATION", "")

        # Resolve placeholders using the represented view element
        floor_name = resolve_placeholder(floor_name_field, view)
        elevation = resolve_placeholder(elevation_field, view) if elevation_field else None
        
        return floor_name, elevation
        
    except Exception as e:
        print("  Warning: Error getting floor info from view {}: {}".format(view_elem_id, e))
        return None, None


def format_sheet_string(sheet_data, municipality, page_number):
    """Format sheet attributes using municipality-specific template.
    
    Args:
        sheet_data: Dictionary with sheet data
        municipality: Municipality name ("Common", "Jerusalem", "Tel-Aviv")
        page_number: Page number for multi-sheet layout (rightmost = 1)
        
    Returns:
        str: Formatted attribute string
    """
    try:
        # Get template for this municipality
        template = DXF_CONFIG[municipality]["string_templates"]["sheet"]
        
        # Get schema fields for defaults
        schema_fields = SHEET_FIELDS.get(municipality, {})
        
        # Apply defaults to sheet_data
        data = with_defaults(sheet_data, schema_fields)
        
        # Resolve placeholders for sheet fields
        # Note: None is passed as element since these are document-level values
        project_value = data.get("PROJECT", "")
        if project_value:
            project_value = resolve_placeholder(project_value, None)
        
        elevation_value = data.get("ELEVATION", "")
        if elevation_value:
            elevation_value = resolve_placeholder(elevation_value, None)
        
        x_value = data.get("X", "")
        if x_value:
            x_value = resolve_placeholder(x_value, None)
        
        y_value = data.get("Y", "")
        if y_value:
            y_value = resolve_placeholder(y_value, None)
        
        # Prepare data with fallbacks
        format_data = {
            "page_number": str(page_number),
            "project": project_value,
            "elevation": elevation_value,
            "building_height": data.get("BUILDING_HEIGHT", ""),
            "x": x_value,
            "y": y_value,
            "lot_area": data.get("LOT_AREA", ""),
            "scale": data.get("scale", "100")
        }
        
        # Format string using template
        return template.format(**format_data)
        
    except Exception as e:
        print("Warning: Error formatting sheet string: {}".format(e))
        return "PAGE_NO={}".format(page_number)


def format_areaplan_string(areaplan_data, municipality, areaplan_elem, calculation_data):
    """Format areaplan attributes using municipality-specific template.
    
    Args:
        areaplan_data: Dictionary with areaplan data
        municipality: Municipality name
        areaplan_elem: DB.ViewPlan element (for placeholder resolution)
        calculation_data: Calculation data dictionary (for inheritance)
        
    Returns:
        str: Formatted attribute string
    """
    try:
        # Get template for this municipality
        template = DXF_CONFIG[municipality]["string_templates"]["areaplan"]
        
        # Apply schema defaults to data
        schema_fields = AREAPLAN_FIELDS.get(municipality, {})
        data = with_defaults(areaplan_data, schema_fields)
        
        # Prepare data based on municipality (resolve placeholders on string fields)
        if municipality == "Jerusalem":
            floor_name = resolve_placeholder(data.get("FLOOR_NAME", ""), areaplan_elem)
            floor_elevation = resolve_placeholder(data.get("FLOOR_ELEVATION", ""), areaplan_elem)
            
            # Check for RepresentedViews and combine floor names/elevations
            represented_views = data.get("RepresentedViews", [])
            if represented_views and isinstance(represented_views, list) and len(represented_views) > 0:
                # Start with current view's values
                floor_names = [floor_name] if floor_name else []
                floor_elevations = [floor_elevation] if floor_elevation else []
                
                # Add represented views' floor names and elevations (using Calculation defaults)
                for view_id in represented_views:
                    rep_floor_name, rep_elevation = get_representedViews_data(view_id, municipality, calculation_data)
                    if rep_floor_name and rep_elevation:
                        floor_names.append(rep_floor_name)
                        floor_elevations.append(rep_elevation)
                
                # Combine with commas
                floor_name = ",".join(floor_names)
                floor_elevation = ",".join(floor_elevations)
            
            format_data = {
                "building_name": resolve_placeholder(data.get("BUILDING_NAME", "1"), areaplan_elem),
                "floor_name": floor_name,
                "floor_elevation": floor_elevation,
                "floor_underground": resolve_placeholder(data.get("FLOOR_UNDERGROUND", "no"), areaplan_elem)
            }
        elif municipality == "Tel-Aviv":
            # Get base floor name
            floor = resolve_placeholder(data.get("FLOOR", ""), areaplan_elem)
            
            # Check for RepresentedViews and combine floor names (using Calculation defaults)
            represented_views = data.get("RepresentedViews", [])
            if represented_views and isinstance(represented_views, list) and len(represented_views) > 0:
                # Start with current view's floor name
                floor_names = [floor] if floor else []
                
                # Add represented views' floor names
                for view_id in represented_views:
                    rep_floor_name, _ = get_representedViews_data(view_id, municipality, calculation_data)
                    if rep_floor_name:
                        floor_names.append(rep_floor_name)
                
                # Combine with commas
                floor = ",".join(floor_names)
            
            format_data = {
                "building": resolve_placeholder(data.get("BUILDING", "1"), areaplan_elem),
                "floor": floor,
                "height": resolve_placeholder(data.get("HEIGHT", ""), areaplan_elem),
                "x": resolve_placeholder(data.get("X", ""), areaplan_elem),
                "y": resolve_placeholder(data.get("Y", ""), areaplan_elem),
                "absolute_height": resolve_placeholder(data.get("Absolute_height", ""), areaplan_elem)
            }
        else:  # Common
            # Get base floor name and elevation
            floor = resolve_placeholder(data.get("FLOOR", ""), areaplan_elem)
            level_elevation = resolve_placeholder(data.get("LEVEL_ELEVATION", ""), areaplan_elem)
            
            # Check for RepresentedViews and combine floor names/elevations
            represented_views = data.get("RepresentedViews", [])
            if represented_views and isinstance(represented_views, list) and len(represented_views) > 0:
                # Start with current view's values
                floor_names = [floor] if floor else []
                level_elevations = [level_elevation] if level_elevation else []
                
                # Add represented views' floor names and elevations
                for view_id in represented_views:
                    rep_floor_name, rep_elevation = get_representedViews_data(view_id, municipality, calculation_data)
                    if rep_floor_name and rep_elevation:
                        floor_names.append(rep_floor_name)
                        level_elevations.append(rep_elevation)
                
                # Combine with commas
                floor = ",".join(floor_names)
                level_elevation = ",".join(level_elevations)
            
            format_data = {
                "building_no": "1",
                "floor": floor,
                "level_elevation": level_elevation,
                "is_underground": str(data.get("IS_UNDERGROUND", 0))
            }
        
        # Format string using template
        return template.format(**format_data)
        
    except Exception as e:
        print("Warning: Error formatting areaplan string: {}".format(e))
        return "FLOOR="


def format_usage_type(value):
    """Format usage type value, converting "0" to empty string.
    
    Args:
        value: Usage type value (string)
        
    Returns:
        str: Empty string if value is "0", otherwise the original value
    """
    return "" if value == "0" else (value or "")


def format_area_string(area_data, municipality, area_elem):
    """Format area attributes using municipality-specific template.
    
    Args:
        area_data: Dictionary with area data (includes UsageType, UsageTypePrev)
        municipality: Municipality name
        area_elem: DB.Area element (for placeholder resolution)
        
    Returns:
        str: Formatted attribute string
    """
    try:
        # Get template for this municipality
        template = DXF_CONFIG[municipality]["string_templates"]["area"]
        
        # Apply schema defaults to data
        schema_fields = AREA_FIELDS.get(municipality, {})
        data = with_defaults(area_data, schema_fields)
        
        # Prepare data based on municipality (resolve placeholders on string fields)
        if municipality == "Jerusalem":
            format_data = {
                "code": format_usage_type(data.get("UsageType", "")),  # Parameter, not JSON field
                "demolition_source_code": format_usage_type(data.get("UsageTypePrev", "")),  # Parameter, not JSON field
                "area": resolve_placeholder(data.get("AREA", ""), area_elem),
                "height1": resolve_placeholder(data.get("HEIGHT", ""), area_elem),
                "appartment_num": resolve_placeholder(data.get("APPARTMENT_NUM", ""), area_elem),
                "height2": resolve_placeholder(data.get("HEIGHT2", ""), area_elem)
            }
        elif municipality == "Tel-Aviv":
            format_data = {
                "code": format_usage_type(data.get("UsageType", "")),  # Parameter, not JSON field
                "code_before": format_usage_type(data.get("UsageTypePrev", "")),  # Parameter, not JSON field
                "id": resolve_placeholder(data.get("ID", ""), area_elem),  # JSON field with placeholder support
                "apartment": resolve_placeholder(data.get("APARTMENT", ""), area_elem),
                "heter": resolve_placeholder(data.get("HETER", "1"), area_elem),
                "height": resolve_placeholder(data.get("HEIGHT", ""), area_elem)
            }
        else:  # Common
            format_data = {
                "usage_type": format_usage_type(data.get("UsageType", "")),  # Parameter, not JSON field
                "usage_type_old": format_usage_type(data.get("UsageTypePrev", "")),  # Parameter, not JSON field
                "area": resolve_placeholder(data.get("AREA", ""), area_elem),
                "asset": resolve_placeholder(data.get("ASSET", ""), area_elem)
            }
        
        # Format string using template
        return template.format(**format_data)
        
    except Exception as e:
        print("Warning: Error formatting area string: {}".format(e))
        return "AREA="


# ============================================================================
# SECTION 6: DXF LAYER & ENTITY CREATION
# ============================================================================

def create_dxf_layers(dxf_doc, municipality):
    """Create DXF layers from municipality configuration.
    
    Args:
        dxf_doc: ezdxf DXF document
        municipality: Municipality name
        
    Returns:
        dict: Layer name mapping for quick access
    """
    try:
        # Get layer configuration
        layers = DXF_CONFIG[municipality]["layers"]
        layer_colors = DXF_CONFIG[municipality]["layer_colors"]
        
        # Create each layer with its color
        for layer_key, layer_name in layers.items():
            if layer_name not in dxf_doc.layers:
                layer = dxf_doc.layers.new(name=layer_name)
                # Set color if defined
                if layer_name in layer_colors:
                    layer.color = layer_colors[layer_name]
        
        return layers
        
    except Exception as e:
        print("Error creating DXF layers: {}".format(e))
        return {}


def add_rectangle(msp, min_point, max_point, layer_name):
    """Add rectangle to DXF using polyline.
    
    Args:
        msp: DXF modelspace
        min_point: (x, y) tuple for bottom-left corner
        max_point: (x, y) tuple for top-right corner
        layer_name: DXF layer name
    """
    try:
        x_min, y_min = min_point
        x_max, y_max = max_point
        
        # Create closed polyline for rectangle
        points = [
            (x_min, y_min),
            (x_max, y_min),
            (x_max, y_max),
            (x_min, y_max)
        ]
        
        polyline = msp.add_lwpolyline(points, dxfattribs={'layer': layer_name})
        polyline.closed = True
        
    except Exception as e:
        print("Warning: Error adding rectangle: {}".format(e))


def add_text(msp, text, position, layer_name, height=10.0):
    """Add text entity to DXF.
    
    Args:
        msp: DXF modelspace
        text: Text string to add
        position: (x, y) tuple for text insertion point
        layer_name: DXF layer name
        height: Text height in DXF units (default 10.0 cm)
    """
    try:
        x, y = position
        msp.add_text(
            text,
            dxfattribs={
                'layer': layer_name,
                'height': height,
                'insert': (x, y, 0),
                'style': 'Standard'
            }
        )
        
    except Exception as e:
        print("Warning: Error adding text: {}".format(e))


def add_polyline_with_arcs(msp, points, layer_name, bulges=None):
    """Add polyline with optional arc segments (bulges) to DXF.
    
    Args:
        msp: DXF modelspace
        points: List of (x, y) tuples
        layer_name: DXF layer name
        bulges: Optional list of bulge values (same length as points)
                Bulge = 0 for straight line, non-zero for arc
    """
    try:
        if not points or len(points) < 2:
            return
        
        # If bulges provided, create polyline with bulges in xyseb format
        if bulges and len(bulges) > 0:
            # Create points_with_bulge list: (x, y, start_width, end_width, bulge)
            points_with_bulge = []
            for i, (x, y) in enumerate(points):
                bulge_val = bulges[i] if i < len(bulges) else 0.0
                points_with_bulge.append((x, y, 0, 0, bulge_val))
            
            # Create polyline with bulge values using xyseb format
            polyline = msp.add_lwpolyline(points_with_bulge, format='xyseb', dxfattribs={'layer': layer_name})
            polyline.closed = True
        else:
            # Simple polyline without arcs
            polyline = msp.add_lwpolyline(points, dxfattribs={'layer': layer_name})
            polyline.closed = True
        
    except Exception as e:
        print("Warning: Error adding polyline: {}".format(e))


def add_dwfx_underlay(dxf_doc, msp, dwfx_filename, insert_point, scale):
    """Add DWFx underlay reference to DXF.
    
    Args:
        dxf_doc: ezdxf DXF document
        msp: DXF modelspace
        dwfx_filename: Filename of DWFx file (without path, just filename.dwfx)
        insert_point: (x, y) tuple for insertion point in DXF coordinates
        scale: Scale factor (e.g., 100 for 1:100, 200 for 1:200)
        
    Returns:
        bool: True if successful, False otherwise
    """
    try:
        # Get or create underlay definitions collection
        if not hasattr(dxf_doc, 'add_underlay_def'):
            print("  Warning: DWFx underlay not supported in this ezdxf version")
            return False
        
        # Add underlay definition (dwfx type)
        # name='1' represents the first sheet in the DWFx file (required for AutoCAD)
        try:
            underlay_def = dxf_doc.add_underlay_def(
                filename=dwfx_filename,
                fmt='dwf',  # DWF format covers both .dwf and .dwfx files
                name='1'  # First sheet in DWFx file
            )
        except Exception as e:
            print("  Warning: Could not create underlay definition: {}".format(e))
            return False
        
        # Add underlay entity to modelspace
        # DWFx is in mm, DXF is in cm, so divide scale by 10
        dwfx_scale = scale / 10.0
        underlay = msp.add_underlay(
            underlay_def,
            insert=insert_point,
            scale=(dwfx_scale, dwfx_scale, dwfx_scale)
        )
        
        # Assign to layer 0
        underlay.dxf.layer = '0'
        
        print("  Added DWFx underlay: {} (scale: {})".format(dwfx_filename, dwfx_scale))
        return True
        
    except Exception as e:
        print("  Warning: Error adding DWFx underlay: {}".format(e))
        return False

# ============================================================================
# SECTION 7: PROCESSING PIPELINE
# ============================================================================

def process_area(area_elem, viewport, msp, scale_factor, offset_x, offset_y, municipality, layers, calculation_data):
    """Process single Area element - add boundary and text to DXF.
    
    Args:
        area_elem: DB.Area element
        viewport: DB.Viewport element (for coordinate transformation)
        msp: DXF modelspace
        scale_factor: REALWORLD_SCALE_FACTOR
        offset_x: Horizontal offset (feet)
        offset_y: Vertical offset (feet)
        municipality: Municipality name
        layers: Layer name mapping
        calculation_data: Calculation data dict (for inheritance)
    """
    try:
        # Get area data with inheritance
        area_data = get_area_data_for_dxf(area_elem, calculation_data, municipality)
        if not area_data:
            return
        
        # Get boundary segments
        boundary_options = DB.SpatialElementBoundaryOptions()
        boundary_segments = area_elem.GetBoundarySegments(boundary_options)
        
        if not boundary_segments or len(boundary_segments) == 0:
            print("  Warning: Area {} has no boundary".format(area_elem.Id))
            return
        
        # Find the exterior boundary loop using signed area (Shoelace formula)
        # Revit does NOT guarantee boundary_segments[0] is the exterior
        # Exterior loop has the largest absolute signed area
        # Use Tessellate() to handle arcs/curves accurately (important for round buildings)
        exterior_loop = None
        max_abs_area = 0.0
        for loop in boundary_segments:
            # Collect tessellated points for accurate area calculation
            pts = []
            for segment in loop:
                curve = segment.GetCurve()
                tess_pts = list(curve.Tessellate())
                # Add all but last point (last point = next segment's first point)
                pts.extend(tess_pts[:-1])
            if len(pts) < 3:
                continue
            # Shoelace formula for signed area
            signed_area = 0.0
            n = len(pts)
            for i in range(n):
                j = (i + 1) % n
                signed_area += pts[i].X * pts[j].Y
                signed_area -= pts[j].X * pts[i].Y
            signed_area /= 2.0
            abs_area = abs(signed_area)
            if abs_area > max_abs_area:
                max_abs_area = abs_area
                exterior_loop = loop
        
        if exterior_loop is None:
            print("  Warning: Area {} could not determine exterior boundary".format(area_elem.Id))
            return
        
        # Collect all boundary points
        boundary_points = []
        bulges = []
        tol = 1e-9
        def _append_pt(pt_sheet, bulge_val):
            if boundary_points:
                prev = boundary_points[-1]
                if abs(prev.X - pt_sheet.X) < tol and abs(prev.Y - pt_sheet.Y) < tol:
                    return
            boundary_points.append(pt_sheet)
            bulges.append(bulge_val)

        # Cache transforms for this viewport/view
        try:
            view = doc.GetElement(viewport.ViewId)
            transform_w_boundary = view.GetModelToProjectionTransforms()[0]
            model_to_proj = transform_w_boundary.GetModelToProjectionTransform()
            proj_to_sheet = viewport.GetProjectionToSheetTransform()
            def _to_sheet(xyz):
                return proj_to_sheet.OfPoint(model_to_proj.OfPoint(xyz))
        except Exception:
            def _to_sheet(xyz):
                return transform_point_to_sheet(xyz, viewport)

        for segment in exterior_loop:
            curve = segment.GetCurve()
            if isinstance(curve, DB.Arc):
                try:
                    start_pt_view = curve.GetEndPoint(0)
                    end_pt_view = curve.GetEndPoint(1)
                    start_pt_sheet = _to_sheet(start_pt_view)
                    center_sheet = _to_sheet(curve.Center)
                    end_pt_sheet = _to_sheet(end_pt_view)

                    tessellated = list(curve.Tessellate())
                    if len(tessellated) >= 2:
                        mid_idx = len(tessellated) // 2
                        mid_sheet = _to_sheet(tessellated[mid_idx])
                    else:
                        start_vec = DB.XYZ(start_pt_sheet.X - center_sheet.X, start_pt_sheet.Y - center_sheet.Y, 0)
                        end_vec = DB.XYZ(end_pt_sheet.X - center_sheet.X, end_pt_sheet.Y - center_sheet.Y, 0)
                        mid_vec_x = (start_vec.X + end_vec.X) / 2.0
                        mid_vec_y = (start_vec.Y + end_vec.Y) / 2.0
                        radius = math.hypot(start_vec.X, start_vec.Y)
                        vec_len = math.hypot(mid_vec_x, mid_vec_y)
                        if vec_len > 0:
                            mid_sheet = DB.XYZ(
                                center_sheet.X + (mid_vec_x / vec_len) * radius,
                                center_sheet.Y + (mid_vec_y / vec_len) * radius,
                                0
                            )
                        else:
                            mid_sheet = start_pt_sheet

                    bulge = calculate_arc_bulge(start_pt_sheet, end_pt_sheet, center_sheet, mid_sheet)
                    _append_pt(start_pt_sheet, bulge)
                except Exception as ex:
                    print("  Arc bulge error: {}".format(str(ex)))
                    try:
                        start_pt_sheet = _to_sheet(curve.GetEndPoint(0))
                        _append_pt(start_pt_sheet, 0.0)
                    except Exception:
                        pass

            else:
                try:
                    tessellated_points = list(curve.Tessellate())
                    for pt_view in tessellated_points[:-1]:
                        pt_sheet = _to_sheet(pt_view)
                        _append_pt(pt_sheet, 0.0)
                except Exception as ex:
                    print("  Warning: Failed to process line segment: {}".format(str(ex)))
        
        # Close the boundary
        if len(boundary_points) > 0:
            _append_pt(boundary_points[0], 0.0)
        
        # Transform SHEET coordinates to DXF coordinates
        transformed_points = [
            convert_point_to_realworld(pt, scale_factor, offset_x, offset_y)
            for pt in boundary_points
        ]
        
        # Add boundary polyline
        add_polyline_with_arcs(
            msp, 
            transformed_points, 
            layers['area_boundary'],
            bulges if any(b != 0 for b in bulges) else None
        )
        
        # Add area text at Location.Point
        location = area_elem.Location
        if location and isinstance(location, DB.LocationPoint):
            loc_pt_view = location.Point
            
            # Transform VIEW to SHEET coordinates
            loc_pt_sheet = transform_point_to_sheet(loc_pt_view, viewport)
            
            # Transform SHEET to DXF coordinates
            text_pos = convert_point_to_realworld(
                loc_pt_sheet, scale_factor, offset_x, offset_y
            )
            
            # Format area string
            area_string = format_area_string(area_data, municipality, area_elem)
            
            # Add text
            add_text(msp, area_string, text_pos, layers['area_text'])
        
    except Exception as e:
        print("  Warning: Error processing area {}: {}".format(area_elem.Id, e))


def process_areaplan_viewport(viewport, msp, scale_factor, offset_x, offset_y, municipality, layers, calculation_data):
    """Process AreaPlan viewport - add crop boundary, plan text, and all areas.
    
    Args:
        viewport: DB.Viewport element
        msp: DXF modelspace
        scale_factor: REALWORLD_SCALE_FACTOR
        offset_x: Horizontal offset (cm)
        offset_y: Vertical offset (cm)
        municipality: Municipality name
        layers: Layer name mapping
        calculation_data: Calculation data dict (for inheritance)
    """
    try:
        # Get the view from viewport
        view_id = viewport.ViewId
        view = doc.GetElement(view_id)
        
        if not view or not isinstance(view, DB.ViewPlan):
            return
        
        # Check if it's an area plan
        if view.ViewType != DB.ViewType.AreaPlan:
            return
        
        print("  Processing AreaPlan: {}".format(view.Name))
        
        # Get areaplan data with inheritance
        areaplan_data = get_areaplan_data_for_dxf(view, calculation_data, municipality)
        
        # Get crop boundary
        crop_manager = view.GetCropRegionShapeManager()
        if crop_manager:
            crop_shape = crop_manager.GetCropShape()
            # GetCropShape() returns IList[CurveLoop]
            if crop_shape and crop_shape.Count > 0:
                # Collect crop boundary points from all curve loops (in VIEW coordinates)
                crop_points_view = []
                for curve_loop in crop_shape:
                    for curve in curve_loop:
                        start_pt = curve.GetEndPoint(0)
                        crop_points_view.append(start_pt)
                
                # Transform VIEW coordinates to SHEET coordinates
                crop_points_sheet = []
                for pt_view in crop_points_view:
                    pt_sheet = transform_point_to_sheet(pt_view, viewport)
                    crop_points_sheet.append(pt_sheet)
                
                # Close the boundary
                if len(crop_points_sheet) > 0:
                    crop_points_sheet.append(crop_points_sheet[0])
                    
                    # Transform SHEET coordinates to DXF coordinates
                    transformed_crop = [
                        convert_point_to_realworld(pt, scale_factor, offset_x, offset_y)
                        for pt in crop_points_sheet
                    ]
                    
                    # Add crop boundary rectangle/polyline
                    add_polyline_with_arcs(msp, transformed_crop, layers['areaplan_frame'])
                    
                    # Add areaplan text at top-right corner
                    if len(transformed_crop) > 0:
                        # Find top-right corner in DXF space
                        max_x_dxf = max(x for x, y in transformed_crop)
                        max_y_dxf = max(y for x, y in transformed_crop)
                        # Position text 200 cm below and to the left in DXF space
                        text_pos = (max_x_dxf - 200.0, max_y_dxf - 200.0)

                        # Use Calculation-aware formatting so represented views inherit
                        # FLOOR/LEVEL_ELEVATION defaults from the Calculation, just like in CalculationSetup
                        areaplan_string = format_areaplan_string(areaplan_data, municipality, view, calculation_data)
                        add_text(msp, areaplan_string, text_pos, layers['areaplan_text'])
        
        # Get all areas in this view
        collector = DB.FilteredElementCollector(doc, view_id)
        areas = collector.OfCategory(DB.BuiltInCategory.OST_Areas).WhereElementIsNotElementType().ToElements()
        
        print("    Found {} areas".format(len(areas)))
        
        # Process each area
        for area in areas:
            if isinstance(area, DB.Area):
                process_area(area, viewport, msp, scale_factor, offset_x, offset_y, municipality, layers, calculation_data)
        
    except Exception as e:
        print("  Warning: Error processing viewport: {}".format(e))


def process_sheet(sheet_elem, dxf_doc, msp, horizontal_offset, page_number, view_scale, valid_viewports):
    """Process entire sheet with horizontal offset for multi-sheet layout.
    
    Args:
        sheet_elem: DB.ViewSheet element
        dxf_doc: ezdxf DXF document
        msp: DXF modelspace
        horizontal_offset: Horizontal offset for this sheet (Revit feet)
        page_number: Page number (rightmost = 1)
        view_scale: Validated uniform view scale for entire export
        valid_viewports: List of pre-validated DB.Viewport elements to process
        
    Returns:
        float: Width of this sheet in Revit feet (for next sheet's offset)
    """
    try:
        print("\n" + "-"*60)
        print("Processing Sheet: {} - {}".format(sheet_elem.SheetNumber, sheet_elem.Name))
        
        # Get sheet data (includes calculation_data)
        sheet_data = get_sheet_data_for_dxf(sheet_elem)
        if not sheet_data:
            print("  Warning: No sheet data found")
            return 0.0
        
        # Extract components from sheet_data
        municipality = sheet_data.get("Municipality", "Common")
        calculation_data = sheet_data.get("calculation_data")
        print("  Municipality: {}".format(municipality))
        
        # Create layers based on municipality
        layers = create_dxf_layers(dxf_doc, municipality)
        
        print("  Using validated scale: 1:{}".format(int(view_scale)))
        print("  Processing {} valid viewports".format(len(valid_viewports)))
        
        # Calculate scale factor
        scale_factor = calculate_realworld_scale_factor(view_scale)
        print("  Scale factor: {}".format(scale_factor))
        
        # Get sheet dimensions
        titleblock = DB.FilteredElementCollector(doc, sheet_elem.Id)\
            .OfCategory(DB.BuiltInCategory.OST_TitleBlocks)\
            .FirstElement()
        
        if not titleblock:
            print("  Warning: No titleblock found")
            return 0.0
        
        bbox = titleblock.get_BoundingBox(sheet_elem)
        if not bbox:
            print("  Warning: Could not get titleblock bounding box")
            return 0.0
        
        # Calculate sheet width in Revit feet
        sheet_width = bbox.Max.X - bbox.Min.X
        sheet_height = bbox.Max.Y - bbox.Min.Y
        
        # For display, convert to cm
        sheet_width_cm = sheet_width * scale_factor
        sheet_height_cm = sheet_height * scale_factor
        print("  Sheet size: {:.1f} x {:.1f} cm ({:.3f} x {:.3f} ft)".format(
            sheet_width_cm, sheet_height_cm, sheet_width, sheet_height))
        
        # Calculate offsets to move bottom-left corner to DXF origin
        # For multi-sheet layout, also apply horizontal offset
        offset_x = bbox.Min.X - horizontal_offset
        offset_y = bbox.Min.Y
        
        print("  Offset: X={:.3f} ft, Y={:.3f} ft (horizontal_offset={:.3f} ft)".format(
            offset_x, offset_y, horizontal_offset))
        
        # Add DWFx underlay (background reference)
        print("  Attempting to add DWFx underlay...")
        
        # Use custom DWFx filename if provided, otherwise generate default
        custom_dwfx = sheet_data.get("DWFx_UnderlayFilename")
        if custom_dwfx and custom_dwfx.strip():
            # User provided a custom filename - use basename only (same folder as DXF)
            import os
            dwfx_filename = os.path.basename(custom_dwfx.strip())
            # Ensure .dwfx extension
            if not dwfx_filename.lower().endswith('.dwfx'):
                dwfx_filename += '.dwfx'
            print("  DWFx filename (custom): {}".format(dwfx_filename))
        else:
            # Generate default filename
            dwfx_filename = export_utils.generate_dwfx_filename(doc.Title, sheet_elem.SheetNumber) + ".dwfx"
            print("  DWFx filename (generated): {}".format(dwfx_filename))
        underlay_insert_point = convert_point_to_realworld(bbox.Min, scale_factor, offset_x, offset_y)
        print("  Underlay insert point: {}".format(underlay_insert_point))
        result = add_dwfx_underlay(
            dxf_doc, 
            msp, 
            dwfx_filename, 
            underlay_insert_point,
            scale=view_scale
        )
        print("  DWFx underlay result: {}".format(result))
        
        # Add sheet frame rectangle (titleblock outline)
        if titleblock and bbox:
            min_point = convert_point_to_realworld(bbox.Min, scale_factor, offset_x, offset_y)
            max_point = convert_point_to_realworld(bbox.Max, scale_factor, offset_x, offset_y)
            add_rectangle(msp, min_point, max_point, layers['sheet_frame'])
        
        # Add sheet text at top-right corner (use calculation_data for fields)
        sheet_string = format_sheet_string(calculation_data, municipality, page_number)
        if titleblock and bbox:
            # Get top-right corner in DXF space
            max_point_dxf = convert_point_to_realworld(bbox.Max, scale_factor, offset_x, offset_y)
            # Position text 10 cm below and to the left in DXF space
            text_pos = (max_point_dxf[0] - 10.0, max_point_dxf[1] - 10.0)
            add_text(msp, sheet_string, text_pos, layers['sheet_text'])
        
        # Process pre-validated viewports (pass calculation_data)
        for viewport in valid_viewports:
            process_areaplan_viewport(
                viewport, msp, scale_factor, offset_x, offset_y, municipality, layers, calculation_data
            )
        
        return sheet_width
        
    except Exception as e:
        print("Error processing sheet: {}".format(e))
        return 0.0

# ============================================================================
# SECTION 8: SHEET SELECTION & SORTING
# ============================================================================

def get_valid_areaplans_and_uniform_scale(sheets):
    """Comprehensive validation: AreaScheme uniformity, valid AreaPlan views, and uniform scale.
    
    Performs validation in single pass:
    1. Validates all sheets belong to same AreaScheme (via Calculation hierarchy)
    2. Filters valid AreaPlan views (has municipality, has areas, has scale)
    3. Validates uniform scale across all valid views
    
    Only includes views that:
    - Belong to an AreaScheme with a defined municipality
    - Contain areas
    - Have a valid scale
    
    Args:
        sheets: List of DB.ViewSheet elements
        
    Returns:
        tuple: (uniform_scale, valid_viewports_dict)
            - uniform_scale: float, the validated uniform scale
            - valid_viewports_dict: {sheet.Id: [list of valid DB.Viewport]}
        
    Raises:
        ValueError: If no valid views found or mixed scales detected
    """
    try:
        # ===== PHASE 1: Validate uniform AreaScheme across all sheets =====
        schemes_found = {}  # {area_scheme_id: [sheet_numbers]}
        missing_scheme_sheets = []  # [sheet_numbers]
        
        for sheet in sheets:
            # Get sheet JSON data
            sheet_data = get_json_data(sheet)
            if not sheet_data:
                missing_scheme_sheets.append(sheet.SheetNumber)
                continue
            
            # Get CalculationGuid (v2.0) or AreaSchemeId (v1.0 fallback)
            calculation_guid = sheet_data.get("CalculationGuid")
            area_scheme_id_legacy = sheet_data.get("AreaSchemeId")
            
            if not calculation_guid and not area_scheme_id_legacy:
                missing_scheme_sheets.append(sheet.SheetNumber)
                continue
            
            # Resolve AreaScheme via viewports
            area_scheme = None
            view_ids = sheet.GetAllPlacedViews()
            if view_ids and view_ids.Count > 0:
                first_view_id = list(view_ids)[0]
                view = doc.GetElement(first_view_id)
                if hasattr(view, 'AreaScheme'):
                    area_scheme = view.AreaScheme
            
            if not area_scheme:
                missing_scheme_sheets.append(sheet.SheetNumber)
                continue
            
            # Get AreaScheme ElementId value
            area_scheme_id = str(get_element_id_value(area_scheme.Id))
            
            # Track which sheets belong to which scheme
            if area_scheme_id not in schemes_found:
                schemes_found[area_scheme_id] = []
            schemes_found[area_scheme_id].append(sheet.SheetNumber)
        
        # Error if sheets are missing AreaScheme
        if missing_scheme_sheets:
            error_msg = "ERROR: Sheets without AreaScheme detected!\n\n"
            error_msg += "The following sheets have no CalculationGuid or AreaScheme:\n"
            for sheet_num in missing_scheme_sheets:
                error_msg += "  - Sheet {}\n".format(sheet_num)
            error_msg += "\nAll sheets must belong to an AreaScheme for DXF export."
            raise ValueError(error_msg)
        
        # Error if no schemes found at all
        if len(schemes_found) == 0:
            error_msg = "No AreaScheme found in selected sheets.\n\n"
            error_msg += "All sheets must belong to an AreaScheme."
            raise ValueError(error_msg)
        
        # Error if multiple AreaSchemes detected
        if len(schemes_found) > 1:
            error_msg = "ERROR: Multiple AreaSchemes detected!\n\n"
            error_msg += "All selected sheets must belong to the same AreaScheme for DXF export.\n\n"
            error_msg += "AreaSchemes found:\n"
            for scheme_id, sheet_numbers in schemes_found.items():
                # Get scheme name
                scheme = get_area_scheme_by_id(scheme_id)
                scheme_name = scheme.Name if scheme else "Unknown"
                error_msg += "\n  AreaScheme '{}' (ID: {}):\n".format(scheme_name, scheme_id)
                for sheet_num in sheet_numbers:
                    error_msg += "    - Sheet {}\n".format(sheet_num)
            error_msg += "\nPlease select sheets from the same AreaScheme only."
            raise ValueError(error_msg)
        
        # All sheets have same AreaScheme - log success
        uniform_scheme_id = list(schemes_found.keys())[0]
        scheme = get_area_scheme_by_id(uniform_scheme_id)
        scheme_name = scheme.Name if scheme else "Unknown"
        print("AreaScheme validation passed:")
        print("  - AreaScheme: {} (ID: {})".format(scheme_name, uniform_scheme_id))
        print("  - Sheets: {}".format(len(sheets)))
        
        # ===== PHASE 2: Filter valid AreaPlan views and validate scale =====
        scales_found = {}  # {scale: [(sheet_number, view_name)]}
        valid_viewports = {}  # {sheet.Id: [viewport]}
        
        for sheet in sheets:
            sheet_valid_viewports = []
            viewports = list(DB.FilteredElementCollector(doc, sheet.Id)
                            .OfClass(DB.Viewport)
                            .ToElements())
            
            for viewport in viewports:
                view = doc.GetElement(viewport.ViewId)
                
                # Must be AreaPlan
                if not view or view.ViewType != DB.ViewType.AreaPlan:
                    continue
                
                # Must have AreaScheme with municipality
                try:
                    areascheme = view.AreaScheme
                except Exception:
                    continue
                
                if not areascheme:
                    continue
                
                municipality = get_municipality_from_areascheme(areascheme)
                if not municipality:
                    continue
                
                # Must have areas
                areas = list(DB.FilteredElementCollector(doc, view.Id)
                            .OfCategory(DB.BuiltInCategory.OST_Areas)
                            .WhereElementIsNotElementType()
                            .ToElements())
                if not areas or len(areas) == 0:
                    continue
                
                # Must have valid scale
                if not hasattr(view, 'Scale'):
                    continue
                scale = float(view.Scale)
                
                # This viewport is VALID!
                sheet_valid_viewports.append(viewport)
                
                # Track scale for validation
                if scale not in scales_found:
                    scales_found[scale] = []
                scales_found[scale].append({
                    'sheet': sheet.SheetNumber,
                    'view': view.Name
                })
            
            # Store valid viewports for this sheet
            if len(sheet_valid_viewports) > 0:
                valid_viewports[sheet.Id] = sheet_valid_viewports
        
        # Check if we found any valid views
        if len(scales_found) == 0:
            error_msg = "No valid AreaPlan views found.\n\n"
            error_msg += "Valid views must:\n"
            error_msg += "- Belong to an AreaScheme with defined municipality\n"
            error_msg += "- Contain areas\n"
            error_msg += "- Have a defined scale"
            raise ValueError(error_msg)
        
        # Check for uniform scale
        if len(scales_found) > 1:
            error_msg = "ERROR: Mixed scales detected in valid AreaPlan views!\n\n"
            error_msg += "All valid AreaPlan views must have the same scale for DXF export.\n\n"
            error_msg += "Scales found:\n"
            for scale, locations in sorted(scales_found.items()):
                error_msg += "\n  Scale 1:{}:\n".format(int(scale))
                for loc in locations:
                    error_msg += "    - Sheet {} / {}\n".format(loc['sheet'], loc['view'])
            error_msg += "\nPlease ensure all AreaPlan views use the same scale before exporting."
            raise ValueError(error_msg)
        
        # All valid views have same scale - perfect!
        uniform_scale = list(scales_found.keys())[0]
        total_viewports = sum(len(v) for v in valid_viewports.values())
        print("Validation passed:")
        print("  - Uniform scale: 1:{}".format(int(uniform_scale)))
        print("  - Valid viewports: {}".format(total_viewports))
        
        return uniform_scale, valid_viewports
        
    except ValueError:
        # Re-raise validation errors
        raise
    except Exception as e:
        error_msg = "Error during validation: {}".format(e)
        print(error_msg)
        raise ValueError(error_msg)


def get_selected_sheets():
    """Get sheets from project browser selection or active view.
    
    Returns:
        list: List of DB.ViewSheet elements, or None if no valid selection
    """
    try:
        # Check if active view is a sheet
        active_view = doc.ActiveView
        if isinstance(active_view, DB.ViewSheet):
            print("Using active sheet: {}".format(active_view.SheetNumber))
            return [active_view]
        
        # Try to get selection from project browser
        uidoc = revit.uidoc
        selection = uidoc.Selection
        selected_ids = selection.GetElementIds()
        
        if selected_ids and len(selected_ids) > 0:
            sheets = []
            for elem_id in selected_ids:
                element = doc.GetElement(elem_id)
                if isinstance(element, DB.ViewSheet):
                    sheets.append(element)
            
            if len(sheets) > 0:
                print("Found {} selected sheets".format(len(sheets)))
                return sheets
        
        # No valid selection - ask user to select sheets
        print("No sheets selected. Please select sheets in project browser or open a sheet.")
        return None
        
    except Exception as e:
        print("Error getting selected sheets: {}".format(e))
        return None


def extract_sheet_number_for_sorting(sheet):
    """Extract numeric portion from sheet number for sorting.
    
    Args:
        sheet: DB.ViewSheet element
        
    Returns:
        tuple: (numeric_part, full_sheet_number) for sorting
    """
    try:
        sheet_number = sheet.SheetNumber
        
        # Try to extract leading number
        match = re.match(r'^(\d+)', sheet_number)
        if match:
            return (int(match.group(1)), sheet_number)
        
        # No leading number, sort alphabetically
        return (999999, sheet_number)
        
    except:
        return (999999, "")


def sort_sheets_by_number(sheets, descending=True):
    """Sort sheets by sheet number (rightmost = page 1).
    
    Args:
        sheets: List of DB.ViewSheet elements
        descending: If True, highest number = page 1 (rightmost in layout)
        
    Returns:
        list: Sorted list of sheets
    """
    try:
        # Sort by extracted number
        sorted_sheets = sorted(sheets, 
                              key=extract_sheet_number_for_sorting,
                              reverse=descending)
        
        print("\nSheet order (left to right):")
        for i, sheet in enumerate(sorted_sheets):
            print("  {} - {} (Page {})".format(
                sheet.SheetNumber, sheet.Name, len(sorted_sheets) - i))
        
        return sorted_sheets
        
    except Exception as e:
        print("Error sorting sheets: {}".format(e))
        return sheets


def group_sheets_by_calculation(initial_sheets):
    """Group sheets by their CalculationGuid.
    
    Args:
        initial_sheets: List of DB.ViewSheet elements
        
    Returns:
        dict: {calculation_guid: [sheets], ...}
              None key = sheets without CalculationGuid (legacy v1.0)
    """
    try:
        groups = {}
        
        for sheet in initial_sheets:
            sheet_data = get_json_data(sheet) or {}
            calc_guid = sheet_data.get("CalculationGuid")
            
            # Use None as key for legacy sheets
            key = calc_guid if calc_guid else None
            
            if key not in groups:
                groups[key] = []
            groups[key].append(sheet)
        
        return groups
        
    except Exception as e:
        print("Warning: Error grouping sheets by calculation: {}".format(e))
        # Return all sheets in one group as fallback
        return {None: initial_sheets}


def expand_calculation_sheets(calculation_guid):
    """Get all sheets in the project that belong to a specific Calculation.
    
    Args:
        calculation_guid: Calculation GUID string, or None for legacy sheets
        
    Returns:
        list: List of DB.ViewSheet elements
    """
    try:
        if calculation_guid is None:
            # Can't expand legacy sheets - return empty list
            return []
        
        # Get schema
        schema_guid = System.Guid(SCHEMA_GUID)
        schema = ESSchema.Lookup(schema_guid)
        
        if not schema:
            print("Warning: pyArea schema not found")
            return []
        
        # Query only elements with pyArea schema (much faster than all sheets)
        from Autodesk.Revit.DB.ExtensibleStorage import ExtensibleStorageFilter
        storage_filter = ExtensibleStorageFilter(schema_guid)
        elements_with_schema = list(
            DB.FilteredElementCollector(doc)
            .WherePasses(storage_filter)
            .ToElements()
        )
        
        # Filter to sheets only and match CalculationGuid
        matching_sheets = []
        for element in elements_with_schema:
            if isinstance(element, DB.ViewSheet):
                sheet_data = get_json_data(element) or {}
                guid = sheet_data.get("CalculationGuid")
                if guid == calculation_guid:
                    matching_sheets.append(element)
        
        print("  Matched {} sheets with CalculationGuid {}".format(len(matching_sheets), calculation_guid[:8]))
        return matching_sheets
        
    except Exception as e:
        print("Warning: Error expanding calculation sheets: {}".format(e))
        import traceback
        traceback.print_exc()
        return []


# ============================================================================
# SECTION 9: MAIN ORCHESTRATION
# ============================================================================

if __name__ == '__main__':
    try:
        print("="*60)
        print("ExportDXF - Area Plans to DXF Export")
        print("="*60)
        
        # 1. Get sheets (active or selected)
        initial_sheets = get_selected_sheets()
        if not initial_sheets or len(initial_sheets) == 0:
            MessageBox.Show(
                "No sheets to export. Please select sheets or open a sheet view.",
                "No Sheets Selected",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning
            )
            sys.exit()
        
        # 2. Group sheets by Calculation
        print("\nGrouping sheets by Calculation...")
        calc_groups = group_sheets_by_calculation(initial_sheets)
        print("Found {} Calculation group(s)".format(len(calc_groups)))
        
        # 3. Load preferences once (used by all exports)
        print("\nLoading preferences...")
        preferences = load_preferences()
        export_folder = export_utils.get_export_folder_path(preferences["ExportFolder"])
        
        # Create export folder if it doesn't exist
        if not os.path.exists(export_folder):
            os.makedirs(export_folder)
            print("Created export folder: {}".format(export_folder))
        
        # 4. Process each Calculation group separately
        exported_files = []
        
        for calc_guid, group_sheets in calc_groups.items():
            print("\n" + "="*60)
            if calc_guid:
                print("Processing Calculation: {}".format(calc_guid[:8]))
            else:
                print("Processing sheets without Calculation (legacy v1.0)")
            print("="*60)
            
            # Expand to all sheets in this Calculation
            if calc_guid:
                all_calc_sheets = expand_calculation_sheets(calc_guid)
                print("Initial sheets in group: {}".format(len(group_sheets)))
                print("Total sheets in Calculation: {}".format(len(all_calc_sheets)))
                sheets_to_export = all_calc_sheets
            else:
                # Legacy sheets - can't expand, just use what was selected
                sheets_to_export = group_sheets
            
            if not sheets_to_export or len(sheets_to_export) == 0:
                print("Warning: No sheets to export in this group, skipping...")
                continue
            
            # Sort sheets (descending - rightmost = page 1)
            sorted_sheets = sort_sheets_by_number(sheets_to_export, descending=True)
            
            # Comprehensive validation: AreaScheme + AreaPlan views + uniform scale
            print("\nValidating sheets and AreaPlan views...")
            try:
                view_scale, valid_viewports_map = get_valid_areaplans_and_uniform_scale(sorted_sheets)
            except ValueError as e:
                # Validation failed - show warning and skip this group
                print("\nValidation failed for this Calculation:")
                print(str(e))
                print("Skipping this Calculation group...\n")
                continue
            
            # Create DXF document
            print("\nCreating DXF document...")
            dxf_doc = ezdxf.new('R2010')  # AutoCAD 2010 format (widely compatible)
            dxf_doc.header['$INSUNITS'] = 5  # 5 = centimeters
            dxf_doc.styles.add('Standard', font='Arial.ttf')
            msp = dxf_doc.modelspace()
            
            # Process each sheet with horizontal offset
            horizontal_offset = 0.0  # In Revit feet
            total_sheets = len(sorted_sheets)
            
            for i, sheet in enumerate(sorted_sheets):
                page_number = total_sheets - i  # Rightmost = page 1
                
                # Get pre-validated viewports for this sheet
                valid_viewports = valid_viewports_map.get(sheet.Id, [])
                
                # Only process sheets with valid viewports
                if len(valid_viewports) > 0:
                    sheet_width = process_sheet(
                        sheet, dxf_doc, msp, horizontal_offset, page_number, view_scale, valid_viewports
                    )
                    
                    # Update horizontal offset for next sheet (add sheet width in feet)
                    horizontal_offset += sheet_width
            
            # Generate filename with Calculation name/guid
            sheet_numbers = [s.SheetNumber for s in sorted_sheets]
            
            # Get Calculation name for filename
            calc_name_part = ""
            if calc_guid:
                # Try to get Calculation name from first sheet's AreaScheme
                try:
                    first_sheet_data = get_sheet_data_for_dxf(sorted_sheets[0])
                    if first_sheet_data:
                        area_scheme = first_sheet_data.get("area_scheme")
                        calc_data = first_sheet_data.get("calculation_data")
                        if calc_data and "Name" in calc_data:
                            calc_name = calc_data["Name"]
                            # Sanitize name for filename
                            calc_name_safe = re.sub(r'[^\w\-_]', '_', calc_name)
                            calc_name_part = "_" + calc_name_safe
                except Exception:
                    # Fall back to guid prefix
                    calc_name_part = "_" + calc_guid[:8]
            
            filename = export_utils.generate_dxf_filename(doc.Title, sheet_numbers) + calc_name_part
            
            dxf_path = os.path.join(export_folder, filename + ".dxf")
            dat_path = os.path.join(export_folder, filename + ".dat")
            
            # Save DXF file
            print("\nSaving DXF file...")
            dxf_doc.saveas(dxf_path)
            print("DXF saved: {}".format(dxf_path))
            
            # Create .dat file with DWFx_SCALE value (if enabled)
            if preferences["DXF_CreateDatFile"]:
                # DWFx files are in millimeters, DXF is in centimeters (real-world scale)
                # When XREFing DWFx into DXF, need to scale by: view_scale / 10
                # Example: 1:100 scale → DWFx_SCALE = 100/10 = 10
                dwfx_scale = int(view_scale / 10)
                print("Creating .dat file...")
                print("  DWFx_SCALE = {} (view scale 1:{})".format(dwfx_scale, int(view_scale)))
                with open(dat_path, 'w') as f:
                    f.write("DWFx_SCALE={}\n".format(dwfx_scale))
                print("DAT saved: {}".format(dat_path))
            else:
                print("\nSkipping .dat file creation (disabled in preferences)")
            
            # Track exported files
            exported_files.append({
                'dxf': os.path.basename(dxf_path),
                'dat': os.path.basename(dat_path) if preferences["DXF_CreateDatFile"] else None,
                'sheets': len(sorted_sheets)
            })
        
        # Report overall results
        print("\n" + "="*60)
        print("EXPORT COMPLETE")
        print("="*60)
        print("Calculation groups processed: {}".format(len(exported_files)))
        print("Output folder: {}".format(export_folder))
        print("\nExported files:")
        for i, file_info in enumerate(exported_files, 1):
            print("\n  {}. {} ({} sheets)".format(i, file_info['dxf'], file_info['sheets']))
            if file_info['dat']:
                print("     {}".format(file_info['dat']))
        print("="*60)
        
    except Exception as e:
        import traceback
        error_msg = "Error during export:\n\n{}".format(str(e))
        print("\n" + "="*60)
        print("ERROR")
        print("="*60)
        print(error_msg)
        print("\nFull traceback:")
        traceback.print_exc()
        print("="*60)
        
        MessageBox.Show(
            error_msg,
            "Export Error",
            MessageBoxButtons.OK,
            MessageBoxIcon.Error
        )