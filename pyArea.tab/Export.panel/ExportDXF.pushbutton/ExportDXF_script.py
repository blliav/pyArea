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

# Current Revit document
doc = revit.doc


# ============================================================================
# SECTION 2: CONSTANTS & CONFIGURATION
# ============================================================================

# Import schema identification from schema_guids.py
from schema_guids import SCHEMA_GUID, SCHEMA_NAME, FIELD_NAME

# Coordinate conversion constants
FEET_TO_CM = 30.48          # Revit internal units (feet) to centimeters
DEFAULT_VIEW_SCALE = 100.0  # Default scale (1:100) if not found

# Import municipality-specific configuration
from municipality_schemas import MUNICIPALITIES, DXF_CONFIG


# ============================================================================
# SECTION 3: DATA EXTRACTION (JSON + Revit API)
# ============================================================================

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
    """Extract sheet data + municipality for DXF export.
    
    Args:
        sheet_elem: DB.ViewSheet element
        
    Returns:
        dict: Sheet data including municipality, or None if error
    """
    try:
        # Get JSON data from sheet
        sheet_data = get_json_data(sheet_elem)
        
        # Get parent AreaScheme
        area_scheme_id = sheet_data.get("AreaSchemeId")
        if not area_scheme_id:
            print("Warning: Sheet {} has no AreaSchemeId".format(sheet_elem.Id))
            return None
        
        area_scheme = get_area_scheme_by_id(area_scheme_id)
        if not area_scheme:
            print("Warning: Could not find AreaScheme {} for sheet {}".format(
                area_scheme_id, sheet_elem.Id))
            return None
        
        # Get municipality
        municipality = get_municipality_from_areascheme(area_scheme)
        
        # Add municipality to data
        sheet_data["Municipality"] = municipality
        
        # Add sheet element reference
        sheet_data["_element"] = sheet_elem
        
        return sheet_data
        
    except Exception as e:
        print("Error getting sheet data: {}".format(e))
        return None


def get_areaplan_data_for_dxf(areaplan_elem):
    """Extract areaplan (view) data for DXF export.
    
    Args:
        areaplan_elem: DB.ViewPlan element (AreaPlan type)
        
    Returns:
        dict: AreaPlan data with element reference
    """
    try:
        # Get JSON data from view
        areaplan_data = get_json_data(areaplan_elem)
        
        # Add element reference
        areaplan_data["_element"] = areaplan_elem
        
        return areaplan_data
        
    except Exception as e:
        print("Warning: Error getting areaplan data for view {}: {}".format(
            areaplan_elem.Id, e))
        return {}


def get_area_data_for_dxf(area_elem):
    """Extract area data + parameters for DXF export.
    
    Args:
        area_elem: DB.Area element
        
    Returns:
        dict: Area data including Usage Type parameters
    """
    try:
        # Get JSON data from area
        area_data = get_json_data(area_elem)
        
        # Get shared parameters
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
        return {}


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


def calculate_arc_bulge(start, end, mid):
    """Calculate DXF bulge value for arc segment.
    
    The bulge is the tangent of 1/4 the included angle of the arc.
    Positive bulge = counterclockwise arc, negative = clockwise.
    
    Args:
        start: Start point (x, y) tuple
        end: End point (x, y) tuple  
        mid: Mid point on arc (x, y) tuple
        
    Returns:
        float: Bulge value for DXF polyline, or 0 if calculation fails
    """
    try:
        # Convert to vectors
        start_x, start_y = start
        end_x, end_y = end
        mid_x, mid_y = mid
        
        # Calculate chord vector
        chord_x = end_x - start_x
        chord_y = end_y - start_y
        chord_length = math.sqrt(chord_x**2 + chord_y**2)
        
        if chord_length < 1e-6:
            return 0.0
        
        # Calculate midpoint of chord
        chord_mid_x = (start_x + end_x) / 2.0
        chord_mid_y = (start_y + end_y) / 2.0
        
        # Calculate sagitta (perpendicular distance from chord midpoint to arc)
        sagitta_x = mid_x - chord_mid_x
        sagitta_y = mid_y - chord_mid_y
        sagitta = math.sqrt(sagitta_x**2 + sagitta_y**2)
        
        # Determine sign (cross product for orientation)
        cross = chord_x * (mid_y - start_y) - chord_y * (mid_x - start_x)
        sign = 1.0 if cross > 0 else -1.0
        
        # Calculate bulge
        # bulge = tan(angle/4) = sagitta / (chord_length/2)
        bulge = sign * (2.0 * sagitta) / chord_length
        
        return bulge
        
    except Exception as e:
        print("Warning: Error calculating arc bulge: {}".format(e))
        return 0.0


# ============================================================================
# SECTION 5: STRING FORMATTING (Municipality-specific)
# ============================================================================

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
        
        # Prepare data with fallbacks
        format_data = {
            "page_number": str(page_number),
            "project": sheet_data.get("PROJECT", ""),
            "elevation": sheet_data.get("ELEVATION", ""),
            "building_height": sheet_data.get("BUILDING_HEIGHT", ""),
            "x": sheet_data.get("X", ""),
            "y": sheet_data.get("Y", ""),
            "lot_area": sheet_data.get("LOT_AREA", ""),
            "scale": sheet_data.get("scale", "100")
        }
        
        # Format string using template
        return template.format(**format_data)
        
    except Exception as e:
        print("Warning: Error formatting sheet string: {}".format(e))
        return "PAGE_NO={}".format(page_number)


def format_areaplan_string(areaplan_data, municipality):
    """Format areaplan attributes using municipality-specific template.
    
    Args:
        areaplan_data: Dictionary with areaplan data
        municipality: Municipality name
        
    Returns:
        str: Formatted attribute string
    """
    try:
        # Get template for this municipality
        template = DXF_CONFIG[municipality]["string_templates"]["areaplan"]
        
        # Prepare data based on municipality
        if municipality == "Jerusalem":
            format_data = {
                "building_name": areaplan_data.get("BUILDING_NAME", "1"),
                "floor_name": areaplan_data.get("FLOOR_NAME", ""),
                "floor_elevation": areaplan_data.get("FLOOR_ELEVATION", ""),
                "floor_underground": areaplan_data.get("FLOOR_UNDERGROUND", "no")
            }
        elif municipality == "Tel-Aviv":
            format_data = {
                "floor": areaplan_data.get("FLOOR", ""),
                "height": areaplan_data.get("HEIGHT", ""),
                "x": areaplan_data.get("X", ""),
                "y": areaplan_data.get("Y", ""),
                "absolute_height": areaplan_data.get("Absolute_height", "")
            }
        else:  # Common
            format_data = {
                "building_no": "1",
                "floor": areaplan_data.get("FLOOR", ""),
                "level_elevation": areaplan_data.get("LEVEL_ELEVATION", ""),
                "is_underground": str(areaplan_data.get("IS_UNDERGROUND", 0))
            }
        
        # Format string using template
        return template.format(**format_data)
        
    except Exception as e:
        print("Warning: Error formatting areaplan string: {}".format(e))
        return "FLOOR="


def format_area_string(area_data, municipality):
    """Format area attributes using municipality-specific template.
    
    Args:
        area_data: Dictionary with area data (includes UsageType, UsageTypePrev)
        municipality: Municipality name
        
    Returns:
        str: Formatted attribute string
    """
    try:
        # Get template for this municipality
        template = DXF_CONFIG[municipality]["string_templates"]["area"]
        
        # Prepare data based on municipality
        if municipality == "Jerusalem":
            format_data = {
                "code": area_data.get("UsageType", ""),
                "demolition_source_code": area_data.get("UsageTypePrev", ""),
                "area": area_data.get("AREA", ""),
                "height1": area_data.get("HEIGHT", ""),
                "appartment_num": area_data.get("APPARTMENT_NUM", ""),
                "height2": area_data.get("HEIGHT2", "")
            }
        elif municipality == "Tel-Aviv":
            format_data = {
                "apartment": area_data.get("APARTMENT", ""),
                "heter": area_data.get("HETER", "1"),
                "height": area_data.get("HEIGHT", "")
            }
        else:  # Common
            format_data = {
                "usage_type": area_data.get("UsageType", ""),
                "usage_type_old": area_data.get("UsageTypePrev", ""),
                "area": area_data.get("AREA", ""),
                "asset": area_data.get("ASSET", "")
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
                'insert': (x, y, 0)
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
        bulges: Optional list of bulge values (same length as points-1)
                Bulge = 0 for straight line, non-zero for arc
    """
    try:
        if not points or len(points) < 2:
            return
        
        # If bulges provided, create polyline with bulges
        if bulges and len(bulges) > 0:
            # Create polyline with bulge values
            polyline = msp.add_lwpolyline(points, dxfattribs={'layer': layer_name})
            # Set bulge for each segment
            for i, bulge in enumerate(bulges):
                if i < len(polyline):
                    polyline[i] = (polyline[i][0], polyline[i][1], 0, 0, bulge)
            polyline.closed = True
        else:
            # Simple polyline without arcs
            polyline = msp.add_lwpolyline(points, dxfattribs={'layer': layer_name})
            polyline.closed = True
        
    except Exception as e:
        print("Warning: Error adding polyline: {}".format(e))


def add_dwfx_underlay(msp, dwfx_path, insert_point, scale_factor):
    """Add DWFX underlay reference to DXF (optional feature).
    
    Args:
        msp: DXF modelspace
        dwfx_path: Path to DWFX file
        insert_point: (x, y) tuple for insertion point
        scale_factor: Scale factor for underlay
        
    Note: This is an advanced feature. May not be supported by all DXF viewers.
    """
    try:
        # DWFX underlay support in ezdxf
        # This is optional and may require additional configuration
        print("Note: DWFX underlay support not implemented in this version")
        pass
        
    except Exception as e:
        print("Warning: Error adding DWFX underlay: {}".format(e))


# ============================================================================
# SECTION 7: PROCESSING PIPELINE
# ============================================================================

def process_area(area_elem, viewport, msp, scale_factor, offset_x, offset_y, municipality, layers):
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
    """
    try:
        # Get area data
        area_data = get_area_data_for_dxf(area_elem)
        if not area_data:
            return
        
        # Get boundary segments
        boundary_options = DB.SpatialElementBoundaryOptions()
        boundary_segments = area_elem.GetBoundarySegments(boundary_options)
        
        if not boundary_segments or len(boundary_segments) == 0:
            print("  Warning: Area {} has no boundary".format(area_elem.Id))
            return
        
        # Process only the exterior boundary loop (ignore holes)
        # boundary_segments[0] is the exterior, [1..n] are holes
        exterior_loop = boundary_segments[0]
        
        # Collect all boundary points
        boundary_points = []
        bulges = []
        
        for segment in exterior_loop:
            # Extract curve from BoundarySegment (in VIEW coordinates)
            curve = segment.GetCurve()
            start_pt_view = curve.GetEndPoint(0)
            end_pt_view = curve.GetEndPoint(1)
            
            # Transform VIEW to SHEET coordinates
            start_pt_sheet = transform_point_to_sheet(start_pt_view, viewport)
            boundary_points.append(start_pt_sheet)
            
            # Check if curve is an arc
            if isinstance(curve, DB.Arc):
                try:
                    # Get midpoint for bulge calculation (in VIEW coordinates)
                    mid_param = (curve.GetEndParameter(0) + curve.GetEndParameter(1)) / 2.0
                    mid_pt_view = curve.Evaluate(mid_param, True)
                    
                    # Transform to SHEET coordinates
                    mid_pt_sheet = transform_point_to_sheet(mid_pt_view, viewport)
                    end_pt_sheet = transform_point_to_sheet(end_pt_view, viewport)
                    
                    # Calculate bulge using SHEET coordinates
                    bulge = calculate_arc_bulge(
                        (start_pt_sheet.X, start_pt_sheet.Y),
                        (end_pt_sheet.X, end_pt_sheet.Y),
                        (mid_pt_sheet.X, mid_pt_sheet.Y)
                    )
                    bulges.append(bulge)
                except:
                    bulges.append(0.0)  # Fallback to straight line
            else:
                bulges.append(0.0)  # Straight line
        
        # Close the boundary
        if len(boundary_points) > 0:
            boundary_points.append(boundary_points[0])
            bulges.append(0.0)
        
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
            area_string = format_area_string(area_data, municipality)
            
            # Add text
            add_text(msp, area_string, text_pos, layers['area_text'])
        
    except Exception as e:
        print("  Warning: Error processing area {}: {}".format(area_elem.Id, e))


def process_areaplan_viewport(viewport, msp, scale_factor, offset_x, offset_y, municipality, layers):
    """Process AreaPlan viewport - add crop boundary, plan text, and all areas.
    
    Args:
        viewport: DB.Viewport element
        msp: DXF modelspace
        scale_factor: REALWORLD_SCALE_FACTOR
        offset_x: Horizontal offset (cm)
        offset_y: Vertical offset (cm)
        municipality: Municipality name
        layers: Layer name mapping
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
        
        # Get areaplan data
        areaplan_data = get_areaplan_data_for_dxf(view)
        
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
                        
                        areaplan_string = format_areaplan_string(areaplan_data, municipality)
                        add_text(msp, areaplan_string, text_pos, layers['areaplan_text'])
        
        # Get all areas in this view
        collector = DB.FilteredElementCollector(doc, view_id)
        areas = collector.OfCategory(DB.BuiltInCategory.OST_Areas).WhereElementIsNotElementType().ToElements()
        
        print("    Found {} areas".format(len(areas)))
        
        # Process each area
        for area in areas:
            if isinstance(area, DB.Area):
                process_area(area, viewport, msp, scale_factor, offset_x, offset_y, municipality, layers)
        
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
        
        # Get sheet data
        sheet_data = get_sheet_data_for_dxf(sheet_elem)
        if not sheet_data:
            print("  Warning: No sheet data found")
            return 0.0
        
        # Get municipality from sheet data
        municipality = sheet_data.get("Municipality", "Common")
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
        
        # Add sheet frame rectangle (titleblock outline)
        if titleblock and bbox:
            min_point = convert_point_to_realworld(bbox.Min, scale_factor, offset_x, offset_y)
            max_point = convert_point_to_realworld(bbox.Max, scale_factor, offset_x, offset_y)
            add_rectangle(msp, min_point, max_point, layers['sheet_frame'])
        
        # Add sheet text at top-right corner
        sheet_string = format_sheet_string(sheet_data, municipality, page_number)
        if titleblock and bbox:
            # Get top-right corner in DXF space
            max_point_dxf = convert_point_to_realworld(bbox.Max, scale_factor, offset_x, offset_y)
            # Position text 10 cm below and to the left in DXF space
            text_pos = (max_point_dxf[0] - 10.0, max_point_dxf[1] - 10.0)
            add_text(msp, sheet_string, text_pos, layers['sheet_text'])
        
        # Process pre-validated viewports
        for viewport in valid_viewports:
            process_areaplan_viewport(
                viewport, msp, scale_factor, offset_x, offset_y, municipality, layers
            )
        
        return sheet_width
        
    except Exception as e:
        print("Error processing sheet: {}".format(e))
        return 0.0


# ============================================================================
# SECTION 8: SHEET SELECTION & SORTING
# ============================================================================

def get_valid_areaplans_and_uniform_scale(sheets):
    """Validate AreaPlan views and ensure uniform scale.
    
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


# ============================================================================
# SECTION 9: MAIN ORCHESTRATION
# ============================================================================

if __name__ == '__main__':
    try:
        print("="*60)
        print("ExportDXF - Area Plans to DXF Export")
        print("="*60)
        
        # 1. Get sheets (active or selected)
        sheets = get_selected_sheets()
        if not sheets or len(sheets) == 0:
            MessageBox.Show(
                "No sheets to export. Please select sheets or open a sheet view.",
                "No Sheets Selected",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning
            )
            sys.exit()
        
        # 2. Sort sheets (descending - rightmost = page 1)
        sorted_sheets = sort_sheets_by_number(sheets, descending=True)
        
        # 3. Validate AreaPlan views and get uniform scale
        print("\nValidating AreaPlan views...")
        try:
            view_scale, valid_viewports_map = get_valid_areaplans_and_uniform_scale(sorted_sheets)
        except ValueError as e:
            # Validation failed - show error and exit
            print("\n" + str(e))
            MessageBox.Show(
                str(e),
                "Validation Error",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error
            )
            sys.exit()
        
        # 4. Create DXF document
        print("\nCreating DXF document...")
        dxf_doc = ezdxf.new('R2010')  # AutoCAD 2010 format (widely compatible)
        dxf_doc.header['$INSUNITS'] = 5  # 5 = centimeters
        msp = dxf_doc.modelspace()
        
        # 5. Process each sheet with horizontal offset
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
        
        # 6. Determine output path
        # Use Desktop/Export/ folder
        desktop = os.path.join(os.path.expanduser("~"), "Desktop")
        export_folder = os.path.join(desktop, "Export")
        
        # Create export folder if it doesn't exist
        if not os.path.exists(export_folder):
            os.makedirs(export_folder)
        
        # Generate filename: <modelname>-<firstsheet>..<lastsheet>
        model_name = doc.Title if doc.Title else "Model"
        model_name = model_name.replace(" ", "_").replace("/", "_").replace("\\", "_")
        
        first_sheet = sorted_sheets[0]
        last_sheet = sorted_sheets[-1]
        
        # Build sheet range string
        if len(sorted_sheets) == 1:
            sheet_range = first_sheet.SheetNumber.replace("/", "_")
        else:
            sheet_range = "{}..{}".format(
                first_sheet.SheetNumber.replace("/", "_"),
                last_sheet.SheetNumber.replace("/", "_")
            )
        
        # Add timestamp (commented out)
        # from datetime import datetime
        # timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        # filename = "{}-{}_{}".format(model_name, sheet_range, timestamp)
        filename = "{}-{}".format(model_name, sheet_range)
        
        dxf_path = os.path.join(export_folder, filename + ".dxf")
        dat_path = os.path.join(export_folder, filename + ".dat")
        
        # 7. Save DXF file
        print("\nSaving DXF file...")
        dxf_doc.saveas(dxf_path)
        print("DXF saved: {}".format(dxf_path))
        
        # 8. Create .dat file with DWFX_SCALE value
        # DWFX files are in millimeters, DXF is in centimeters (real-world scale)
        # When XREFing DWFX into DXF, need to scale by: view_scale / 10
        # Example: 1:100 scale → DWFX_SCALE = 100/10 = 10
        dwfx_scale = int(view_scale / 10)
        print("Creating .dat file...")
        print("  DWFX_SCALE = {} (view scale 1:{})".format(dwfx_scale, int(view_scale)))
        with open(dat_path, 'w') as f:
            f.write("DWFX_SCALE={}\n".format(dwfx_scale))
        print("DAT saved: {}".format(dat_path))
        
        # 9. Report results
        print("\n" + "="*60)
        print("EXPORT COMPLETE")
        print("="*60)
        print("Sheets exported: {}".format(len(sorted_sheets)))
        print("Output folder: {}".format(export_folder))
        print("DXF file: {}".format(os.path.basename(dxf_path)))
        print("DAT file: {}".format(os.path.basename(dat_path)))
        print("="*60)
        
        # Show success message
        MessageBox.Show(
            "Export successful!\n\n"
            "Exported {} sheet(s)\n"
            "Files saved to: {}\n\n"
            "{}\n{}".format(
                len(sorted_sheets),
                export_folder,
                os.path.basename(dxf_path),
                os.path.basename(dat_path)
            ),
            "Export Complete",
            MessageBoxButtons.OK,
            MessageBoxIcon.Information
        )
        
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