#! python3
# -*- coding: utf-8 -*-
"""
Export Floor Plans to DXF for MOT (Ministry of Tourism).
Exports room boundaries with layers based on room's "MOT classification" parameter.
No municipalities, no blocks, no JSON data - simple geometry export only.
"""

# ============================================================================
# SECTION 1: IMPORTS & SETUP
# ============================================================================

# Standard library imports
import sys
import os
import math
import re

# Script directory for path resolution
script_dir = os.path.dirname(__file__)

# Add lib directory (contains export_utils, python_utils)
lib_dir = os.path.join(script_dir, '..', '..', 'lib')
lib_dir = os.path.abspath(lib_dir)
if lib_dir not in sys.path:
    sys.path.insert(0, lib_dir)

# pyRevit imports (CPython compatible)
from pyrevit import revit, DB, script

# External package - auto-download from PyPI if missing
from python_utils import ensure_vendor_cpython_in_path, install_packages_from_pypi

# Packages required for DXF export (ezdxf + all dependencies)
EZDXF_PACKAGES = [
    'typing_extensions',
    'pyparsing', 
    'numpy',
    'fontTools',  # ezdxf dependency for font handling
    'ezdxf'
]

# Ensure vendor_cpython directory is in path
ensure_vendor_cpython_in_path()

try:
    import ezdxf
except ImportError:
    print("pyArea: Installing ezdxf (first-time setup)...")
    install_packages_from_pypi(EZDXF_PACKAGES)
    # Clear import cache so Python finds newly installed packages
    import importlib
    importlib.invalidate_caches()
    import ezdxf

# .NET interop
import clr
import System
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
clr.AddReference('System.Windows.Forms')
from System.Windows.Forms import MessageBox, MessageBoxButtons, MessageBoxIcon, DialogResult

# Import export utilities
import export_utils

# Current Revit document
doc = revit.doc


# ============================================================================
# SECTION 2: CONSTANTS & CONFIGURATION
# ============================================================================

# Coordinate conversion constants
FEET_TO_CM = 30.48          # Revit internal units (feet) to centimeters
FEET_TO_METERS = 0.3048     # Revit internal units (feet) to meters
DEFAULT_VIEW_SCALE = 100.0  # Default scale (1:100) if not found

# Layer configuration for MOT export
FRAME_LAYER = 'Frame'
FRAME_LAYER_COLOR = 7  # White

# True-color (RGB) per MOT classification. Keys must match the parameter value exactly.
MOT_LAYER_COLORS = {
    'יחידות אירוח': (44, 160, 44),
    'יחידות אירוח-מעברים ואחסון': (152, 223, 138),
    'יחידות אירוח-רחצה': (174, 199, 232),
}

# Fallback color for any classification not listed above. A MERM room usually carries
# a non-guestroom classification, so this color and MERM_LAYER_COLOR end up drawn
# concentrically - they were picked as a pair to stay readable in that case.
DEFAULT_MOT_LAYER_COLOR = (255, 0, 255)  # Magenta

# Global (constant) width applied to room classification polylines, in DXF units (cm)
POLYLINE_GLOBAL_WIDTH = 5.0

# Dedicated layer for protected spaces, matched by room NAME instead of classification.
# A room can be exported twice: once by classification, once here.
MERM_LAYER = 'MERM'
MERM_LAYER_COLOR = (218, 145, 0)  # Ochre - warm, separates from the magenta core
MERM_POLYLINE_WIDTH = 20.0
MERM_NAME_KEYWORDS = ['ממ"מ', 'ממ"ק', 'ממ"ד', 'מקלט']

# DWFx underlay fade percentage. DXF allows 0-80; higher is more faded.
UNDERLAY_FADE = 70


# ============================================================================
# SECTION 3: UTILITIES
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


def load_preferences():
    """
    Load export preferences from user (AppData).
    User prefs: ExportFolder stored in %APPDATA%\\pyArea\\preferences.json
    
    Returns:
        dict: Preferences dictionary
    """
    try:
        # Start with defaults
        prefs = export_utils.get_default_preferences()
        
        # Load user preferences from AppData (ExportFolder only)
        appdata = os.environ.get('APPDATA', os.path.expanduser('~'))
        prefs_file = os.path.join(appdata, 'pyArea', 'preferences.json')
        if os.path.exists(prefs_file):
            import json
            with open(prefs_file, 'r', encoding='utf-8') as f:
                user_prefs = json.load(f)
            if "ExportFolder" in user_prefs:
                prefs["ExportFolder"] = user_prefs["ExportFolder"]
        
        return prefs
        
    except Exception as e:
        print("Warning: Failed to load preferences, using defaults: {}".format(str(e)))
        return export_utils.get_default_preferences()


# ============================================================================
# SECTION 4: COORDINATE & GEOMETRY UTILITIES
# ============================================================================

def calculate_realworld_scale_factor(view_scale):
    """Compute real-world scale factor: FEET_TO_CM * view_scale.
    
    Args:
        view_scale: View scale number (e.g., 100 for 1:100, 200 for 1:200)
        
    Returns:
        float: Scale factor to convert sheet feet to real-world cm
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
    """
    return ((xyz.X - offset_x) * scale_factor, (xyz.Y - offset_y) * scale_factor)


def transform_point_to_sheet(view_point, viewport):
    """Transform point from view coordinates to sheet coordinates.
    
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
# SECTION 5: DXF LAYER & ENTITY CREATION
# ============================================================================

def create_frame_layer(dxf_doc):
    """Create Frame layer for titleblock and view crop boundaries.
    
    Args:
        dxf_doc: ezdxf DXF document
    """
    try:
        if FRAME_LAYER not in dxf_doc.layers:
            layer = dxf_doc.layers.new(name=FRAME_LAYER)
            layer.color = FRAME_LAYER_COLOR
    except Exception as e:
        print("Error creating Frame layer: {}".format(e))


def create_merm_layer(dxf_doc):
    """Create MERM layer for protected-space room boundaries.
    
    Args:
        dxf_doc: ezdxf DXF document
    """
    try:
        if MERM_LAYER not in dxf_doc.layers:
            layer = dxf_doc.layers.new(name=MERM_LAYER)
            layer.rgb = MERM_LAYER_COLOR
    except Exception as e:
        print("Error creating MERM layer: {}".format(e))


def sanitize_layer_name(layer_name):
    """Sanitize layer name for AutoCAD compatibility.
    
    AutoCAD layer names cannot contain: < > / \ " : ; ? * | = ` , (comma)
    Hebrew and other Unicode characters are supported by AutoCAD.
    
    Args:
        layer_name: Original layer name
        
    Returns:
        str: Sanitized layer name safe for AutoCAD
    """
    if not layer_name:
        return "Layer_0"
    
    # Remove only the characters that AutoCAD doesn't allow in layer names
    # Keep Hebrew and other Unicode characters - AutoCAD supports them
    sanitized = re.sub(r'[<>/\\\":;?*|=`,]', '_', layer_name)
    
    # Clean up multiple underscores
    sanitized = re.sub(r'_+', '_', sanitized)
    
    # Remove leading/trailing underscores and spaces
    sanitized = sanitized.strip('_ ')
    
    # Ensure it's not empty
    if not sanitized:
        sanitized = "Layer_0"
    
    # Limit length to 255 chars
    if len(sanitized) > 255:
        sanitized = sanitized[:255]
    
    return sanitized


def ensure_layer_exists(dxf_doc, layer_name):
    """Ensure a layer exists in the DXF document with sanitized name.
    
    Args:
        dxf_doc: ezdxf DXF document
        layer_name: Layer name to create (will be sanitized)
        
    Returns:
        str: The actual sanitized layer name that was created/used
    """
    try:
        if not layer_name:
            return "0"  # Default layer
        
        # Sanitize the layer name for AutoCAD compatibility
        safe_name = sanitize_layer_name(layer_name)
        
        if safe_name not in dxf_doc.layers:
            layer = dxf_doc.layers.new(name=safe_name)
            layer.rgb = MOT_LAYER_COLORS.get(layer_name, DEFAULT_MOT_LAYER_COLOR)
            if safe_name != layer_name:
                print("      Note: Layer '{}' sanitized to '{}'".format(layer_name, safe_name))
        
        return safe_name
        
    except Exception as e:
        print("Warning: Error creating layer '{}': {}".format(layer_name, e))
        return "0"  # Fall back to default layer


def add_rectangle(msp, min_point, max_point, layer_name, width=0.0):
    """Add rectangle to DXF using polyline.
    
    Args:
        msp: DXF modelspace
        min_point: (x, y) tuple for bottom-left corner
        max_point: (x, y) tuple for top-right corner
        layer_name: DXF layer name
        width: Constant polyline width in DXF units (0 = hairline)
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
        
        polyline = msp.add_lwpolyline(points, dxfattribs=_polyline_attribs(layer_name, width))
        polyline.closed = True
        
    except Exception as e:
        print("Warning: Error adding rectangle: {}".format(e))


def _polyline_attribs(layer_name, width):
    """Build LWPOLYLINE dxfattribs, omitting const_width when width is 0 (hairline).
    
    Args:
        layer_name: DXF layer name
        width: Constant polyline width in DXF units
        
    Returns:
        dict: dxfattribs for add_lwpolyline
    """
    attribs = {'layer': layer_name}
    if width:
        attribs['const_width'] = width
    return attribs


def add_polyline_with_arcs(msp, points, layer_name, bulges=None, width=0.0):
    """Add polyline with optional arc segments (bulges) to DXF.
    
    Args:
        msp: DXF modelspace
        points: List of (x, y) tuples
        layer_name: DXF layer name
        bulges: Optional list of bulge values (same length as points)
        width: Constant polyline width in DXF units (0 = hairline)
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
            polyline = msp.add_lwpolyline(points_with_bulge, format='xyseb',
                                         dxfattribs=_polyline_attribs(layer_name, width))
            polyline.closed = True
        else:
            # Simple polyline without arcs
            polyline = msp.add_lwpolyline(points, dxfattribs=_polyline_attribs(layer_name, width))
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
        
        # Fade it back so the exported room boundaries read on top of it
        underlay.dxf.fade = UNDERLAY_FADE
        
        print("  Added DWFx underlay: {} (scale: {}, fade: {})".format(
            dwfx_filename, dwfx_scale, UNDERLAY_FADE))
        return True
        
    except Exception as e:
        print("  Warning: Error adding DWFx underlay: {}".format(e))
        return False


def send_underlays_to_back(msp):
    """Force all underlay entities to the back of the draw order.
    
    Creation order alone is not enough: underlays are added per sheet, so a later
    sheet's underlay would otherwise sit above an earlier sheet's geometry. This
    writes an explicit SORTENTSTABLE instead.
    
    Sort-handles are taken from the pool of existing entity handles, sorted
    ascending, so every assigned value is a valid handle. Handle "0" is avoided
    because AutoCAD draws it last rather than first.
    
    Args:
        msp: DXF modelspace
    """
    try:
        entities = list(msp)
        if not entities:
            return
        
        underlays = [e for e in entities if e.dxftype().endswith('UNDERLAY')]
        if not underlays:
            return
        others = [e for e in entities if not e.dxftype().endswith('UNDERLAY')]
        
        # Ascending pool of valid sort-handles, one per entity
        pool = sorted((e.dxf.handle for e in entities), key=lambda h: int(h, 16))
        
        # Underlays first, remaining entities keep their relative creation order
        ordered = underlays + others
        msp.set_redraw_order(
            [(e.dxf.handle, pool[i]) for i, e in enumerate(ordered)]
        )
        
        print("Draw order: {} underlay(s) sent to back".format(len(underlays)))
        
    except Exception as e:
        print("Warning: Could not set underlay draw order: {}".format(e))


# ============================================================================
# SECTION 6: ROOM PROCESSING
# ============================================================================

def get_mot_classification(room_elem):
    """Get MOT classification parameter value from room.
    
    Args:
        room_elem: DB.Architecture.Room element
        
    Returns:
        str: MOT classification value, or None if not set
    """
    try:
        param = room_elem.LookupParameter("MOT Classification")
        if not param or not param.HasValue:
            return None
        
        # Try different methods to get the value based on storage type
        storage_type = param.StorageType
        value = None
        
        if storage_type == DB.StorageType.String:
            value = param.AsString()
        elif storage_type == DB.StorageType.Integer:
            value = str(param.AsInteger())
        elif storage_type == DB.StorageType.Double:
            value = str(param.AsDouble())
        elif storage_type == DB.StorageType.ElementId:
            # For element ID parameters, get the element and its name
            elem_id = param.AsElementId()
            if elem_id and elem_id != DB.ElementId.InvalidElementId:
                elem = doc.GetElement(elem_id)
                if elem:
                    value = elem.Name
        
        if value and str(value).strip():
            return str(value).strip()
        
        return None
        
    except Exception as e:
        print("  Warning: Error reading MOT classification for room {}: {}".format(room_elem.Id, e))
        return None


def get_room_name(room_elem):
    """Get the room's Name parameter value.
    
    Uses the ROOM_NAME built-in parameter rather than Room.Name, which in some
    Revit versions returns the name concatenated with the room number.
    
    Args:
        room_elem: DB.Architecture.Room element
        
    Returns:
        str: Room name, or None if not set
    """
    try:
        param = room_elem.get_Parameter(DB.BuiltInParameter.ROOM_NAME)
        if param and param.HasValue:
            value = param.AsString()
            if value and value.strip():
                return value.strip()
        return None
    except Exception as e:
        print("  Warning: Error reading name for room {}: {}".format(room_elem.Id, e))
        return None


def is_merm_room(room_name):
    """Check whether a room name marks it as a protected space.
    
    Matches any of MERM_NAME_KEYWORDS as a substring. Hebrew gershayim (U+05F4)
    and typographic quotes are normalized to ASCII so both spellings match.
    
    Args:
        room_name: Room name string, or None
        
    Returns:
        bool: True if the room belongs on the MERM layer
    """
    if not room_name:
        return False
    
    normalized = room_name.replace('\u05f4', '"').replace('\u201c', '"').replace('\u201d', '"')
    return any(keyword in normalized for keyword in MERM_NAME_KEYWORDS)


def get_room_exterior_loop(room_elem):
    """Get the exterior boundary loop of a room as the raw BoundarySegment list.
    
    Uses signed area (Shoelace formula) to identify the exterior loop.
    
    Args:
        room_elem: DB.Architecture.Room element
        
    Returns:
        The exterior BoundarySegment loop, or None if not found
    """
    try:
        boundary_options = DB.SpatialElementBoundaryOptions()
        boundary_segments = room_elem.GetBoundarySegments(boundary_options)
        
        if not boundary_segments or len(boundary_segments) == 0:
            return None
        
        exterior_loop = None
        max_abs_area = 0.0
        for loop in boundary_segments:
            pts = []
            for segment in loop:
                curve = segment.GetCurve()
                tess_pts = list(curve.Tessellate())
                pts.extend(tess_pts[:-1])
            if len(pts) < 3:
                continue
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
        
        return exterior_loop
        
    except Exception as e:
        print("  Warning: Error getting exterior loop for room {}: {}".format(room_elem.Id, e))
        return None


def get_room_boundary_polyline_dxf(room_elem, viewport, scale_factor, offset_x, offset_y):
    """Build room's exterior boundary polyline in DXF coordinates with bulges.
    
    Args:
        room_elem: DB.Architecture.Room element
        viewport: DB.Viewport element
        scale_factor: REALWORLD_SCALE_FACTOR
        offset_x: Horizontal offset (feet)
        offset_y: Vertical offset (feet)

    Returns:
        tuple: (dxf_pts, bulges)
            dxf_pts: closed list of (x, y) in DXF coordinates
            bulges: list of bulge values (or None if no arcs)
    """
    exterior_loop = get_room_exterior_loop(room_elem)
    if exterior_loop is None:
        return None, None

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

    if not boundary_points:
        return None, None

    transformed_points = [
        convert_point_to_realworld(pt, scale_factor, offset_x, offset_y)
        for pt in boundary_points
    ]
    return transformed_points, bulges


def process_room(room_elem, viewport, msp, scale_factor, offset_x, offset_y, dxf_doc):
    """Process single Room element - add boundary to DXF on MOT classification layer.
    
    Args:
        room_elem: DB.Architecture.Room element
        viewport: DB.Viewport element (for coordinate transformation)
        msp: DXF modelspace
        scale_factor: REALWORLD_SCALE_FACTOR
        offset_x: Horizontal offset (feet)
        offset_y: Vertical offset (feet)
        dxf_doc: ezdxf DXF document
    """
    try:
        # Get MOT classification
        mot_class = get_mot_classification(room_elem)
        if not mot_class:
            # Skip rooms without MOT classification
            return
        
        # Get boundary
        transformed_points, bulges = get_room_boundary_polyline_dxf(
            room_elem, viewport, scale_factor, offset_x, offset_y)
        if not transformed_points:
            print("      Warning: Room {} has no boundary".format(room_elem.Id))
            return
        
        print("      Room boundary has {} points".format(len(transformed_points)))
        
        # Ensure layer exists (returns sanitized layer name)
        safe_layer_name = ensure_layer_exists(dxf_doc, mot_class)
        
        # Add boundary polyline on MOT classification layer
        add_polyline_with_arcs(
            msp, 
            transformed_points, 
            safe_layer_name,
            bulges if any(b != 0 for b in bulges) else None,
            width=POLYLINE_GLOBAL_WIDTH
        )
        print("      Added room boundary to layer '{}'".format(safe_layer_name))
        
    except Exception as e:
        print("  Warning: Error processing room {}: {}".format(room_elem.Id, e))


def process_merm_room(room_elem, viewport, msp, scale_factor, offset_x, offset_y):
    """Add a protected-space room's boundary to the MERM layer.
    
    Independent of the classification pass - the same room may be exported by both.
    
    Args:
        room_elem: DB.Architecture.Room element
        viewport: DB.Viewport element (for coordinate transformation)
        msp: DXF modelspace
        scale_factor: REALWORLD_SCALE_FACTOR
        offset_x: Horizontal offset (feet)
        offset_y: Vertical offset (feet)
        
    Returns:
        bool: True if a boundary was added
    """
    try:
        transformed_points, bulges = get_room_boundary_polyline_dxf(
            room_elem, viewport, scale_factor, offset_x, offset_y)
        if not transformed_points:
            print("      Warning: MERM room {} has no boundary".format(room_elem.Id))
            return False
        
        add_polyline_with_arcs(
            msp,
            transformed_points,
            MERM_LAYER,
            bulges if any(b != 0 for b in bulges) else None,
            width=MERM_POLYLINE_WIDTH
        )
        return True
        
    except Exception as e:
        print("  Warning: Error processing MERM room {}: {}".format(room_elem.Id, e))
        return False


def get_view_elevation(view):
    """Get the elevation of a plan view's associated level, in feet.
    
    Args:
        view: DB.ViewPlan element
        
    Returns:
        float: Level elevation in feet, or 0.0 if it cannot be determined
    """
    try:
        level = view.GenLevel
        if level is not None:
            return float(level.Elevation)
    except Exception:
        pass
    
    try:
        level = doc.GetElement(view.LevelId)
        if level is not None and hasattr(level, 'Elevation'):
            return float(level.Elevation)
    except Exception:
        pass
    
    return 0.0


def build_room_view_assignments(valid_viewports_map):
    """Assign each room to the single lowest plan view that shows it.
    
    A room spanning more than one level (double-height hall, atrium) appears in
    every plan whose view range it crosses, so it would otherwise be exported once
    per plan. Each room is pinned to the viewport whose level sits lowest; ties are
    broken by viewport id so the result is deterministic across runs.
    
    Args:
        valid_viewports_map: {sheet.Id: [DB.Viewport]} from validation
        
    Returns:
        dict: {room id value: viewport id value} - the only viewport allowed to
              export that room
    """
    best = {}
    occurrences = {}
    
    try:
        for viewports in valid_viewports_map.values():
            for viewport in viewports:
                view = doc.GetElement(viewport.ViewId)
                if not view:
                    continue
                
                candidate_key = (get_view_elevation(view),
                                 get_element_id_value(viewport.Id))
                
                rooms = DB.FilteredElementCollector(doc, view.Id)\
                    .OfCategory(DB.BuiltInCategory.OST_Rooms)\
                    .WhereElementIsNotElementType()\
                    .ToElements()
                
                for room in rooms:
                    if not isinstance(room, DB.Architecture.Room):
                        continue
                    room_key = get_element_id_value(room.Id)
                    occurrences[room_key] = occurrences.get(room_key, 0) + 1
                    if room_key not in best or candidate_key < best[room_key]:
                        best[room_key] = candidate_key
        
        multi_level = sum(1 for count in occurrences.values() if count > 1)
        print("Room/view assignment: {} room(s) total, {} shown in more than one plan".format(
            len(best), multi_level))
        
    except Exception as e:
        print("Warning: Could not build room/view assignments: {}".format(e))
        return {}
    
    return dict((room_key, key[1]) for room_key, key in best.items())


def process_floorplan_viewport(viewport, msp, scale_factor, offset_x, offset_y, dxf_doc,
                               room_assignments=None):
    """Process FloorPlan or AreaPlan viewport - add crop boundary and all rooms.
    
    Args:
        viewport: DB.Viewport element
        msp: DXF modelspace
        scale_factor: REALWORLD_SCALE_FACTOR
        offset_x: Horizontal offset (feet)
        offset_y: Vertical offset (feet)
        dxf_doc: ezdxf DXF document
        room_assignments: {room id value: viewport id value} restricting each room
                          to one viewport; None exports every room in the view
    """
    try:
        # Get the view from viewport
        view_id = viewport.ViewId
        view = doc.GetElement(view_id)
        
        if not view or not isinstance(view, DB.ViewPlan):
            return
        
        view_type_name = "AreaPlan" if view.ViewType == DB.ViewType.AreaPlan else "FloorPlan"
        print("  Processing {}: {}".format(view_type_name, view.Name))
        
        # Get all rooms in this view
        collector = DB.FilteredElementCollector(doc, view_id)
        rooms = collector.OfCategory(DB.BuiltInCategory.OST_Rooms).WhereElementIsNotElementType().ToElements()
        room_list = [r for r in rooms if isinstance(r, DB.Architecture.Room)]
        
        print("    Found {} rooms".format(len(room_list)))
        
        # Drop rooms pinned to a lower plan, so a multi-level room is exported once.
        # Applied here, before both passes, so MERM and classification stay in sync.
        if room_assignments:
            vp_key = get_element_id_value(viewport.Id)
            kept = [r for r in room_list
                    if room_assignments.get(get_element_id_value(r.Id), vp_key) == vp_key]
            skipped = len(room_list) - len(kept)
            if skipped > 0:
                print("    Skipping {} room(s) already exported on a lower plan".format(skipped))
            room_list = kept
        
        # MERM pass first: entities added earlier sit behind later ones in AutoCAD's
        # draw order, so these wide outlines stay behind the classification boundaries
        merm_added = 0
        for room in room_list:
            if is_merm_room(get_room_name(room)):
                if process_merm_room(room, viewport, msp, scale_factor, offset_x, offset_y):
                    merm_added += 1
        if merm_added > 0:
            print("    Added {} room(s) to layer '{}'".format(merm_added, MERM_LAYER))
        
        # Process each room
        rooms_processed = 0
        for room in room_list:
            process_room(room, viewport, msp, scale_factor, offset_x, offset_y, dxf_doc)
            rooms_processed += 1
        
        print("    Processed {} rooms".format(rooms_processed))
        
        # Draw view crop boundary on Frame layer
        crop_manager = view.GetCropRegionShapeManager()
        if not crop_manager:
            return
        crop_shape = crop_manager.GetCropShape()
        if not crop_shape or crop_shape.Count == 0:
            return
        
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
        if len(crop_points_sheet) == 0:
            return
        crop_points_sheet.append(crop_points_sheet[0])
        
        # Transform SHEET coordinates to DXF coordinates
        transformed_crop = [
            convert_point_to_realworld(pt, scale_factor, offset_x, offset_y)
            for pt in crop_points_sheet
        ]
        
        # Add crop boundary on Frame layer
        add_polyline_with_arcs(msp, transformed_crop, FRAME_LAYER)
        
    except Exception as e:
        print("  Warning: Error processing viewport: {}".format(e))
        import traceback
        traceback.print_exc()


def process_sheet(sheet_elem, dxf_doc, msp, horizontal_offset, view_scale, valid_viewports,
                  room_assignments=None):
    """Process entire sheet with horizontal offset for multi-sheet layout.
    
    Args:
        sheet_elem: DB.ViewSheet element
        dxf_doc: ezdxf DXF document
        msp: DXF modelspace
        horizontal_offset: Horizontal offset for this sheet (Revit feet)
        view_scale: Validated uniform view scale for entire export
        valid_viewports: List of pre-validated DB.Viewport elements to process
        room_assignments: {room id value: viewport id value} restricting each room
                          to a single viewport
        
    Returns:
        float: Width of this sheet in Revit feet (for next sheet's offset)
    """
    try:
        print("\n" + "-"*60)
        print("Processing Sheet: {} - {}".format(sheet_elem.SheetNumber, sheet_elem.Name))
        
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
        offset_x = bbox.Min.X - horizontal_offset
        offset_y = bbox.Min.Y
        
        print("  Offset: X={:.3f} ft, Y={:.3f} ft (horizontal_offset={:.3f} ft)".format(
            offset_x, offset_y, horizontal_offset))
        
        # Add DWFx underlay (background reference)
        print("  Attempting to add DWFx underlay...")
        dwfx_filename = export_utils.generate_dwfx_filename(doc, sheet_elem) + ".dwfx"
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
        
        # Add sheet frame rectangle (titleblock outline) on Frame layer
        if titleblock and bbox:
            min_point = convert_point_to_realworld(bbox.Min, scale_factor, offset_x, offset_y)
            max_point = convert_point_to_realworld(bbox.Max, scale_factor, offset_x, offset_y)
            add_rectangle(msp, min_point, max_point, FRAME_LAYER)
        
        # Process pre-validated viewports
        for viewport in valid_viewports:
            process_floorplan_viewport(
                viewport, msp, scale_factor, offset_x, offset_y, dxf_doc,
                room_assignments=room_assignments
            )
        
        return sheet_width
        
    except Exception as e:
        print("Error processing sheet: {}".format(e))
        return 0.0


# ============================================================================
# SECTION 7: SHEET SELECTION & VALIDATION
# ============================================================================

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
        
        # No valid selection
        print("No sheets selected. Please select sheets in project browser or open a sheet.")
        return None
        
    except Exception as e:
        print("Error getting selected sheets: {}".format(e))
        return None


def get_valid_floorplans_and_scale(sheets):
    """Validate FloorPlan and AreaPlan views and get uniform scale (or prompt user if mixed).
    
    Args:
        sheets: List of DB.ViewSheet elements
        
    Returns:
        tuple: (uniform_scale, valid_viewports_dict)
            - uniform_scale: float, the validated uniform scale
            - valid_viewports_dict: {sheet.Id: [list of valid DB.Viewport]}
        
    Raises:
        ValueError: If no valid views found
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

                # Only process FloorPlan or AreaPlan views
                if not view or (view.ViewType != DB.ViewType.FloorPlan and view.ViewType != DB.ViewType.AreaPlan):
                    continue
                
                # Must have rooms
                rooms = list(DB.FilteredElementCollector(doc, view.Id)
                            .OfCategory(DB.BuiltInCategory.OST_Rooms)
                            .WhereElementIsNotElementType()
                            .ToElements())
                if not rooms or len(rooms) == 0:
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
            error_msg = "No valid FloorPlan or AreaPlan views found.\n\n"
            error_msg += "Valid views must:\n"
            error_msg += "- Be FloorPlan or AreaPlan views\n"
            error_msg += "- Contain rooms\n"
            error_msg += "- Have a defined scale"
            raise ValueError(error_msg)
        
        # Check for uniform scale
        if len(scales_found) == 1:
            # All views have same scale - perfect!
            uniform_scale = list(scales_found.keys())[0]
            total_viewports = sum(len(v) for v in valid_viewports.values())
            print("Validation passed:")
            print("  - Uniform scale: 1:{}".format(int(uniform_scale)))
            print("  - Valid viewports: {}".format(total_viewports))
            return uniform_scale, valid_viewports
        
        # Multiple scales found - prompt user to choose
        print("\nMultiple scales detected in views:")
        for scale, locations in sorted(scales_found.items()):
            print("  Scale 1:{}: {} view(s)".format(int(scale), len(locations)))
        
        # Build prompt message
        prompt_msg = "Multiple scales detected in views!\n\n"
        prompt_msg += "Please select which scale to export:\n\n"
        
        scale_list = sorted(scales_found.keys())
        for i, scale in enumerate(scale_list, 1):
            locations = scales_found[scale]
            prompt_msg += "{}. Scale 1:{} ({} view(s))\n".format(i, int(scale), len(locations))
            for loc in locations[:3]:  # Show first 3 examples
                prompt_msg += "   - Sheet {} / {}\n".format(loc['sheet'], loc['view'])
            if len(locations) > 3:
                prompt_msg += "   ... and {} more\n".format(len(locations) - 3)
            prompt_msg += "\n"
        
        # Show dialog with numbered options
        from System.Windows.Forms import Form, Label, Button, FormBorderStyle, FormStartPosition
        from System.Drawing import Point, Size
        
        selected_scale = [None]  # Use list to capture in closure
        
        form = Form()
        form.Text = "Select Export Scale"
        form.Width = 500
        form.Height = 200 + len(scale_list) * 40
        form.FormBorderStyle = FormBorderStyle.FixedDialog
        form.StartPosition = FormStartPosition.CenterScreen
        form.MaximizeBox = False
        form.MinimizeBox = False
        
        label = Label()
        label.Text = "Multiple scales found. Select which scale to export:"
        label.Location = Point(20, 20)
        label.Size = Size(450, 40)
        form.Controls.Add(label)
        
        y_pos = 70
        for i, scale in enumerate(scale_list):
            btn = Button()
            btn.Text = "1:{} ({} views)".format(int(scale), len(scales_found[scale]))
            btn.Location = Point(20, y_pos)
            btn.Size = Size(450, 30)
            btn.Tag = scale
            def on_click(sender, args):
                selected_scale[0] = sender.Tag
                form.DialogResult = DialogResult.OK
                form.Close()
            btn.Click += on_click
            form.Controls.Add(btn)
            y_pos += 40
        
        # Cancel button
        cancel_btn = Button()
        cancel_btn.Text = "Cancel"
        cancel_btn.Location = Point(20, y_pos + 10)
        cancel_btn.Size = Size(450, 30)
        cancel_btn.DialogResult = DialogResult.Cancel
        form.Controls.Add(cancel_btn)
        form.CancelButton = cancel_btn
        
        result = form.ShowDialog()
        
        if result != DialogResult.OK or selected_scale[0] is None:
            raise ValueError("Export cancelled by user")
        
        chosen_scale = selected_scale[0]
        print("\nUser selected scale: 1:{}".format(int(chosen_scale)))
        
        # Filter viewports to only those with chosen scale
        filtered_viewports = {}
        for sheet_id, vp_list in valid_viewports.items():
            filtered_list = []
            for vp in vp_list:
                view = doc.GetElement(vp.ViewId)
                if view and hasattr(view, 'Scale') and float(view.Scale) == chosen_scale:
                    filtered_list.append(vp)
            if len(filtered_list) > 0:
                filtered_viewports[sheet_id] = filtered_list
        
        total_viewports = sum(len(v) for v in filtered_viewports.values())
        print("  - Filtered to {} viewports at scale 1:{}".format(total_viewports, int(chosen_scale)))
        
        return chosen_scale, filtered_viewports
        
    except ValueError:
        # Re-raise validation errors
        raise
    except Exception as e:
        error_msg = "Error during validation: {}".format(e)
        print(error_msg)
        import traceback
        traceback.print_exc()
        raise ValueError(error_msg)


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
            print("  {} - {}".format(sheet.SheetNumber, sheet.Name))
        
        return sorted_sheets
        
    except Exception as e:
        print("Error sorting sheets: {}".format(e))
        return sheets


# ============================================================================
# SECTION 8: MAIN ORCHESTRATION
# ============================================================================

if __name__ == '__main__':
    try:
        print("="*60)
        print("ExportDXFforMOT - Floor Plans to DXF Export for MOT")
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
        
        # 2. Load preferences
        print("\nLoading preferences...")
        preferences = load_preferences()
        export_folder = export_utils.get_export_folder_path(preferences["ExportFolder"])
        
        # Create export folder if it doesn't exist
        if not os.path.exists(export_folder):
            os.makedirs(export_folder)
            print("Created export folder: {}".format(export_folder))
        
        # 3. Sort sheets (descending - rightmost = page 1)
        sorted_sheets = sort_sheets_by_number(initial_sheets, descending=True)
        
        # 4. Validate FloorPlan/AreaPlan views + uniform scale (or user choice)
        print("\nValidating sheets and views...")
        try:
            view_scale, valid_viewports_map = get_valid_floorplans_and_scale(sorted_sheets)
        except ValueError as e:
            # Validation failed - show warning and exit
            print("\nValidation failed:")
            print(str(e))
            MessageBox.Show(
                str(e),
                "Validation Error",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning
            )
            sys.exit()
        
        # 5. Create DXF document
        print("\nCreating DXF document...")
        dxf_doc = ezdxf.new('R2010')  # AutoCAD 2010 format (widely compatible)
        dxf_doc.header['$INSUNITS'] = 5  # 5 = centimeters
        dxf_doc.header['$SORTENTS'] = 127  # Honor SORTENTSTABLE draw order for all ops
        dxf_doc.styles.add('Standard', font='Arial.ttf')
        msp = dxf_doc.modelspace()
        
        # Create fixed layers
        create_frame_layer(dxf_doc)
        create_merm_layer(dxf_doc)
        
        # 6. Pin every room to the lowest plan that shows it, so a room spanning
        #    several levels is exported once instead of once per plan
        print("")
        room_assignments = build_room_view_assignments(valid_viewports_map)
        
        # 7. Process each sheet with horizontal offset
        horizontal_offset = 0.0  # In Revit feet
        
        for sheet in sorted_sheets:
            # Get pre-validated viewports for this sheet
            valid_viewports = valid_viewports_map.get(sheet.Id, [])
            
            # Only process sheets with valid viewports
            if len(valid_viewports) > 0:
                sheet_width = process_sheet(
                    sheet, dxf_doc, msp, horizontal_offset, view_scale, valid_viewports,
                    room_assignments=room_assignments
                )
                
                # Update horizontal offset for next sheet (add sheet width in feet)
                horizontal_offset += sheet_width
        
        # Push underlays behind all exported geometry (all sheets are now placed)
        send_underlays_to_back(msp)
        
        # 8. Generate filename
        sheet_numbers = [s.SheetNumber for s in sorted_sheets]
        if len(sheet_numbers) == 1:
            sheets_part = export_utils.sanitize_filename_part(sheet_numbers[0])
        else:
            sheets_part = "{}..{}".format(
                export_utils.sanitize_filename_part(sheet_numbers[0]),
                export_utils.sanitize_filename_part(sheet_numbers[-1])
            )
        
        project_label = export_utils.sanitize_filename_part(
            export_utils._get_project_label(doc))
        filename = "{}-MOT-{}".format(project_label, sheets_part)
        
        dxf_path = os.path.join(export_folder, filename + ".dxf")
        
        # 9. Save DXF file (binary format)
        print("\nSaving DXF file...")
        dxf_doc.saveas(dxf_path, fmt='bin')
        print("DXF saved: {}".format(dxf_path))
        
        # 10. Report results
        print("\n" + "="*60)
        print("EXPORT COMPLETE")
        print("="*60)
        print("Output folder: {}".format(export_folder))
        print("\nExported file:")
        print("  - {} ({} sheets)".format(os.path.basename(dxf_path), len(sorted_sheets)))
        print("="*60)
        
        MessageBox.Show(
            "Export complete!\n\nFiles saved to:\n{}".format(export_folder),
            "Export Successful",
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
