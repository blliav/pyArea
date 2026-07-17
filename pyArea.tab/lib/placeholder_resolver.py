# -*- coding: utf-8 -*-
"""Centralized placeholder resolution for pyArea.

Resolves placeholder strings (e.g. "<View Name>", "<by Shared Coordinates>")
to their actual values from Revit elements.

Compatible with both IronPython 2.7 and CPython 3 pyRevit engines.

Used by:
- ExportDXF (CPython 3) - resolving values during DXF export
- CalculationSetup (IronPython 2.7) - showing resolved values in the UI
"""

from pyrevit import DB

# IronPython 2.7 / CPython 3 string type compatibility
try:
    string_types = basestring  # noqa: F821 (IronPython 2.7)
except NameError:
    string_types = str

# Coordinate conversion constant
FEET_TO_METERS = 0.3048  # Revit internal units (feet) to meters


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


def get_project_base_point(doc):
    """Get the Project Base Point element.

    Args:
        doc: Revit document

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


def get_shared_coordinates(doc, point):
    """Convert a point from Revit internal coordinates to shared coordinates.

    Generic function that transforms any point to shared coordinate system.
    Returns all three coordinates (X, Y, Z) in meters.

    Args:
        doc: Revit document
        point: DB.XYZ point in Revit internal (project) coordinates,
               where Z is measured from the internal origin.
               For levels, pass DB.XYZ(0, 0, level.ProjectElevation).

    Returns:
        tuple: (x_meters, y_meters, z_meters) or (None, None, None) if error

    Examples:
        # Internal origin
        x, y, z = get_shared_coordinates(doc, DB.XYZ(0, 0, 0))

        # Project Base Point
        pbp = get_project_base_point(doc)
        x, y, z = get_shared_coordinates(doc, pbp.Position)

        # Level (use ProjectElevation, not Elevation)
        x, y, z = get_shared_coordinates(doc, DB.XYZ(0, 0, level.ProjectElevation))
    """
    try:
        # Get active project location
        project_location = doc.ActiveProjectLocation
        if not project_location:
            print("Warning: No active project location found")
            return None, None, None

        # X/Y: GetProjectPosition handles rotation + offset correctly
        project_position = project_location.GetProjectPosition(point)
        x_meters = project_position.EastWest * FEET_TO_METERS
        y_meters = project_position.NorthSouth * FEET_TO_METERS

        # Z: GetProjectPosition.Elevation is unreliable for the shared Z offset.
        # Formula: shared_Z = (internal_Z - PBP_internal_Z + PBP_shared_Z) * FTM
        # This is equivalent to: internal_Z * FTM + InternalOrigin_shared_elevation
        pbp_elev_feet = 0.0
        pbp_internal_z_feet = 0.0
        try:
            pbp_points = DB.FilteredElementCollector(doc)\
                .OfCategory(DB.BuiltInCategory.OST_ProjectBasePoint)\
                .ToElements()
            if pbp_points and len(pbp_points) > 0:
                p = pbp_points[0].get_Parameter(DB.BuiltInParameter.BASEPOINT_ELEVATION_PARAM)
                if p and p.HasValue:
                    pbp_elev_feet = p.AsDouble()
                pbp_pos = pbp_points[0].Position
                if pbp_pos is not None:
                    pbp_internal_z_feet = pbp_pos.Z
        except Exception:
            pass
        z_meters = (point.Z - pbp_internal_z_feet + pbp_elev_feet) * FEET_TO_METERS

        return x_meters, y_meters, z_meters

    except Exception as e:
        print("Warning: Error converting point to shared coordinates: {}".format(e))
        return None, None, None


def is_placeholder(value):
    """Check whether a value is a placeholder string like "<...>".

    Args:
        value: Value to check

    Returns:
        bool: True if value is a placeholder string
    """
    if not value or not isinstance(value, string_types):
        return False
    return value.startswith("<") and value.endswith(">")


def resolve_placeholder(placeholder_value, element, doc, context=None):
    """Resolve a placeholder string to its actual value.

    Simple direct resolution - just pass the element and get the value.

    Args:
        placeholder_value: String that may be a placeholder (e.g., "<View Name>")
        element: Revit element to extract value from (View, Sheet, Area, etc.)
        doc: Revit document
        context: Optional dict with extra resolution context. Keys:
            "floor_elevations" - sorted list of (elevation_feet, level_id)
                                 for <by Floor Above>
            "auto_number"      - sequential area counter for <AutoNumber>

    Returns:
        str: Resolved value, or original value if not a placeholder
    """
    if context is None:
        context = {}

    if not placeholder_value or not isinstance(placeholder_value, string_types):
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
            # Compute level height above PBP using internal coordinates (elevation-base-independent)
            if hasattr(element, 'GenLevel'):
                level = element.GenLevel
                if level:
                    pbp = get_project_base_point(doc)
                    pbp_internal_z = pbp.Position.Z if pbp else 0.0
                    elevation_feet = level.ProjectElevation - pbp_internal_z
                    return format_meters(elevation_feet * FEET_TO_METERS)
            return ""

        elif placeholder_value == "<by Shared Coordinates>":
            # Get level elevation in shared coordinate system
            if hasattr(element, 'GenLevel'):
                level = element.GenLevel
                if level:
                    # Create a point at the level's elevation (in internal coordinates)
                    level_point = DB.XYZ(0, 0, level.ProjectElevation)
                    # Transform to shared coordinates and get Z component
                    _, _, z_meters = get_shared_coordinates(doc, level_point)
                    return format_meters(z_meters)
            return ""

        elif placeholder_value == "<by Floor Above>":
            # Height = distance to next floor above (in meters)
            # Uses only AreaPlan levels in current calculation (via context)
            if hasattr(element, 'GenLevel'):
                level = element.GenLevel
                if level:
                    floor_elevations = context.get("floor_elevations", [])
                    for elev, lid in floor_elevations:
                        if elev > level.ProjectElevation + 0.001:  # tolerance
                            diff_meters = (elev - level.ProjectElevation) * FEET_TO_METERS
                            return format_meters(diff_meters)
                    return "3.00"  # topmost floor default: 3m
            return ""

        elif placeholder_value == "<AutoNumber>":
            # Sequential number assigned per view (set via context)
            if "auto_number" in context:
                return str(context["auto_number"])
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
            x, _, _ = get_shared_coordinates(doc, DB.XYZ(0, 0, 0))
            return format_meters(x)

        elif placeholder_value == "<N/S@InternalOrigin>":
            _, y, _ = get_shared_coordinates(doc, DB.XYZ(0, 0, 0))
            return format_meters(y)

        elif placeholder_value == "<E/W@ProjectBasePoint>":
            pbp = get_project_base_point(doc)
            if pbp and hasattr(pbp, 'Position'):
                x, _, _ = get_shared_coordinates(doc, pbp.Position)
                return format_meters(x)
            return ""

        elif placeholder_value == "<N/S@ProjectBasePoint>":
            pbp = get_project_base_point(doc)
            if pbp and hasattr(pbp, 'Position'):
                _, y, _ = get_shared_coordinates(doc, pbp.Position)
                return format_meters(y)
            return ""

        elif placeholder_value == "<SharedElevation@ProjectBasePoint>":
            pbp = get_project_base_point(doc)
            if pbp and hasattr(pbp, 'Position'):
                _, _, z = get_shared_coordinates(doc, pbp.Position)
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


def build_floor_elevations_context(doc, view_ids):
    """Build the "floor_elevations" context list for <by Floor Above>.

    Collects unique level elevations from the given AreaPlan views,
    sorted ascending.

    Args:
        doc: Revit document
        view_ids: Iterable of view ElementIds (or int ids) of AreaPlan views

    Returns:
        list: Sorted list of (elevation_feet, level_id_int) tuples
    """
    floor_elevations = []
    seen_level_ids = set()
    for vid in view_ids:
        try:
            if isinstance(vid, DB.ElementId):
                view = doc.GetElement(vid)
            else:
                view = doc.GetElement(DB.ElementId(System_Int64(vid)))
            if view and hasattr(view, 'GenLevel') and view.GenLevel:
                lid = view.GenLevel.Id.Value
                if lid not in seen_level_ids:
                    seen_level_ids.add(lid)
                    floor_elevations.append((view.GenLevel.ProjectElevation, lid))
        except Exception:
            pass
    floor_elevations.sort()
    return floor_elevations


# Lazy Int64 conversion helper (only needed when ids are passed as ints)
def System_Int64(value):
    from System import Int64
    return Int64(int(value))
