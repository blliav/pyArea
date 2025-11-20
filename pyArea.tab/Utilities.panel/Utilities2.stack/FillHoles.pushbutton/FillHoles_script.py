# -*- coding: utf-8 -*-
"""Fill holes in area boundaries."""

__title__ = "Fill\nHoles"
__author__ = "Your Name"

from pyrevit import revit, DB, forms, script
import os
import sys
import time

SCRIPT_DIR = os.path.dirname(__file__)
LIB_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(SCRIPT_DIR))), "lib")

if LIB_DIR not in sys.path:
    sys.path.insert(0, LIB_DIR)

import data_manager

logger = script.get_logger()

def _get_active_areaplan_view():
    uidoc = revit.uidoc
    doc = revit.doc
    if uidoc is None or doc is None:
        forms.alert("No active document.", exitscript=True)
    view = uidoc.ActiveView
    try:
        if not isinstance(view, DB.ViewPlan) or view.ViewType != DB.ViewType.AreaPlan:
            forms.alert("Please run this tool from an Area Plan view.", exitscript=True)
    except Exception:
        forms.alert("Please run this tool from an Area Plan view.", exitscript=True)
    return doc, view


def _get_areas_in_view(doc, areaplan_view):
    areas = []
    try:
        area_scheme = getattr(areaplan_view, "AreaScheme", None)
        gen_level = getattr(areaplan_view, "GenLevel", None)
        if area_scheme is None or gen_level is None:
            return []
        target_scheme_id = area_scheme.Id
        target_level_id = gen_level.Id
        collector = DB.FilteredElementCollector(doc)
        collector = collector.OfCategory(DB.BuiltInCategory.OST_Areas)
        collector = collector.WhereElementIsNotElementType()
        for area in collector:
            try:
                if getattr(area, "AreaScheme", None) is None:
                    continue
                if area.AreaScheme.Id != target_scheme_id:
                    continue
                if area.LevelId != target_level_id:
                    continue
                areas.append(area)
            except Exception:
                continue
    except Exception:
        return []
    return areas


def _get_boundary_loops(area):
    try:
        options = DB.SpatialElementBoundaryOptions()
        loops = area.GetBoundarySegments(options)
        return list(loops) if loops else []
    except Exception:
        return []


def _point_inside_curveloop(point, curve_loop):
    if not point or not curve_loop or curve_loop.NumberOfCurves() == 0:
        return False

    polygon = []
    prev = None
    for curve in curve_loop:
        try:
            pt = curve.GetEndPoint(0)
            if prev is None or not pt.IsAlmostEqualTo(prev):
                polygon.append((pt.X, pt.Y))
                prev = pt
        except Exception:
            continue
    if len(polygon) < 3:
        return False

    x, y = point.X, point.Y
    inside = False
    n = len(polygon)
    p1x, p1y = polygon[0]
    for i in range(n + 1):
        p2x, p2y = polygon[i % n]
        if min(p1y, p2y) < y <= max(p1y, p2y) and x <= max(p1x, p2x):
            if p1y != p2y:
                xinters = (y - p1y) * (p2x - p1x) / (p2y - p1y) + p1x
            else:
                xinters = p1x
            if p1x == p2x or x <= xinters:
                inside = not inside
        p1x, p1y = p2x, p2y

    return inside


def _boundary_loop_to_curveloop(boundary_loop):
    """Convert a boundary segment loop to a CurveLoop."""
    try:
        curve_loop = DB.CurveLoop()
        for segment in boundary_loop:
            curve = segment.GetCurve()
            if curve:
                curve_loop.Append(curve)
        return curve_loop if curve_loop.NumberOfCurves() > 0 else None
    except Exception as e:
        print("  Warning: Failed to create CurveLoop - {}".format(e))
        return None


def _get_curveloop_centroid(curve_loop):
    """Get centroid of a CurveLoop using bounding box center.
    
    This is more reliable than averaging curve midpoints.
    """
    if not curve_loop or curve_loop.NumberOfCurves() == 0:
        return None
    
    min_x = float("inf")
    max_x = float("-inf")
    min_y = float("inf")
    max_y = float("-inf")
    has_points = False
    
    # Sample multiple points along each curve to build accurate bounding box
    for curve in curve_loop:
        try:
            # Sample at 0%, 25%, 50%, 75%, 100% along each curve
            for t in [0.0, 0.25, 0.5, 0.75, 1.0]:
                try:
                    p = curve.Evaluate(t, True)
                    if p.X < min_x:
                        min_x = p.X
                    if p.X > max_x:
                        max_x = p.X
                    if p.Y < min_y:
                        min_y = p.Y
                    if p.Y > max_y:
                        max_y = p.Y
                    has_points = True
                except Exception:
                    continue
        except Exception:
            continue
    
    if not has_points:
        return None
    
    center_x = (min_x + max_x) / 2.0
    center_y = (min_y + max_y) / 2.0
    
    return DB.XYZ(center_x, center_y, 0.0)


def _curveloop_to_solid(curve_loop, extrusion_height=1.0):
    """Convert a CurveLoop to a thin planar Solid by extruding it."""
    try:
        # Create a list containing just this one loop
        loops = [curve_loop]
        # Extrude the loop to create a solid
        solid = DB.GeometryCreationUtilities.CreateExtrusionGeometry(
            loops,
            DB.XYZ(0, 0, 1),  # Extrusion direction (up)
            extrusion_height
        )
        return solid
    except Exception as e:
        logger.debug("Failed to create solid from CurveLoop: %s", e)
        return None


def _solid_to_curveloops(solid):
    """Extract CurveLoops from a Solid's bottom face."""
    try:
        loops = []
        min_z = float('inf')
        bottom_face = None
        
        # Find the face with the lowest Z coordinate (bottom face)
        for face in solid.Faces:
            try:
                # Get a point on the face to check its Z coordinate
                mesh = face.Triangulate()
                if mesh and mesh.Vertices.Count > 0:
                    z = mesh.Vertices[0].Z
                    if z < min_z:
                        min_z = z
                        bottom_face = face
            except Exception:
                continue
        
        # Extract CurveLoops from the bottom face
        if bottom_face:
            edge_loops = bottom_face.GetEdgesAsCurveLoops()
            for loop in edge_loops:
                loops.append(loop)
        
        return loops
    except Exception as e:
        logger.debug("Failed to extract CurveLoops from Solid: %s", e)
        return []


def _subtract_areas_from_hole(hole_curveloop, other_area_curveloops):
    """Subtract other areas' outer boundaries from a hole to find empty regions.
    
    Returns: List of CurveLoops representing the empty parts of the hole
    """
    try:
        # Convert hole to solid
        hole_solid = _curveloop_to_solid(hole_curveloop)
        if not hole_solid:
            return [hole_curveloop]
        
        result_solid = hole_solid
        
        # Subtract each overlapping area
        for other_area_id, other_loop in other_area_curveloops.items():
            try:
                # Convert other area to solid
                other_solid = _curveloop_to_solid(other_loop)
                if not other_solid:
                    continue
                
                # Perform boolean subtraction
                result_solid = DB.BooleanOperationsUtils.ExecuteBooleanOperation(
                    result_solid,
                    other_solid,
                    DB.BooleanOperationsType.Difference
                )
                logger.debug("Boolean subtraction succeeded for Area %s", other_area_id)
                
            except Exception as e:
                logger.debug("Boolean operation failed for Area %s: %s", other_area_id, e)
                continue
        
        # Convert result solid back to CurveLoops
        result_loops = _solid_to_curveloops(result_solid)
        if not result_loops:
            return [hole_curveloop]
        return result_loops
        
    except Exception as e:
        logger.debug("Boolean subtraction failed: %s", e)
        return [hole_curveloop]


def _get_curveloop_area(curve_loop):
    """Calculate approximate area of a CurveLoop."""
    try:
        # Use bounding box as approximation
        min_x = float("inf")
        max_x = float("-inf")
        min_y = float("inf")
        max_y = float("-inf")
        
        for curve in curve_loop:
            for t in [0.0, 0.5, 1.0]:
                try:
                    p = curve.Evaluate(t, True)
                    min_x = min(min_x, p.X)
                    max_x = max(max_x, p.X)
                    min_y = min(min_y, p.Y)
                    max_y = max(max_y, p.Y)
                except Exception:
                    continue
        
        # Area in square feet
        return (max_x - min_x) * (max_y - min_y)
    except Exception:
        return 0.0


def _find_empty_hole_regions(areas, min_area_sqft=0.5):
    """Find empty regions in holes using 2D boolean operations.
    
    For each hole, subtract all overlapping Areas to find the actual empty space.
    Filters out regions smaller than min_area_sqft (default 0.5 sq ft = ~465 sq cm).
    
    Returns: List of (area_id, hole_idx, empty_region_centroids)
    """
    holes_data = []
    
    # Build map of all areas' outer boundaries as CurveLoops
    area_outer_curveloops = {}
    area_locations = {}
    for area in areas:
        loops = _get_boundary_loops(area)
        if loops and len(loops) > 0:
            outer_curveloop = _boundary_loop_to_curveloop(loops[0])
            if outer_curveloop:
                area_outer_curveloops[area.Id] = outer_curveloop
        try:
            location = getattr(area, "Location", None)
            if location and hasattr(location, "Point"):
                area_locations[area.Id] = location.Point
        except Exception:
            continue

    # Process each area's holes
    for area in areas:
        loops = _get_boundary_loops(area)
        if not loops or len(loops) <= 1:
            continue
        
        # Process each interior loop (hole)
        for loop_idx, loop in enumerate(loops[1:]):
            hole_curveloop = _boundary_loop_to_curveloop(loop)
            if not hole_curveloop:
                continue
            
            # Collect other areas' outer boundaries that overlap with this hole
            # Check BOTH location points AND boundary points to detect overlap
            other_curveloops = {}
            for aid, cl in area_outer_curveloops.items():
                if aid == area.Id:
                    continue
                
                # First check: location point inside hole
                loc_point = area_locations.get(aid)
                if loc_point and _point_inside_curveloop(loc_point, hole_curveloop):
                    other_curveloops[aid] = cl
                    logger.debug("Area %s location is inside hole %s", aid, loop_idx + 1)
                    continue
                
                # Second check: any boundary point inside hole (catches areas extending into hole)
                overlaps = False
                for curve in cl:
                    try:
                        # Check start, middle, and end points of each curve
                        for t in [0.0, 0.25, 0.5, 0.75, 1.0]:
                            pt = curve.Evaluate(t, True)
                            if _point_inside_curveloop(pt, hole_curveloop):
                                overlaps = True
                                logger.debug("Area %s boundary overlaps hole %s", aid, loop_idx + 1)
                                break
                        if overlaps:
                            break
                    except Exception:
                        continue
                
                if overlaps:
                    other_curveloops[aid] = cl
            
            # If no areas overlap, this hole is completely empty - add its centroid
            if not other_curveloops:
                centroid = _get_curveloop_centroid(hole_curveloop)
                if centroid:
                    logger.debug("Hole %s of Area %s: No overlapping areas, completely empty", loop_idx + 1, area.Id)
                    holes_data.append((area.Id, loop_idx + 1, [centroid]))
                else:
                    logger.debug("Hole %s of Area %s: No centroid calculated", loop_idx + 1, area.Id)
                continue
            else:
                logger.debug("Hole %s of Area %s: Found %s overlapping areas", loop_idx + 1, area.Id, len(other_curveloops))
            
            # Otherwise, subtract overlapping areas to find empty regions
            empty_regions = _subtract_areas_from_hole(hole_curveloop, other_curveloops)
            
            # Get centroid of each empty region (filter out tiny regions)
            centroids = []
            for region in empty_regions:
                region_area = _get_curveloop_area(region)
                if region_area >= min_area_sqft:
                    centroid = _get_curveloop_centroid(region)
                    if centroid:
                        centroids.append(centroid)

            if centroids:
                holes_data.append((area.Id, loop_idx + 1, centroids))

    return holes_data


def _get_hole_usage_for_municipality(doc, view):
    municipality, variant = data_manager.get_municipality_from_view(doc, view)
    usage_type_value = None
    usage_type_name = None

    logger.debug("Detected municipality='{}', variant='{}'".format(municipality, variant))
    print("FillHoles: Municipality = {} | Variant = {}".format(municipality, variant))
    if municipality == "Jerusalem":
        usage_type_value = "70"
        usage_type_name = u"הורדה"
    elif municipality == "Common":
        if variant == "Gross":
            usage_type_value = "700"
        else:
            usage_type_value = "300"
        usage_type_name = u"הורדה"
    elif municipality == "Tel-Aviv":
        usage_type_value = "-1"
        usage_type_name = u"חלל"
    return usage_type_value, usage_type_name


def fill_holes():
    """
    Fill holes in area boundaries.
    
    This tool will:
    1. Find all area boundaries with holes
    2. Create filled regions to cover the holes
    3. Optionally create new area elements in the holes
    """
    # TODO: Implement hole filling logic
    overall_start = time.time()
    doc, view = _get_active_areaplan_view()
    area_scheme = getattr(view, "AreaScheme", None)
    if area_scheme is None:
        forms.alert("Active Area Plan has no Area Scheme.", exitscript=True)

    perf_data = []

    areas_start = time.time()
    areas = _get_areas_in_view(doc, view)
    if not areas:
        forms.alert("No areas found in the active Area Plan view.", exitscript=True)
    perf_data.append(("Collect Areas", time.time() - areas_start))

    holes_start = time.time()
    holes_data = _find_empty_hole_regions(areas)
    if not holes_data:
        forms.alert("No holes detected in area boundaries in this view.", exitscript=True)
    perf_data.append(("Detect Empty Holes", time.time() - holes_start))
    
    # Debug output: show what was detected
    print("\nHole Detection Results:")
    for area_id, hole_idx, centroids in holes_data:
        area_id_val = area_id.IntegerValue if hasattr(area_id, 'IntegerValue') else int(area_id.Value)
        print("  Area {}: Hole {} has {} empty region(s)".format(area_id_val, hole_idx, len(centroids)))
        for i, c in enumerate(centroids):
            print("    Region {}: centroid at ({:.2f}, {:.2f})".format(i+1, c.X, c.Y))

    usage_type_value, usage_type_name = _get_hole_usage_for_municipality(doc, view)
    logger.debug(
        "Using usage_type_value='%s', usage_type_name='%s' for new areas",
        usage_type_value,
        usage_type_name,
    )

    created_count = 0
    failed_count = 0
    skipped_count = 0
    created_area_ids = []  # Store created area integer IDs for linkify output
    
    # Get output for linkify
    output = script.get_output()
    
    # Create set of existing area IDs for fast lookup
    existing_area_ids = set(a.Id.IntegerValue if hasattr(a.Id, 'IntegerValue') else int(a.Id.Value) for a in areas)
    
    print("\nProcessing {} holes with {} empty regions to fill...".format(
        len(holes_data),
        sum(len(centroids) for _, _, centroids in holes_data)
    ))
    
    creation_start = time.time()
    
    with revit.Transaction("Fill Area Holes"):
        for area_id, hole_idx, centroids in holes_data:
            # Create an Area for each empty region
            for region_idx, centroid in enumerate(centroids):
                try:
                    uv = DB.UV(centroid.X, centroid.Y)
                    
                    # Create Area using Revit API
                    new_area = doc.Create.NewArea(view, uv)

                    if new_area is None:
                        failed_count += 1
                        continue
                    
                    # Check if area was actually created (not overlapping existing area)
                    # If area is placed on top of existing area, Revit returns the existing area
                    new_area_id_value = new_area.Id.IntegerValue if hasattr(new_area.Id, 'IntegerValue') else int(new_area.Id.Value)
                    
                    # Check if this is an existing area (Revit returned existing area instead of creating new)
                    if new_area_id_value in existing_area_ids:
                        logger.debug("Area %s already exists, skipping", new_area_id_value)
                        print("  Skipped existing Area {} (hole {} region {})".format(new_area_id_value, hole_idx, region_idx + 1))
                        skipped_count += 1
                        continue
                    
                    # Check if area has zero size (creation failed - likely overlapping)
                    area_size = new_area.Area
                    if area_size == 0:
                        logger.debug("Area %s has zero size at centroid (%.3f, %.3f), deleting to prevent warning", 
                                   new_area_id_value, centroid.X, centroid.Y)
                        # Delete the zero-size area to prevent "Multiple Areas in same region" warning
                        try:
                            doc.Delete(new_area.Id)
                        except Exception as del_ex:
                            logger.debug("Could not delete zero-size area: %s", del_ex)
                        print("  Skipped zero-size area at ({:.2f}, {:.2f}) - overlaps existing (hole {} region {})".format(
                            centroid.X, centroid.Y, hole_idx, region_idx + 1))
                        skipped_count += 1
                        continue

                    # Successfully created - now set parameters
                    if usage_type_value:
                        try:
                            param = new_area.LookupParameter("Usage Type")
                            if param and not param.IsReadOnly:
                                param.Set(usage_type_value)

                            if usage_type_name:
                                name_param = new_area.LookupParameter("Name")
                                if name_param and not name_param.IsReadOnly:
                                    name_param.Set(usage_type_name)

                                usage_type_name_param = new_area.LookupParameter("Usage Type Name")
                                if usage_type_name_param and not usage_type_name_param.IsReadOnly:
                                    usage_type_name_param.Set(usage_type_name)
                        except Exception as param_ex:
                            logger.debug("Failed to set parameters on Area %s: %s", new_area.Id, param_ex)

                    created_count += 1
                    
                    # Store created area integer ID for linkify output after transaction
                    created_area_ids.append(new_area_id_value)
                    logger.debug("Successfully created Area %s", new_area_id_value)
                    print("  Created Area {} (hole {} region {})".format(new_area_id_value, hole_idx, region_idx + 1))
                    
                except Exception as create_ex:
                    logger.debug("Failed to create Area for hole %s region %s: %s", hole_idx, region_idx + 1, create_ex)
                    failed_count += 1
                    continue
        
        commit_start = time.time()
        perf_data.append(("Create Areas (inside txn)", commit_start - creation_start))
    
    # Transaction commits when exiting the 'with' block above
    perf_data.append(("Transaction Commit", time.time() - commit_start))

    # Output linkify for created areas AFTER transaction commits
    if created_area_ids:
        print("\n" + "="*50)
        print("Created Areas (click to select):")
        print("="*50)
        for area_id_value in created_area_ids:
            # Linkify requires ElementId object, reconstruct from integer value
            try:
                # Create ElementId from integer value
                elem_id = data_manager.create_element_id(area_id_value)
                output.linkify(elem_id, title="Area {}".format(area_id_value))
                print(" in view '{}'".format(view.Name))
            except Exception as e:
                logger.debug("Failed to linkify area %s: %s", area_id_value, e)
                print("Area {} in view '{}' (link failed)".format(area_id_value, view.Name))

    # Simple fast report using print instead of markdown rendering
    print("\n" + "="*50)
    print("FillHoles Report")
    print("="*50)
    print("Areas scanned: {}".format(len(areas)))
    print("Holes detected: {}".format(len(holes_data)))
    print("New areas created: {}".format(created_count))
    if skipped_count > 0:
        print("Skipped (overlapping): {}".format(skipped_count))
    if failed_count > 0:
        print("Failed: {}".format(failed_count))
    if created_count > 0:
        print("\n✓ Successfully filled {} hole(s)".format(created_count))
    
    total_time = time.time() - overall_start
    perf_data.append(("Total Runtime", total_time))
    
    print("\nTiming: " + ", ".join(
        "{}: {:.2f}s".format(label, duration) for label, duration in perf_data
    ))
    print("="*50)


if __name__ == '__main__':
    fill_holes()
