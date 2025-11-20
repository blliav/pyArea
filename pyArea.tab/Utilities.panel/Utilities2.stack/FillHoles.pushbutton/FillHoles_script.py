# -*- coding: utf-8 -*-
"""Fill holes in area boundaries using boolean union approach."""

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


# ============================================================
# UTILITY FUNCTIONS
# ============================================================

def _get_element_id_value(element_id):
    """Get integer value from ElementId - compatible with Revit 2024, 2025 and 2026+"""
    try:
        return element_id.IntegerValue
    except AttributeError:
        return int(element_id.Value)


def _get_active_areaplan_view():
    """Get the active Area Plan view and document."""
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
    """Get all PLACED areas in the specified Area Plan view."""
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
                
                # Filter out unplaced areas (no boundaries)
                loops = _get_boundary_loops(area)
                if not loops or len(loops) == 0:
                    continue
                
                areas.append(area)
            except Exception:
                continue
    except Exception:
        return []
    
    return areas


def _get_boundary_loops(area):
    """Get boundary loops from an area element."""
    try:
        options = DB.SpatialElementBoundaryOptions()
        loops = area.GetBoundarySegments(options)
        return list(loops) if loops else []
    except Exception:
        return []


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
        logger.debug("Failed to create CurveLoop: %s", e)
        return None


def _get_curveloop_centroid(curve_loop):
    """Get centroid of a CurveLoop using bounding box center."""
    if not curve_loop or curve_loop.NumberOfCurves() == 0:
        return None
    
    min_x = float("inf")
    max_x = float("-inf")
    min_y = float("inf")
    max_y = float("-inf")
    has_points = False
    
    # Sample multiple points along each curve
    for curve in curve_loop:
        try:
            for t in [0.0, 0.25, 0.5, 0.75, 1.0]:
                try:
                    p = curve.Evaluate(t, True)
                    min_x = min(min_x, p.X)
                    max_x = max(max_x, p.X)
                    min_y = min(min_y, p.Y)
                    max_y = max(max_y, p.Y)
                    has_points = True
                except Exception:
                    continue
        except Exception:
            continue
    
    if not has_points:
        return None
    
    return DB.XYZ((min_x + max_x) / 2.0, (min_y + max_y) / 2.0, 0.0)


def _get_curveloop_bbox_area(curve_loop):
    """Approximate area of a CurveLoop using its 2D bounding box.
    Used only to classify exterior vs hole loops on the bottom face.
    """
    if not curve_loop or curve_loop.NumberOfCurves() == 0:
        return 0.0
    try:
        min_x = float("inf")
        max_x = float("-inf")
        min_y = float("inf")
        max_y = float("-inf")
        for curve in curve_loop:
            try:
                for t in [0.0, 0.5, 1.0]:
                    p = curve.Evaluate(t, True)
                    min_x = min(min_x, p.X)
                    max_x = max(max_x, p.X)
                    min_y = min(min_y, p.Y)
                    max_y = max(max_y, p.Y)
            except Exception:
                continue
        if min_x == float("inf") or max_x == float("-inf"):
            return 0.0
        return max(0.0, (max_x - min_x) * (max_y - min_y))
    except Exception:
        return 0.0


def _generate_grid_points(curve_loop, grid_size=5):
    """Generate a grid of test points within the bounding box of a CurveLoop.
    
    Args:
        curve_loop: The CurveLoop to generate points for
        grid_size: Number of points per dimension (default 5x5 = 25 points)
    
    Returns:
        List of (x, y) tuples
    """
    if not curve_loop or curve_loop.NumberOfCurves() == 0:
        return []
    
    try:
        # Get bounding box
        min_x = float("inf")
        max_x = float("-inf")
        min_y = float("inf")
        max_y = float("-inf")
        
        for curve in curve_loop:
            try:
                for t in [0.0, 0.25, 0.5, 0.75, 1.0]:
                    p = curve.Evaluate(t, True)
                    min_x = min(min_x, p.X)
                    max_x = max(max_x, p.X)
                    min_y = min(min_y, p.Y)
                    max_y = max(max_y, p.Y)
            except Exception:
                continue
        
        if min_x == float("inf") or max_x == float("-inf"):
            return []
        
        # Generate grid with margin (avoid edges)
        margin = 0.1  # feet
        min_x += margin
        max_x -= margin
        min_y += margin
        max_y -= margin
        
        if max_x <= min_x or max_y <= min_y:
            return []
        
        points = []
        for i in range(grid_size):
            for j in range(grid_size):
                x = min_x + (max_x - min_x) * i / (grid_size - 1)
                y = min_y + (max_y - min_y) * j / (grid_size - 1)
                points.append((x, y))
        
        return points
    except Exception as e:
        logger.debug("Failed to generate grid points: %s", e)
        return []


def _is_point_inside_area(point_x, point_y, area_elem):
    """Check if a point (x, y) is inside an area's boundary.
    
    Args:
        point_x, point_y: Point coordinates in feet
        area_elem: Area element to check
    
    Returns:
        True if point is inside area boundary, False otherwise
    """
    try:
        loops = _get_boundary_loops(area_elem)
        if not loops or len(loops) == 0:
            return False
        
        # Get the exterior boundary (first loop)
        exterior_loop = loops[0]
        curve_loop = _boundary_loop_to_curveloop(exterior_loop)
        
        if not curve_loop:
            return False
        
        # Use Revit's IsInside method on the planar loop
        test_point = DB.XYZ(point_x, point_y, 0)
        
        # Try to determine if point is inside using ray casting
        # Count intersections with boundary curves
        # Odd = inside, Even = outside
        intersections = 0
        ray_end = DB.XYZ(point_x + 1000.0, point_y, 0)  # Long horizontal ray
        
        for curve in curve_loop:
            try:
                # Create intersection result
                result = curve.Project(test_point)
                if result and result.Distance < 0.01:  # Point is ON the boundary
                    return False  # Consider boundary points as outside
                
                # Simple ray casting - count crossings
                # This is approximate but works for most cases
                p0 = curve.GetEndPoint(0)
                p1 = curve.GetEndPoint(1)
                
                # Check if ray crosses this curve segment
                if ((p0.Y <= point_y < p1.Y) or (p1.Y <= point_y < p0.Y)):
                    # Calculate x coordinate of intersection
                    if abs(p1.Y - p0.Y) > 1e-6:
                        x_intersect = p0.X + (point_y - p0.Y) * (p1.X - p0.X) / (p1.Y - p0.Y)
                        if x_intersect > point_x:
                            intersections += 1
            except Exception:
                continue
        
        # Odd number of intersections = inside
        return (intersections % 2) == 1
    except Exception as e:
        logger.debug("Failed to check point inside area: %s", e)
        return False

def _get_hole_usage_for_municipality(doc, view):
    """Get the appropriate usage type for holes based on municipality."""
    municipality, variant = data_manager.get_municipality_from_view(doc, view)
    usage_type_value = None
    usage_type_name = None

    logger.debug("Detected municipality='%s', variant='%s'", municipality, variant)
    print("Municipality: {} | Variant: {}".format(municipality, variant))
    
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


# ============================================================
# BOOLEAN UNION APPROACH
# ============================================================

def _create_union_of_areas(areas):
    """Create a boolean union of all area boundaries.
    
    Returns: A Solid representing the union, or None if failed
    """
    if not areas:
        return None
    
    print("\n" + "="*50)
    print("STEP 1: Creating Boolean Union")
    print("="*50)
    
    # Convert all area boundaries to solids
    solids = []
    for area in areas:
        try:
            area_id = _get_element_id_value(area.Id)
            loops = _get_boundary_loops(area)
            
            if not loops or len(loops) == 0:
                continue
            
            # Convert ALL loops (outer + holes) to CurveLoops
            curve_loops = []
            for loop in loops:
                curve_loop = _boundary_loop_to_curveloop(loop)
                if curve_loop:
                    curve_loops.append(curve_loop)
            
            if not curve_loops:
                continue
            
            # Extrude to create solid (1 foot height)
            # First loop is outer, rest are holes - Revit will create solid with holes
            solid = DB.GeometryCreationUtilities.CreateExtrusionGeometry(
                curve_loops,  # Include outer + holes
                DB.XYZ(0, 0, 1),  # Extrusion direction (up)
                1.0  # Height in feet
            )
            if solid:
                solids.append(solid)
                logger.debug("Converted Area %s to solid", area_id)
        except Exception as e:
            logger.debug("Failed to convert Area %s to solid: %s", _get_element_id_value(area.Id), e)
            continue
    
    if not solids:
        print("ERROR: No valid area boundaries found")
        return None
    
    print("Converted {} area(s) to solids".format(len(solids)))
    
    # Union all solids
    try:
        union_solid = solids[0]
        for i, solid in enumerate(solids[1:], 1):
            try:
                union_solid = DB.BooleanOperationsUtils.ExecuteBooleanOperation(
                    union_solid,
                    solid,
                    DB.BooleanOperationsType.Union
                )
                logger.debug("Unioned solid %d/%d", i, len(solids) - 1)
            except Exception as e:
                logger.debug("Boolean union failed for solid %d: %s", i, e)
                # Continue with partial union
                continue
        
        print("Boolean union completed")
        return union_solid
    except Exception as e:
        logger.debug("Failed to create union: %s", e)
        print("ERROR: Boolean union failed - {}".format(e))
        return None


def _extract_holes_from_union(union_solid):
    """Extract interior holes from the union solid.
    
    Returns: List of CurveLoop objects representing holes
    """
    if not union_solid:
        return []
    
    print("\n" + "="*50)
    print("STEP 2: Extracting Holes from Union")
    print("="*50)
    
    holes = []
    
    try:
        # Find the bottom face (lowest Z)
        min_z = float('inf')
        bottom_face = None
        
        for face in union_solid.Faces:
            try:
                mesh = face.Triangulate()
                if mesh and mesh.Vertices.Count > 0:
                    z = mesh.Vertices[0].Z
                    if z < min_z:
                        min_z = z
                        bottom_face = face
            except Exception:
                continue
        
        if not bottom_face:
            print("ERROR: No bottom face found in union")
            return []
        
        # Get curve loops from the bottom face
        edge_loops = bottom_face.GetEdgesAsCurveLoops()

        # Classify loops by approximate area: largest = exterior, others = holes
        loops = [loop for loop in edge_loops]
        if len(loops) <= 1:
            print("No holes detected")
            return []

        loop_areas = [_get_curveloop_bbox_area(loop) for loop in loops]
        if not any(a > 0.0 for a in loop_areas):
            print("No holes detected")
            return []

        outer_index = max(range(len(loops)), key=lambda i: loop_areas[i])

        for idx, loop in enumerate(loops):
            if idx == outer_index:
                continue
            holes.append(loop)

        if holes:
            print("Found {} hole(s)".format(len(holes)))
        else:
            print("No holes detected")

        return list(holes)
    except Exception as e:
        logger.debug("Failed to extract holes: %s", e)
        print("ERROR: Failed to extract holes - {}".format(e))
        return []


def _create_areas_in_holes(doc, view, holes, usage_type_value, usage_type_name):
    """Create new Area elements to fill holes using optimized multi-point strategy.
    
    For each hole:
      1. Try centroid first (fast path)
      2. If hole not fully filled, try grid points
      3. Skip points inside already-created areas
      4. Continue until no new areas created
    
    Returns: (created_count, failed_count, created_area_ids)
    """
    if not holes:
        return 0, 0, []
    
    print("\n" + "="*50)
    print("STEP 3: Filling Holes (Multi-Point Strategy)")
    print("="*50)
    
    total_created = 0
    total_failed = 0
    all_created_area_ids = []  # store all created ElementId objects
    
    # Use a single main transaction with subtransactions for each hole
    with revit.Transaction("Fill Area Holes"):
        for hole_idx, hole in enumerate(holes, 1):
            print("\nProcessing Hole {}/{}:".format(hole_idx, len(holes)))
            
            hole_created_count = 0
            hole_created_areas = []  # Track areas created for THIS hole
            
            # PHASE 1: Try centroid first (fast path)
            centroid = _get_curveloop_centroid(hole)
            if centroid:
                logger.debug("Hole %d: Trying centroid at (%.3f, %.3f)", hole_idx, centroid.X, centroid.Y)
                
                sub_txn = DB.SubTransaction(doc)
                try:
                    sub_txn.Start()
                    
                    uv = DB.UV(centroid.X, centroid.Y)
                    temp_area = doc.Create.NewArea(view, uv)
                    
                    if temp_area:
                        try:
                            doc.Regenerate()
                        except Exception:
                            pass
                        
                        if temp_area.Area > 0:
                            # Set parameters
                            _set_area_parameters(temp_area, usage_type_value, usage_type_name)
                            
                            sub_txn.Commit()
                            hole_created_areas.append(temp_area)
                            hole_created_count += 1
                            all_created_area_ids.append(temp_area.Id)
                            logger.debug("Hole %d: Centroid SUCCESS - Area %s created", hole_idx, _get_element_id_value(temp_area.Id))
                            print("  Centroid: Created area {} ({:.2f} sqm)".format(
                                _get_element_id_value(temp_area.Id),
                                temp_area.Area * 0.09290304
                            ))
                        else:
                            doc.Delete(temp_area.Id)
                            sub_txn.RollBack()
                    else:
                        sub_txn.RollBack()
                except Exception as e:
                    logger.debug("Hole %d: Centroid failed - %s", hole_idx, e)
                    try:
                        if sub_txn.HasStarted() and not sub_txn.HasEnded():
                            sub_txn.RollBack()
                    except Exception:
                        pass
            
            # PHASE 2: Try grid points (handles subdivided holes)
            # Generate grid of test points
            grid_points = _generate_grid_points(hole, grid_size=5)
            
            if grid_points:
                logger.debug("Hole %d: Generated %d grid points", hole_idx, len(grid_points))
                
                consecutive_failures = 0
                max_consecutive_failures = 10  # Stop after 10 consecutive failures
                
                for point_idx, (test_x, test_y) in enumerate(grid_points):
                    # Check if we should stop (too many consecutive failures)
                    if consecutive_failures >= max_consecutive_failures:
                        logger.debug("Hole %d: Stopping after %d consecutive failures", hole_idx, consecutive_failures)
                        break
                    
                    # Skip if point is inside any already-created area for this hole
                    skip = False
                    for existing_area in hole_created_areas:
                        if _is_point_inside_area(test_x, test_y, existing_area):
                            logger.debug("Hole %d Point %d: Skipping - inside existing area %s", 
                                       hole_idx, point_idx, _get_element_id_value(existing_area.Id))
                            skip = True
                            break
                    
                    if skip:
                        continue
                    
                    # Try creating area at this point
                    sub_txn = DB.SubTransaction(doc)
                    try:
                        sub_txn.Start()
                        
                        uv = DB.UV(test_x, test_y)
                        temp_area = doc.Create.NewArea(view, uv)
                        
                        if temp_area:
                            try:
                                doc.Regenerate()
                            except Exception:
                                pass
                            
                            if temp_area.Area > 0:
                                # Check if this is a duplicate of existing area
                                is_duplicate = False
                                for existing_area in hole_created_areas:
                                    if _get_element_id_value(existing_area.Id) == _get_element_id_value(temp_area.Id):
                                        is_duplicate = True
                                        break
                                
                                if is_duplicate:
                                    doc.Delete(temp_area.Id)
                                    sub_txn.RollBack()
                                    consecutive_failures += 1
                                else:
                                    # New area created!
                                    _set_area_parameters(temp_area, usage_type_value, usage_type_name)
                                    
                                    sub_txn.Commit()
                                    hole_created_areas.append(temp_area)
                                    hole_created_count += 1
                                    all_created_area_ids.append(temp_area.Id)
                                    consecutive_failures = 0  # Reset counter
                                    
                                    logger.debug("Hole %d Point %d: SUCCESS - Area %s created", 
                                               hole_idx, point_idx, _get_element_id_value(temp_area.Id))
                                    print("  Grid point {}: Created area {} ({:.2f} sqm)".format(
                                        point_idx,
                                        _get_element_id_value(temp_area.Id),
                                        temp_area.Area * 0.09290304
                                    ))
                            else:
                                doc.Delete(temp_area.Id)
                                sub_txn.RollBack()
                                consecutive_failures += 1
                        else:
                            sub_txn.RollBack()
                            consecutive_failures += 1
                    except Exception as e:
                        logger.debug("Hole %d Point %d failed: %s", hole_idx, point_idx, e)
                        try:
                            if sub_txn.HasStarted() and not sub_txn.HasEnded():
                                sub_txn.RollBack()
                        except Exception:
                            pass
                        consecutive_failures += 1
            
            # Summary for this hole
            if hole_created_count > 0:
                print("  Total: {} area(s) created for this hole".format(hole_created_count))
                total_created += hole_created_count
            else:
                print("  FAILED: No areas created")
                total_failed += 1
    
    return total_created, total_failed, all_created_area_ids


def _set_area_parameters(area_elem, usage_type_value, usage_type_name):
    """Set usage type parameters on an area element."""
    if not usage_type_value:
        return
    
    try:
        param = area_elem.LookupParameter("Usage Type")
        if param and not param.IsReadOnly:
            param.Set(usage_type_value)
        
        if usage_type_name:
            name_param = area_elem.LookupParameter("Name")
            if name_param and not name_param.IsReadOnly:
                name_param.Set(usage_type_name)
            
            usage_type_name_param = area_elem.LookupParameter("Usage Type Name")
            if usage_type_name_param and not usage_type_name_param.IsReadOnly:
                usage_type_name_param.Set(usage_type_name)
    except Exception as e:
        logger.debug("Failed to set parameters on Area %s: %s", _get_element_id_value(area_elem.Id), e)


# ============================================================
# MAIN FUNCTION
# ============================================================

def fill_holes():
    """
    Fill holes in area boundaries using boolean union approach.
    
    This tool will:
    1. Create a boolean union of all area boundaries
    2. Extract holes from the unified region
    3. Create new areas at the centroid of each hole
    """
    overall_start = time.time()
    
    # Get active view
    doc, view = _get_active_areaplan_view()
    area_scheme = getattr(view, "AreaScheme", None)
    if area_scheme is None:
        forms.alert("Active Area Plan has no Area Scheme.", exitscript=True)
    
    output = script.get_output()
    perf_data = []
    print("Active View: '{}'".format(view.Name))
    print("Area Scheme: '{}'".format(area_scheme.Name))
    
    # Get areas
    areas_start = time.time()
    areas = _get_areas_in_view(doc, view)
    if not areas:
        forms.alert("No areas found in the active Area Plan view.", exitscript=True)
    print("Found {} placed area(s) in view".format(len(areas)))
    
    
    perf_data.append(("Collect Areas", time.time() - areas_start))
    
    # Get municipality settings
    usage_type_value, usage_type_name = _get_hole_usage_for_municipality(doc, view)
    
    # Create union
    union_start = time.time()
    union_solid = _create_union_of_areas(areas)
    if not union_solid:
        forms.alert("Failed to create boolean union of areas.", exitscript=True)
    perf_data.append(("Create Union", time.time() - union_start))
    
    # Extract holes
    holes_start = time.time()
    holes = _extract_holes_from_union(union_solid)
    if not holes:
        forms.alert("No holes detected in the unified area region.", exitscript=True)
    perf_data.append(("Extract Holes", time.time() - holes_start))
    
    # Create areas in holes
    create_start = time.time()
    created_count, failed_count, created_area_ids = _create_areas_in_holes(
        doc, view, holes, usage_type_value, usage_type_name
    )
    perf_data.append(("Create Areas", time.time() - create_start))
    
    # Output linkify for created areas (use stored ElementIds directly)
    if created_area_ids:
        print("\n" + "="*50)
        print("Created Areas (click to select)")
        print("="*50)
        for elem_id in created_area_ids:
            try:
                link = output.linkify(elem_id)
                area_elem = doc.GetElement(elem_id)
                
                # Get area in square meters
                area_sqft = area_elem.Area
                area_sqm = area_sqft * 0.09290304  # Convert sq ft to sq m
                
                # Get level name
                level_id = area_elem.LevelId
                level = doc.GetElement(level_id)
                level_name = level.Name if level else "N/A"
                
                # Format: linkified_id | View: view_name | Level: level_name | Area: X.XX sqm
                print("{} | View: {} | Level: {} | Area: {:.2f} sqm".format(
                    link, view.Name, level_name, area_sqm
                ))
            except Exception as e:
                logger.debug("Failed to linkify area %s: %s", elem_id, e)
    
    # Final report
    total_time = time.time() - overall_start
    perf_data.append(("Total Runtime", total_time))
    
    print("\n" + "="*50)
    if created_count > 0:
        print("Successfully filled {} hole(s)".format(created_count))
    else:
        print("No holes were filled")
    if failed_count > 0:
        print("Failed: {}".format(failed_count))
    
    print("\nTiming: " + ", ".join(
        "{}: {:.2f}s".format(label, duration) for label, duration in perf_data
    ))
    print("="*50)


if __name__ == '__main__':
    fill_holes()
