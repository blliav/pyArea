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
    """Create new Area elements at the centroid of each hole.
    
    Returns: (created_count, failed_count, created_area_ids)
    """
    if not holes:
        return 0, 0, []
    
    print("\n" + "="*50)
    print("STEP 3: Filling Holes")
    print("="*50)
    
    created_count = 0
    failed_count = 0
    created_area_ids = []  # store ElementId objects for created areas
    
    # Use a single main transaction with subtransactions for each hole
    # This groups all hole filling into one undo operation
    with revit.Transaction("Fill Area Holes"):
        for hole_idx, hole in enumerate(holes, 1):
            # Use subtransaction for each hole
            # This ensures proper regeneration between holes while keeping one undo entry
            sub_txn = DB.SubTransaction(doc)
            try:
                sub_txn.Start()
                
                # Get centroid of hole
                centroid = _get_curveloop_centroid(hole)
                if not centroid:
                    logger.debug("Hole %d: Failed to calculate centroid", hole_idx)
                    print("  ERROR: Could not calculate centroid")
                    sub_txn.RollBack()
                    failed_count += 1
                    continue
                
                logger.debug("Hole %d: Centroid at (%.3f, %.3f)", hole_idx, centroid.X, centroid.Y)
                
                # Try centroid first, then various offsets if centroid fails
                # This handles rotated/irregular holes where centroid might be outside
                test_points = [
                    (centroid.X, centroid.Y),  # Centroid
                ]
                
                # Add radial offsets in 8 directions at multiple distances
                for distance in [0.1, 0.3, 0.5]:  # feet
                    test_points.extend([
                        (centroid.X + distance, centroid.Y),          # East
                        (centroid.X - distance, centroid.Y),          # West
                        (centroid.X, centroid.Y + distance),          # North
                        (centroid.X, centroid.Y - distance),          # South
                        (centroid.X + distance*0.7, centroid.Y + distance*0.7),  # NE
                        (centroid.X - distance*0.7, centroid.Y + distance*0.7),  # NW
                        (centroid.X + distance*0.7, centroid.Y - distance*0.7),  # SE
                        (centroid.X - distance*0.7, centroid.Y - distance*0.7),  # SW
                    ])
                
                new_area = None
                area_size = 0
                attempted = 0
                
                for test_x, test_y in test_points:
                    attempted += 1
                    try:
                        # Create UV point for area placement
                        uv = DB.UV(test_x, test_y)
                        
                        # Create new area
                        temp_area = doc.Create.NewArea(view, uv)
                        
                        if temp_area is None:
                            continue
                        
                        # Force Revit to regenerate and calculate area size
                        try:
                            doc.Regenerate()
                        except Exception:
                            pass
                        
                        # Check if area has valid size
                        temp_size = temp_area.Area
                        
                        if temp_size > 0:
                            # Success!
                            new_area = temp_area
                            area_size = temp_size
                            break
                        else:
                            # Zero size - delete and try next point
                            try:
                                doc.Delete(temp_area.Id)
                            except Exception:
                                pass
                    except Exception as ex:
                        logger.debug("Hole %d position %d failed: %s", hole_idx, attempted, ex)
                        continue
                
                if new_area is None or area_size == 0:
                    logger.debug("Hole %d: All %d test points failed", hole_idx, attempted)
                    sub_txn.RollBack()
                    failed_count += 1
                    continue
                
                area_elem_id = new_area.Id
                area_id_value = _get_element_id_value(area_elem_id)
                
                # Set usage type parameters on the new area
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
                        logger.debug("Failed to set parameters on Area %s: %s", area_id_value, param_ex)
                
                # Success! Commit the subtransaction
                sub_txn.Commit()
                created_area_ids.append(area_elem_id)
                created_count += 1
                logger.debug("Successfully created Area %s", area_id_value)
                
            except Exception as e:
                logger.debug("Failed to create area for hole %d: %s", hole_idx, e)
                try:
                    if sub_txn.HasStarted() and not sub_txn.HasEnded():
                        sub_txn.RollBack()
                except Exception:
                    pass
                failed_count += 1
    
    return created_count, failed_count, created_area_ids


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
