# -*- coding: utf-8 -*-
"""Fill holes in area boundaries using boolean union approach."""

__title__ = "Fill\nHoles"
__author__ = "Your Name"

from pyrevit import revit, DB, forms, script
import os
import sys
import time

# Import WPF for dialog
import clr
clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
import System
from System.Windows import Window
from System.Windows.Controls import CheckBox, StackPanel
from System.Windows.Media.Imaging import BitmapImage, BitmapCacheOption

SCRIPT_DIR = os.path.dirname(__file__)
LIB_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(SCRIPT_DIR))), "lib")

if LIB_DIR not in sys.path:
    sys.path.insert(0, LIB_DIR)

import data_manager

logger = script.get_logger()

# ============================================================
# CONSTANTS
# ============================================================

EXTRUSION_HEIGHT = 1.0  # Feet - height for solid extrusion in boolean operations
SQFT_TO_SQM = 0.09290304  # Square feet to square meters conversion
MIN_VOLUME_THRESHOLD = 0.001  # Minimum volume to consider solid as non-empty
MAX_RECURSION_DEPTH = 10  # Maximum depth for recursive hole filling


# ============================================================
# VIEW SELECTION DIALOG
# ============================================================

class ViewSelectionDialog(forms.WPFWindow):
    """Dialog for selecting AreaPlan views to process"""
    
    def __init__(self, area_schemes, selected_scheme_index=0, preselected_view_ids=None):
        """
        Args:
            area_schemes: List of (AreaScheme element, municipality) tuples
            selected_scheme_index: Index of the area scheme to select by default
            preselected_view_ids: Set of view ElementIds to preselect
        """
        forms.WPFWindow.__init__(self, 'ViewSelectionDialog.xaml')
        
        self._doc = revit.doc
        self._area_schemes = area_schemes
        self._preselected_view_ids = preselected_view_ids or set()
        self._view_checkboxes = {}  # {view_id_value: checkbox}
        
        # Wire up events
        self.btn_ok.Click += self.on_ok_clicked
        self.btn_cancel.Click += self.on_cancel_clicked
        self.btn_select_all.Click += self.on_select_all_clicked
        self.combo_areascheme.SelectionChanged += self.on_areascheme_changed
        
        # Populate area scheme dropdown
        for area_scheme, municipality in area_schemes:
            display_text = "{} ({})".format(area_scheme.Name, municipality)
            self.combo_areascheme.Items.Add(display_text)
        
        # Select the specified area scheme
        if area_schemes:
            self.combo_areascheme.SelectedIndex = selected_scheme_index
        
        # Show area scheme selector only when multiple municipalities exist
        try:
            municipalities = set()
            for _, municipality in area_schemes:
                if municipality:
                    municipalities.add(municipality)
            visibility = System.Windows.Visibility
            if len(municipalities) > 1:
                # Multiple municipalities defined in model - show selector
                self.combo_areascheme.Visibility = visibility.Visible
            else:
                # Single municipality (or none) - hide selector, still use selected index internally
                self.combo_areascheme.Visibility = visibility.Collapsed
        except Exception:
            # Fallback: keep default visibility from XAML
            pass
        
        # Load mode icons
        self._set_mode_images()

        # Result
        self.selected_views = None
        self.only_donut_holes = False
    
    def on_areascheme_changed(self, sender, args):
        """Handle area scheme selection change"""
        if self.combo_areascheme.SelectedIndex < 0:
            return
        
        # Rebuild view list for selected area scheme
        self._populate_view_list()
    
    def _populate_view_list(self):
        """Populate the list of AreaPlan views for the selected area scheme"""
        # Clear existing checkboxes
        self.panel_views.Children.Clear()
        self._view_checkboxes = {}
        
        if self.combo_areascheme.SelectedIndex < 0:
            return
        
        # Get selected area scheme
        area_scheme, municipality = self._area_schemes[self.combo_areascheme.SelectedIndex]
        
        # Get all AreaPlan views for this area scheme
        collector = DB.FilteredElementCollector(self._doc)
        all_views = collector.OfClass(DB.View).ToElements()
        
        area_plan_views = []
        for view in all_views:
            try:
                # Must be AreaPlan with matching scheme
                if not hasattr(view, 'AreaScheme'):
                    continue
                if not view.AreaScheme or view.AreaScheme.Id != area_scheme.Id:
                    continue
                
                # Check if view has placed areas
                areas = _get_areas_in_view(self._doc, view)
                if not areas:
                    continue
                
                area_plan_views.append(view)
            except:
                continue
        
        # Sort by elevation (level) from lowest to highest
        area_plan_views.sort(key=lambda v: v.Origin.Z if hasattr(v, 'Origin') else 0)
        
        # Create checkboxes for each view
        if not area_plan_views:
            no_views_text = System.Windows.Controls.TextBlock()
            no_views_text.Text = "No Area Plan views with placed areas found for this area scheme."
            no_views_text.Foreground = System.Windows.Media.Brushes.Gray
            no_views_text.FontStyle = System.Windows.FontStyles.Italic
            no_views_text.Margin = System.Windows.Thickness(5)
            self.panel_views.Children.Add(no_views_text)
            return
        
        for view in area_plan_views:
            view_id_value = _get_element_id_value(view.Id)
            
            # Get level name
            level_name = "N/A"
            if hasattr(view, 'GenLevel') and view.GenLevel:
                level_name = view.GenLevel.Name
            
            # Create checkbox
            checkbox = CheckBox()
            checkbox.Content = "{} (Level: {})".format(view.Name, level_name)
            checkbox.Margin = System.Windows.Thickness(5, 3, 5, 3)
            checkbox.Tag = view_id_value
            
            # Preselect if in preselected list
            if view.Id in self._preselected_view_ids:
                checkbox.IsChecked = True
            
            self.panel_views.Children.Add(checkbox)
            self._view_checkboxes[view_id_value] = checkbox
    
    def on_select_all_clicked(self, sender, args):
        """Select all views"""
        for checkbox in self._view_checkboxes.values():
            checkbox.IsChecked = True
    
    def on_ok_clicked(self, sender, args):
        """Handle OK button click"""
        if self.combo_areascheme.SelectedIndex < 0:
            forms.alert("Please select an area scheme.", exitscript=False)
            return
        
        # Get selected views
        selected_view_ids = []
        for view_id_value, checkbox in self._view_checkboxes.items():
            if checkbox.IsChecked:
                selected_view_ids.append(view_id_value)
        
        if not selected_view_ids:
            forms.alert("Please select at least one view.", exitscript=False)
            return
        
        # Get area scheme
        area_scheme, municipality = self._area_schemes[self.combo_areascheme.SelectedIndex]
        
        # Get view elements
        selected_views = []
        for view_id_value in selected_view_ids:
            view = self._doc.GetElement(DB.ElementId(System.Int64(int(view_id_value))))
            if view:
                selected_views.append(view)
        
        self.selected_views = selected_views
        
        # Get radio button state
        self.only_donut_holes = bool(self.rb_fill_donut_holes.IsChecked)
        
        self.DialogResult = True
        self.Close()
    
    def on_cancel_clicked(self, sender, args):
        """Handle Cancel button click"""
        self.DialogResult = False
        self.Close()

    def _set_mode_images(self):
        """Load diagram icons for hole filling modes."""
        icon_pairs = [
            (getattr(self, "img_mode_all_holes", None), "FillHolesIcon_split.png"),
            (getattr(self, "img_mode_donut_holes", None), "FillHolesIcon.png"),
        ]
        for image_control, filename in icon_pairs:
            if image_control is None:
                continue
            try:
                image_path = os.path.join(SCRIPT_DIR, filename)
                if not os.path.exists(image_path):
                    logger.debug("Mode icon not found: %s", image_path)
                    continue
                bitmap = BitmapImage()
                bitmap.BeginInit()
                bitmap.UriSource = System.Uri(image_path)
                bitmap.CacheOption = BitmapCacheOption.OnLoad
                bitmap.EndInit()
                image_control.Source = bitmap
            except Exception as exc:
                logger.debug("Failed to load mode icon %s: %s", filename, exc)


# ============================================================
# UTILITY FUNCTIONS
# ============================================================

def _get_element_id_value(element_id):
    """Get integer value from ElementId - compatible with Revit 2024, 2025 and 2026+"""
    try:
        return element_id.IntegerValue
    except AttributeError:
        return int(element_id.Value)


def _get_context_element():
    """Get context element(s) from selection or active view.
    
    Returns:
        tuple: (elements, element_type) where:
               - elements is a list of View elements (for "views" type) or single ViewSheet (for "sheet" type)
               - element_type is "views", "sheet", or None
               Returns (None, None) if no valid context
    """
    try:
        doc = revit.doc
        uidoc = revit.uidoc
        
        # Get current selection
        selection = uidoc.Selection
        selected_ids = selection.GetElementIds()
        
        if selected_ids and len(selected_ids) > 0:
            # Collect all selected AreaPlan views (from viewports or project browser)
            selected_area_plan_views = []
            selected_sheets = []
            
            for elem_id in selected_ids:
                elem = doc.GetElement(elem_id)
                
                # Check if it's a viewport (view on sheet)
                if isinstance(elem, DB.Viewport):
                    view_id = elem.ViewId
                    view = doc.GetElement(view_id)
                    # Check if it's an area plan with defined municipality
                    if hasattr(view, 'AreaScheme') and view.AreaScheme:
                        if data_manager.get_municipality(view.AreaScheme):
                            selected_area_plan_views.append(view)
                
                # Check if it's a view (selected in project browser)
                elif isinstance(elem, DB.View) and not isinstance(elem, DB.ViewSheet):
                    if hasattr(elem, 'AreaScheme') and elem.AreaScheme:
                        if data_manager.get_municipality(elem.AreaScheme):
                            selected_area_plan_views.append(elem)
                
                # Check if it's a sheet
                elif isinstance(elem, DB.ViewSheet):
                    selected_sheets.append(elem)
            
            # Return selected AreaPlan views if any were found
            if selected_area_plan_views:
                return (selected_area_plan_views, "views")
            
            # Return first selected sheet if any
            if selected_sheets:
                return (selected_sheets[0], "sheet")
        
        # Priority 2: Check active view if nothing is selected
        active_view = uidoc.ActiveView
        
        # Check if active view is a sheet
        if isinstance(active_view, DB.ViewSheet):
            return (active_view, "sheet")
        
        # Check if active view is an area plan
        if hasattr(active_view, 'AreaScheme') and active_view.AreaScheme:
            if data_manager.get_municipality(active_view.AreaScheme):
                return ([active_view], "views")
        
    except Exception:
        pass  # Silently fail
    
    return (None, None)


def _get_defined_area_schemes():
    """Get all area schemes with municipality defined.
    
    Returns:
        List of (AreaScheme element, municipality) tuples
    """
    doc = revit.doc
    collector = DB.FilteredElementCollector(doc)
    area_schemes = list(collector.OfClass(DB.AreaScheme).ToElements())
    
    defined_schemes = []
    for scheme in area_schemes:
        municipality = data_manager.get_municipality(scheme)
        if municipality:
            defined_schemes.append((scheme, municipality))
    
    return defined_schemes


def _get_views_on_sheet(sheet):
    """Get AreaPlan views placed on a sheet.
    
    Args:
        sheet: ViewSheet element
        
    Returns:
        List of View elements (AreaPlan views only)
    """
    try:
        view_ids = sheet.GetAllPlacedViews()
        area_plan_views = []
        
        for view_id in view_ids:
            view = revit.doc.GetElement(view_id)
            if hasattr(view, 'AreaScheme') and view.AreaScheme:
                area_plan_views.append(view)
        
        return area_plan_views
    except:
        return []


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


def _get_curveloop_bbox(curve_loop, sample_points=None):
    """Get 2D bounding box of a CurveLoop.
    
    Args:
        curve_loop: CurveLoop to measure
        sample_points: List of t values to sample (default: [0.0, 0.25, 0.5, 0.75, 1.0])
    
    Returns:
        Tuple of (min_x, max_x, min_y, max_y) or None if failed
    """
    if not curve_loop or curve_loop.NumberOfCurves() == 0:
        return None
    
    if sample_points is None:
        sample_points = [0.0, 0.25, 0.5, 0.75, 1.0]
    
    min_x = float("inf")
    max_x = float("-inf")
    min_y = float("inf")
    max_y = float("-inf")
    has_points = False
    
    for curve in curve_loop:
        try:
            for t in sample_points:
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
    
    if not has_points or min_x == float("inf"):
        return None
    
    return (min_x, max_x, min_y, max_y)


def _get_curveloop_centroid(curve_loop):
    """Get centroid of a CurveLoop using bounding box center."""
    bbox = _get_curveloop_bbox(curve_loop)
    if not bbox:
        return None
    
    min_x, max_x, min_y, max_y = bbox
    return DB.XYZ((min_x + max_x) / 2.0, (min_y + max_y) / 2.0, 0.0)


def _get_curveloop_bbox_area(curve_loop):
    """Approximate area of a CurveLoop using its 2D bounding box.
    Used to classify exterior vs hole loops.
    """
    bbox = _get_curveloop_bbox(curve_loop, sample_points=[0.0, 0.5, 1.0])
    if not bbox:
        return 0.0
    
    min_x, max_x, min_y, max_y = bbox
    return max(0.0, (max_x - min_x) * (max_y - min_y))


def _curveloops_to_solid(curve_loops):
    """Convert CurveLoop(s) to a solid for boolean operations.
    
    Args:
        curve_loops: Single CurveLoop or list of CurveLoops
    
    Returns:
        Solid object or None if conversion fails
    """
    try:
        # Normalize to list
        if isinstance(curve_loops, DB.CurveLoop):
            curve_loops = [curve_loops]
        
        if not curve_loops:
            return None
        
        # Extrude to create solid
        solid = DB.GeometryCreationUtilities.CreateExtrusionGeometry(
            curve_loops,
            DB.XYZ(0, 0, 1),  # Extrusion direction (up)
            EXTRUSION_HEIGHT
        )
        return solid
    except Exception as e:
        logger.debug("Failed to convert CurveLoop(s) to solid: %s", e)
        return None


def _area_to_solid(area_elem):
    """Convert an area element to a solid for boolean operations."""
    try:
        loops = _get_boundary_loops(area_elem)
        if not loops:
            return None
        
        # Convert all loops to CurveLoops
        curve_loops = []
        for loop in loops:
            curve_loop = _boundary_loop_to_curveloop(loop)
            if curve_loop:
                curve_loops.append(curve_loop)
        
        return _curveloops_to_solid(curve_loops)
    except Exception as e:
        logger.debug("Failed to convert area to solid: %s", e)
        return None


def _get_bottom_face(solid):
    """Get the bottom face (lowest Z) from a solid.
    
    Args:
        solid: Solid to search
    
    Returns:
        Face object or None if not found
    """
    try:
        min_z = float('inf')
        bottom_face = None
        
        for face in solid.Faces:
            try:
                mesh = face.Triangulate()
                if mesh and mesh.Vertices.Count > 0:
                    z = mesh.Vertices[0].Z
                    if z < min_z:
                        min_z = z
                        bottom_face = face
            except Exception:
                continue
        
        return bottom_face
    except Exception as e:
        logger.debug("Failed to find bottom face: %s", e)
        return None


def _get_loops_from_solid(solid, holes_only=False):
    """Get CurveLoops from a solid's bottom face.
    
    Args:
        solid: Solid to extract loops from
        holes_only: If True, return only interior holes. If False, return all loops.
    
    Returns:
        List of CurveLoop objects
    """
    try:
        bottom_face = _get_bottom_face(solid)
        if not bottom_face:
            return []
        
        # Get all curve loops from the bottom face
        edge_loops = bottom_face.GetEdgesAsCurveLoops()
        loops = [loop for loop in edge_loops]
        
        if not loops:
            return []
        
        # If requesting all loops, return them
        if not holes_only:
            return loops
        
        # Otherwise, filter to get only holes (exclude largest loop = exterior)
        if len(loops) <= 1:
            return []  # No holes
        
        # Classify by area: largest = exterior, others = holes
        loop_areas = [_get_curveloop_bbox_area(loop) for loop in loops]
        if not any(a > 0.0 for a in loop_areas):
            return []
        
        outer_index = max(range(len(loops)), key=lambda i: loop_areas[i])
        
        return [loop for idx, loop in enumerate(loops) if idx != outer_index]
    except Exception as e:
        logger.debug("Failed to get loops from solid: %s", e)
        return []

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
    
    # Convert all areas to solids
    solids = []
    for area in areas:
        solid = _area_to_solid(area)
        if solid:
            solids.append(solid)
    
    if not solids:
        return None
    
    # Create union
    try:
        union = solids[0]
        for solid in solids[1:]:
            try:
                union = DB.BooleanOperationsUtils.ExecuteBooleanOperation(
                    union, solid, DB.BooleanOperationsType.Union
                )
            except Exception as e:
                logger.debug("Failed to union solid: %s", e)
                continue
        return union
    except Exception as e:
        logger.debug("Failed to create union: %s", e)
        return None


def _extract_holes_from_union(union_solid):
    """Extract interior holes from the union solid.
    
    Returns: List of CurveLoop objects representing holes
    """
    if not union_solid:
        return []
    
    try:
        holes = _get_loops_from_solid(union_solid, holes_only=True)
        return holes
    except Exception as e:
        logger.debug("Failed to extract holes: %s", e)
        return []


def _point_in_curveloop_2d(point, curve_loop):
    """Test if a 2D point (X,Y) is inside a CurveLoop using ray-casting.
    
    Handles all curve types including lines, arcs, splines, and ellipses.
    
    Args:
        point: XYZ point to test
        curve_loop: CurveLoop boundary
    
    Returns:
        True if point is inside, False otherwise
    """
    try:
        # Ray-casting algorithm: count intersections with a ray from point to infinity
        px, py, pz = point.X, point.Y, point.Z
        intersections = 0
        
        # Create a horizontal ray from point to a far point to the right
        # Use a very large X value to ensure the ray extends beyond all geometry
        far_x = px + 1000000.0  # 1 million feet should be far enough
        ray_start = DB.XYZ(px, py, pz)
        ray_end = DB.XYZ(far_x, py, pz)
        
        # Create a Line for the ray
        try:
            ray_line = DB.Line.CreateBound(ray_start, ray_end)
        except Exception:
            # If points are too close, test fails - assume outside
            return False
        
        for curve in curve_loop:
            try:
                # Use Revit's SetComparisonResult to find intersections
                # This properly handles all curve types (lines, arcs, splines, ellipses)
                result = curve.Intersect(ray_line)
                
                if result == DB.SetComparisonResult.Overlap:
                    # Curves overlap - get intersection results
                    intersection_result_array = clr.Reference[DB.IntersectionResultArray]()
                    result = curve.Intersect(ray_line, intersection_result_array)
                    
                    if result == DB.SetComparisonResult.Overlap and intersection_result_array.Value:
                        # Count valid intersections to the right of the point
                        for i in range(intersection_result_array.Value.Size):
                            int_result = intersection_result_array.Value.get_Item(i)
                            int_point = int_result.XYZPoint
                            
                            # Only count if intersection is to the right of test point
                            if int_point.X > px + 0.0001:  # Small tolerance
                                intersections += 1
            except Exception:
                # If intersection test fails, skip this curve
                continue
        
        # Odd number of intersections = inside
        return (intersections % 2) == 1
    except Exception as e:
        logger.debug("Point-in-polygon test failed: %s", e)
        return False


def _get_areas_inside_boundary(all_areas, donut_area):
    """Find all areas whose centroids lie inside a donut area's boundary.
    
    Uses the donut area's exterior boundary for testing, which includes
    considering areas that are in the hole regions.
    
    Args:
        all_areas: List of all Area elements to check
        donut_area: The donut Area element whose boundary to test against
    
    Returns:
        List of Area elements that are inside the boundary (excluding the donut itself)
    """
    try:
        # Get the donut's exterior boundary for containment testing
        boundary_loops = _get_boundary_loops(donut_area)
        if not boundary_loops:
            logger.debug("Failed to get boundary loops from donut area")
            return []
        
        exterior_loop = _boundary_loop_to_curveloop(boundary_loops[0])
        if not exterior_loop:
            logger.debug("Failed to convert exterior boundary to CurveLoop")
            return []
        
        contained_areas = []
        
        for area in all_areas:
            try:
                # Skip the donut area itself
                if area.Id == donut_area.Id:
                    continue
                
                # Get area's location point (centroid)
                location = area.Location
                if not location or not hasattr(location, 'Point'):
                    continue
                
                point = location.Point
                
                # Test if point is inside the exterior boundary using 2D ray-casting
                is_inside = _point_in_curveloop_2d(point, exterior_loop)
                
                if is_inside:
                    contained_areas.append(area)
                    logger.debug("Area %s is inside the boundary", _get_element_id_value(area.Id))
            
            except Exception as e:
                logger.debug("Failed to test area %s containment: %s", 
                            _get_element_id_value(area.Id), e)
                continue
        
        return contained_areas
    
    except Exception as e:
        logger.debug("Failed to find contained areas: %s", e)
        return []


def _identify_donut_areas(areas):
    """Identify areas that have interior holes (donuts).
    
    Args:
        areas: List of Area elements
    
    Returns:
        List of (area, exterior_loop, hole_count) tuples for donut areas
    """
    donut_areas = []
    
    for area in areas:
        try:
            # Get boundary loops for this area
            boundary_loops = _get_boundary_loops(area)
            
            if len(boundary_loops) <= 1:
                # No holes in this area - skip
                continue
            
            # This area is a DONUT (has holes)
            # boundary_loops[0] = exterior boundary
            # boundary_loops[1:] = interior holes
            exterior_loop = _boundary_loop_to_curveloop(boundary_loops[0])
            if exterior_loop:
                hole_count = len(boundary_loops) - 1
                donut_areas.append((area, exterior_loop, hole_count))
                logger.debug("Area %s is a donut with %d hole(s)", 
                            _get_element_id_value(area.Id), hole_count)
        
        except Exception as e:
            logger.debug("Failed to check Area %s for donut: %s", 
                        _get_element_id_value(area.Id), e)
            continue
    
    return donut_areas


def _fill_hole_recursive(doc, view, hole_loop, usage_type_value, usage_type_name, created_area_ids, depth=0, max_depth=10):
    """Recursively fill a hole by:
    1. Place area at hole centroid
    2. Convert area to solid and subtract from hole
    3. Find remaining holes
    4. Recursively fill them
    
    Args:
        doc: Revit document
        view: AreaPlan view
        hole_loop: CurveLoop of the hole (for centroid calculation)
        usage_type_value: Usage type value to set
        usage_type_name: Usage type name to set
        created_area_ids: List to accumulate created area IDs
        depth: Current recursion depth
        max_depth: Maximum recursion depth
    
    Returns:
        Number of areas created in this branch
    """
    if depth >= max_depth:
        logger.debug("Max recursion depth reached")
        return 0
    
    # Get centroid of hole
    centroid = _get_curveloop_centroid(hole_loop)
    if not centroid:
        logger.debug("Depth %d: Failed to calculate centroid", depth)
        return 0
    
    logger.debug("Depth %d: Trying centroid at (%.3f, %.3f)", depth, centroid.X, centroid.Y)
    
    # Try creating area at centroid
    sub_txn = DB.SubTransaction(doc)
    try:
        sub_txn.Start()
        
        uv = DB.UV(centroid.X, centroid.Y)
        new_area = doc.Create.NewArea(view, uv)
        
        if not new_area:
            sub_txn.RollBack()
            return 0
        
        # Regenerate to get boundaries
        try:
            doc.Regenerate()
        except Exception:
            pass
        
        if new_area.Area <= 0:
            doc.Delete(new_area.Id)
            sub_txn.RollBack()
            return 0
        
        # Set parameters
        _set_area_parameters(new_area, usage_type_value, usage_type_name)
        
        # SUCCESS - commit this area
        sub_txn.Commit()
        created_area_ids.append(new_area.Id)
        
        area_id_value = _get_element_id_value(new_area.Id)
        logger.debug("Depth %d: Created area %s (%.2f sqm)", depth, area_id_value, new_area.Area * SQFT_TO_SQM)
        
        areas_created = 1
        
        # Try to find remaining unfilled space after this area placement
        hole_solid = _curveloops_to_solid(hole_loop)
        area_solid = _area_to_solid(new_area)
        
        if not hole_solid or not area_solid:
            logger.debug("Depth %d: Failed to convert to solids - stopping recursion", depth)
            return areas_created
        
        # Subtract area from hole
        try:
            remainder_solid = DB.BooleanOperationsUtils.ExecuteBooleanOperation(
                hole_solid,
                area_solid,
                DB.BooleanOperationsType.Difference
            )
            
            if not remainder_solid or remainder_solid.Volume < MIN_VOLUME_THRESHOLD:
                # Hole completely filled!
                logger.debug("Depth %d: Hole completely filled", depth)
                return areas_created
            
            # Get all loops from remainder (each represents unfilled space)
            remaining_loops = _get_loops_from_solid(remainder_solid, holes_only=False)
            
            if not remaining_loops:
                logger.debug("Depth %d: No remaining space after subtraction", depth)
                return areas_created
            
            logger.debug("Depth %d: Found %d remaining loop(s) to fill", depth, len(remaining_loops))
            
            # Recursively fill all remaining loops
            for remaining_loop in remaining_loops:
                sub_areas = _fill_hole_recursive(
                    doc, view,
                    remaining_loop,
                    usage_type_value, usage_type_name,
                    created_area_ids,
                    depth + 1,
                    max_depth
                )
                areas_created += sub_areas
            
            return areas_created
            
        except Exception as e:
            logger.debug("Depth %d: Boolean subtraction failed - %s", depth, e)
            return areas_created
        
    except Exception as e:
        logger.debug("Depth %d: Failed to create area - %s", depth, e)
        try:
            if sub_txn.HasStarted() and not sub_txn.HasEnded():
                sub_txn.RollBack()
        except Exception:
            pass
        return 0


def _create_areas_in_holes(doc, view, holes, usage_type_value, usage_type_name):
    """Create new Area elements to fill holes using recursive boolean subtraction.
    
    For each hole:
      1. Place area at centroid
      2. Subtract area from hole
      3. Recursively fill remaining holes
    
    Returns: (created_count, failed_count, created_area_ids)
    """
    if not holes:
        return 0, 0, []
    
    total_created = 0
    total_failed = 0
    all_created_area_ids = []  # store all created ElementId objects
    
    # Use a single main transaction
    with revit.Transaction("Fill Area Holes"):
        for hole_loop in holes:
            # Recursively fill this hole
            areas_created = _fill_hole_recursive(
                doc, view,
                hole_loop,
                usage_type_value, usage_type_name,
                all_created_area_ids,
                depth=0,
                max_depth=MAX_RECURSION_DEPTH
            )
            
            if areas_created > 0:
                total_created += areas_created
            else:
                total_failed += 1
    
    return total_created, total_failed, all_created_area_ids


def _union_and_fill_holes(doc, view, areas_to_union, usage_type_value, usage_type_name, group_label=""):
    """Create union of areas, extract holes, and fill them.
    
    This is the core hole-filling method used by both modes:
    - All Gaps mode: Called once with all areas in view
    - Islands Only mode: Called per donut area with contained areas
    
    Args:
        doc: Revit document
        view: AreaPlan view
        areas_to_union: List of Area elements to union
        usage_type_value: Usage type value for created areas
        usage_type_name: Usage type name for created areas
        group_label: Label for console output (e.g., "Donut Area 123")
    
    Returns:
        (created_count, failed_count, created_area_ids)
    """
    if not areas_to_union:
        return 0, 0, []
    
    union_solid = _create_union_of_areas(areas_to_union)
    if not union_solid:
        return 0, 0, []
    
    # Extract holes from union
    holes = _extract_holes_from_union(union_solid)
    if not holes:
        return 0, 0, []
    
    # Fill the holes
    created_count, failed_count, created_area_ids = _create_areas_in_holes(
        doc, view, holes, usage_type_value, usage_type_name
    )
    
    return created_count, failed_count, created_area_ids


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
    1. Show dialog to select AreaPlan views to process
    2. Create a boolean union of all area boundaries in each view
    3. Extract holes from the unified region
    4. Create new areas at the centroid of each hole
    """
    overall_start = time.time()
    doc = revit.doc
    
    # Get defined area schemes
    defined_schemes = _get_defined_area_schemes()
    if not defined_schemes:
        forms.alert(
            "No area schemes with municipality defined.\n\n"
            "Please define an area scheme using the Calculation Setup tool first.",
            exitscript=True
        )
    
    # Get context element to determine default area scheme
    context_elem, context_type = _get_context_element()
    
    # Determine which area scheme to select by default
    selected_scheme_index = 0
    preselected_view_ids = set()
    
    if context_elem:
        if context_type == "views":
            # Multiple AreaPlan views selected (or single active view)
            # Use the first view's area scheme
            first_view = context_elem[0]
            view_scheme = first_view.AreaScheme
            for i, (scheme, _) in enumerate(defined_schemes):
                if scheme.Id == view_scheme.Id:
                    selected_scheme_index = i
                    break
            # Preselect all selected views
            for view in context_elem:
                preselected_view_ids.add(view.Id)
        
        elif context_type == "sheet":
            # Get area plans on this sheet
            views_on_sheet = _get_views_on_sheet(context_elem)
            if views_on_sheet:
                # Use the first view's area scheme
                first_view_scheme = views_on_sheet[0].AreaScheme
                for i, (scheme, _) in enumerate(defined_schemes):
                    if scheme.Id == first_view_scheme.Id:
                        selected_scheme_index = i
                        break
                # Preselect all views on sheet
                for view in views_on_sheet:
                    preselected_view_ids.add(view.Id)
    
    # Show view selection dialog
    dialog = ViewSelectionDialog(
        defined_schemes,
        selected_scheme_index,
        preselected_view_ids
    )
    
    if not dialog.ShowDialog():
        return  # User cancelled
    
    selected_views = dialog.selected_views
    if not selected_views:
        return
    
    # Sort processing order by elevation from lowest to highest
    selected_views.sort(key=lambda v: v.Origin.Z if hasattr(v, 'Origin') else 0)
    
    # Get the checkbox option
    only_donut_holes = dialog.only_donut_holes
    
    # Process each selected view
    output = script.get_output()
    perf_data = []
    total_created_all = 0
    total_failed_all = 0
    
    for view in selected_views:
        area_scheme = getattr(view, "AreaScheme", None)
        if area_scheme is None:
            continue
        
        created, failed = _process_view_holes(doc, view, output, perf_data, only_donut_holes)
        total_created_all += created
        total_failed_all += failed
    
    # Final summary - single line
    print("\nTotal: {} area(s) created | {} failed".format(total_created_all, total_failed_all))


def _process_view_holes(doc, view, output, perf_data, only_donut_holes=False):
    """Process holes for a single view.
    
    Args:
        doc: Revit document
        view: AreaPlan view to process
        output: Script output for linkifying
        perf_data: List to accumulate performance data
        only_donut_holes: If True, only fill holes within individual areas (donuts),
                         not gaps between areas
    
    Returns:
        tuple: (created_count, failed_count)
    """
    
    # Get areas
    areas_start = time.time()
    areas = _get_areas_in_view(doc, view)
    if not areas:
        return 0, 0
    
    perf_data.append(("Collect Areas ({})".format(view.Name), time.time() - areas_start))
    
    # Print view header
    view_link = output.linkify(view.Id)
    print("-" * 80)
    print("{} {}".format(view_link, view.Name))
    
    # Get municipality settings
    usage_type_value, usage_type_name = _get_hole_usage_for_municipality(doc, view)
    
    # Process based on mode
    process_start = time.time()
    all_created_area_ids = []
    total_created = 0
    total_failed = 0
    
    if only_donut_holes:
        # ISLANDS ONLY MODE: Process each donut area individually
        donut_areas = _identify_donut_areas(areas)
        
        if not donut_areas:
            return 0, 0
        
        perf_data.append(("Identify Donuts ({})".format(view.Name), time.time() - process_start))
        
        # Process each donut area
        for donut_area, exterior_loop, hole_count in donut_areas:
            # Find all areas contained within this donut's boundary
            contained_areas = _get_areas_inside_boundary(areas, donut_area)
            
            # Union the donut area with all contained areas
            areas_to_union = [donut_area] + contained_areas
            
            # Extract and fill holes in this union
            created, failed, created_ids = _union_and_fill_holes(
                doc, view, areas_to_union, 
                usage_type_value, usage_type_name
            )
            
            total_created += created
            total_failed += failed
            all_created_area_ids.extend(created_ids)
        
        perf_data.append(("Process Donuts ({})".format(view.Name), time.time() - process_start))
    
    else:
        # ALL GAPS MODE: Process all areas at once
        created, failed, created_ids = _union_and_fill_holes(
            doc, view, areas,
            usage_type_value, usage_type_name
        )
        
        total_created = created
        total_failed = failed
        all_created_area_ids = created_ids
        
        perf_data.append(("Union and Fill ({})".format(view.Name), time.time() - process_start))
    
    # Use the accumulated results
    created_count = total_created
    failed_count = total_failed
    created_area_ids = all_created_area_ids
    
    # Summary with linkified areas
    if created_area_ids:
        for elem_id in created_area_ids:
            try:
                link = output.linkify(elem_id)
                area_elem = doc.GetElement(elem_id)
                area_sqm = area_elem.Area * SQFT_TO_SQM if area_elem else 0
                # Use HTML non-breaking spaces for indentation (two spaces)
                output.print_html("&nbsp;&nbsp;{} ({:.2f} sqm)".format(link, area_sqm))
            except Exception:
                pass
    
    return created_count, failed_count


if __name__ == '__main__':
    fill_holes()
