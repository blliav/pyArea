# -*- coding: utf-8 -*-
"""Fill holes in area boundaries using boolean union approach.

DEBUG MODE:
    Set DEBUG_VERBOSE = True to enable detailed console output
    Set DEBUG_VISUALIZATION = True to show holes in 3D using dc3dserver

KNOWN ISSUES AND POTENTIAL SOLUTIONS:

1. THIN SLIVERS BETWEEN AREAS:
   - Problem: Very thin gaps (< 0.01 ft) between adjacent areas cause boolean 
     operations to fail or produce invalid geometry.
   - Solution A: Offset curve loops outward by a small amount (0.01-0.05 ft) 
     before creating solids. This ensures overlapping geometry that booleans 
     can handle. After union, the result will have clean boundaries.
   - Solution B: Use 2D polygon analysis (e.g., Clipper library) which handles
     thin geometry better than 3D boolean operations.

2. TANGENT/COINCIDENT EDGES:
   - Problem: When two areas share an exact edge, boolean union may fail due
     to numerical precision issues.
   - Solution: Apply small random perturbation to one set of curves, or use
     a robust 2D polygon library.

3. SELF-INTERSECTING BOUNDARIES:
   - Problem: Area boundaries that cross themselves cannot be extruded.
   - Solution: Validate and repair boundaries before processing.

4. ALTERNATIVE APPROACHES:
   - 2D Polygon Analysis: Use a library like Clipper (Python: pyclipper) to
     perform 2D polygon union/difference operations. This is more robust for
     thin slivers and complex intersections.
   - Raster/Grid Analysis: Convert boundaries to a grid, perform operations
     on the grid, then vectorize back. Handles any geometry.
   - Edge Graph Analysis: Build a graph of all boundary edges, find enclosed
     regions using graph traversal algorithms.

TODO for improved reliability:
   - [ ] Add curve loop offset option for thin sliver handling
   - [ ] Integrate pyclipper for 2D polygon operations
   - [ ] Add boundary validation before processing
   - [ ] Try multiple placement points if centroid fails
"""

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

# Import 2D polygon operations (WPF-based, IronPython 2.7 compatible)
try:
    from polygon_2d import (Polygon2D, find_gap_points_2d, find_all_gap_regions_2d, 
                            find_all_gap_regions_2d_from_polygons)
    POLYGON_2D_AVAILABLE = True
except ImportError as e:
    POLYGON_2D_AVAILABLE = False
    print("[WARNING] polygon_2d module not available: {}".format(e))

logger = script.get_logger()

# ============================================================
# CONSTANTS
# ============================================================

EXTRUSION_HEIGHT = 1.0  # Feet - height for solid extrusion in boolean operations
SQFT_TO_SQM = 0.09290304  # Square feet to square meters conversion
MIN_VOLUME_THRESHOLD = 0.001  # Minimum volume to consider solid as non-empty
MAX_RECURSION_DEPTH = 10  # Maximum depth for recursive hole filling

# Debug visualization - set to True to enable 3D visualization of failed operations
DEBUG_VISUALIZATION = True
DEBUG_VERBOSE = True  # Print detailed debug info for each failure

# Track failures for diagnostics
class DebugStats:
    """Track debug statistics for failed operations."""
    def __init__(self):
        self.reset()
    
    def reset(self):
        self.failed_area_to_solid = []  # List of (area_id, area_name, error_msg, boundary_info)
        self.failed_boolean_ops = []  # List of (solid1_info, solid2_info, error_msg)
        self.failed_holes = []  # List of (hole_loop, centroid, error_msg)
        self.successful_solids = []  # List of solid objects for visualization
        self.hole_loops = []  # List of hole CurveLoops for visualization
        self.all_area_solids = []  # List of (area_id, solid) tuples
    
    def print_summary(self):
        """Print diagnostic summary of failures."""
        print("\n" + "=" * 80)
        print("DEBUG DIAGNOSTIC SUMMARY")
        print("=" * 80)
        
        if self.failed_area_to_solid:
            print("\n[FAILED] Area to Solid Conversions: {}".format(len(self.failed_area_to_solid)))
            for area_id, area_name, error_msg, boundary_info in self.failed_area_to_solid:
                print("  - Area {} ({}): {}".format(area_id, area_name, error_msg))
                if boundary_info:
                    print("    Boundary Info: {}".format(boundary_info))
        
        if self.failed_boolean_ops:
            print("\n[FAILED] Boolean Operations: {}".format(len(self.failed_boolean_ops)))
            for solid1_info, solid2_info, error_msg in self.failed_boolean_ops:
                print("  - Union failed: {} + {}".format(solid1_info, solid2_info))
                print("    Error: {}".format(error_msg))
        
        if self.failed_holes:
            print("\n[FAILED] Hole Filling: {}".format(len(self.failed_holes)))
            for hole_info, centroid, error_msg in self.failed_holes:
                if centroid:
                    print("  - Hole at ({:.2f}, {:.2f}): {}".format(centroid.X, centroid.Y, error_msg))
                else:
                    print("  - Hole (no centroid): {}".format(error_msg))
        
        total_failures = len(self.failed_area_to_solid) + len(self.failed_boolean_ops) + len(self.failed_holes)
        if total_failures == 0:
            print("\n[OK] No failures detected.")
        else:
            print("\n[ANALYSIS] Total failures: {}".format(total_failures))
            print("\nPossible causes and solutions:")
            if self.failed_area_to_solid:
                print("  1. AREA TO SOLID FAILURES:")
                print("     - Self-intersecting boundary curves")
                print("     - Very thin slivers (< 0.01 ft width)")
                print("     - Open or disconnected boundary loops")
                print("     Solution: Check area boundaries for geometry issues")
            if self.failed_boolean_ops:
                print("  2. BOOLEAN OPERATION FAILURES:")
                print("     - Tangent or near-tangent edges between areas")
                print("     - Coincident edges with opposite orientations")
                print("     - Very small gaps or overlaps")
                print("     Solution: Offset curves slightly to ensure overlap")
            if self.failed_holes:
                print("  3. HOLE FILLING FAILURES:")
                print("     - Hole centroid falls outside valid area boundary")
                print("     - Hole too small to place area")
                print("     Solution: Consider using 2D polygon analysis")
        
        print("=" * 80 + "\n")

# Global debug stats instance
debug_stats = DebugStats()


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


def _boundary_loop_to_curveloop(boundary_loop, debug=False):
    """Convert a boundary segment loop to a CurveLoop.
    
    Simple approach: just collect the curves from segments.
    """
    try:
        if not boundary_loop:
            return None
        
        curves = []
        for segment in boundary_loop:
            try:
                curve = segment.GetCurve()
                if curve:
                    curves.append(curve)
            except:
                pass
        
        if not curves:
            return None
        
        # Create CurveLoop using CreateViaCopy which is more tolerant
        try:
            curve_loop = DB.CurveLoop.CreateViaCopy(curves)
            return curve_loop
        except:
            # Fallback: create CurveLoop and add curves
            curve_loop = DB.CurveLoop()
            for curve in curves:
                try:
                    curve_loop.Append(curve)
                except:
                    pass
            return curve_loop if curve_loop.NumberOfCurves() > 0 else None
            
    except Exception as e:
        if debug:
            print("    [DEBUG] Failed to create CurveLoop: {}".format(e))
        return None


def _boundary_loop_to_points(boundary_loop):
    """Extract 2D points from a boundary segment loop.
    
    This bypasses CurveLoop creation entirely - just gets the points
    from each curve for 2D polygon operations.
    
    Args:
        boundary_loop: Revit boundary segment loop
    
    Returns:
        List of (x, y) tuples, or None if failed
    """
    try:
        if not boundary_loop:
            return None
        
        points = []
        for segment in boundary_loop:
            try:
                curve = segment.GetCurve()
                if curve:
                    # Tessellate the curve to get points
                    # This handles lines, arcs, splines, etc.
                    tessellated = curve.Tessellate()
                    for i, pt in enumerate(tessellated):
                        # Skip the last point of each curve (it's the start of the next)
                        if i < len(tessellated) - 1:
                            points.append((pt.X, pt.Y))
            except:
                pass
        
        # Remove duplicates while preserving order
        if points:
            cleaned = [points[0]]
            for p in points[1:]:
                if abs(p[0] - cleaned[-1][0]) > 0.001 or abs(p[1] - cleaned[-1][1]) > 0.001:
                    cleaned.append(p)
            return cleaned if len(cleaned) >= 3 else None
        
        return None
        
    except:
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


def _get_curveloop_info(curve_loop):
    """Get diagnostic info about a CurveLoop."""
    try:
        if not curve_loop:
            return "None"
        num_curves = curve_loop.NumberOfCurves()
        bbox = _get_curveloop_bbox(curve_loop)
        if bbox:
            min_x, max_x, min_y, max_y = bbox
            width = max_x - min_x
            height = max_y - min_y
            return "curves={}, width={:.3f}ft, height={:.3f}ft".format(num_curves, width, height)
        return "curves={}, no bbox".format(num_curves)
    except Exception as e:
        return "error: {}".format(e)


def _curveloops_to_solid(curve_loops, source_info=None):
    """Convert CurveLoop(s) to a solid for boolean operations.
    
    Args:
        curve_loops: Single CurveLoop or list of CurveLoops
        source_info: Optional string describing source for debug output
    
    Returns:
        Solid object or None if conversion fails
    """
    try:
        # Normalize to list
        if isinstance(curve_loops, DB.CurveLoop):
            curve_loops = [curve_loops]
        
        if not curve_loops:
            if DEBUG_VERBOSE:
                print("  [DEBUG] _curveloops_to_solid: No curve loops provided (source: {})".format(source_info or "unknown"))
            return None
        
        # Debug info about curve loops
        if DEBUG_VERBOSE:
            for i, cl in enumerate(curve_loops):
                info = _get_curveloop_info(cl)
                print("  [DEBUG] CurveLoop {}: {}".format(i, info))
        
        # Extrude to create solid
        solid = DB.GeometryCreationUtilities.CreateExtrusionGeometry(
            curve_loops,
            DB.XYZ(0, 0, 1),  # Extrusion direction (up)
            EXTRUSION_HEIGHT
        )
        return solid
    except Exception as e:
        if DEBUG_VERBOSE:
            print("  [DEBUG] Failed to convert CurveLoop(s) to solid: {} (source: {})".format(e, source_info or "unknown"))
        logger.debug("Failed to convert CurveLoop(s) to solid: %s", e)
        return None


def _area_to_solid(area_elem):
    """Convert an area element to a solid for boolean operations."""
    area_id = _get_element_id_value(area_elem.Id)
    area_name = "N/A"
    try:
        name_param = area_elem.LookupParameter("Name")
        if name_param:
            area_name = name_param.AsString() or "N/A"
    except:
        pass
    
    try:
        loops = _get_boundary_loops(area_elem)
        if not loops:
            error_msg = "No boundary loops found"
            if DEBUG_VERBOSE:
                print("  [DEBUG] Area {} ({}): {}".format(area_id, area_name, error_msg))
            debug_stats.failed_area_to_solid.append((area_id, area_name, error_msg, None))
            return None
        
        # Convert all loops to CurveLoops
        curve_loops = []
        boundary_info = []
        for i, loop in enumerate(loops):
            curve_loop = _boundary_loop_to_curveloop(loop)
            if curve_loop:
                curve_loops.append(curve_loop)
                boundary_info.append("Loop {}: {}".format(i, _get_curveloop_info(curve_loop)))
            else:
                boundary_info.append("Loop {}: FAILED to convert".format(i))
        
        if not curve_loops:
            error_msg = "All boundary loops failed to convert"
            if DEBUG_VERBOSE:
                print("  [DEBUG] Area {} ({}): {}".format(area_id, area_name, error_msg))
            debug_stats.failed_area_to_solid.append((area_id, area_name, error_msg, "; ".join(boundary_info)))
            return None
        
        source_info = "Area {}".format(area_id)
        solid = _curveloops_to_solid(curve_loops, source_info)
        
        if not solid:
            error_msg = "Extrusion geometry creation failed"
            if DEBUG_VERBOSE:
                print("  [DEBUG] Area {} ({}): {}".format(area_id, area_name, error_msg))
            debug_stats.failed_area_to_solid.append((area_id, area_name, error_msg, "; ".join(boundary_info)))
            return None
        
        # Track successful solid
        debug_stats.all_area_solids.append((area_id, solid))
        return solid
    except Exception as e:
        error_msg = str(e)
        if DEBUG_VERBOSE:
            print("  [DEBUG] Area {} ({}) exception: {}".format(area_id, area_name, error_msg))
        debug_stats.failed_area_to_solid.append((area_id, area_name, error_msg, None))
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
    
    if DEBUG_VERBOSE:
        print("\n[DEBUG] Creating union of {} areas...".format(len(areas)))
    
    # Convert all areas to solids with tracking
    solids = []
    area_id_map = {}  # Map solid index to area id for debug
    
    for area in areas:
        area_id = _get_element_id_value(area.Id)
        solid = _area_to_solid(area)
        if solid:
            area_id_map[len(solids)] = area_id
            solids.append(solid)
        else:
            if DEBUG_VERBOSE:
                print("  [DEBUG] Area {} failed to convert to solid".format(area_id))
    
    if not solids:
        if DEBUG_VERBOSE:
            print("  [DEBUG] No solids created from areas!")
        return None
    
    if DEBUG_VERBOSE:
        print("  [DEBUG] Successfully converted {} of {} areas to solids".format(len(solids), len(areas)))
    
    # Create union with detailed tracking
    try:
        union = solids[0]
        union_area_ids = [area_id_map.get(0, "unknown")]
        failed_unions = 0
        
        for i, solid in enumerate(solids[1:], start=1):
            area_id = area_id_map.get(i, "unknown")
            try:
                prev_volume = union.Volume if union else 0
                solid_volume = solid.Volume if solid else 0
                
                new_union = DB.BooleanOperationsUtils.ExecuteBooleanOperation(
                    union, solid, DB.BooleanOperationsType.Union
                )
                
                if new_union and new_union.Volume > MIN_VOLUME_THRESHOLD:
                    union = new_union
                    union_area_ids.append(area_id)
                    if DEBUG_VERBOSE:
                        print("  [DEBUG] Union {}/{}: Added area {}, volume: {:.3f} -> {:.3f}".format(
                            i, len(solids)-1, area_id, prev_volume, union.Volume))
                else:
                    error_msg = "Result volume too small or None"
                    if DEBUG_VERBOSE:
                        print("  [DEBUG] Union {}/{}: FAILED for area {} - {}".format(i, len(solids)-1, area_id, error_msg))
                    debug_stats.failed_boolean_ops.append(
                        ("Union of {} areas".format(len(union_area_ids)), 
                         "Area {}".format(area_id), 
                         error_msg))
                    failed_unions += 1
                    
            except Exception as e:
                error_msg = str(e)
                if DEBUG_VERBOSE:
                    print("  [DEBUG] Union {}/{}: EXCEPTION for area {} - {}".format(i, len(solids)-1, area_id, error_msg))
                debug_stats.failed_boolean_ops.append(
                    ("Union of {} areas".format(len(union_area_ids)), 
                     "Area {}".format(area_id), 
                     error_msg))
                failed_unions += 1
                logger.debug("Failed to union solid: %s", e)
                continue
        
        if DEBUG_VERBOSE:
            print("  [DEBUG] Union complete: {} successful, {} failed".format(
                len(solids) - failed_unions, failed_unions))
            if union:
                print("  [DEBUG] Final union volume: {:.3f} cubic ft".format(union.Volume))
        
        debug_stats.successful_solids.append(union)
        return union
    except Exception as e:
        if DEBUG_VERBOSE:
            print("  [DEBUG] Union creation failed: {}".format(e))
        logger.debug("Failed to create union: %s", e)
        return None


def _extract_holes_from_union(union_solid):
    """Extract interior holes from the union solid.
    
    Returns: List of CurveLoop objects representing holes
    """
    if not union_solid:
        if DEBUG_VERBOSE:
            print("  [DEBUG] _extract_holes_from_union: No union solid provided")
        return []
    
    try:
        holes = _get_loops_from_solid(union_solid, holes_only=True)
        
        if DEBUG_VERBOSE:
            print("\n[DEBUG] Extracted {} holes from union solid".format(len(holes)))
            for i, hole in enumerate(holes):
                info = _get_curveloop_info(hole)
                centroid = _get_curveloop_centroid(hole)
                if centroid:
                    print("  [DEBUG] Hole {}: {} | Centroid: ({:.2f}, {:.2f})".format(
                        i, info, centroid.X, centroid.Y))
                else:
                    print("  [DEBUG] Hole {}: {} | No centroid".format(i, info))
        
        # Track holes for visualization
        debug_stats.hole_loops.extend(holes)
        
        return holes
    except Exception as e:
        if DEBUG_VERBOSE:
            print("  [DEBUG] Failed to extract holes: {}".format(e))
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
    hole_info = _get_curveloop_info(hole_loop)
    
    if depth >= max_depth:
        error_msg = "Max recursion depth reached"
        if DEBUG_VERBOSE:
            print("  [DEBUG] Depth {}: {} - {}".format(depth, error_msg, hole_info))
        debug_stats.failed_holes.append((hole_info, None, error_msg))
        logger.debug("Max recursion depth reached")
        return 0
    
    # Get centroid of hole
    centroid = _get_curveloop_centroid(hole_loop)
    if not centroid:
        error_msg = "Failed to calculate centroid"
        if DEBUG_VERBOSE:
            print("  [DEBUG] Depth {}: {} - {}".format(depth, error_msg, hole_info))
        debug_stats.failed_holes.append((hole_info, None, error_msg))
        logger.debug("Depth %d: Failed to calculate centroid", depth)
        return 0
    
    if DEBUG_VERBOSE:
        print("  [DEBUG] Depth {}: Trying centroid at ({:.3f}, {:.3f}) for hole: {}".format(
            depth, centroid.X, centroid.Y, hole_info))
    
    logger.debug("Depth %d: Trying centroid at (%.3f, %.3f)", depth, centroid.X, centroid.Y)
    
    # Try creating area at centroid
    sub_txn = DB.SubTransaction(doc)
    try:
        sub_txn.Start()
        
        uv = DB.UV(centroid.X, centroid.Y)
        new_area = doc.Create.NewArea(view, uv)
        
        if not new_area:
            error_msg = "NewArea returned None"
            if DEBUG_VERBOSE:
                print("  [DEBUG] Depth {}: {} at ({:.3f}, {:.3f})".format(depth, error_msg, centroid.X, centroid.Y))
            debug_stats.failed_holes.append((hole_info, centroid, error_msg))
            sub_txn.RollBack()
            return 0
        
        # Regenerate to get boundaries
        try:
            doc.Regenerate()
        except Exception:
            pass
        
        if new_area.Area <= 0:
            error_msg = "Created area has zero or negative area"
            if DEBUG_VERBOSE:
                print("  [DEBUG] Depth {}: {} at ({:.3f}, {:.3f})".format(depth, error_msg, centroid.X, centroid.Y))
            debug_stats.failed_holes.append((hole_info, centroid, error_msg))
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


def _find_holes_using_2d(areas):
    """Find holes/gaps between areas using 2D polygon boolean operations.
    
    ALGORITHM:
    =========
    1. EXTRACT BOUNDARIES: For each area, get boundary segments and tessellate
       curves to (x,y) points. Identify exterior (largest) vs interior loops.
    
    2. CREATE POLYGONS: Build WPF Polygon2D objects from exterior point loops.
       Interior loops are donut holes = direct gaps to fill.
    
    3. FIND GAPS BETWEEN AREAS:
       - Create bounding box around all area polygons
       - Subtract each polygon from the bounding box
       - Remaining contours (except outer margin) = gaps
    
    4. PROCESS GAPS: For each gap contour, find an interior point (handles
       non-convex shapes) and return as centroid for area creation.
    
    Uses WPF geometry (System.Windows.Media) - IronPython 2.7 compatible.
    
    Args:
        areas: List of Area elements
    
    Returns:
        List of gap regions, each containing 'centroid' (x, y, z) and 'area' (sqft)
    """
    if not POLYGON_2D_AVAILABLE:
        if DEBUG_VERBOSE:
            print("  [2D] polygon_2d module not available, falling back to 3D")
        return []
    
    if not areas:
        return []
    
    if DEBUG_VERBOSE:
        print("\n[2D POLYGON ANALYSIS] Processing {} areas...".format(len(areas)))
    
    # ===========================================
    # SIMPLE ALGORITHM:
    # 1. Convert each area to Polygon2D (with holes)
    # 2. Union ALL polygons
    # 3. Interior loops in union = gaps to fill
    # ===========================================
    
    from polygon_2d import _find_interior_point
    
    area_polygons = []
    failed_areas = []
    
    # Step 1: Convert each area to Polygon2D (with its interior holes)
    for area in areas:
        area_id = _get_element_id_value(area.Id)
        try:
            loops = _get_boundary_loops(area)
            if not loops:
                failed_areas.append((area_id, "no boundary loops"))
                continue
            
            # Convert each loop to points
            loop_data = []
            for i, loop in enumerate(loops):
                points = _boundary_loop_to_points(loop)
                if points and len(points) >= 3:
                    area_val = abs(Polygon2D._calculate_contour_area(points))
                    loop_data.append({'points': points, 'area': area_val})
            
            if not loop_data:
                failed_areas.append((area_id, "no valid point loops"))
                continue
            
            # Sort by area descending - largest is exterior
            loop_data.sort(key=lambda x: x['area'], reverse=True)
            
            # Create polygon WITH holes
            ext_pts = loop_data[0]['points']
            hole_pts_list = [ld['points'] for ld in loop_data[1:] if ld['area'] >= 0.5]
            
            try:
                if hole_pts_list:
                    poly = Polygon2D.from_points_with_holes(ext_pts, hole_pts_list)
                else:
                    poly = Polygon2D(points=ext_pts)
                    
                if not poly.is_empty:
                    area_polygons.append(poly)
            except Exception as pe:
                if DEBUG_VERBOSE:
                    print("  [2D] Failed to create polygon for area {}: {}".format(area_id, pe))
                        
        except Exception as e:
            failed_areas.append((area_id, str(e)))
            continue
    
    if DEBUG_VERBOSE and failed_areas:
        print("  [2D] WARNING: {} areas failed".format(len(failed_areas)))
    
    if not area_polygons:
        if DEBUG_VERBOSE:
            print("  [2D] No valid area polygons")
        return []
    
    if DEBUG_VERBOSE:
        print("  [2D] Created {} area polygons".format(len(area_polygons)))
    
    # Step 2: Union ALL polygons
    union = Polygon2D.union_all(area_polygons)
    if union.is_empty:
        if DEBUG_VERBOSE:
            print("  [2D] Union is empty")
        return []
    
    if DEBUG_VERBOSE:
        print("  [2D] Union created successfully")
    
    # Step 3: Get contours from union - interior loops = holes
    all_contours = union.get_contours()
    if DEBUG_VERBOSE:
        print("  [2D] Found {} contours in union".format(len(all_contours)))
    
    # Calculate areas and sort - largest is exterior, rest are holes
    contour_data = []
    for contour in all_contours:
        if len(contour) >= 3:
            area_val = abs(Polygon2D._calculate_contour_area(contour))
            contour_data.append({'contour': contour, 'area': area_val})
    
    contour_data.sort(key=lambda x: x['area'], reverse=True)
    
    # Skip largest (exterior boundary), rest are holes/gaps
    all_gap_regions = []
    for i, cd in enumerate(contour_data):
        if i == 0:
            if DEBUG_VERBOSE:
                print("  [2D] Exterior boundary: area={:.2f} sqft".format(cd['area']))
            continue  # Skip exterior
        
        contour = cd['contour']
        area_val = cd['area']
        
        # Filter tiny regions
        if area_val < 0.5:
            continue
        
        # Find interior point
        interior_pt = _find_interior_point(contour, debug=False)
        if interior_pt:
            cx, cy = interior_pt
            all_gap_regions.append({
                'contour': contour,
                'centroid': (cx, cy, 0.0),
                'area': area_val,
                'source': 'union_hole'
            })
            if DEBUG_VERBOSE:
                print("  [2D] Hole: centroid=({:.2f}, {:.2f}), area={:.2f} sqft".format(cx, cy, area_val))
    
    if DEBUG_VERBOSE:
        print("  [2D] Total holes found: {}".format(len(all_gap_regions)))
    
    # DEBUG: Visualize the 2D geometry
    if DEBUG_VISUALIZATION:
        try:
            from polygon_2d import visualize_2d_geometry
            gap_contours = [r['contour'] for r in all_gap_regions if 'contour' in r]
            centroids = [(r['centroid'][0], r['centroid'][1]) for r in all_gap_regions if 'centroid' in r]
            visualize_2d_geometry(area_polygons, gap_contours, centroids, 
                                  title="2D: {} areas, {} holes".format(len(area_polygons), len(all_gap_regions)))
        except Exception as e:
            if DEBUG_VERBOSE:
                print("  [2D] Visualization failed: {}".format(e))
    
    return all_gap_regions


def _create_areas_at_gap_points(doc, view, gap_regions, usage_type_value, usage_type_name):
    """Create areas at detected gap region centroids.
    
    Args:
        doc: Revit document
        view: AreaPlan view
        gap_regions: List of dicts with 'centroid' and 'area' keys
        usage_type_value: Usage type value for created areas
        usage_type_name: Usage type name for created areas
    
    Returns:
        (created_count, failed_count, created_area_ids, failed_regions)
    """
    if not gap_regions:
        return 0, 0, [], []
    
    total_created = 0
    total_failed = 0
    created_area_ids = []
    failed_regions = []  # Track failures for visualization
    
    # Sort by area descending - fill larger gaps first
    sorted_regions = sorted(gap_regions, key=lambda r: r.get('area', 0), reverse=True)
    
    with revit.Transaction("Fill Gaps (2D)"):
        for region in sorted_regions:
            centroid = region.get('centroid')
            if not centroid:
                total_failed += 1
                continue
            
            cx, cy, cz = centroid
            
            # Use SubTransaction for each attempt
            sub_txn = DB.SubTransaction(doc)
            try:
                sub_txn.Start()
                
                uv = DB.UV(cx, cy)
                new_area = doc.Create.NewArea(view, uv)
                
                if not new_area:
                    if DEBUG_VERBOSE:
                        print("  [2D] Failed to create area at ({:.2f}, {:.2f})".format(cx, cy))
                    sub_txn.RollBack()
                    total_failed += 1
                    failed_regions.append(region)
                    continue
                
                # Regenerate to get area value
                try:
                    doc.Regenerate()
                except:
                    pass
                
                if new_area.Area <= 0:
                    if DEBUG_VERBOSE:
                        print("  [2D] Area at ({:.2f}, {:.2f}) has zero area - deleting".format(cx, cy))
                    doc.Delete(new_area.Id)
                    sub_txn.RollBack()
                    total_failed += 1
                    failed_regions.append(region)
                    continue
                
                # Set parameters
                _set_area_parameters(new_area, usage_type_value, usage_type_name)
                
                sub_txn.Commit()
                created_area_ids.append(new_area.Id)
                total_created += 1
                
                if DEBUG_VERBOSE:
                    print("  [2D] Created area at ({:.2f}, {:.2f}): {:.2f} sqm".format(
                        cx, cy, new_area.Area * SQFT_TO_SQM))
                
            except Exception as e:
                if DEBUG_VERBOSE:
                    print("  [2D] Exception creating area at ({:.2f}, {:.2f}): {}".format(cx, cy, e))
                try:
                    if sub_txn.HasStarted() and not sub_txn.HasEnded():
                        sub_txn.RollBack()
                except:
                    pass
                total_failed += 1
                failed_regions.append(region)
    
    return total_created, total_failed, created_area_ids, failed_regions


def _union_and_fill_holes(doc, view, areas_to_union, usage_type_value, usage_type_name, group_label=""):
    """Create union of areas, extract holes, and fill them.
    
    This is the core hole-filling method used by both modes:
    - All Gaps mode: Called once with all areas in view
    - Islands Only mode: Called per donut area with contained areas
    
    STRATEGY:
    1. PRIMARY: Use 2D polygon boolean operations (WPF-based)
       - More robust, handles complex geometry
       - No Revit 3D solid failures
    2. FALLBACK: Use 3D boolean union approach
       - Used if 2D module unavailable
       - May fail on complex geometry
    
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
    
    total_created = 0
    total_failed = 0
    all_created_ids = []
    
    # ========================================
    # PRIMARY APPROACH: 2D Polygon Operations
    # ========================================
    if POLYGON_2D_AVAILABLE:
        gap_regions = _find_holes_using_2d(areas_to_union)
        
        if gap_regions:
            created, failed, ids, failed_regions = _create_areas_at_gap_points(
                doc, view, gap_regions, usage_type_value, usage_type_name
            )
            total_created += created
            total_failed += failed
            all_created_ids.extend(ids)
            
            # DEBUG: Visualize failed gaps
            if DEBUG_VISUALIZATION and failed_regions:
                try:
                    from polygon_2d import visualize_2d_geometry
                    
                    # Get area polygons for context
                    area_polygons = []
                    for area in areas_to_union:
                        loops = _get_boundary_loops(area)
                        for loop in loops:
                            points = _boundary_loop_to_points(loop)
                            if points:
                                from polygon_2d import Polygon2D
                                poly = Polygon2D(points=points)
                                if not poly.is_empty:
                                    area_polygons.append(poly)
                                break  # Only exterior loop
                    
                    # Extract failed gap contours and centroids
                    failed_contours = [r['contour'] for r in failed_regions if 'contour' in r]
                    failed_centroids = [(r['centroid'][0], r['centroid'][1]) for r in failed_regions if 'centroid' in r]
                    
                    visualize_2d_geometry(
                        area_polygons, 
                        failed_contours, 
                        failed_centroids,
                        title="FAILED GAPS - {} failures (Red=Failed, Green=Centroid)".format(len(failed_regions))
                    )
                except Exception as e:
                    if DEBUG_VERBOSE:
                        print("  [2D] Failed gap visualization error: {}".format(e))
            
            if total_created > 0:
                # 2D approach succeeded
                return total_created, total_failed, all_created_ids
        
        if DEBUG_VERBOSE:
            print("  [2D] No gaps found with 2D analysis")
    
    # ========================================
    # FALLBACK: 3D Boolean Union Approach (DISABLED for testing)
    # ========================================
    # Commented out to isolate 2D solution testing
    # 
    # union_solid = _create_union_of_areas(areas_to_union)
    # if not union_solid:
    #     return total_created, total_failed, all_created_ids
    # 
    # # Extract holes from union
    # holes = _extract_holes_from_union(union_solid)
    # if not holes:
    #     return total_created, total_failed, all_created_ids
    # 
    # # Fill the holes using 3D approach
    # created_count, failed_count, created_area_ids = _create_areas_in_holes(
    #     doc, view, holes, usage_type_value, usage_type_name
    # )
    # 
    # total_created += created_count
    # total_failed += failed_count
    # all_created_ids.extend(created_area_ids)
    
    return total_created, total_failed, all_created_ids


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
# DEBUG VISUALIZATION (using dc3dserver)
# ============================================================

# Global server reference to keep it alive
_debug_server = None


def _curveloop_to_vertices(curve_loop, z_offset=0.0):
    """Convert a CurveLoop to a list of vertices for visualization."""
    vertices = []
    try:
        for curve in curve_loop:
            # Sample points along the curve
            start = curve.GetEndPoint(0)
            end = curve.GetEndPoint(1)
            vertices.append(DB.XYZ(start.X, start.Y, start.Z + z_offset))
            
            # For arcs and complex curves, add midpoints
            if not isinstance(curve, DB.Line):
                try:
                    mid = curve.Evaluate(0.5, True)
                    vertices.append(DB.XYZ(mid.X, mid.Y, mid.Z + z_offset))
                except:
                    pass
    except Exception as e:
        if DEBUG_VERBOSE:
            print("  [DEBUG] Failed to convert curve loop to vertices: {}".format(e))
    return vertices


def _create_edges_from_vertices(vertices, color):
    """Create Edge objects from a list of vertices for dc3dserver."""
    edges = []
    if len(vertices) < 2:
        return edges
    try:
        for i in range(len(vertices)):
            edges.append(revit.dc3dserver.Edge(vertices[i - 1], vertices[i], color))
    except Exception as e:
        if DEBUG_VERBOSE:
            print("  [DEBUG] Failed to create edges: {}".format(e))
    return edges


def _solid_to_edges(solid, color, z_offset=0.0):
    """Convert a solid's edges to dc3dserver Edge objects with color."""
    edges = []
    try:
        for edge in solid.Edges:
            try:
                curve = edge.AsCurve()
                if curve:
                    start = curve.GetEndPoint(0)
                    end = curve.GetEndPoint(1)
                    
                    # Apply z offset if specified
                    if z_offset != 0.0:
                        start = DB.XYZ(start.X, start.Y, start.Z + z_offset)
                        end = DB.XYZ(end.X, end.Y, end.Z + z_offset)
                    
                    # Create edge - try different approaches
                    try:
                        dc_edge = revit.dc3dserver.Edge(start, end, color)
                        edges.append(dc_edge)
                    except Exception as edge_err:
                        if DEBUG_VERBOSE and len(edges) == 0:
                            print("  [DEBUG] Edge creation error: {}".format(edge_err))
            except:
                continue
    except Exception as e:
        if DEBUG_VERBOSE:
            print("  [DEBUG] Failed to extract edges from solid: {}".format(e))
    return edges


def _mesh_from_solid_with_color(solid, color):
    """Create a dc3dserver Mesh from a solid with a custom color.
    
    Based on pyRevit's Mesh.from_solid but allows specifying color.
    """
    try:
        Triangle = revit.dc3dserver.Triangle
        Edge = revit.dc3dserver.Edge
        Mesh = revit.dc3dserver.Mesh
        
        triangles = []
        edges = []
        edge_color = DB.ColorWithTransparency(0, 0, 0, 0)  # Black edges
        
        for face in solid.Faces:
            face_mesh = face.Triangulate()
            triangle_count = face_mesh.NumTriangles
            
            for idx in range(triangle_count):
                mesh_triangle = face_mesh.get_Triangle(idx)
                
                # Get normal based on distribution type
                if face_mesh.DistributionOfNormals == DB.DistributionOfNormals.OnePerFace:
                    normal = face_mesh.GetNormal(0)
                elif face_mesh.DistributionOfNormals == DB.DistributionOfNormals.OnEachFacet:
                    normal = face_mesh.GetNormal(idx)
                elif face_mesh.DistributionOfNormals == DB.DistributionOfNormals.AtEachPoint:
                    normal = (
                        face_mesh.GetNormal(mesh_triangle.get_Index(0)) +
                        face_mesh.GetNormal(mesh_triangle.get_Index(1)) +
                        face_mesh.GetNormal(mesh_triangle.get_Index(2))
                    ).Normalize()
                else:
                    normal = Mesh.calculate_triangle_normal(
                        mesh_triangle.get_Vertex(0),
                        mesh_triangle.get_Vertex(1),
                        mesh_triangle.get_Vertex(2)
                    )
                
                triangles.append(Triangle(
                    mesh_triangle.get_Vertex(0),
                    mesh_triangle.get_Vertex(1),
                    mesh_triangle.get_Vertex(2),
                    normal,
                    color  # Use our custom color!
                ))
        
        # Add edges
        for edge in solid.Edges:
            pts = edge.Tessellate()
            for i in range(len(pts) - 1):
                edges.append(Edge(pts[i], pts[i + 1], edge_color))
        
        return Mesh(edges, triangles)
    except Exception as e:
        if DEBUG_VERBOSE:
            print("  [DEBUG] _mesh_from_solid_with_color failed: {}".format(e))
        return None


def _visualize_debug_geometry():
    """Visualize all solids and holes using dc3dserver with a persistent dialog."""
    global _debug_server
    
    try:
        # Check if dc3dserver is available
        if not hasattr(revit, 'dc3dserver'):
            print("[DEBUG] dc3dserver not available - skipping visualization")
            return
        
        # Clean up any previous visualization first
        try:
            old_server = revit.dc3dserver.Server(register=False)
            old_server.remove_server()
            revit.uidoc.RefreshActiveView()
        except:
            pass
        
        doc = revit.doc
        meshes = []
        edges = []
        
        # Color definitions (RGBA with transparency)
        COLOR_HOLE = DB.ColorWithTransparency(255, 0, 0, 100)  # Red for holes
        COLOR_AREA_SOLID = DB.ColorWithTransparency(0, 150, 255, 150)  # Blue for area solids
        COLOR_UNION_SOLID = DB.ColorWithTransparency(0, 255, 0, 180)  # Green for union result
        
        # Debug: print what we have to visualize
        print("\n[DEBUG] Geometry to visualize:")
        print("  - all_area_solids: {}".format(len(debug_stats.all_area_solids)))
        print("  - successful_solids: {}".format(len(debug_stats.successful_solids)))
        print("  - hole_loops: {}".format(len(debug_stats.hole_loops)))
        
        # Visualize all area solids as COLORED MESHES (blue)
        solid_count = 0
        for area_id, solid in debug_stats.all_area_solids:
            try:
                if solid and solid.Volume > MIN_VOLUME_THRESHOLD:
                    mesh = _mesh_from_solid_with_color(solid, COLOR_AREA_SOLID)
                    if mesh:
                        meshes.append(mesh)
                        solid_count += 1
                        if DEBUG_VERBOSE and solid_count == 1:
                            # Print first solid's bounding info
                            bbox = solid.GetBoundingBox()
                            if bbox:
                                print("  [DEBUG] First solid bbox: min=({:.1f},{:.1f},{:.1f}) max=({:.1f},{:.1f},{:.1f})".format(
                                    bbox.Min.X, bbox.Min.Y, bbox.Min.Z,
                                    bbox.Max.X, bbox.Max.Y, bbox.Max.Z))
            except Exception as e:
                if DEBUG_VERBOSE:
                    print("  [DEBUG] Failed to create mesh for area {}: {}".format(area_id, e))
        
        print("  - Created {} BLUE meshes from area solids".format(solid_count))
        
        # Visualize union result solids as COLORED MESHES (green)
        union_count = 0
        for solid in debug_stats.successful_solids:
            try:
                if solid and solid.Volume > MIN_VOLUME_THRESHOLD:
                    mesh = _mesh_from_solid_with_color(solid, COLOR_UNION_SOLID)
                    if mesh:
                        meshes.append(mesh)
                        union_count += 1
            except Exception as e:
                if DEBUG_VERBOSE:
                    print("  [DEBUG] Failed to create mesh for union solid: {}".format(e))
        
        print("  - Created {} GREEN meshes from union solids".format(union_count))
        
        # Visualize detected holes as edges
        z_offset = 0.5  # Slight offset to make visible above floor
        hole_count = 0
        for hole_loop in debug_stats.hole_loops:
            try:
                vertices = _curveloop_to_vertices(hole_loop, z_offset)
                if vertices:
                    edges.extend(_create_edges_from_vertices(vertices, COLOR_HOLE))
                    hole_count += 1
            except Exception as e:
                if DEBUG_VERBOSE:
                    print("  [DEBUG] Failed to visualize hole: {}".format(e))
        
        # Create and register server if we have anything to show
        if meshes or edges:
            print("\n[DEBUG VISUALIZATION]")
            print("  - BLUE meshes: {} area solids".format(solid_count))
            print("  - GREEN meshes: {} union result".format(union_count))
            print("  - RED edges: {} detected holes".format(hole_count))
            print("  - TOTAL: {} meshes, {} edges".format(len(meshes), len(edges)))
            
            try:
                # Create server and keep reference
                _debug_server = revit.dc3dserver.Server()
                _debug_server.meshes = meshes
                _debug_server.edges = edges
                
                print("\n  [DEBUG] Server created with {} meshes".format(len(meshes)))
                
                # Refresh view
                revit.uidoc.RefreshActiveView()
                
                # Also try updating all open views
                try:
                    for view_id in revit.uidoc.GetOpenUIViews():
                        try:
                            view_id.RefreshView()
                        except:
                            pass
                except:
                    pass
                
                print("\n  Visualization is now active in 3D views.")
                print("  Navigate to a 3D view to see the geometry.")
                
                # Show non-modal window to keep server alive
                _show_debug_visualization_window(_debug_server, solid_count, union_count, hole_count)
                
            except Exception as e:
                print("  [DEBUG] Failed to register visualization server: {}".format(e))
        else:
            print("\n[DEBUG VISUALIZATION] No geometry to visualize")
            
    except Exception as e:
        print("[DEBUG] Visualization failed: {}".format(e))


def _show_debug_visualization_window(server, solid_count, union_count, hole_count):
    """Show non-modal window to keep visualization alive."""
    from System.Windows import Window, WindowStartupLocation, ResizeMode, Thickness
    from System.Windows.Controls import Grid, RowDefinition, TextBlock, Button, StackPanel
    from System.Windows import GridLength, GridUnitType, FontWeights, HorizontalAlignment
    from System.Windows.Media import Brushes
    
    # Use list to hold references (closure-friendly)
    server_ref = [server]
    uidoc_ref = [revit.uidoc]  # Capture uidoc reference for closure
    
    try:
        # Create window programmatically
        window = Window()
        window.Title = "Fill Holes - Debug Visualization"
        window.Height = 220
        window.Width = 380
        window.WindowStartupLocation = WindowStartupLocation.CenterScreen
        window.Topmost = True
        window.ResizeMode = ResizeMode.NoResize
        
        # Create grid
        grid = Grid()
        grid.Margin = Thickness(15)
        
        # Row definitions
        grid.RowDefinitions.Add(RowDefinition())
        grid.RowDefinitions[0].Height = GridLength(1, GridUnitType.Auto)
        grid.RowDefinitions.Add(RowDefinition())
        grid.RowDefinitions[1].Height = GridLength(1, GridUnitType.Star)
        grid.RowDefinitions.Add(RowDefinition())
        grid.RowDefinitions[2].Height = GridLength(1, GridUnitType.Auto)
        
        # Title
        title = TextBlock()
        title.Text = "Debug Visualization Active"
        title.FontWeight = FontWeights.Bold
        title.FontSize = 14
        title.Margin = Thickness(0, 0, 0, 10)
        Grid.SetRow(title, 0)
        grid.Children.Add(title)
        
        # Info panel
        panel = StackPanel()
        Grid.SetRow(panel, 1)
        
        info = TextBlock()
        info.Text = (
            "BLUE = Area solids ({})\n"
            "GREEN = Union result ({})\n"
            "RED = Detected holes ({})"
        ).format(solid_count, union_count, hole_count)
        info.TextWrapping = System.Windows.TextWrapping.Wrap
        panel.Children.Add(info)
        
        hint = TextBlock()
        hint.Text = "Switch to a 3D view to see the geometry.\nClose this window to clear visualization."
        hint.TextWrapping = System.Windows.TextWrapping.Wrap
        hint.Foreground = Brushes.Gray
        hint.Margin = Thickness(0, 10, 0, 0)
        panel.Children.Add(hint)
        
        grid.Children.Add(panel)
        
        # Close button
        btn = Button()
        btn.Content = "Close Visualization"
        btn.Padding = Thickness(15, 8, 15, 8)
        btn.Margin = Thickness(0, 15, 0, 0)
        btn.HorizontalAlignment = HorizontalAlignment.Center
        Grid.SetRow(btn, 2)
        
        def on_click(s, e):
            window.Close()
        btn.Click += on_click
        
        grid.Children.Add(btn)
        
        window.Content = grid
        
        # Handle window close - clean up server using closure
        def on_closed(s, e):
            try:
                if server_ref[0]:
                    server_ref[0].remove_server()
                    uidoc_ref[0].RefreshActiveView()
                    print("[DEBUG] Visualization cleared.")
                    server_ref[0] = None
            except Exception as ex:
                print("[DEBUG] Error cleaning up: {}".format(ex))
        
        window.Closed += on_closed
        
        # Show non-modal
        window.Show()
        
    except Exception as e:
        # Fallback: just print instructions
        print("  [DEBUG] Could not create visualization window: {}".format(e))
        print("  Run this to clear visualization manually:")
        print("  revit.dc3dserver.Server().remove_server()")
        # Keep server alive by not removing it


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
    # Reset debug stats at start of each run
    debug_stats.reset()
    
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
    
    # Print debug diagnostic summary
    if DEBUG_VERBOSE:
        debug_stats.print_summary()
    
    # Visualize failed geometry if enabled
    if DEBUG_VISUALIZATION:
        _visualize_debug_geometry()


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
