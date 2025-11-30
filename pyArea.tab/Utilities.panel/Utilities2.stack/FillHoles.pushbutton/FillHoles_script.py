# -*- coding: utf-8 -*-
"""Fill Holes - Creates areas in gaps between existing areas.

Uses 2D polygon boolean operations to detect gaps:
1. Convert area boundaries to 2D polygons
2. Union all polygons
3. Interior loops in union = gaps to fill
4. Create areas at each gap's interior point
"""

__title__ = "Fill\nHoles"
__author__ = "pyArea"

from pyrevit import revit, DB, forms, script
import os
import sys

import clr
clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
import System
from System.Windows.Controls import CheckBox
from System.Windows.Media.Imaging import BitmapImage, BitmapCacheOption

SCRIPT_DIR = os.path.dirname(__file__)
LIB_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(SCRIPT_DIR))), "lib")
if LIB_DIR not in sys.path:
    sys.path.insert(0, LIB_DIR)

import data_manager
from polygon_2d import Polygon2D, _find_interior_point

logger = script.get_logger()

SQFT_TO_SQM = 0.09290304

# Global to track failed regions across views
_failed_regions = []


# ============================================================
# UTILITIES
# ============================================================

def get_element_id_value(element_id):
    """Get integer value from ElementId (Revit 2024-2026+ compatible)."""
    try:
        return element_id.IntegerValue
    except AttributeError:
        return int(element_id.Value)


def get_areas_in_view(doc, view):
    """Get all placed areas in the specified AreaPlan view."""
    if not hasattr(view, 'AreaScheme') or not view.AreaScheme:
        return []
    if not hasattr(view, 'GenLevel') or not view.GenLevel:
        return []
    
    target_scheme_id = view.AreaScheme.Id
    target_level_id = view.GenLevel.Id
    
    collector = DB.FilteredElementCollector(doc)
    collector = collector.OfCategory(DB.BuiltInCategory.OST_Areas)
    collector = collector.WhereElementIsNotElementType()
    
    areas = []
    for area in collector:
        try:
            if not hasattr(area, 'AreaScheme') or area.AreaScheme is None:
                continue
            if area.AreaScheme.Id != target_scheme_id:
                continue
            if area.LevelId != target_level_id:
                continue
            # Must have boundaries (placed)
            opts = DB.SpatialElementBoundaryOptions()
            loops = area.GetBoundarySegments(opts)
            if loops and len(list(loops)) > 0:
                areas.append(area)
        except Exception:
            continue
    return areas


def boundary_loop_to_points(boundary_loop):
    """Convert boundary segment loop to (x, y) points list."""
    if not boundary_loop:
        return None
    
    points = []
    for segment in boundary_loop:
        try:
            curve = segment.GetCurve()
            if curve:
                tessellated = curve.Tessellate()
                for i, pt in enumerate(tessellated):
                    if i < len(tessellated) - 1:
                        points.append((pt.X, pt.Y))
        except Exception:
            pass
    
    if len(points) < 3:
        return None
    
    # Remove duplicates
    cleaned = [points[0]]
    for p in points[1:]:
        if abs(p[0] - cleaned[-1][0]) > 0.001 or abs(p[1] - cleaned[-1][1]) > 0.001:
            cleaned.append(p)
    
    return cleaned if len(cleaned) >= 3 else None


def get_usage_type_for_municipality(doc, view):
    """Get usage type value and name based on municipality."""
    municipality, variant = data_manager.get_municipality_from_view(doc, view)
    
    if municipality == "Jerusalem":
        return "70", u"הורדה"
    elif municipality == "Common":
        return "700" if variant == "Gross" else "300", u"הורדה"
    elif municipality == "Tel-Aviv":
        return "-1", u"חלל"
    return None, None


def set_area_parameters(area_elem, usage_type_value, usage_type_name):
    """Set usage type parameters on an area element."""
    if not usage_type_value:
        return
    try:
        param = area_elem.LookupParameter("Usage Type")
        if param and not param.IsReadOnly:
            param.Set(usage_type_value)
        if usage_type_name:
            for param_name in ["Name", "Usage Type Name"]:
                p = area_elem.LookupParameter(param_name)
                if p and not p.IsReadOnly:
                    p.Set(usage_type_name)
    except Exception:
        pass


# ============================================================
# HOLE DETECTION (2D)
# ============================================================

def find_gaps_in_areas(areas):
    """Find gaps between areas using 2D polygon boolean operations.
    
    Algorithm:
    1. Convert each area to Polygon2D (exterior + holes)
    2. Union all polygons
    3. Interior loops in union = gaps
    
    Returns:
        List of dicts with 'centroid' (x, y, z) and 'area' (sqft)
    """
    if not areas:
        return []
    
    area_polygons = []
    
    for area in areas:
        try:
            opts = DB.SpatialElementBoundaryOptions()
            loops = list(area.GetBoundarySegments(opts))
            if not loops:
                continue
            
            # Convert loops to point lists
            loop_data = []
            for loop in loops:
                points = boundary_loop_to_points(loop)
                if points and len(points) >= 3:
                    area_val = abs(Polygon2D._calculate_contour_area(points))
                    loop_data.append({'points': points, 'area': area_val})
            
            if not loop_data:
                continue
            
            # Largest = exterior, rest = holes
            loop_data.sort(key=lambda x: x['area'], reverse=True)
            ext_pts = loop_data[0]['points']
            hole_pts = [ld['points'] for ld in loop_data[1:] if ld['area'] >= 0.5]
            
            if hole_pts:
                poly = Polygon2D.from_points_with_holes(ext_pts, hole_pts)
            else:
                poly = Polygon2D(points=ext_pts)
            
            if not poly.is_empty:
                area_polygons.append(poly)
                
        except Exception:
            continue
    
    if not area_polygons:
        return []
    
    # Union all polygons
    union = Polygon2D.union_all(area_polygons)
    if union.is_empty:
        return []
    
    # Get contours - largest is exterior, rest are holes/gaps
    contours = union.get_contours()
    if len(contours) < 2:
        return []  # No interior holes
    
    contour_data = []
    for contour in contours:
        if len(contour) >= 3:
            area_val = abs(Polygon2D._calculate_contour_area(contour))
            contour_data.append({'contour': contour, 'area': area_val})
    
    contour_data.sort(key=lambda x: x['area'], reverse=True)
    
    # Skip exterior (largest), process interior holes
    gap_regions = []
    for cd in contour_data[1:]:
        contour = cd['contour']
        area_val = cd['area']
        
        if area_val < 0.5:  # Skip tiny regions
            continue
        
        interior_pt = _find_interior_point(contour)
        if interior_pt:
            gap_regions.append({
                'centroid': (interior_pt[0], interior_pt[1], 0.0),
                'area': area_val,
                'contour': contour  # Store for visualization
            })
    
    return gap_regions


# ============================================================
# DONUT MODE HELPERS
# ============================================================

def identify_donut_areas(areas):
    """Find areas that have interior holes (donuts).
    
    Returns:
        List of (area_elem, hole_count) tuples
    """
    donuts = []
    for area in areas:
        try:
            opts = DB.SpatialElementBoundaryOptions()
            loops = list(area.GetBoundarySegments(opts))
            if len(loops) > 1:
                donuts.append((area, len(loops) - 1))
        except Exception:
            continue
    return donuts


def point_in_polygon_2d(px, py, polygon_points):
    """Ray-casting algorithm to check if point is inside polygon."""
    n = len(polygon_points)
    inside = False
    j = n - 1
    for i in range(n):
        xi, yi = polygon_points[i]
        xj, yj = polygon_points[j]
        if ((yi > py) != (yj > py)) and (px < (xj - xi) * (py - yi) / (yj - yi) + xi):
            inside = not inside
        j = i
    return inside


def get_areas_inside_donut(all_areas, donut_area):
    """Find areas whose centroids are inside the donut's exterior boundary."""
    try:
        opts = DB.SpatialElementBoundaryOptions()
        loops = list(donut_area.GetBoundarySegments(opts))
        if not loops:
            return []
        
        exterior_points = boundary_loop_to_points(loops[0])
        if not exterior_points:
            return []
        
        contained = []
        for area in all_areas:
            if area.Id == donut_area.Id:
                continue
            try:
                loc = area.Location
                if loc and hasattr(loc, 'Point'):
                    pt = loc.Point
                    if point_in_polygon_2d(pt.X, pt.Y, exterior_points):
                        contained.append(area)
            except Exception:
                continue
        return contained
    except Exception:
        return []


# ============================================================
# AREA CREATION
# ============================================================

def create_areas_at_gaps(doc, view, gap_regions, usage_type_value, usage_type_name):
    """Create new areas at detected gap centroids.
    
    Returns:
        (created_count, failed_count, created_ids)
    """
    global _failed_regions
    
    if not gap_regions:
        return 0, 0, []
    
    created_count = 0
    failed_count = 0
    created_ids = []
    
    # Sort by area descending (fill larger gaps first)
    sorted_regions = sorted(gap_regions, key=lambda r: r.get('area', 0), reverse=True)
    
    with revit.Transaction("Fill Holes"):
        for region in sorted_regions:
            centroid = region.get('centroid')
            if not centroid:
                failed_count += 1
                region['view'] = view
                _failed_regions.append(region)
                continue
            
            cx, cy, _ = centroid
            
            sub_txn = DB.SubTransaction(doc)
            try:
                sub_txn.Start()
                
                uv = DB.UV(cx, cy)
                new_area = doc.Create.NewArea(view, uv)
                
                if not new_area:
                    sub_txn.RollBack()
                    failed_count += 1
                    region['view'] = view
                    _failed_regions.append(region)
                    continue
                
                doc.Regenerate()
                
                if new_area.Area <= 0:
                    doc.Delete(new_area.Id)
                    sub_txn.RollBack()
                    failed_count += 1
                    region['view'] = view
                    _failed_regions.append(region)
                    continue
                
                set_area_parameters(new_area, usage_type_value, usage_type_name)
                sub_txn.Commit()
                
                created_ids.append(new_area.Id)
                created_count += 1
                
            except Exception:
                try:
                    if sub_txn.HasStarted() and not sub_txn.HasEnded():
                        sub_txn.RollBack()
                except Exception:
                    pass
                failed_count += 1
                region['view'] = view
                _failed_regions.append(region)
    
    return created_count, failed_count, created_ids


# ============================================================
# VIEW SELECTION DIALOG
# ============================================================

class ViewSelectionDialog(forms.WPFWindow):
    """Dialog for selecting AreaPlan views and fill mode."""
    
    def __init__(self, area_schemes, selected_index=0, preselected_ids=None):
        forms.WPFWindow.__init__(self, 'ViewSelectionDialog.xaml')
        
        self._doc = revit.doc
        self._area_schemes = area_schemes
        self._preselected_ids = preselected_ids or set()
        self._checkboxes = {}
        
        self.btn_ok.Click += self._on_ok
        self.btn_cancel.Click += self._on_cancel
        self.btn_select_all.Click += self._on_select_all
        self.combo_areascheme.SelectionChanged += self._on_scheme_changed
        
        for scheme, municipality in area_schemes:
            self.combo_areascheme.Items.Add("{} ({})".format(scheme.Name, municipality))
        
        if area_schemes:
            self.combo_areascheme.SelectedIndex = selected_index
        
        # Hide scheme selector if only one municipality
        municipalities = set(m for _, m in area_schemes if m)
        if len(municipalities) <= 1:
            self.combo_areascheme.Visibility = System.Windows.Visibility.Collapsed
        
        self._load_mode_icons()
        self.selected_views = None
        self.only_donut_holes = False
    
    def _on_scheme_changed(self, sender, args):
        if self.combo_areascheme.SelectedIndex < 0:
            return
        self._populate_views()
    
    def _populate_views(self):
        self.panel_views.Children.Clear()
        self._checkboxes = {}
        
        if self.combo_areascheme.SelectedIndex < 0:
            return
        
        scheme, _ = self._area_schemes[self.combo_areascheme.SelectedIndex]
        
        collector = DB.FilteredElementCollector(self._doc).OfClass(DB.View)
        views = []
        
        for view in collector:
            try:
                if not hasattr(view, 'AreaScheme') or not view.AreaScheme:
                    continue
                if view.AreaScheme.Id != scheme.Id:
                    continue
                areas = get_areas_in_view(self._doc, view)
                if areas:
                    views.append(view)
            except Exception:
                continue
        
        views.sort(key=lambda v: v.Origin.Z if hasattr(v, 'Origin') else 0)
        
        if not views:
            tb = System.Windows.Controls.TextBlock()
            tb.Text = "No Area Plan views with placed areas found."
            tb.Foreground = System.Windows.Media.Brushes.Gray
            self.panel_views.Children.Add(tb)
            return
        
        for view in views:
            level_name = view.GenLevel.Name if hasattr(view, 'GenLevel') and view.GenLevel else "N/A"
            
            cb = CheckBox()
            cb.Content = "{} (Level: {})".format(view.Name, level_name)
            cb.Margin = System.Windows.Thickness(5, 3, 5, 3)
            cb.Tag = get_element_id_value(view.Id)
            cb.IsChecked = view.Id in self._preselected_ids
            
            self.panel_views.Children.Add(cb)
            self._checkboxes[cb.Tag] = cb
    
    def _on_select_all(self, sender, args):
        for cb in self._checkboxes.values():
            cb.IsChecked = True
    
    def _on_ok(self, sender, args):
        if self.combo_areascheme.SelectedIndex < 0:
            forms.alert("Please select an area scheme.", exitscript=False)
            return
        
        selected_ids = [vid for vid, cb in self._checkboxes.items() if cb.IsChecked]
        if not selected_ids:
            forms.alert("Please select at least one view.", exitscript=False)
            return
        
        self.selected_views = []
        for vid in selected_ids:
            view = self._doc.GetElement(DB.ElementId(System.Int64(int(vid))))
            if view:
                self.selected_views.append(view)
        
        self.only_donut_holes = bool(self.rb_fill_donut_holes.IsChecked)
        self.DialogResult = True
        self.Close()
    
    def _on_cancel(self, sender, args):
        self.DialogResult = False
        self.Close()
    
    def _load_mode_icons(self):
        icons = [
            (getattr(self, "img_mode_all_holes", None), "FillHolesIcon_split.png"),
            (getattr(self, "img_mode_donut_holes", None), "FillHolesIcon.png"),
        ]
        for img, filename in icons:
            if img is None:
                continue
            try:
                path = os.path.join(SCRIPT_DIR, filename)
                if os.path.exists(path):
                    bmp = BitmapImage()
                    bmp.BeginInit()
                    bmp.UriSource = System.Uri(path)
                    bmp.CacheOption = BitmapCacheOption.OnLoad
                    bmp.EndInit()
                    img.Source = bmp
            except Exception:
                pass


# ============================================================
# FAILED GAPS NAVIGATOR
# ============================================================

class FailedGapsNavigator(object):
    """Non-modal dialog to navigate between failed gap regions with ghost geometry."""
    
    def __init__(self, failed_regions):
        self._regions = failed_regions
        self._current_index = 0
        self._window = None
        self._uidoc = revit.uidoc
        self._doc = revit.doc
        
        if not self._regions:
            return
        
        self._create_window()
        self._show_current_failure()
    
    def _create_window(self):
        """Create the navigation window programmatically."""
        from System.Windows import Window, WindowStartupLocation, ResizeMode, Thickness
        from System.Windows import HorizontalAlignment, VerticalAlignment, FontWeights
        from System.Windows.Controls import Grid, RowDefinition, ColumnDefinition, TextBlock, Button, StackPanel
        from System.Windows import GridLength, GridUnitType
        from System.Windows.Media import Brushes
        
        window = Window()
        window.Title = "Failed Gaps Navigator"
        window.Width = 340
        window.Height = 220
        window.WindowStartupLocation = WindowStartupLocation.Manual
        window.Left = 50
        window.Top = 100
        window.Topmost = True
        window.ResizeMode = ResizeMode.NoResize
        
        grid = Grid()
        grid.Margin = Thickness(15)
        
        # Rows
        for _ in range(3):
            grid.RowDefinitions.Add(RowDefinition())
        grid.RowDefinitions[0].Height = GridLength(1, GridUnitType.Auto)
        grid.RowDefinitions[1].Height = GridLength(1, GridUnitType.Star)
        grid.RowDefinitions[2].Height = GridLength(1, GridUnitType.Auto)
        
        # Title
        title = TextBlock()
        title.Text = "Failed Gap Locations"
        title.FontWeight = FontWeights.Bold
        title.FontSize = 14
        title.Margin = Thickness(0, 0, 0, 10)
        Grid.SetRow(title, 0)
        grid.Children.Add(title)
        
        # Info panel
        info_panel = StackPanel()
        Grid.SetRow(info_panel, 1)
        
        self._counter_text = TextBlock()
        self._counter_text.FontSize = 16
        self._counter_text.HorizontalAlignment = HorizontalAlignment.Center
        self._counter_text.Margin = Thickness(0, 5, 0, 5)
        info_panel.Children.Add(self._counter_text)
        
        self._info_text = TextBlock()
        self._info_text.FontSize = 11
        self._info_text.Foreground = Brushes.Gray
        self._info_text.HorizontalAlignment = HorizontalAlignment.Center
        self._info_text.TextWrapping = System.Windows.TextWrapping.Wrap
        info_panel.Children.Add(self._info_text)
        
        self._error_text = TextBlock()
        self._error_text.FontSize = 10
        self._error_text.Foreground = Brushes.Red
        self._error_text.HorizontalAlignment = HorizontalAlignment.Center
        self._error_text.TextWrapping = System.Windows.TextWrapping.Wrap
        self._error_text.Margin = Thickness(0, 5, 0, 0)
        info_panel.Children.Add(self._error_text)
        
        grid.Children.Add(info_panel)
        
        # Navigation buttons
        btn_grid = Grid()
        Grid.SetRow(btn_grid, 2)
        btn_grid.Margin = Thickness(0, 10, 0, 0)
        
        for _ in range(3):
            btn_grid.ColumnDefinitions.Add(ColumnDefinition())
        
        btn_prev = Button()
        btn_prev.Content = u"◀ Prev"
        btn_prev.Width = 80
        btn_prev.Height = 30
        btn_prev.Click += self._on_prev
        Grid.SetColumn(btn_prev, 0)
        btn_grid.Children.Add(btn_prev)
        
        btn_close = Button()
        btn_close.Content = "Close"
        btn_close.Width = 80
        btn_close.Height = 30
        btn_close.Click += self._on_close
        Grid.SetColumn(btn_close, 1)
        btn_grid.Children.Add(btn_close)
        
        btn_next = Button()
        btn_next.Content = u"Next ▶"
        btn_next.Width = 80
        btn_next.Height = 30
        btn_next.Click += self._on_next
        Grid.SetColumn(btn_next, 2)
        btn_grid.Children.Add(btn_next)
        
        grid.Children.Add(btn_grid)
        
        window.Content = grid
        window.Closed += self._on_window_closed
        
        self._window = window
    
    def _on_prev(self, sender, args):
        try:
            self._error_text.Text = ""  # Clear previous errors
            if self._current_index > 0:
                self._current_index -= 1
            else:
                self._current_index = len(self._regions) - 1
            self._show_current_failure()
        except Exception as e:
            self._error_text.Text = "Error: {}".format(str(e))
    
    def _on_next(self, sender, args):
        try:
            print("Next button clicked - current index: {}".format(self._current_index))
            self._error_text.Text = ""  # Clear previous errors
            if self._current_index < len(self._regions) - 1:
                self._current_index += 1
            else:
                self._current_index = 0
            print("New index: {}".format(self._current_index))
            self._show_current_failure()
        except Exception as e:
            error_msg = "Error: {}".format(str(e))
            print(error_msg)
            self._error_text.Text = error_msg
    
    def _on_close(self, sender, args):
        self._window.Close()
    
    def _on_window_closed(self, sender, args):
        pass  # Nothing to clean up
    
    def _show_current_failure(self):
        """Navigate to and highlight current failed region."""
        try:
            print("_show_current_failure called")
            if not self._regions:
                print("No regions!")
                return
            
            region = self._regions[self._current_index]
            centroid = region.get('centroid')
            area_sqft = region.get('area', 0)
            view = region.get('view')
            contour = region.get('contour')
            
            print("Region data - view: {}, centroid: {}, area: {}".format(
                view.Name if view else "None", 
                centroid if centroid else "None",
                area_sqft
            ))
            
            # Update UI
            self._counter_text.Text = "{} / {}".format(self._current_index + 1, len(self._regions))
            
            # Convert sqft to sqm (0.09290304 conversion factor)
            area_sqm = area_sqft * 0.09290304
            info = "Area: {:.2f} sqm".format(area_sqm)
            if view:
                info += "\nView: {}".format(view.Name)
            else:
                info += "\nView: None"
            if centroid:
                info += "\nPoint: ({:.1f}, {:.1f})".format(centroid[0], centroid[1])
            else:
                info += "\nPoint: None"
            self._info_text.Text = info
            
            # Switch to the view and zoom
            if view and centroid:
                print("Calling _zoom_to_point...")
                self._zoom_to_point(view, centroid, contour)
            else:
                error_msg = "Missing view or centroid"
                print(error_msg)
                self._error_text.Text = error_msg
            
            # Ghost geometry disabled - doesn't work from modeless dialogs
            # The zoom itself should be sufficient to locate the failure
            print("Done with _show_current_failure")
            
        except Exception as e:
            error_msg = "Error: {}".format(str(e))
            print(error_msg)
            self._error_text.Text = error_msg
    
    def _zoom_to_point(self, view, centroid, contour=None):
        """Zoom the view to center on the centroid."""
        try:
            from Autodesk.Revit import DB as RevitDB
            
            print("_zoom_to_point: Switching to view {}".format(view.Name))
            # Switch to the view
            self._uidoc.ActiveView = view
            print("Active view set successfully")
            
            # Calculate zoom extents from contour or use default
            if contour and len(contour) >= 3:
                min_x = min(p[0] for p in contour)
                max_x = max(p[0] for p in contour)
                min_y = min(p[1] for p in contour)
                max_y = max(p[1] for p in contour)
                margin = 5.0  # feet margin
                print("Using contour bounds: ({:.1f},{:.1f}) to ({:.1f},{:.1f})".format(min_x, min_y, max_x, max_y))
            else:
                cx, cy, _ = centroid
                margin = 10.0
                min_x, max_x = cx - margin, cx + margin
                min_y, max_y = cy - margin, cy + margin
                print("Using centroid bounds: ({:.1f},{:.1f}) to ({:.1f},{:.1f})".format(min_x, min_y, max_x, max_y))
            
            # Get view's Z level
            z = 0
            if hasattr(view, 'GenLevel') and view.GenLevel:
                z = view.GenLevel.Elevation
            print("View Z level: {}".format(z))
            
            # Create bounding box for zoom
            min_pt = RevitDB.XYZ(min_x - margin, min_y - margin, z - 1)
            max_pt = RevitDB.XYZ(max_x + margin, max_y + margin, z + 1)
            
            # Get all open UIViews and find the one matching our view
            ui_views = self._uidoc.GetOpenUIViews()
            print("Found {} open UIViews".format(len(list(ui_views))))
            
            target_ui_view = None
            for ui_view in ui_views:
                if ui_view.ViewId == view.Id:
                    target_ui_view = ui_view
                    print("Found matching UIView")
                    break
            
            if target_ui_view:
                print("Calling ZoomAndCenterRectangle...")
                target_ui_view.ZoomAndCenterRectangle(min_pt, max_pt)
                print("Zoom completed")
            else:
                error_msg = "View not found in open views"
                print(error_msg)
                self._error_text.Text = error_msg
            
        except Exception as e:
            error_msg = "Zoom failed: {}".format(str(e))
            print(error_msg)
            self._error_text.Text = error_msg
    
    
    def show(self):
        """Show the non-modal dialog."""
        if self._window:
            self._window.Show()


# ============================================================
# MAIN
# ============================================================

def get_defined_area_schemes():
    """Get area schemes with municipality defined."""
    doc = revit.doc
    collector = DB.FilteredElementCollector(doc).OfClass(DB.AreaScheme)
    
    defined = []
    for scheme in collector:
        municipality = data_manager.get_municipality(scheme)
        if municipality:
            defined.append((scheme, municipality))
    return defined


def get_context_element():
    """Get context from selection or active view."""
    try:
        doc = revit.doc
        uidoc = revit.uidoc
        selection = uidoc.Selection.GetElementIds()
        
        if selection:
            views = []
            sheets = []
            for eid in selection:
                elem = doc.GetElement(eid)
                if isinstance(elem, DB.Viewport):
                    view = doc.GetElement(elem.ViewId)
                    if hasattr(view, 'AreaScheme') and view.AreaScheme:
                        if data_manager.get_municipality(view.AreaScheme):
                            views.append(view)
                elif isinstance(elem, DB.View) and not isinstance(elem, DB.ViewSheet):
                    if hasattr(elem, 'AreaScheme') and elem.AreaScheme:
                        if data_manager.get_municipality(elem.AreaScheme):
                            views.append(elem)
                elif isinstance(elem, DB.ViewSheet):
                    sheets.append(elem)
            
            if views:
                return views, "views"
            if sheets:
                return sheets[0], "sheet"
        
        active = uidoc.ActiveView
        if isinstance(active, DB.ViewSheet):
            return active, "sheet"
        if hasattr(active, 'AreaScheme') and active.AreaScheme:
            if data_manager.get_municipality(active.AreaScheme):
                return [active], "views"
    except Exception:
        pass
    return None, None


def get_views_on_sheet(sheet):
    """Get AreaPlan views placed on a sheet."""
    views = []
    try:
        for vid in sheet.GetAllPlacedViews():
            view = revit.doc.GetElement(vid)
            if hasattr(view, 'AreaScheme') and view.AreaScheme:
                views.append(view)
    except Exception:
        pass
    return views


def process_view(doc, view, output, only_donut_holes):
    """Process a single view and fill holes.
    
    Returns:
        (created_count, failed_count)
    """
    areas = get_areas_in_view(doc, view)
    if not areas:
        return 0, 0
    
    view_link = output.linkify(view.Id)
    print("-" * 60)
    print("{} {}".format(view_link, view.Name))
    
    usage_value, usage_name = get_usage_type_for_municipality(doc, view)
    
    total_created = 0
    total_failed = 0
    all_created_ids = []
    
    if only_donut_holes:
        # Islands Only mode: process each donut individually
        donuts = identify_donut_areas(areas)
        if not donuts:
            print("  No donut areas found")
            return 0, 0
        
        for donut_area, hole_count in donuts:
            contained = get_areas_inside_donut(areas, donut_area)
            areas_to_process = [donut_area] + contained
            
            gaps = find_gaps_in_areas(areas_to_process)
            if gaps:
                created, failed, ids = create_areas_at_gaps(
                    doc, view, gaps, usage_value, usage_name
                )
                total_created += created
                total_failed += failed
                all_created_ids.extend(ids)
    else:
        # All Gaps mode: process all areas together
        gaps = find_gaps_in_areas(areas)
        if gaps:
            created, failed, ids = create_areas_at_gaps(
                doc, view, gaps, usage_value, usage_name
            )
            total_created = created
            total_failed = failed
            all_created_ids = ids
    
    # Print created areas
    for eid in all_created_ids:
        try:
            link = output.linkify(eid)
            area_elem = doc.GetElement(eid)
            sqm = area_elem.Area * SQFT_TO_SQM if area_elem else 0
            output.print_html("&nbsp;&nbsp;{} ({:.2f} sqm)".format(link, sqm))
        except Exception:
            pass
    
    return total_created, total_failed


def fill_holes():
    """Main entry point."""
    global _failed_regions
    _failed_regions = []  # Reset failures
    
    doc = revit.doc
    
    # Get defined area schemes
    schemes = get_defined_area_schemes()
    if not schemes:
        forms.alert(
            "No area schemes with municipality defined.\n\n"
            "Please define an area scheme using the Calculation Setup tool first.",
            exitscript=True
        )
    
    # Determine default selection from context
    context, context_type = get_context_element()
    selected_index = 0
    preselected_ids = set()
    
    if context:
        if context_type == "views":
            first_scheme = context[0].AreaScheme
            for i, (scheme, _) in enumerate(schemes):
                if scheme.Id == first_scheme.Id:
                    selected_index = i
                    break
            preselected_ids = set(v.Id for v in context)
        elif context_type == "sheet":
            views = get_views_on_sheet(context)
            if views:
                first_scheme = views[0].AreaScheme
                for i, (scheme, _) in enumerate(schemes):
                    if scheme.Id == first_scheme.Id:
                        selected_index = i
                        break
                preselected_ids = set(v.Id for v in views)
    
    # Show dialog
    dialog = ViewSelectionDialog(schemes, selected_index, preselected_ids)
    if not dialog.ShowDialog():
        return
    
    selected_views = dialog.selected_views
    if not selected_views:
        return
    
    only_donut_holes = dialog.only_donut_holes
    
    # Sort by elevation
    selected_views.sort(key=lambda v: v.Origin.Z if hasattr(v, 'Origin') else 0)
    
    # Process views
    output = script.get_output()
    total_created = 0
    total_failed = 0
    
    for view in selected_views:
        created, failed = process_view(doc, view, output, only_donut_holes)
        total_created += created
        total_failed += failed
    
    print("\nTotal: {} area(s) created | {} failed".format(total_created, total_failed))
    
    # Show navigator for failed regions
    if _failed_regions:
        navigator = FailedGapsNavigator(_failed_regions)
        navigator.show()


if __name__ == '__main__':
    fill_holes()
