# -*- coding: utf-8 -*-
"""Detect Bad Boundaries - Identifies areas with invalid boundary loops.

Checks each area's boundary segments to detect gaps where curves don't connect.
Displays an interactive visualization window with:
- Red circles marking gap locations (turn green when fixed)
- TreeView navigation between gaps across views
- Auto-detection of fixes when using trim/modify tools
"""

__title__ = "Detect\nBad Boundaries"
__author__ = "pyArea"
__persistentengine__ = True

import os
import sys
import math

from pyrevit import revit, DB, forms, HOST_APP, UI

# Import InvalidOperationException for proper error handling in ExternalEvent
try:
    from Autodesk.Revit.Exceptions import InvalidOperationException
except ImportError:
    InvalidOperationException = Exception

import clr
clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')
clr.AddReference('System.Windows.Forms')
import System
from System.Windows.Controls import CheckBox
from System.Windows.Forms import SendKeys
from System.Threading import Thread, ThreadStart

# Document references
doc = HOST_APP.doc
uidoc = HOST_APP.uidoc

# Add lib path for data_manager
SCRIPT_DIR = os.path.dirname(__file__)
LIB_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(SCRIPT_DIR))), "lib")
if LIB_DIR not in sys.path:
    sys.path.insert(0, LIB_DIR)

import data_manager

# Tolerance for point matching (feet) - Revit's internal tolerance is ~1/256 inch
POINT_TOLERANCE = 0.001



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


def points_equal(pt1, pt2, tolerance=POINT_TOLERANCE):
    """Check if two XYZ points are equal within tolerance."""
    return (abs(pt1.X - pt2.X) < tolerance and 
            abs(pt1.Y - pt2.Y) < tolerance and 
            abs(pt1.Z - pt2.Z) < tolerance)


def gap_point_key(pt):
    """Create a hashable key from an XYZ point for gap grouping."""
    # Round to avoid floating point issues, but gaps from same boundary will be identical
    return (round(pt.X, 6), round(pt.Y, 6), round(pt.Z, 6))


# ============================================================
# GAP CLASS
# ============================================================

class Gap(object):
    """Represents a gap between two area boundary segments.
    
    A gap occurs where two boundary elements fail to connect, breaking
    an area's boundary loop. Each gap has:
    - Exactly 1-2 boundary elements forming the gap
    - 1-2 areas affected by this gap
    - A specific location (center point between disconnected endpoints)
    - A length (distance between disconnected endpoints)
    """
    
    # Tolerance for matching gap locations (feet)
    # Using 0.5 ft (~15 cm) to account for gap center shifting when boundaries are modified
    LOCATION_TOLERANCE = 0.5
    
    def __init__(self, center, length_ft, boundary_element_ids, area_ids, view_id):
        """
        Args:
            center: XYZ point at gap center
            length_ft: Gap length in feet
            boundary_element_ids: Set of integer element IDs (1-2 elements)
            area_ids: Set of integer area IDs (1-2 areas)
            view_id: ElementId of the AreaPlan view
        """
        self.center = center
        self.length_ft = length_ft
        self.length_cm = length_ft * 30.48
        self.boundary_element_ids = set(boundary_element_ids) if boundary_element_ids else set()
        self.area_ids = set(area_ids) if area_ids else set()
        self.view_id = view_id
        self.fixed = False
        self.index = -1  # Set when added to gap list
    
    def contains_boundary_element(self, element_id_int):
        """Check if this gap involves the given boundary element."""
        return element_id_int in self.boundary_element_ids
    
    def contains_area(self, area_id_int):
        """Check if this gap affects the given area."""
        return area_id_int in self.area_ids
    
    def is_near_location(self, point, tolerance=None):
        """Check if a point is near this gap's center."""
        if tolerance is None:
            tolerance = self.LOCATION_TOLERANCE
        try:
            return self.center.DistanceTo(point) < tolerance
        except Exception:
            return False
    
    def recheck(self, doc):
        """Recheck if this gap still exists by examining its areas.
        
        Returns True if the gap is now fixed (no longer present).
        """
        if self.fixed:
            return False
        
        if not self.area_ids or not self.boundary_element_ids:
            return False
        
        areas_checked = 0
        
        for area_id in self.area_ids:
            try:
                elem_id = DB.ElementId(System.Int64(int(area_id)))
                area = doc.GetElement(elem_id)
            except Exception:
                continue
            
            if not area:
                continue
            
            try:
                result = check_area_boundaries(area)
            except Exception:
                continue
            
            areas_checked += 1
            gap_data = result.get('gap_data') or []
            
            for gap_info in gap_data:
                gap_elements = gap_info.get('elements') or []
                gap_element_ids = set()
                for eid in gap_elements:
                    if eid and eid != DB.ElementId.InvalidElementId:
                        gap_element_ids.add(get_element_id_value(eid))
                
                if self.boundary_element_ids == gap_element_ids:
                    return False
        
        if areas_checked > 0:
            self.fixed = True
            return True
        
        return False
    
    @staticmethod
    def create_from_group(group_dict, view_id):
        """Create a Gap from a gap group dictionary.
        
        Args:
            group_dict: Dict with 'center', 'length', 'areas', 'boundary_elements'
            view_id: ElementId of the view
        """
        center = group_dict.get('center')
        if not center:
            return None
        
        length_ft = group_dict.get('length', 0) or 0
        
        # Extract boundary element IDs as integers
        boundary_ids = set()
        for eid in group_dict.get('boundary_elements', set()):
            if eid and eid != DB.ElementId.InvalidElementId:
                boundary_ids.add(get_element_id_value(eid))
        
        # Extract area IDs as integers
        area_ids = set()
        for area in group_dict.get('areas', set()):
            if area:
                area_ids.add(get_element_id_value(area.Id))
        
        return Gap(center, length_ft, boundary_ids, area_ids, view_id)


def group_gaps_by_location(bad_areas):
    """Group gaps by exact location. Same gap affects multiple areas."""
    gap_dict = {}  # key -> group
    nonloc_areas = []
    
    for area_data in bad_areas:
        area = area_data.get('area')
        gap_data_list = area_data.get('gap_data') or []
        
        if not gap_data_list:
            nonloc_areas.append(area_data)
            continue
        
        for gap_info in gap_data_list:
            gap_point = gap_info.get('location')
            gap_length = gap_info.get('length', 0)
            gap_elements = gap_info.get('elements') or []
            if not gap_point:
                continue
                
            key = gap_point_key(gap_point)
            
            if key not in gap_dict:
                gap_dict[key] = {
                    'center': gap_point,
                    'length': gap_length,
                    'areas': set(),
                    'boundary_elements': set(),
                }
            
            group = gap_dict[key]
            if area is not None:
                group['areas'].add(area)
            for elem_id in gap_elements:
                if elem_id and elem_id != DB.ElementId.InvalidElementId:
                    group['boundary_elements'].add(elem_id)
    
    return list(gap_dict.values()), nonloc_areas


# ============================================================
# BAD BOUNDARY DETECTION
# ============================================================

def is_loop_closed(boundary_loop):
    """Check if a boundary loop's curves form a closed chain.
    
    A loop is closed if:
    - It has at least one segment
    - Each curve's end point connects to the next curve's start point
    - The last curve's end point connects to the first curve's start point
    
    Returns:
        (is_closed, error_details, problematic_element_ids, gap_data) tuple
        gap_data is a list of dicts with 'location' (XYZ) and 'length' (feet)
    """
    if not boundary_loop:
        return False, "Empty loop", [], []
    
    segments = list(boundary_loop)
    if not segments:
        return False, "No segments in loop", [], []
    
    # Collect curves with their source element IDs
    curve_data = []
    for segment in segments:
        try:
            curve = segment.GetCurve()
            elem_id = segment.ElementId  # The boundary element (wall, area boundary, etc.)
            if curve:
                curve_data.append({'curve': curve, 'element_id': elem_id})
        except Exception:
            pass
    
    if not curve_data:
        return False, "No valid curves in loop", [], []
    
    if len(curve_data) == 1:
        # Single curve - check if it's closed (e.g., a circle)
        curve = curve_data[0]['curve']
        elem_id = curve_data[0]['element_id']
        if curve.IsBound:
            start = curve.GetEndPoint(0)
            end = curve.GetEndPoint(1)
            if not points_equal(start, end):
                # Gap is at the end of the single curve
                gap_length = start.DistanceTo(end)
                gap_elements = []
                if elem_id and elem_id != DB.ElementId.InvalidElementId:
                    gap_elements.append(elem_id)
                gap_data = [{'location': end, 'length': gap_length, 'elements': gap_elements}]
                return False, "Single curve is not closed", gap_elements, gap_data
        return True, None, [], []
    
    # Multiple curves - check end-to-end connectivity
    problematic_ids = []
    gap_data = []
    
    for i in range(len(curve_data)):
        current = curve_data[i]
        next_data = curve_data[(i + 1) % len(curve_data)]
        
        current_end = current['curve'].GetEndPoint(1)
        next_start = next_data['curve'].GetEndPoint(0)
        
        if not points_equal(current_end, next_start):
            # Gap location is midpoint between the two disconnected points
            gap_point = DB.XYZ(
                (current_end.X + next_start.X) / 2,
                (current_end.Y + next_start.Y) / 2,
                (current_end.Z + next_start.Z) / 2
            )
            gap_length = current_end.DistanceTo(next_start)
            gap_elements = []
            
            # Add both elements involved in the gap
            cur_eid = current['element_id']
            if cur_eid and cur_eid != DB.ElementId.InvalidElementId:
                gap_elements.append(cur_eid)
                if cur_eid not in problematic_ids:
                    problematic_ids.append(cur_eid)

            next_eid = next_data['element_id']
            if next_eid and next_eid != DB.ElementId.InvalidElementId:
                if next_eid not in gap_elements:
                    gap_elements.append(next_eid)
                if next_eid not in problematic_ids:
                    problematic_ids.append(next_eid)

            gap_data.append({'location': gap_point, 'length': gap_length, 'elements': gap_elements})
    
    if gap_data:
        return False, "Found {} gap(s) in boundary".format(len(gap_data)), problematic_ids, gap_data
    
    return True, None, [], []


def check_area_boundaries(area):
    """Check all boundary loops of an area for validity.
    
    Returns:
        dict with 'valid', 'issues', 'problematic_elements', 'gap_data'
        gap_data is list of dicts with 'location' and 'length'
    """
    result = {
        'valid': True,
        'issues': [],
        'loop_count': 0,
        'problematic_elements': [],
        'gap_data': []
    }
    
    try:
        opts = DB.SpatialElementBoundaryOptions()
        loops = area.GetBoundarySegments(opts)
        
        if not loops:
            result['valid'] = False
            result['issues'].append("No boundary segments found")
            return result
        
        loop_list = list(loops)
        result['loop_count'] = len(loop_list)
        
        if result['loop_count'] == 0:
            result['valid'] = False
            result['issues'].append("Empty boundary collection")
            return result
        
        for i, loop in enumerate(loop_list):
            is_closed, error, problem_ids, gap_data = is_loop_closed(loop)
            if not is_closed:
                result['valid'] = False
                result['issues'].append("Loop {}: {}".format(i + 1, error))
                result['problematic_elements'].extend(problem_ids)
                result['gap_data'].extend(gap_data)
    
    except Exception as e:
        result['valid'] = False
        result['issues'].append("Error reading boundaries: {}".format(str(e)))
    
    return result


# ============================================================
# VIEW SELECTION DIALOG
# ============================================================

class ViewSelectionDialog(forms.WPFWindow):
    """Dialog for selecting AreaPlan views to check."""
    
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
        
        self.selected_views = None
    
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
            cb.Checked += self._on_view_checkbox_changed
            cb.Unchecked += self._on_view_checkbox_changed

            self.panel_views.Children.Add(cb)
            self._checkboxes[cb.Tag] = cb
        
        self._update_views_header(len([cb for cb in self._checkboxes.values() if cb.IsChecked]))
    
    def _on_select_all(self, sender, args):
        for cb in self._checkboxes.values():
            cb.IsChecked = True
        self._update_views_header(len(self._checkboxes))
    
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
        
        self.DialogResult = True
        self.Close()
    
    def _on_cancel(self, sender, args):
        self.DialogResult = False
        self.Close()

    def _on_view_checkbox_changed(self, sender, args):
        self._update_views_header(len([cb for cb in self._checkboxes.values() if cb.IsChecked]))

    def _update_views_header(self, selected_count):
        base_text = "Area Plans to check:"
        if hasattr(self, 'txt_views_header_label') and self.txt_views_header_label is not None:
            self.txt_views_header_label.Text = base_text

        if hasattr(self, 'txt_views_header_count') and self.txt_views_header_count is not None:
            if selected_count > 0:
                self.txt_views_header_count.Text = u"{} selected".format(selected_count)
            else:
                self.txt_views_header_count.Text = ""


# ============================================================
# MODELESS VISUALIZATION WINDOW
# ============================================================

# Global state for ExternalEvent communication
_views_to_restore = []
_dc3d_server_to_clear = None
_initial_open_view_ids = set()
_candidate_view_ids = set()
_nav_flat_list = []
_nav_index = 0
_view_to_open_id = None
_recheck_gap_indices = []  # Gap indices to recheck
_active_window = None  # Reference to the active visualization window
_original_sheet_state = None  # Stores (sheet_id, zoom_corners) when escaping from activated viewport


class SimpleEventHandler(UI.IExternalEventHandler):
    """ExternalEvent handler that executes a provided callable."""
    def __init__(self, do_this):
        self.do_this = do_this

    def Execute(self, uiapp):
        try:
            self.do_this()
        except Exception:
            pass

    def GetName(self):
        return "SimpleEventHandler"


def _restore_views_action():
    """Restore view states and close views opened by this command."""
    global _views_to_restore, _dc3d_server_to_clear, _initial_open_view_ids, _candidate_view_ids, _original_sheet_state
    
    if _dc3d_server_to_clear:
        try:
            _dc3d_server_to_clear.meshes = []
        except Exception:
            pass
        _dc3d_server_to_clear = None
    
    if _views_to_restore:
        view_ids = list(_views_to_restore)
        _views_to_restore = []
        try:
            with revit.Transaction("Restore View States"):
                for view_id in view_ids:
                    try:
                        view = revit.doc.GetElement(view_id)
                        if view and view.IsValidObject:
                            view.DisableTemporaryViewMode(DB.TemporaryViewMode.TemporaryViewProperties)
                    except Exception:
                        pass
        except Exception:
            pass
    
    # Close views opened by this command
    try:
        uidoc = HOST_APP.uidoc
        if uidoc:
            for uiv in list(uidoc.GetOpenUIViews()):
                vid = uiv.ViewId
                if _candidate_view_ids and vid in _candidate_view_ids and vid not in _initial_open_view_ids:
                    try:
                        uiv.Close()
                    except Exception:
                        pass
    except Exception:
        pass
    
    # Restore original sheet view if we escaped from one
    if _original_sheet_state:
        try:
            from System.Windows.Forms import Application
            doc = HOST_APP.doc
            uidoc = HOST_APP.uidoc
            
            sheet_id = _original_sheet_state.get('sheet_id')
            if sheet_id and uidoc:
                sheet = doc.GetElement(sheet_id)
                if sheet and sheet.IsValidObject:
                    uidoc.ActiveView = sheet
                    Application.DoEvents()
                    for uiv in uidoc.GetOpenUIViews():
                        if uiv.ViewId == sheet_id:
                            uiv.ZoomToFit()
                            break
        except Exception:
            pass
        
        _original_sheet_state = None


def _navigate_to_gap():
    """Navigate to the current gap (runs in Revit API context)."""
    global _nav_flat_list, _nav_index
    
    if not _nav_flat_list or _nav_index < 0 or _nav_index >= len(_nav_flat_list):
        return
    
    item = _nav_flat_list[_nav_index]
    view_id, center = item.get('view_id'), item.get('center')
    if not view_id or not center:
        return
    
    doc, uidoc = revit.doc, revit.uidoc
    view = doc.GetElement(view_id)
    if not view:
        return
    
    # Detect if we're in an activated viewport on a sheet by checking window title
    # Only do the escape dance if we're actually on a sheet
    is_in_activated_viewport = False
    try:
        import ctypes
        user32 = ctypes.windll.user32
        revit_hwnd = HOST_APP.proc_window
        if revit_hwnd:
            length = user32.GetWindowTextLengthW(revit_hwnd)
            if length > 0:
                buffer = ctypes.create_unicode_buffer(length + 1)
                user32.GetWindowTextW(revit_hwnd, buffer, length + 1)
                title = buffer.value
                # Window title contains "Sheet:" when viewing a sheet (even with activated viewport)
                if "Sheet:" in title or "- Sheet:" in title:
                    is_in_activated_viewport = True
    except Exception:
        pass
    
    # If in activated viewport, escape it by closing and reopening the view
    if is_in_activated_viewport:
        global _original_sheet_state
        try:
            from System.Windows.Forms import Application
            
            current_uiv = None
            for uiv in uidoc.GetOpenUIViews():
                if uiv.ViewId == view.Id:
                    current_uiv = uiv
                    break
            
            if current_uiv:
                # Save original sheet state for restoration on Done
                if _original_sheet_state is None:
                    try:
                        sheet_collector = DB.FilteredElementCollector(doc).OfClass(DB.ViewSheet)
                        for sheet in sheet_collector:
                            try:
                                for vp_id in sheet.GetAllViewports():
                                    vp = doc.GetElement(vp_id)
                                    if vp and vp.ViewId == view.Id:
                                        _original_sheet_state = {'sheet_id': sheet.Id}
                                        break
                                if _original_sheet_state:
                                    break
                            except:
                                continue
                    except:
                        pass
                
                # Find another view to open temporarily
                other_view = None
                for v in DB.FilteredElementCollector(doc).OfClass(DB.View):
                    try:
                        if v.Id != view.Id and not v.IsTemplate and v.CanBePrinted:
                            other_view = v
                            break
                    except:
                        continue
                
                if other_view:
                    uidoc.ActiveView = other_view
                    for _ in range(3):
                        Application.DoEvents()
                    try:
                        current_uiv.Close()
                    except:
                        pass
                    for _ in range(3):
                        Application.DoEvents()
                    uidoc.ActiveView = view
                    for _ in range(3):
                        Application.DoEvents()
        except:
            pass
    
    # Activate view
    try:
        uidoc.ActiveView = view
    except Exception:
        pass
    
    # Pump the Windows message queue to let Revit fully process view activation
    try:
        from System.Windows.Forms import Application
        Application.DoEvents()
    except Exception:
        pass
    
    # Find UIView
    target_uiv = None
    for uiv in uidoc.GetOpenUIViews():
        if uiv.ViewId == view.Id:
            target_uiv = uiv
            break
    
    if not target_uiv:
        return
    
    # Zoom to gap location
    half = 5.0
    min_pt = DB.XYZ(center.X - half, center.Y - half, center.Z)
    max_pt = DB.XYZ(center.X + half, center.Y + half, center.Z)
    
    try:
        target_uiv.ZoomToFit()
    except Exception:
        pass
    
    for _ in range(3):
        try:
            target_uiv.ZoomAndCenterRectangle(min_pt, max_pt)
        except Exception:
            pass


def _open_view_and_fit():
    """Open a view and zoom to fit (runs in Revit API context)."""
    global _view_to_open_id
    if not _view_to_open_id:
        return
    
    doc, uidoc = revit.doc, revit.uidoc
    view = doc.GetElement(_view_to_open_id)
    if not view:
        return
    
    try:
        uidoc.ActiveView = view
        for uiv in uidoc.GetOpenUIViews():
            if uiv.ViewId == view.Id:
                uiv.ZoomToFit()
                break
    except Exception:
        pass


# Create ExternalEvents at module level
_restore_handler = SimpleEventHandler(_restore_views_action)
_restore_event = UI.ExternalEvent.Create(_restore_handler)
_nav_handler = SimpleEventHandler(_navigate_to_gap)
_nav_event = UI.ExternalEvent.Create(_nav_handler)
_view_open_handler = SimpleEventHandler(_open_view_and_fit)
_view_open_event = UI.ExternalEvent.Create(_view_open_handler)


def _recheck_gaps_action():
    """Recheck gaps after boundary elements were modified (runs in Revit API context)."""
    global _recheck_gap_indices, _active_window
    if not _active_window or not _recheck_gap_indices:
        return
    
    try:
        _active_window._recheck_gaps(list(_recheck_gap_indices))
    except Exception:
        pass
    finally:
        _recheck_gap_indices = []


_recheck_handler = SimpleEventHandler(_recheck_gaps_action)
_recheck_event = UI.ExternalEvent.Create(_recheck_handler)


# ============================================================
# VIEW STATE MANAGEMENT (Using Temporary View Properties Mode)
# ============================================================

def get_subcategory(doc, parent_category, subcategory_name):
    """Get a subcategory by name from a parent category."""
    try:
        parent_cat = doc.Settings.Categories.get_Item(parent_category)
        if parent_cat and parent_cat.SubCategories:
            for subcat in parent_cat.SubCategories:
                if subcat.Name == subcategory_name:
                    return subcat
    except Exception:
        pass
    return None


def apply_temporary_visibility(doc, view):
    """Apply temporary visibility settings using Revit's Temporary View Properties mode.
    
    This enables the purple frame indicator and auto-restores when disabled.
    
    Settings applied:
    - Area boundaries: visible
    - Areas: visible
    - Area Interior Fill: visible (the hatch pattern)
    - Area Color Fill: hidden (the colored regions from color schemes)
    - Display Model: Halftone
    """
    if not view or not view.IsValidObject:
        return False
    
    try:
        # Enable temporary view properties mode - shows purple frame
        # All changes made after this are temporary and auto-restore when disabled
        if not view.IsTemporaryViewPropertiesModeEnabled():
            view.EnableTemporaryViewPropertiesMode(view.Id)
    except Exception:
        return False
    
    try:
        # Set Display Model to Halftone (0=Normal, 1=Halftone, 2=Do not display)
        display_model_param = view.get_Parameter(DB.BuiltInParameter.VIEW_MODEL_DISPLAY_MODE)
        if display_model_param and not display_model_param.IsReadOnly:
            display_model_param.Set(1)  # Halftone
    except Exception:
        pass
    
    try:
        # Show area boundaries (OST_AreaSchemeLines)
        view.SetCategoryHidden(
            DB.ElementId(DB.BuiltInCategory.OST_AreaSchemeLines),
            False  # visible
        )
    except Exception:
        pass
    
    try:
        # Show areas (OST_Areas)
        view.SetCategoryHidden(
            DB.ElementId(DB.BuiltInCategory.OST_Areas),
            False  # visible
        )
    except Exception:
        pass
    
    # For subcategories, we need to get them by name and use their actual category ID
    try:
        # Get Interior Fill subcategory and show it
        interior_fill_cat = get_subcategory(doc, DB.BuiltInCategory.OST_Areas, "Interior Fill")
        if interior_fill_cat:
            view.SetCategoryHidden(interior_fill_cat.Id, False)  # visible
    except Exception:
        pass
    
    try:
        # Get Color Fill subcategory and hide it
        color_fill_cat = get_subcategory(doc, DB.BuiltInCategory.OST_Areas, "Color Fill")
        if color_fill_cat:
            view.SetCategoryHidden(color_fill_cat.Id, True)  # hidden
    except Exception:
        pass
    
    return True


class VisualizationControlWindow(forms.WPFWindow):
    """Modeless window for navigating bad boundaries."""
    
    XAML = '''
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Bad Boundaries Visualization"
        Width="320" SizeToContent="Height"
        Topmost="True" ShowInTaskbar="True"
        ResizeMode="NoResize"
        WindowStartupLocation="Manual">
    <StackPanel Margin="15">
        <TextBlock x:Name="status_text" TextWrapping="Wrap" Margin="0,0,0,10"/>
        <TextBlock x:Name="instructions_text" Foreground="Gray" TextWrapping="Wrap" Margin="0,0,0,10"/>
        <TreeView x:Name="gaps_tree" Margin="0,0,0,10" Height="200"/>
        <StackPanel Orientation="Horizontal" HorizontalAlignment="Center" Margin="0,0,0,10">
            <Button x:Name="prev_btn" Content="Previous gap" Width="110" Margin="0,0,5,0"/>
            <Button x:Name="next_btn" Content="Next gap" Width="110" Margin="5,0,0,0"/>
        </StackPanel>
        <Button x:Name="done_btn" Content="Done - Clear Visualization" Height="32"/>
    </StackPanel>
</Window>
'''
    
    def __init__(self, bad_count, gaps_by_view, dc3d_server, temp_view_ids=None):
        forms.WPFWindow.__init__(self, self.XAML, literal_string=True)
        
        global _active_window
        _active_window = self
        
        self._server = dc3d_server
        self._gaps_by_view = gaps_by_view
        self._temp_view_ids = temp_view_ids or []
        self._states_restored = False
        self._bad_count = bad_count
        self._current_index = -1
        self._ignore_selection = False
        self._doc_changed_handler = None
        
        # Gap objects list - each Gap tracks its own state
        self._gaps = []  # List of Gap objects
        self._gap_items = {}  # flat_index -> TreeViewItem
        self._gap_display_numbers = {}  # flat_index -> display number
        
        # Index for fast lookup: boundary_element_id -> list of Gap objects
        self._boundary_to_gaps = {}  # int -> [Gap, ...]
        
        # Build Gap objects from gap groups
        self._structured_gaps = []
        for view_id, gap_groups in gaps_by_view.items():
            view = doc.GetElement(view_id)
            if not view:
                continue
            sorted_groups = sorted(gap_groups, key=lambda g: g.get('length', 0), reverse=True)
            view_gaps = []
            for group in sorted_groups:
                gap_obj = Gap.create_from_group(group, view_id)
                if gap_obj:
                    gap_obj.index = len(self._gaps)
                    self._gaps.append(gap_obj)
                    view_gaps.append(gap_obj)
                    
                    # Build boundary element index for fast lookup
                    for bid in gap_obj.boundary_element_ids:
                        if bid not in self._boundary_to_gaps:
                            self._boundary_to_gaps[bid] = []
                        self._boundary_to_gaps[bid].append(gap_obj)
            
            if view_gaps:
                self._structured_gaps.append({'view_id': view_id, 'view_name': view.Name, 'gaps': view_gaps})
        
        # Position window
        self.Left = System.Windows.SystemParameters.WorkArea.Width - self.Width - 20
        self.Top = 80
        
        # Set text
        self.status_text.Text = "Found {} bad boundaries.\nRed circles mark gap locations.".format(bad_count)
        self.instructions_text.Text = "Click a gap to zoom, or use Next/Previous.\nDouble-click view header to zoom to fit."
        
        # Wire events
        self.done_btn.Click += self._on_done
        self.Closed += self._on_closed
        self.prev_btn.Click += self._on_prev
        self.next_btn.Click += self._on_next
        self.gaps_tree.SelectedItemChanged += self._on_selection_changed
        self.gaps_tree.MouseDoubleClick += self._on_double_click
        
        # Setup DC3D server - clear any stale data from previous runs
        self._server.meshes = []  # Clear previous visualization
        self._server.enabled_view_types = [
            DB.ViewType.ThreeD, DB.ViewType.Elevation, DB.ViewType.Section,
            DB.ViewType.FloorPlan, DB.ViewType.CeilingPlan, DB.ViewType.EngineeringPlan, DB.ViewType.AreaPlan,
        ]
        self._server.add_server()
        self._create_visualization()
        self._build_tree()
        self._update_state()
        self._subscribe_doc_events()
    
    def _create_visualization(self):
        """Create circles at gap locations - red for unfixed, green for fixed."""
        self._rebuild_dc3d_visualization()
    
    def _rebuild_dc3d_visualization(self):
        """Rebuild DC3D visualization with current gap states.
        
        Red circles for unfixed gaps, green circles for fixed gaps.
        """
        # Colors: red for unfixed, green for fixed
        color_unfixed = DB.ColorWithTransparency(255, 0, 0, 0)  # Red
        color_fixed = DB.ColorWithTransparency(0, 180, 0, 0)    # Green
        
        edges = []
        for gap in self._gaps:
            center = gap.center
            color = color_fixed if gap.fixed else color_unfixed
            for i in range(64):
                a1, a2 = 2 * math.pi * i / 64, 2 * math.pi * (i + 1) / 64
                p1 = DB.XYZ(center.X + math.cos(a1), center.Y + math.sin(a1), center.Z)
                p2 = DB.XYZ(center.X + math.cos(a2), center.Y + math.sin(a2), center.Z)
                edges.append(revit.dc3dserver.Edge(p1, p2, color))
        
        if edges:
            self._server.meshes = [revit.dc3dserver.Mesh(edges, [])]
        else:
            self._server.meshes = []
    
    def _build_tree(self):
        """Build TreeView with views and gaps."""
        self.gaps_tree.Items.Clear()
        self._gap_items = {}
        self._gap_display_numbers = {}
        gap_num = 1
        for view_info in self._structured_gaps:
            view_item = System.Windows.Controls.TreeViewItem()
            view_item.Header = "{} ({} gap{})".format(
                view_info['view_name'], len(view_info['gaps']),
                "" if len(view_info['gaps']) == 1 else "s")
            view_item.IsExpanded = True
            view_item.Tag = "view:{}".format(get_element_id_value(view_info['view_id']))
            for gap in view_info['gaps']:
                flat_idx = gap.index
                self._gap_display_numbers[flat_idx] = gap_num
                gap_item = System.Windows.Controls.TreeViewItem()
                gap_item.Header = self._format_gap_header(gap)
                gap_item.Tag = "gap:{}".format(flat_idx)
                self._gap_items[flat_idx] = gap_item
                view_item.Items.Add(gap_item)
                gap_num += 1
            self.gaps_tree.Items.Add(view_item)
    
    def _format_gap_header(self, gap):
        """Format gap header text. gap can be Gap object or index."""
        if isinstance(gap, int):
            # Called with index - get the Gap object
            if gap < 0 or gap >= len(self._gaps):
                return "Gap {}".format(gap + 1)
            gap = self._gaps[gap]
        
        display_index = self._gap_display_numbers.get(gap.index, gap.index + 1)
        length_cm = gap.length_cm or 0
        if length_cm:
            header = "Gap {} ({:.1f} cm)".format(display_index, length_cm)
        else:
            header = "Gap {}".format(display_index)
        if gap.fixed:
            header += " [FIXED]"
        return header
    
    def _update_gap_visual(self, gap):
        """Update TreeViewItem appearance for a gap. gap can be Gap object or index."""
        if isinstance(gap, int):
            if gap < 0 or gap >= len(self._gaps):
                return
            gap = self._gaps[gap]
        
        gap_item = self._gap_items.get(gap.index)
        if not gap_item:
            return
        gap_item.Header = self._format_gap_header(gap)
        if gap.fixed:
            try:
                gap_item.Foreground = System.Windows.Media.Brushes.Gray
            except Exception:
                pass
            try:
                gap_item.FontStyle = System.Windows.FontStyles.Italic
            except Exception:
                pass
    
    def _subscribe_doc_events(self):
        try:
            app = doc.Application
        except Exception:
            app = None
        if not app:
            return
        try:
            self._doc_changed_handler = self._on_doc_changed
            app.DocumentChanged += self._doc_changed_handler
        except Exception:
            self._doc_changed_handler = None
    
    def _unsubscribe_doc_events(self):
        global _active_window
        try:
            app = doc.Application
        except Exception:
            app = None
        if app and self._doc_changed_handler:
            try:
                app.DocumentChanged -= self._doc_changed_handler
            except Exception:
                pass
        self._doc_changed_handler = None
        if _active_window is self:
            _active_window = None
    
    def _on_doc_changed(self, sender, args):
        """Handle document changes - recheck gaps whose boundary elements were modified."""
        try:
            # Collect all changed element IDs
            changed_ids = set()
            try:
                for eid in args.GetModifiedElementIds():
                    changed_ids.add(eid)
            except Exception:
                pass
            try:
                for eid in args.GetAddedElementIds():
                    changed_ids.add(eid)
            except Exception:
                pass
            try:
                for eid in args.GetDeletedElementIds():
                    changed_ids.add(eid)
            except Exception:
                pass
            
            if not changed_ids:
                return
            
            # Convert to integer IDs for lookup
            changed_int_ids = set()
            for eid in changed_ids:
                try:
                    changed_int_ids.add(get_element_id_value(eid))
                except Exception:
                    pass
            
            if not changed_int_ids:
                return
            
            # Find gaps that reference the changed boundary elements
            gaps_to_recheck = set()
            for bid in changed_int_ids:
                if bid in self._boundary_to_gaps:
                    for gap in self._boundary_to_gaps[bid]:
                        if not gap.fixed:
                            gaps_to_recheck.add(gap)
            
            if gaps_to_recheck:
                self._recheck_gaps(gaps_to_recheck)
        except Exception:
            pass
    
    def _recheck_gaps(self, gaps_to_recheck):
        """Recheck a set of Gap objects to see if they are now fixed."""
        if not gaps_to_recheck:
            return
        
        doc_local = revit.doc
        any_changed = False
        
        for gap in gaps_to_recheck:
            if gap.recheck(doc_local):
                any_changed = True
                self._update_gap_visual(gap)
        
        if any_changed:
            # Update DC3D visualization - fixed gaps become green circles
            try:
                self._rebuild_dc3d_visualization()
            except Exception:
                pass
            
            # Auto-navigate to next unfixed gap if current gap was fixed
            try:
                if self._current_index >= 0:
                    current_gap = self._gaps[self._current_index]
                    if current_gap.fixed:
                        self._navigate_to_next_unfixed()
            except Exception:
                pass
            
            try:
                fixed_count = len([g for g in self._gaps if g.fixed])
                total = len(self._gaps)
                self.status_text.Text = "Found {} bad boundaries.\n{} of {} gaps marked as fixed.".format(
                    self._bad_count, fixed_count, total)
            except Exception:
                pass
    
    def _update_state(self):
        """Update button states and status text."""
        has_gaps = bool(self._gaps)
        self.prev_btn.IsEnabled = has_gaps and len(self._gaps) > 1
        self.next_btn.IsEnabled = has_gaps and len(self._gaps) > 1
        if has_gaps and self._current_index >= 0:
            self.status_text.Text = "Found {} bad boundaries.\nGap {} of {}.".format(
                self._bad_count, self._current_index + 1, len(self._gaps))
        else:
            self.status_text.Text = "Found {} bad boundaries.\nRed circles mark gap locations.".format(self._bad_count)
    
    def _sync_selection(self):
        """Sync TreeView selection to current index and scroll to ensure visibility."""
        if self._current_index < 0:
            return
        target = "gap:{}".format(self._current_index)
        self._ignore_selection = True
        try:
            target_item = None
            for view_item in self.gaps_tree.Items:
                for gap_item in view_item.Items:
                    if str(gap_item.Tag) == target:
                        gap_item.IsSelected = True
                        view_item.IsExpanded = True
                        target_item = gap_item
                    else:
                        gap_item.IsSelected = False
            
            # Auto-scroll to make selected item visible
            if target_item:
                try:
                    # Use BringIntoView to scroll the item into visibility
                    target_item.BringIntoView()
                    # Alternative approach if BringIntoView doesn't work well:
                    # target_item.Focus()
                except Exception:
                    # Fallback: try to scroll the parent view item
                    try:
                        parent = target_item.Parent
                        if parent:
                            parent.BringIntoView()
                    except Exception:
                        pass
        finally:
            self._ignore_selection = False
    
    def _navigate(self):
        """Raise navigation event to zoom to current gap."""
        global _nav_flat_list, _nav_index, _nav_event
        if self._current_index >= 0 and self._gaps:
            # Build flat list from Gap objects for navigation
            _nav_flat_list = [{'view_id': g.view_id, 'center': g.center} for g in self._gaps]
            _nav_index = self._current_index
            _nav_event.Raise()
    
    def _navigate_to_next_unfixed(self):
        """Navigate to the next unfixed gap, starting from current position.
        
        This is called when a gap is fixed (e.g., via trim tool). To allow
        navigation to work while the user is in a tool mode, we send ESC key
        from a background thread with a delay to exit the tool first, then
        trigger navigation after the tool has exited.
        """
        if not self._gaps:
            return
        
        # Search forward from current position for an unfixed gap
        start = self._current_index if self._current_index >= 0 else 0
        n = len(self._gaps)
        
        next_unfixed_idx = None
        for i in range(1, n + 1):
            idx = (start + i) % n
            if not self._gaps[idx].fixed:
                next_unfixed_idx = idx
                break
        
        if next_unfixed_idx is None:
            # All gaps are fixed - stay on current
            return
        
        # Update UI immediately
        self._current_index = next_unfixed_idx
        self._update_state()
        self._sync_selection()
        
        # Send ESC key from a background thread to exit tool mode (e.g., trim),
        # then trigger navigation AFTER the ESC has been processed.
        # The delay ensures Revit finishes its current operation and is ready.
        window_ref = self  # Capture reference for the thread
        
        def send_esc_then_navigate():
            try:
                # Delay to let Revit finish processing the current tool action
                Thread.Sleep(150)
                # Send ESC key to exit current tool mode
                SendKeys.SendWait("{ESC}")
                # Wait for ESC to be processed
                Thread.Sleep(100)
                # Now trigger navigation via ExternalEvent
                window_ref._navigate()
            except Exception:
                pass
        
        try:
            esc_thread = Thread(ThreadStart(send_esc_then_navigate))
            esc_thread.IsBackground = True
            esc_thread.Start()
        except Exception:
            # Fallback: try navigating directly (may not work if tool is active)
            self._navigate()
    
    def _on_selection_changed(self, sender, args):
        """Handle gap selection from TreeView."""
        if self._ignore_selection:
            return
        tvi = self.gaps_tree.SelectedItem
        if not tvi:
            return
        tag = str(getattr(tvi, 'Tag', '') or '')
        if tag.startswith("gap:"):
            try:
                idx = int(tag.split(":")[1])
                if 0 <= idx < len(self._gaps) and idx != self._current_index:
                    self._current_index = idx
                    self._update_state()
                    # Ensure the selected item is visible (auto-scroll)
                    try:
                        tvi.BringIntoView()
                    except Exception:
                        pass
                    self._navigate()
            except ValueError:
                pass
        elif self._current_index >= 0:
            # Spurious event selected wrong item (view header) - re-sync to correct gap
            self._sync_selection()
    
    def _on_double_click(self, sender, args):
        """Handle double-click on view header to zoom to fit."""
        global _view_to_open_id, _view_open_event
        try:
            args.Handled = True
        except Exception:
            pass
        tvi = self.gaps_tree.SelectedItem
        if not tvi:
            return
        tag = str(getattr(tvi, 'Tag', '') or '')
        if tag.startswith("view:"):
            try:
                _view_to_open_id = DB.ElementId(int(tag.split(":")[1]))
                _view_open_event.Raise()
            except Exception:
                pass
    
    def _on_next(self, sender, args):
        if not self._gaps:
            return
        self._current_index = (self._current_index + 1) % len(self._gaps) if self._current_index >= 0 else 0
        self._update_state()
        self._sync_selection()
        self._navigate()
    
    def _on_prev(self, sender, args):
        if not self._gaps:
            return
        self._current_index = (self._current_index - 1) % len(self._gaps) if self._current_index >= 0 else len(self._gaps) - 1
        self._update_state()
        self._sync_selection()
        self._navigate()
    
    def _on_done(self, sender, args):
        """Clean up and close."""
        global _views_to_restore, _dc3d_server_to_clear
        self._unsubscribe_doc_events()
        self._states_restored = True
        
        # Clear DC3D visualization
        try:
            self._server.meshes = []
            self._server.remove_server()
        except Exception:
            pass
        
        _dc3d_server_to_clear = self._server
        _views_to_restore = list(self._temp_view_ids)
        _restore_event.Raise()
        self.Close()
    
    def _on_closed(self, sender, args):
        """Ensure cleanup on window close."""
        global _views_to_restore, _dc3d_server_to_clear
        self._unsubscribe_doc_events()
        if not self._states_restored:
            _dc3d_server_to_clear = self._server
            _views_to_restore = list(self._temp_view_ids)
            _restore_event.Raise()


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


def process_view(doc, view):
    """Process a single view and check for bad boundaries.
    
    Returns:
        (total_count, bad_count, bad_areas)
    """
    areas = get_areas_in_view(doc, view)
    if not areas:
        return 0, 0, []
    
    bad_areas = []
    for area in areas:
        result = check_area_boundaries(area)
        if not result['valid']:
            bad_areas.append({
                'area': area,
                'issues': result['issues'],
                'problematic_elements': result['problematic_elements'],
                'gap_data': result['gap_data']
            })
    
    return len(areas), len(bad_areas), bad_areas


def detect_bad_boundaries():
    """Main entry point - detect bad boundaries and show visualization window."""
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
    
    # Show view selection dialog
    dialog = ViewSelectionDialog(schemes, selected_index, preselected_ids)
    if not dialog.ShowDialog():
        return None
    
    selected_views = dialog.selected_views
    if not selected_views:
        return None
    
    # Sort views by elevation
    selected_views.sort(key=lambda v: v.Origin.Z if hasattr(v, 'Origin') else 0)
    
    # Process views and collect gaps
    total_areas = 0
    total_bad = 0
    gaps_by_view = {}
    
    for view in selected_views:
        count, bad_count, bad_areas = process_view(doc, view)
        total_areas += count
        total_bad += bad_count
        
        if not bad_areas:
            continue
        
        # Group gaps by location and build gap data for this view
        gap_groups, nonloc_areas = group_gaps_by_location(bad_areas)
        gap_groups_for_view = list(gap_groups)
        
        # Add non-located areas (those without specific gap points) using area center
        for area_data in nonloc_areas:
            area = area_data.get('area')
            if not area:
                continue
            location = getattr(area, 'Location', None)
            if not location or not hasattr(location, 'Point'):
                continue
            elem_ids = [eid for eid in (area_data.get('problematic_elements') or [])
                        if eid and eid != DB.ElementId.InvalidElementId]
            gap_groups_for_view.append({
                'center': location.Point,
                'areas': set([area]),
                'boundary_elements': set(elem_ids),
            })
        
        if gap_groups_for_view:
            gaps_by_view[view.Id] = gap_groups_for_view
    
    # Check results
    if total_bad == 0:
        forms.alert(
            "All {} areas have valid boundaries.".format(total_areas),
            title="No Issues Found"
        )
        return None
    
    # Apply temporary visibility using Revit's Temporary View Properties mode
    temp_view_ids = []
    with revit.Transaction("Enable Temporary View Properties"):
        for view in selected_views:
            if apply_temporary_visibility(doc, view):
                temp_view_ids.append(view.Id)
    
    return total_bad, gaps_by_view, temp_view_ids


# ============================================================
# RUN
# ============================================================

def _cleanup_dc3d_servers():
    """Clear any existing DC3D visualization before running."""
    try:
        import revit.dc3dserver
        for server in list(revit.dc3dserver.Server.GetRegisteredServers()):
            try:
                server.remove_server()
            except Exception:
                pass
    except Exception:
        pass


def _pre_open_views(uidoc, view_ids):
    """Pre-open all views with gaps to improve first-click zoom behavior."""
    if not uidoc or not view_ids:
        return
    
    original_view = uidoc.ActiveView
    for vid in view_ids:
        try:
            view = revit.doc.GetElement(vid)
            if view:
                uidoc.ActiveView = view
                for uiv in uidoc.GetOpenUIViews():
                    if uiv.ViewId == vid:
                        uiv.ZoomToFit()
                        break
        except Exception:
            pass
    
    # Restore original view
    if original_view:
        try:
            uidoc.ActiveView = original_view
        except Exception:
            pass


def _show_visualization(total_bad, gaps_by_view, temp_view_ids):
    """Create and show the visualization window."""
    global _initial_open_view_ids, _candidate_view_ids
    
    uidoc = HOST_APP.uidoc
    
    # Track initially open views for cleanup on Done
    _initial_open_view_ids = set(uiv.ViewId for uiv in uidoc.GetOpenUIViews()) if uidoc else set()
    _candidate_view_ids = set(gaps_by_view.keys())
    
    # Pre-open views to improve navigation
    _pre_open_views(uidoc, _candidate_view_ids)
    
    # Create DC3D server and visualization window
    dc3d_server = revit.dc3dserver.Server()
    dc3d_server.meshes = []
    
    window = VisualizationControlWindow(total_bad, gaps_by_view, dc3d_server, temp_view_ids)
    window.Show()


# Entry point
_cleanup_dc3d_servers()
result = detect_bad_boundaries()

if result:
    total_bad, gaps_by_view, temp_view_ids = result
    _show_visualization(total_bad, gaps_by_view, temp_view_ids)
