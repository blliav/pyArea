# -*- coding: utf-8 -*-
"""Detect Bad Boundaries - Identifies areas with invalid boundary loops.

Checks each area's boundary segments to detect:
1. Unclosed loops (curves don't connect end-to-end)
2. Empty boundaries (no curves at all)

Reports problematic areas with clickable links and temporary visualization.
"""

__title__ = "Detect\nBad Boundaries"
__author__ = "pyArea"
__persistentengine__ = True

import os
import sys
import math

from pyrevit import revit, DB, forms, script, HOST_APP, UI

# Import InvalidOperationException for proper error handling in ExternalEvent
try:
    from Autodesk.Revit.Exceptions import InvalidOperationException
except ImportError:
    InvalidOperationException = Exception

import clr
clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')
import System
from System.Windows.Controls import CheckBox

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
    global _views_to_restore, _dc3d_server_to_clear, _initial_open_view_ids, _candidate_view_ids
    
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
    
    # Activate view
    try:
        uidoc.ActiveView = view
    except Exception:
        pass
    
    # WORKAROUND: Pump the Windows message queue to let Revit fully process view activation.
    # Without this, when navigating to the first gap in a view, ZoomAndCenterRectangle
    # silently fails or zooms to the wrong location. This happens because setting
    # uidoc.ActiveView is asynchronous - Revit queues the view switch but doesn't
    # complete it before we call GetOpenUIViews() and zoom methods. By calling
    # Application.DoEvents(), we force Revit to process pending UI messages,
    # ensuring the view is fully activated before we attempt to zoom.
    # Alternative approaches that DON'T work: time.sleep(), doc.Regenerate(),
    # forced property access. Only message queue pumping resolves this.
    try:
        from System.Windows.Forms import Application
        Application.DoEvents()
    except Exception:
        pass
    
    # Find UIView
    target_uiv = None
    for attempt in range(2):
        open_views = list(uidoc.GetOpenUIViews())
        for uiv in open_views:
            if uiv.ViewId == view.Id:
                target_uiv = uiv
                break
        if target_uiv:
            break
        try:
            uidoc.ActiveView = view
        except Exception:
            pass
    
    if not target_uiv:
        return
    
    # Zoom
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
        
        self._server = dc3d_server
        self._gaps_by_view = gaps_by_view
        self._temp_view_ids = temp_view_ids or []
        self._states_restored = False
        self._bad_count = bad_count
        self._current_index = -1
        self._ignore_selection = False
        
        # Build flat list of gaps for navigation
        self._flat_gaps = []
        self._structured_gaps = []
        for view_id, gap_groups in gaps_by_view.items():
            view = doc.GetElement(view_id)
            if not view:
                continue
            sorted_groups = sorted(gap_groups, key=lambda g: g.get('length', 0), reverse=True)
            view_gaps = []
            for group in sorted_groups:
                center = group.get('center')
                if center:
                    view_gaps.append({
                        'flat_index': len(self._flat_gaps),
                        'length_cm': (group.get('length', 0) or 0) * 30.48,
                    })
                    self._flat_gaps.append({'view_id': view_id, 'center': center})
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
        
        # Setup DC3D server
        self._server.enabled_view_types = [
            DB.ViewType.ThreeD, DB.ViewType.Elevation, DB.ViewType.Section,
            DB.ViewType.FloorPlan, DB.ViewType.CeilingPlan, DB.ViewType.EngineeringPlan, DB.ViewType.AreaPlan,
        ]
        self._server.add_server()
        self._create_visualization()
        self._build_tree()
        self._update_state()
    
    def _create_visualization(self):
        """Create red circles at gap locations."""
        color = DB.ColorWithTransparency(255, 0, 0, 0)
        edges = []
        for gap in self._flat_gaps:
            center = gap['center']
            for i in range(64):
                a1, a2 = 2 * math.pi * i / 64, 2 * math.pi * (i + 1) / 64
                p1 = DB.XYZ(center.X + math.cos(a1), center.Y + math.sin(a1), center.Z)
                p2 = DB.XYZ(center.X + math.cos(a2), center.Y + math.sin(a2), center.Z)
                edges.append(revit.dc3dserver.Edge(p1, p2, color))
        if edges:
            self._server.meshes = [revit.dc3dserver.Mesh(edges, [])]
    
    def _build_tree(self):
        """Build TreeView with views and gaps."""
        self.gaps_tree.Items.Clear()
        gap_num = 1
        for view_info in self._structured_gaps:
            view_item = System.Windows.Controls.TreeViewItem()
            view_item.Header = "{} ({} gap{})".format(
                view_info['view_name'], len(view_info['gaps']),
                "" if len(view_info['gaps']) == 1 else "s")
            view_item.IsExpanded = True
            view_item.Tag = "view:{}".format(get_element_id_value(view_info['view_id']))
            for gap in view_info['gaps']:
                gap_item = System.Windows.Controls.TreeViewItem()
                length = gap.get('length_cm', 0)
                gap_item.Header = "Gap {} ({:.1f} cm)".format(gap_num, length) if length else "Gap {}".format(gap_num)
                gap_item.Tag = "gap:{}".format(gap['flat_index'])
                view_item.Items.Add(gap_item)
                gap_num += 1
            self.gaps_tree.Items.Add(view_item)
    
    def _update_state(self):
        """Update button states and status text."""
        has_gaps = bool(self._flat_gaps)
        self.prev_btn.IsEnabled = has_gaps and len(self._flat_gaps) > 1
        self.next_btn.IsEnabled = has_gaps and len(self._flat_gaps) > 1
        if has_gaps and self._current_index >= 0:
            self.status_text.Text = "Found {} bad boundaries.\nGap {} of {}.".format(
                self._bad_count, self._current_index + 1, len(self._flat_gaps))
        else:
            self.status_text.Text = "Found {} bad boundaries.\nRed circles mark gap locations.".format(self._bad_count)
    
    def _sync_selection(self):
        """Sync TreeView selection to current index."""
        if self._current_index < 0:
            return
        target = "gap:{}".format(self._current_index)
        self._ignore_selection = True
        try:
            for view_item in self.gaps_tree.Items:
                for gap_item in view_item.Items:
                    if str(gap_item.Tag) == target:
                        gap_item.IsSelected = True
                        view_item.IsExpanded = True
                    else:
                        gap_item.IsSelected = False
        finally:
            self._ignore_selection = False
    
    def _navigate(self):
        """Raise navigation event to zoom to current gap."""
        global _nav_flat_list, _nav_index, _nav_event
        if self._current_index >= 0 and self._flat_gaps:
            _nav_flat_list = list(self._flat_gaps)
            _nav_index = self._current_index
            _nav_event.Raise()
    
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
                if 0 <= idx < len(self._flat_gaps) and idx != self._current_index:
                    self._current_index = idx
                    self._update_state()
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
        if not self._flat_gaps:
            return
        self._current_index = (self._current_index + 1) % len(self._flat_gaps) if self._current_index >= 0 else 0
        self._update_state()
        self._sync_selection()
        self._navigate()
    
    def _on_prev(self, sender, args):
        if not self._flat_gaps:
            return
        self._current_index = (self._current_index - 1) % len(self._flat_gaps) if self._current_index >= 0 else len(self._flat_gaps) - 1
        self._update_state()
        self._sync_selection()
        self._navigate()
    
    def _on_done(self, sender, args):
        """Clean up and close."""
        global _views_to_restore, _dc3d_server_to_clear
        self._states_restored = True
        _dc3d_server_to_clear = self._server
        _views_to_restore = list(self._temp_view_ids)
        _restore_event.Raise()
        self.Close()
    
    def _on_closed(self, sender, args):
        """Ensure cleanup on window close."""
        global _views_to_restore, _dc3d_server_to_clear
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


def process_view(doc, view, output):
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
    """Main entry point."""
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
    
    # Sort by elevation
    selected_views.sort(key=lambda v: v.Origin.Z if hasattr(v, 'Origin') else 0)
    
    # Process views
    output = script.get_output()
    
    total_areas = 0
    total_bad = 0
    gaps_by_view = {}
    view_results = []
    
    for view in selected_views:
        count, bad_count, bad_areas = process_view(doc, view, output)
        total_areas += count
        total_bad += bad_count
        if count or bad_count:
            view_results.append((view, count, bad_areas))
    
    first_view = True
    for view, area_count, bad_areas in view_results:
        if not bad_areas:
            continue
        
        view_link = output.linkify(view.Id)
        
        gap_groups, nonloc_areas = group_gaps_by_location(bad_areas)
        gap_groups_for_view = list(gap_groups)
        
        # Count unique affected areas across all gaps
        affected_areas = set()
        for g in gap_groups:
            affected_areas.update(g['areas'])
        bad_area_count = len(affected_areas)
        
        # Separator line between views (not before first)
        if not first_view:
            output.print_html('<hr style="margin:15px 0;border:none;border-top:1px solid #ccc;">')
        first_view = False
        
        # View header - larger text for emphasis, view name before link
        output.print_html(
            '<h3 style="margin-top:10px;margin-bottom:5px;">{} ({} bad area{}, {} gap{}) {}</h3>'.format(
                view.Name,
                bad_area_count,
                "" if bad_area_count == 1 else "s",
                len(gap_groups),
                "" if len(gap_groups) == 1 else "s",
                view_link
            )
        )
        
        # Sort gaps by size (largest first)
        sorted_gaps = sorted(gap_groups, key=lambda g: g.get('length', 0), reverse=True)
        
        for idx, group in enumerate(sorted_gaps, 1):
            areas_in_gap = sorted(
                [a for a in group['areas'] if a is not None],
                key=lambda a: get_element_id_value(a.Id)
            )
            boundary_ids = sorted(
                list(group['boundary_elements']),
                key=lambda eid: get_element_id_value(eid)
            )
            
            # Get gap length in cm (Revit internal units are feet)
            gap_length_ft = group.get('length', 0)
            gap_length_cm = gap_length_ft * 30.48  # feet to cm
            
            # Create gap link that selects all boundary elements at once
            # Gap number and size before link
            if boundary_ids:
                gap_link = output.linkify(list(boundary_ids), title="Select boundary elements")
                output.print_html(
                    "&nbsp;&nbsp;<b>Gap {} ({:.1f} cm)</b> {}".format(idx, gap_length_cm, gap_link)
                )
            else:
                output.print_html(
                    "&nbsp;&nbsp;<b>Gap {} ({:.1f} cm)</b>".format(idx, gap_length_cm)
                )
            
            if areas_in_gap:
                area_links = []
                for area in areas_in_gap:
                    area_link = output.linkify(area.Id)
                    area_name_param = area.LookupParameter("Name")
                    area_name = area_name_param.AsString() if area_name_param else "Unnamed"
                    area_number_param = area.LookupParameter("Number")
                    area_number = area_number_param.AsString() if area_number_param else ""
                    area_links.append(
                        "{} {} - {}".format(
                            area_link,
                            area_number,
                            area_name
                        )
                    )
                if area_links:
                    output.print_html(
                        "&nbsp;&nbsp;&nbsp;&nbsp;<i>Affected areas:</i> {}".format(
                            ", ".join(area_links)
                        )
                    )
        
        for area_data in nonloc_areas:
            area = area_data.get('area')
            issues = area_data.get('issues', [])
            if not area:
                continue
            area_link = output.linkify(area.Id)
            area_name_param = area.LookupParameter("Name")
            area_name = area_name_param.AsString() if area_name_param else "Unnamed"
            area_number_param = area.LookupParameter("Number")
            area_number = area_number_param.AsString() if area_number_param else ""
            output.print_html(
                "&nbsp;&nbsp;<b>{}</b> {} - {} | {}".format(
                    area_link,
                    area_number,
                    area_name,
                    "; ".join(issues)
                )
            )
            
            problematic_elements = area_data.get('problematic_elements', [])
            if problematic_elements:
                boundary_links = []
                for elem_id in problematic_elements:
                    if elem_id and elem_id != DB.ElementId.InvalidElementId:
                        elem = doc.GetElement(elem_id)
                        if elem:
                            elem_link = output.linkify(elem_id)
                            elem_type = elem.Category.Name if elem.Category else "Element"
                            boundary_links.append("{} ({})".format(elem_link, elem_type))
                if boundary_links:
                    output.print_html(
                        "&nbsp;&nbsp;&nbsp;&nbsp;<i>Boundary elements causing issues:</i> {}".format(
                            ", ".join(boundary_links)
                        )
                    )
        
        for area_data in nonloc_areas:
            area = area_data.get('area')
            if not area:
                continue
            location = getattr(area, 'Location', None)
            if not location or not hasattr(location, 'Point'):
                continue
            center = location.Point
            elem_ids = [eid for eid in (area_data.get('problematic_elements') or [])
                        if eid and eid != DB.ElementId.InvalidElementId]
            gap_groups_for_view.append(
                {
                    'center': center,
                    'areas': set([area]),
                    'boundary_elements': set(elem_ids),
                }
            )
        
        if gap_groups_for_view:
            gaps_by_view[view.Id] = gap_groups_for_view
    
    # Separator before summary
    output.print_html('<hr style="margin:15px 0;border:none;border-top:1px solid #ccc;">')
    
    # Summary - condensed
    if total_bad == 0:
        output.print_html('<p style="color:green;margin:0;">All {} areas have valid boundaries.</p>'.format(total_areas))
        return None, None, None
    else:
        total_gaps = sum(len(gs) for gs in gaps_by_view.values())
        
        # Apply temporary visibility using Revit's Temporary View Properties mode
        # This shows a purple frame and auto-restores when disabled
        temp_view_ids = []
        with revit.Transaction("Enable Temporary View Properties"):
            for view in selected_views:
                if apply_temporary_visibility(doc, view):
                    temp_view_ids.append(view.Id)
        
        output.print_html(
            '<p style="margin:0;"><b>Summary:</b> {} area(s) with bad boundaries (out of {} total), {} distinct gap(s).<br>'
            '<b>Visualization active</b> - Red circles mark gap locations.<br>'
            '<span style="color:#9370DB;"><b>Temporary View Properties enabled</b> (purple frame). '
            'Original settings will be restored when you click Done.</span><br>'
            'Click element links above to navigate. Click Done when finished.</p>'.format(
                total_bad, total_areas, total_gaps
            )
        )
        
        return total_bad, gaps_by_view, temp_view_ids


# ============================================================
# RUN
# ============================================================

result = detect_bad_boundaries()

if result and result[0] is not None:
    total_bad, gaps_by_view, temp_view_ids = result
    uidoc = HOST_APP.uidoc
    
    # Track initially open views
    _initial_open_view_ids = set(uiv.ViewId for uiv in uidoc.GetOpenUIViews()) if uidoc else set()
    _candidate_view_ids = set(gaps_by_view.keys())
    
    # Pre-open all views with gaps (improves first-click zoom behavior)
    if uidoc and _candidate_view_ids:
        original_view = uidoc.ActiveView
        for vid in _candidate_view_ids:
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
        if original_view:
            try:
                uidoc.ActiveView = original_view
            except Exception:
                pass
    
    # Show modeless dialog
    server = revit.dc3dserver.Server(register=False)
    main_window = VisualizationControlWindow(total_bad, gaps_by_view, server, temp_view_ids)
    main_window.show()
