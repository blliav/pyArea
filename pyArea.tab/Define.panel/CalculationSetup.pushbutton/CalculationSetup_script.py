# -*- coding: utf-8 -*-
"""Define pyArea Schema Data - Hierarchy Manager
Manages the complete hierarchy: AreaScheme > Sheet > AreaPlan > RepresentedAreaPlans
"""

import sys
import os
from pyrevit import revit, DB, forms, script
from collections import OrderedDict

# Add lib folder to path
lib_path = os.path.join(os.path.dirname(__file__), "..", "..", "lib")
if lib_path not in sys.path:
    sys.path.insert(0, lib_path)

import data_manager
from schemas import municipality_schemas

# Import WPF
import clr
clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
import System
from System import Int64
from System.Windows import Window
from System.Windows.Controls import TextBox, ComboBox, CheckBox, StackPanel, Grid, TextBlock, Button, RowDefinition, ColumnDefinition
from System.Windows.Media import VisualTreeHelper
from System.Collections.ObjectModel import ObservableCollection


class TreeNode(object):
    """Represents a node in the hierarchy tree"""
    
    def __init__(self, element, element_type, display_name, parent=None):
        self.Element = element  # Revit element
        self.ElementType = element_type  # "AreaScheme", "Sheet", "AreaPlan", "RepresentedAreaPlan"
        self.DisplayName = display_name
        self.Parent = parent
        self.Children = ObservableCollection[TreeNode]()
        self.Icon = self._get_icon()
        self.Status = ""
        self.FontWeight = "Normal"
        
    def _get_icon(self):
        """Get icon for element type"""
        icons = {
            "AreaScheme": "📐",
            "Sheet": "📄",
            "AreaPlan": "■",  # Solid square - on sheet
            "AreaPlan_NotOnSheet": "□",  # Hollow square - not on sheet
            "RepresentedAreaPlan": "🔗"
        }
        return icons.get(self.ElementType, "📦")
    
    def add_child(self, child_node):
        """Add a child node"""
        child_node.Parent = self
        self.Children.Add(child_node)
        return child_node
    
    def remove_child(self, child_node):
        """Remove a child node"""
        if child_node in self.Children:
            self.Children.Remove(child_node)
            child_node.Parent = None


class CalculationSetupWindow(forms.WPFWindow):
    """Hierarchy Manager Dialog"""
    
    def __init__(self):
        forms.WPFWindow.__init__(self, 'CalculationSetupWindow.xaml')
        
        self._doc = revit.doc
        self._field_controls = {}
        self._selected_node = None
        self._tree_nodes = ObservableCollection[TreeNode]()
        
        # Wire up events
        self.tree_hierarchy.SelectedItemChanged += self.on_tree_selection_changed
        self.tree_hierarchy.MouseLeftButtonDown += self.on_tree_mouse_down
        self.btn_add.Click += self.on_add_clicked
        self.btn_remove.Click += self.on_remove_clicked
        self.btn_close.Click += self.on_close_clicked
        
        # Run cleanup on startup to fix any existing nested represented views
        self._cleanup_nested_represented_views()
        
        # Build initial tree
        self.build_tree()
        
        # Set initial button text
        self._update_add_button_text()
        
        # Load saved expansion state or expand all by default
        self._restore_expansion_state()
        
        # Apply context awareness AFTER tree is expanded (preselect based on selection or active view)
        self._apply_context_awareness()
        
    def _cleanup_nested_represented_views(self):
        """Clean up any existing nested represented views and remove empty RepresentedViews arrays"""
        try:
            # Get all views
            collector = DB.FilteredElementCollector(self._doc)
            all_views = collector.OfClass(DB.View).ToElements()
            
            # Build set of views that are on sheets
            views_on_sheets = set()
            sheets_collector = DB.FilteredElementCollector(self._doc)
            sheets = sheets_collector.OfClass(DB.ViewSheet).ToElements()
            for sheet in sheets:
                try:
                    view_ids = sheet.GetAllPlacedViews()
                    for vid in view_ids:
                        views_on_sheets.add(vid)
                except:
                    pass
            
            changes_made = False
            
            with revit.Transaction("Cleanup Nested RepresentedViews"):
                for view in all_views:
                    view_data = data_manager.get_data(view)
                    if not view_data or "RepresentedViews" not in view_data:
                        continue
                    
                    represented_ids = view_data.get("RepresentedViews", [])
                    if not represented_ids:
                        # Remove empty RepresentedViews array
                        print("  - Removing empty RepresentedViews from '{}' (ID: {})".format(
                            view.Name if hasattr(view, 'Name') else "?",
                            view.Id.Value
                        ))
                        view_data.pop("RepresentedViews", None)
                        data_manager.set_data(view, view_data)
                        changes_made = True
                        continue
                    
                    # NOTE: Parent views (with RepresentedViews) CAN be on sheets
                    # We only need to validate that the REPRESENTED views themselves aren't on sheets
                    
                    # Check for nested represented views and flatten them
                    all_represented_ids = list(represented_ids)  # Start with direct children
                    ids_to_clean = []
                    
                    for rep_id in represented_ids:
                        try:
                            rep_view = self._doc.GetElement(DB.ElementId(Int64(int(rep_id))))
                            if not rep_view:
                                continue
                            
                            # Check if represented view is on a sheet (invalid)
                            if rep_view.Id in views_on_sheets:
                                print("  - Removing '{}' (ID: {}) from represented list - it's on a sheet".format(
                                    rep_view.Name if hasattr(rep_view, 'Name') else "?",
                                    rep_id
                                ))
                                ids_to_clean.append(rep_id)
                                continue
                            
                            # Check if represented view has its own represented views (nested)
                            rep_data = data_manager.get_data(rep_view)
                            if rep_data and "RepresentedViews" in rep_data:
                                nested_ids = rep_data.get("RepresentedViews", [])
                                if nested_ids:
                                    print("  - Flattening nested represented views from '{}' (ID: {})".format(
                                        rep_view.Name if hasattr(rep_view, 'Name') else "?",
                                        rep_id
                                    ))
                                    # Add nested views to parent's list
                                    for nested_id in nested_ids:
                                        if nested_id not in all_represented_ids:
                                            all_represented_ids.append(nested_id)
                                    
                                    # Remove RepresentedViews from child
                                    rep_data.pop("RepresentedViews", None)
                                    data_manager.set_data(rep_view, rep_data)
                                    changes_made = True
                                elif "RepresentedViews" in rep_data:
                                    # Remove empty RepresentedViews array
                                    rep_data.pop("RepresentedViews", None)
                                    data_manager.set_data(rep_view, rep_data)
                                    changes_made = True
                        except:
                            pass
                    
                    # Remove invalid IDs (views on sheets)
                    for rep_id in ids_to_clean:
                        if rep_id in all_represented_ids:
                            all_represented_ids.remove(rep_id)
                    
                    # Update parent if list changed
                    if set(all_represented_ids) != set(represented_ids) or ids_to_clean:
                        if all_represented_ids:
                            view_data["RepresentedViews"] = all_represented_ids
                        else:
                            view_data.pop("RepresentedViews", None)
                        data_manager.set_data(view, view_data)
                        changes_made = True
            
            # Cleanup completed silently
            pass
        
        except Exception as e:
            print("Error during cleanup: {}".format(e))
    
    def _get_context_element(self):
        """Get context element from selection or active view
        
        Returns:
            tuple: (element, element_type) where element_type is "view" or "sheet"
                   Returns (None, None) if no valid context
        """
        try:
            # Get current selection
            selection = revit.uidoc.Selection
            selected_ids = selection.GetElementIds()
            
            # Priority 1: Check if a view or sheet is selected in project browser
            # or if a viewport is selected on a sheet
            for elem_id in selected_ids:
                elem = self._doc.GetElement(elem_id)
                
                # Check if it's a viewport (view on sheet)
                if isinstance(elem, DB.Viewport):
                    view_id = elem.ViewId
                    view = self._doc.GetElement(view_id)
                    # Check if it's an area plan (views on sheets are shown even without explicit data)
                    if hasattr(view, 'AreaScheme') and view.AreaScheme:
                        # Check if the area scheme has a municipality (only defined schemes are shown)
                        if data_manager.get_municipality(view.AreaScheme):
                            return (view, "view")
                
                # Check if it's a view (selected in project browser)
                if isinstance(elem, DB.View) and not isinstance(elem, DB.ViewSheet):
                    if hasattr(elem, 'AreaScheme') and elem.AreaScheme:
                        # Must have municipality and either be on a sheet or have explicit data
                        if data_manager.get_municipality(elem.AreaScheme):
                            # Check if it's on a sheet OR has explicit data
                            if data_manager.has_data(elem) or self._is_view_on_sheet(elem):
                                return (elem, "view")
                
                # Check if it's a sheet with data
                if isinstance(elem, DB.ViewSheet):
                    if data_manager.has_data(elem):
                        return (elem, "sheet")
            
            # Priority 2: Check active view if nothing is selected
            active_view = revit.uidoc.ActiveView
            
            # Check if active view is a sheet with data
            if isinstance(active_view, DB.ViewSheet):
                if data_manager.has_data(active_view):
                    return (active_view, "sheet")
            
            # Check if active view is an area plan
            if hasattr(active_view, 'AreaScheme') and active_view.AreaScheme:
                # Must have municipality and either be on a sheet or have explicit data
                if data_manager.get_municipality(active_view.AreaScheme):
                    if data_manager.has_data(active_view) or self._is_view_on_sheet(active_view):
                        return (active_view, "view")
            
        except Exception as e:
            pass  # Silently fail
        
        return (None, None)
    
    def _is_view_on_sheet(self, view):
        """Check if a view is placed on any sheet
        
        Args:
            view: View element to check
            
        Returns:
            bool: True if view is on a sheet, False otherwise
        """
        try:
            sheets_collector = DB.FilteredElementCollector(self._doc)
            sheets = sheets_collector.OfClass(DB.ViewSheet).ToElements()
            for sheet in sheets:
                try:
                    view_ids = sheet.GetAllPlacedViews()
                    if view.Id in view_ids:
                        return True
                except:
                    pass
        except:
            pass
        
        return False
    
    def _find_node_by_element_id(self, element_id):
        """Find a node in the tree by element ID
        
        Args:
            element_id: Revit ElementId to search for
            
        Returns:
            TreeNode if found, None otherwise
        """
        def search_node(node):
            """Recursively search through node and children"""
            if node.Element.Id == element_id:
                return node
            
            for child in node.Children:
                result = search_node(child)
                if result:
                    return result
            
            return None
        
        # Search through all root nodes
        for root_node in self._tree_nodes:
            result = search_node(root_node)
            if result:
                return result
        
        return None
    
    def _select_and_expand_node(self, target_node):
        """Select and expand a node in the tree
        
        Args:
            target_node: TreeNode to select
        """
        try:
            import System.Windows.Threading as Threading
            
            def do_select():
                try:
                    # Build path from root to target
                    path_nodes = []
                    current = target_node
                    while current:
                        path_nodes.insert(0, current)
                        current = current.Parent
                    
                    # Expand all parent nodes (not the target itself)
                    for i in range(len(path_nodes) - 1):
                        node = path_nodes[i]
                        container = self._get_container_for_node_simple(node)
                        if container:
                            if not container.IsExpanded:
                                container.IsExpanded = True
                                container.UpdateLayout()
                    
                    # Select the target node
                    target_container = self._get_container_for_node_simple(target_node)
                    if target_container:
                        target_container.IsSelected = True
                        target_container.BringIntoView()
                
                except Exception as e:
                    pass  # Silently fail
            
            # Use Dispatcher to delay selection until after expansion is complete
            self.tree_hierarchy.Dispatcher.BeginInvoke(
                Threading.DispatcherPriority.ContextIdle,
                System.Action(do_select)
            )
        
        except Exception as e:
            pass  # Silently fail
    
    def _get_container_for_node_simple(self, node):
        """Get TreeViewItem container using TreeView's own methods
        
        Args:
            node: TreeNode to find container for
            
        Returns:
            TreeViewItem container or None
        """
        try:
            # For root nodes
            if not node.Parent:
                for i in range(self.tree_hierarchy.Items.Count):
                    if self.tree_hierarchy.Items[i] == node:
                        container = self.tree_hierarchy.ItemContainerGenerator.ContainerFromItem(node)
                        return container
            else:
                # For child nodes, get parent container first
                parent_container = self._get_container_for_node_simple(node.Parent)
                if parent_container and parent_container.ItemContainerGenerator:
                    return parent_container.ItemContainerGenerator.ContainerFromItem(node)
        except:
            pass
        
        return None
    
    def _apply_context_awareness(self):
        """Apply context awareness by detecting and selecting the current view/sheet"""
        try:
            # Get context element
            context_elem, context_type = self._get_context_element()
            
            if not context_elem:
                return  # No context to apply
            
            # Find the node in the tree
            node = self._find_node_by_element_id(context_elem.Id)
            
            if node:
                # Select and expand to this node
                self._select_and_expand_node(node)
        
        except Exception as e:
            pass  # Silently fail - don't disrupt normal workflow
    
    def _reselect_after_add(self, element_id):
        """Re-select an element after adding it to the tree
        
        Args:
            element_id: ElementId of the newly added element to select
        """
        try:
            import System.Windows.Threading as Threading
            
            def do_reselect():
                try:
                    node = self._find_node_by_element_id(element_id)
                    if node:
                        self._select_and_expand_node(node)
                except:
                    pass
            
            # Use Dispatcher to delay selection until tree is fully rendered
            self.tree_hierarchy.Dispatcher.BeginInvoke(
                Threading.DispatcherPriority.ContextIdle,
                System.Action(do_reselect)
            )
        except:
            pass  # Silently fail
    
    def rebuild_tree(self):
        """Rebuild tree and restore expansion state"""
        self.build_tree()
        self._restore_expansion_state()
    
    def build_tree(self):
        """Build the hierarchy tree from Revit elements"""
        self._tree_nodes.Clear()
        
        # Get all AreaSchemes
        collector = DB.FilteredElementCollector(self._doc)
        area_schemes = collector.OfClass(DB.AreaScheme).ToElements()
        
        for area_scheme in area_schemes:
            # Get municipality - only show if defined
            municipality = data_manager.get_municipality(area_scheme)
            if not municipality:
                continue  # Skip undefined AreaSchemes
            
            # Create AreaScheme node
            scheme_node = TreeNode(
                area_scheme,
                "AreaScheme",
                area_scheme.Name
            )
            
            scheme_node.FontWeight = "Bold"
            
            # Get sheets for this AreaScheme
            self._add_sheets_to_scheme(scheme_node)
            
            self._tree_nodes.Add(scheme_node)
        
        # Set tree source
        self.tree_hierarchy.ItemsSource = self._tree_nodes
    
    def _expand_all_nodes(self):
        """Expand all tree nodes"""
        try:
            import System.Windows.Threading as Threading
            
            def do_expand():
                try:
                    # Expand all top-level items
                    for i in range(self.tree_hierarchy.Items.Count):
                        container = self.tree_hierarchy.ItemContainerGenerator.ContainerFromIndex(i)
                        if container:
                            self._expand_node_recursive(container)
                except:
                    pass
            
            # Use Dispatcher to delay expansion until UI is ready
            self.tree_hierarchy.Dispatcher.BeginInvoke(
                Threading.DispatcherPriority.Background,
                System.Action(do_expand)
            )
        except:
            pass  # Silently fail if expansion doesn't work
    
    def _expand_node_recursive(self, item_container):
        """Recursively expand a tree node and its children"""
        try:
            item_container.IsExpanded = True
            item_container.UpdateLayout()
            
            # Expand children
            if hasattr(item_container, 'Items'):
                for child_item in item_container.Items:
                    child_container = item_container.ItemContainerGenerator.ContainerFromItem(child_item)
                    if child_container:
                        self._expand_node_recursive(child_container)
        except:
            pass
    
    def _add_sheets_to_scheme(self, scheme_node):
        """Add sheets and AreaPlans that belong to this AreaScheme"""
        area_scheme = scheme_node.Element
        area_scheme_id = str(area_scheme.Id.Value)
        
        # Get all sheets
        collector = DB.FilteredElementCollector(self._doc)
        sheets = collector.OfClass(DB.ViewSheet).ToElements()
        
        # Build set of views that are on sheets
        views_on_sheets = set()
        for sheet in sheets:
            try:
                view_ids = sheet.GetAllPlacedViews()
                for vid in view_ids:
                    views_on_sheets.add(vid)
            except:
                pass
        
        # Add sheets
        for sheet in sheets:
            # Check if sheet belongs to this AreaScheme
            sheet_data = data_manager.get_data(sheet)
            if sheet_data and sheet_data.get("AreaSchemeId") == area_scheme_id:
                sheet_name = "{} - {}".format(
                    sheet.SheetNumber if hasattr(sheet, 'SheetNumber') else "?",
                    sheet.Name if hasattr(sheet, 'Name') else "Unnamed"
                )
                sheet_node = scheme_node.add_child(TreeNode(
                    sheet,
                    "Sheet",
                    sheet_name
                ))
                
                # Add AreaPlans on this sheet (indented under sheet)
                self._add_views_to_sheet(sheet_node, area_scheme, views_on_sheets)
        
        # Add AreaPlans that have data but are NOT on any sheet (at scheme level)
        self._add_standalone_views(scheme_node, area_scheme, views_on_sheets)
    
    def _add_views_to_sheet(self, sheet_node, area_scheme, views_on_sheets):
        """Add AreaPlan views that are on this sheet"""
        try:
            view_ids = sheet_node.Element.GetAllPlacedViews()
            
            # Collect views first
            views_to_add = []
            for view_id in view_ids:
                view = self._doc.GetElement(view_id)
                
                # Check if it's an AreaPlan view with matching AreaScheme
                if hasattr(view, 'AreaScheme') and view.AreaScheme and view.AreaScheme.Id == area_scheme.Id:
                    views_to_add.append(view)
            
            # Sort by elevation (Z coordinate of view origin)
            views_to_add.sort(key=lambda v: v.Origin.Z if hasattr(v, 'Origin') else 0)
            
            # Add sorted views to tree
            for view in views_to_add:
                view_name = view.Name if hasattr(view, 'Name') else "Unnamed View"
                view_node = sheet_node.add_child(TreeNode(
                    view,
                    "AreaPlan",  # Solid square - on sheet
                    view_name
                ))
                
                # Add RepresentedViews
                self._add_represented_views(view_node)
        except:
            pass
    
    def _add_standalone_views(self, scheme_node, area_scheme, views_on_sheets):
        """Add AreaPlan views with data that are NOT on any sheet"""
        # Get all views
        collector = DB.FilteredElementCollector(self._doc)
        all_views = collector.OfClass(DB.View).ToElements()
        
        # Collect views that meet criteria first
        views_to_add = []
        for view in all_views:
            try:
                # Must be AreaPlan with matching scheme
                if not hasattr(view, 'AreaScheme'):
                    continue
                if not view.AreaScheme or view.AreaScheme.Id != area_scheme.Id:
                    continue
                
                # Must have data (user added it)
                if not data_manager.has_data(view):
                    continue
                
                # Must NOT be on any sheet
                if view.Id in views_on_sheets:
                    continue
                
                # Must NOT be used as RepresentedView
                # (Check all views to see if this view is in their RepresentedViews list)
                is_represented = False
                for other_view in all_views:
                    other_data = data_manager.get_data(other_view)
                    if other_data and "RepresentedViews" in other_data:
                        rep_ids = other_data.get("RepresentedViews", [])
                        if str(view.Id.Value) in rep_ids:
                            is_represented = True
                            break
                
                if is_represented:
                    continue
                
                # Add to collection
                views_to_add.append(view)
            except:
                continue
        
        # Sort by elevation (Z coordinate of view origin)
        views_to_add.sort(key=lambda v: v.Origin.Z if hasattr(v, 'Origin') else 0)
        
        # Add sorted views to tree
        for view in views_to_add:
            view_name = view.Name if hasattr(view, 'Name') else "Unnamed View"
            view_node = scheme_node.add_child(TreeNode(
                view,
                "AreaPlan_NotOnSheet",  # Hollow square - not on sheet
                view_name
            ))
            
            # These can also have RepresentedViews
            self._add_represented_views(view_node)
    
    def _add_represented_views(self, view_node):
        """Add represented area plans for this AreaPlan"""
        view_data = data_manager.get_data(view_node.Element)
        if view_data and "RepresentedViews" in view_data:
            represented_ids = view_data.get("RepresentedViews", [])
            
            # Build set of views that are on sheets (to detect edge case)
            views_on_sheets = set()
            collector = DB.FilteredElementCollector(self._doc)
            sheets = collector.OfClass(DB.ViewSheet).ToElements()
            for sheet in sheets:
                try:
                    view_ids = sheet.GetAllPlacedViews()
                    for vid in view_ids:
                        views_on_sheets.add(vid)
                except:
                    pass
            
            # Track which IDs to remove (views that are now on sheets)
            ids_to_remove = []
            valid_rep_views = []
            
            for rep_id in represented_ids:
                try:
                    rep_view = self._doc.GetElement(DB.ElementId(Int64(int(rep_id))))
                    if rep_view:
                        # EDGE CASE: Check if this represented view is actually on a sheet
                        if rep_view.Id in views_on_sheets:
                            # This view is now on a sheet, should not be a represented view
                            ids_to_remove.append(rep_id)
                            # Also clean up the represented view's own RepresentedViews data
                            rep_data = data_manager.get_data(rep_view)
                            if rep_data and "RepresentedViews" in rep_data:
                                rep_data.pop("RepresentedViews", None)
                                with revit.Transaction("Clean up nested RepresentedViews"):
                                    data_manager.set_data(rep_view, rep_data)
                        else:
                            # Valid represented view - collect for sorting
                            valid_rep_views.append(rep_view)
                except:
                    pass
            
            # Sort represented views by elevation
            valid_rep_views.sort(key=lambda v: v.Origin.Z if hasattr(v, 'Origin') else 0)
            
            # Add sorted represented views to tree
            for rep_view in valid_rep_views:
                rep_name = rep_view.Name if hasattr(rep_view, 'Name') else "Unnamed"
                view_node.add_child(TreeNode(
                    rep_view,
                    "RepresentedAreaPlan",
                    rep_name
                ))
            
            # Clean up: remove invalid represented view IDs
            if ids_to_remove:
                for rep_id in ids_to_remove:
                    represented_ids.remove(rep_id)
                view_data["RepresentedViews"] = represented_ids
                with revit.Transaction("Clean up invalid RepresentedViews"):
                    data_manager.set_data(view_node.Element, view_data)
    
    def on_tree_mouse_down(self, sender, args):
        """Handle mouse click on tree - deselect if clicking empty space"""
        try:
            # Check if sender is TreeViewItem - if so, we clicked on an item
            if isinstance(sender, System.Windows.Controls.TreeViewItem):
                return
            
            # We clicked on the TreeView background - clear selection
            # Need to set IsSelected = False on the container, not just SelectedItem = None
            if self.tree_hierarchy.SelectedItem:
                # Get the container for the selected item
                container = self.tree_hierarchy.ItemContainerGenerator.ContainerFromItem(
                    self.tree_hierarchy.SelectedItem
                )
                if container:
                    container.IsSelected = False
        except:
            pass
    
    def on_tree_selection_changed(self, sender, args):
        """Handle tree selection change"""
        selected_item = self.tree_hierarchy.SelectedItem
        if not selected_item:
            self._selected_node = None
            self._update_add_button_text()
            self._clear_properties_panel()
            return
        
        self._selected_node = selected_item
        self._update_add_button_text()
        self.update_properties_panel()
    
    def _clear_properties_panel(self):
        """Clear the properties panel when nothing is selected"""
        self.text_fields_title.Text = "Select an element from the tree"
        self.text_fields_subtitle.Text = ""
        self.panel_fields.Children.Clear()
        self._field_controls = {}
        self.text_json.Text = "Select an element to view its JSON data..."
        self.text_json.Foreground = System.Windows.Media.Brushes.Gray
        self.text_json.Background = System.Windows.Media.Brushes.LightGray
    
    def _update_add_button_text(self):
        """Update Add and Remove button text and enabled state based on selection"""
        if not self._selected_node:
            self.btn_add.Content = "➕ Scheme"
            self.btn_add.IsEnabled = True
            self.btn_remove.IsEnabled = False
        elif self._selected_node.ElementType == "AreaScheme":
            self.btn_add.Content = "➕ Sheet"
            self.btn_add.IsEnabled = True
            self.btn_remove.IsEnabled = True
        elif self._selected_node.ElementType == "Sheet":
            self.btn_add.Content = "➕ AreaPlan"
            self.btn_add.IsEnabled = True
            self.btn_remove.IsEnabled = True
        elif self._selected_node.ElementType == "AreaPlan":
            # AreaPlan on sheet - can add RepresentedViews but can't remove (it's on a sheet)
            self.btn_add.Content = "➕ Represented AreaPlan"
            self.btn_add.IsEnabled = True
            self.btn_remove.IsEnabled = False
        elif self._selected_node.ElementType == "AreaPlan_NotOnSheet":
            # AreaPlan not on sheet - can set representing view or remove
            self.btn_add.Content = "🔗 Set Representing View"
            self.btn_add.IsEnabled = True
            self.btn_remove.IsEnabled = True
        elif self._selected_node.ElementType == "RepresentedAreaPlan":
            # RepresentedAreaPlans can be moved to a different parent or removed
            self.btn_add.Content = "🔗 Set Representing View"
            self.btn_add.IsEnabled = True
            self.btn_remove.IsEnabled = True
        else:
            self.btn_add.Content = "➕"
            self.btn_add.IsEnabled = True
            self.btn_remove.IsEnabled = True
    
    def update_properties_panel(self):
        """Update the right panel with selected element's properties"""
        if not self._selected_node:
            return
        
        node = self._selected_node
        
        # Get municipality and variant
        municipality = self._get_municipality_for_node(node)
        variant = self._get_variant_for_node(node)
        
        # Update title with element name on first line, details on second line
        self._update_fields_title(node.DisplayName, node.ElementType, municipality, variant)
        
        # Update JSON viewer
        self._update_json_viewer(node)
        
        # Clear fields
        self.panel_fields.Children.Clear()
        self._field_controls = {}
        
        # Build fields based on element type
        self._build_fields_for_node(node)
        
        # Auto-save default values for new AreaSchemes (no existing data)
        if node.ElementType == "AreaScheme" and not data_manager.has_data(node.Element):
            self._save_default_areascheme_values()
    
    def _get_municipality_for_node(self, node):
        """Get municipality for a node"""
        if node.ElementType == "AreaScheme":
            return data_manager.get_municipality(node.Element)
        elif node.ElementType == "Sheet":
            return data_manager.get_municipality_from_sheet(self._doc, node.Element)
        elif node.ElementType in ["AreaPlan", "AreaPlan_NotOnSheet", "RepresentedAreaPlan"]:
            # get_municipality_from_view returns (municipality, variant) tuple
            municipality, variant = data_manager.get_municipality_from_view(self._doc, node.Element)
            return municipality
        return None
    
    def _get_variant_for_node(self, node):
        """Get variant for a node"""
        if node.ElementType == "AreaScheme":
            return data_manager.get_variant(node.Element)
        elif node.ElementType == "Sheet":
            # Sheets inherit variant from their AreaScheme
            sheet_data = data_manager.get_data(node.Element)
            if sheet_data and "AreaSchemeId" in sheet_data:
                area_scheme_id = DB.ElementId(Int64(int(sheet_data["AreaSchemeId"])))
                area_scheme = self._doc.GetElement(area_scheme_id)
                if area_scheme:
                    return data_manager.get_variant(area_scheme)
        elif node.ElementType in ["AreaPlan", "AreaPlan_NotOnSheet", "RepresentedAreaPlan"]:
            # get_municipality_from_view returns (municipality, variant) tuple
            municipality, variant = data_manager.get_municipality_from_view(self._doc, node.Element)
            return variant
        return None
    
    def _update_fields_title(self, name, element_type, municipality, variant=None):
        """Update the fields panel title with name and details in separate TextBlocks"""
        
        # Set the element name (bold)
        self.text_fields_title.Text = name
        
        # Build the details text (Type | Municipality | Variant)
        details_parts = [element_type]
        if municipality:
            details_parts.append(municipality)
            if variant and variant != "Default":
                details_parts.append(variant)
        
        details_text = " | ".join(details_parts)
        self.text_fields_subtitle.Text = details_text
    
    def _build_fields_for_node(self, node):
        """Build input fields for the selected node"""
        municipality = self._get_municipality_for_node(node)
        
        # Get field definitions
        if node.ElementType == "AreaScheme":
            fields = municipality_schemas.AREASCHEME_FIELDS
        elif node.ElementType == "Sheet":
            if not municipality:
                self._show_no_municipality_message()
                return
            fields = municipality_schemas.SHEET_FIELDS.get(municipality, {})
        elif node.ElementType in ["AreaPlan", "AreaPlan_NotOnSheet", "RepresentedAreaPlan"]:
            if not municipality:
                self._show_no_municipality_message()
                return
            # RepresentedAreaPlans are AreaPlans too, just referenced by another view
            # They have all the same fields EXCEPT RepresentedAreaPlans (no nesting)
            fields = municipality_schemas.AREAPLAN_FIELDS.get(municipality, {})
        else:
            return
        
        # Load existing data
        existing_data = data_manager.get_data(node.Element) or {}
        
        # Create field controls (skip internal fields)
        for field_name, field_props in fields.items():
            # Skip internal fields that shouldn't be shown to user
            if field_name == "AreaSchemeId":
                continue
            # Skip RepresentedViews field - managed via Add/Remove buttons, not direct editing
            if field_name == "RepresentedViews":
                continue
            self._create_field_control(field_name, field_props, existing_data.get(field_name))
    
    def _show_no_municipality_message(self):
        """Show message when municipality is not defined"""
        msg = TextBlock()
        msg.Text = "No municipality defined. Please define AreaScheme first."
        msg.Foreground = System.Windows.Media.Brushes.Red
        msg.FontWeight = System.Windows.FontWeights.Bold
        self.panel_fields.Children.Add(msg)
    
    def _create_field_control(self, field_name, field_props, current_value):
        """Create a field control with horizontal layout: label left, input right"""
        # Main container grid
        main_grid = Grid()
        main_grid.Margin = System.Windows.Thickness(0, 4, 0, 4)
        
        # Define columns: Label column, Input column
        main_grid.ColumnDefinitions.Add(ColumnDefinition())
        main_grid.ColumnDefinitions.Add(ColumnDefinition())
        main_grid.ColumnDefinitions[0].Width = System.Windows.GridLength(140)  # Fixed width for labels
        main_grid.ColumnDefinitions[1].Width = System.Windows.GridLength(1, System.Windows.GridUnitType.Star)
        
        # Label column - StackPanel with English on top, Hebrew on bottom
        label_panel = StackPanel()
        label_panel.Orientation = System.Windows.Controls.Orientation.Vertical
        label_panel.VerticalAlignment = System.Windows.VerticalAlignment.Center
        Grid.SetColumn(label_panel, 0)
        
        # Top row: English label with required indicator
        top_panel = StackPanel()
        top_panel.Orientation = System.Windows.Controls.Orientation.Horizontal
        
        # English label
        label_en = TextBlock()
        label_en.Text = field_name
        label_en.FontSize = 10
        label_en.FontWeight = System.Windows.FontWeights.SemiBold
        label_en.Foreground = System.Windows.Media.Brushes.Black
        label_en.ToolTip = field_props.get("description", "")
        label_en.Margin = System.Windows.Thickness(0, 0, 3, 0)
        top_panel.Children.Add(label_en)
        
        # Required indicator
        if field_props.get("required", False):
            required_label = TextBlock()
            required_label.Text = "*"
            required_label.FontSize = 10
            required_label.FontWeight = System.Windows.FontWeights.Bold
            required_label.Foreground = System.Windows.Media.Brushes.Red
            required_label.Margin = System.Windows.Thickness(0, 0, 0, 0)
            top_panel.Children.Add(required_label)
        
        label_panel.Children.Add(top_panel)
        
        # Bottom row: Hebrew label (if available)
        hebrew_name = field_props.get("hebrew_name", "")
        if hebrew_name:
            label_he = TextBlock()
            label_he.Text = hebrew_name
            label_he.FontSize = 9
            label_he.FontWeight = System.Windows.FontWeights.Normal
            label_he.Foreground = System.Windows.Media.Brushes.Gray
            label_he.Margin = System.Windows.Thickness(0, 1, 0, 0)
            label_panel.Children.Add(label_he)
        
        main_grid.Children.Add(label_panel)
        
        # Get default value
        default_value = field_props.get("default", "")
        
        # Input control - create appropriate control based on field type
        field_type = field_props.get("type")
        
        if field_name == "Municipality" or field_name == "Variant" or (field_type == "string" and "options" in field_props):
            # ComboBox for Municipality, Variant, or options
            combo = ComboBox()
            combo.FontSize = 11
            combo.Height = 26
            combo.Margin = System.Windows.Thickness(5, 0, 0, 0)
            combo.VerticalAlignment = System.Windows.VerticalAlignment.Center
            if field_name == "Municipality":
                for muni in ["Common", "Jerusalem", "Tel-Aviv"]:
                    combo.Items.Add(muni)
                if current_value:
                    combo.SelectedItem = current_value
                else:
                    combo.SelectedIndex = 0
                # Wire up handler to update Variant dropdown when Municipality changes
                combo.SelectionChanged += self.on_municipality_changed
            elif field_name == "Variant":
                # Variant options depend on Municipality
                # Get current municipality value from the selected node
                node_data = data_manager.get_data(self._selected_node.Element) or {}
                municipality_value = node_data.get("Municipality", "Common")
                variants = municipality_schemas.MUNICIPALITY_VARIANTS.get(municipality_value, ["Default"])
                for variant in variants:
                    combo.Items.Add(variant)
                if current_value:
                    combo.SelectedItem = current_value
                else:
                    combo.SelectedIndex = 0  # Default
            else:
                for option in field_props["options"]:
                    combo.Items.Add(option)
                if current_value:
                    combo.SelectedItem = current_value
                else:
                    combo.SelectedIndex = 0
            Grid.SetColumn(combo, 1)
            main_grid.Children.Add(combo)
            self._field_controls[field_name] = combo
            combo.SelectionChanged += self.on_field_changed
            
        elif field_name in ["IS_UNDERGROUND", "FLOOR_UNDERGROUND"]:
            # CheckBox for boolean fields - align to left to match textboxes
            checkbox = CheckBox()
            checkbox.HorizontalAlignment = System.Windows.HorizontalAlignment.Left
            checkbox.Margin = System.Windows.Thickness(5, 0, 0, 0)
            checkbox.VerticalAlignment = System.Windows.VerticalAlignment.Center
            if current_value:
                # Handle both "yes"/"no" strings and 1/0 integers
                if isinstance(current_value, str):
                    checkbox.IsChecked = current_value.lower() == "yes"
                else:
                    checkbox.IsChecked = bool(current_value)
            Grid.SetColumn(checkbox, 1)
            main_grid.Children.Add(checkbox)
            self._field_controls[field_name] = checkbox
            checkbox.Checked += self.on_field_changed
            checkbox.Unchecked += self.on_field_changed
            
        else:
            # Check if field supports placeholders
            field_placeholders = field_props.get("placeholders", [])
            has_placeholders = len(field_placeholders) > 0
            
            if has_placeholders:
                # Use editable ComboBox with placeholders
                combo = ComboBox()
                combo.IsEditable = True
                combo.FontSize = 11
                combo.Height = 26
                combo.Margin = System.Windows.Thickness(5, 0, 0, 0)
                combo.VerticalAlignment = System.Windows.VerticalAlignment.Center
                combo.ToolTip = field_props.get("description", "")
                
                # Add placeholder options
                for placeholder in field_placeholders:
                    combo.Items.Add(placeholder)
                
                # Set current value or default
                if current_value is not None:
                    combo.Text = str(current_value)
                elif default_value:
                    combo.Text = default_value
                    combo.Foreground = System.Windows.Media.Brushes.Gray
                    combo.Tag = "showing_default"
                
                # Create handlers with closure to capture default_value
                def create_combo_handlers(cb, def_val):
                    # Clear default on focus
                    def on_got_focus(sender, args):
                        if sender.Tag == "showing_default":
                            sender.Text = ""
                            sender.Foreground = System.Windows.Media.Brushes.Black
                            sender.Tag = None
                    
                    # Reset to default if empty on lost focus
                    def on_lost_focus(sender, args):
                        if not sender.Text or sender.Text.strip() == "":
                            if def_val:
                                sender.Text = def_val
                                sender.Foreground = System.Windows.Media.Brushes.Gray
                                sender.Tag = "showing_default"
                        self.on_field_changed(sender, args)
                    
                    return on_got_focus, on_lost_focus
                
                got_focus_handler, lost_focus_handler = create_combo_handlers(combo, default_value)
                combo.GotFocus += got_focus_handler
                combo.LostFocus += lost_focus_handler
                
                Grid.SetColumn(combo, 1)
                main_grid.Children.Add(combo)
                self._field_controls[field_name] = combo
                
                # Handle selection changes
                combo.SelectionChanged += self.on_field_changed
            else:
                # Regular TextBox for fields without placeholders
                textbox = TextBox()
                textbox.FontSize = 11
                textbox.Height = 26
                textbox.Margin = System.Windows.Thickness(5, 0, 0, 0)
                textbox.VerticalAlignment = System.Windows.VerticalAlignment.Center
                textbox.ToolTip = field_props.get("description", "")
                
                # Set value or show default in gray
                if current_value is not None:
                    textbox.Text = str(current_value)
                    textbox.Foreground = System.Windows.Media.Brushes.Black
                elif default_value:
                    textbox.Text = default_value
                    textbox.Foreground = System.Windows.Media.Brushes.Gray
                    textbox.Tag = "showing_default"
                
                # Create handlers with closure to capture default_value
                def create_textbox_handlers(tb, def_val):
                    # Clear default on focus
                    def on_got_focus(sender, args):
                        if sender.Tag == "showing_default":
                            sender.Text = ""
                            sender.Foreground = System.Windows.Media.Brushes.Black
                            sender.Tag = None
                    
                    # Reset to default if empty on lost focus
                    def on_lost_focus(sender, args):
                        if not sender.Text or sender.Text.strip() == "":
                            if def_val:
                                sender.Text = def_val
                                sender.Foreground = System.Windows.Media.Brushes.Gray
                                sender.Tag = "showing_default"
                        self.on_field_changed(sender, args)
                    
                    return on_got_focus, on_lost_focus
                
                got_focus_handler, lost_focus_handler = create_textbox_handlers(textbox, default_value)
                textbox.GotFocus += got_focus_handler
                textbox.LostFocus += lost_focus_handler
                
                Grid.SetColumn(textbox, 1)
                main_grid.Children.Add(textbox)
                self._field_controls[field_name] = textbox
        
        self.panel_fields.Children.Add(main_grid)
    
    def on_municipality_changed(self, sender, args):
        """Update Variant dropdown when Municipality changes"""
        if not self._selected_node:
            return
        
        # Get the new municipality value
        municipality_combo = self._field_controls.get("Municipality")
        variant_combo = self._field_controls.get("Variant")
        
        if not municipality_combo or not variant_combo:
            return
        
        selected_municipality = municipality_combo.SelectedItem
        if not selected_municipality:
            return
        
        # Get available variants for this municipality
        variants = municipality_schemas.MUNICIPALITY_VARIANTS.get(selected_municipality, ["Default"])
        
        # Store current selection
        current_variant = variant_combo.SelectedItem
        
        # Update Variant combo items
        variant_combo.Items.Clear()
        for variant in variants:
            variant_combo.Items.Add(variant)
        
        # Try to restore previous selection, or default to first item
        if current_variant in variants:
            variant_combo.SelectedItem = current_variant
        else:
            variant_combo.SelectedIndex = 0
        
        # Call the regular field changed handler to save
        self.on_field_changed(sender, args)
    
    def _save_default_areascheme_values(self):
        """Save default Municipality and Variant values for a new AreaScheme
        
        This is called automatically when displaying a new AreaScheme's properties
        to ensure the default dropdown values are saved even if the user doesn't
        interact with them.
        """
        if not self._selected_node or self._selected_node.ElementType != "AreaScheme":
            return
        
        # Collect Municipality and Variant from dropdowns
        data_dict = {}
        
        # Get Municipality value (should always have a default)
        if "Municipality" in self._field_controls:
            muni_control = self._field_controls["Municipality"]
            if isinstance(muni_control, ComboBox) and muni_control.SelectedItem:
                data_dict["Municipality"] = muni_control.SelectedItem
        
        # Get Variant value (should always have a default)
        if "Variant" in self._field_controls:
            variant_control = self._field_controls["Variant"]
            if isinstance(variant_control, ComboBox) and variant_control.SelectedItem:
                data_dict["Variant"] = variant_control.SelectedItem
        
        # Only save if we have at least Municipality
        if not data_dict.get("Municipality"):
            return
        
        # Save to element
        try:
            with revit.Transaction("Initialize AreaScheme Data"):
                success = data_manager.set_data(self._selected_node.Element, data_dict)
            
            if success:
                # Update JSON viewer to reflect changes
                self._update_json_viewer(self._selected_node)
                
                # Check if this node is actually in the tree (vs being a temporary node from _add_area_scheme)
                element_id = self._selected_node.Element.Id
                existing_node = self._find_node_by_element_id(element_id)
                
                if existing_node:
                    # Node already in tree - no need to do anything (already saved above)
                    pass
                else:
                    # Temporary node - need to rebuild tree to show it, then re-select
                    self.rebuild_tree()
                    
                    # Re-select using Dispatcher for proper timing
                    import System.Windows.Threading as Threading
                    
                    def do_reselect():
                        try:
                            node = self._find_node_by_element_id(element_id)
                            if node:
                                self._select_and_expand_node(node)
                        except:
                            pass
                    
                    self.tree_hierarchy.Dispatcher.BeginInvoke(
                        Threading.DispatcherPriority.ContextIdle,
                        System.Action(do_reselect)
                    )
        except Exception as e:
            print("Error saving default AreaScheme values: {}".format(e))
    
    def on_field_changed(self, sender, args):
        """Auto-save when a field changes"""
        if not self._selected_node:
            return
        
        # Collect data from all fields
        data_dict = {}
        for field_name, control in self._field_controls.items():
            if isinstance(control, TextBox):
                # Skip if showing default placeholder
                if control.Tag == "showing_default":
                    continue
                text = control.Text.strip()
                if text:
                    data_dict[field_name] = text
            elif isinstance(control, ComboBox):
                # Skip if showing default placeholder
                if control.Tag == "showing_default":
                    continue
                # For editable ComboBox, use Text property; for regular ComboBox, use SelectedItem
                if control.IsEditable:
                    text = control.Text.strip() if control.Text else ""
                    if text:
                        data_dict[field_name] = text
                else:
                    if control.SelectedItem:
                        data_dict[field_name] = control.SelectedItem
            elif isinstance(control, CheckBox):
                # FLOOR_UNDERGROUND uses "yes"/"no", IS_UNDERGROUND uses 1/0
                if field_name == "FLOOR_UNDERGROUND":
                    data_dict[field_name] = "yes" if control.IsChecked else "no"
                else:
                    data_dict[field_name] = 1 if control.IsChecked else 0
        
        # Save to element
        try:
            with revit.Transaction("Update pyArea Data"):
                success = data_manager.set_data(self._selected_node.Element, data_dict)
            
            if success:
                # Update JSON viewer to reflect changes
                self._update_json_viewer(self._selected_node)
        except Exception as e:
            print("Error saving data: {}".format(e))
    
    def _save_pending_changes(self):
        """Save any pending field changes before closing dialog"""
        if not self._selected_node or not self._field_controls:
            return
        
        # Trigger field changed handler to save current state
        # This ensures any fields that haven't lost focus yet are saved
        try:
            self.on_field_changed(None, None)
        except Exception as e:
            print("Error saving pending changes: {}".format(e))
    
    def on_add_clicked(self, sender, args):
        """Add new element to hierarchy - context-aware based on selection"""
        if not self._selected_node:
            # Nothing selected - add AreaScheme
            self._add_area_scheme()
        elif self._selected_node.ElementType == "AreaScheme":
            # AreaScheme selected - add Sheet
            self._add_sheet()
        elif self._selected_node.ElementType == "Sheet":
            # Sheet selected - add AreaPlan to sheet
            self._add_areaplan_to_sheet()
        elif self._selected_node.ElementType == "AreaPlan":
            # AreaPlan on sheet - add RepresentedAreaPlan
            self._add_represented_areaplan()
        elif self._selected_node.ElementType in ["AreaPlan_NotOnSheet", "RepresentedAreaPlan"]:
            # AreaPlan not on sheet or RepresentedAreaPlan - set representing view (move to parent)
            self._set_representing_view()
        else:
            # Fallback - show dialog
            options = ["AreaScheme", "Sheet", "AreaPlan", "RepresentedAreaPlan"]
            selected = forms.CommandSwitchWindow.show(
                options,
                message="What would you like to add?"
            )
            
            if not selected:
                return
            
            # Execute based on selection
            if selected == "AreaScheme":
                self._add_area_scheme()
            elif selected == "Sheet":
                self._add_sheet()
            elif selected == "AreaPlan":
                self._add_areaplan_to_sheet()
            elif selected == "RepresentedAreaPlan":
                self._add_represented_areaplan()
    
    def _add_area_scheme(self):
        """Add a new AreaScheme (define municipality for undefined schemes)"""
        # Get all existing area schemes
        collector = DB.FilteredElementCollector(self._doc)
        area_schemes = list(collector.OfClass(DB.AreaScheme).ToElements())
        
        if not area_schemes:
            forms.alert("No AreaSchemes found in the project. Please create one in Revit first.")
            return
        
        # Filter to only undefined AreaSchemes
        undefined_schemes = []
        for scheme in area_schemes:
            municipality = data_manager.get_municipality(scheme)
            if not municipality:
                undefined_schemes.append(scheme)
        
        if not undefined_schemes:
            forms.alert("All AreaSchemes already have municipality defined.")
            return
        
        # Let user pick an undefined AreaScheme
        scheme_dict = OrderedDict()
        for scheme in undefined_schemes:
            scheme_dict[scheme.Name] = scheme
        
        selected_name = forms.SelectFromList.show(
            sorted(scheme_dict.keys()),
            title="Select AreaScheme to Define",
            button_name="Select"
        )
        
        if selected_name:
            selected_scheme = scheme_dict[selected_name]
            # Select this scheme in tree and show properties
            self._selected_node = TreeNode(selected_scheme, "AreaScheme", selected_name)
            self.update_properties_panel()
    
    def _add_sheet(self):
        """Add a Sheet to selected AreaScheme"""
        if not self._selected_node or self._selected_node.ElementType != "AreaScheme":
            forms.alert("Please select an AreaScheme first.")
            return
        
        area_scheme = self._selected_node.Element
        area_scheme_id = str(area_scheme.Id.Value)
        
        # Get all sheets
        collector = DB.FilteredElementCollector(self._doc)
        all_sheets = list(collector.OfClass(DB.ViewSheet).ToElements())
        
        if not all_sheets:
            forms.alert("No sheets found in the project. Please create sheets in Revit first.")
            return
        
        # Categorize sheets
        sheets_with_areaplans = []  # Sheets with AreaPlans from this scheme
        sheets_already_assigned = []  # Sheets already assigned to this scheme
        other_sheets = []  # Other sheets
        
        for sheet in all_sheets:
            # Check if already assigned to this AreaScheme
            sheet_data = data_manager.get_data(sheet)
            if sheet_data and sheet_data.get("AreaSchemeId") == area_scheme_id:
                sheets_already_assigned.append(sheet)
                continue
            
            # Check if has AreaPlans from this scheme
            has_areaplans = False
            try:
                view_ids = sheet.GetAllPlacedViews()
                for view_id in view_ids:
                    view = self._doc.GetElement(view_id)
                    if hasattr(view, 'AreaScheme') and view.AreaScheme.Id == area_scheme.Id:
                        has_areaplans = True
                        break
            except:
                pass
            
            if has_areaplans:
                sheets_with_areaplans.append(sheet)
            else:
                other_sheets.append(sheet)
        
        # Build selection list with smart ordering using TemplateListItem
        class SheetOption(forms.TemplateListItem):
            def __init__(self, sheet, has_areaplans=False):
                # Store the sheet as the item
                super(SheetOption, self).__init__(sheet, checked=has_areaplans)
                self.has_areaplans = has_areaplans
            
            @property
            def name(self):
                """Display name for the list"""
                sheet = self.item
                sheet_name = "{} - {}".format(
                    sheet.SheetNumber if hasattr(sheet, 'SheetNumber') else "?",
                    sheet.Name if hasattr(sheet, 'Name') else "Unnamed"
                )
                if self.has_areaplans:
                    return "{} (has AreaPlans)".format(sheet_name)
                else:
                    return sheet_name
        
        # Build options list - sheets with AreaPlans first (and pre-checked)
        options = []
        for sheet in sheets_with_areaplans:
            options.append(SheetOption(sheet, has_areaplans=True))
        for sheet in other_sheets:
            options.append(SheetOption(sheet, has_areaplans=False))
        
        if not options:
            if sheets_already_assigned:
                forms.alert("All sheets are already assigned to this AreaScheme.")
            else:
                forms.alert("No sheets available to assign.")
            return
        
        # Show selection dialog with pre-checked sheets
        selected_options = forms.SelectFromList.show(
            options,
            title="Select Sheets for {}".format(area_scheme.Name),
            multiselect=True,
            button_name="Add Sheets"
        )
        
        if not selected_options:
            return
        
        # Map back to sheets (item property contains the actual sheet)
        selected_sheets = []
        for opt in selected_options:
            # Check if it's a SheetOption or the raw sheet
            if isinstance(opt, SheetOption):
                selected_sheets.append(opt.item)
            else:
                # Sometimes pyRevit returns the item directly
                selected_sheets.append(opt)
        
        # Assign sheets to AreaScheme
        with revit.Transaction("Assign Sheets to AreaScheme"):
            success_count = 0
            for sheet in selected_sheets:
                # Set AreaSchemeId on sheet
                sheet_data = data_manager.get_data(sheet) or {}
                sheet_data["AreaSchemeId"] = area_scheme_id
                if data_manager.set_data(sheet, sheet_data):
                    success_count += 1
        
        # Refresh tree and select first added sheet
        self.rebuild_tree()
        
        if selected_sheets:
            self._reselect_after_add(selected_sheets[0].Id)
    
    def _add_areaplan_to_sheet(self):
        """Add AreaPlan views to selected Sheet"""
        if not self._selected_node or self._selected_node.ElementType != "Sheet":
            forms.alert("Please select a Sheet first.")
            return
        
        sheet = self._selected_node.Element
        
        # Get the AreaScheme from the sheet's parent
        if not self._selected_node.Parent or self._selected_node.Parent.ElementType != "AreaScheme":
            forms.alert("Cannot determine AreaScheme for this sheet.")
            return
        
        area_scheme = self._selected_node.Parent.Element
        
        # Get all AreaPlan views with the same AreaScheme
        collector = DB.FilteredElementCollector(self._doc)
        all_views = collector.OfClass(DB.View).ToElements()
        
        # Get views already on this sheet
        views_on_this_sheet = set()
        try:
            view_ids = sheet.GetAllPlacedViews()
            for vid in view_ids:
                views_on_this_sheet.add(vid)
        except:
            pass
        
        # Filter to AreaPlan views with same scheme that are NOT already in the tree
        available_views = []
        views_already_on_sheet = []
        
        for view in all_views:
            try:
                if not hasattr(view, 'AreaScheme'):
                    continue
                
                view_area_scheme = view.AreaScheme
                if view_area_scheme is None or view_area_scheme.Id != area_scheme.Id:
                    continue
                
                # Skip views that already have data (already in tree)
                if data_manager.has_data(view):
                    continue
                
                # Check if already on this sheet (but no data yet)
                if view.Id in views_on_this_sheet:
                    views_already_on_sheet.append(view)
                else:
                    available_views.append(view)
            except:
                continue
        
        if not available_views and not views_already_on_sheet:
            forms.alert("No AreaPlan views found for this AreaScheme.\n\nCreate AreaPlan views in Revit first.")
            return
        
        # Build selection list
        class ViewOption(forms.TemplateListItem):
            def __init__(self, view, on_sheet=False):
                super(ViewOption, self).__init__(view, checked=on_sheet)
                self.on_sheet = on_sheet
            
            @property
            def name(self):
                view = self.item
                view_name = view.Name if hasattr(view, 'Name') else "Unnamed View"
                if self.on_sheet:
                    return "■ {} (already on sheet)".format(view_name)
                else:
                    return "□ {}".format(view_name)
        
        # Build options - views already on sheet first (pre-checked)
        options = []
        for view in views_already_on_sheet:
            options.append(ViewOption(view, on_sheet=True))
        for view in available_views:
            options.append(ViewOption(view, on_sheet=False))
        
        if not options:
            forms.alert("No views available.")
            return
        
        # Show selection dialog
        selected_options = forms.SelectFromList.show(
            options,
            title="Select AreaPlan Views for Sheet {}".format(
                sheet.SheetNumber if hasattr(sheet, 'SheetNumber') else "?"
            ),
            multiselect=True,
            button_name="Update Sheet"
        )
        
        if selected_options is None:
            return
        
        # Get selected views
        selected_views = []
        for opt in selected_options:
            if isinstance(opt, ViewOption):
                selected_views.append(opt.item)
            else:
                selected_views.append(opt)
        
        # Store the selected views that should be tracked for this sheet
        # Views already on the sheet are auto-detected
        # But we also track views that user wants to define even if not placed yet
        with revit.Transaction("Add AreaPlans to Tracking"):
            for view in selected_views:
                # Ensure view has data (even if empty) so it shows in tree
                view_data = data_manager.get_data(view) or {}
                # Mark it as belonging to this AreaScheme
                if not view_data:
                    # Initialize with empty data to mark it as "defined"
                    data_manager.set_data(view, {})
                
                # EDGE CASE: Check if this view was a represented view of any unplaced view
                # If so, we need to remove it from that unplaced view's RepresentedViews
                # because it's now placed on a sheet
                view_id_str = str(view.Id.Value)
                collector = DB.FilteredElementCollector(self._doc)
                all_views = collector.OfClass(DB.View).ToElements()
                
                for check_view in all_views:
                    check_data = data_manager.get_data(check_view)
                    if check_data and "RepresentedViews" in check_data:
                        rep_views = check_data.get("RepresentedViews", [])
                        if view_id_str in rep_views:
                            # Remove this view from the represented views list
                            rep_views.remove(view_id_str)
                            check_data["RepresentedViews"] = rep_views
                            data_manager.set_data(check_view, check_data)
        
        # Refresh tree to show updated state
        self.rebuild_tree()
        
        if selected_views:
            self._reselect_after_add(selected_views[0].Id)
    
    def _set_representing_view(self):
        """Set which view this AreaPlan represents (move to parent or pool)"""
        if not self._selected_node or self._selected_node.ElementType not in ["RepresentedAreaPlan", "AreaPlan_NotOnSheet"]:
            forms.alert("Please select an AreaPlan (not on sheet) or Represented AreaPlan first.")
            return
        
        represented_view = self._selected_node.Element
        current_parent = self._selected_node.Parent
        
        # For RepresentedAreaPlan, we need to get the current parent
        # For AreaPlan_NotOnSheet, current_parent might be AreaScheme (no parent view)
        has_current_parent = False
        if current_parent and current_parent.ElementType in ["AreaPlan", "AreaPlan_NotOnSheet"]:
            has_current_parent = True
        
        # Get the AreaScheme
        if not hasattr(represented_view, 'AreaScheme'):
            forms.alert("Selected view is not an AreaPlan.")
            return
        
        area_scheme = represented_view.AreaScheme
        
        # Get all AreaPlan views with the same AreaScheme (potential parents)
        collector = DB.FilteredElementCollector(self._doc)
        all_views = collector.OfClass(DB.View).ToElements()
        
        # Build set of views that are on sheets
        sheets_collector = DB.FilteredElementCollector(self._doc)
        all_sheets = list(sheets_collector.OfClass(DB.ViewSheet).ToElements())
        views_on_sheets = set()
        for sheet in all_sheets:
            try:
                view_ids = sheet.GetAllPlacedViews()
                for vid in view_ids:
                    views_on_sheets.add(vid)
            except:
                pass
        
        # Build set of ALL represented view IDs (views already represented by any parent)
        all_represented_ids = set()
        for check_view in all_views:
            check_data = data_manager.get_data(check_view)
            if check_data and "RepresentedViews" in check_data:
                rep_ids = check_data.get("RepresentedViews", [])
                # Convert string IDs to ElementIds for comparison
                for rep_id_str in rep_ids:
                    try:
                        rep_elem_id = DB.ElementId(Int64(int(rep_id_str)))
                        all_represented_ids.add(rep_elem_id)
                    except:
                        pass
        
        # Filter to valid parent candidates
        available_parents = []
        for view in all_views:
            try:
                if not hasattr(view, 'AreaScheme'):
                    continue
                
                view_area_scheme = view.AreaScheme
                if view_area_scheme is None or view_area_scheme.Id != area_scheme.Id:
                    continue
                
                # Skip the represented view itself
                if view.Id == represented_view.Id:
                    continue
                
                # Skip the current parent (if any)
                if has_current_parent and view.Id == current_parent.Element.Id:
                    continue
                
                # ONLY show views that are placed on sheets
                if view.Id not in views_on_sheets:
                    continue
                
                # Skip views that are already represented by another view
                # (unless it's the current view being moved)
                if view.Id in all_represented_ids and view.Id != represented_view.Id:
                    continue
                
                available_parents.append(view)
            except:
                continue
        
        if not available_parents:
            forms.alert("No available AreaPlan views found.\n\nEligible views must be:\n- Same AreaScheme\n- Placed on a sheet\n- Not already representing another view")
            return
        
        # Build selection list
        class ParentOption(forms.TemplateListItem):
            def __init__(self, view):
                super(ParentOption, self).__init__(view, checked=False)
            
            @property
            def name(self):
                view = self.item
                view_name = view.Name if hasattr(view, 'Name') else "Unnamed View"
                return "■ {}".format(view_name)
        
        # Add "Remove from all parents" option at the top
        options = ["──────────────────────────"]
        options.append("↺ Move to pool (remove from parent)")
        options.append("──────────────────────────")
        
        # Add parent options (all are on sheets now)
        for view in available_parents:
            option = ParentOption(view)
            options.append(option)
        
        # Show selection dialog
        selected = forms.SelectFromList.show(
            options,
            title="Move '{}' to...".format(represented_view.Name),
            button_name="Move"
        )
        
        if not selected:
            return
        
        # Handle selection
        try:
            with revit.Transaction("Set Representing View"):
                view_id_str = str(represented_view.Id.Value)
                
                # Remove from current parent (if any)
                if has_current_parent:
                    parent_data = data_manager.get_data(current_parent.Element) or {}
                    represented_ids = parent_data.get("RepresentedViews", [])
                    
                    if view_id_str in represented_ids:
                        represented_ids.remove(view_id_str)
                    
                    # Clean up empty RepresentedViews array
                    if represented_ids:
                        parent_data["RepresentedViews"] = represented_ids
                    else:
                        parent_data.pop("RepresentedViews", None)
                    
                    data_manager.set_data(current_parent.Element, parent_data)
                
                # Add to new parent or move to pool
                if selected == "↺ Move to pool (remove from parent)":
                    # Ensure the view has data so it shows as AreaPlan_NotOnSheet
                    view_data = data_manager.get_data(represented_view) or {}
                    if not view_data:
                        data_manager.set_data(represented_view, {})
                elif selected not in ["──────────────────────────"]:
                    # Get the new parent view
                    new_parent_view = selected.item if isinstance(selected, ParentOption) else selected
                    
                    # Add to new parent's RepresentedViews
                    new_parent_data = data_manager.get_data(new_parent_view) or {}
                    new_represented_ids = new_parent_data.get("RepresentedViews", [])
                    
                    if not isinstance(new_represented_ids, list):
                        new_represented_ids = []
                    
                    if view_id_str not in new_represented_ids:
                        new_represented_ids.append(view_id_str)
                    
                    new_parent_data["RepresentedViews"] = new_represented_ids
                    data_manager.set_data(new_parent_view, new_parent_data)
            
            # Refresh tree and re-select the moved view
            self.rebuild_tree()
            self._reselect_after_add(represented_view.Id)
        
        except Exception as e:
            print("Error setting representing view: {}".format(e))
    
    def _add_represented_areaplan(self):
        """Add RepresentedAreaPlan to selected AreaPlan"""
        if not self._selected_node or self._selected_node.ElementType not in ["AreaPlan", "AreaPlan_NotOnSheet"]:
            forms.alert("Please select an AreaPlan view first.")
            return
        
        current_view = self._selected_node.Element
        
        # Get the AreaScheme from the current view
        if not hasattr(current_view, 'AreaScheme'):
            forms.alert("Selected view is not an AreaPlan.")
            return
        
        area_scheme = current_view.AreaScheme
        
        # Get all AreaPlan views with the same AreaScheme
        collector = DB.FilteredElementCollector(self._doc)
        all_views = collector.OfClass(DB.View).ToElements()
        
        # Get all sheets once to check which views are placed
        sheets_collector = DB.FilteredElementCollector(self._doc)
        all_sheets = list(sheets_collector.OfClass(DB.ViewSheet).ToElements())
        
        # Build set of view IDs that are on sheets
        views_on_sheets = set()
        for sheet in all_sheets:
            try:
                view_ids = sheet.GetAllPlacedViews()
                for vid in view_ids:
                    views_on_sheets.add(vid)
            except:
                pass
        
        # Build set of ALL represented view IDs (from any view)
        all_represented_ids = set()
        for check_view in all_views:
            check_data = data_manager.get_data(check_view)
            if check_data and "RepresentedViews" in check_data:
                rep_ids = check_data.get("RepresentedViews", [])
                all_represented_ids.update(rep_ids)
        
        # Filter to AreaPlan views that are available to be represented
        available_views = []
        for view in all_views:
            try:
                if not hasattr(view, 'AreaScheme'):
                    continue
                
                # Check if AreaScheme is not None
                view_area_scheme = view.AreaScheme
                if view_area_scheme is None:
                    continue
                
                if view_area_scheme.Id != area_scheme.Id:
                    continue
                
                if view.Id == current_view.Id:
                    continue  # Skip the current view itself
                
                # Check if view is on any sheet
                if view.Id in views_on_sheets:
                    continue
                
                view_id_str = str(view.Id.Value)
                
                # Skip if already represented by ANY view
                if view_id_str in all_represented_ids:
                    continue
                
                # Views with data that are standalone (AreaPlan_NotOnSheet) are OK to add as represented
                # Only exclude if they don't meet the above criteria
                available_views.append(view)
            except:
                # Skip views that cause errors
                continue
        
        if not available_views:
            forms.alert("No available AreaPlan views found.\n\nRepresented AreaPlans must be:\n- Same AreaScheme as current view\n- Not placed on any sheet")
            return
        
        # Build selection list
        class ViewOption(forms.TemplateListItem):
            def __init__(self, view):
                super(ViewOption, self).__init__(view, checked=False)
            
            @property
            def name(self):
                view = self.item
                return view.Name if hasattr(view, 'Name') else "Unnamed View"
        
        options = [ViewOption(view) for view in available_views]
        
        # Show selection dialog
        selected_options = forms.SelectFromList.show(
            options,
            title="Select Represented AreaPlans for {}".format(current_view.Name),
            multiselect=True,
            button_name="Add Represented AreaPlans"
        )
        
        if not selected_options:
            return
        
        # Get selected views
        selected_views = []
        for opt in selected_options:
            if isinstance(opt, ViewOption):
                selected_views.append(opt.item)
            else:
                selected_views.append(opt)
        
        # Update RepresentedViews list
        try:
            view_data = data_manager.get_data(current_view) or {}
            represented_ids = view_data.get("RepresentedViews", [])
            
            # Ensure it's a list
            if not isinstance(represented_ids, list):
                represented_ids = []
            
            # Add new view IDs and handle nested represented views
            success = False
            with revit.Transaction("Add RepresentedViews"):
                for view in selected_views:
                    view_id_str = str(view.Id.Value)
                    if view_id_str not in represented_ids:
                        represented_ids.append(view_id_str)
                    
                    # EDGE CASE: Check if this view has its own represented views (nested)
                    # If so, flatten the hierarchy by adding them to the parent and removing from child
                    nested_view_data = data_manager.get_data(view)
                    if nested_view_data and "RepresentedViews" in nested_view_data:
                        nested_ids = nested_view_data.get("RepresentedViews", [])
                        if nested_ids:
                            # Add nested views to parent's list
                            for nested_id in nested_ids:
                                if nested_id not in represented_ids:
                                    represented_ids.append(nested_id)
                            
                            # Remove RepresentedViews from the child view (flatten hierarchy)
                            nested_view_data.pop("RepresentedViews", None)
                            data_manager.set_data(view, nested_view_data)
                
                # Save parent's updated RepresentedViews list
                view_data["RepresentedViews"] = represented_ids
                success = data_manager.set_data(current_view, view_data)
            
            # Refresh tree AFTER transaction and expand the node
            if success:
                # Save the path of the current node to ensure it stays expanded
                self._ensure_node_expanded_after_rebuild(self._selected_node)
                self.rebuild_tree()
                
                # Re-select the first added represented view
                if selected_views:
                    self._reselect_after_add(selected_views[0].Id)
            else:
                print("✗ WARNING: Failed to save RepresentedViews data")
        
        except Exception as e:
            print("Error adding Represented AreaPlans: {}".format(e))
    
    def on_remove_clicked(self, sender, args):
        """Remove data from selected element"""
        if not self._selected_node:
            forms.alert("Please select an element to remove data from.")
            return
        
        node = self._selected_node
        element_name = node.DisplayName
        element_type = node.ElementType
        
        # Confirm removal
        if element_type == "AreaScheme":
            message = "Remove municipality data from AreaScheme '{}'?\n\nThis will also remove all associated Sheet and AreaPlan data.".format(element_name)
        elif element_type == "Sheet":
            message = "Remove data from Sheet '{}'?\n\nThis will unlink it from the AreaScheme.".format(element_name)
        elif element_type == "AreaPlan":
            message = "Remove data from AreaPlan '{}'?".format(element_name)
        elif element_type == "RepresentedAreaPlan":
            message = "Remove '{}' from Represented AreaPlans list?".format(element_name)
        else:
            message = "Remove data from '{}'?".format(element_name)
        
        if not forms.alert(message, yes=True, no=True):
            return
        
        try:
            with revit.Transaction("Remove pyArea Data"):
                if element_type == "RepresentedAreaPlan":
                    # Remove from parent's RepresentedViews list only - don't delete the view's data
                    # This allows it to reappear as AreaPlan_NotOnSheet in the tree
                    if node.Parent and node.Parent.ElementType in ["AreaPlan", "AreaPlan_NotOnSheet"]:
                        parent_view = node.Parent.Element
                        view_data = data_manager.get_data(parent_view) or {}
                        represented_ids = view_data.get("RepresentedViews", [])
                        
                        # Remove this view's ID
                        view_id_str = str(node.Element.Id.Value)
                        if view_id_str in represented_ids:
                            represented_ids.remove(view_id_str)
                        
                        # Clean up: remove RepresentedViews field if empty
                        if represented_ids:
                            view_data["RepresentedViews"] = represented_ids
                        else:
                            view_data.pop("RepresentedViews", None)
                        
                        success = data_manager.set_data(parent_view, view_data)
                        
                        # Ensure the removed view has data so it shows as AreaPlan_NotOnSheet
                        if success:
                            removed_view_data = data_manager.get_data(node.Element) or {}
                            if not removed_view_data:
                                # Initialize with empty data to keep it in tree
                                data_manager.set_data(node.Element, {})
                    else:
                        success = False
                
                elif element_type == "AreaScheme":
                    # Remove data from AreaScheme and all associated Sheets and AreaPlans
                    area_scheme_id = str(node.Element.Id.Value)
                    removed_count = 0
                    
                    # Remove from all sheets
                    collector = DB.FilteredElementCollector(self._doc)
                    sheets = collector.OfClass(DB.ViewSheet).ToElements()
                    for sheet in sheets:
                        sheet_data = data_manager.get_data(sheet)
                        if sheet_data and sheet_data.get("AreaSchemeId") == area_scheme_id:
                            if data_manager.delete_data(sheet):
                                removed_count += 1
                    
                    # Remove from all AreaPlan views
                    views_collector = DB.FilteredElementCollector(self._doc)
                    views = views_collector.OfClass(DB.View).ToElements()
                    for view in views:
                        try:
                            if hasattr(view, 'AreaScheme') and view.AreaScheme and view.AreaScheme.Id == node.Element.Id:
                                if data_manager.delete_data(view):
                                    removed_count += 1
                        except:
                            pass
                    
                    # Remove from AreaScheme itself
                    success = data_manager.delete_data(node.Element)
                    if success:
                        removed_count += 1
                
                else:
                    # Remove data from element
                    success = data_manager.delete_data(node.Element)
            
            if success:
                self.rebuild_tree()
        
        except Exception as e:
            print("Error removing data: {}".format(e))
    
    def on_close_clicked(self, sender, args):
        """Close dialog"""
        # Save any pending field changes before closing
        self._save_pending_changes()
        
        # Save expansion state before closing
        self._save_expansion_state()
        self.Close()
    
    def _save_expansion_state(self):
        """Save which tree nodes are expanded"""
        try:
            expanded_paths = []
            
            # Collect paths of expanded nodes
            for i in range(self.tree_hierarchy.Items.Count):
                container = self.tree_hierarchy.ItemContainerGenerator.ContainerFromIndex(i)
                if container and container.IsExpanded:
                    node = self.tree_hierarchy.Items[i]
                    path = self._get_node_path(node)
                    expanded_paths.append(path)
                    # Recursively check children
                    self._collect_expanded_paths(container, path, expanded_paths)
            
            # Save to pyRevit config
            cfg = script.get_config()
            cfg.expanded_nodes = ','.join(expanded_paths) if expanded_paths else ''
            script.save_config()
        except:
            pass  # Silently fail if save doesn't work
    
    def _collect_expanded_paths(self, container, parent_path, expanded_paths):
        """Recursively collect expanded node paths"""
        try:
            if hasattr(container, 'Items'):
                for i in range(container.Items.Count):
                    child_container = container.ItemContainerGenerator.ContainerFromIndex(i)
                    if child_container and child_container.IsExpanded:
                        child_node = container.Items[i]
                        child_path = parent_path + '/' + child_node.DisplayName
                        expanded_paths.append(child_path)
                        self._collect_expanded_paths(child_container, child_path, expanded_paths)
        except:
            pass
    
    def _get_node_path(self, node):
        """Get unique path for a node (e.g., 'AreaScheme/Sheet/View')"""
        return node.DisplayName
    
    def _get_full_node_path(self, node):
        """Get full hierarchical path for a node (e.g., 'AreaScheme/Sheet/View')"""
        path_parts = []
        current = node
        while current:
            path_parts.insert(0, current.DisplayName)
            current = current.Parent
        return '/'.join(path_parts)
    
    def _ensure_node_expanded_after_rebuild(self, node):
        """Ensure a specific node path is expanded after rebuild"""
        try:
            # Get the full path of the node
            full_path = self._get_full_node_path(node)
            
            # Load current expansion state
            cfg = script.get_config()
            expanded_str = cfg.get_option('expanded_nodes', '')
            expanded_paths = set(expanded_str.split(',')) if expanded_str else set()
            
            # Add this path and all parent paths
            path_parts = full_path.split('/')
            for i in range(1, len(path_parts) + 1):
                partial_path = '/'.join(path_parts[:i])
                expanded_paths.add(partial_path)
            
            # Save back
            cfg.expanded_nodes = ','.join(expanded_paths)
            script.save_config()
        except:
            pass  # Silently fail if save doesn't work
    
    def _restore_expansion_state(self):
        """Restore saved expansion state"""
        try:
            # Load from pyRevit config
            cfg = script.get_config()
            expanded_str = cfg.get_option('expanded_nodes', '')
            
            if not expanded_str:
                # No saved state - expand all by default
                self._expand_all_nodes()
                return
            
            expanded_paths = set(expanded_str.split(','))
            
            # Use Dispatcher to delay expansion until UI is ready
            import System.Windows.Threading as Threading
            
            def do_restore():
                try:
                    for i in range(self.tree_hierarchy.Items.Count):
                        container = self.tree_hierarchy.ItemContainerGenerator.ContainerFromIndex(i)
                        if container:
                            node = self.tree_hierarchy.Items[i]
                            path = self._get_node_path(node)
                            # Expand if in saved state OR if it's an AreaScheme (always expand top level)
                            if path in expanded_paths or node.ElementType == "AreaScheme":
                                container.IsExpanded = True
                                container.UpdateLayout()
                                self._restore_children_expansion(container, path, expanded_paths, auto_expand_sheets=True)
                except:
                    pass
            
            self.tree_hierarchy.Dispatcher.BeginInvoke(
                Threading.DispatcherPriority.Background,
                System.Action(do_restore)
            )
        except:
            # If restore fails, expand all
            self._expand_all_nodes()
    
    def _restore_children_expansion(self, container, parent_path, expanded_paths, auto_expand_sheets=False):
        """Recursively restore expansion state for children"""
        try:
            if hasattr(container, 'Items'):
                for i in range(container.Items.Count):
                    child_container = container.ItemContainerGenerator.ContainerFromIndex(i)
                    if child_container:
                        child_node = container.Items[i]
                        child_path = parent_path + '/' + child_node.DisplayName
                        # Expand if in saved state OR if auto_expand_sheets is True and it's a Sheet
                        if child_path in expanded_paths or (auto_expand_sheets and child_node.ElementType == "Sheet"):
                            child_container.IsExpanded = True
                            child_container.UpdateLayout()
                            self._restore_children_expansion(child_container, child_path, expanded_paths, auto_expand_sheets)
        except:
            pass
    
    def _update_json_viewer(self, node):
        """Update JSON viewer with element's data"""
        try:
            import json
            # Get data from element
            data = data_manager.get_data(node.Element)
            
            # Set gray background for advanced data panel
            gray_brush = System.Windows.Media.BrushConverter().ConvertFromString("#F5F5F5")
            self.text_json.Background = gray_brush
            
            if data:
                # Pretty print JSON
                json_str = json.dumps(data, indent=2, ensure_ascii=False)
                self.text_json.Text = json_str
                self.text_json.Foreground = System.Windows.Media.Brushes.Black
            else:
                self.text_json.Text = "{}\n\n(No data stored)"
                self.text_json.Foreground = System.Windows.Media.Brushes.Gray
        except Exception as e:
            self.text_json.Text = "Error loading JSON: {}".format(e)
            self.text_json.Foreground = System.Windows.Media.Brushes.Red


if __name__ == '__main__':
    # Show dialog
    window = CalculationSetupWindow()
    window.ShowDialog()
