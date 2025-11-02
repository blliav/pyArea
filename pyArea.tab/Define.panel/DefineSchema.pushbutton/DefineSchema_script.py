# -*- coding: utf-8 -*-
"""Define pyArea Schema Data - Hierarchy Manager

Manages the complete hierarchy: AreaScheme > Sheet > AreaPlan > RepresentedAreaPlans
"""
__title__ = "Define Schema"
__doc__ = "Manage pyArea data hierarchy"

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
from System.Windows import Window
from System.Windows.Controls import TextBox, ComboBox, CheckBox, StackPanel, Grid, TextBlock
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


class DefineSchemaWindow(forms.WPFWindow):
    """Hierarchy Manager Dialog"""
    
    def __init__(self):
        forms.WPFWindow.__init__(self, 'DefineSchemaWindow.xaml')
        
        self._doc = revit.doc
        self._field_controls = {}
        self._selected_node = None
        self._tree_nodes = ObservableCollection[TreeNode]()
        
        # Wire up events
        self.tree_hierarchy.SelectedItemChanged += self.on_tree_selection_changed
        self.tree_hierarchy.MouseLeftButtonDown += self.on_tree_mouse_down
        self.btn_add.Click += self.on_add_clicked
        self.btn_remove.Click += self.on_remove_clicked
        self.btn_refresh.Click += self.on_refresh_clicked
        self.btn_close.Click += self.on_close_clicked
        
        # Build initial tree
        self.build_tree()
        
        # Set initial button text
        self._update_add_button_text()
        
        # Load saved expansion state
        self._restore_expansion_state()
        
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
            
            scheme_node.Status = "({})".format(municipality)
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
        area_scheme_id = str(area_scheme.Id.IntegerValue)
        
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
            
            for view_id in view_ids:
                view = self._doc.GetElement(view_id)
                
                # Check if it's an AreaPlan view with matching AreaScheme
                if hasattr(view, 'AreaScheme') and view.AreaScheme and view.AreaScheme.Id == area_scheme.Id:
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
                        if str(view.Id.IntegerValue) in rep_ids:
                            is_represented = True
                            break
                
                if is_represented:
                    continue
                
                # Add as standalone view at scheme level
                view_name = view.Name if hasattr(view, 'Name') else "Unnamed View"
                view_node = scheme_node.add_child(TreeNode(
                    view,
                    "AreaPlan_NotOnSheet",  # Hollow square - not on sheet
                    view_name
                ))
                
                # These can also have RepresentedViews
                self._add_represented_views(view_node)
            except:
                continue
    
    def _add_represented_views(self, view_node):
        """Add represented area plans for this AreaPlan"""
        view_data = data_manager.get_data(view_node.Element)
        if view_data and "RepresentedViews" in view_data:
            represented_ids = view_data.get("RepresentedViews", [])
            for rep_id in represented_ids:
                try:
                    rep_view = self._doc.GetElement(DB.ElementId(int(rep_id)))
                    if rep_view:
                        rep_name = rep_view.Name if hasattr(rep_view, 'Name') else "Unnamed"
                        view_node.add_child(TreeNode(
                            rep_view,
                            "RepresentedAreaPlan",
                            rep_name
                        ))
                except:
                    pass
    
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
        self.text_element_type.Text = "No selection"
        self.text_element_name.Text = ""
        self.group_municipality.Visibility = System.Windows.Visibility.Collapsed
        self.panel_fields.Children.Clear()
        self._field_controls = {}
        self.text_json.Text = "Select an element to view its JSON data..."
        self.text_json.Foreground = System.Windows.Media.Brushes.Gray
    
    def _update_add_button_text(self):
        """Update Add and Remove button text and enabled state based on selection"""
        if not self._selected_node:
            self.btn_add.Content = "➕ Add Scheme"
            self.btn_add.IsEnabled = True
            self.btn_remove.IsEnabled = False
        elif self._selected_node.ElementType == "AreaScheme":
            self.btn_add.Content = "➕ Add Sheet"
            self.btn_add.IsEnabled = True
            self.btn_remove.IsEnabled = True
        elif self._selected_node.ElementType == "Sheet":
            self.btn_add.Content = "➕ Add AreaPlan"
            self.btn_add.IsEnabled = True
            self.btn_remove.IsEnabled = True
        elif self._selected_node.ElementType == "AreaPlan":
            # AreaPlan on sheet - can add RepresentedViews but can't remove (it's on a sheet)
            self.btn_add.Content = "➕ Add Represented AreaPlan"
            self.btn_add.IsEnabled = True
            self.btn_remove.IsEnabled = False
        elif self._selected_node.ElementType == "AreaPlan_NotOnSheet":
            # AreaPlan not on sheet - can add RepresentedViews and can remove
            self.btn_add.Content = "➕ Add Represented AreaPlan"
            self.btn_add.IsEnabled = True
            self.btn_remove.IsEnabled = True
        elif self._selected_node.ElementType == "RepresentedAreaPlan":
            # RepresentedAreaPlans can't have nested RepresentedAreaPlans but can be removed
            self.btn_add.Content = "➕ Add Represented AreaPlan"
            self.btn_add.IsEnabled = False
            self.btn_remove.IsEnabled = True
        else:
            self.btn_add.Content = "➕ Add"
            self.btn_add.IsEnabled = True
            self.btn_remove.IsEnabled = True
    
    def update_properties_panel(self):
        """Update the right panel with selected element's properties"""
        if not self._selected_node:
            return
        
        node = self._selected_node
        
        # Update element info
        self.text_element_type.Text = node.ElementType
        self.text_element_name.Text = node.DisplayName
        
        # Update JSON viewer
        self._update_json_viewer(node)
        
        # Clear fields
        self.panel_fields.Children.Clear()
        self._field_controls = {}
        
        # Show municipality for Sheet/AreaPlan/RepresentedAreaPlan
        if node.ElementType in ["Sheet", "AreaPlan", "AreaPlan_NotOnSheet", "RepresentedAreaPlan"]:
            self.group_municipality.Visibility = System.Windows.Visibility.Visible
            municipality = self._get_municipality_for_node(node)
            if municipality:
                self.text_municipality.Text = municipality
                self.text_municipality.Foreground = System.Windows.Media.Brushes.DarkGreen
            else:
                self.text_municipality.Text = "Not detected"
                self.text_municipality.Foreground = System.Windows.Media.Brushes.Red
        else:
            self.group_municipality.Visibility = System.Windows.Visibility.Collapsed
        
        # Build fields based on element type
        self._build_fields_for_node(node)
    
    def _get_municipality_for_node(self, node):
        """Get municipality for a node"""
        if node.ElementType == "AreaScheme":
            return data_manager.get_municipality(node.Element)
        elif node.ElementType == "Sheet":
            return data_manager.get_municipality_from_sheet(self._doc, node.Element)
        elif node.ElementType in ["AreaPlan", "AreaPlan_NotOnSheet", "RepresentedAreaPlan"]:
            return data_manager.get_municipality_from_view(self._doc, node.Element)
        return None
    
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
        """Create a field control (same as before)"""
        grid = Grid()
        grid.ColumnDefinitions.Add(System.Windows.Controls.ColumnDefinition())
        grid.ColumnDefinitions.Add(System.Windows.Controls.ColumnDefinition())
        grid.ColumnDefinitions[0].Width = System.Windows.GridLength(150)
        
        # Label
        label = TextBlock()
        label.Text = "{}:".format(field_name)
        label.ToolTip = field_props.get("description", "")
        Grid.SetColumn(label, 0)
        grid.Children.Add(label)
        
        field_type = field_props.get("type")
        
        if field_name == "Municipality" or (field_type == "string" and "options" in field_props):
            # ComboBox for Municipality or options
            combo = ComboBox()
            if field_name == "Municipality":
                for muni in ["Common", "Jerusalem", "Tel-Aviv"]:
                    combo.Items.Add(muni)
                if current_value:
                    combo.SelectedItem = current_value
                else:
                    combo.SelectedIndex = 0
            else:
                for option in field_props["options"]:
                    combo.Items.Add(option)
                combo.SelectedIndex = 0
            Grid.SetColumn(combo, 1)
            grid.Children.Add(combo)
            self._field_controls[field_name] = combo
            # Auto-save on change
            combo.SelectionChanged += self.on_field_changed
            
        elif field_type == "int" and field_name in ["IS_UNDERGROUND"]:
            # CheckBox for boolean int
            checkbox = CheckBox()
            checkbox.Content = field_props.get("description", "")
            if current_value:
                checkbox.IsChecked = bool(current_value)
            Grid.SetColumn(checkbox, 1)
            grid.Children.Add(checkbox)
            self._field_controls[field_name] = checkbox
            # Auto-save on change
            checkbox.Checked += self.on_field_changed
            checkbox.Unchecked += self.on_field_changed
            
        else:
            # TextBox for everything else
            textbox = TextBox()
            textbox.ToolTip = field_props.get("description", "")
            # Set default or current value
            if current_value is not None:
                textbox.Text = str(current_value)
            else:
                default_value = field_props.get("default", "")
                if default_value:
                    textbox.Text = default_value
            Grid.SetColumn(textbox, 1)
            grid.Children.Add(textbox)
            self._field_controls[field_name] = textbox
            # Auto-save on lost focus
            textbox.LostFocus += self.on_field_changed
        
        self.panel_fields.Children.Add(grid)
    
    def on_field_changed(self, sender, args):
        """Auto-save when a field changes"""
        if not self._selected_node:
            return
        
        # Collect data from all fields
        data_dict = {}
        for field_name, control in self._field_controls.items():
            if isinstance(control, TextBox):
                text = control.Text.strip()
                if text:
                    data_dict[field_name] = text
            elif isinstance(control, ComboBox):
                if control.SelectedItem:
                    data_dict[field_name] = control.SelectedItem
            elif isinstance(control, CheckBox):
                data_dict[field_name] = 1 if control.IsChecked else 0
        
        # Save to element
        try:
            with revit.Transaction("Update pyArea Data"):
                success = data_manager.set_data(self._selected_node.Element, data_dict)
            
            if success:
                # Only rebuild tree if Municipality changed (new AreaScheme appears)
                # For other fields, just save without rebuilding to keep selection
                if "Municipality" in data_dict:
                    self.rebuild_tree()
        except Exception as e:
            print("Error saving data: {}".format(e))
    
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
        elif self._selected_node.ElementType in ["AreaPlan", "AreaPlan_NotOnSheet"]:
            # AreaPlan selected - add RepresentedAreaPlan
            self._add_represented_areaplan()
        elif self._selected_node.ElementType == "RepresentedAreaPlan":
            # RepresentedAreaPlan selected - add another RepresentedAreaPlan to parent AreaPlan
            # Find parent AreaPlan
            if self._selected_node.Parent and self._selected_node.Parent.ElementType in ["AreaPlan", "AreaPlan_NotOnSheet"]:
                # Temporarily select parent for adding
                original_selection = self._selected_node
                self._selected_node = self._selected_node.Parent
                self._add_represented_areaplan()
            else:
                forms.alert("Cannot determine parent AreaPlan.")
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
        area_scheme_id = str(area_scheme.Id.IntegerValue)
        
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
        
        # Refresh tree
        self.rebuild_tree()
    
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
        
        # Filter to AreaPlan views with same scheme
        available_views = []
        views_already_on_sheet = []
        
        for view in all_views:
            try:
                if not hasattr(view, 'AreaScheme'):
                    continue
                
                view_area_scheme = view.AreaScheme
                if view_area_scheme is None or view_area_scheme.Id != area_scheme.Id:
                    continue
                
                # Check if already on this sheet
                if view.Id in views_on_this_sheet:
                    views_already_on_sheet.append(view)
                else:
                    # Check if not used as RepresentedView anywhere
                    is_represented = False
                    # For now, allow all views - we'll just show them differently
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
        
        # Refresh tree to show updated state
        self.rebuild_tree()
    
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
        
        # Filter to AreaPlan views with same scheme, not on any sheet
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
                if view.Id not in views_on_sheets:
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
            
            # Add new view IDs
            for view in selected_views:
                view_id_str = str(view.Id.IntegerValue)
                if view_id_str not in represented_ids:
                    represented_ids.append(view_id_str)
            
            view_data["RepresentedViews"] = represented_ids
            
            # Save
            with revit.Transaction("Add RepresentedViews"):
                success = data_manager.set_data(current_view, view_data)
            
            # Refresh tree AFTER transaction
            if success:
                self.rebuild_tree()
        
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
                    # Remove from parent's RepresentedViews list
                    if node.Parent and node.Parent.ElementType == "AreaPlan":
                        parent_view = node.Parent.Element
                        view_data = data_manager.get_data(parent_view) or {}
                        represented_ids = view_data.get("RepresentedViews", [])
                        
                        # Remove this view's ID
                        view_id_str = str(node.Element.Id.IntegerValue)
                        if view_id_str in represented_ids:
                            represented_ids.remove(view_id_str)
                        
                        view_data["RepresentedViews"] = represented_ids
                        success = data_manager.set_data(parent_view, view_data)
                    else:
                        success = False
                
                elif element_type == "AreaScheme":
                    # Remove data from AreaScheme and all associated Sheets and AreaPlans
                    area_scheme_id = str(node.Element.Id.IntegerValue)
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
    
    def on_refresh_clicked(self, sender, args):
        """Refresh tree from Revit"""
        self.rebuild_tree()
    
    def on_close_clicked(self, sender, args):
        """Close dialog"""
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
    window = DefineSchemaWindow()
    window.ShowDialog()
