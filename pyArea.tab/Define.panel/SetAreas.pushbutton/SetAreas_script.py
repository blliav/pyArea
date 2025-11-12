# -*- coding: utf-8 -*-
"""Set Usage Type and Usage Type Prev for Area Elements

This dialog allows bulk editing of Usage Type and Usage Type Previous parameters,
as well as extensible schema fields (municipality-specific data) for Area elements.

FEATURES:
    - Colored dropdown selection for usage types
    - "Varies" detection across multiple selected areas
    - Extensible schema data entry with municipality-specific fields
    - Change detection with smart Apply button state management
    - Clearing fields (including required fields with defaults)

PERFORMANCE OPTIMIZATIONS APPLIED:
    - Direct Revit API imports instead of pyRevit wrappers (saves ~0.3s)
    - Minimal pyRevit imports (WPFWindow only, not entire forms module)
    - Single-pass parameter reading for multiple parameters
    - Cached number-to-text lookup dictionary
    
See DEVELOPMENT_LOG.md for detailed performance optimization history.
"""

import sys
import csv
import os

# Direct Revit API imports instead of pyRevit wrappers for better performance
import clr
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
from Autodesk.Revit import DB
from Autodesk.Revit.UI import TaskDialog

# Get document and UI document from pyRevit host
try:
    doc = __revit__.ActiveUIDocument.Document
    uidoc = __revit__.ActiveUIDocument
except:
    doc = None
    uidoc = None

# Minimal pyRevit imports - only what's absolutely needed
from pyrevit.forms import WPFWindow

from colored_combobox import ColoredComboBox

# Add lib folder to path
lib_path = os.path.join(os.path.dirname(__file__), "..", "..", "lib")
if lib_path not in sys.path:
    sys.path.insert(0, lib_path)

import data_manager

from schemas import municipality_schemas

# Import WPF for dynamic field controls (CLR already imported above)
clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
import System
from System.Windows.Controls import TextBox, ComboBox, CheckBox, TextBlock, Grid, RowDefinition, ColumnDefinition, StackPanel


class SetAreasWindow(WPFWindow):
    """Dialog window for setting area usage types"""
    
    def __init__(self, area_elements, options_list, municipality):
        WPFWindow.__init__(self, 'SetAreasWindow.xaml')
        
        self._areas = area_elements
        self._options = options_list
        self._municipality = municipality
        self.usage_type_value = None
        self.usage_type_prev_value = None
        self._schema_field_controls = {}
        self._initial_schema_values = {}  # Track initial values for change detection
        
        # Track initial state for "Varies" detection
        self._initial_usage_type_varies = False
        self._initial_usage_type_prev_varies = False
        
        # Track if any changes have been made
        self._has_changes = False
        
        # Setup comboboxes with colored options and color swatches
        self._combo_usage_type = ColoredComboBox(
            self.combo_usage_type, 
            options_list,
            self.color_swatch_usage_type,
            rtl=True
        )
        self._combo_usage_type_prev = ColoredComboBox(
            self.combo_usage_type_prev, 
            options_list,
            self.color_swatch_usage_type_prev,
            rtl=True
        )
        
        self._combo_usage_type.populate()
        self._combo_usage_type_prev.populate()
        
        # Set initial values - check if values vary across selected areas
        # OPTIMIZATION: Read both parameters in one pass through areas
        if area_elements:
            usage_type_value, usage_type_prev_value = self._get_initial_parameter_values(
                ["Usage Type", "Usage Type Prev"]
            )
            
            # Set Usage Type
            if usage_type_value == "<Varies>":
                self._initial_usage_type_varies = True
                self._combo_usage_type.set_initial_value("Varies")
                self._combo_usage_type._update_swatch(None)  # Clear color swatch
            elif usage_type_value:
                display_value = self._find_display_text_by_number(usage_type_value)
                if display_value:
                    self._combo_usage_type.set_initial_value(display_value)
                else:
                    self._combo_usage_type.set_initial_value("Not defined")
            else:
                self._combo_usage_type.set_initial_value("Not defined")
            
            # Set Usage Type Prev
            if usage_type_prev_value == "<Varies>":
                self._initial_usage_type_prev_varies = True
                self._combo_usage_type_prev.set_initial_value("Varies")
                self._combo_usage_type_prev._update_swatch(None)  # Clear color swatch
            elif usage_type_prev_value:
                display_value = self._find_display_text_by_number(usage_type_prev_value)
                if display_value:
                    self._combo_usage_type_prev.set_initial_value(display_value)
                else:
                    self._combo_usage_type_prev.set_initial_value("Not defined")
            else:
                self._combo_usage_type_prev.set_initial_value("Not defined")
        
        # Update dialog title with area count and municipality
        area_count = len(area_elements)
        self.Title = "Set {} Areas | {}".format(area_count, municipality)
        
        # Hide the info text at the bottom since we now show it in the title
        import System.Windows
        self.text_info.Visibility = System.Windows.Visibility.Collapsed
        
        # Build schema fields
        self._build_schema_fields()
        
        # Calculate and set dynamic window height based on field count
        self._set_dynamic_window_height()
        
        # Initialize Apply button state (disabled until changes are made)
        # Set directly to disabled before wiring events to avoid false triggers
        self.btn_apply.IsEnabled = False
        
        # Wire up change detection for Usage Type combo boxes
        self.combo_usage_type.SelectionChanged += self._on_field_changed
        self.combo_usage_type_prev.SelectionChanged += self._on_field_changed
    
    def _get_initial_parameter_values(self, param_names):
        """
        OPTIMIZED: Get initial parameter values from selected areas for multiple parameters in one pass.
        Returns a tuple of values in the same order as param_names.
        Each value is either the parameter value (if consistent), "<Varies>" (if different), or None (if empty/missing).
        
        Args:
            param_names: List of parameter names to read
            
        Returns:
            Tuple of values corresponding to each parameter name
        """
        if not self._areas:
            return tuple([None] * len(param_names))
        
        # Dictionary to store sets of values for each parameter
        param_values = {name: set() for name in param_names}
        
        # Single pass through all areas
        for area in self._areas:
            for param_name in param_names:
                param = area.LookupParameter(param_name)
                if param and param.HasValue:
                    value = param.AsString()
                    param_values[param_name].add(value if value else "")
                else:
                    param_values[param_name].add("")
        
        # Process results for each parameter
        results = []
        for param_name in param_names:
            values = param_values[param_name]
            if len(values) == 1:
                value = list(values)[0]
                results.append(value if value else None)
            else:
                # Values vary across areas
                results.append("<Varies>")
        
        return tuple(results)
    
    def _get_initial_parameter_value(self, param_name):
        """
        Get initial parameter value from selected areas.
        Returns the value if all areas have the same value, or "<Varies>" if they differ.
        Returns None if parameter doesn't exist or is empty.
        
        Note: For multiple parameters, use _get_initial_parameter_values() instead for better performance.
        """
        if not self._areas:
            return None
        
        values = set()
        for area in self._areas:
            param = area.LookupParameter(param_name)
            if param and param.HasValue:
                value = param.AsString()
                values.add(value if value else "")
            else:
                values.add("")
        
        # If all values are the same, return it
        if len(values) == 1:
            value = list(values)[0]
            return value if value else None
        else:
            # Values vary across areas
            return "<Varies>"
    
    def _find_display_text_by_number(self, number):
        """Find the full display text (number. name) from just the number"""
        # OPTIMIZATION: Cache the lookup dictionary for faster repeated lookups
        if not hasattr(self, '_number_to_text_cache'):
            self._number_to_text_cache = {}
            for option in self._options:
                # Options can be tuples (text, color) or just strings
                text = option[0] if isinstance(option, tuple) else option
                # Text format is "number. name"
                if ". " in text:
                    num_part = text.split(". ", 1)[0]
                    self._number_to_text_cache[num_part] = text
        
        return self._number_to_text_cache.get(number)
    
    def _build_schema_fields(self):
        """Build UI controls for extensible schema fields"""
        # Get field definitions for Area type
        self.panel_schema_fields.Children.Clear()
        
        # Get field definitions for Area type
        if not self._municipality or self._municipality == "Common":
            # Common has minimal fields
            if self._municipality == "Common":
                fields = municipality_schemas.AREA_FIELDS.get("Common", {})
            else:
                return
        else:
            fields = municipality_schemas.AREA_FIELDS.get(self._municipality, {})
        
        if not fields:
            return
        
        # Load existing data from all areas and detect variance
        field_values = {}
        if self._areas:
            # Collect values for each field from all areas
            for field_name in fields.keys():
                field_values[field_name] = set()
            
            for area in self._areas:
                area_data = data_manager.get_area_data(area) or {}
                for field_name in fields.keys():
                    value = area_data.get(field_name)
                    # Store None as empty string for consistency
                    field_values[field_name].add(str(value) if value is not None else "")
            
            # Process: if all same, use value; if different, use "<Varies>"
            existing_data = {}
            for field_name, values in field_values.items():
                if len(values) == 1:
                    # All areas have same value
                    val = list(values)[0]
                    existing_data[field_name] = val if val != "" else None
                else:
                    # Values vary across areas
                    existing_data[field_name] = "<Varies>"
        else:
            existing_data = {}
        
        # Store initial values for change detection
        self._initial_schema_values = dict(existing_data)
        
        # Create field controls
        for field_name, field_props in fields.items():
            self._create_field_control(field_name, field_props, existing_data.get(field_name))
    
    def _set_dynamic_window_height(self):
        """Calculate and set window height based on number of additional data fields"""
        # Base height for fixed elements (Usage Type section + buttons + margins)
        base_height = 240
        
        # Height per field (compact horizontal layout)
        field_height = 40
        
        # Border padding and extra space
        border_overhead = 60
        
        # Count number of fields
        field_count = self.panel_schema_fields.Children.Count
        
        if field_count == 0:
            # No additional fields - use minimal height
            calculated_height = base_height
        else:
            # Calculate height based on number of fields with extra padding
            calculated_height = base_height + border_overhead + (field_count * field_height)
        
        # Set minimum and maximum heights
        min_height = 300
        max_height = 700
        
        # Apply calculated height within bounds
        self.Height = max(min_height, min(calculated_height, max_height))
    
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
        field_type = field_props.get("type", "string")
        
        if field_type == "string" and "options" in field_props:
            # ComboBox for fields with predefined options
            combo = ComboBox()
            combo.FontSize = 11
            combo.Height = 26
            combo.Margin = System.Windows.Thickness(5, 0, 0, 0)
            combo.VerticalAlignment = System.Windows.VerticalAlignment.Center
            
            # Add "Varies" option if needed
            if current_value == "<Varies>":
                combo.Items.Add("<Varies>")
            
            for option in field_props["options"]:
                combo.Items.Add(option)
            
            if current_value:
                combo.SelectedItem = current_value
                # If showing Varies, style it
                if current_value == "<Varies>":
                    combo.FontStyle = System.Windows.FontStyles.Italic
                    combo.Foreground = System.Windows.Media.Brushes.Gray
            
            Grid.SetColumn(combo, 1)
            main_grid.Children.Add(combo)
            self._schema_field_controls[field_name] = combo
        else:
            # Check if field supports placeholders
            field_placeholders = field_props.get("placeholders", [])
            has_placeholders = len(field_placeholders) > 0
            
            if has_placeholders:
                # Use editable ComboBox with placeholders
                combo = ComboBox()
                combo.IsEditable = True
                combo.IsTextSearchEnabled = False  # Disable text search to prevent typing issues
                combo.FontSize = 11
                combo.Height = 26
                combo.Margin = System.Windows.Thickness(5, 0, 0, 0)
                combo.VerticalAlignment = System.Windows.VerticalAlignment.Center
                combo.ToolTip = field_props.get("description", "")
                
                # Add placeholder options
                for placeholder in field_placeholders:
                    combo.Items.Add(placeholder)
                
                # Set current value, varies indicator, or default
                if current_value == "<Varies>":
                    combo.Text = "<Varies>"
                    combo.FontStyle = System.Windows.FontStyles.Italic
                    combo.Foreground = System.Windows.Media.Brushes.Gray
                    combo.Tag = "showing_varies"
                elif current_value is not None:
                    combo.Text = str(current_value)
                elif default_value:
                    combo.Text = default_value
                    combo.Foreground = System.Windows.Media.Brushes.Gray
                    combo.Tag = "showing_default"
                
                # Create handlers with closure to capture default_value
                def create_combo_handlers(cb, def_val):
                    # Clear default or varies on focus
                    def on_got_focus(sender, args):
                        if sender.Tag == "showing_default" or sender.Tag == "showing_varies":
                            sender.Text = ""
                            sender.FontStyle = System.Windows.FontStyles.Normal
                            sender.Foreground = System.Windows.Media.Brushes.Black
                            sender.Tag = None
                    
                    # Reset to default if empty on lost focus (but not to Varies)
                    def on_lost_focus(sender, args):
                        if not sender.Text or sender.Text.strip() == "":
                            if def_val:
                                sender.Text = def_val
                                sender.Foreground = System.Windows.Media.Brushes.Gray
                                sender.Tag = "showing_default"
                    
                    return on_got_focus, on_lost_focus
                
                got_focus_handler, lost_focus_handler = create_combo_handlers(combo, default_value)
                combo.GotFocus += got_focus_handler
                combo.LostFocus += lost_focus_handler
                
                Grid.SetColumn(combo, 1)
                main_grid.Children.Add(combo)
                self._schema_field_controls[field_name] = combo
            else:
                # Regular TextBox for fields without placeholders
                textbox = TextBox()
                textbox.FontSize = 11
                textbox.Height = 26
                textbox.Margin = System.Windows.Thickness(5, 0, 0, 0)
                textbox.VerticalAlignment = System.Windows.VerticalAlignment.Center
                textbox.ToolTip = field_props.get("description", "")
                
                # Set value, varies indicator, or show default in gray
                if current_value == "<Varies>":
                    textbox.Text = "<Varies>"
                    textbox.FontStyle = System.Windows.FontStyles.Italic
                    textbox.Foreground = System.Windows.Media.Brushes.Gray
                    textbox.Tag = "showing_varies"
                elif current_value is not None:
                    textbox.Text = str(current_value)
                    textbox.Foreground = System.Windows.Media.Brushes.Black
                elif default_value:
                    textbox.Text = default_value
                    textbox.Foreground = System.Windows.Media.Brushes.Gray
                    textbox.Tag = "showing_default"
                
                # Create handlers with closure to capture default_value
                def create_textbox_handlers(tb, def_val):
                    # Clear default or varies on focus
                    def on_got_focus(sender, args):
                        if sender.Tag == "showing_default" or sender.Tag == "showing_varies":
                            sender.Text = ""
                            sender.FontStyle = System.Windows.FontStyles.Normal
                            sender.Foreground = System.Windows.Media.Brushes.Black
                            sender.Tag = None
                    
                    # Reset to default if empty on lost focus (but not to Varies)
                    def on_lost_focus(sender, args):
                        if not sender.Text or sender.Text.strip() == "":
                            if def_val:
                                sender.Text = def_val
                                sender.Foreground = System.Windows.Media.Brushes.Gray
                                sender.Tag = "showing_default"
                    
                    return on_got_focus, on_lost_focus
                
                got_focus_handler, lost_focus_handler = create_textbox_handlers(textbox, default_value)
                textbox.GotFocus += got_focus_handler
                textbox.LostFocus += lost_focus_handler
                
                Grid.SetColumn(textbox, 1)
                main_grid.Children.Add(textbox)
                self._schema_field_controls[field_name] = textbox
        
        self.panel_schema_fields.Children.Add(main_grid)
        
        # Wire up change detection for this field
        control = self._schema_field_controls.get(field_name)
        if control:
            if isinstance(control, TextBox):
                control.TextChanged += self._on_field_changed
            elif isinstance(control, ComboBox):
                control.SelectionChanged += self._on_field_changed
                # For editable ComboBox, also detect changes on LostFocus
                # (when user types without selecting from dropdown)
                if control.IsEditable:
                    control.LostFocus += self._on_field_changed
    
    def _on_field_changed(self, sender, args):
        """Event handler for any field change"""
        self._has_changes = True
        self._update_apply_button_state()
    
    def _update_apply_button_state(self):
        """Enable or disable Apply button based on whether changes have been made"""
        if not hasattr(self, 'btn_apply'):
            return
        
        has_changes = self._check_for_changes()
        self.btn_apply.IsEnabled = has_changes
    
    def _check_for_changes(self):
        """Check if any changes have been made to the form"""
        # Check Usage Type combo boxes
        usage_type_text = self._combo_usage_type.get_text()
        usage_type_prev_text = self._combo_usage_type_prev.get_text()
        
        # Check if Usage Type changed (including clearing to "Not defined")
        if usage_type_text and usage_type_text != "Varies":
            if not self._initial_usage_type_varies or usage_type_text != "":
                return True
        
        # Check if Usage Type Prev changed (including clearing to "Not defined")
        if usage_type_prev_text and usage_type_prev_text != "Varies":
            if not self._initial_usage_type_prev_varies or usage_type_prev_text != "":
                return True
        
        # Check schema fields for changes
        for field_name, control in self._schema_field_controls.items():
            current_value = self._get_schema_field_value(control)
            initial_value = self._initial_schema_values.get(field_name)
            
            # Normalize for comparison (None and "" are equivalent)
            current_normalized = current_value if current_value else None
            initial_normalized = initial_value if initial_value not in [None, "", "<Varies>"] else None
            
            # If values differ, there's a change
            if current_normalized != initial_normalized:
                return True
            
            # Special case: clearing a field that had a value (including <Varies>)
            if initial_value and initial_value != "<Varies>" and not current_value:
                return True
            
            # Special case: <Varies> was replaced with empty (user cleared it)
            if initial_value == "<Varies>" and current_value is None:
                # Check if the field was actually cleared (not just showing placeholder)
                # This includes fields showing default placeholder OR truly empty fields
                if isinstance(control, TextBox):
                    if control.Tag == "showing_default" or (control.Tag != "showing_varies" and not control.Text):
                        return True
                elif isinstance(control, ComboBox):
                    if control.Tag == "showing_default" or (control.Tag != "showing_varies" and not control.Text):
                        return True
        
        return False
    
    def apply_clicked(self, sender, args):
        """Handle Apply button click"""
        usage_type_text = self._combo_usage_type.get_text()
        usage_type_prev_text = self._combo_usage_type_prev.get_text()
        
        # Note: Apply button is only enabled when changes are detected
        # So we don't need validation here - button shouldn't be clickable without changes
        
        # Parse the usage type values (format: "number. name")
        # Track if user explicitly selected "Not defined" to clear parameters
        # Track if user left "Varies" unchanged to skip updates
        self.skip_usage_type = (usage_type_text == "Varies" and self._initial_usage_type_varies)
        self.skip_usage_type_prev = (usage_type_prev_text == "Varies" and self._initial_usage_type_prev_varies)
        self.clear_usage_type = (usage_type_text == "Not defined")
        self.clear_usage_type_prev = (usage_type_prev_text == "Not defined")
        
        if usage_type_text and usage_type_text != "Not defined" and usage_type_text != "Varies":
            parts = usage_type_text.split('. ', 1)  # Split on first ". " only
            if len(parts) == 2:
                self.usage_type_value = parts[0].strip()  # number
                self.usage_type_name = parts[1].strip()   # name
            else:
                self.usage_type_value = usage_type_text
                self.usage_type_name = None
        else:
            self.usage_type_value = None
            self.usage_type_name = None
        
        if usage_type_prev_text and usage_type_prev_text != "Not defined" and usage_type_prev_text != "Varies":
            parts = usage_type_prev_text.split('. ', 1)  # Split on first ". " only
            if len(parts) == 2:
                self.usage_type_prev_value = parts[0].strip()  # number
                self.usage_type_prev_name = parts[1].strip()   # name
            else:
                self.usage_type_prev_value = usage_type_prev_text
                self.usage_type_prev_name = None
        else:
            self.usage_type_prev_value = None
            self.usage_type_prev_name = None
        
        # Collect schema data - include cleared fields
        self.schema_data = {}
        self.schema_fields_to_clear = []
        
        for field_name, control in self._schema_field_controls.items():
            value = self._get_schema_field_value(control)
            initial_value = self._initial_schema_values.get(field_name)
            
            if value:
                # Field has a value - save it
                self.schema_data[field_name] = value
            elif initial_value and initial_value != "<Varies>":
                # Field was cleared (had value before, now empty)
                self.schema_fields_to_clear.append(field_name)
            elif initial_value == "<Varies>":
                # Was varies, check if user cleared it
                # This includes: empty fields OR fields showing default placeholder
                if isinstance(control, TextBox):
                    if control.Tag == "showing_default" or (control.Tag != "showing_varies" and not control.Text):
                        # User cleared the varies field (now showing default or empty)
                        self.schema_fields_to_clear.append(field_name)
                elif isinstance(control, ComboBox):
                    if control.Tag == "showing_default" or (control.Tag != "showing_varies" and not control.Text):
                        # User cleared the varies field (now showing default or empty)
                        self.schema_fields_to_clear.append(field_name)
        
        self.DialogResult = True
        self.Close()
    
    def _get_schema_field_value(self, control):
        """Get value from a schema field control"""
        if isinstance(control, TextBox):
            # Skip if showing default placeholder or varies indicator
            if control.Tag == "showing_default" or control.Tag == "showing_varies":
                return None
            text = control.Text.strip() if control.Text else ""
            # Don't save "<Varies>" as a value
            if text == "<Varies>":
                return None
            return text if text else None
        elif isinstance(control, ComboBox):
            # Skip if showing default placeholder or varies indicator
            if control.Tag == "showing_default" or control.Tag == "showing_varies":
                return None
            # For editable ComboBox, use Text property; for regular ComboBox, use SelectedItem
            if control.IsEditable:
                text = control.Text.strip() if control.Text else ""
                # Don't save "<Varies>" as a value
                if text == "<Varies>":
                    return None
                return text if text else None
            else:
                if control.SelectedItem:
                    selected = str(control.SelectedItem)
                    # Don't save "<Varies>" as a value
                    if selected == "<Varies>":
                        return None
                    return selected
        return None
    
    def cancel_clicked(self, sender, args):
        """Handle Cancel button click"""
        self.DialogResult = False
        self.Close()


def load_usage_types_from_csv(doc, active_view):
    """
    Load usage type options from CSV file based on municipality of the view's area scheme.
    
    Args:
        doc: Revit document
        active_view: Current active view
    
    Returns:
        tuple: (municipality, options_list) where options_list is list of tuples: (text, (R, G, B)) or just text strings
    """
    municipality = None
    variant = "Default"
    
    # Check if active view is an AreaPlan view
    if hasattr(active_view, 'AreaScheme'):
        municipality, variant = data_manager.get_municipality_from_view(doc, active_view)
    
    # Default to Common if no municipality detected
    if not municipality:
        municipality = "Common"
        variant = "Default"
    
    # Build path to CSV file based on municipality and variant
    script_dir = os.path.dirname(__file__)
    tab_path = os.path.dirname(os.path.dirname(script_dir))  # Go up to pyArea.tab
    csv_filename = municipality_schemas.get_usage_type_csv_filename(municipality, variant)
    csv_path = os.path.join(tab_path, "lib", csv_filename)
    
    # Fallback to lib folder if CSV not found directly
    if not os.path.exists(csv_path):
        csv_path = os.path.join(tab_path, "lib", "schemas", csv_filename)
    
    options = []
    
    # Add "Not defined" option at the top (without color)
    options.append("Not defined")
    
    try:
        import codecs
        with codecs.open(csv_path, 'r', encoding='utf-8-sig') as f:
            reader = csv.DictReader(f)
            for row in reader:
                if row.get('usage_type') and row.get('name'):
                    try:
                        number = row['usage_type'].strip()
                        name = row['name'].strip()
                        
                        # Format: "number. name"
                        display_text = u"{}. {}".format(number, name)
                        
                        # Get RGB values if available
                        r = int(row['R'].strip()) if row.get('R') and row['R'].strip() else None
                        g = int(row['G'].strip()) if row.get('G') and row['G'].strip() else None
                        b = int(row['B'].strip()) if row.get('B') and row['B'].strip() else None
                        
                        if r is not None and g is not None and b is not None:
                            # Add as tuple with color
                            options.append((display_text, (r, g, b)))
                        else:
                            # Add as string without color
                            options.append(display_text)
                    
                    except (ValueError, KeyError):
                        continue
    
    except Exception as e:
        TaskDialog.Show("CSV Error", "Error loading CSV file: {}\nPath: {}\nMunicipality: {}".format(str(e), csv_path, municipality))
        sys.exit()
    
    return municipality, options


def main():
    """Main function - entry point for the script"""
    # Document already imported at module level
    active_view = doc.ActiveView
    
    # Get selected area elements
    selection_ids = uidoc.Selection.GetElementIds()
    area_elements = [doc.GetElement(el_id) for el_id in selection_ids 
                     if isinstance(doc.GetElement(el_id), DB.Area)]
    
    if not area_elements:
        TaskDialog.Show("Selection Error", "Please select at least one area element.")
        sys.exit()
    
    # Load usage type options from CSV based on municipality
    municipality, options_list = load_usage_types_from_csv(doc, active_view)
    
    if not options_list:
        TaskDialog.Show("CSV Error", "No usage types found in CSV file for municipality: {}".format(municipality))
        sys.exit()
    
    # Show dialog
    dialog = SetAreasWindow(area_elements, options_list, municipality)
    dialog.ShowDialog()
    
    # If user clicked Apply
    if dialog.DialogResult:
        # Start transaction to modify parameters and schema data
        t = DB.Transaction(doc, "Set Area Usage Types and Data")
        t.Start()
        try:
            updated_usage_type = 0
            updated_usage_type_prev = 0
            failed_usage_type = 0
            failed_usage_type_prev = 0
            updated_schema_data = 0
            failed_schema_data = 0
            
            for area in area_elements:
                # Update or clear Usage Type
                if dialog.skip_usage_type:
                    # Skip - user left "Varies" unchanged
                    pass
                elif dialog.clear_usage_type:
                    # Clear Usage Type parameters
                    param = area.LookupParameter("Usage Type")
                    if param and not param.IsReadOnly:
                        param.Set("")
                        updated_usage_type += 1
                    
                    name_param = area.LookupParameter("Name")
                    if name_param and not name_param.IsReadOnly:
                        name_param.Set("")
                    
                    usage_type_name_param = area.LookupParameter("Usage Type Name")
                    if usage_type_name_param and not usage_type_name_param.IsReadOnly:
                        usage_type_name_param.Set("")
                
                elif dialog.usage_type_value:
                    # Set Usage Type parameter (number only)
                    param = area.LookupParameter("Usage Type")
                    if param and not param.IsReadOnly:
                        param.Set(dialog.usage_type_value)
                        updated_usage_type += 1
                    else:
                        failed_usage_type += 1
                    
                    # Set Name parameter (name text)
                    if dialog.usage_type_name:
                        name_param = area.LookupParameter("Name")
                        if name_param and not name_param.IsReadOnly:
                            name_param.Set(dialog.usage_type_name)
                        
                        # Set Usage Type Name parameter if it exists
                        usage_type_name_param = area.LookupParameter("Usage Type Name")
                        if usage_type_name_param and not usage_type_name_param.IsReadOnly:
                            usage_type_name_param.Set(dialog.usage_type_name)
                
                # Update or clear Usage Type Prev
                if dialog.skip_usage_type_prev:
                    # Skip - user left "Varies" unchanged
                    pass
                elif dialog.clear_usage_type_prev:
                    # Clear Usage Type Prev parameters
                    param = area.LookupParameter("Usage Type Prev")
                    if param and not param.IsReadOnly:
                        param.Set("")
                        updated_usage_type_prev += 1
                    
                    usage_type_prev_name_param = area.LookupParameter("Usage Type Prev. Name")
                    if usage_type_prev_name_param and not usage_type_prev_name_param.IsReadOnly:
                        usage_type_prev_name_param.Set("")
                
                elif dialog.usage_type_prev_value:
                    # Set Usage Type Prev parameter (number only)
                    param = area.LookupParameter("Usage Type Prev")
                    if param and not param.IsReadOnly:
                        param.Set(dialog.usage_type_prev_value)
                        updated_usage_type_prev += 1
                    else:
                        failed_usage_type_prev += 1
                    
                    # Set Usage Type Prev. Name parameter if it exists
                    if dialog.usage_type_prev_name:
                        usage_type_prev_name_param = area.LookupParameter("Usage Type Prev. Name")
                        if usage_type_prev_name_param and not usage_type_prev_name_param.IsReadOnly:
                            usage_type_prev_name_param.Set(dialog.usage_type_prev_name)
                
                # Update extensible schema data if provided or fields were cleared
                if dialog.schema_data or dialog.schema_fields_to_clear:
                    # Get existing data
                    existing_data = data_manager.get_area_data(area) or {}
                    
                    # Merge with new/updated data
                    if dialog.schema_data:
                        existing_data.update(dialog.schema_data)
                    
                    # Remove cleared fields
                    if dialog.schema_fields_to_clear:
                        for field_name in dialog.schema_fields_to_clear:
                            if field_name in existing_data:
                                del existing_data[field_name]
                    
                    # Save to extensible schema
                    success, errors = data_manager.set_area_data(area, existing_data, municipality)
                    if success:
                        updated_schema_data += 1
                    else:
                        failed_schema_data += 1
            
            t.Commit()
        except Exception as ex:
            t.RollBack()
            TaskDialog.Show("Transaction Error", "Failed to save changes: {}".format(str(ex)))


if __name__ == "__main__":
    main()
