# -*- coding: utf-8 -*-
"""Set Usage Type and Usage Type Prev for Area Elements"""
__title__ = "Set Areas"
__doc__ = "Bulk set Usage Type and Usage Type Prev parameters for selected areas"

import sys
import csv
import os
from pyrevit import revit, DB, forms, script
from colored_combobox import ColoredComboBox

# Add lib folder to path
lib_path = os.path.join(os.path.dirname(__file__), "..", "..", "lib")
if lib_path not in sys.path:
    sys.path.insert(0, lib_path)

import data_manager
from schemas import municipality_schemas

# Import WPF for dynamic field controls
import clr
clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
import System
from System.Windows.Controls import TextBox, ComboBox, CheckBox, TextBlock, Grid, RowDefinition, ColumnDefinition, StackPanel


class SetAreasWindow(forms.WPFWindow):
    """Dialog window for setting area usage types"""
    
    def __init__(self, area_elements, options_list, municipality):
        forms.WPFWindow.__init__(self, 'SetAreasWindow.xaml')
        self._areas = area_elements
        self._options = options_list
        self._municipality = municipality
        self.usage_type_value = None
        self.usage_type_prev_value = None
        self._schema_field_controls = {}
        
        # Track initial state for "Varies" detection
        self._initial_usage_type_varies = False
        self._initial_usage_type_prev_varies = False
        
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
        if area_elements:
            # Check Usage Type values
            usage_type_value = self._get_initial_parameter_value("Usage Type")
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
            
            # Check Usage Type Prev values
            usage_type_prev_value = self._get_initial_parameter_value("Usage Type Prev")
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
    
    def _get_initial_parameter_value(self, param_name):
        """
        Get initial parameter value from selected areas.
        Returns the value if all areas have the same value, or "<Varies>" if they differ.
        Returns None if parameter doesn't exist or is empty.
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
        for option in self._options:
            # Options can be tuples (text, color) or just strings
            text = option[0] if isinstance(option, tuple) else option
            # Text format is "number. name"
            if text.startswith(number + ". "):
                return text
        return None
    
    def _build_schema_fields(self):
        """Build input fields for extensible schema data based on municipality"""
        # Clear existing fields
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
        
        # Load existing data from first area (if available)
        existing_data = {}
        if self._areas:
            existing_data = data_manager.get_area_data(self._areas[0]) or {}
        
        # Create field controls
        for field_name, field_props in fields.items():
            self._create_field_control(field_name, field_props, existing_data.get(field_name))
    
    def _create_field_control(self, field_name, field_props, current_value):
        """Create a field control following CalculationSettings UI pattern"""
        # Main container grid
        main_grid = Grid()
        main_grid.Margin = System.Windows.Thickness(0, 8, 0, 2)
        
        # Define rows: Label row, Input row
        main_grid.RowDefinitions.Add(RowDefinition())
        main_grid.RowDefinitions.Add(RowDefinition())
        main_grid.RowDefinitions[0].Height = System.Windows.GridLength(1, System.Windows.GridUnitType.Auto)
        main_grid.RowDefinitions[1].Height = System.Windows.GridLength(1, System.Windows.GridUnitType.Auto)
        
        # Label row with English name on left, Hebrew on right
        label_grid = Grid()
        label_grid.ColumnDefinitions.Add(ColumnDefinition())
        label_grid.ColumnDefinitions.Add(ColumnDefinition())
        label_grid.ColumnDefinitions[0].Width = System.Windows.GridLength(1, System.Windows.GridUnitType.Star)
        label_grid.ColumnDefinitions[1].Width = System.Windows.GridLength(1, System.Windows.GridUnitType.Auto)
        Grid.SetRow(label_grid, 0)
        
        # Left side: English label with required indicator
        left_panel = StackPanel()
        left_panel.Orientation = System.Windows.Controls.Orientation.Horizontal
        Grid.SetColumn(left_panel, 0)
        
        # English label
        label_en = TextBlock()
        label_en.Text = field_name
        label_en.FontSize = 11
        label_en.FontWeight = System.Windows.FontWeights.SemiBold
        label_en.Foreground = System.Windows.Media.Brushes.Black
        label_en.ToolTip = field_props.get("description", "")
        label_en.Margin = System.Windows.Thickness(0, 0, 5, 3)
        left_panel.Children.Add(label_en)
        
        # Required indicator
        if field_props.get("required", False):
            required_label = TextBlock()
            required_label.Text = "*"
            required_label.FontSize = 11
            required_label.FontWeight = System.Windows.FontWeights.Bold
            required_label.Foreground = System.Windows.Media.Brushes.Red
            required_label.Margin = System.Windows.Thickness(0, 0, 0, 3)
            left_panel.Children.Add(required_label)
        
        label_grid.Children.Add(left_panel)
        
        # Right side: Hebrew label (if available)
        hebrew_name = field_props.get("hebrew_name", "")
        if hebrew_name:
            label_he = TextBlock()
            label_he.Text = hebrew_name
            label_he.FontSize = 11
            label_he.FontWeight = System.Windows.FontWeights.SemiBold
            label_he.Foreground = System.Windows.Media.Brushes.Black
            label_he.HorizontalAlignment = System.Windows.HorizontalAlignment.Right
            label_he.Margin = System.Windows.Thickness(0, 0, 0, 3)
            Grid.SetColumn(label_he, 1)
            label_grid.Children.Add(label_he)
        
        main_grid.Children.Add(label_grid)
        
        # Get default value
        default_value = field_props.get("default", "")
        
        # Input row - create appropriate control based on field type
        field_type = field_props.get("type", "string")
        
        if field_type == "string" and "options" in field_props:
            # ComboBox for fields with predefined options
            combo = ComboBox()
            combo.FontSize = 11
            combo.Height = 26
            combo.Margin = System.Windows.Thickness(0, 2, 0, 0)
            for option in field_props["options"]:
                combo.Items.Add(option)
            if current_value:
                combo.SelectedItem = current_value
            Grid.SetRow(combo, 1)
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
                combo.Margin = System.Windows.Thickness(0, 2, 0, 0)
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
                    
                    return on_got_focus, on_lost_focus
                
                got_focus_handler, lost_focus_handler = create_combo_handlers(combo, default_value)
                combo.GotFocus += got_focus_handler
                combo.LostFocus += lost_focus_handler
                
                Grid.SetRow(combo, 1)
                main_grid.Children.Add(combo)
                self._schema_field_controls[field_name] = combo
            else:
                # Regular TextBox for fields without placeholders
                textbox = TextBox()
                textbox.FontSize = 11
                textbox.Height = 26
                textbox.Margin = System.Windows.Thickness(0, 2, 0, 0)
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
                    
                    return on_got_focus, on_lost_focus
                
                got_focus_handler, lost_focus_handler = create_textbox_handlers(textbox, default_value)
                textbox.GotFocus += got_focus_handler
                textbox.LostFocus += lost_focus_handler
                
                Grid.SetRow(textbox, 1)
                main_grid.Children.Add(textbox)
                self._schema_field_controls[field_name] = textbox
        
        self.panel_schema_fields.Children.Add(main_grid)
    
    def apply_clicked(self, sender, args):
        """Handle Apply button click"""
        usage_type_text = self._combo_usage_type.get_text()
        usage_type_prev_text = self._combo_usage_type_prev.get_text()
        
        # At least one field must be changed or filled
        # "Varies" without change doesn't count as input
        has_usage_type_change = (usage_type_text and 
                                 usage_type_text != "Varies" and 
                                 usage_type_text != "Not defined")
        has_usage_type_prev_change = (usage_type_prev_text and 
                                      usage_type_prev_text != "Varies" and 
                                      usage_type_prev_text != "Not defined")
        # User can also explicitly choose "Not defined" to clear
        has_clear_action = (usage_type_text == "Not defined" or 
                           usage_type_prev_text == "Not defined")
        has_schema_data = any(self._get_schema_field_value(ctrl) 
                             for ctrl in self._schema_field_controls.values())
        
        if not has_usage_type_change and not has_usage_type_prev_change and not has_clear_action and not has_schema_data:
            forms.alert("Please make at least one change.", exitscript=False)
            return
        
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
        
        # Collect schema data
        self.schema_data = {}
        for field_name, control in self._schema_field_controls.items():
            value = self._get_schema_field_value(control)
            if value:
                self.schema_data[field_name] = value
        
        self.DialogResult = True
        self.Close()
    
    def _get_schema_field_value(self, control):
        """Get value from a schema field control"""
        if isinstance(control, TextBox):
            # Skip if showing default placeholder
            if control.Tag == "showing_default":
                return None
            text = control.Text.strip() if control.Text else ""
            return text if text else None
        elif isinstance(control, ComboBox):
            # Skip if showing default placeholder
            if control.Tag == "showing_default":
                return None
            # For editable ComboBox, use Text property; for regular ComboBox, use SelectedItem
            if control.IsEditable:
                text = control.Text.strip() if control.Text else ""
                return text if text else None
            else:
                if control.SelectedItem:
                    return str(control.SelectedItem)
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
    # Detect municipality from the active view's area scheme
    municipality = None
    
    # Check if active view is an AreaPlan view
    if hasattr(active_view, 'AreaScheme'):
        municipality = data_manager.get_municipality_from_view(doc, active_view)
    
    # Default to Common if no municipality detected
    if not municipality:
        municipality = "Common"
    
    # Build path to CSV file based on municipality
    script_dir = os.path.dirname(__file__)
    tab_path = os.path.dirname(os.path.dirname(script_dir))  # Go up to pyArea.tab
    csv_filename = "UsageType_{}.csv".format(municipality)
    csv_path = os.path.join(tab_path, "lib", "schemas", csv_filename)
    
    # Fallback to lib folder if not in schemas folder
    if not os.path.exists(csv_path):
        csv_path = os.path.join(tab_path, "lib", csv_filename)
    
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
        forms.alert("Error loading CSV file: {}\nPath: {}\nMunicipality: {}".format(str(e), csv_path, municipality), exitscript=True)
    
    return municipality, options


def main():
    """Main function - entry point for the script"""
    doc = revit.doc
    active_view = doc.ActiveView
    logger = script.get_logger()
    
    # Get selected area elements
    area_elements = [el for el in revit.get_selection() 
                     if isinstance(el, DB.Area) or isinstance(el, DB.Architecture.Room)]
    
    if not area_elements:
        forms.alert("Please select at least one area element.", exitscript=True)
    
    # Load usage type options from CSV based on municipality
    municipality, options_list = load_usage_types_from_csv(doc, active_view)
    
    logger.info("Detected municipality: {}".format(municipality))
    logger.info("Loaded {} usage type options".format(len(options_list)))
    
    if not options_list:
        forms.alert("No usage types found in CSV file for municipality: {}".format(municipality), exitscript=True)
    
    # Show dialog
    dialog = SetAreasWindow(area_elements, options_list, municipality)
    dialog.ShowDialog()
    
    # If user clicked Apply
    if dialog.DialogResult:
        # Start transaction to modify parameters and schema data
        with revit.Transaction("Set Area Usage Types and Data"):
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
                
                # Update extensible schema data if provided
                if dialog.schema_data:
                    # Get existing data and merge with new data
                    existing_data = data_manager.get_area_data(area) or {}
                    existing_data.update(dialog.schema_data)
                    
                    # Save to extensible schema
                    success, errors = data_manager.set_area_data(area, existing_data, municipality)
                    if success:
                        updated_schema_data += 1
                    else:
                        failed_schema_data += 1
                        logger.warning("Failed to save schema data for area {}: {}".format(
                            area.Id, ", ".join(errors)))
            
            # Log results
            logger = script.get_logger()
            if dialog.usage_type_value:
                logger.info("Usage Type: Updated {} area(s), Failed: {}".format(
                    updated_usage_type, failed_usage_type))
            
            if dialog.usage_type_prev_value:
                logger.info("Usage Type Prev: Updated {} area(s), Failed: {}".format(
                    updated_usage_type_prev, failed_usage_type_prev))
            
            if dialog.schema_data:
                logger.info("Schema Data: Updated {} area(s), Failed: {}".format(
                    updated_schema_data, failed_schema_data))


if __name__ == "__main__":
    main()
