# -*- coding: utf-8 -*-
"""Define pyArea Schema Data on Elements

Allows users to define extensible storage data on AreaSchemes, Sheets, and AreaPlans.
"""
__title__ = "Define Schema"
__doc__ = "Define pyArea extensible storage data on selected elements"

import sys
import os
from pyrevit import revit, DB, forms, script

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


class DefineSchemaWindow(forms.WPFWindow):
    """Dialog window for defining schema data"""
    
    def __init__(self, elements, element_type):
        forms.WPFWindow.__init__(self, 'DefineSchemaWindow.xaml')
        
        self._elements = elements
        self._element_type = element_type
        self._field_controls = {}
        self._current_municipality = None
        
        # Update element information
        self.text_element_type.Text = element_type
        self.text_selected_count.Text = "{} element(s)".format(len(elements))
        
        # Show element names
        element_names = []
        for elem in elements:
            try:
                if hasattr(elem, 'Name'):
                    name = elem.Name
                elif hasattr(elem, 'ViewName'):
                    name = elem.ViewName
                else:
                    name = "Id: {}".format(elem.Id.IntegerValue)
                element_names.append(name)
            except:
                element_names.append("Id: {}".format(elem.Id.IntegerValue))
        
        if len(element_names) <= 5:
            self.text_element_names.Text = ", ".join(element_names)
        else:
            self.text_element_names.Text = "{}, ... ({} more)".format(
                ", ".join(element_names[:5]), 
                len(element_names) - 5
            )
        
        # Wire up events
        self.btn_load.Click += self.on_load_clicked
        self.btn_apply.Click += self.on_apply_clicked
        self.btn_close.Click += self.on_close_clicked
        
        # Initial UI update
        self.update_ui()
    
    def update_ui(self):
        """Update UI based on element type"""
        # Hide municipality section for AreaScheme (it's a data field)
        # Show it for Sheet/AreaPlan (read-only, detected from AreaScheme)
        if self._element_type == "AreaScheme":
            self.group_municipality.Visibility = System.Windows.Visibility.Collapsed
        else:
            self.group_municipality.Visibility = System.Windows.Visibility.Visible
            # Get and display municipality from elements
            municipality = self._get_municipality_from_elements()
            if municipality:
                self.text_municipality.Text = municipality
                self.text_municipality.Foreground = System.Windows.Media.Brushes.DarkGreen
            else:
                self.text_municipality.Text = "Not detected - Please define AreaScheme first"
                self.text_municipality.Foreground = System.Windows.Media.Brushes.Red
        
        self.update_fields()
    
    def update_fields(self):
        """Update field inputs based on element type and municipality"""
        # Clear existing fields
        self.panel_fields.Children.Clear()
        self._field_controls = {}
        
        element_type = self._element_type
        
        # Get municipality (only needed for Sheet/AreaPlan)
        if element_type == "AreaScheme":
            # For AreaScheme, municipality is a data field, not from combobox
            municipality = None
            self._current_municipality = None
        else:
            # For Sheet/AreaPlan, try to get municipality from first element
            municipality = self._get_municipality_from_elements()
            self._current_municipality = municipality
            
            if not municipality:
                msg = TextBlock()
                msg.Text = "No municipality defined. Please define AreaScheme first."
                msg.Foreground = System.Windows.Media.Brushes.Red
                msg.TextWrapping = System.Windows.TextWrapping.Wrap
                self.panel_fields.Children.Add(msg)
                return
        
        # Get field definitions
        try:
            if element_type == "AreaScheme":
                fields = municipality_schemas.AREASCHEME_FIELDS
            elif element_type == "Sheet":
                fields = municipality_schemas.SHEET_FIELDS.get(municipality, {})
            elif element_type == "AreaPlan (View)":
                fields = municipality_schemas.AREAPLAN_FIELDS.get(municipality, {})
            else:
                fields = {}
            
            if not fields:
                msg = TextBlock()
                msg.Text = "No fields defined for this configuration"
                msg.FontStyle = System.Windows.FontStyle.Italic
                msg.Foreground = System.Windows.Media.Brushes.Gray
                self.panel_fields.Children.Add(msg)
                return
            
            # Create field inputs
            for field_name, field_props in fields.items():
                self._create_field_input(field_name, field_props)
            
        except Exception as e:
            self._show_status("Error updating fields: {}".format(e), error=True)
    
    def _create_field_input(self, field_name, field_props):
        """Create input control for a field"""
        # Create grid for field
        grid = Grid()
        grid.Margin = System.Windows.Thickness(0, 5, 0, 5)
        
        # Define columns
        from System.Windows.Controls import ColumnDefinition
        col1 = ColumnDefinition()
        col1.Width = System.Windows.GridLength(150)
        col2 = ColumnDefinition()
        col2.Width = System.Windows.GridLength(1, System.Windows.GridUnitType.Star)
        grid.ColumnDefinitions.Add(col1)
        grid.ColumnDefinitions.Add(col2)
        
        # Label
        label = TextBlock()
        label.Text = field_name + ":"
        if field_props.get("required"):
            label.Text += " *"
            label.FontWeight = System.Windows.FontWeights.Bold
        Grid.SetColumn(label, 0)
        grid.Children.Add(label)
        
        # Input control
        field_type = field_props.get("type", "string")
        
        if field_name == "Municipality" and "options" in field_props:
            # ComboBox for municipality
            combo = ComboBox()
            for option in field_props["options"]:
                combo.Items.Add(option)
            combo.SelectedIndex = 0
            Grid.SetColumn(combo, 1)
            grid.Children.Add(combo)
            self._field_controls[field_name] = combo
            
        elif field_type == "int" and field_name in ["IS_UNDERGROUND"]:
            # CheckBox for boolean int
            checkbox = CheckBox()
            checkbox.Content = field_props.get("description", "")
            Grid.SetColumn(checkbox, 1)
            grid.Children.Add(checkbox)
            self._field_controls[field_name] = checkbox
            
        else:
            # TextBox for everything else
            textbox = TextBox()
            textbox.ToolTip = field_props.get("description", "")
            # Set default value if available
            default_value = field_props.get("default", "")
            if default_value:
                textbox.Text = default_value
            Grid.SetColumn(textbox, 1)
            grid.Children.Add(textbox)
            self._field_controls[field_name] = textbox
        
        self.panel_fields.Children.Add(grid)
    
    def _get_municipality_from_elements(self):
        """Get municipality from first element's AreaScheme"""
        if not self._elements:
            return None
        
        doc = revit.doc
        first_elem = self._elements[0]
        
        if self._element_type == "Sheet":
            return data_manager.get_municipality_from_sheet(doc, first_elem)
        elif self._element_type == "AreaPlan (View)":
            return data_manager.get_municipality_from_view(doc, first_elem)
        
        return None
    
    def on_load_clicked(self, sender, args):
        """Load existing data from first selected element"""
        if not self._elements:
            self._show_status("No elements selected", error=True)
            return
        
        try:
            first_elem = self._elements[0]
            data = data_manager.get_data(first_elem)
            
            if not data:
                self._show_status("No existing data found on element", error=False)
                return
            
            # Populate fields
            for field_name, control in self._field_controls.items():
                if field_name in data:
                    value = data[field_name]
                    
                    if isinstance(control, TextBox):
                        control.Text = str(value) if value is not None else ""
                    elif isinstance(control, ComboBox):
                        for item in control.Items:
                            if item == value:
                                control.SelectedItem = item
                                break
                    elif isinstance(control, CheckBox):
                        control.IsChecked = bool(value)
            
            self._show_status("Data loaded successfully", error=False)
            
        except Exception as e:
            self._show_status("Error loading data: {}".format(e), error=True)
    
    def on_apply_clicked(self, sender, args):
        """Apply data to selected elements"""
        if not self._elements:
            self._show_status("No elements selected", error=True)
            return
        
        try:
            # Collect data from fields
            data_dict = {}
            
            for field_name, control in self._field_controls.items():
                if isinstance(control, TextBox):
                    text = control.Text.strip()
                    if text:
                        # Try to convert to appropriate type
                        field_type = self._get_field_type(field_name)
                        if field_type == "float":
                            try:
                                data_dict[field_name] = float(text)
                            except:
                                data_dict[field_name] = text
                        elif field_type == "int":
                            try:
                                data_dict[field_name] = int(text)
                            except:
                                data_dict[field_name] = text
                        else:
                            data_dict[field_name] = text
                            
                elif isinstance(control, ComboBox):
                    if control.SelectedItem:
                        data_dict[field_name] = control.SelectedItem
                        
                elif isinstance(control, CheckBox):
                    data_dict[field_name] = 1 if control.IsChecked else 0
            
            # Apply to elements
            with revit.Transaction("Define pyArea Schema"):
                success_count = 0
                errors = []
                
                for elem in self._elements:
                    if self._element_type == "AreaScheme":
                        success = data_manager.set_data(elem, data_dict)
                    else:
                        success, err = self._apply_with_validation(elem, data_dict)
                        if not success:
                            errors.extend(err)
                    
                    if success:
                        success_count += 1
                
                if errors:
                    error_msg = "\n".join(set(errors))
                    self._show_status("Applied to {}/{} elements. Errors:\n{}".format(
                        success_count, len(self._elements), error_msg), error=True)
                else:
                    self._show_status("Successfully applied to {} element(s)".format(success_count), error=False)
        
        except Exception as e:
            self._show_status("Error applying data: {}".format(e), error=True)
    
    def _apply_with_validation(self, element, data_dict):
        """Apply data with validation based on element type"""
        element_type = self._element_type
        
        if element_type == "Sheet":
            return data_manager.set_sheet_data(element, data_dict, self._current_municipality)
        elif element_type == "AreaPlan (View)":
            return data_manager.set_areaplan_data(element, data_dict, self._current_municipality)
        else:
            success = data_manager.set_data(element, data_dict)
            return success, [] if success else ["Failed to store data"]
    
    def _get_field_type(self, field_name):
        """Get field type from schema definitions"""
        element_type = self._element_type
        
        try:
            if element_type == "AreaScheme":
                fields = municipality_schemas.AREASCHEME_FIELDS
            elif element_type == "Sheet":
                fields = municipality_schemas.SHEET_FIELDS.get(self._current_municipality, {})
            elif element_type == "AreaPlan (View)":
                fields = municipality_schemas.AREAPLAN_FIELDS.get(self._current_municipality, {})
            else:
                return "string"
            
            return fields.get(field_name, {}).get("type", "string")
            
        except:
            return "string"
    
    def _show_status(self, message, error=False):
        """Show status message"""
        self.text_status.Text = message
        if error:
            self.text_status.Foreground = System.Windows.Media.Brushes.Red
        else:
            self.text_status.Foreground = System.Windows.Media.Brushes.Green
    
    def on_close_clicked(self, sender, args):
        """Close the window"""
        self.Close()


# ==================== Main Script ====================

def get_all_area_schemes():
    """Get all AreaScheme elements in the document"""
    doc = revit.doc
    collector = DB.FilteredElementCollector(doc)
    area_schemes = collector.OfClass(DB.AreaScheme).ToElements()
    return list(area_schemes)


def get_selected_elements_or_pick():
    """Get selected elements or let user pick AreaSchemes"""
    doc = revit.doc
    selection = revit.get_selection()
    
    # Check if user has selected elements
    if selection:
        elements = list(selection)
        first_elem = elements[0]
        
        if isinstance(first_elem, DB.ViewSheet):
            element_type = "Sheet"
            # Filter to only Sheets
            elements = [e for e in elements if isinstance(e, DB.ViewSheet)]
            return elements, element_type
            
        elif isinstance(first_elem, DB.View):
            element_type = "AreaPlan (View)"
            # Filter to only Views
            elements = [e for e in elements if isinstance(e, DB.View)]
            return elements, element_type
    
    # No selection or invalid selection - ask user what they want to define
    options = ["AreaScheme", "Sheet", "AreaPlan (View)"]
    selected_type = forms.CommandSwitchWindow.show(
        options,
        message="What do you want to define?"
    )
    
    if not selected_type:
        script.exit()
    
    if selected_type == "AreaScheme":
        # Get all area schemes and let user pick
        area_schemes = get_all_area_schemes()
        
        if not area_schemes:
            forms.alert("No AreaSchemes found in the document.", exitscript=True)
        
        # Create display names for area schemes
        area_scheme_dict = {}
        for scheme in area_schemes:
            name = scheme.Name
            # Check if it already has municipality defined
            municipality = data_manager.get_municipality(scheme)
            if municipality:
                display_name = "{} [{}]".format(name, municipality)
            else:
                display_name = "{} [Not defined]".format(name)
            area_scheme_dict[display_name] = scheme
        
        # Let user select area schemes
        selected_names = forms.SelectFromList.show(
            sorted(area_scheme_dict.keys()),
            title="Select AreaSchemes",
            multiselect=True,
            button_name="Select"
        )
        
        if not selected_names:
            script.exit()
        
        elements = [area_scheme_dict[name] for name in selected_names]
        return elements, "AreaScheme"
    
    elif selected_type == "Sheet":
        forms.alert("Please select Sheet elements in Revit and run the tool again.", exitscript=True)
    
    elif selected_type == "AreaPlan (View)":
        forms.alert("Please select AreaPlan view elements in Revit and run the tool again.", exitscript=True)


if __name__ == '__main__':
    # Get elements (selected or picked)
    elements, element_type = get_selected_elements_or_pick()
    
    # Show dialog
    window = DefineSchemaWindow(elements, element_type)
    window.ShowDialog()
