# -*- coding: utf-8 -*-
__title__ = "Set Areas"
__doc__ = "Bulk set Usage Type and Usage Type Prev parameters for selected areas"

import csv
import os
import clr

from pyrevit import forms, revit, DB, script
from System.Collections.ObjectModel import ObservableCollection

# Get the current document
doc = revit.doc
uidoc = revit.uidoc

# Path to CSV file
lib_path = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(__file__))), "lib")
csv_path = os.path.join(lib_path, "UsageType_Common.csv")


class UsageTypeItem(object):
    """Class to represent a usage type item with color, number, and name"""
    def __init__(self, number, name, r=None, g=None, b=None):
        self.number = number
        self.name = name
        self.r = r
        self.g = g
        self.b = b
        self.has_color = r is not None and g is not None and b is not None
        
    def get_display_text(self):
        """Returns formatted display text for UI"""
        color_indicator = ""
        if self.has_color:
            color_indicator = "[RGB:{},{},{}] ".format(self.r, self.g, self.b)
        return "{}{} -> {}".format(color_indicator, self.number, self.name)
    
    @property
    def display_name(self):
        """Property for SelectFromList name_attr"""
        return self.get_display_text()
    
    def __str__(self):
        return self.get_display_text()
    
    def __repr__(self):
        return self.get_display_text()


def load_usage_types_from_csv():
    """Load usage types from CSV file"""
    usage_types = []
    try:
        # IronPython 2.7 doesn't support encoding parameter
        import codecs
        with codecs.open(csv_path, 'r', encoding='utf-8-sig') as f:
            reader = csv.DictReader(f)
            for row in reader:
                if row.get('usage_type') and row.get('name'):
                    try:
                        number = row['usage_type'].strip()
                        name = row['name'].strip()
                        r = int(row['R'].strip()) if row.get('R') and row['R'].strip() else None
                        g = int(row['G'].strip()) if row.get('G') and row['G'].strip() else None
                        b = int(row['B'].strip()) if row.get('B') and row['B'].strip() else None
                        
                        item = UsageTypeItem(number, name, r, g, b)
                        usage_types.append(item)
                    except ValueError:
                        continue
    except Exception as e:
        forms.alert("Error loading CSV file: {}".format(str(e)), exitscript=True)
    
    return usage_types


def get_current_parameter_value(areas, param_name):
    """Get the current parameter value from areas, returns 'Varies' if different"""
    values = set()
    for area in areas:
        param = area.LookupParameter(param_name)
        if param and not param.IsReadOnly:
            value = param.AsString() if param.StorageType == DB.StorageType.String else str(param.AsInteger()) if param.StorageType == DB.StorageType.Integer else None
            if value:
                values.add(value)
    
    if len(values) == 0:
        return None
    elif len(values) == 1:
        return list(values)[0]
    else:
        return "Varies"


# WPF Dialog removed - using SelectFromList instead

# Keeping this class definition for reference but not used
class _OldSetAreasWindow(forms.WPFWindow):
    """Dialog for setting area usage types"""
    
    def __init__(self, usage_types, current_usage_type=None, current_usage_type_prev=None):
        # XAML layout
        xaml_file = """
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Set Area Usage Types" Height="300" Width="650"
        WindowStartupLocation="CenterScreen" ResizeMode="NoResize"
        ShowInTaskbar="False">
    <Grid Margin="20">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="20"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        
        <!-- Usage Type -->
        <Label Grid.Row="0" Content="Usage Type:" FontWeight="Bold" FontSize="13"/>
        <ComboBox Grid.Row="1" x:Name="usage_type_cb" Height="30" FontSize="11"
                  IsEditable="True" IsTextSearchEnabled="False"/>
        
        <!-- Usage Type Prev -->
        <Label Grid.Row="3" Content="Usage Type Prev:" FontWeight="Bold" FontSize="13"/>
        <ComboBox Grid.Row="4" x:Name="usage_type_prev_cb" Height="30" FontSize="11"
                  IsEditable="True" IsTextSearchEnabled="False"/>
        
        <!-- Buttons -->
        <StackPanel Grid.Row="6" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,20,0,0">
            <Button Content="Apply" Width="100" Height="32" Margin="0,0,10,0"
                    x:Name="apply_btn" IsDefault="True"/>
            <Button Content="Cancel" Width="100" Height="32"
                    x:Name="cancel_btn" IsCancel="True"/>
        </StackPanel>
    </Grid>
</Window>
"""
        forms.WPFWindow.__init__(self, xaml_file, literal_string=True)
        
        self.usage_types = usage_types
        self.current_usage_type = current_usage_type
        self.current_usage_type_prev = current_usage_type_prev
        self.selected_usage_type = None
        self.selected_usage_type_prev = None
        
        # Store full list for filtering
        self.all_items_text = []
        for item in usage_types:
            self.all_items_text.append(item.get_display_text())
        
        # Populate ComboBoxes using ItemsSource
        self._populate_combo(self.usage_type_cb, usage_types, current_usage_type)
        self._populate_combo(self.usage_type_prev_cb, usage_types, current_usage_type_prev)
        
        # Wire up events for filtering and dropdown
        self.usage_type_cb.PreviewKeyDown += lambda s, e: self.on_key_down(self.usage_type_cb, e)
        self.usage_type_cb.DropDownOpened += self.dropdown_opened
        
        # Add TextChanged for live filtering - use AddHandler to get proper event
        import System
        from System.Windows.Controls import TextChangedEventHandler
        self.usage_type_cb.AddHandler(
            System.Windows.Controls.Primitives.TextBoxBase.TextChangedEvent,
            TextChangedEventHandler(lambda s, e: self.on_text_changed(self.usage_type_cb))
        )
        
        self.usage_type_prev_cb.PreviewKeyDown += lambda s, e: self.on_key_down(self.usage_type_prev_cb, e)
        self.usage_type_prev_cb.DropDownOpened += self.dropdown_opened
        self.usage_type_prev_cb.AddHandler(
            System.Windows.Controls.Primitives.TextBoxBase.TextChangedEvent,
            TextChangedEventHandler(lambda s, e: self.on_text_changed(self.usage_type_prev_cb))
        )
        
        # Wire up button events
        self.apply_btn.Click += self.apply_click
        self.cancel_btn.Click += self.cancel_click
    
    def _populate_combo(self, combo, items, current_value):
        """Populate combobox with items"""
        # Create observable collection for ItemsSource
        collection = ObservableCollection[object]()
        
        # Add "No Change" option
        collection.Add("-- No Change --")
        
        # Add current value indicator if available
        if current_value and current_value != "Varies":
            collection.Add("-- Current: {} --".format(current_value))
        elif current_value == "Varies":
            collection.Add("-- Current: Varies --")
        
        # Add all usage type items
        for item in items:
            collection.Add(item.get_display_text())
        
        # Set ItemsSource
        combo.ItemsSource = collection
        combo.SelectedIndex = 0
    
    def dropdown_opened(self, sender, e):
        """Handle dropdown opened event"""
        pass
    
    def on_key_down(self, combo, e):
        """Open dropdown when user starts typing"""
        # Open dropdown if not already open and user is typing
        if not combo.IsDropDownOpen:
            # Check if it's a character key (not navigation keys)
            from System.Windows.Input import Key
            if e.Key not in [Key.Tab, Key.Enter, Key.Escape, Key.Up, Key.Down, 
                            Key.Left, Key.Right, Key.Home, Key.End]:
                combo.IsDropDownOpen = True
    
    def on_text_changed(self, combo):
        """Filter combo items based on typed text"""
        search_text = combo.Text
        
        # Skip if empty or starts with "--"
        if not search_text or search_text.startswith("--"):
            return
        
        # Filter items that contain the search text (case-insensitive)
        search_lower = search_text.lower()
        filtered = [item for item in self.all_items_text 
                   if search_lower in item.lower()]
        
        # Only update if we have results
        if filtered:
            # Store the current text before updating ItemsSource
            current_text = search_text
            
            # Update ItemsSource with filtered items
            collection = ObservableCollection[object]()
            collection.Add("-- No Change --")
            
            for item_text in filtered:
                collection.Add(item_text)
            
            # Temporarily disable events to prevent recursion
            combo.ItemsSource = collection
            
            # Restore the text that user typed
            if combo.Text != current_text:
                combo.Text = current_text
            
            # Keep dropdown open
            if not combo.IsDropDownOpen:
                combo.IsDropDownOpen = True
    
    def apply_click(self, sender, e):
        """Handle apply button click"""
        # Get selected values from ComboBoxes
        usage_type_text = self.usage_type_cb.Text
        usage_type_prev_text = self.usage_type_prev_cb.Text
        
        # Parse selections
        if usage_type_text and not usage_type_text.startswith("--"):
            # Find matching item
            for item in self.usage_types:
                if item.get_display_text() == usage_type_text:
                    self.selected_usage_type = item
                    break
        
        if usage_type_prev_text and not usage_type_prev_text.startswith("--"):
            # Find matching item
            for item in self.usage_types:
                if item.get_display_text() == usage_type_prev_text:
                    self.selected_usage_type_prev = item
                    break
        
        self.DialogResult = True
        self.Close()
    
    def cancel_click(self, sender, e):
        """Handle cancel button click"""
        self.DialogResult = False
        self.Close()




def set_area_parameters(areas, usage_type_item, usage_type_prev_item):
    """Set parameters for selected areas"""
    count = 0
    errors = []
    
    with revit.Transaction("Set Area Usage Types"):
        for area in areas:
            try:
                # Set Usage Type parameters
                if usage_type_item:
                    # Set Usage Type (number only)
                    usage_type_param = area.LookupParameter("Usage Type")
                    if usage_type_param and not usage_type_param.IsReadOnly:
                        usage_type_param.Set(usage_type_item.number)
                    
                    # Set Usage Type Name (name string from CSV)
                    usage_type_name_param = area.LookupParameter("Usage Type Name")
                    if usage_type_name_param and not usage_type_name_param.IsReadOnly:
                        usage_type_name_param.Set(usage_type_item.name)
                    
                    # Set built-in Name parameter (name string from CSV)
                    name_param = area.LookupParameter("Name")
                    if name_param and not name_param.IsReadOnly:
                        name_param.Set(usage_type_item.name)
                    else:
                        # Try built-in parameter for Rooms and Areas
                        try:
                            if isinstance(area, DB.Architecture.Room):
                                bip = area.get_Parameter(DB.BuiltInParameter.ROOM_NAME)
                            elif isinstance(area, DB.Area):
                                bip = area.get_Parameter(DB.BuiltInParameter.AREA_NAME)
                            else:
                                bip = None
                            if bip and not bip.IsReadOnly:
                                bip.Set(usage_type_item.name)
                        except Exception:
                            pass
                
                # Set Usage Type Prev parameters
                if usage_type_prev_item:
                    # Set Usage Type Prev (number only)
                    usage_type_prev_param = area.LookupParameter("Usage Type Prev")
                    if usage_type_prev_param and not usage_type_prev_param.IsReadOnly:
                        usage_type_prev_param.Set(usage_type_prev_item.number)
                    
                    # Set Usage Type Prev. Name (name string from CSV)
                    usage_type_prev_name_param = area.LookupParameter("Usage Type Prev. Name")
                    if usage_type_prev_name_param and not usage_type_prev_name_param.IsReadOnly:
                        usage_type_prev_name_param.Set(usage_type_prev_item.name)
                
                count += 1
                
            except Exception as e:
                errors.append("Area {}: {}".format(area.Id.IntegerValue, str(e)))
    
    return count, errors


# Main execution
if __name__ == "__main__":
    # Get current selection
    selection = uidoc.Selection.GetElementIds()
    
    if not selection:
        forms.alert("Please select at least one area.", exitscript=True)
    
    # Filter for areas only
    areas = []
    for elem_id in selection:
        elem = doc.GetElement(elem_id)
        if isinstance(elem, DB.Architecture.Room) or isinstance(elem, DB.Area):
            areas.append(elem)
    
    if not areas:
        forms.alert("No areas found in selection. Please select areas.", exitscript=True)
    
    # Load usage types from CSV
    usage_types = load_usage_types_from_csv()
    
    if not usage_types:
        forms.alert("No usage types found in CSV file.", exitscript=True)
    
    # Get current parameter values from selected areas
    current_usage_type = get_current_parameter_value(areas, "Usage Type")
    current_usage_type_prev = get_current_parameter_value(areas, "Usage Type Prev")
    
    # Show first selection for Usage Type
    usage_type_options = ["-- No Change --"]
    if current_usage_type and current_usage_type != "Varies":
        usage_type_options.append("-- Current: {} --".format(current_usage_type))
    elif current_usage_type == "Varies":
        usage_type_options.append("-- Current: Varies --")
    usage_type_options.extend(usage_types)
    
    selected_usage_type_obj = forms.SelectFromList.show(
        usage_type_options,
        name_attr='display_name' if hasattr(usage_type_options[0], 'display_name') else None,
        title='Select Usage Type (searchable)',
        width=700,
        height=600,
        button_name='Select',
        multiselect=False
    )
    
    # Parse Usage Type selection
    selected_usage_type = None
    if selected_usage_type_obj and isinstance(selected_usage_type_obj, UsageTypeItem):
        selected_usage_type = selected_usage_type_obj
    
    # Show second selection for Usage Type Prev
    usage_type_prev_options = ["-- No Change --"]
    if current_usage_type_prev and current_usage_type_prev != "Varies":
        usage_type_prev_options.append("-- Current: {} --".format(current_usage_type_prev))
    elif current_usage_type_prev == "Varies":
        usage_type_prev_options.append("-- Current: Varies --")
    usage_type_prev_options.extend(usage_types)
    
    selected_usage_type_prev_obj = forms.SelectFromList.show(
        usage_type_prev_options,
        name_attr='display_name' if hasattr(usage_type_prev_options[0], 'display_name') else None,
        title='Select Usage Type Prev (searchable)',
        width=700,
        height=600,
        button_name='Select',
        multiselect=False
    )
    
    # Parse Usage Type Prev selection
    selected_usage_type_prev = None
    if selected_usage_type_prev_obj and isinstance(selected_usage_type_prev_obj, UsageTypeItem):
        selected_usage_type_prev = selected_usage_type_prev_obj
    
    # Check if any changes selected
    if selected_usage_type is None and selected_usage_type_prev is None:
        forms.alert("No changes selected.")
        script.exit()
    
    # Apply changes
    count, errors = set_area_parameters(
        areas, 
        selected_usage_type, 
        selected_usage_type_prev
    )
    
    # Show results only if there are errors
    if errors:
        error_msg = "Updated {} areas with {} errors:\n\n{}".format(
            count, len(errors), "\n".join(errors[:10]))
        forms.alert(error_msg)
