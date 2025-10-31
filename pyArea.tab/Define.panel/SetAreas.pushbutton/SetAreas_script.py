# -*- coding: utf-8 -*-
"""Set Usage Type and Usage Type Prev for Area Elements"""
__title__ = "Set Areas"
__doc__ = "Bulk set Usage Type and Usage Type Prev parameters for selected areas"

import csv
import os
from pyrevit import revit, DB, forms, script
from colored_combobox import ColoredComboBox


class SetAreasWindow(forms.WPFWindow):
    """Dialog window for setting area usage types"""
    
    def __init__(self, area_elements, options_list):
        forms.WPFWindow.__init__(self, 'SetAreasWindow.xaml')
        self._areas = area_elements
        self._options = options_list
        self.usage_type_value = None
        self.usage_type_prev_value = None
        
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
        
        # Set initial values from the first selected area's parameters
        if area_elements:
            first_area = area_elements[0]
            
            # Get Usage Type parameter value (number only)
            usage_type_param = first_area.LookupParameter("Usage Type")
            if usage_type_param and usage_type_param.HasValue:
                number = usage_type_param.AsString()
                # Find matching option and set the full "number. name" display text
                display_value = self._find_display_text_by_number(number)
                if display_value:
                    self._combo_usage_type.set_initial_value(display_value)
            
            # Get Usage Type Prev parameter value (number only)
            usage_type_prev_param = first_area.LookupParameter("Usage Type Prev")
            if usage_type_prev_param and usage_type_prev_param.HasValue:
                number = usage_type_prev_param.AsString()
                # Find matching option and set the full "number. name" display text
                display_value = self._find_display_text_by_number(number)
                if display_value:
                    self._combo_usage_type_prev.set_initial_value(display_value)
        
        # Update info text
        area_count = len(area_elements)
        if area_count > 0:
            self.text_info.Text = "{} area element(s) selected".format(area_count)
    
    def _find_display_text_by_number(self, number):
        """Find the full display text (number. name) from just the number"""
        for option in self._options:
            # Options can be tuples (text, color) or just strings
            text = option[0] if isinstance(option, tuple) else option
            # Text format is "number. name"
            if text.startswith(number + ". "):
                return text
        return None
    
    def apply_clicked(self, sender, args):
        """Handle Apply button click"""
        usage_type_text = self._combo_usage_type.get_text()
        usage_type_prev_text = self._combo_usage_type_prev.get_text()
        
        # At least one field must be filled
        if not usage_type_text and not usage_type_prev_text:
            forms.alert("Please select or enter at least one usage type value.", exitscript=False)
            return
        
        # Parse the values (format: "number. name")
        if usage_type_text:
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
        
        if usage_type_prev_text:
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
        
        self.DialogResult = True
        self.Close()
    
    def cancel_clicked(self, sender, args):
        """Handle Cancel button click"""
        self.DialogResult = False
        self.Close()


def load_usage_types_from_csv():
    """
    Load usage type options from CSV file.
    
    Returns list of tuples: (text, (R, G, B)) or just text strings
    """
    # Path to CSV file - lib folder is in the pyArea.tab directory
    # __file__ is in: pyArea.tab\Define.panel\SetAreas.pushbutton\
    # CSV is in: pyArea.tab\lib\
    script_dir = os.path.dirname(__file__)
    tab_path = os.path.dirname(os.path.dirname(script_dir))  # Go up to pyArea.tab
    csv_path = os.path.join(tab_path, "lib", "UsageType_Common.csv")
    
    options = []
    
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
        forms.alert("Error loading CSV file: {}\nPath: {}".format(str(e), csv_path), exitscript=True)
    
    return options


def main():
    # Get selected area elements
    area_elements = [el for el in revit.get_selection() 
                     if isinstance(el, DB.Area) or isinstance(el, DB.Architecture.Room)]
    
    if not area_elements:
        forms.alert("Please select at least one area element.", exitscript=True)
    
    # Load usage type options from CSV
    options_list = load_usage_types_from_csv()
    
    if not options_list:
        forms.alert("No usage types found in CSV file.", exitscript=True)
    
    # Show dialog
    dialog = SetAreasWindow(area_elements, options_list)
    dialog.ShowDialog()
    
    # If user clicked Apply
    if dialog.DialogResult:
        # Start transaction to modify parameters
        with revit.Transaction("Set Area Usage Types"):
            updated_usage_type = 0
            updated_usage_type_prev = 0
            failed_usage_type = 0
            failed_usage_type_prev = 0
            
            for area in area_elements:
                # Update Usage Type if provided
                if dialog.usage_type_value:
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
                
                # Update Usage Type Prev if provided
                if dialog.usage_type_prev_value:
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
            
            # Log results
            logger = script.get_logger()
            if dialog.usage_type_value:
                logger.info("Usage Type: Updated {} area(s), Failed: {}".format(
                    updated_usage_type, failed_usage_type))
            
            if dialog.usage_type_prev_value:
                logger.info("Usage Type Prev: Updated {} area(s), Failed: {}".format(
                    updated_usage_type_prev, failed_usage_type_prev))


if __name__ == "__main__":
    main()
