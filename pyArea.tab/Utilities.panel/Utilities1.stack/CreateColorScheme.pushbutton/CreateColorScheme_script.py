# -*- coding: utf-8 -*-
"""Create or update color scheme for areas."""

__title__ = "Create\nColor\nScheme"
__author__ = "Your Name"

import os
import csv
from pyrevit import revit, DB, forms, script

# Get the directory where this script is located
SCRIPT_DIR = os.path.dirname(__file__)
LIB_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(SCRIPT_DIR))), "lib")

logger = script.get_logger()


class CreateColorSchemeWindow(forms.WPFWindow):
    """Window for creating/updating area color schemes."""
    
    def __init__(self, xaml_file_name):
        forms.WPFWindow.__init__(self, xaml_file_name)
        self._area_schemes = []
        self._parameters = []
        self._existing_color_schemes = {}
        
        self.setup_area_schemes()
        self.setup_municipalities()
        self.setup_parameters()
        
        # Populate color scheme dropdown after all dropdowns are set up
        self.update_color_scheme_dropdown()
        
    def setup_area_schemes(self):
        """Populate area scheme dropdown."""
        # Get all area schemes in the document
        area_schemes = DB.FilteredElementCollector(revit.doc)\
                         .OfClass(DB.AreaScheme)\
                         .ToElements()
        
        self._area_schemes = sorted(area_schemes, key=lambda x: x.Name)
        scheme_names = [scheme.Name for scheme in self._area_schemes]
        
        if scheme_names:
            self.area_scheme_cb.ItemsSource = scheme_names
            self.area_scheme_cb.SelectedIndex = 0
        else:
            forms.alert("No area schemes found in the document.", exitscript=True)
    
    def setup_municipalities(self):
        """Populate municipality dropdown."""
        municipalities = ["Common", "Jerusalem", "Tel-Aviv"]
        self.municipality_cb.ItemsSource = municipalities
        self.municipality_cb.SelectedIndex = 0
    
    def setup_parameters(self):
        """Populate parameter dropdown with Area element parameters."""
        # Get a sample area to extract parameters
        areas = DB.FilteredElementCollector(revit.doc)\
                  .OfCategory(DB.BuiltInCategory.OST_Areas)\
                  .WhereElementIsNotElementType()\
                  .ToElements()
        
        if not areas:
            forms.alert("No areas found in the document. Please create areas first.", 
                       exitscript=True)
        
        sample_area = areas[0]
        self._sample_area = sample_area
        params = []
        
        # Get all parameters from the area
        for param in sample_area.Parameters:
            if param.StorageType == DB.StorageType.String or \
               param.StorageType == DB.StorageType.Integer:
                params.append(param.Definition.Name)
        
        self._parameters = sorted(set(params))
        
        if self._parameters:
            self.parameter_cb.ItemsSource = self._parameters
            
            # Try to set default to 'Usage Type', otherwise use first parameter
            default_param = "Usage Type"
            if default_param in self._parameters:
                self.parameter_cb.SelectedItem = default_param
                logger.debug("Set default parameter to: {}".format(default_param))
            else:
                self.parameter_cb.SelectedIndex = 0
                logger.debug("'Usage Type' not found, using first parameter: {}".format(self._parameters[0]))
        else:
            forms.alert("No suitable parameters found on Area elements.", 
                       exitscript=True)
    
    def get_existing_color_schemes(self, area_scheme, parameter_name):
        """Get existing color schemes for the selected area scheme (all parameters)."""
        color_schemes = []
        
        try:
            # Get all color fill schemes
            color_fill_schemes = DB.FilteredElementCollector(revit.doc)\
                                   .OfClass(DB.ColorFillScheme)\
                                   .ToElements()
            
            area_category = DB.Category.GetCategory(revit.doc, DB.BuiltInCategory.OST_Areas)
            area_category_int = area_category.Id.Value if area_category else None
            target_area_id_int = area_scheme.Id.Value
            
            for scheme in color_fill_schemes:
                try:
                    scheme_name = getattr(scheme, 'Name', '<unnamed>')
                    # Check category match
                    scheme_cat = getattr(scheme, 'CategoryId', None)
                    if not scheme_cat:
                        continue
                    if area_category_int is not None and scheme_cat.Value != area_category_int:
                        continue
                    
                    # Check area scheme match
                    scheme_area_id = getattr(scheme, 'AreaSchemeId', None)
                    if not scheme_area_id:
                        continue
                    if scheme_area_id.Value != target_area_id_int:
                        continue
                    
                    # Add all schemes for this area scheme, regardless of parameter
                    color_schemes.append(scheme_name)
                    logger.debug("Found existing scheme '{}' for area scheme '{}'".format(
                        scheme_name, area_scheme.Name))
                except Exception as e2:
                    logger.debug("Error checking scheme {}: {}".format(getattr(scheme, 'Name', 'Unknown'), e2))
        except Exception as e:
            logger.warning("Error getting color schemes: {}".format(e))
        
        return sorted(color_schemes)
    
    def area_scheme_changed(self, sender, args):
        """Handle area scheme selection change."""
        self.update_color_scheme_dropdown()
    
    def parameter_changed(self, sender, args):
        """Handle parameter selection change."""
        # No need to update color scheme dropdown - it shows all schemes regardless of parameter
        pass
    
    def update_color_scheme_dropdown(self):
        """Update the color scheme dropdown based on selected area scheme."""
        if self.area_scheme_cb.SelectedIndex >= 0:
            area_scheme = self._area_schemes[self.area_scheme_cb.SelectedIndex]
            # parameter_name is not used anymore, but kept for compatibility
            parameter_name = self.parameter_cb.SelectedItem if self.parameter_cb.SelectedIndex >= 0 else None
            
            existing_schemes = self.get_existing_color_schemes(area_scheme, parameter_name)
            self.scheme_name_cb.ItemsSource = existing_schemes
            
            # Clear the text field if no existing schemes
            if not existing_schemes:
                self.scheme_name_cb.Text = ""
            # Or set to first scheme if available
            elif self.scheme_name_cb.Text not in existing_schemes:
                self.scheme_name_cb.Text = existing_schemes[0]
    
    def cancel_click(self, sender, args):
        """Handle cancel button click."""
        self.Close()
    
    def create_click(self, sender, args):
        """Handle create/update button click."""
        # Validate inputs
        if self.area_scheme_cb.SelectedIndex < 0:
            forms.alert("Please select an area scheme.")
            return
        
        if self.municipality_cb.SelectedIndex < 0:
            forms.alert("Please select a municipality.")
            return
        
        if self.parameter_cb.SelectedIndex < 0:
            forms.alert("Please select a parameter.")
            return
        
        if not self.scheme_name_cb.Text:
            forms.alert("Please enter a color scheme name.")
            return
        
        # Get selected values
        area_scheme = self._area_schemes[self.area_scheme_cb.SelectedIndex]
        municipality = self.municipality_cb.SelectedItem
        parameter_name = self.parameter_cb.SelectedItem
        scheme_name = self.scheme_name_cb.Text
        
        # Close the window
        self.Close()
        
        # Process the color scheme creation/update
        process_color_scheme(area_scheme, municipality, parameter_name, scheme_name)


def read_csv_data(municipality):
    """Read CSV data for the specified municipality."""
    csv_filename = "UsageType_{}.csv".format(municipality)
    csv_path = os.path.join(LIB_DIR, csv_filename)
    
    logger.debug("Looking for CSV file: {}".format(csv_path))
    
    if not os.path.exists(csv_path):
        forms.alert("CSV file not found: {}".format(csv_filename), exitscript=True)
    
    logger.debug("CSV file found, reading data...")
    
    data = []
    row_count = 0
    skipped_count = 0
    
    try:
        import codecs
        with codecs.open(csv_path, 'r', encoding='utf-8-sig') as f:  # utf-8-sig handles BOM
            reader = csv.DictReader(f)
            logger.debug("CSV headers: {}".format(reader.fieldnames))
            
            for row in reader:
                row_count += 1
                logger.debug("Row {}: {}".format(row_count, row))
                
                # Get values and strip whitespace
                usage_type = row.get('usage_type', '').strip()
                r_val = row.get('R', '').strip()
                g_val = row.get('G', '').strip()
                b_val = row.get('B', '').strip()
                
                # Skip rows with empty usage_type or missing RGB values
                if not usage_type or not r_val or not g_val or not b_val:
                    logger.debug("Skipping row {} - missing data: usage_type='{}', R='{}', G='{}', B='{}'".format(
                        row_count, usage_type, r_val, g_val, b_val))
                    skipped_count += 1
                    continue
                
                try:
                    data.append({
                        'usage_type': usage_type,
                        'name': row.get('name', '').strip(),
                        'R': int(r_val),
                        'G': int(g_val),
                        'B': int(b_val)
                    })
                    logger.debug("Added entry: {} - {} ({}, {}, {})".format(
                        usage_type, row.get('name', '').strip(), r_val, g_val, b_val))
                except ValueError as e:
                    logger.warning("Skipping row {} with invalid RGB values: {} - Error: {}".format(
                        row_count, row, e))
                    skipped_count += 1
                    
    except Exception as e:
        logger.error("Error reading CSV file: {}".format(e))
        forms.alert("Error reading CSV file: {}".format(e), exitscript=True)
    
    logger.debug("CSV reading complete: {} valid entries, {} skipped".format(len(data), skipped_count))
    print("CSV Data Summary: {} valid entries found from {} total rows".format(len(data), row_count))
    
    return data, csv_filename


def get_parameter_id(parameter_name):
    """Get the parameter ID for the given parameter name."""
    # Try to find the parameter definition
    areas = DB.FilteredElementCollector(revit.doc)\
              .OfCategory(DB.BuiltInCategory.OST_Areas)\
              .WhereElementIsNotElementType()\
              .FirstElement()
    
    if areas:
        for param in areas.Parameters:
            if param.Definition.Name == parameter_name:
                # Check if it's a shared parameter
                if hasattr(param, 'IsShared') and param.IsShared:
                    return param.Id
                # For project parameters, we need to find the parameter element
                else:
                    # Try to find parameter element by name
                    param_elements = DB.FilteredElementCollector(revit.doc)\
                                       .OfClass(DB.ParameterElement)\
                                       .ToElements()
                    for pe in param_elements:
                        if hasattr(pe, 'Name') and pe.Name == parameter_name:
                            return pe.Id
    
    return None


def process_color_scheme(area_scheme, municipality, parameter_name, scheme_name):
    """Create or update the color scheme."""
    logger.debug("=== Starting Color Scheme Process ===")
    logger.debug("Area Scheme: {}".format(area_scheme.Name))
    logger.debug("Municipality: {}".format(municipality))
    logger.debug("Parameter: {}".format(parameter_name))
    logger.debug("Scheme Name: {}".format(scheme_name))
    
    # Read CSV data
    csv_data, csv_filename = read_csv_data(municipality)
    
    logger.debug("CSV data entries: {}".format(len(csv_data)))
    
    if not csv_data:
        logger.error("No valid data found in CSV file")
        forms.alert("No valid data found in CSV file.", exitscript=True)
    
    print("Successfully loaded {} entries from CSV".format(len(csv_data)))
    
    # Get parameter ID
    logger.debug("Looking for parameter: {}".format(parameter_name))
    param_id = get_parameter_id(parameter_name)
    
    if not param_id:
        logger.error("Could not find parameter: {}".format(parameter_name))
        forms.alert("Could not find parameter: {}".format(parameter_name), exitscript=True)
    
    logger.debug("Parameter ID found: {}".format(param_id))
    
    # Check if color scheme already exists
    existing_scheme = None
    color_fill_schemes = DB.FilteredElementCollector(revit.doc)\
                           .OfClass(DB.ColorFillScheme)\
                           .ToElements()
    
    for scheme in color_fill_schemes:
        if scheme.Name == scheme_name:
            existing_scheme = scheme
            break
    
    logger.debug("Checking for existing scheme...")
    if existing_scheme:
        logger.debug("Found existing scheme: {}".format(existing_scheme.Name))
    else:
        logger.debug("No existing scheme found, will create new")
    
    # Start transaction
    with revit.Transaction("Create/Update Color Scheme"):
        try:
            if existing_scheme:
                # Update existing scheme
                color_scheme = existing_scheme
                logger.info("Updating existing color scheme: {}".format(scheme_name))
                print("Updating existing color scheme: {}".format(scheme_name))
            else:
                # Create new color scheme by duplicating an existing one
                logger.debug("Creating new color scheme...")
                
                # Find an existing color scheme for Areas to duplicate
                area_category_id = DB.Category.GetCategory(revit.doc, 
                                                          DB.BuiltInCategory.OST_Areas).Id
                template_scheme = None
                
                # CRITICAL: Must find a scheme with the SAME AreaSchemeId
                # because AreaSchemeId is read-only and cannot be changed after duplication
                logger.debug("Looking for color schemes for Area Scheme: {} (ID: {})".format(
                    area_scheme.Name, area_scheme.Id))
                logger.debug("Total color fill schemes found: {}".format(len(color_fill_schemes)))

                # Pick the FIRST scheme that:
                # - belongs to Areas category
                # - has AreaSchemeId matching the selected area scheme
                # Parameter can be changed later via ParameterDefinition
                for scheme in color_fill_schemes:
                    try:
                        if not hasattr(scheme, 'CategoryId'):
                            continue
                        if scheme.CategoryId.Value != area_category_id.Value:
                            continue
                        if not hasattr(scheme, 'AreaSchemeId'):
                            continue
                        if scheme.AreaSchemeId.Value != area_scheme.Id.Value:
                            continue

                        template_scheme = scheme
                    except Exception as e:
                        logger.debug("Error checking scheme {}: {}".format(getattr(scheme, 'Name', 'Unknown'), e))
                        import traceback
                        logger.debug(traceback.format_exc())
                
                logger.debug("Template scheme found: {}".format(template_scheme.Name if template_scheme else "None"))
                
                if not template_scheme:
                    # No template found with matching area scheme
                    logger.error("No Area color scheme found for area scheme: {}".format(area_scheme.Name))
                    
                    error_msg = "No existing color scheme found for Area Scheme '{}'.\n\n".format(area_scheme.Name)
                    
                    error_msg += "To create a new color scheme:\n"
                    error_msg += "1. Open an Area Plan view for '{}'\n".format(area_scheme.Name)
                    error_msg += "2. Go to View tab > Color Fill Legend\n"
                    error_msg += "3. Create a color scheme manually (any parameter)\n"
                    error_msg += "4. Then run this tool again to populate it with data.\n\n"
                    error_msg += "Note: The color scheme must be created for the specific Area Scheme."
                    
                    forms.alert(error_msg, exitscript=True)
                
                # Duplicate the scheme (AreaSchemeId is inherited and cannot be changed)
                logger.debug("Duplicating scheme: {}".format(template_scheme.Name))
                new_scheme_id = template_scheme.Duplicate(scheme_name)
                color_scheme = revit.doc.GetElement(new_scheme_id)
                logger.debug("Duplicated scheme created with ID: {}".format(new_scheme_id))
                
                # Verify the AreaSchemeId is correct (it should be inherited from template)
                logger.debug("Verifying AreaSchemeId...")
                if color_scheme.AreaSchemeId.Value == area_scheme.Id.Value:
                    logger.debug("AreaSchemeId is correct: {}".format(area_scheme.Name))
                else:
                    logger.error("AreaSchemeId mismatch!")
                    forms.alert("Error: Duplicated scheme has wrong Area Scheme association.", exitscript=True)
                
                logger.info("Created new color scheme: {}".format(scheme_name))
            
            # Track old parameter for reporting (only for existing schemes)
            old_parameter_name = None
            parameter_changed = False
            if existing_scheme:
                try:
                    # Get the current parameter before changing it
                    old_param_id = color_scheme.ParameterDefinition
                    if old_param_id:
                        try:
                            old_param_elem = revit.doc.GetElement(old_param_id)
                            if old_param_elem and hasattr(old_param_elem, 'Name'):
                                old_parameter_name = old_param_elem.Name
                            else:
                                # Try to get from definition
                                try:
                                    defn = old_param_elem.GetDefinition()
                                    if defn:
                                        old_parameter_name = defn.Name
                                except:
                                    pass
                        except:
                            pass
                        # Check if parameter is actually changing
                        if old_param_id.Value != param_id.Value:
                            parameter_changed = True
                            logger.debug("Parameter will change from '{}' to '{}'".format(old_parameter_name, parameter_name))
                except Exception as e:
                    logger.debug("Could not get old parameter: {}".format(e))
            
            # Set the parameter for both new and existing schemes
            logger.debug("Setting parameter to: {}".format(parameter_name))
            try:
                color_scheme.ParameterDefinition = param_id
                logger.debug("Parameter set successfully")
            except Exception as e:
                error_msg = str(e)
                logger.error("Failed to set parameter: {}".format(error_msg))
                
                # Check if this is the "paramId cannot be applied" error
                if "paramId cannot be applied" in error_msg or "Colors are not preserved" in error_msg:
                    # This happens when changing parameter would invalidate existing colors
                    msg = "Cannot change parameter on existing color scheme '{}' from '{}' to '{}'.\n\n".format(
                        scheme_name, old_parameter_name or "unknown", parameter_name)
                    msg += "Revit does not allow changing the parameter when it would invalidate existing color entries.\n\n"
                    msg += "Options:\n"
                    msg += "1. Create a new color scheme with a different name\n"
                    msg += "2. Manually delete the existing scheme in Revit and create a new one\n"
                    msg += "3. Keep the existing parameter and update only the colors"
                    
                    forms.alert(msg, exitscript=True)
                else:
                    forms.alert("Failed to set parameter on color scheme: {}".format(error_msg), exitscript=True)
            
            # Set the title to the parameter name
            logger.debug("Setting title to parameter name: {}".format(parameter_name))
            try:
                color_scheme.Title = parameter_name
                logger.debug("Title set successfully")
            except Exception as e:
                logger.warning("Could not set title: {}".format(e))
            
            # Simplified update logic
            storage_type = color_scheme.StorageType
            is_duplicated = not existing_scheme
            
            # Helper functions
            def get_value_key(entry):
                try:
                    if storage_type == DB.StorageType.String:
                        return entry.GetStringValue()
                    elif storage_type == DB.StorageType.Integer:
                        return entry.GetIntegerValue()
                    elif storage_type == DB.StorageType.Double:
                        return entry.GetDoubleValue()
                except:
                    pass
                return None
            
            def get_color_tuple(entry):
                try:
                    c = entry.Color
                    return (c.Red, c.Green, c.Blue)
                except:
                    return (0, 0, 0)
            
            # Prepare CSV entries first
            csv_by_key = {}
            for row in csv_data:
                try:
                    if storage_type == DB.StorageType.String:
                        key = row['usage_type']
                    elif storage_type == DB.StorageType.Integer:
                        key = int(row['usage_type'])
                    elif storage_type == DB.StorageType.Double:
                        key = float(row['usage_type'])
                    csv_by_key[key] = row
                except:
                    pass
            
            # Track changes
            created = []
            modified = []
            deleted = []
            non_compliant = []
            
            # Get solid fill pattern
            solid_pattern = None
            for fp in DB.FilteredElementCollector(revit.doc).OfClass(DB.FillPatternElement):
                if fp.GetFillPattern().IsSolidFill:
                    solid_pattern = fp
                    break
            
            # Get existing entries and build lookup BEFORE any modifications
            existing = list(color_scheme.GetEntries())
            existing_by_key = {}
            for e in existing:
                key = get_value_key(e)
                if key is not None:
                    existing_by_key[key] = e
            
            # Now process deletions
            if is_duplicated:
                # For new schemes: clear what we can, report the rest as non-compliant
                for key, entry in existing_by_key.items():
                    try:
                        color_scheme.RemoveEntry(entry)
                    except:
                        # Can't remove - Revit auto-created it
                        if key not in csv_by_key:
                            non_compliant.append((key, get_color_tuple(entry), getattr(entry, 'Caption', ''), 'Auto-created by Revit'))
                # Clear the lookup since we deleted everything we could
                existing_by_key = {}
            else:
                # For existing schemes: delete entries not in CSV
                keys_to_remove = []
                for key, entry in existing_by_key.items():
                    if key not in csv_by_key:
                        try:
                            color_scheme.RemoveEntry(entry)
                            deleted.append((key, get_color_tuple(entry), getattr(entry, 'Caption', '')))
                            keys_to_remove.append(key)
                        except:
                            non_compliant.append((key, get_color_tuple(entry), getattr(entry, 'Caption', ''), 'Cannot be removed'))
                            keys_to_remove.append(key)
                # Remove deleted keys from lookup
                for key in keys_to_remove:
                    existing_by_key.pop(key, None)
            
            # Add/update entries from CSV
            for key, row in csv_by_key.items():
                if key in existing_by_key:
                    # Entry already exists - check if we need to update it
                    entry = existing_by_key[key]
                    try:
                        # Capture current state BEFORE any modifications
                        old_color = get_color_tuple(entry)
                        old_caption = getattr(entry, 'Caption', '')
                        old_pattern_id = entry.FillPatternId
                        
                        # Desired state from CSV
                        new_color = (row['R'], row['G'], row['B'])
                        new_caption = row.get('name', '')
                        
                        # Check if anything actually changed
                        color_changed = old_color != new_color
                        caption_changed = new_caption and old_caption != new_caption
                        pattern_changed = solid_pattern and entry.FillPatternId != solid_pattern.Id
                        
                        # Debug logging for this entry
                        logger.debug("Checking entry: {}".format(key))
                        logger.debug("  Old color: {}, New color: {}, Changed: {}".format(old_color, new_color, color_changed))
                        logger.debug("  Old caption: '{}', New caption: '{}', Changed: {}".format(old_caption, new_caption, caption_changed))
                        logger.debug("  Old pattern ID: {}, Solid pattern ID: {}, Changed: {}".format(
                            old_pattern_id, solid_pattern.Id if solid_pattern else None, pattern_changed))
                        
                        if color_changed or caption_changed or pattern_changed:
                            logger.debug("  -> Entry will be MODIFIED (color:{}, caption:{}, pattern:{})".format(
                                color_changed, caption_changed, pattern_changed))
                            
                            # Save old color for reporting BEFORE modifying
                            report_old_color = old_color
                            report_new_color = new_color
                            report_caption = new_caption or old_caption
                            
                            # Now modify the entry
                            if color_changed:
                                entry.Color = DB.Color(*new_color)
                            if caption_changed:
                                entry.Caption = new_caption
                            if pattern_changed:
                                entry.FillPatternId = solid_pattern.Id
                            
                            # Only add to modified list if COLOR actually changed
                            # (caption and pattern changes are silent updates)
                            if color_changed:
                                modified.append((key, report_old_color, report_new_color, report_caption))
                        else:
                            logger.debug("  -> Entry unchanged, skipping")
                    except Exception as ex:
                        logger.error("Error processing entry {}: {}".format(key, ex))
                        import traceback
                        logger.error(traceback.format_exc())
                else:
                    # Create new entry
                    try:
                        entry = DB.ColorFillSchemeEntry(storage_type)
                        entry.Color = DB.Color(row['R'], row['G'], row['B'])
                        if storage_type == DB.StorageType.String:
                            entry.SetStringValue(row['usage_type'])
                        elif storage_type == DB.StorageType.Integer:
                            entry.SetIntegerValue(int(row['usage_type']))
                        elif storage_type == DB.StorageType.Double:
                            entry.SetDoubleValue(float(row['usage_type']))
                        if row['name']:
                            entry.Caption = row['name']
                        if solid_pattern:
                            entry.FillPatternId = solid_pattern.Id
                        color_scheme.AddEntry(entry)
                        created.append((key, (row['R'], row['G'], row['B']), row.get('name', '')))
                    except:
                        pass  # Silently ignore creation errors
            
            added_count = len(created)
            
            logger.info("Successfully added {} entries".format(added_count))
            # Plain text headings, color squares via print_html per line
            out = script.get_output()
            def square_html(rgb):
                return '<span style="display:inline-block;width:10px;height:10px;border:1px solid #666;margin-right:6px;background-color: rgb({},{},{});"></span>'.format(rgb[0], rgb[1], rgb[2])

            print("Color Scheme Update Report")
            print("Color scheme: {} | Area Scheme: {} | Parameter: {}".format(scheme_name, area_scheme.Name, parameter_name))
            if parameter_changed and old_parameter_name:
                print("Parameter changed from '{}' to '{}'".format(old_parameter_name, parameter_name))
            print("Source CSV: {}".format(csv_filename))

            print("\nCreated ({}):".format(len(created)))
            for val, rgb, cap in created:
                text = " {}{}{}".format(val, " - " if cap else "", cap or "")
                out.print_html(square_html(rgb) + text)

            # Show Modified section only when updating existing scheme
            if not is_duplicated:
                if modified:
                    print("\nModified ({}):".format(len(modified)))
                    for val, oldrgb, newrgb, cap in modified:
                        text = " {}{}{}".format(val, " - " if cap else "", cap or "")
                        out.print_html(square_html(oldrgb) + '&nbsp;→&nbsp;' + square_html(newrgb) + text)

                print("\nDeleted ({}):".format(len(deleted)))
                for val, rgb, cap in deleted:
                    text = " {}{}{}".format(val, " - " if cap else "", cap or "")
                    out.print_html(square_html(rgb) + text)

            # Report non-compliant entries
            if non_compliant:
                print("\nNon-compliant entries ({}):".format(len(non_compliant)))
                for val, rgb, cap, err in non_compliant:
                    text = " {}{}{} ({})".format(val, " - " if cap else "", cap or "", err)
                    out.print_html(square_html(rgb) + text)
            
        except Exception as e:
            import traceback
            logger.error("Error creating/updating color scheme: {}".format(e))
            logger.error("Traceback: {}".format(traceback.format_exc()))
            print("ERROR: {}".format(e))
            print(traceback.format_exc())
            forms.alert("Error: {}".format(e))


if __name__ == '__main__':
    # Show the window
    CreateColorSchemeWindow('CreateColorSchemeWindow.xaml').ShowDialog()
