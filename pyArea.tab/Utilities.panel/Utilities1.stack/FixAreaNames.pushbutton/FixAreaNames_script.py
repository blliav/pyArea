# -*- coding: utf-8 -*-
"""Correct area names by usage types."""

__title__ = "Fix Area\nNames"
__author__ = "Your Name"

from pyrevit import revit, DB, forms, script
import os
import sys
import csv
import codecs

# Import WPF for dialog
import clr
clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
import System
from System.Windows import Window
from System.Windows.Controls import CheckBox, StackPanel

SCRIPT_DIR = os.path.dirname(__file__)
LIB_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(SCRIPT_DIR))), "lib")

if LIB_DIR not in sys.path:
    sys.path.insert(0, LIB_DIR)

import data_manager
from schemas.municipality_schemas import get_usage_type_csv_filename

logger = script.get_logger()


# ============================================================
# VIEW SELECTION DIALOG (reused from FillHoles)
# ============================================================

class ViewSelectionDialog(forms.WPFWindow):
    """Dialog for selecting AreaPlan views to process"""
    
    def __init__(self, area_schemes, selected_scheme_index=0, preselected_view_ids=None):
        """
        Args:
            area_schemes: List of (AreaScheme element, municipality) tuples
            selected_scheme_index: Index of the area scheme to select by default
            preselected_view_ids: Set of view ElementIds to preselect
        """
        forms.WPFWindow.__init__(self, 'ViewSelectionDialog.xaml')
        
        self._doc = revit.doc
        self._area_schemes = area_schemes
        self._preselected_view_ids = preselected_view_ids or set()
        self._view_checkboxes = {}  # {view_id_value: checkbox}
        
        # Wire up events
        self.btn_ok.Click += self.on_ok_clicked
        self.btn_cancel.Click += self.on_cancel_clicked
        self.btn_select_all.Click += self.on_select_all_clicked
        self.combo_areascheme.SelectionChanged += self.on_areascheme_changed
        
        # Populate area scheme dropdown
        for area_scheme, municipality in area_schemes:
            display_text = "{} ({})".format(area_scheme.Name, municipality)
            self.combo_areascheme.Items.Add(display_text)
        
        # Select the specified area scheme
        if area_schemes:
            self.combo_areascheme.SelectedIndex = selected_scheme_index
        
        # Show area scheme selector only when multiple municipalities exist
        try:
            municipalities = set()
            for _, municipality in area_schemes:
                if municipality:
                    municipalities.add(municipality)
            visibility = System.Windows.Visibility
            if len(municipalities) > 1:
                self.combo_areascheme.Visibility = visibility.Visible
            else:
                self.combo_areascheme.Visibility = visibility.Collapsed
        except Exception:
            pass
        
        # Hide mode selection (not needed for FixAreaNames)
        try:
            if hasattr(self, 'panel_mode'):
                self.panel_mode.Visibility = System.Windows.Visibility.Collapsed
        except Exception:
            pass
        
        # Result
        self.selected_views = None
    
    def on_areascheme_changed(self, sender, args):
        """Handle area scheme selection change"""
        if self.combo_areascheme.SelectedIndex < 0:
            return
        
        # Rebuild view list for selected area scheme
        self._populate_view_list()
    
    def _populate_view_list(self):
        """Populate the list of AreaPlan views for the selected area scheme"""
        # Clear existing checkboxes
        self.panel_views.Children.Clear()
        self._view_checkboxes = {}
        
        if self.combo_areascheme.SelectedIndex < 0:
            return
        
        # Get selected area scheme
        area_scheme, municipality = self._area_schemes[self.combo_areascheme.SelectedIndex]
        
        # Get all AreaPlan views for this area scheme
        collector = DB.FilteredElementCollector(self._doc)
        all_views = collector.OfClass(DB.View).ToElements()
        
        area_plan_views = []
        for view in all_views:
            try:
                # Must be AreaPlan with matching scheme
                if not hasattr(view, 'AreaScheme'):
                    continue
                if not view.AreaScheme or view.AreaScheme.Id != area_scheme.Id:
                    continue
                
                # Check if view has placed areas
                areas = _get_areas_in_view(self._doc, view)
                if not areas:
                    continue
                
                area_plan_views.append(view)
            except:
                continue
        
        # Sort by elevation (level)
        area_plan_views.sort(key=lambda v: v.Origin.Z if hasattr(v, 'Origin') else 0, reverse=True)
        
        # Create checkboxes for each view
        if not area_plan_views:
            no_views_text = System.Windows.Controls.TextBlock()
            no_views_text.Text = "No Area Plan views with placed areas found for this area scheme."
            no_views_text.Foreground = System.Windows.Media.Brushes.Gray
            no_views_text.FontStyle = System.Windows.FontStyles.Italic
            no_views_text.Margin = System.Windows.Thickness(5)
            self.panel_views.Children.Add(no_views_text)
            return
        
        for view in area_plan_views:
            view_id_value = _get_element_id_value(view.Id)
            
            # Get level name
            level_name = "N/A"
            if hasattr(view, 'GenLevel') and view.GenLevel:
                level_name = view.GenLevel.Name
            
            # Create checkbox
            checkbox = CheckBox()
            checkbox.Content = "{} (Level: {})".format(view.Name, level_name)
            checkbox.Margin = System.Windows.Thickness(5, 3, 5, 3)
            checkbox.Tag = view_id_value
            
            # Preselect if in preselected list
            if view.Id in self._preselected_view_ids:
                checkbox.IsChecked = True
            
            self.panel_views.Children.Add(checkbox)
            self._view_checkboxes[view_id_value] = checkbox
    
    def on_select_all_clicked(self, sender, args):
        """Select all views"""
        for checkbox in self._view_checkboxes.values():
            checkbox.IsChecked = True
    
    def on_ok_clicked(self, sender, args):
        """Handle OK button click"""
        if self.combo_areascheme.SelectedIndex < 0:
            forms.alert("Please select an area scheme.", exitscript=False)
            return
        
        # Get selected views
        selected_view_ids = []
        for view_id_value, checkbox in self._view_checkboxes.items():
            if checkbox.IsChecked:
                selected_view_ids.append(view_id_value)
        
        if not selected_view_ids:
            forms.alert("Please select at least one view.", exitscript=False)
            return
        
        # Get area scheme
        area_scheme, municipality = self._area_schemes[self.combo_areascheme.SelectedIndex]
        
        # Get view elements
        selected_views = []
        for view_id_value in selected_view_ids:
            view = self._doc.GetElement(DB.ElementId(System.Int64(int(view_id_value))))
            if view:
                selected_views.append(view)
        
        self.selected_views = selected_views
        self.DialogResult = True
        self.Close()
    
    def on_cancel_clicked(self, sender, args):
        """Handle Cancel button click"""
        self.DialogResult = False
        self.Close()


# ============================================================
# UTILITY FUNCTIONS
# ============================================================

def _get_element_id_value(element_id):
    """Get integer value from ElementId - compatible with Revit 2024, 2025 and 2026+"""
    try:
        return element_id.IntegerValue
    except AttributeError:
        return int(element_id.Value)


def _get_context_element():
    """Get context element(s) from selection or active view.
    
    Returns:
        tuple: (elements, element_type) where:
               - elements is a list of View elements (for "views" type) or single ViewSheet (for "sheet" type)
               - element_type is "views", "sheet", or None
               Returns (None, None) if no valid context
    """
    try:
        doc = revit.doc
        uidoc = revit.uidoc
        
        # Get current selection
        selection = uidoc.Selection
        selected_ids = selection.GetElementIds()
        
        if selected_ids and len(selected_ids) > 0:
            # Collect all selected AreaPlan views (from viewports or project browser)
            selected_area_plan_views = []
            selected_sheets = []
            
            for elem_id in selected_ids:
                elem = doc.GetElement(elem_id)
                
                # Check if it's a viewport (view on sheet)
                if isinstance(elem, DB.Viewport):
                    view_id = elem.ViewId
                    view = doc.GetElement(view_id)
                    # Check if it's an area plan with defined municipality
                    if hasattr(view, 'AreaScheme') and view.AreaScheme:
                        if data_manager.get_municipality(view.AreaScheme):
                            selected_area_plan_views.append(view)
                
                # Check if it's a view (selected in project browser)
                elif isinstance(elem, DB.View) and not isinstance(elem, DB.ViewSheet):
                    if hasattr(elem, 'AreaScheme') and elem.AreaScheme:
                        if data_manager.get_municipality(elem.AreaScheme):
                            selected_area_plan_views.append(elem)
                
                # Check if it's a sheet
                elif isinstance(elem, DB.ViewSheet):
                    selected_sheets.append(elem)
            
            # Return selected AreaPlan views if any were found
            if selected_area_plan_views:
                return (selected_area_plan_views, "views")
            
            # Return first selected sheet if any
            if selected_sheets:
                return (selected_sheets[0], "sheet")
        
        # Priority 2: Check active view if nothing is selected
        active_view = uidoc.ActiveView
        
        # Check if active view is a sheet
        if isinstance(active_view, DB.ViewSheet):
            return (active_view, "sheet")
        
        # Check if active view is an area plan
        if hasattr(active_view, 'AreaScheme') and active_view.AreaScheme:
            if data_manager.get_municipality(active_view.AreaScheme):
                return ([active_view], "views")
        
    except Exception:
        pass  # Silently fail
    
    return (None, None)


def _get_defined_area_schemes():
    """Get all area schemes with municipality defined.
    
    Returns:
        List of (AreaScheme element, municipality) tuples
    """
    doc = revit.doc
    collector = DB.FilteredElementCollector(doc)
    area_schemes = list(collector.OfClass(DB.AreaScheme).ToElements())
    
    defined_schemes = []
    for scheme in area_schemes:
        municipality = data_manager.get_municipality(scheme)
        if municipality:
            defined_schemes.append((scheme, municipality))
    
    return defined_schemes


def _get_views_on_sheet(sheet):
    """Get AreaPlan views placed on a sheet.
    
    Args:
        sheet: ViewSheet element
        
    Returns:
        List of View elements (AreaPlan views only)
    """
    try:
        view_ids = sheet.GetAllPlacedViews()
        area_plan_views = []
        
        for view_id in view_ids:
            view = revit.doc.GetElement(view_id)
            if hasattr(view, 'AreaScheme') and view.AreaScheme:
                area_plan_views.append(view)
        
        return area_plan_views
    except:
        return []


def _get_areas_in_view(doc, areaplan_view):
    """Get all PLACED areas in the specified Area Plan view."""
    areas = []
    try:
        area_scheme = getattr(areaplan_view, "AreaScheme", None)
        gen_level = getattr(areaplan_view, "GenLevel", None)
        if area_scheme is None or gen_level is None:
            return []
        
        target_scheme_id = area_scheme.Id
        target_level_id = gen_level.Id
        
        collector = DB.FilteredElementCollector(doc)
        collector = collector.OfCategory(DB.BuiltInCategory.OST_Areas)
        collector = collector.WhereElementIsNotElementType()
        
        for area in collector:
            try:
                if getattr(area, "AreaScheme", None) is None:
                    continue
                if area.AreaScheme.Id != target_scheme_id:
                    continue
                if area.LevelId != target_level_id:
                    continue
                
                # Filter out unplaced areas (no boundaries)
                options = DB.SpatialElementBoundaryOptions()
                loops = area.GetBoundarySegments(options)
                if not loops or len(loops) == 0:
                    continue
                
                areas.append(area)
            except Exception:
                continue
    except Exception:
        return []
    
    return areas


def _load_usage_type_catalog(municipality, variant):
    """Load usage type catalog from CSV file.
    
    Args:
        municipality: Municipality name
        variant: Variant name
    
    Returns:
        dict: {usage_type_value: usage_type_name}
    """
    csv_filename = get_usage_type_csv_filename(municipality, variant)
    csv_path = os.path.join(LIB_DIR, csv_filename)
    
    catalog = {}
    
    if not os.path.exists(csv_path):
        logger.warning("Usage type CSV not found: %s", csv_path)
        return catalog
    
    try:
        with codecs.open(csv_path, 'r', encoding='utf-8-sig') as f:
            reader = csv.DictReader(f)
            for row in reader:
                usage_type = row.get('usage_type', '').strip()
                name = row.get('name', '').strip()
                if usage_type and name:
                    catalog[usage_type] = name
    except Exception as e:
        logger.error("Failed to load usage type CSV: %s", e)
    
    return catalog


def _is_area_in_group(area_elem):
    """Check if area is part of a group."""
    try:
        group_id = area_elem.GroupId
        if group_id and group_id != DB.ElementId.InvalidElementId:
            return True
    except:
        pass
    return False


# ============================================================
# VALIDATION AND FIXING
# ============================================================

class AreaValidationResult:
    """Result of area validation"""
    def __init__(self, area_elem):
        self.area_elem = area_elem
        self.area_id = _get_element_id_value(area_elem.Id)
        self.usage_type = None
        self.usage_type_prev = None
        self.name = None
        self.is_in_group = _is_area_in_group(area_elem)
        self.issues = []
        self.actions_taken = []
        self.has_issues = False


def _validate_and_fix_area(area_elem, catalog, doc):
    """Validate and fix a single area.
    
    Args:
        area_elem: Area element
        catalog: Usage type catalog {usage_type_value: usage_type_name}
        doc: Revit document
    
    Returns:
        AreaValidationResult
    """
    result = AreaValidationResult(area_elem)
    
    # Get current values
    try:
        usage_type_param = area_elem.LookupParameter("Usage Type")
        if usage_type_param:
            result.usage_type = usage_type_param.AsString()
    except:
        pass
    
    try:
        usage_type_prev_param = area_elem.LookupParameter("Usage Type Prev")
        if usage_type_prev_param:
            result.usage_type_prev = usage_type_prev_param.AsString()
    except:
        pass
    
    try:
        name_param = area_elem.LookupParameter("Name")
        if name_param:
            result.name = name_param.AsString()
    except:
        pass
    
    # Validate Usage Type
    if result.usage_type:
        if result.usage_type not in catalog:
            result.has_issues = True
            result.issues.append("Usage Type '{}' not in catalog".format(result.usage_type))
            
            # Special case: if value is "0", empty it
            if result.usage_type == "0" and not result.is_in_group:
                try:
                    usage_type_param = area_elem.LookupParameter("Usage Type")
                    if usage_type_param and not usage_type_param.IsReadOnly:
                        usage_type_param.Set("")
                        result.actions_taken.append("Cleared Usage Type (was '0')")
                except Exception as e:
                    result.actions_taken.append("Failed to clear Usage Type: {}".format(str(e)))
        else:
            # Check if name matches
            expected_name = catalog[result.usage_type]
            if result.name != expected_name:
                result.has_issues = True
                result.issues.append("Name mismatch: '{}' should be '{}'".format(
                    result.name or "(empty)", expected_name))
                
                # Try to fix if not in group
                if not result.is_in_group:
                    try:
                        name_param = area_elem.LookupParameter("Name")
                        if name_param and not name_param.IsReadOnly:
                            name_param.Set(expected_name)
                            result.actions_taken.append("Fixed Name to '{}'".format(expected_name))
                    except Exception as e:
                        result.actions_taken.append("Failed to fix Name: {}".format(str(e)))
    
    # Validate Usage Type Prev
    if result.usage_type_prev:
        if result.usage_type_prev not in catalog:
            result.has_issues = True
            result.issues.append("Usage Type Prev '{}' not in catalog".format(result.usage_type_prev))
            
            # Special case: if value is "0", empty it
            if result.usage_type_prev == "0" and not result.is_in_group:
                try:
                    usage_type_prev_param = area_elem.LookupParameter("Usage Type Prev")
                    if usage_type_prev_param and not usage_type_prev_param.IsReadOnly:
                        usage_type_prev_param.Set("")
                        result.actions_taken.append("Cleared Usage Type Prev (was '0')")
                except Exception as e:
                    result.actions_taken.append("Failed to clear Usage Type Prev: {}".format(str(e)))
        else:
            # Check if Usage Type Prev. Name matches
            expected_prev_name = catalog[result.usage_type_prev]
            
            # Get current Usage Type Prev. Name
            usage_type_prev_name = None
            try:
                usage_type_prev_name_param = area_elem.LookupParameter("Usage Type Prev. Name")
                if usage_type_prev_name_param:
                    usage_type_prev_name = usage_type_prev_name_param.AsString()
            except:
                pass
            
            if usage_type_prev_name != expected_prev_name:
                result.has_issues = True
                result.issues.append("Usage Type Prev. Name mismatch: '{}' should be '{}'".format(
                    usage_type_prev_name or "(empty)", expected_prev_name))
                
                # Try to fix if not in group
                if not result.is_in_group:
                    try:
                        usage_type_prev_name_param = area_elem.LookupParameter("Usage Type Prev. Name")
                        if usage_type_prev_name_param and not usage_type_prev_name_param.IsReadOnly:
                            usage_type_prev_name_param.Set(expected_prev_name)
                            result.actions_taken.append("Fixed Usage Type Prev. Name to '{}'".format(expected_prev_name))
                    except Exception as e:
                        result.actions_taken.append("Failed to fix Usage Type Prev. Name: {}".format(str(e)))
    
    # Special case: if in group, note that we can't fix
    if result.is_in_group and result.has_issues:
        result.actions_taken.append("Cannot fix - area is in a group")
    
    return result


def _process_view(doc, view, catalog, output):
    """Process all areas in a view.
    
    Args:
        doc: Revit document
        view: AreaPlan view
        catalog: Usage type catalog
        output: Script output for linkify
    
    Returns:
        tuple: (total_areas, problematic_areas, fixed_count)
    """
    areas = _get_areas_in_view(doc, view)
    
    if not areas:
        return 0, 0, 0
    
    problematic_results = []
    fixed_count = 0
    
    with revit.Transaction("Fix Area Names"):
        for area in areas:
            result = _validate_and_fix_area(area, catalog, doc)
            
            if result.has_issues:
                problematic_results.append(result)
                if result.actions_taken and not result.is_in_group:
                    fixed_count += 1
    
    # Report problematic areas - compact format
    if problematic_results:
        # Add separator line before view
        print("-" * 80)
        
        # View header - single line with linkified view
        view_link = output.linkify(view.Id)
        print("{} {}".format(view_link, view.Name))
        
        # Each area on one line, indented with visible character
        for result in problematic_results:
            try:
                area_link = output.linkify(result.area_elem.Id)
                
                # Compact format: ID | UT | UTPrev | Issues | Actions
                parts = []
                parts.append(area_link)
                parts.append("UT:{}".format(result.usage_type or ""))
                parts.append("UTPrev:{}".format(result.usage_type_prev or ""))
                parts.append("Issues:[{}]".format("; ".join(result.issues)))
                
                if result.actions_taken:
                    parts.append("Actions:[{}]".format("; ".join(result.actions_taken)))
                else:
                    parts.append("Actions:[None]")
                
                line_text = u"  \u2022 " + " | ".join(parts)
                
                # Check if this area needs manual fixing (in group or no successful actions)
                needs_manual_fix = result.is_in_group or not result.actions_taken or \
                                   any("Cannot fix" in action or "Failed" in action for action in result.actions_taken)
                
                if needs_manual_fix:
                    # Color the line red for manual fixes
                    output.print_html('<span style="color: red;">{}</span>'.format(line_text))
                else:
                    print(line_text)
            except Exception as e:
                logger.debug("Failed to process result for area %s: %s", result.area_id, e)
    
    return len(areas), len(problematic_results), fixed_count


# ============================================================
# MAIN FUNCTION
# ============================================================

def fix_area_names():
    """
    Validate and fix area names based on usage types.
    
    This tool will:
    1. Show dialog to select AreaPlan views to process
    2. Load usage type catalog for the municipality
    3. Validate each area's Usage Type and Name
    4. Fix mismatches where possible (not in groups)
    5. Report all problematic areas
    """
    doc = revit.doc
    output = script.get_output()
    
    # Set output window to be wider for long lines
    output.set_width(1400)
    
    # Get defined area schemes
    defined_schemes = _get_defined_area_schemes()
    if not defined_schemes:
        forms.alert(
            "No area schemes with municipality defined.\n\n"
            "Please define an area scheme using the Calculation Setup tool first.",
            exitscript=True
        )
    
    # Get context element to determine default area scheme
    context_elem, context_type = _get_context_element()
    
    # Determine which area scheme to select by default
    selected_scheme_index = 0
    preselected_view_ids = set()
    
    if context_elem:
        if context_type == "views":
            # Multiple AreaPlan views selected (or single active view)
            first_view = context_elem[0]
            view_scheme = first_view.AreaScheme
            for i, (scheme, _) in enumerate(defined_schemes):
                if scheme.Id == view_scheme.Id:
                    selected_scheme_index = i
                    break
            # Preselect all selected views
            for view in context_elem:
                preselected_view_ids.add(view.Id)
        
        elif context_type == "sheet":
            # Get area plans on this sheet
            views_on_sheet = _get_views_on_sheet(context_elem)
            if views_on_sheet:
                first_view_scheme = views_on_sheet[0].AreaScheme
                for i, (scheme, _) in enumerate(defined_schemes):
                    if scheme.Id == first_view_scheme.Id:
                        selected_scheme_index = i
                        break
                # Preselect all views on sheet
                for view in views_on_sheet:
                    preselected_view_ids.add(view.Id)
    
    # Show view selection dialog
    dialog = ViewSelectionDialog(
        defined_schemes,
        selected_scheme_index,
        preselected_view_ids
    )
    
    if not dialog.ShowDialog():
        return  # User cancelled
    
    selected_views = dialog.selected_views
    if not selected_views:
        return
    
    # Get municipality and variant from first view
    first_view = selected_views[0]
    municipality, variant = data_manager.get_municipality_from_view(doc, first_view)
    
    # Load usage type catalog
    catalog = _load_usage_type_catalog(municipality, variant)
    if not catalog:
        forms.alert(
            "Failed to load usage type catalog for {} ({})".format(municipality, variant),
            exitscript=True
        )
    
    print("Municipality: {} | Variant: {} | Catalog: {} types\n".format(
        municipality, variant, len(catalog)))
    
    # Sort views by elevation (lowest to highest)
    selected_views.sort(key=lambda v: v.Origin.Z if hasattr(v, 'Origin') else 0)
    
    # Process each selected view
    total_areas = 0
    total_problematic = 0
    total_fixed = 0
    
    for view in selected_views:
        area_scheme = getattr(view, "AreaScheme", None)
        if area_scheme is None:
            continue
        
        view_total, view_problematic, view_fixed = _process_view(doc, view, catalog, output)
        total_areas += view_total
        total_problematic += view_problematic
        total_fixed += view_fixed
    
    # Final summary - single line
    print("\nTotal: {} areas | {} problematic | {} fixed | {} need manual fix".format(
        total_areas, total_problematic, total_fixed, total_problematic - total_fixed))


if __name__ == '__main__':
    fix_area_names()
