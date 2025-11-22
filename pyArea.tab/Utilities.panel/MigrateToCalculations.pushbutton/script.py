# -*- coding: utf-8 -*-
"""Migrate old pyArea projects to Calculation hierarchy (Schema v2.0)

Converts projects from Schema v1.0 (Sheet-level fields) to v2.0 (Calculation-based).
Groups sheets with identical metadata into shared Calculations.
"""

__title__ = "Migrate to\nCalculations"
__doc__ = """Migrate old pyArea data to Calculation hierarchy.

Converts projects from Schema v1.0 to v2.0:
- Groups sheets with identical metadata
- Creates Calculations on AreaSchemes
- Updates sheets to reference Calculations
- Sets SchemaVersion to "2.0"

Safe to run multiple times - skips if already migrated.
"""

import sys
import os

# Add lib to path
script_dir = os.path.dirname(__file__)
lib_path = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(script_dir))), "lib")
if lib_path not in sys.path:
    sys.path.insert(0, lib_path)

from pyrevit import revit, DB, forms
import data_manager
from schemas import municipality_schemas


# Helper function for Revit 2026 compatibility - ElementId.IntegerValue was removed
def get_element_id_value(element_id):
    """Get integer value from ElementId - compatible with Revit 2024, 2025 and 2026+"""
    try:
        # Revit 2024-2025
        return element_id.IntegerValue
    except AttributeError:
        # Revit 2026+ - IntegerValue removed, use Value instead
        return int(element_id.Value)


def get_sheet_metadata_key(sheet_data, municipality):
    """Generate a unique key for sheet metadata grouping.
    
    Args:
        sheet_data: Sheet data dictionary
        municipality: Municipality name
        
    Returns:
        str: Unique key for grouping (e.g., "PROJECT=A|ELEVATION=100|...")
    """
    # Get fields that should be moved to Calculation
    calc_fields = municipality_schemas.get_fields_for_element_type("Calculation", municipality)
    
    # Build key from field values
    key_parts = []
    for field_name in sorted(calc_fields.keys()):
        if field_name in ["CalculationGuid", "Name", "AreaPlanDefaults", "AreaDefaults"]:
            continue  # Skip these
        
        value = sheet_data.get(field_name, "")
        key_parts.append("{}={}".format(field_name, value))
    
    return "|".join(key_parts) if key_parts else "EMPTY"


def migrate_project(doc):
    """Migrate project from Schema v1.0 to v2.0.
    
    Args:
        doc: Revit document
        
    Returns:
        tuple: (success, message)
    """
    # Check current schema version
    current_version = data_manager.get_schema_version(doc)
    if current_version == "2.0":
        return True, "Project already migrated to Schema v2.0"
    
    # Collect all sheets
    collector = DB.FilteredElementCollector(doc).OfClass(DB.ViewSheet)
    all_sheets = list(collector)
    
    if not all_sheets:
        return False, "No sheets found in project"
    
    # Find sheets with pyArea data
    sheets_to_migrate = []
    for sheet in all_sheets:
        sheet_data = data_manager.get_sheet_data(sheet)
        if sheet_data and "AreaSchemeId" in sheet_data:
            sheets_to_migrate.append(sheet)
    
    if not sheets_to_migrate:
        return False, "No pyArea sheets found to migrate"
    
    # Group sheets by AreaScheme
    sheets_by_areascheme = {}
    
    for sheet in sheets_to_migrate:
        sheet_data = data_manager.get_sheet_data(sheet)
        area_scheme_id_str = sheet_data.get("AreaSchemeId")
        
        if not area_scheme_id_str:
            continue
        
        # Get AreaScheme
        area_scheme = data_manager.get_area_scheme_by_id(doc, area_scheme_id_str)
        if not area_scheme:
            print("WARNING: Sheet {} references non-existent AreaScheme {}".format(
                sheet.SheetNumber, area_scheme_id_str))
            continue
        
        area_scheme_key = str(get_element_id_value(area_scheme.Id))
        
        if area_scheme_key not in sheets_by_areascheme:
            sheets_by_areascheme[area_scheme_key] = {
                "area_scheme": area_scheme,
                "sheets": []
            }
        
        sheets_by_areascheme[area_scheme_key]["sheets"].append(sheet)
    
    if not sheets_by_areascheme:
        return False, "No valid sheets found for migration"
    
    # Process each AreaScheme
    total_calculations = 0
    total_sheets = 0
    
    with revit.Transaction("Migrate to Calculations"):
        for area_scheme_data in sheets_by_areascheme.values():
            area_scheme = area_scheme_data["area_scheme"]
            sheets = area_scheme_data["sheets"]
            
            # Get municipality
            municipality = data_manager.get_municipality(area_scheme)
            if not municipality:
                print("WARNING: AreaScheme {} has no municipality set, skipping".format(
                    area_scheme.Name))
                continue
            
            # Group sheets by metadata
            metadata_groups = {}
            
            for sheet in sheets:
                sheet_data = data_manager.get_sheet_data(sheet)
                metadata_key = get_sheet_metadata_key(sheet_data, municipality)
                
                if metadata_key not in metadata_groups:
                    metadata_groups[metadata_key] = {
                        "sheets": [],
                        "metadata": {}
                    }
                
                metadata_groups[metadata_key]["sheets"].append(sheet)
                
                # Store metadata (use first sheet's data as representative)
                if not metadata_groups[metadata_key]["metadata"]:
                    # Extract Calculation fields from sheet data
                    calc_fields = municipality_schemas.get_fields_for_element_type(
                        "Calculation", municipality)
                    
                    for field_name in calc_fields.keys():
                        if field_name in ["CalculationGuid", "Name", "AreaPlanDefaults", "AreaDefaults"]:
                            continue
                        if field_name in sheet_data:
                            metadata_groups[metadata_key]["metadata"][field_name] = sheet_data[field_name]
            
            # Create Calculation for each metadata group
            for group_index, (metadata_key, group_data) in enumerate(metadata_groups.items()):
                # Generate Calculation
                calculation_guid = data_manager.generate_calculation_guid()
                
                # Create calculation name
                sheet_numbers = [s.SheetNumber for s in group_data["sheets"]]
                if len(sheet_numbers) == 1:
                    calc_name = "Calculation {}".format(sheet_numbers[0])
                else:
                    calc_name = "Calculation {}-{}".format(
                        sheet_numbers[0], sheet_numbers[-1])
                
                # Build calculation data
                calculation_data = {
                    "Name": calc_name
                }
                
                # Add metadata fields
                calculation_data.update(group_data["metadata"])
                
                # Ensure required fields without defaults are populated
                calc_fields = municipality_schemas.get_fields_for_element_type(
                    "Calculation", municipality)
                
                for field_name, field_def in calc_fields.items():
                    # Skip special fields
                    if field_name in ["CalculationGuid", "Name", "AreaPlanDefaults", "AreaDefaults"]:
                        continue
                    
                    # If required field is missing and has no default, provide empty string
                    if field_def.get("required") and field_name not in calculation_data:
                        if "default" not in field_def and "placeholders" not in field_def:
                            calculation_data[field_name] = ""
                            print("INFO: Setting empty default for missing field '{}' in Calculation '{}'".format(
                                field_name, calc_name))
                
                # Set empty defaults
                calculation_data["AreaPlanDefaults"] = {}
                calculation_data["AreaDefaults"] = {}
                
                # Save Calculation to AreaScheme
                success, errors = data_manager.set_calculation(
                    area_scheme, calculation_guid, calculation_data, municipality)
                
                if not success:
                    print("ERROR creating Calculation: {}".format("; ".join(errors)))
                    continue
                
                total_calculations += 1
                
                # Update all sheets in this group (only store CalculationGuid)
                for sheet in group_data["sheets"]:
                    if data_manager.set_sheet_data(sheet, calculation_guid):
                        total_sheets += 1
                    else:
                        print("ERROR updating sheet {}".format(sheet.SheetNumber))
        
        # Set schema version
        data_manager.set_schema_version(doc, "2.0")
    
    message = "Migration complete!\n\n"
    message += "Created {} Calculations\n".format(total_calculations)
    message += "Updated {} Sheets\n".format(total_sheets)
    message += "Schema version set to 2.0"
    
    return True, message


def main():
    """Main entry point."""
    doc = revit.doc
    
    if not doc:
        forms.alert("No active document", exitscript=True)
    
    # Confirm migration
    current_version = data_manager.get_schema_version(doc)
    
    if current_version == "2.0":
        forms.alert("Project already migrated to Schema v2.0", exitscript=True)
    
    message = "Migrate project to Calculation hierarchy (Schema v2.0)?\n\n"
    message += "This will:\n"
    message += "- Group sheets with identical metadata\n"
    message += "- Create Calculations on AreaSchemes\n"
    message += "- Update all sheets to reference Calculations\n\n"
    message += "This operation can be undone with Ctrl+Z.\n\n"
    message += "Current schema version: {}".format(current_version or "1.0 (implicit)")
    
    if not forms.alert(message, yes=True, no=True):
        return
    
    # Perform migration
    success, result_message = migrate_project(doc)
    
    if success:
        forms.alert(result_message, title="Migration Successful")
    else:
        forms.alert(result_message, title="Migration Failed")


if __name__ == "__main__":
    main()
