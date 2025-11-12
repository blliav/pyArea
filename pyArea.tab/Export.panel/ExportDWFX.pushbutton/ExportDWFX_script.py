# -*- coding: utf-8 -*-
"""DWFX Exporter with Configurable Settings
Exports each sheet as a separate DWFX file with optional opaque white removal.
"""

__title__ = "Export\nDWFX"
__author__ = "pyArea"

import os
import sys
import shutil
import tempfile

# Add lib to path
script_dir = os.path.dirname(__file__)
lib_path = os.path.join(script_dir, "..", "..", "lib")
if lib_path not in sys.path:
    sys.path.insert(0, lib_path)

from pyrevit import revit, DB, forms
from data_manager import get_preferences
import export_utils

doc = revit.doc
uidoc = revit.uidoc


# ============================================================
# SHEET SELECTION
# ============================================================

def get_selected_sheets():
    """Get sheets from project browser selection or active view.
    
    Returns:
        list: List of DB.ViewSheet elements, or None if no valid selection
    """
    try:
        # Check if active view is a sheet
        active_view = doc.ActiveView
        if isinstance(active_view, DB.ViewSheet):
            print("Using active sheet: {}".format(active_view.SheetNumber))
            return [active_view]
        
        # Try to get selection from project browser
        selection = uidoc.Selection
        selected_ids = selection.GetElementIds()
        
        if selected_ids and len(selected_ids) > 0:
            sheets = []
            for elem_id in selected_ids:
                element = doc.GetElement(elem_id)
                if isinstance(element, DB.ViewSheet):
                    sheets.append(element)
            
            if len(sheets) > 0:
                print("Found {} selected sheets".format(len(sheets)))
                return sheets
        
        # No valid selection - ask user to select sheets
        print("No sheets selected. Please select sheets in project browser or open a sheet.")
        return None
        
    except Exception as e:
        print("Error getting selected sheets: {}".format(e))
        return None


# ============================================================
# VALIDATION
# ============================================================

def validate_sheets_uniform_areascheme(sheets):
    """
    Validates all sheets belong to same AreaScheme.
    Raises ValueError if validation fails.
    
    Returns:
        str: Uniform AreaScheme ID
    """
    from data_manager import get_sheet_data
    
    schemes_found = {}
    missing_scheme_sheets = []
    
    for sheet in sheets:
        sheet_data = get_sheet_data(sheet)
        if not sheet_data:
            missing_scheme_sheets.append(sheet.SheetNumber)
            continue
        
        area_scheme_id = sheet_data.get("AreaSchemeId")
        if not area_scheme_id:
            missing_scheme_sheets.append(sheet.SheetNumber)
            continue
        
        if area_scheme_id not in schemes_found:
            schemes_found[area_scheme_id] = []
        schemes_found[area_scheme_id].append(sheet.SheetNumber)
    
    # Error if sheets missing AreaSchemeId
    if missing_scheme_sheets:
        error_msg = "ERROR: Sheets without AreaScheme detected!\n\n"
        error_msg += "The following sheets have no AreaSchemeId:\n"
        for num in missing_scheme_sheets:
            error_msg += "  - Sheet {}\n".format(num)
        error_msg += "\nAll sheets must belong to an AreaScheme for DWFX export."
        raise ValueError(error_msg)
    
    # Error if multiple schemes
    if len(schemes_found) > 1:
        error_msg = "ERROR: Multiple AreaSchemes detected!\n\n"
        error_msg += "All selected sheets must belong to the same AreaScheme.\n\n"
        for scheme_id, sheet_nums in schemes_found.items():
            scheme_elem = doc.GetElement(DB.ElementId(int(scheme_id)))
            scheme_name = scheme_elem.Name if scheme_elem else "Unknown"
            error_msg += "\n  AreaScheme '{}' (ID: {}):\n".format(scheme_name, scheme_id)
            for num in sheet_nums:
                error_msg += "    - Sheet {}\n".format(num)
        error_msg += "\nPlease select sheets from the same AreaScheme only."
        raise ValueError(error_msg)
    
    # Error if no schemes found
    if len(schemes_found) == 0:
        error_msg = "No AreaScheme found in selected sheets.\n\n"
        error_msg += "All sheets must belong to an AreaScheme."
        raise ValueError(error_msg)
    
    # Return uniform scheme ID
    uniform_scheme_id = list(schemes_found.keys())[0]
    scheme_elem = doc.GetElement(DB.ElementId(int(uniform_scheme_id)))
    scheme_name = scheme_elem.Name if scheme_elem else "Unknown"
    print("AreaScheme validation passed: {}".format(scheme_name))
    return uniform_scheme_id


# ============================================================
# MAIN EXPORT LOGIC
# ============================================================

def main():
    try:
        print("="*60)
        print("DWFX EXPORT")
        print("="*60)
        
        # 1. Load preferences
        print("\nLoading preferences...")
        preferences = get_preferences()
        print("  Export folder: {}".format(preferences["ExportFolder"]))
        print("  Element Data: {}".format(preferences["DWFX_ExportElementData"]))
        print("  Quality: {}".format(preferences["DWFX_Quality"]))
        print("  Remove opaque white: {}".format(preferences["DWFX_RemoveOpaqueWhite"]))
        
        # 2. Get sheets (active or selected)
        print("\nGetting sheets...")
        sheets = get_selected_sheets()
        if not sheets or len(sheets) == 0:
            forms.alert(
                "No sheets to export. Please select sheets or open a sheet view.",
                title="No Sheets Selected"
            )
            return
        
        print("Selected {} sheet(s)".format(len(sheets)))
        for sheet in sheets:
            print("  - {} - {}".format(sheet.SheetNumber, sheet.Name))
        
        # 3. Validate uniform AreaScheme
        print("\nValidating sheets...")
        area_scheme_id = validate_sheets_uniform_areascheme(sheets)
        
        # 4. Get export folder
        export_folder = export_utils.get_export_folder_path(preferences["ExportFolder"])
        if not os.path.exists(export_folder):
            os.makedirs(export_folder)
            print("Created export folder: {}".format(export_folder))
        
        # 5. Setup DWFX options
        print("\nConfiguring DWFX export options...")
        dwfx_options = DB.DWFXExportOptions()
        dwfx_options.ExportingAreas = False  # Always false per requirements
        dwfx_options.ExportObjectData = preferences["DWFX_ExportElementData"]  # Boolean property
        
        # Try to set quality using enum (Revit API version dependent)
        try:
            quality_map = {
                "Low": DB.DWFImageQuality.Low,
                "Medium": DB.DWFImageQuality.Medium,
                "High": DB.DWFImageQuality.High
            }
            if preferences["DWFX_Quality"] in quality_map:
                dwfx_options.ImageQuality = quality_map[preferences["DWFX_Quality"]]
                print("  Quality: {}".format(preferences["DWFX_Quality"]))
        except AttributeError:
            # ImageQuality enum not available in this Revit version
            print("  Quality: Skipped (not supported in this Revit version)")
        
        print("  ExportingAreas: False")
        print("  ExportObjectData: {}".format("Yes" if preferences["DWFX_ExportElementData"] else "No"))
        
        # 6. Determine if we need temp directory for white removal
        use_temp = preferences["DWFX_RemoveOpaqueWhite"]
        temp_dir = None
        
        if use_temp:
            temp_dir = tempfile.mkdtemp(prefix='dwfx_export_')
            print("  Using temp directory for white removal: {}".format(temp_dir))
        
        # 7. Export each sheet
        print("\n" + "="*60)
        print("EXPORTING SHEETS")
        print("="*60)
        
        exported_count = 0
        failed_sheets = []
        
        for sheet in sheets:
            try:
                # Generate filename
                filename = export_utils.generate_dwfx_filename(doc.Title, sheet.SheetNumber)
                
                # Determine export path (temp or final)
                if use_temp:
                    export_path = temp_dir
                    temp_filepath = os.path.join(temp_dir, filename + ".dwfx")
                else:
                    export_path = export_folder
                
                final_filepath = os.path.join(export_folder, filename + ".dwfx")
                
                # Create ViewSet with single sheet
                view_set = DB.ViewSet()
                view_set.Insert(sheet)
                
                # Export to temp or final location (requires transaction)
                t = DB.Transaction(doc, "Export DWFX")
                t.Start()
                try:
                    doc.Export(export_path, filename, view_set, dwfx_options)
                    t.Commit()
                except:
                    t.RollBack()
                    raise
                
                # Post-process if white removal enabled
                if use_temp:
                    print("  Exported {} to temp...".format(sheet.SheetNumber))
                    
                    # Apply white removal
                    changes, success = export_utils.fix_dwfx_file(temp_filepath)
                    
                    if success:
                        if changes > 0:
                            print("    Removed {} white fills".format(changes))
                        else:
                            print("    No white fills found")
                        
                        # Move from temp to final location
                        if os.path.exists(final_filepath):
                            os.remove(final_filepath)
                        shutil.move(temp_filepath, final_filepath)
                        print("  \u2713 Completed: {} \u2192 {}".format(sheet.SheetNumber, filename + ".dwfx"))
                    else:
                        print("    WARNING: White removal failed, using original export")
                        if os.path.exists(final_filepath):
                            os.remove(final_filepath)
                        shutil.move(temp_filepath, final_filepath)
                        print("  \u2713 Completed: {} \u2192 {}".format(sheet.SheetNumber, filename + ".dwfx"))
                else:
                    print("  \u2713 Exported: {} \u2192 {}".format(sheet.SheetNumber, filename + ".dwfx"))
                
                exported_count += 1
                
            except Exception as e:
                print("  \u2717 Failed: {} - {}".format(sheet.SheetNumber, str(e)))
                failed_sheets.append((sheet.SheetNumber, str(e)))
        
        # 8. Cleanup temp directory
        if temp_dir and os.path.exists(temp_dir):
            try:
                shutil.rmtree(temp_dir, ignore_errors=True)
            except:
                pass
        
        # 9. Report results
        print("\n" + "="*60)
        print("EXPORT COMPLETE")
        print("="*60)
        print("Exported: {} of {} sheets".format(exported_count, len(sheets)))
        print("Export folder: {}".format(export_folder))
        
        if failed_sheets:
            print("\nFailed sheets:")
            for num, error in failed_sheets:
                print("  - Sheet {}: {}".format(num, error))
            forms.alert(
                "Export completed with errors.\n\n"
                "Successfully exported: {} of {}\n\n"
                "Check console for details.".format(exported_count, len(sheets)),
                title="Export Completed with Errors"
            )
        else:
            forms.alert(
                "Successfully exported {} sheet(s) to:\n\n{}".format(
                    exported_count, export_folder
                ),
                title="Export Complete"
            )
        
    except ValueError as e:
        # Validation errors
        print("\n" + "="*60)
        print("VALIDATION ERROR")
        print("="*60)
        print(str(e))
        forms.alert(str(e), title="Validation Error")
        
    except Exception as e:
        # Unexpected errors
        print("\n" + "="*60)
        print("ERROR")
        print("="*60)
        print(str(e))
        import traceback
        traceback.print_exc()
        forms.alert(
            "Unexpected error during export:\n\n{}".format(str(e)),
            title="Error"
        )


if __name__ == "__main__":
    main()
