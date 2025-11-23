# -*- coding: utf-8 -*-
"""DWFx Exporter with Configurable Settings
Exports each sheet as a separate DWFx file with optional opaque white removal.
"""

__title__ = "Export\nDWFx"
__author__ = "pyArea"

import os
import sys
import shutil
import tempfile
import subprocess
import datetime
import System

# Add lib to path
script_dir = os.path.dirname(__file__)
lib_path = os.path.join(script_dir, "..", "..", "lib")
if lib_path not in sys.path:
    sys.path.insert(0, lib_path)

from pyrevit import revit, DB, forms
from pyrevit import script
from data_manager import get_preferences
import export_utils
from python_utils import find_python_executable

doc = revit.doc
uidoc = revit.uidoc
output = script.get_output()


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
    Validates sheets and categorizes them by AreaScheme.
    Allows sheets without AreaScheme if user approves.
    Raises ValueError if validation fails.
    
    Returns:
        tuple: (uniform_scheme_id or None, sheets_with_scheme, sheets_without_scheme)
    """
    from data_manager import get_calculation_from_sheet
    
    schemes_found = {}
    missing_scheme_sheets = []
    sheets_with_scheme = []
    sheets_without_scheme = []
    
    for sheet in sheets:
        # Get AreaScheme from sheet's viewport (v2.0 approach)
        area_scheme, calculation_data = get_calculation_from_sheet(doc, sheet)
        
        if not area_scheme:
            missing_scheme_sheets.append(sheet.SheetNumber)
            sheets_without_scheme.append(sheet)
            continue
        
        area_scheme_id = str(area_scheme.Id.Value)
        
        if area_scheme_id not in schemes_found:
            schemes_found[area_scheme_id] = []
        schemes_found[area_scheme_id].append(sheet.SheetNumber)
        sheets_with_scheme.append(sheet)
    
    # If sheets missing AreaScheme - ask for user approval
    if missing_scheme_sheets:
        warning_msg = "WARNING: Sheets undefined by pyArea detected!\n\n"
        warning_msg += "The following sheets are not associated with any AreaScheme/Calculation:\n"
        for num in missing_scheme_sheets:
            warning_msg += "  - Sheet {}\n".format(num)
        warning_msg += "\nDo you want to proceed with exporting these sheets?"
        
        user_approved = forms.alert(
            warning_msg,
            title="Sheets Undefined by pyArea",
            yes=True,
            no=True
        )
        
        if not user_approved:
            raise ValueError("Export cancelled by user.")
        
        print("User approved export of {} sheet(s) undefined by pyArea".format(len(missing_scheme_sheets)))
    
    # Error if multiple schemes (only among sheets WITH schemes)
    if len(schemes_found) > 1:
        error_msg = "ERROR: Multiple AreaSchemes detected!\n\n"
        error_msg += "All selected sheets with AreaSchemes must belong to the same AreaScheme.\n\n"
        for scheme_id, sheet_nums in schemes_found.items():
            scheme_elem = doc.GetElement(DB.ElementId(System.Int64(int(scheme_id))))
            scheme_name = scheme_elem.Name if scheme_elem else "Unknown"
            error_msg += "\n  AreaScheme '{}' (ID: {}):\n".format(scheme_name, scheme_id)
            for num in sheet_nums:
                error_msg += "    - Sheet {}\n".format(num)
        error_msg += "\nPlease select sheets from the same AreaScheme only."
        raise ValueError(error_msg)
    
    # Return uniform scheme ID (or None if all sheets lack schemes)
    uniform_scheme_id = None
    if len(schemes_found) > 0:
        uniform_scheme_id = list(schemes_found.keys())[0]
        scheme_elem = doc.GetElement(DB.ElementId(System.Int64(int(uniform_scheme_id))))
        scheme_name = scheme_elem.Name if scheme_elem else "Unknown"
        print("AreaScheme validation passed: {}".format(scheme_name))
    else:
        print("No AreaScheme found - proceeding with basic export")
    
    return uniform_scheme_id, sheets_with_scheme, sheets_without_scheme


# ============================================================
# MAIN EXPORT LOGIC
# ============================================================

def main():
    try:
        print("="*60)
        print("DWFx EXPORT")
        print("="*60)
        
        # 1. Load preferences
        print("\nLoading preferences...")
        preferences = get_preferences()
        print("  Export folder: {}".format(preferences["ExportFolder"]))
        print("  Element Data: {}".format(preferences["DWFx_ExportElementData"]))
        print("  Quality: {}".format(preferences["DWFx_Quality"]))
        print("  Remove opaque white: {}".format(preferences["DWFx_RemoveOpaqueWhite"]))
        
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
        area_scheme_id, sheets_with_scheme, sheets_without_scheme = validate_sheets_uniform_areascheme(sheets)
        
        if sheets_without_scheme:
            print("  {} sheet(s) without AreaScheme will be exported".format(len(sheets_without_scheme)))
        if sheets_with_scheme:
            print("  {} sheet(s) with AreaScheme will be exported".format(len(sheets_with_scheme)))
        
        # 4. Get export folder
        export_folder = export_utils.get_export_folder_path(preferences["ExportFolder"])
        if not os.path.exists(export_folder):
            os.makedirs(export_folder)
            print("Created export folder: {}".format(export_folder))
        
        # 5. Setup DWFx options
        print("\nConfiguring DWFx export options...")
        dwfx_options = DB.DWFXExportOptions()
        dwfx_options.ExportingAreas = False  # Always false per requirements
        dwfx_options.ExportObjectData = preferences["DWFx_ExportElementData"]  # Boolean property
        
        # Try to set quality using enum (Revit API version dependent)
        try:
            quality_map = {
                "Low": DB.DWFImageQuality.Low,
                "Medium": DB.DWFImageQuality.Medium,
                "High": DB.DWFImageQuality.High
            }
            if preferences["DWFx_Quality"] in quality_map:
                dwfx_options.ImageQuality = quality_map[preferences["DWFx_Quality"]]
                print("  Quality: {}".format(preferences["DWFx_Quality"]))
        except AttributeError:
            # ImageQuality enum not available in this Revit version
            print("  Quality: Skipped (not supported in this Revit version)")
        
        print("  ExportingAreas: False")
        print("  ExportObjectData: {}".format("Yes" if preferences["DWFx_ExportElementData"] else "No"))
        
        # 6. Setup background processing if white removal enabled
        use_background_processing = preferences["DWFx_RemoveOpaqueWhite"]
        file_list_path = None
        temp_export_folder = None
        temp_files = []  # Track temp files for background processing
        
        if use_background_processing:
            # Create temp directory for export
            temp_export_folder = tempfile.mkdtemp(prefix='dwfx_export_')
            
            # Create file list for background processor
            timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            file_list_path = os.path.join(temp_export_folder, "dwfx_file_list_{}.txt".format(timestamp))
            print("  Background processing enabled")
            print("  Temp export folder: {}".format(temp_export_folder))
            print("  File list: {}".format(file_list_path))
        
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
                
                # Determine export location
                if use_background_processing:
                    # Export to temp folder
                    export_target_folder = temp_export_folder
                    temp_filepath = os.path.join(temp_export_folder, filename + ".dwfx")
                else:
                    # Export directly to final folder
                    export_target_folder = export_folder
                
                # Create ViewSet with single sheet
                view_set = DB.ViewSet()
                view_set.Insert(sheet)
                
                # Print status before export starts
                print("Exporting sheet {} - {}...".format(sheet.SheetNumber, sheet.Name))
                
                # Export (requires transaction)
                t = DB.Transaction(doc, "Export DWFx")
                t.Start()
                try:
                    doc.Export(export_target_folder, filename, view_set, dwfx_options)
                    t.Commit()
                    
                    # Verify file was actually created (ESC cancellation doesn't throw exception)
                    expected_file = os.path.join(export_target_folder, filename + ".dwfx")
                    if not os.path.exists(expected_file):
                        raise Exception("Export was cancelled or file was not created")
                    
                    # Only report success and count if file exists
                    if use_background_processing:
                        output.print_md("✅ Exported to temp: **{}** → **{}**".format(sheet.SheetNumber, filename + ".dwfx"))
                        # Track temp file for background processing
                        temp_files.append(temp_filepath)
                    else:
                        output.print_md("✅ Exported: **{}** → **{}**".format(sheet.SheetNumber, filename + ".dwfx"))
                    
                    exported_count += 1
                    
                except Exception as export_error:
                    # Only rollback if transaction is still active (ESC auto-rolls back)
                    if t.GetStatus() == DB.TransactionStatus.Started:
                        t.RollBack()
                    raise export_error
                
            except Exception as e:
                output.print_md("❌ Failed: **{}** - {}".format(sheet.SheetNumber, str(e)))
                failed_sheets.append((sheet.SheetNumber, str(e)))
        
        # 8. Launch background processor if white removal enabled
        if use_background_processing and temp_files:
            try:
                # Write file list (temp file paths)
                with open(file_list_path, 'w') as f:
                    for temp_path in temp_files:
                        f.write(temp_path + "\n")
                
                # Find Python interpreter and postprocessor script
                processor_script = os.path.join(lib_path, "DWFx_postprocessor.py")
                
                if not os.path.exists(processor_script):
                    print("\nWARNING: Background processor not found")
                    print("Files will remain in temp folder: {}".format(temp_export_folder))
                else:
                    # Launch external Python process (detached)
                    python_exe = find_python_executable(prefer_pythonw=True)
                    
                    # Start process detached (no wait)
                    if os.name == 'nt':  # Windows
                        subprocess.Popen(
                            [python_exe, processor_script, file_list_path, export_folder],
                            creationflags=subprocess.CREATE_NO_WINDOW if hasattr(subprocess, 'CREATE_NO_WINDOW') else 0,
                            close_fds=True
                        )
                    else:  # Unix-like
                        subprocess.Popen(
                            [python_exe, processor_script, file_list_path, export_folder],
                            close_fds=True,
                            start_new_session=True
                        )
                    
                    print("\nBackground processing started!")
                    print("Processing {} file(s) to remove white backgrounds...".format(len(temp_files)))
                    print("Files will be moved to: {}".format(export_folder))
                    print("\nYou can continue working in Revit.")
                    
            except Exception as e:
                print("\nWARNING: Failed to start background processor: {}".format(str(e)))
                print("Files will remain in temp folder: {}".format(temp_export_folder))
        
        # 9. Report results
        print("\n" + "="*60)
        print("EXPORT COMPLETE")
        print("="*60)
        print("Exported: {} of {} sheets".format(exported_count, len(sheets)))
        
        print("Export folder: {}".format(export_folder))
        
        if use_background_processing and temp_files:
            print("\nNote: Background processing active - files will move to final location when complete.")
        
        if failed_sheets:
            print("\nFailed sheets:")
            for num, error in failed_sheets:
                print("  - Sheet {}: {}".format(num, error))
        
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
