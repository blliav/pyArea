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
import subprocess
import datetime
import System

# Add lib to path
script_dir = os.path.dirname(__file__)
lib_path = os.path.join(script_dir, "..", "..", "lib")
if lib_path not in sys.path:
    sys.path.insert(0, lib_path)

from pyrevit import revit, DB, forms
from data_manager import get_preferences
import export_utils
from python_utils import find_python_executable

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
            scheme_elem = doc.GetElement(DB.ElementId(System.Int64(int(scheme_id))))
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
    scheme_elem = doc.GetElement(DB.ElementId(System.Int64(int(uniform_scheme_id))))
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
        
        # 6. Setup background processing if white removal enabled
        use_background_processing = preferences["DWFX_RemoveOpaqueWhite"]
        file_list_path = None
        exported_files = []  # Track exported files for background processing
        
        if use_background_processing:
            # Create file list for background processor
            timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            file_list_path = os.path.join(export_folder, "dwfx_processing_queue_{}.txt".format(timestamp))
            print("  Background processing enabled")
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
                final_filepath = os.path.join(export_folder, filename + ".dwfx")
                
                # Create ViewSet with single sheet
                view_set = DB.ViewSet()
                view_set.Insert(sheet)
                
                # Export to final location (requires transaction)
                t = DB.Transaction(doc, "Export DWFX")
                t.Start()
                try:
                    doc.Export(export_folder, filename, view_set, dwfx_options)
                    t.Commit()
                except:
                    t.RollBack()
                    raise
                
                print("  \u2713 Exported: {} \u2192 {}".format(sheet.SheetNumber, filename + ".dwfx"))
                
                # Track file for background processing
                if use_background_processing:
                    exported_files.append(final_filepath)
                
                exported_count += 1
                
            except Exception as e:
                print("  \u2717 Failed: {} - {}".format(sheet.SheetNumber, str(e)))
                failed_sheets.append((sheet.SheetNumber, str(e)))
        
        # 8. Launch background processor if white removal enabled
        if use_background_processing and exported_files:
            try:
                # Write file list
                with open(file_list_path, 'w') as f:
                    for filepath in exported_files:
                        f.write(filepath + "\n")
                
                # Find Python interpreter and postprocessor script
                processor_script = os.path.join(lib_path, "dwfx_postprocessor.py")
                
                if not os.path.exists(processor_script):
                    print("\nWARNING: Background processor not found at: {}".format(processor_script))
                    print("White background removal will not be applied.")
                else:
                    # Launch external Python process (detached)
                    python_exe = find_python_executable(prefer_pythonw=True)
                    
                    # Start process detached (no wait)
                    if os.name == 'nt':  # Windows
                        # Use CREATE_NO_WINDOW flag to hide console
                        subprocess.Popen(
                            [python_exe, processor_script, file_list_path],
                            creationflags=subprocess.CREATE_NO_WINDOW if hasattr(subprocess, 'CREATE_NO_WINDOW') else 0,
                            close_fds=True
                        )
                    else:  # Unix-like
                        subprocess.Popen(
                            [python_exe, processor_script, file_list_path],
                            close_fds=True,
                            start_new_session=True
                        )
                    
                    log_filename = "dwfx_postprocessing_{}.log".format(
                        datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
                    )
                    log_path = os.path.join(export_folder, log_filename)
                    
                    print("\nBackground processing started!")
                    print("Using Python: {}".format(python_exe))
                    print("Processing {} file(s) to remove white backgrounds...".format(len(exported_files)))
                    print("\nYou can continue working in Revit.")
                    print("Check log file for progress: {}".format(log_path))
                    
            except Exception as e:
                print("\nWARNING: Failed to start background processor: {}".format(str(e)))
                print("White background removal will not be applied.")
                # Clean up file list
                if file_list_path and os.path.exists(file_list_path):
                    try:
                        os.remove(file_list_path)
                    except:
                        pass
        
        # 9. Report results
        print("\n" + "="*60)
        print("EXPORT COMPLETE")
        print("="*60)
        print("Exported: {} of {} sheets".format(exported_count, len(sheets)))
        print("Export folder: {}".format(export_folder))
        
        if use_background_processing and exported_files:
            print("\nNote: White background removal is processing in the background.")
        
        if failed_sheets:
            print("\nFailed sheets:")
            for num, error in failed_sheets:
                print("  - Sheet {}: {}".format(num, error))
            
            if use_background_processing and exported_files:
                forms.alert(
                    "Export completed with errors.\n\n"
                    "Successfully exported: {} of {}\n\n"
                    "White background removal is processing in the background.\n"
                    "Check console and log file for details.".format(exported_count, len(sheets)),
                    title="Export Completed with Errors"
                )
            else:
                forms.alert(
                    "Export completed with errors.\n\n"
                    "Successfully exported: {} of {}\n\n"
                    "Check console for details.".format(exported_count, len(sheets)),
                    title="Export Completed with Errors"
                )
        else:
            if use_background_processing and exported_files:
                forms.alert(
                    "Successfully exported {} sheet(s) to:\n\n{}\n\n"
                    "White background removal is processing in the background.\n"
                    "You can continue working - check the log file for progress.".format(
                        exported_count, export_folder
                    ),
                    title="Export Complete"
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
