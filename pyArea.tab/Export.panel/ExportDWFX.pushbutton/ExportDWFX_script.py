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
    Validates all sheets belong to same AreaScheme.
    Raises ValueError if validation fails.
    
    Returns:
        str: Uniform AreaScheme ID
    """
    from data_manager import get_calculation_from_sheet
    
    schemes_found = {}
    missing_scheme_sheets = []
    
    for sheet in sheets:
        # Get AreaScheme from sheet's viewport (v2.0 approach)
        area_scheme, calculation_data = get_calculation_from_sheet(doc, sheet)
        
        if not area_scheme:
            missing_scheme_sheets.append(sheet.SheetNumber)
            continue
        
        area_scheme_id = str(area_scheme.Id.Value)
        
        if area_scheme_id not in schemes_found:
            schemes_found[area_scheme_id] = []
        schemes_found[area_scheme_id].append(sheet.SheetNumber)
    
    # Error if sheets missing AreaScheme
    if missing_scheme_sheets:
        error_msg = "ERROR: Sheets without AreaScheme detected!\n\n"
        error_msg += "The following sheets have no AreaScheme:\n"
        for num in missing_scheme_sheets:
            error_msg += "  - Sheet {}\n".format(num)
        error_msg += "\nAll sheets must have AreaPlan views for DWFX export."
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
                
                # Export (requires transaction)
                t = DB.Transaction(doc, "Export DWFX")
                t.Start()
                try:
                    doc.Export(export_target_folder, filename, view_set, dwfx_options)
                    t.Commit()
                except:
                    t.RollBack()
                    raise
                
                if use_background_processing:
                    output.print_md("✅ Exported to temp: **{}** → **{}**".format(sheet.SheetNumber, filename + ".dwfx"))
                    # Track temp file for background processing
                    temp_files.append(temp_filepath)
                else:
                    output.print_md("✅ Exported: **{}** → **{}**".format(sheet.SheetNumber, filename + ".dwfx"))
                
                exported_count += 1
                
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
                processor_script = os.path.join(lib_path, "dwfx_postprocessor.py")
                
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
