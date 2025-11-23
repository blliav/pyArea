# -*- coding: utf-8 -*-
"""Export utilities for DXF and DWFx exporters
Pure Python - compatible with both CPython and IronPython
"""

import re
import os
import io
import zipfile
import shutil
import tempfile

# ============================================================
# CONFIGURATION DEFAULTS
# ============================================================

DEFAULT_EXPORT_FOLDER = "Desktop/Export"
DEFAULT_DXF_CREATE_DAT = True
DEFAULT_DWFx_EXPORT_ELEMENT_DATA = True
DEFAULT_DWFx_QUALITY = "Medium"
DEFAULT_DWFx_REMOVE_OPAQUE_WHITE = True

DWFx_QUALITY_MAP = {
    "Low": {"ImageQuality": 1},
    "Medium": {"ImageQuality": 2},
    "High": {"ImageQuality": 3}
}


def get_default_preferences():
    """Returns default preferences dictionary"""
    return {
        "ExportFolder": DEFAULT_EXPORT_FOLDER,
        "DXF_CreateDatFile": DEFAULT_DXF_CREATE_DAT,
        "DWFx_ExportElementData": DEFAULT_DWFx_EXPORT_ELEMENT_DATA,
        "DWFx_Quality": DEFAULT_DWFx_QUALITY,
        "DWFx_RemoveOpaqueWhite": DEFAULT_DWFx_REMOVE_OPAQUE_WHITE
    }


# ============================================================
# NAMING UTILITIES
# ============================================================

def sanitize_filename_part(text):
    """Replaces invalid filename characters with underscores"""
    invalid_chars = r'[<>:"/\\|?*]'
    return re.sub(invalid_chars, '_', str(text)).strip()


def format_sheet_range(sheet_numbers):
    """
    Formats sheet numbers for filename.
    Single sheet: "A101"
    Multiple sheets: "A101..A105"
    """
    if len(sheet_numbers) == 1:
        return sanitize_filename_part(sheet_numbers[0])
    return "{}..{}".format(
        sanitize_filename_part(sheet_numbers[0]),
        sanitize_filename_part(sheet_numbers[-1])
    )


def generate_dxf_filename(model_name, sheet_numbers):
    """
    Generate DXF filename with sheet range.
    Format: ModelName-SheetRange
    Example: MyProject-A101..A105
    """
    model = sanitize_filename_part(model_name if model_name else "Model")
    sheets = format_sheet_range(sheet_numbers)
    return "{}-{}".format(model, sheets)


def generate_dwfx_filename(model_name, sheet_number):
    """
    Generate DWFx filename for single sheet.
    Format: ModelName-Sheet
    Example: MyProject-A101
    """
    model = sanitize_filename_part(model_name if model_name else "Model")
    sheet = sanitize_filename_part(sheet_number)
    return "{}-{}".format(model, sheet)


def get_export_folder_path(config_folder=None):
    """
    Resolves export folder path.
    Handles Desktop/ relative paths.
    
    Args:
        config_folder: Configured folder path (can be relative or absolute)
    
    Returns:
        Absolute path to export folder
    """
    if not config_folder:
        config_folder = DEFAULT_EXPORT_FOLDER
    
    # Handle absolute paths
    if os.path.isabs(config_folder):
        return config_folder
    
    # Handle Desktop/ relative paths
    if config_folder.startswith("Desktop"):
        desktop = os.path.join(os.path.expanduser("~"), "Desktop")
        parts = config_folder.replace("\\", "/").split("/", 1)
        if len(parts) > 1:
            return os.path.join(desktop, parts[1])
        return desktop
    
    return config_folder


# ============================================================
# DWFx OPAQUE WHITE REMOVAL
# ============================================================

def find_fpage_files(root_dir):
    """Find all .fpage files in the directory structure"""
    fpage_files = []
    for dirpath, dirnames, filenames in os.walk(root_dir):
        for filename in filenames:
            if filename.endswith('.fpage'):
                fpage_files.append(os.path.join(dirpath, filename))
    return fpage_files


def process_fpage_file(fpage_path):
    """
    Change all white fills to transparent in .fpage file.
    Replaces fill="#FFFFFF" with fill="#00FFFFFF" in Path elements.
    
    Returns:
        Number of changes made
    """
    try:
        with io.open(fpage_path, 'r', encoding='utf-8') as f:
            content = f.read()
        
        # Replace fill="#FFFFFF" within Path elements only
        modified = re.sub(r'(<Path[^>]*\s+fill=")#FFFFFF(")', r'\1#00FFFFFF\2', content, flags=re.IGNORECASE)
        modified = re.sub(r'(<Path[^>]*\s+Fill=")#FFFFFF(")', r'\1#00FFFFFF\2', modified, flags=re.IGNORECASE)
        
        if content != modified:
            with io.open(fpage_path, 'w', encoding='utf-8') as f:
                f.write(modified)
            return len(re.findall(r'#00FFFFFF', modified)) - len(re.findall(r'#00FFFFFF', content))
        
        return 0
    except Exception as e:
        print("Warning: Failed to process fpage file {}: {}".format(fpage_path, str(e)))
        return 0


def fix_dwfx_file(dwfx_path):
    """
    Remove opaque white background from DWFx file.
    
    Extracts DWFx (zip format), processes all .fpage files to replace
    white fills (#FFFFFF) with transparent (#00FFFFFF), then re-zips.
    
    Args:
        dwfx_path: Path to DWFx file to process
    
    Returns:
        tuple: (total_changes, success)
    """
    temp_dir = None
    try:
        # Create temp directory
        temp_dir = tempfile.mkdtemp(prefix='dwfx_fix_')
        
        # Extract DWFx (it's a zip file)
        with zipfile.ZipFile(dwfx_path, 'r') as zip_ref:
            zip_ref.extractall(temp_dir)
        
        # Process all .fpage files
        total_changes = 0
        fpage_files = find_fpage_files(temp_dir)
        
        for fpage in fpage_files:
            changes = process_fpage_file(fpage)
            total_changes += changes
        
        # Re-zip only if changes were made
        if total_changes > 0:
            with zipfile.ZipFile(dwfx_path, 'w', zipfile.ZIP_DEFLATED) as zip_ref:
                for root, dirs, files in os.walk(temp_dir):
                    for file in files:
                        file_path = os.path.join(root, file)
                        arcname = os.path.relpath(file_path, temp_dir).replace('\\', '/')
                        zip_ref.write(file_path, arcname)
        
        return (total_changes, True)
        
    except Exception as e:
        print("ERROR: Failed to fix DWFx file {}: {}".format(dwfx_path, str(e)))
        return (0, False)
        
    finally:
        # Cleanup temp directory
        if temp_dir and os.path.exists(temp_dir):
            shutil.rmtree(temp_dir, ignore_errors=True)
