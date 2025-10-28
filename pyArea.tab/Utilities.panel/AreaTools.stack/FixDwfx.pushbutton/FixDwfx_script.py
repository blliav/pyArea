# -*- coding: utf-8 -*-
"""Remove white background from DWFX floor plans."""

__title__ = "Fix\nDWFX"
__author__ = "Your Name"

import os
import io
import zipfile
import shutil
import tempfile
import re
import xml.etree.ElementTree as ET
from pyrevit import forms

def select_dwfx_files():
    """Prompt user to select one or multiple DWFX files."""
    files = forms.pick_file(file_ext='dwfx', multi_file=True, title='Select DWFX File(s)')
    return files if files else None


def find_fpage_files(root_dir):
    """Find all .fpage files in the directory structure."""
    fpage_files = []
    for dirpath, dirnames, filenames in os.walk(root_dir):
        for filename in filenames:
            if filename.endswith('.fpage'):
                fpage_files.append(os.path.join(dirpath, filename))
    return fpage_files


def get_sheet_name_from_xml(fpage_path):
    """Extract sheet name from descriptor.xml in the same directory."""
    directory = os.path.dirname(fpage_path)
    descriptor_path = os.path.join(directory, 'descriptor.xml')
    
    # Read descriptor.xml
    if os.path.exists(descriptor_path):
        try:
            tree = ET.parse(descriptor_path)
            root = tree.getroot()
            # Get name attribute from root element
            sheet_name = root.get('name')
            if sheet_name:
                return sheet_name
        except:
            pass
    
    # Fallback to directory name if not found
    return os.path.basename(directory)


def process_fpage_file(fpage_path):
    """Change all white fills to transparent in .fpage file."""
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
    except:
        return 0


def process_dwfx_file(dwfx_path):
    """Process DWFX file - change all white fills to transparent."""
    print("\nProcessing: {}".format(os.path.basename(dwfx_path)))
    
    temp_dir = tempfile.mkdtemp(prefix='dwfx_')
    
    try:
        # Extract and process
        with zipfile.ZipFile(dwfx_path, 'r') as zip_ref:
            zip_ref.extractall(temp_dir)
        
        total = 0
        sheet_results = []
        for fpage in find_fpage_files(temp_dir):
            changes = process_fpage_file(fpage)
            if changes > 0:
                sheet_name = get_sheet_name_from_xml(fpage)
                sheet_results.append((sheet_name, changes))
                print("  Sheet '{}' - Changes made: {}".format(sheet_name, changes))
            total += changes
        
        # Re-zip if modified
        if total > 0:
            backup = dwfx_path + '.backup'
            if not os.path.exists(backup):
                shutil.copy2(dwfx_path, backup)
            
            with zipfile.ZipFile(dwfx_path, 'w', zipfile.ZIP_DEFLATED) as zip_ref:
                for root, dirs, files in os.walk(temp_dir):
                    for file in files:
                        file_path = os.path.join(root, file)
                        arcname = os.path.relpath(file_path, temp_dir).replace('\\', '/')
                        zip_ref.write(file_path, arcname)
        else:
            print("No white fills found - file unchanged")
        
        return (total, sheet_results)
        
    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)


def fix_dwfx_background():
    """Main function - change white fills to transparent in DWFX files."""
    dwfx_files = select_dwfx_files()
    if not dwfx_files:
        forms.alert("No files selected.", exitscript=True)
    
    # Process files
    for dwfx_file in dwfx_files:
        try:
            process_dwfx_file(dwfx_file)
        except Exception as e:
            print("ERROR: {}".format(str(e)))


if __name__ == '__main__':
    fix_dwfx_background()
