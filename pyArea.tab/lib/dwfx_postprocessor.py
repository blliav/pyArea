# -*- coding: utf-8 -*-
"""Standalone DWFX Post-Processor

Processes DWFX files to remove opaque white backgrounds.
Runs as external process (CPython) independent of Revit.

Usage:
    python dwfx_postprocessor.py <file_list_path>
    
Where file_list_path is a text file containing one DWFX filepath per line.
"""

import sys
import os
import io
import zipfile
import shutil
import tempfile
import re
import datetime
import traceback


# ============================================================
# DWFX PROCESSING LOGIC
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
        print("  WARNING: Failed to process fpage file {}: {}".format(fpage_path, str(e)))
        return 0


def fix_dwfx_file(dwfx_path):
    """
    Remove opaque white background from DWFX file (in-place).
    
    Extracts DWFX (zip format), processes all .fpage files to replace
    white fills (#FFFFFF) with transparent (#00FFFFFF), then re-zips.
    
    Args:
        dwfx_path: Path to DWFX file to process (will be modified in-place)
    
    Returns:
        tuple: (total_changes, success, error_message)
    """
    temp_dir = None
    try:
        # Create temp directory
        temp_dir = tempfile.mkdtemp(prefix='dwfx_fix_')
        
        # Extract DWFX (it's a zip file)
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
        
        return (total_changes, True, None)
        
    except Exception as e:
        error_msg = "Failed to fix DWFX file: {}".format(str(e))
        return (0, False, error_msg)
        
    finally:
        # Cleanup temp directory
        if temp_dir and os.path.exists(temp_dir):
            shutil.rmtree(temp_dir, ignore_errors=True)


# ============================================================
# BATCH PROCESSING
# ============================================================

def process_file_list(file_list_path, log_path):
    """
    Process all DWFX files listed in the file_list_path.
    Writes detailed log to log_path.
    
    Args:
        file_list_path: Path to text file containing DWFX file paths (one per line)
        log_path: Path to write log file
    
    Returns:
        tuple: (total_processed, total_succeeded, total_failed)
    """
    # Read file list
    try:
        with open(file_list_path, 'r') as f:
            files = [line.strip() for line in f if line.strip()]
    except Exception as e:
        with open(log_path, 'w') as log:
            log.write("ERROR: Failed to read file list: {}\n".format(str(e)))
        return (0, 0, 1)
    
    if not files:
        with open(log_path, 'w') as log:
            log.write("ERROR: No files found in file list\n")
        return (0, 0, 1)
    
    # Process files and log results
    total_processed = 0
    total_succeeded = 0
    total_failed = 0
    
    with open(log_path, 'w') as log:
        log.write("DWFX Post-Processing Log\n")
        log.write("Started: {}\n".format(datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
        log.write("="*60 + "\n\n")
        
        for i, dwfx_path in enumerate(files, 1):
            filename = os.path.basename(dwfx_path)
            log.write("[{}/{}] Processing: {}\n".format(i, len(files), filename))
            
            # Check if file exists
            if not os.path.exists(dwfx_path):
                log.write("  ERROR: File not found\n\n")
                total_failed += 1
                continue
            
            # Process file
            try:
                changes, success, error = fix_dwfx_file(dwfx_path)
                total_processed += 1
                
                if success:
                    if changes > 0:
                        log.write("  SUCCESS: Removed {} white fills\n".format(changes))
                    else:
                        log.write("  SUCCESS: No white fills found (file unchanged)\n")
                    total_succeeded += 1
                else:
                    log.write("  ERROR: {}\n".format(error if error else "Unknown error"))
                    total_failed += 1
                    
            except Exception as e:
                log.write("  ERROR: Unexpected error: {}\n".format(str(e)))
                log.write("  Traceback:\n")
                log.write(traceback.format_exc())
                total_failed += 1
            
            log.write("\n")
            log.flush()  # Flush after each file for real-time logging
        
        # Summary
        log.write("="*60 + "\n")
        log.write("SUMMARY\n")
        log.write("="*60 + "\n")
        log.write("Total files: {}\n".format(len(files)))
        log.write("Processed: {}\n".format(total_processed))
        log.write("Succeeded: {}\n".format(total_succeeded))
        log.write("Failed: {}\n".format(total_failed))
        log.write("\nCompleted: {}\n".format(datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
    
    return (total_processed, total_succeeded, total_failed)


# ============================================================
# MAIN ENTRY POINT
# ============================================================

def main():
    """Main entry point for standalone execution"""
    if len(sys.argv) != 2:
        print("Usage: python dwfx_postprocessor.py <file_list_path>")
        print("\nWhere file_list_path is a text file containing one DWFX filepath per line.")
        sys.exit(1)
    
    file_list_path = sys.argv[1]
    
    # Generate log file path (same directory as file list)
    log_dir = os.path.dirname(file_list_path)
    log_filename = "dwfx_postprocessing_{}.log".format(
        datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    )
    log_path = os.path.join(log_dir, log_filename)
    
    print("DWFX Post-Processor")
    print("="*60)
    print("File list: {}".format(file_list_path))
    print("Log file: {}".format(log_path))
    print("\nProcessing in background...")
    print("Check log file for progress and results.")
    print("="*60)
    
    # Process files
    total, succeeded, failed = process_file_list(file_list_path, log_path)
    
    # Print summary to console
    print("\nProcessing complete!")
    print("Processed: {} | Succeeded: {} | Failed: {}".format(total, succeeded, failed))
    print("\nSee log file for details: {}".format(log_path))
    
    # Exit with code 0 if all succeeded, 1 if any failed
    sys.exit(0 if failed == 0 else 1)


if __name__ == "__main__":
    main()
