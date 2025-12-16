# -*- coding: utf-8 -*-
"""Standalone DWFx Post-Processor

Processes DWFx files to remove opaque white backgrounds.
Runs as external process (CPython) independent of Revit.

Usage:
    python DWFx_postprocessor.py <file_list_path> [final_folder]
    
Args:
    file_list_path: Text file with one DWFx filepath per line
    final_folder: (Optional) If provided, treats files as temp files,
                  moves processed files to final_folder, and cleans up temp files
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
# DWFx PROCESSING LOGIC
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
    Remove opaque white background from DWFx file (in-place).
    
    Extracts DWFx (zip format), processes all .fpage files to replace
    white fills (#FFFFFF) with transparent (#00FFFFFF), then re-zips.
    
    Args:
        dwfx_path: Path to DWFx file to process (will be modified in-place)
    
    Returns:
        tuple: (total_changes, success, error_message)
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
        
        return (total_changes, True, None)
        
    except Exception as e:
        error_msg = "Failed to fix DWFx file: {}".format(str(e))
        return (0, False, error_msg)
        
    finally:
        # Cleanup temp directory
        if temp_dir and os.path.exists(temp_dir):
            shutil.rmtree(temp_dir, ignore_errors=True)


# ============================================================
# BATCH PROCESSING
# ============================================================

def process_file_list(file_list_path, log_path, final_folder=None):
    """
    Process all DWFx files listed in the file_list_path.
    Writes detailed log to log_path and prints to console.
    
    Args:
        file_list_path: Path to text file containing DWFx file paths (one per line)
        log_path: Path to write log file
        final_folder: (Optional) If provided, moves processed files here and deletes temp files
    
    Returns:
        tuple: (total_processed, total_succeeded, total_failed)
    """
    # Read file list
    try:
        with io.open(file_list_path, 'r', encoding='utf-8') as f:
            files = [line.strip() for line in f if line.strip()]
    except Exception as e:
        with io.open(log_path, 'w', encoding='utf-8') as log:
            log.write("ERROR: Failed to read file list: {}\n".format(str(e)))
        return (0, 0, 1)
    
    if not files:
        with io.open(log_path, 'w', encoding='utf-8') as log:
            log.write("ERROR: No files found in file list\n")
        return (0, 0, 1)
    
    # Process files and log results
    total_processed = 0
    total_succeeded = 0
    total_failed = 0
    temp_files_to_delete = []  # Track temp files for cleanup
    
    # Ensure final_folder exists if provided
    if final_folder and not os.path.exists(final_folder):
        try:
            os.makedirs(final_folder)
        except Exception as e:
            with io.open(log_path, 'w', encoding='utf-8') as log:
                log.write("ERROR: Failed to create final folder: {}\n".format(str(e)))
            return (0, 0, 1)
    
    # Print header to console
    print("="*60)
    print("DWFx Post-Processing")
    print("="*60)
    print("Started: {}".format(datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
    if final_folder:
        print("Mode: Temp -> Final")
        print("Output folder: {}".format(final_folder))
    else:
        print("Mode: In-place")
    print("Files to process: {}".format(len(files)))
    print("="*60)
    print("")
    
    with io.open(log_path, 'w', encoding='utf-8') as log:
        log.write("DWFx Post-Processing Log\n")
        log.write("Started: {}\n".format(datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
        if final_folder:
            log.write("Mode: Temp -> Final ({})\n".format(final_folder))
        else:
            log.write("Mode: In-place\n")
        log.write("="*60 + "\n\n")
        
        for i, dwfx_path in enumerate(files, 1):
            filename = os.path.basename(dwfx_path)
            print("[{}/{}] Processing: {}".format(i, len(files), filename))
            log.write("[{}/{}] Processing: {}\n".format(i, len(files), filename))
            
            # Check if file exists
            if not os.path.exists(dwfx_path):
                print("  ERROR: File not found")
                log.write("  ERROR: File not found\n\n")
                total_failed += 1
                continue
            
            # Process file
            try:
                changes, success, error = fix_dwfx_file(dwfx_path)
                total_processed += 1
                
                if success:
                    if changes > 0:
                        print("  SUCCESS: Removed {} white fills".format(changes))
                        log.write("  SUCCESS: Removed {} white fills\n".format(changes))
                    else:
                        print("  SUCCESS: No white fills found")
                        log.write("  SUCCESS: No white fills found\n")
                    
                    # If final_folder provided, move file there
                    if final_folder:
                        final_path = os.path.join(final_folder, filename)
                        try:
                            shutil.move(dwfx_path, final_path)
                            print("  Moved to final folder")
                            log.write("  Moved to: {}\n".format(final_path))
                            # Don't add to delete list since move already removed it
                        except Exception as move_error:
                            print("  WARNING: Failed to move file")
                            log.write("  WARNING: Failed to move file: {}\n".format(str(move_error)))
                            # Mark for deletion anyway
                            temp_files_to_delete.append(dwfx_path)
                    
                    total_succeeded += 1
                else:
                    print("  ERROR: {}".format(error if error else "Unknown error"))
                    log.write("  ERROR: {}\n".format(error if error else "Unknown error"))
                    total_failed += 1
                    
            except Exception as e:
                print("  ERROR: Unexpected error: {}".format(str(e)))
                log.write("  ERROR: Unexpected error: {}\n".format(str(e)))
                log.write("  Traceback:\n")
                log.write(traceback.format_exc())
                total_failed += 1
            
            print("")  # Blank line between files
            log.write("\n")
            log.flush()  # Flush after each file for real-time logging
        
        # Summary
        print("="*60)
        print("SUMMARY")
        print("="*60)
        print("Total files: {}".format(len(files)))
        print("Processed: {}".format(total_processed))
        print("Succeeded: {}".format(total_succeeded))
        print("Failed: {}".format(total_failed))
        print("")
        print("Completed: {}".format(datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
        print("Log file: {}".format(log_path))
        print("="*60)
        
        log.write("="*60 + "\n")
        log.write("SUMMARY\n")
        log.write("="*60 + "\n")
        log.write("Total files: {}\n".format(len(files)))
        log.write("Processed: {}\n".format(total_processed))
        log.write("Succeeded: {}\n".format(total_succeeded))
        log.write("Failed: {}\n".format(total_failed))
        log.write("\nCompleted: {}\n".format(datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
    
    # Cleanup remaining temp files (if any failed to move)
    if temp_files_to_delete:
        for temp_file in temp_files_to_delete:
            try:
                if os.path.exists(temp_file):
                    os.remove(temp_file)
            except:
                pass
    
    return (total_processed, total_succeeded, total_failed)


# ============================================================
# MAIN ENTRY POINT
# ============================================================

def main():
    """Main entry point for standalone execution"""
    if len(sys.argv) not in (2, 3):
        print("Usage: python DWFx_postprocessor.py <file_list_path> [final_folder]")
        print("\nArgs:")
        print("  file_list_path: Text file with one DWFx filepath per line")
        print("  final_folder: (Optional) Move processed files here and cleanup temp files")
        sys.exit(1)
    
    file_list_path = sys.argv[1]
    final_folder = sys.argv[2] if len(sys.argv) == 3 else None
    
    # Generate log file path (same directory as file list)
    log_dir = os.path.dirname(file_list_path)
    log_filename = "dwfx_postprocessing_{}.log".format(
        datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    )
    log_path = os.path.join(log_dir, log_filename)
    
    # Note: Minimal console output since this runs in background with pythonw
    # All details are written to log file
    
    # Process files
    total, succeeded, failed = process_file_list(file_list_path, log_path, final_folder)
    
    # Pause for user if running in console (python.exe)
    # Will be skipped when running with pythonw.exe (no console)
    if sys.stdout and hasattr(sys.stdout, 'isatty') and sys.stdout.isatty():
        print("")
        try:
            # Python 2/3 compatibility
            try:
                input_func = raw_input
            except NameError:
                input_func = input
            input_func("Press Enter to close this window...")
        except (EOFError, KeyboardInterrupt):
            # Handle case where stdin is not available or user interrupts
            pass
    
    # Cleanup file list and temp directory if using final_folder mode
    if final_folder:
        try:
            if os.path.exists(file_list_path):
                temp_dir = os.path.dirname(file_list_path)
                os.remove(file_list_path)
                
                # Try to remove temp directory if it's empty
                try:
                    if os.path.exists(temp_dir) and not os.listdir(temp_dir):
                        os.rmdir(temp_dir)
                except:
                    pass
        except:
            pass  # Cleanup is best-effort
    
    # Exit with code 0 if all succeeded, 1 if any failed
    sys.exit(0 if failed == 0 else 1)


if __name__ == "__main__":
    main()
