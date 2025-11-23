# -*- coding: utf-8 -*-
"""Remove white background from DWFX floor plans."""

__title__ = "Fix\nDWFX"

import os
import sys
import subprocess
import datetime
from pyrevit import forms

# Add lib to path
script_dir = os.path.dirname(__file__)
lib_path = os.path.join(script_dir, "..", "..", "..", "lib")
if lib_path not in sys.path:
    sys.path.insert(0, lib_path)

from python_utils import find_python_executable

def select_dwfx_files():
    """Prompt user to select one or multiple DWFX files."""
    files = forms.pick_file(file_ext='dwfx', multi_file=True, title='Select DWFX File(s)')
    return files if files else None


def launch_background_processor(dwfx_files):
    """Launch external processor for DWFX files in background."""
    try:
        # Create temp file list
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        
        # Use temp directory for file list
        import tempfile
        temp_dir = tempfile.gettempdir()
        file_list_path = os.path.join(temp_dir, "dwfx_fix_queue_{}.txt".format(timestamp))
        
        # Write file list
        with open(file_list_path, 'w') as f:
            for filepath in dwfx_files:
                f.write(filepath + "\n")
        
        # Find Python interpreter and postprocessor script
        processor_script = os.path.join(lib_path, "dwfx_postprocessor.py")
        
        if not os.path.exists(processor_script):
            forms.alert(
                "Background processor not found at:\n\n{}\n\n"
                "Please check your installation.".format(processor_script),
                title="Processor Not Found"
            )
            return False
        
        # Launch external Python process in visible console window
        python_exe = find_python_executable(prefer_pythonw=False)
        
        # Start process with visible console (detached)
        if os.name == 'nt':  # Windows
            # Use CREATE_NEW_CONSOLE to show console window
            subprocess.Popen(
                [python_exe, processor_script, file_list_path],
                creationflags=subprocess.CREATE_NEW_CONSOLE,
                close_fds=True
            )
        else:  # Unix-like
            subprocess.Popen(
                [python_exe, processor_script, file_list_path],
                close_fds=True,
                start_new_session=True
            )
        
        # Show minimal confirmation - detailed progress is in console window
        print("\nConsole window opened with processing progress.")
        print("You can continue working in Revit.")
        print("\n{} file(s) queued for processing.".format(len(dwfx_files)))
        
        return True
        
    except Exception as e:
        forms.alert(
            "Failed to start background processor:\n\n{}\n\n"
            "Check console for details.".format(str(e)),
            title="Error"
        )
        import traceback
        traceback.print_exc()
        return False


def fix_dwfx_background():
    """Main function - launch background processor for DWFX white removal."""
    # Select files
    dwfx_files = select_dwfx_files()
    if not dwfx_files:
        print("No files selected.")
        return
    
    # Launch background processor (shows progress in console window)
    launch_background_processor(dwfx_files)


if __name__ == '__main__':
    fix_dwfx_background()
