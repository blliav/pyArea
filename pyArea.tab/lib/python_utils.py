# -*- coding: utf-8 -*-
"""Python Utilities

Helper functions for finding pyRevit's embedded Python.
Used by DWFx export/processing scripts that need external Python process.
"""

import os
import sys


def find_python_executable(prefer_pythonw=True):
    """
    Find Python executable - uses pyRevit's embedded Python.
    
    Works in both IronPython and CPython contexts within pyRevit.
    
    Args:
        prefer_pythonw: If True, prefer pythonw.exe over python.exe (no console window)
    
    Returns:
        str: Path to Python executable
    
    Raises:
        RuntimeError: If no Python executable found
    """
    exe_name = "pythonw.exe" if prefer_pythonw else "python.exe"
    alt_name = "python.exe" if prefer_pythonw else "pythonw.exe"
    
    # Method 1: Use pyRevit's HOME_DIR if available
    try:
        from pyrevit import HOME_DIR
        # pyRevit CPython is in: HOME_DIR/bin/cengines/CPY3XXX/
        cengines = os.path.join(HOME_DIR, 'bin', 'cengines')
        if os.path.exists(cengines):
            for engine in os.listdir(cengines):
                if engine.upper().startswith('CPY'):
                    python_exe = os.path.join(cengines, engine, exe_name)
                    if os.path.exists(python_exe):
                        return python_exe
                    python_alt = os.path.join(cengines, engine, alt_name)
                    if os.path.exists(python_alt):
                        return python_alt
    except ImportError:
        pass
    
    # Method 2: If running in CPython context, sys.executable points to Python
    if sys.executable and os.path.exists(sys.executable):
        exe_dir = os.path.dirname(sys.executable)
        python_exe = os.path.join(exe_dir, exe_name)
        if os.path.exists(python_exe):
            return python_exe
        python_alt = os.path.join(exe_dir, alt_name)
        if os.path.exists(python_alt):
            return python_alt
    
    # Method 3: Fallback - search common pyRevit locations
    appdata = os.environ.get('APPDATA', '')
    if appdata:
        for pyrevit_folder in ['pyRevit-Master', 'pyRevit', 'pyRevit-Dev']:
            cengines = os.path.join(appdata, pyrevit_folder, 'bin', 'cengines')
            if os.path.exists(cengines):
                try:
                    for engine in os.listdir(cengines):
                        if engine.upper().startswith('CPY'):
                            python_exe = os.path.join(cengines, engine, exe_name)
                            if os.path.exists(python_exe):
                                return python_exe
                            python_alt = os.path.join(cengines, engine, alt_name)
                            if os.path.exists(python_alt):
                                return python_alt
                except:
                    pass
    
    raise RuntimeError("Could not find pyRevit's embedded Python. Check pyRevit installation.")
