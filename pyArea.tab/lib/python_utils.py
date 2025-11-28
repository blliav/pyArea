# -*- coding: utf-8 -*-
"""Python Utilities

Helper functions for finding and working with Python installations.
Compatible with both IronPython (pyRevit) and CPython environments.

Prefers the version-specific Python installed by InstallDependencies script.
"""

import os
import sys

# Configuration matching InstallDependencies
PYTHON_BASE_DIR = os.path.join(os.environ.get('LOCALAPPDATA', ''), 'Programs', 'Python')


def get_pyrevit_python_version():
    """Get pyRevit's Python major.minor version."""
    return "{}.{}".format(sys.version_info.major, sys.version_info.minor)


def get_python_dir_for_version(version):
    """Get Python installation directory for version (e.g., 3.12 -> Python312)."""
    major, minor = version.split('.')
    return os.path.join(PYTHON_BASE_DIR, "Python{}{}".format(major, minor))


def find_python_executable(prefer_pythonw=True):
    """
    Find Python executable on the system.
    
    Priority:
    1. Version-matching Python installed by InstallDependencies (in LOCALAPPDATA)
    2. PYTHONPATH environment variable
    3. sys.executable (standard Python environment)
    4. Search common installation locations
    5. "python" command as last resort
    
    Args:
        prefer_pythonw: If True, prefer pythonw.exe over python.exe (no console window)
    
    Returns:
        str: Path to Python executable, or "python" as fallback
    """
    exe_name = "pythonw.exe" if prefer_pythonw else "python.exe"
    
    # Priority 1: Version-matching Python (InstallDependencies location)
    try:
        pyrevit_version = get_pyrevit_python_version()
        python_dir = get_python_dir_for_version(pyrevit_version)
        python_exe = os.path.join(python_dir, exe_name)
        
        if os.path.exists(python_exe):
            return python_exe
        
        # Try python.exe if pythonw.exe not found
        if prefer_pythonw:
            python_exe_alt = os.path.join(python_dir, "python.exe")
            if os.path.exists(python_exe_alt):
                return python_exe_alt
    except:
        pass  # Continue to next method
    
    # Priority 2: Check PYTHONPATH environment variable
    pythonpath = os.environ.get('PYTHONPATH', '')
    if pythonpath:
        # PYTHONPATH typically points to site-packages, go up to Python root
        paths = pythonpath.split(os.pathsep)
        for path in paths:
            if 'Python' in path and 'site-packages' in path:
                # Extract Python root (e.g., ...\Python312)
                python_root = path.split('Lib')[0].rstrip(os.sep)
                python_exe = os.path.join(python_root, exe_name)
                if os.path.exists(python_exe):
                    return python_exe
                # Try python.exe if pythonw.exe not found
                if prefer_pythonw:
                    python_exe_alt = os.path.join(python_root, "python.exe")
                    if os.path.exists(python_exe_alt):
                        return python_exe_alt
    
    # Priority 3: sys.executable (standard Python environment)
    if sys.executable:
        if prefer_pythonw and sys.executable.endswith("python.exe"):
            pythonw = sys.executable.replace("python.exe", "pythonw.exe")
            if os.path.exists(pythonw):
                return pythonw
        return sys.executable
    
    # Priority 4: Search common installation locations (fallback)
    possible_paths = []
    
    # User AppData installations (InstallDependencies location)
    if os.path.exists(PYTHON_BASE_DIR):
        for version in ["Python313", "Python312", "Python311", "Python310", "Python39", "Python38"]:
            possible_paths.append(os.path.join(PYTHON_BASE_DIR, version, exe_name))
    
    # Direct installations (C:\PythonXX\)
    for version in ["313", "312", "311", "310", "39", "38"]:
        possible_paths.append(r"C:\Python{}\{}".format(version, exe_name))
    
    # Program Files installations
    for version in ["Python313", "Python312", "Python311", "Python310", "Python39", "Python38"]:
        possible_paths.append(r"C:\Program Files\{}\{}".format(version, exe_name))
    
    # Search for first existing path
    for path in possible_paths:
        if os.path.exists(path):
            return path
    
    # Priority 5: Fallback to 'python' command (relies on PATH)
    return "python"


def get_python_version(python_exe):
    """
    Get Python version string from executable.
    
    Args:
        python_exe: Path to Python executable
    
    Returns:
        str: Version string (e.g., "3.9.13") or None if error
    """
    try:
        import subprocess
        result = subprocess.run(
            [python_exe, "--version"],
            capture_output=True,
            text=True,
            timeout=5
        )
        if result.returncode == 0:
            # Output format: "Python 3.9.13"
            return result.stdout.strip().replace("Python ", "")
        return None
    except:
        return None


def verify_python_executable(python_exe):
    """
    Verify that a Python executable is valid and accessible.
    
    Args:
        python_exe: Path to Python executable
    
    Returns:
        bool: True if valid, False otherwise
    """
    if not python_exe or python_exe == "python":
        # Can't verify "python" command without executing
        return True
    
    return os.path.exists(python_exe) and os.path.isfile(python_exe)
