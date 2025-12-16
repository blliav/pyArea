# -*- coding: utf-8 -*-
"""Python Utilities

Helper functions for finding pyRevit's embedded Python and managing dependencies.
Used by DWFx export/processing scripts that need external Python process.

COMPATIBILITY: This module is compatible with both IronPython 2.7 and CPython 3.x.
- IronPython functions: find_python_executable()
- CPython functions: install_packages_from_pypi(), get_vendor_cpython_dir(), ensure_vendor_cpython_in_path()
"""

import os
import sys

# Import CPython-specific modules only when needed (not available in IronPython 2.7)
try:
    import urllib.request
    import zipfile
    import tempfile
    import json as json_module
    _CPYTHON_AVAILABLE = True
except ImportError:
    _CPYTHON_AVAILABLE = False


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


def get_vendor_cpython_dir():
    """Get the vendor_cpython directory path for CPython external packages.
    
    Returns:
        str: Absolute path to pyArea.tab/lib/vendor_cpython/
    """
    lib_dir = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(lib_dir, 'vendor_cpython')


def install_packages_from_pypi(packages, target_dir=None):
    """Download and install packages from PyPI without pip.
    
    Downloads wheel files directly from PyPI and extracts them to the target directory.
    Handles both pure Python wheels and platform-specific wheels (for numpy, etc.).
    
    Args:
        packages: List of package names to install (e.g., ['ezdxf', 'numpy'])
        target_dir: Directory to extract packages into. Defaults to lib/vendor_cpython/
    
    Returns:
        bool: True if all packages installed successfully
        
    Raises:
        ImportError: If no compatible wheel found for a package or CPython modules unavailable
    """
    if not _CPYTHON_AVAILABLE:
        raise ImportError("CPython modules (urllib, zipfile) not available. This function requires CPython.")
    
    if target_dir is None:
        target_dir = get_vendor_cpython_dir()
    
    os.makedirs(target_dir, exist_ok=True)
    
    # Get Python version for wheel compatibility (e.g., "cp310" for 3.10)
    py_version = "cp{}{}".format(sys.version_info.major, sys.version_info.minor)
    
    for package in packages:
        print("pyArea: Downloading {}...".format(package))
        
        # Get package info from PyPI JSON API
        api_url = "https://pypi.org/pypi/{}/json".format(package)
        with urllib.request.urlopen(api_url, timeout=30) as response:
            data = json_module.loads(response.read().decode())
        
        # Find compatible wheel
        wheel_url = None
        for file_info in data['urls']:
            filename = file_info['filename']
            # Try pure Python wheel first
            if filename.endswith('-py3-none-any.whl'):
                wheel_url = file_info['url']
                break
            # For numpy: need platform-specific wheel (Windows 64-bit)
            if package == 'numpy' and py_version in filename and 'win_amd64' in filename:
                wheel_url = file_info['url']
                break
        
        if not wheel_url:
            raise ImportError("No compatible wheel found for {} (Python {})".format(package, py_version))
        
        # Download wheel to temp file
        with tempfile.NamedTemporaryFile(suffix='.whl', delete=False) as tmp:
            tmp_path = tmp.name
            urllib.request.urlretrieve(wheel_url, tmp_path)
        
        # Extract wheel (it's just a zip file)
        with zipfile.ZipFile(tmp_path, 'r') as whl:
            whl.extractall(target_dir)
        
        # Cleanup
        os.remove(tmp_path)
        print("pyArea: {} installed".format(package))
    
    # Add to path if not already there
    if target_dir not in sys.path:
        sys.path.insert(0, target_dir)
    
    print("pyArea: All dependencies installed successfully")
    return True


def ensure_vendor_cpython_in_path():
    """Ensure the vendor_cpython directory is in sys.path.
    
    Call this before importing CPython external packages that live in vendor_cpython/.
    
    Returns:
        str: Path to vendor_cpython directory
    """
    vendor_dir = get_vendor_cpython_dir()
    if vendor_dir not in sys.path:
        sys.path.insert(0, vendor_dir)
    return vendor_dir
