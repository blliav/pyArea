#! python3
# -*- coding: utf-8 -*-
"""Install Python matching pyRevit's version and install ezdxf."""

__context__ = 'zero-doc'
__title__ = 'Install\nDependencies'

import sys
import os
import subprocess
import urllib.request
import tempfile
import winreg

# Add lib to path
script_dir = os.path.dirname(__file__)
lib_path = os.path.join(script_dir, "..", "..", "..", "lib")
if lib_path not in sys.path:
    sys.path.insert(0, lib_path)

from python_utils import get_pyrevit_python_version, get_python_dir_for_version, PYTHON_BASE_DIR

# Configuration
AVAILABLE_VERSIONS = {
    "3.8": "3.8.10",
    "3.9": "3.9.13", 
    "3.10": "3.10.11",
    "3.11": "3.11.9",
    "3.12": "3.12.7",
    "3.13": "3.13.0"
}

def check_python_exists(version):
    """Check if matching Python exists."""
    python_dir = get_python_dir_for_version(version)
    python_exe = os.path.join(python_dir, "python.exe")
    return os.path.exists(python_exe), python_exe

def download_and_install_python(version):
    """Download and install Python silently."""
    full_version = AVAILABLE_VERSIONS.get(version)
    if not full_version:
        print("❌ Python {} not available".format(version))
        return False
    
    url = "https://www.python.org/ftp/python/{}/python-{}-amd64.exe".format(
        full_version, full_version)
    target_dir = get_python_dir_for_version(version)
    
    print("📥 Downloading Python {}...".format(full_version))
    
    # Download installer
    temp_dir = tempfile.mkdtemp()
    installer = os.path.join(temp_dir, "python-installer.exe")
    
    try:
        urllib.request.urlretrieve(url, installer)
        
        # Install silently
        print("📦 Installing to {}...".format(target_dir))
        os.makedirs(target_dir, exist_ok=True)
        
        cmd = [
            installer, "/quiet",
            "InstallAllUsers=0",
            "PrependPath=0",
            "Include_test=0",
            "TargetDir={}".format(target_dir)
        ]
        
        result = subprocess.run(cmd, timeout=600)
        
        # Cleanup
        try:
            os.remove(installer)
            os.rmdir(temp_dir)
        except:
            pass
        
        return result.returncode == 0
        
    except Exception as e:
        print("❌ Installation failed: {}".format(str(e)))
        return False

def check_and_install_ezdxf(python_exe):
    """Check if ezdxf exists, install if not."""
    # Check if ezdxf exists
    result = subprocess.run(
        [python_exe, "-c", "import ezdxf; print('OK')"],
        capture_output=True, text=True, timeout=10
    )
    
    if result.returncode == 0 and "OK" in result.stdout:
        print("✅ ezdxf already installed")
        return True
    
    # Install ezdxf
    print("📦 Installing ezdxf...")
    result = subprocess.run(
        [python_exe, "-m", "pip", "install", "ezdxf"],
        capture_output=True, text=True, timeout=600
    )
    
    if result.returncode == 0:
        print("✅ ezdxf installed successfully")
        return True
    else:
        print("❌ ezdxf installation failed")
        return False

def set_pythonpath(python_dir):
    """Set PYTHONPATH to Python's site-packages."""
    site_packages = os.path.join(python_dir, "Lib", "site-packages")
    
    try:
        # Open registry
        key = winreg.OpenKey(
            winreg.HKEY_CURRENT_USER, r"Environment", 
            0, winreg.KEY_ALL_ACCESS
        )
        
        # Get existing PYTHONPATH
        try:
            existing, _ = winreg.QueryValueEx(key, "PYTHONPATH")
        except FileNotFoundError:
            existing = ""
        
        # Clean out old Python paths
        paths = existing.split(os.pathsep) if existing else []
        cleaned = [p for p in paths if not ('Python' in p and p != site_packages)]
        
        # Add new path at beginning
        if site_packages not in cleaned:
            cleaned.insert(0, site_packages)
        
        new_pythonpath = os.pathsep.join(cleaned)
        
        # Set in registry
        winreg.SetValueEx(key, "PYTHONPATH", 0, winreg.REG_EXPAND_SZ, new_pythonpath)
        winreg.CloseKey(key)
        
        # Set for current session
        os.environ['PYTHONPATH'] = new_pythonpath
        
        print("✅ PYTHONPATH set to: {}".format(site_packages))
        return True
        
    except Exception as e:
        print("❌ Failed to set PYTHONPATH: {}".format(str(e)))
        return False

def main():
    """Main function."""
    print("=" * 60)
    print("pyRevit Dependencies Installer")
    print("=" * 60)
    
    # Step 1: Get pyRevit version
    pyrevit_version = get_pyrevit_python_version()
    print("\n📌 pyRevit Python version: {}".format(pyrevit_version))
    
    # Step 2: Check/Install matching Python
    print("\n🔍 Checking for Python {}...".format(pyrevit_version))
    exists, python_exe = check_python_exists(pyrevit_version)
    
    if not exists:
        print("❌ Python {} not found".format(pyrevit_version))
        if not download_and_install_python(pyrevit_version):
            print("\n❌ FAILED: Could not install Python")
            return
        exists, python_exe = check_python_exists(pyrevit_version)
        if not exists:
            print("\n❌ FAILED: Python installation verification failed")
            return
    
    print("✅ Python {} ready at: {}".format(pyrevit_version, python_exe))
    
    # Step 3: Check/Install ezdxf
    print("\n🔍 Checking ezdxf...")
    if not check_and_install_ezdxf(python_exe):
        print("\n⚠️  WARNING: ezdxf installation had issues")
    
    # Step 4: Set PYTHONPATH
    print("\n🔧 Setting PYTHONPATH...")
    python_dir = get_python_dir_for_version(pyrevit_version)
    set_pythonpath(python_dir)
    
    print("\n" + "=" * 60)
    print("✅ SETUP COMPLETE")
    print("=" * 60)
    print("Python {} is configured for pyRevit".format(pyrevit_version))
    print("Restart Revit for changes to take effect")

if __name__ == "__main__":
    main()
