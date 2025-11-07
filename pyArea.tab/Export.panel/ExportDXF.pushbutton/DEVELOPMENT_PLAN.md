# ExportDXF Script Development Plan

**Date:** November 6, 2025  
**Script Type:** CPython 3 (pyRevit)  
**Purpose:** Export Area Plans to DXF with municipality-specific formatting using JSON-based extensible storage

---

## 1. Script Architecture Overview

```
ExportDXF_script.py (CPython 3)
├── Imports & Setup
├── Constants & Configuration
├── Utility Functions
│   ├── Coordinate Transformation
│   ├── Geometry Processing
│   └── Data Retrieval
├── DXF Creation Functions
│   ├── Layer Management
│   ├── Entity Creation (polylines, text, rectangles)
│   └── Underlay Management
├── Data Extraction Functions
│   ├── JSON Schema Reading
│   ├── Municipality Detection
│   └── String Formatting
├── Processing Functions
│   ├── Area Processing
│   ├── AreaPlan Processing
│   └── Sheet Processing
└── Main Export Orchestration
```

---

## 2. Naming Conventions

### Functions (snake_case)

#### Utility Functions
```python
convert_point_to_realworld()         # Convert Revit XYZ to real-world cm in DXF (applies scale + offset)
transform_point_to_sheet()           # Transform view point to sheet coordinates
calculate_arc_bulge()                # Calculate DXF bulge value for arcs
calculate_realworld_scale_factor()   # Calculate scale factor: Revit feet → real-world cm at view scale
```

#### Data Retrieval (JSON-based)
```python
get_json_data()                      # Get JSON from extensible storage
extract_municipality()               # Extract municipality from AreaScheme
get_area_data_for_dxf()             # Get Area JSON data
get_areaplan_data_for_dxf()         # Get AreaPlan JSON data
get_sheet_data_for_dxf()            # Get Sheet JSON data
```

#### String Formatting (Municipality-specific)
```python
format_area_string()                # Format area attributes string
format_areaplan_string()            # Format areaplan attributes string
format_sheet_string()               # Format sheet attributes string
```

#### DXF Creation
```python
create_dxf_layers()                 # Create DXF layers based on municipality
add_rectangle()                     # Add rectangle to DXF
add_text()                          # Add text to DXF
add_polyline_with_arcs()           # Add polyline with arc segments (bulge values)
add_dwfx_underlay()                # Add DWFX underlay reference
```

#### Processing Pipeline
```python
process_area()                      # Process single Area element
process_areaplan_viewport()         # Process AreaPlan viewport
process_sheet()                     # Process entire sheet
```

#### Sheet Management
```python
get_selected_sheets()               # Get sheets from project browser selection
sort_sheets_by_number()            # Sort sheets numerically
extract_sheet_number_for_sorting()  # Extract numeric portion for sorting
```

### Variables (snake_case)

#### Constants (UPPER_CASE)
```python
REALWORLD_SCALE_FACTOR             # Scale factor: Revit feet → real-world cm (accounts for view scale)
OFFSET_X, OFFSET_Y                 # Sheet origin offsets (for multi-sheet and coordinate alignment)
MUNICIPALITY_TYPE                  # Current municipality ("Common", "Jerusalem", "Tel-Aviv")
SCHEMA_GUID                        # Extensible storage schema GUID
SCHEMA_NAME                        # Schema name ("pyArea")
FIELD_NAME                         # Field name ("Data")
```

#### Local Variables
```python
area_boundary_curves               # List of boundary curve segments
sheet_width                        # Sheet width in Revit feet
view_scale                         # View scale number (e.g., 100 for 1:100)
municipality                       # Municipality string

# Type-specific naming
area_elem                          # DB.Area element
areaplan_view                      # DB.ViewPlan (AreaPlan type)
sheet_elem                         # DB.ViewSheet
area_scheme_elem                   # DB.AreaScheme

# Data dictionaries
area_data                          # JSON dict from Area
areaplan_data                      # JSON dict from AreaPlan
sheet_data                         # JSON dict from Sheet
```

---

### Scaling Logic - Critical Understanding

**Purpose:** Scale the sheet in DXF modelspace so that view content appears at real-world dimensions in centimeters.

**Formula:**
```python
# Base conversion: Revit internal units (feet) to centimeters
FEET_TO_CM = 30.48  # 1 foot = 30.48 cm

# View scale factor (from AreaPlan view or Sheet JSON)
# For 1:100 scale, view_scale = 100
# For 1:200 scale, view_scale = 200

# Combined scale factor:
REALWORLD_SCALE_FACTOR = FEET_TO_CM * view_scale

# Example calculations:
# At 1:100 scale: 30.48 * 100 = 3048
# At 1:200 scale: 30.48 * 200 = 6096
# At 1:50 scale:  30.48 * 50  = 1524

# Why this works:
# - Sheet coordinates are already scaled by view scale (e.g., 10-foot wall → 0.1 feet on sheet at 1:100)
# - To convert back to real-world cm: sheet_feet * view_scale * FEET_TO_CM
# - Combined: sheet_feet * (FEET_TO_CM * view_scale) = sheet_feet * REALWORLD_SCALE_FACTOR
```

**Result:** When you measure in the DXF (in cm), it matches the real-world dimensions that the view scale represents.

---

## 3. Interaction with Other Extension Modules

### Shared Module Strategy

**✅ CAN Import (Pure Python):**
- `municipality_schemas.py` - Field definitions, DXF config, validators
  - No Revit dependencies, pure Python data structures
  - Works in both CPython and IronPython

**❌ CANNOT Import (IronPython Only):**
- `schema_manager.py` - Uses IronPython-specific Revit API calls
- `data_manager.py` - Depends on schema_manager.py

### Reading JSON Data (CPython Compatible)

**Solution:** Implement direct JSON reading using Revit API:

```python
def get_json_data(element):
    """Read JSON data from extensible storage (CPython compatible)"""
    from Autodesk.Revit.DB.ExtensibleStorage import Schema
    import System
    import json
    
    # Schema constants
    SCHEMA_GUID = "A7B3C9D1-E5F2-4A8B-9C3D-1E2F3A4B5C6D"
    FIELD_NAME = "Data"
    
    # Get schema by GUID
    schema_guid = System.Guid(SCHEMA_GUID)
    schema = Schema.Lookup(schema_guid)
    
    if not schema:
        return {}
    
    entity = element.GetEntity(schema)
    if not entity.IsValid():
        return {}
    
    # Read JSON string
    json_string = entity.Get[str](FIELD_NAME)
    
    if not json_string:
        return {}
    
    return json.loads(json_string)
```

### Import Shared Definitions

**Import from `municipality_schemas.py`:**
```python
from municipality_schemas import (
    MUNICIPALITIES,           # List of valid municipalities
    DXF_CONFIG,              # DXF export configuration (layers, templates)
    SHEET_FIELDS,            # Field definitions (optional, for reference)
    AREAPLAN_FIELDS,
    AREA_FIELDS
)
```

### Replicate Helper Functions

```python
def get_municipality_from_areascheme(area_scheme):
    """Replicate data_manager functionality in CPython"""
    data = get_json_data(area_scheme)
    return data.get("Municipality", "Common")

def get_area_scheme_by_id(doc, element_id):
    """Get AreaScheme element by ID"""
    try:
        if isinstance(element_id, str):
            element_id = int(element_id)
        elem_id = DB.ElementId(System.Int64(element_id))
        element = doc.GetElement(elem_id)
        if isinstance(element, DB.AreaScheme):
            return element
        return None
    except:
        return None
```

---

## 4. CPython-Specific Considerations

### Imports Structure

```python
#! python3
# -*- coding: utf-8 -*-
"""
Export Area Plans to DXF with municipality-specific formatting.
Uses JSON-based extensible storage schema.
"""

# Add local lib directory for ezdxf package
import sys
import os
script_dir = os.path.dirname(__file__)
lib_dir = os.path.join(script_dir, 'lib')
if lib_dir not in sys.path:
    sys.path.insert(0, lib_dir)

# pyRevit imports (CPython compatible)
from pyrevit import revit, DB, UI, forms, script

# Standard Python libraries
import json
import math
import re

# External package (bundled in lib folder)
import ezdxf

# .NET interop
import clr
import System
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
from Autodesk.Revit.DB.ExtensibleStorage import Schema as ESSchema

doc = revit.doc
```

### Bundle Structure

```
ExportDXF.pushbutton/
├── ExportDXF_script.py          # Main CPython script
├── bundle.yaml                   # pyRevit button configuration
├── DEVELOPMENT_PLAN.md          # This document
└── lib/
    └── ezdxf/                   # CPython ezdxf package (bundled)
```

### String Handling

```python
# CPython 3: str is Unicode by default
text = "שלום"  # Hebrew works fine, no u"" prefix needed

# For .NET String interop:
from System import String
net_string = String(text)
```

### Key Differences from IronPython Scripts

| Aspect | IronPython Scripts | ExportDXF (CPython) |
|--------|-------------------|---------------------|
| Import shared modules | ✅ Can import `data_manager`, `schema_manager` | ❌ Cannot import (incompatible) |
| JSON reading | Via `schema_manager.get_data()` | Direct ExtensibleStorage API |
| External packages | Limited (.NET only) | ✅ Can use `ezdxf` |
| String handling | Mixed bytes/unicode | Pure Unicode |
| pyRevit API | Available | Available |

---

## 5. Data Flow & Processing Order

### Main Execution Flow

```
1. INITIALIZATION
   └─ get_selected_sheets() or use active sheet
   └─ sort_sheets_by_number(descending=True)
   └─ Create DXF document

2. FOR EACH SHEET (left-to-right layout):
   
   a. GET SHEET METADATA
      └─ get_sheet_data_for_dxf(sheet)
         ├─ Read JSON from sheet extensible storage
         ├─ Find parent AreaScheme via AreaSchemeId
         ├─ Get municipality from AreaScheme
         └─ Extract scale, coordinates, project info
   
   b. SETUP DXF ENVIRONMENT
      └─ create_dxf_layers(dwg, municipality)
      └─ Calculate and set REALWORLD_SCALE_FACTOR, OFFSET_X, OFFSET_Y
   
   c. PROCESS SHEET CONTENT
      └─ process_sheet(sheet, horizontal_offset, page_number):
         
         ├─ Get titleblock bounding box
         ├─ Add DWFX underlay (background)
         ├─ Add titleblock frame rectangle
         ├─ Add sheet string (formatted per municipality)
         
         └─ FOR EACH viewport on sheet:
            
            ├─ IF viewport contains AreaPlan view:
               └─ process_areaplan_viewport(viewport, view):
                  
                  ├─ Get crop boundary points
                  ├─ Add crop boundary rectangle
                  ├─ Get areaplan data from JSON
                  ├─ Add areaplan string (formatted per municipality)
                  
                  └─ FOR EACH area in view:
                     └─ process_area(area):
                        ├─ Get boundary curve segments
                        ├─ Transform to sheet coordinates
                        ├─ Create polyline with arc bulge values
                        ├─ Get area data (JSON + parameters)
                        └─ Add area string (formatted per municipality)

3. SAVE OUTPUT
   └─ Save DXF file (Desktop/Export/<filename>.dxf)
   └─ Create .dat file with DWFX_SCALE value
```

### Data Sources & Schema Structure

#### AreaScheme (Root)
```json
{
    "Municipality": "Jerusalem"
}
```
**Purpose:** Store municipality type that applies to all child elements

#### Sheet
```json
{
    "AreaSchemeId": "123456",
    "PROJECT": "Project Name",
    "ELEVATION": "0.00",
    "BUILDING_HEIGHT": "15.5",
    "X": "123456.78",
    "Y": "234567.89",
    "LOT_AREA": "500.0"
}
```
**Purpose:** Sheet-level attributes for DXF export

#### AreaPlan (View)
```json
{
    "BUILDING_NAME": "A",
    "FLOOR_NAME": "Floor 1",
    "FLOOR_ELEVATION": "0.00",
    "FLOOR_UNDERGROUND": "no",
    "RepresentedViews": ["123", "456"]
}
```
**Purpose:** Floor-level attributes, may represent multiple floors

#### Area
```json
{
    "HEIGHT": "2.80",
    "AREA": "25.5",
    "APPARTMENT_NUM": "A1",
    "HEIGHT2": ""
}
```
**Plus shared parameters:**
- `Usage Type` (parameter, not JSON)
- `Usage Type Prev` (parameter, not JSON)

---

## 6. Municipality-Specific Configuration

### Single Source of Truth: `municipality_schemas.DXF_CONFIG`

**All DXF export settings are defined in:** `lib/schemas/municipality_schemas.py`

This ensures both CPython (ExportDXF) and IronPython scripts (SetAreas, CalculationSetup) share the same field definitions.

### Usage in ExportDXF:

```python
# Import shared schema definitions
import sys
import os
schemas_path = os.path.join(os.path.dirname(__file__), '..', '..', 'lib', 'schemas')
sys.path.insert(0, schemas_path)

from municipality_schemas import MUNICIPALITIES, DXF_CONFIG

# Get configuration for current municipality (direct dictionary access)
dxf_config = DXF_CONFIG[municipality]

# Access layers
layers = dxf_config["layers"]
layer_colors = dxf_config["layer_colors"]

# Access string templates
sheet_template = dxf_config["string_templates"]["sheet"]
areaplan_template = dxf_config["string_templates"]["areaplan"]
area_template = dxf_config["string_templates"]["area"]

# Optional: Use .get() for safe fallback to Common if needed
# dxf_config = DXF_CONFIG.get(municipality, DXF_CONFIG["Common"])
```

### Configuration Structure:

Each municipality has:
- **`layers`** - DXF layer names for each entity type
- **`layer_colors`** - AutoCAD color numbers (1=Red, 3=Green, 7=White)
- **`string_templates`** - Python format strings for sheet/areaplan/area attributes

**Example for Jerusalem:**
```python
DXF_CONFIG["Jerusalem"] = {
    "layers": {
        'sheet_frame': 'AREA_PLAN_MAIN_FRAME',
        'area_boundary': 'AREA_PLAN_BORDER',
        # ...
    },
    "layer_colors": {
        'AREA_PLAN_MAIN_FRAME': 7,    # White
        'AREA_PLAN_BORDER': 1,         # Red
        # ...
    },
    "string_templates": {
        "sheet": "PROJECT={project}&&&ELEVATION={elevation}&&&...",
        "areaplan": "BUILDING_NAME={building_name}&&&FLOOR_NAME={floor_name}&&&...",
        "area": "CODE={code}&&&DEMOLITION_SOURCE_CODE={demolition_source_code}&&&..."
    }
}
```

**Benefits:**
- ✅ Single point of modification for all municipalities
- ✅ No code duplication between scripts
- ✅ Changes to field definitions automatically sync
- ✅ Pure Python module - works in both CPython and IronPython

---

## 7. Error Handling & Robustness

### Graceful Degradation Strategy

```python
def format_area_string_safe(area, municipality):
    """Safely create area string with fallbacks"""
    try:
        return format_area_string(area, municipality)
    except Exception as e:
        print("  Warning: Error formatting area {}: {}".format(area.Id, e))
        # Return minimal valid string
        if municipality == "Jerusalem":
            return "CODE=&&&DEMOLITION_SOURCE_CODE=&&&AREA=&&&HEIGHT1=&&&APPARTMENT_NUM=&&&HEIGHT2="
        else:
            return "USAGE_TYPE=&&&USAGE_TYPE_OLD=&&&AREA=&&&ASSET="
```

### Missing Data Handling

```python
# Always provide fallback values
municipality = data.get("Municipality", "Common")
elevation = data.get("ELEVATION", "0.00")
building_name = data.get("BUILDING_NAME", "1")

# Handle missing parameters gracefully
usage_type = ""
param = area.LookupParameter("Usage Type")
if param and param.HasValue:
    usage_type = param.AsString() or ""
```

### Geometry Error Handling

```python
# Try arc with bulge, fallback to line segments
try:
    bulge = calculate_arc_bulge(start, end, center, mid)
    points.append((x, y, 0, 0, bulge))
except Exception as e:
    print("  Arc processing error: {}, using line".format(e))
    points.append((x, y, 0, 0, 0))  # Straight line
```

---

## 8. Testing Strategy

### Test Scenarios

1. **Single Sheet Export - Common Municipality**
   - Verify layer names (RZ_*)
   - Verify string formats
   - Check scale calculation

2. **Single Sheet Export - Jerusalem Municipality**
   - Verify layer names (AREA_PLAN_*)
   - Verify extended string formats with all fields
   - Check coordinate extraction

3. **Multi-Sheet Export**
   - Verify horizontal offset calculation
   - Verify PAGE_NO numbering (rightmost = 1)
   - Check sheet sorting

4. **Geometry Handling**
   - Areas with arc boundaries
   - Areas with spline/ellipse boundaries (tessellation)
   - Multiple boundary loops (interior holes)

5. **Missing Data Scenarios**
   - Areas without Usage Type parameters
   - Views without JSON data
   - Sheets without AreaScheme reference

6. **Edge Cases**
   - Empty sheets (no viewports)
   - Non-AreaPlan viewports on sheet
   - Mixed municipalities (if supported)

---

## 9. Implementation Phases

### Phase 1: Foundation
- [ ] Imports and constants
- [ ] JSON reading functions
- [ ] Basic utility functions (coordinate conversion)

### Phase 2: Data Extraction
- [ ] `get_json_data()` - Core JSON reading
- [ ] Municipality detection from AreaScheme
- [ ] Data getters for Sheet, AreaPlan, Area

### Phase 3: String Formatting
- [ ] `format_sheet_string()` - Municipality-specific
- [ ] `format_areaplan_string()` - Municipality-specific
- [ ] `format_area_string()` - Municipality-specific

### Phase 4: DXF Creation
- [ ] Layer management
- [ ] Rectangle/text/polyline functions
- [ ] Arc bulge calculation
- [ ] DWFX underlay support

### Phase 5: Processing Pipeline
- [ ] `process_area()` - Single area processing
- [ ] `process_areaplan_viewport()` - View processing
- [ ] `process_sheet()` - Sheet processing with offset

### Phase 6: Main Orchestration
- [ ] Sheet selection and sorting
- [ ] Multi-sheet layout logic
- [ ] File saving (.dxf and .dat)
- [ ] Error reporting

### Phase 7: Testing & Refinement
- [ ] Test all municipalities
- [ ] Test multi-sheet export
- [ ] Handle edge cases
- [ ] Performance optimization

---

## 10. Constants Reference

```python
# Schema identification
SCHEMA_GUID = "A7B3C9D1-E5F2-4A8B-9C3D-1E2F3A4B5C6D"
SCHEMA_NAME = "pyArea"
FIELD_NAME = "Data"

# Coordinate conversion constants
FEET_TO_CM = 30.48          # Revit internal units to centimeters
DEFAULT_VIEW_SCALE = 100.0  # Default scale (1:100) if not found
```

### pyRevit Button Configuration (bundle.yaml)

```yaml
title: "Export\nDXF"
tooltip: "Export Area Plans to DXF for municipality submission"
context:
  - active-view-is-sheet
```

---

**End of Development Plan**
