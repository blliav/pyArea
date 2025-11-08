# ExportDXF Script Development Plan

**Date:** November 8, 2025 (Updated)  
**Script Type:** CPython 3 (pyRevit)  
**Purpose:** Export Area Plans to DXF with municipality-specific formatting using JSON-based extensible storage  

**Architectural Approach:** Single procedural script with clear sections. After evaluating Clean Architecture (DTOs/Services/Adapters), chose procedural approach as better fit for this tool's complexity (~600-1000 lines ETL flow). Accept ~100 lines of code duplication vs IronPython modules rather than IPC or bridge module complexity.

---

## 1. Script Architecture Overview

**Approach:** Single procedural script with clear sections and logical function grouping.

```python
# ExportDXF_script.py (CPython 3)

# ============================================================================
# SECTION 1: IMPORTS & SETUP
# ============================================================================
# - Standard library imports (json, math, os, sys)
# - pyRevit imports (revit, DB, UI, script) - NOTE: forms NOT compatible with CPython
# - External packages (ezdxf)
# - .NET interop (System, ExtensibleStorage, System.Windows.Forms)
# - Path setup for lib/ and lib/schemas/

# ============================================================================
# SECTION 2: CONSTANTS & CONFIGURATION
# ============================================================================
# - Import SCHEMA_GUID, SCHEMA_NAME, FIELD_NAME from schema_guids
# - FEET_TO_CM = 30.48
# - DEFAULT_VIEW_SCALE = 100.0
# - Import DXF_CONFIG from municipality_schemas

# ============================================================================
# SECTION 3: DATA EXTRACTION (JSON + Revit API)
# ============================================================================
def get_json_data(element)              # Core JSON reading from ExtensibleStorage
def get_municipality_from_areascheme()  # Extract municipality
def get_area_scheme_by_id()             # Helper to get AreaScheme element
def get_sheet_data_for_dxf()            # Extract sheet data + municipality
def get_areaplan_data_for_dxf()         # Extract areaplan data
def get_area_data_for_dxf()             # Extract area data + parameters

# ============================================================================
# SECTION 4: COORDINATE & GEOMETRY UTILITIES
# ============================================================================
def calculate_realworld_scale_factor()  # Compute FEET_TO_CM * view_scale
def convert_point_to_realworld()        # Revit XYZ → real-world cm in DXF
def transform_point_to_sheet()          # Transform view coordinates to sheet using Revit matrices
def calculate_arc_bulge()               # Calculate DXF bulge value for arcs

# ============================================================================
# SECTION 5: STRING FORMATTING (Municipality-specific)
# ============================================================================
def format_sheet_string()               # Format sheet attributes using DXF_CONFIG
def format_areaplan_string()            # Format areaplan attributes
def format_area_string()                # Format area attributes

# ============================================================================
# SECTION 6: DXF LAYER & ENTITY CREATION
# ============================================================================
def create_dxf_layers()                 # Create layers from DXF_CONFIG
def add_rectangle()                     # Add rectangle to DXF
def add_text()                          # Add text entity to DXF
def add_polyline_with_arcs()            # Add polyline with arc segments (bulges)
def add_dwfx_underlay()                 # Add DWFX underlay reference (optional)

# ============================================================================
# SECTION 7: PROCESSING PIPELINE
# ============================================================================
def process_area()                      # Process single Area element
def process_areaplan_viewport()         # Process AreaPlan viewport
def process_sheet()                     # Process entire sheet with offset

# ============================================================================
# SECTION 8: SHEET SELECTION, VALIDATION & SORTING
# ============================================================================
def get_valid_areaplans_and_uniform_scale()  # Validate views & uniform scale (CRITICAL)
def get_selected_sheets()               # Get sheets from project browser
def sort_sheets_by_number()             # Sort sheets numerically
def extract_sheet_number_for_sorting()  # Extract numeric portion

# ============================================================================
# SECTION 9: MAIN ORCHESTRATION
# ============================================================================
if __name__ == '__main__':
    # 1. Get sheets (active or selected)
    # 2. Sort sheets (rightmost = page 1)
    # 3. Validate & filter: get valid AreaPlans + uniform scale
    # 4. Create DXF document
    # 5. For each sheet: process with horizontal offset
    # 6. Save .dxf and .dat files
    # 7. Report results
```

**Key Principles:**
- Functions grouped by responsibility (data, geometry, formatting, DXF, processing)
- Clear section markers for easy navigation
- Dependency flow: low-level utilities first, high-level processing last
- Single source of truth: `municipality_schemas.py` for templates and layers

---

## 2. Naming Conventions

**Functions:** snake_case  
**Variables:** snake_case with type hints where helpful  
**Constants:** UPPER_CASE  

### Function Naming by Section

All functions follow the pattern: `verb_noun()` for clarity and searchability.

**Section 3 - Data Extraction:** `get_*()` or `extract_*()`  
**Section 4 - Geometry:** `calculate_*()`, `convert_*()`, `transform_*()`  
**Section 5 - Formatting:** `format_*_string()`  
**Section 6 - DXF Creation:** `create_*()`, `add_*()`  
**Section 7 - Processing:** `process_*()`  
**Section 8 - Selection:** `get_*()`, `sort_*()`  

### Example Functions per Section

```python
# Section 3: Data Extraction
get_json_data(element)                    # Read JSON from ExtensibleStorage
get_municipality_from_areascheme(scheme)  # Extract municipality string
get_sheet_data_for_dxf(sheet)             # Get all sheet data for export

# Section 4: Geometry
calculate_realworld_scale_factor(view_scale)           # FEET_TO_CM * view_scale
convert_point_to_realworld(xyz, scale, offset_x, offset_y)  # Revit XYZ → DXF cm
transform_point_to_sheet(view_point, viewport)         # View coords → Sheet coords
calculate_arc_bulge(start, end, mid)                   # Arc → bulge value

# Section 5: Formatting
format_sheet_string(sheet_data, municipality)    # Build sheet attribute string
format_areaplan_string(plan_data, municipality)  # Build areaplan string
format_area_string(area_data, municipality)      # Build area string

# Section 6: DXF Creation
create_dxf_layers(dxf_doc, municipality)        # Setup layers from DXF_CONFIG
add_polyline_with_arcs(msp, points, layer) # Draw polyline with bulge
add_text(msp, text, point, layer)          # Draw text entity

# Section 7: Processing
process_area(area_elem, msp, scale, offset_x, offset_y, municipality, layers)  # Process single area
process_areaplan_viewport(viewport, msp, scale, offset_x, offset_y, municipality, layers)  # Process viewport
process_sheet(sheet_elem, dxf_doc, msp, horizontal_offset, page_number, view_scale, valid_viewports)  # Process entire sheet

# Section 8: Selection & Validation
get_valid_areaplans_and_uniform_scale(sheets)  # Validate & filter views, check scale
get_selected_sheets()                     # Get sheets from UI
sort_sheets_by_number(sheets)             # Sort numerically
extract_sheet_number_for_sorting(sheet)   # Extract numeric portion
```

### Variables (snake_case)

#### Module Constants (UPPER_CASE)
```python
FEET_TO_CM = 30.48                 # Revit feet to centimeters conversion
DEFAULT_VIEW_SCALE = 100.0         # Default scale if not found
SCHEMA_GUID                        # Extensible storage schema GUID (imported)
SCHEMA_NAME                        # Schema name "pyArea" (imported)
FIELD_NAME                         # Field name "Data" (imported)
```

#### Runtime Variables (Calculated/Passed)
```python
view_scale                         # Validated uniform scale (e.g., 100 for 1:100)
scale_factor                       # Calculated: FEET_TO_CM * view_scale
municipality                       # Municipality string per sheet
offset_x, offset_y                 # Sheet origin offsets (for multi-sheet layout)
valid_viewports                    # Pre-validated viewport list per sheet
area_boundary_curves               # List of boundary curve segments
sheet_width                        # Sheet width in cm (calculated)

# Type-specific naming (all Revit elements use _elem suffix)
area_elem                          # DB.Area element
areaplan_elem                      # DB.ViewPlan (AreaPlan type)
sheet_elem                         # DB.ViewSheet element
areascheme_elem                    # DB.AreaScheme element

# Data dictionaries
area_data                          # JSON dict from Area
areaplan_data                      # JSON dict from AreaPlan
sheet_data                         # JSON dict from Sheet
```

---

### Coordinate Transformation Chain - Critical Understanding

**Three-Step Flow:**
```
VIEW coordinates → SHEET coordinates → DXF coordinates
(Revit feet)        (Revit feet)        (real-world cm)
```

**Step 1: View to Sheet Transformation**
```python
def transform_point_to_sheet(view_point, viewport):
    """Transform using Revit's transformation matrices"""
    view = doc.GetElement(viewport.ViewId)
    transform_w_boundary = view.GetModelToProjectionTransforms()[0]
    model_to_proj = transform_w_boundary.GetModelToProjectionTransform()
    proj_to_sheet = viewport.GetProjectionToSheetTransform()
    
    proj_point = model_to_proj.OfPoint(view_point)
    sheet_point = proj_to_sheet.OfPoint(proj_point)
    return sheet_point
```

**Step 2: Sheet to DXF Transformation**
```python
def convert_point_to_realworld(xyz, scale_factor, offset_x, offset_y):
    """Convert sheet coordinates to DXF real-world cm
    
    Formula: (point - offset) * scale
    Offset applied BEFORE scaling to move origin correctly.
    """
    return ((xyz.X - offset_x) * scale_factor, (xyz.Y - offset_y) * scale_factor)

# Scale factor calculation:
REALWORLD_SCALE_FACTOR = FEET_TO_CM * view_scale
# At 1:100 scale: 30.48 * 100 = 3048
# At 1:200 scale: 30.48 * 200 = 6096
```

**Key Principles:**
- Area boundaries from `GetBoundarySegments()` are in **VIEW coordinates**
- Crop boundaries from `GetCropShape()` are in **VIEW coordinates**
- Must transform to SHEET coordinates before applying DXF scale/offset
- Offset subtraction before scaling ensures correct origin placement

### Text Positioning Strategy

**Principle:** Calculate text positions directly in DXF space after transforming frames.

```python
# Sheet text at top-right corner (10 cm offset)
max_point_dxf = convert_point_to_realworld(bbox.Max, scale_factor, offset_x, offset_y)
text_pos = (max_point_dxf[0] - 10.0, max_point_dxf[1] - 10.0)

# AreaPlan text at top-right corner (200 cm offset)
max_x_dxf = max(x for x, y in transformed_crop)
max_y_dxf = max(y for x, y in transformed_crop)
text_pos = (max_x_dxf - 200.0, max_y_dxf - 200.0)
```

**Benefits:**
- ✅ Work directly in target coordinate system (DXF cm)
- ✅ No unit conversion roundtrips
- ✅ Offsets are exactly what appears in DXF
- ✅ Simpler and more intuitive

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
    
    # Schema constants (imported at module level from schema_guids.py)
    # SCHEMA_GUID, FIELD_NAME already available
    
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

**Import from shared modules:**
```python
# From schema_guids.py
from schema_guids import SCHEMA_GUID, SCHEMA_NAME, FIELD_NAME

# From municipality_schemas.py
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
from pyrevit import revit, DB, UI, script  # NOTE: forms not available in CPython

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
clr.AddReference('System.Windows.Forms')  # For MessageBox dialogs
from Autodesk.Revit.DB.ExtensibleStorage import Schema as ESSchema
from System.Windows.Forms import MessageBox, MessageBoxButtons, MessageBoxIcon

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
      └─ create_dxf_layers(dxf_doc, municipality)
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
                        ├─ Collect ALL boundary points (list)
                        ├─ Batch transform: convert_points_to_realworld()
                        ├─ Create polyline with arc bulge values
                        ├─ Get area Location.Point for text placement
                        ├─ Transform text position
                        ├─ Get area data (JSON + parameters)
                        └─ Add area text at Location.Point

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
   - Empty sheets (no viewports) → Skipped during processing
   - Non-AreaPlan viewports on sheet → Ignored during validation
   - AreaPlan views without municipality → Filtered out with warning
   - AreaPlan views without areas → Filtered out with warning
   - Mixed municipalities (if supported) → Each sheet uses its own municipality
   - **Mixed scales (validated and blocked)** ⚠️ → Export fails with detailed error

7. **Valid AreaPlan Criteria** (Applied During Validation)
   - Must be of ViewType `DB.ViewType.AreaPlan`
   - Must have an AreaSchemeId that resolves to valid AreaScheme element
   - AreaScheme must have municipality defined in ExtensibleStorage JSON
   - Must contain at least one Area element
   - Must have a valid `Scale` property

---

## 9. Main Orchestration Flow

**Section 9 (Main Block) Execution Order:**

1. **Get Sheets** - From selection or active view
2. **Sort Sheets** - By sheet number (descending, rightmost = page 1)
3. **⚠️ Validate & Filter** - Single comprehensive validation pass (CRITICAL)
   - `get_valid_areaplans_and_uniform_scale(sorted_sheets)`
   - Filters valid AreaPlan views (has municipality, has areas, has scale)
   - Validates uniform scale across all valid views
   - Returns: `(uniform_scale, {sheet.Id: [valid_viewports]})`
   - If validation fails → show detailed error and EXIT
4. **Create DXF Document** - Initialize ezdxf document and modelspace
5. **Process Sheets** - Iterate through sorted sheets with horizontal offset
   - Pass validated scale AND pre-validated viewports to each `process_sheet()` call
   - Only process sheets with valid viewports
   - Track cumulative horizontal offset
6. **Generate Filename** - `<modelname>-<firstsheet>..<lastsheet>_<timestamp>`
7. **Save DXF File** - Write to Desktop/Export/
8. **Create DAT File** - Write `DWFX_SCALE = view_scale / 10`
9. **Report Results** - Console output and success dialog

**Key Architectural Points:**
- **Validation happens ONCE** at orchestration level (step 3)
- **Filtering and scale validation unified** - no redundant checks
- Valid viewports are **pre-identified** and passed down
- `process_sheet()` receives both validated scale and valid viewports
- Processing layer has **no validation logic** - only data transformation

---

## 10. Implementation Phases

### Phase 1: Foundation ✅ COMPLETE
- [x] Imports and constants
- [x] JSON reading functions
- [x] Basic utility functions (coordinate conversion)

### Phase 2: Data Extraction ✅ COMPLETE
- [x] `get_json_data()` - Core JSON reading
- [x] Municipality detection from AreaScheme
- [x] Data getters for Sheet, AreaPlan, Area

### Phase 3: String Formatting ✅ COMPLETE
- [x] `format_sheet_string()` - Municipality-specific
- [x] `format_areaplan_string()` - Municipality-specific
- [x] `format_area_string()` - Municipality-specific

### Phase 4: DXF Creation ✅ COMPLETE
- [x] Layer management
- [x] Rectangle/text/polyline functions
- [x] Arc bulge calculation
- [x] DWFX underlay support (placeholder)

### Phase 5: Processing Pipeline ✅ COMPLETE
- [x] `process_area()` - Single area processing
- [x] `process_areaplan_viewport()` - View processing
- [x] `process_sheet()` - Sheet processing with offset

### Phase 6: Main Orchestration ✅ COMPLETE
- [x] Sheet selection and sorting
- [x] Multi-sheet layout logic
- [x] File saving (.dxf and .dat)
- [x] Error reporting

### Phase 7: Testing & Refinement 🔄 IN PROGRESS
- [ ] Test all municipalities
- [ ] Test multi-sheet export
- [ ] Handle edge cases
- [ ] Performance optimization

---

## 11. Implementation Insights & Lessons Learned

**Completed:** November 7, 2025

### Key Implementation Decisions

1. **Unified Validation Architecture**
   - **Decision:** Single validation pass that filters valid views AND validates scale
   - **Rationale:** Avoid redundancy; clear separation between validation and processing
   - **Implementation:** 
     - `get_valid_areaplans_and_uniform_scale()` performs all validation in one pass
     - Filters views by: has municipality, has areas, has scale
     - Validates uniform scale across filtered views
     - Returns: `(uniform_scale, {sheet.Id: [valid_viewports]})`
     - `process_sheet()` receives both scale and pre-validated viewports
   - **Architecture:** Validation at orchestration level, processing is pure transformation
   - **Benefit:** No validation logic duplication between validation and processing layers

2. **Schema Constants Import**
   - **Decision:** Import `SCHEMA_GUID`, `SCHEMA_NAME`, `FIELD_NAME` from `schema_guids.py`
   - **Rationale:** Single source of truth; prevents hardcoded value drift
   - **Updated:** Both script and plan to use imports instead of hardcoding

3. **Coordinate Transformation Order**
   - **Decision:** Apply offset BEFORE scaling: `(point - offset) * scale`
   - **Rationale:** Correct origin placement for multi-sheet layout
   - **Formula:** Subtract sheet minimum from point, then scale to real-world
   - **Implementation:** All coordinate transformations use this pattern consistently

4. **Arc Handling**
   - **Decision:** Calculate bulge from arc midpoint using sagitta formula
   - **Fallback:** Graceful degradation to straight line if calculation fails
   - **Implementation:** `calculate_arc_bulge()` with cross-product for orientation

5. **Error Handling Philosophy**
   - **Approach:** Graceful degradation with warnings, not failures
   - **Examples:** Missing data uses fallbacks, failed geometry continues with defaults
   - **User Experience:** Export completes even with partial data issues

6. **File Naming Convention**
   - **Decision:** Use `<modelname>-<firstsheet>..<lastsheet>_<timestamp>` format
   - **Rationale:** Single DXF file contains all sheets; naming reflects sheet range
   - **Model name:** Retrieved from `doc.Title`
   - **Sheet range:** Shows first..last for multi-sheet exports

7. **DWFX_SCALE Calculation**
   - **Decision:** `DWFX_SCALE = view_scale / 10`
   - **Rationale:** DWFX files are in millimeters, DXF is in centimeters
   - **Formula explanation:** 
     - DXF scaled up by `view_scale` (e.g., 100x for 1:100)
     - DWFX in mm needs: scale by 100, divide by 10 (mm→cm) = 10
   - **Examples:** 1:100→10, 1:200→20, 1:50→5

8. **Unified Validation & Filtering** ⚠️ **CRITICAL**
   - **Decision:** Single validation pass filters valid AreaPlan views AND validates uniform scale
   - **Rationale:** Avoid validation redundancy; clear separation of validation vs processing
   - **Implementation:** `get_valid_areaplans_and_uniform_scale()` performs comprehensive validation
   - **Valid AreaPlan Criteria:**
     1. Must belong to an AreaScheme with defined municipality
     2. Must contain areas (not empty)
     3. Must have a defined scale
   - **Scale Validation:** All valid views must have same scale
   - **Returns:** `(uniform_scale, {sheet.Id: [valid_viewports]})`
   - **Error Handling:** 
     - No valid views found → detailed error
     - Mixed scales detected → detailed error with sheet/view list
   - **Processing:** Only pre-validated viewports are processed (no redundant checks)
   - **User Experience:** Clear error messages showing validation failures upfront

### File Output Structure

```
Desktop/
└── Export/
    ├── <modelname>-<firstsheet>..<lastsheet>_<timestamp>.dxf
    └── <modelname>-<firstsheet>..<lastsheet>_<timestamp>.dat
```

**Naming Convention:**
- **Model name**: From `doc.Title` (spaces and slashes replaced with underscores)
- **Sheet range**: 
  - Single sheet: `A101`
  - Multiple sheets: `A101..A105`
- **Timestamp**: `YYYYMMDD_HHMMSS`

**Examples:**
- Single sheet: `MyProject-A101_20251107_143022.dxf`
- Multiple sheets: `MyProject-A101..A105_20251107_143022.dxf`

**DAT File Content:**
```
DWFX_SCALE=<view_scale / 10>
```

**DWFX_SCALE Calculation:**
- DWFX files are in **millimeters**
- DXF is scaled to **centimeters** (real-world)
- When XREFing DWFX into DXF: `DWFX_SCALE = view_scale / 10`
- Examples:
  - 1:100 scale → `DWFX_SCALE=10.0`
  - 1:200 scale → `DWFX_SCALE=20.0`
  - 1:50 scale → `DWFX_SCALE=5.0`

### Dependencies

```python
# Standard library (built-in)
import sys, os, json, math, re
from datetime import datetime

# pyRevit (provided by host)
from pyrevit import revit, DB, UI, script  # forms not available in CPython

# .NET Windows Forms (for dialogs)
from System.Windows.Forms import MessageBox, MessageBoxButtons, MessageBoxIcon

# External (bundled in lib/)
import ezdxf  # DXF creation library

# .NET (via pythonnet/clr)
import clr, System
from Autodesk.Revit.DB.ExtensibleStorage import Schema
```

### Script Statistics

- **Total Lines:** ~1,320 lines
- **Sections:** 9 clearly marked sections
- **Functions:** 23 functions
- **Error Handlers:** Every function has try/except with meaningful messages
- **Comments:** ~15% of lines are documentation/comments

### DXF Configuration

- **DXF Version:** R2010 (AutoCAD 2010 format - widely compatible)
- **Units:** `$INSUNITS = 5` (centimeters)
- **Coordinate System:** Real-world centimeters in modelspace
- **Polylines:** All set to closed with `polyline.closed = True`
- **Text Height:** 10.0 cm (default for all text entities)

### Corrections & Refinements

**Post-Implementation Updates (Nov 7-8, 2025):**
- ✅ Fixed filename convention to use model name and sheet range
- ✅ Corrected DWFX_SCALE calculation (view_scale / 10 instead of 1.0)
- ✅ Changed DWFX_SCALE to integer format (10 not 10.0)
- ✅ Commented out per-run timestamp in exported filenames to match municipality workflow expectations
- ✅ Corrected crop boundary extraction to iterate CurveLoops from `GetCropShape()`
- ✅ Fixed area boundary processing by calling `BoundarySegment.GetCurve()` before curve evaluation
- ✅ **CRITICAL REFACTOR:** Unified validation architecture
  - Single validation function: `get_valid_areaplans_and_uniform_scale()`
  - Filters valid AreaPlan views (has municipality, has areas, has scale)
  - Validates uniform scale across all valid views
  - Returns both scale and pre-validated viewports map
  - `process_sheet()` receives both validated scale AND valid viewports
  - Eliminates validation redundancy between validation and processing layers
  - Clear architectural separation: validation upfront, processing is pure transformation
- ✅ **CPython Compatibility Fixes (Nov 8, 2025):**
  - Replaced `pyrevit.forms` with `System.Windows.Forms.MessageBox` (forms not available in CPython)
  - Fixed AreaScheme access: use `view.AreaScheme` property instead of `view.AreaSchemeId` + `doc.GetElement()`
  - All dialogs now use .NET MessageBox with proper icons (Warning, Error, Information)
  - Replaced `exitscript=True` with `sys.exit()` for CPython compatibility
- ✅ **Coordinate Transformation Refinements (Nov 9, 2025):**
  - Proper view-to-sheet transformation using Revit's transformation matrices
  - Simplified text positioning by working directly in DXF space
  - Set DXF units explicitly to centimeters ($INSUNITS = 5)
  - All polylines set to closed with `polyline.closed = True`
  - Export only exterior boundary loops for areas (ignore interior holes)

---

## 12. Constants Reference

```python
# Schema identification (imported from schema_guids.py)
from schema_guids import SCHEMA_GUID, SCHEMA_NAME, FIELD_NAME
# SCHEMA_GUID = "A7B3C9D1-E5F2-4A8B-9C3D-1E2F3A4B5C6D"
# SCHEMA_NAME = "pyArea"
# FIELD_NAME = "Data"

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
