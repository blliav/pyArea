# ExportDXF Script Development Plan

**Date:** November 8, 2025 (Updated November 12, 2025; November 20, 2025; November 21, 2025)  
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
def calculate_arc_bulge()               # Calculate DXF bulge value for arcs (uses center + mid-point)
# Curve handling: Arcs (bulge), Splines/Ellipses (tessellated), Lines (direct)

# ============================================================================
# SECTION 5: STRING FORMATTING (Municipality-specific)
# ============================================================================
def resolve_placeholder()               # Resolve placeholder strings (e.g., <Title on Sheet>)
def get_representedViews_data()         # Get floor data from represented views (uses Calculation defaults)
def format_sheet_string()               # Format sheet attributes using DXF_CONFIG
def format_areaplan_string()            # Format areaplan attributes
def format_usage_type()                 # Convert "0" to empty string for usage types
def format_area_string()                # Format area attributes

# ============================================================================
# SECTION 6: DXF LAYER & ENTITY CREATION
# ============================================================================
def create_dxf_layers()                 # Create layers from DXF_CONFIG
def add_rectangle()                     # Add rectangle to DXF
def add_text()                          # Add text entity to DXF
def add_polyline_with_arcs()            # Add polyline with arc segments (bulges)
def add_dwfx_underlay()                 # Add DWFX underlay reference with scale conversion

# ============================================================================
# SECTION 7: PROCESSING PIPELINE
# ============================================================================
def process_area()                      # Process single Area element
def process_areaplan_viewport()         # Process AreaPlan viewport
def process_sheet()                     # Process entire sheet with offset

# ============================================================================
# SECTION 8: SHEET SELECTION, VALIDATION & SORTING
# ============================================================================
def get_valid_areaplans_and_uniform_scale()  # Single-pass validation: AreaScheme + views + scale (CRITICAL)
def get_selected_sheets()               # Get sheets from project browser or active view
def sort_sheets_by_number()             # Sort sheets numerically
def extract_sheet_number_for_sorting()  # Extract numeric portion
def group_sheets_by_calculation()       # Group initial sheets by CalculationGuid
def expand_calculation_sheets()         # Expand a CalculationGuid to all its sheets in the model

# ============================================================================
# SECTION 9: MAIN ORCHESTRATION
# ============================================================================
if __name__ == '__main__':
    # 1. Get sheets (active or selected)
    # 2. Group sheets by CalculationGuid (legacy sheets grouped under None)
    # 3. For each Calculation group:
    #       - Expand to all sheets in that Calculation (v2.0), or keep selected legacy sheets
    #       - Sort sheets (rightmost = page 1)
    #       - Run single-pass validation: AreaScheme + valid AreaPlans + uniform scale
    #       - Create DXF document for this Calculation
    #       - For each sheet in group: process with horizontal offset
    #       - Save .dxf and .dat files (one DXF per Calculation group)
    # 4. Report overall results across all Calculation groups
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
format_usage_type(value)                         # Convert "0" to empty string
format_area_string(area_data, municipality)      # Build area string

# Section 6: DXF Creation
create_dxf_layers(dxf_doc, municipality)                  # Setup layers from DXF_CONFIG
add_polyline_with_arcs(msp, points, layer)                # Draw polyline with bulge
add_text(msp, text, point, layer)                         # Draw text entity
add_dwfx_underlay(dxf_doc, msp, filename, point, scale)   # Add DWFX underlay reference

# Section 7: Processing
process_area(area_elem, msp, scale, offset_x, offset_y, municipality, layers)  # Process single area
process_areaplan_viewport(viewport, msp, scale, offset_x, offset_y, municipality, layers)  # Process viewport
process_sheet(sheet_elem, dxf_doc, msp, horizontal_offset, page_number, view_scale, valid_viewports)  # Process entire sheet

# Section 8: Selection & Validation
get_valid_areaplans_and_uniform_scale(sheets)  # Single-pass: AreaScheme + views + scale
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

### DWFX Underlay Support

**Purpose:** Add DWFX file references as background underlays in DXF exports, allowing CAD users to view the original Revit sheet appearance behind the area boundaries.

**Implementation:**
```python
def add_dwfx_underlay(dxf_doc, msp, dwfx_filename, insert_point, scale):
    """Add DWFX underlay reference to DXF.
    
    Args:
        dxf_doc: ezdxf DXF document
        msp: DXF modelspace
        dwfx_filename: Filename of DWFX file (without path, just filename.dwfx)
        insert_point: (x, y) tuple for insertion point in DXF coordinates
        scale: View scale factor (e.g., 100 for 1:100, 200 for 1:200)
    
    Returns:
        bool: True if successful, False otherwise
    """
    # Scale conversion: DWFX is in mm, DXF is in cm
    dwfx_scale = scale / 10.0  # 1:100 → 10, 1:200 → 20
    
    # Create underlay definition
    underlay_def = dxf_doc.add_underlay_def(
        filename=dwfx_filename,
        fmt='dwf',  # DWF format covers both .dwf and .dwfx files
        name='1'    # First sheet in DWFX file (required for AutoCAD)
    )
    
    # Add underlay entity to modelspace
    underlay = msp.add_underlay(
        underlay_def,
        insert=insert_point,
        scale=(dwfx_scale, dwfx_scale, dwfx_scale)
    )
    
    # Assign to layer 0
    underlay.dxf.layer = '0'
```

**Usage in process_sheet():**
```python
# Add DWFX underlay (background reference)
print("  Attempting to add DWFX underlay...")

# Use custom DWFX filename from sheet if provided, otherwise generate default
custom_dwfx = sheet_data.get("DWFx_UnderlayFilename")
if custom_dwfx and custom_dwfx.strip():
    # User provided a custom filename on the Sheet JSON - use basename only
    import os
    dwfx_filename = os.path.basename(custom_dwfx.strip())
    # Ensure .dwfx extension
    if not dwfx_filename.lower().endswith('.dwfx'):
        dwfx_filename += '.dwfx'
    print("  DWFX filename (custom): {}".format(dwfx_filename))
else:
    # Fallback: generate DWFX filename from model title and sheet number
    dwfx_filename = export_utils.generate_dwfx_filename(doc.Title, sheet_elem.SheetNumber) + ".dwfx"
    print("  DWFX filename (generated): {}".format(dwfx_filename))

underlay_insert_point = convert_point_to_realworld(bbox.Min, scale_factor, offset_x, offset_y)
add_dwfx_underlay(dxf_doc, msp, dwfx_filename, underlay_insert_point, scale=view_scale)
```

**Key Points:**
- **Filename selection precedence:**
  1. If the sheet JSON contains a non-empty `DWFx_UnderlayFilename` field (set via CalculationSetup on the Sheet node), ExportDXF uses that value as the DWFX filename (basename only), ensuring it ends with `.dwfx`.
  2. Otherwise, it falls back to `export_utils.generate_dwfx_filename(doc.Title, sheet_elem.SheetNumber) + ".dwfx"` to match DWFX files created by ExportDWFX.
  - Default format: `{ModelName}-{SheetNumber}.dwfx` (e.g., `MyProject-A101.dwfx`)
- **Scale conversion:** DWFX files are in millimeters, DXF is in centimeters
  - Formula: `dwfx_scale = view_scale / 10.0`
  - Example: 1:100 scale → `dwfx_scale = 10`, 1:200 scale → `dwfx_scale = 20`
- **AutoCAD compatibility:** 
  - `fmt='dwf'` - File format type for DWF/DWFX files
  - `name='1'` - References first sheet in DWFX file (required for AutoCAD to load)
- **Layer assignment:** Underlay is placed on layer `'0'` (default layer)
- **File reference:** DXF stores only the filename (relative reference), DWFX and DXF files must be in same folder

**Coordinate Placement:**
- Underlay is inserted at sheet's bottom-left corner `(bbox.Min)`
- Uses same `convert_point_to_realworld()` transformation as sheet frame
- Positioned before other entities so it appears as background

**Relationship to .dat File:**
- The `.dat` file created during export contains `DWFX_SCALE={int(view_scale/10)}`
- This provides the same scale value for external reference workflows
- See Section 10 (Main Orchestration Flow) step 8 for .dat file creation

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

# External package (auto-downloaded to pyArea.tab/lib/vendor_cpython/ on first run)
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
└── DEVELOPMENT_PLAN.md          # This document
```

**Dependencies:** External packages (ezdxf, numpy, etc.) are auto-downloaded to `pyArea.tab/lib/vendor_cpython/` on first run. This central location is gitignored and shared across CPython scripts.

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
// Common & Jerusalem
{
    "BUILDING_NAME": "A",
    "FLOOR_NAME": "Floor 1",
    "FLOOR_ELEVATION": "0.00",
    "FLOOR_UNDERGROUND": "no",
    "RepresentedViews": ["123", "456"]
}

// Tel-Aviv (includes BUILDING field)
{
    "BUILDING": "1",
    "FLOOR": "Floor 1",
    "HEIGHT": "2.80",
    "X": "123456.78",
    "Y": "234567.89",
    "Absolute_height": "0.00",
    "RepresentedViews": ["123", "456"]
}
```
**Purpose:** Floor-level attributes, may represent multiple floors

#### Area
```json
// Common & Jerusalem
{
    "HEIGHT": "2.80",
    "AREA": "25.5",
    "APPARTMENT_NUM": "A1",
    "HEIGHT2": ""
}

// Tel-Aviv (includes ID field with <AreaNumber> placeholder support)
{
    "ID": "",              // Can use <AreaNumber> placeholder
    "APARTMENT": "1",
    "HETER": "1",
    "HEIGHT": "2.80"
}
```
**Plus shared parameters:**
- `Usage Type` (parameter, not JSON)
- `Usage Type Prev` (parameter, not JSON)

**Note on Usage Type "0" Values:**
- The `format_usage_type()` helper converts "0" values to empty strings
- Applied to all municipalities: Common (usage_type, usage_type_old), Jerusalem (code, demolition_source_code), Tel-Aviv (code, code_before)

---

## 6. Placeholder System

### Overview

Placeholders are special string values (e.g., `<View Name>`, `<AreaNumber>`) that are dynamically resolved at export time. They allow users to reference Revit element properties, project data, and coordinate information without hardcoding values.

### Complete Placeholder List (12 Total)

**Basic Placeholders (3):**
- `<View Name>` - Name of the view
- `<Level Name>` - Associated level name from view
- `<Title on Sheet>` - "Title on Sheet" parameter (with fallback to level name)

**Coordinate Placeholders (5):**
- `<E/W@InternalOrigin>` - East/West (X) shared coordinate at internal origin (meters)
- `<N/S@InternalOrigin>` - North/South (Y) shared coordinate at internal origin (meters)
- `<E/W@ProjectBasePoint>` - East/West (X) shared coordinate at project base point (meters)
- `<N/S@ProjectBasePoint>` - North/South (Y) shared coordinate at project base point (meters)
- `<SharedElevation@ProjectBasePoint>` - Elevation at project base point (meters)

**Level-based Placeholders (2):**
- `<by Project Base Point>` - Level elevation relative to project base point (meters)
- `<by Shared Coordinates>` - Level elevation in shared coordinate system (meters)

**Area-specific Placeholders (1):**
- `<AreaNumber>` - Area number from Revit Area element's `Number` property
  - Returns empty string if no number is assigned
  - Used in Tel-Aviv municipality's `ID` field

**Project-level Placeholders (2):**
- `<Project Name>` - From ProjectInformation element
- `<Project Number>` - From ProjectInformation element

### Implementation

```python
def resolve_placeholder(placeholder_value, element):
    """Resolve a placeholder string to its actual value.
    
    Args:
        placeholder_value: String that may be a placeholder (e.g., "<View Name>")
        element: Revit element to extract value from (View, Sheet, Area, etc.)
        
    Returns:
        str: Resolved value, or original value if not a placeholder
    """
    if not placeholder_value or not isinstance(placeholder_value, str):
        return placeholder_value or ""
    
    # Not a placeholder - return as-is
    if not (placeholder_value.startswith("<") and placeholder_value.endswith(">")):
        return placeholder_value
    
    # Resolve based on placeholder type
    # ... implementation for each placeholder type ...
```

### Helper Functions (Section 3: Data Extraction)

- `get_shared_coordinates(point)` - Transforms any point from project to shared coordinates, returns `(x_meters, y_meters, z_meters)`
- `get_project_base_point()` - Gets Project Base Point element
- `format_meters(value)` - Formats values to 2 decimal places

**Note on Implementation:**
- `<by Project Base Point>`: Uses `level.Elevation` directly (already in project coordinates relative to PBP)
- `<by Shared Coordinates>`: Creates point `DB.XYZ(0, 0, level.Elevation)` and transforms via `get_shared_coordinates()` to get shared Z coordinate

### Usage Examples

**Jerusalem Sheet Configuration:**
```json
{
    "PROJECT": "<Project Name>",
    "ELEVATION": "<by Project Base Point>",
    "X": "<E/W@ProjectBasePoint>",
    "Y": "<N/S@ProjectBasePoint>"
}
```

**Tel-Aviv Area Configuration:**
```json
{
    "ID": "<AreaNumber>",
    "APARTMENT": "1",
    "HETER": "1"
}
```

**Common AreaPlan Configuration:**
```json
{
    "FLOOR": "<View Name>",
    "LEVEL_ELEVATION": "<by Project Base Point>"
}
```

---

## 7. Municipality-Specific Configuration

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
- **`layer_colors`** - AutoCAD color numbers (1=Red, 3=Green, 6=Magenta, 7=White)
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

**Tel-Aviv Configuration (Updated Nov 12, 2025):**
```python
DXF_CONFIG["Tel-Aviv"] = {
    "layers": {
        'sheet_frame': 'AREA_PLAN_MAIN_FRAME',
        'sheet_text': 'AREA_PLAN_MAIN_FRAME',
        'areaplan_frame': 'muni_floor',
        'areaplan_text': 'muni_floor',
        'area_boundary': 'muni_area',
        'area_text': 'muni_area'
    },
    "layer_colors": {
        'AREA_PLAN_MAIN_FRAME': 7,    # White
        'muni_floor': 6,              # Magenta
        'muni_area': 1                # Red
    },
    "string_templates": {
        "sheet": "PAGE_NO={page_number}",
        "areaplan": "BUILDING={building}&&&FLOOR={floor}&&&HEIGHT={height}&&&X={x}&&&Y={y}&&&ABSOLUTE_HEIGHT={absolute_height}",
        "area": "CODE={code}&&&CODE_BEFORE={code_before}&&&ID={id}&&&APARTMENT={apartment}&&&HETER={heter}&&&HEIGHT={height}"
    }
}
```

**Benefits:**
- ✅ Single point of modification for all municipalities
- ✅ No code duplication between scripts
- ✅ Changes to field definitions automatically sync
- ✅ Pure Python module - works in both CPython and IronPython

---

## 8. Error Handling & Robustness

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

## 9. Testing Strategy

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

## 10. Main Orchestration Flow

**Section 9 (Main Block) Execution Order:**

1. **Get Sheets** - From selection or active view
2. **Sort Sheets** - By sheet number (descending, rightmost = page 1)
3. **⚠️ Comprehensive Validation** - Single-pass validation (CRITICAL)
   - `get_valid_areaplans_and_uniform_scale(sorted_sheets)`
   - **Phase 1:** Validates all sheets belong to same AreaScheme (hard error if missing or mixed)
   - **Phase 2:** Filters valid AreaPlan views (has municipality, has areas, has scale)
   - **Phase 3:** Validates uniform scale across all valid views
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

## 11. Implementation Phases

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
- [x] Arc bulge calculation fixes (Nov 13, 2025)
- [x] Complex curve support (splines, ellipses) via tessellation (Nov 13, 2025)
- [x] Performance optimization (cached transforms, dedup helper) (Nov 13, 2025)
- [ ] Test all municipalities
- [ ] Test multi-sheet export
- [ ] Handle edge cases

---

## 12. Implementation Insights & Lessons Learned

**Completed:** November 7, 2025

### Key Implementation Decisions

1. **Unified Validation Architecture** ⚠️ **CRITICAL**
   - **Decision:** Single comprehensive validation pass in three phases
   - **Rationale:** Avoid redundancy; fail-fast with complete error reporting; clear separation between validation and processing
   - **Implementation:** 
     - `get_valid_areaplans_and_uniform_scale()` performs all validation in single function
     - **Phase 1:** Validate uniform AreaScheme (hard error if sheets missing AreaSchemeId or belong to different schemes)
     - **Phase 2:** Filter views by: has municipality, has areas, has scale
     - **Phase 3:** Validate uniform scale across filtered views
     - Returns: `(uniform_scale, {sheet.Id: [valid_viewports]})`
     - `process_sheet()` receives both scale and pre-validated viewports
   - **Architecture:** Validation at orchestration level, processing is pure transformation
   - **Benefit:** No validation logic duplication; comprehensive upfront error reporting; single source of truth

2. **Schema Constants Import**
   - **Decision:** Import `SCHEMA_GUID`, `SCHEMA_NAME`, `FIELD_NAME` from `schema_guids.py`
   - **Rationale:** Single source of truth; prevents hardcoded value drift
   - **Updated:** Both script and plan to use imports instead of hardcoding

3. **Coordinate Transformation Order**
   - **Decision:** Apply offset BEFORE scaling: `(point - offset) * scale`
   - **Rationale:** Correct origin placement for multi-sheet layout
   - **Formula:** Subtract sheet minimum from point, then scale to real-world
   - **Implementation:** All coordinate transformations use this pattern consistently

4. **Arc and Complex Curve Handling** (Updated Nov 13, 2025)
   - **Arc Bulge Calculation:**
     - Uses arc center point and tessellated mid-point for accurate direction detection
     - Tests both CCW and CW directions to determine correct orientation
     - Formula: `bulge = tan(included_angle / 4)`
     - Implementation: `calculate_arc_bulge(start_pt, end_pt, center_pt, mid_pt)`
   - **Complex Curve Support:**
     - Splines, ellipses, NURBS curves tessellated into line segments
     - Curve types: `HermiteSpline`, `NurbSpline`, `Ellipse`, `CylindricalHelix`
     - Uses Revit's `curve.Tessellate()` for accurate point generation
   - **DXF Polyline Creation:**
     - Uses `format='xyseb'` for per-vertex bulge values
     - Format: `(x, y, start_width, end_width, bulge)`
     - Bulges applied at creation time, not post-processed
   - **Fallback:** Graceful degradation to straight line if calculation fails

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

8. **Comprehensive Single-Pass Validation** ⚠️ **CRITICAL**
   - **Decision:** Single validation function with three phases (AreaScheme + views + scale)
   - **Rationale:** Avoid validation redundancy; fail-fast with complete validation errors; clear separation of validation vs processing
   - **Implementation:** `get_valid_areaplans_and_uniform_scale()` performs all validation:
     - **Phase 1: AreaScheme Uniformity**
       - Validates all sheets have AreaSchemeId (hard error if missing)
       - Validates all sheets belong to same AreaScheme (hard error if mixed)
       - Logs AreaScheme name and ID on success
     - **Phase 2: Valid AreaPlan Filtering**
       - Must belong to an AreaScheme with defined municipality
       - Must contain areas (not empty)
       - Must have a defined scale
     - **Phase 3: Uniform Scale Validation**
       - All valid views must have same scale
       - Returns: `(uniform_scale, {sheet.Id: [valid_viewports]})`
   - **Error Handling:** 
     - Sheets without AreaSchemeId → hard error with sheet list
     - Multiple AreaSchemes → hard error with grouped sheet lists per scheme
     - No valid views found → detailed error
     - Mixed scales detected → detailed error with sheet/view list
   - **Processing:** Only pre-validated viewports are processed (no redundant checks)
   - **User Experience:** Clear, comprehensive error messages showing all validation failures upfront

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

# External (auto-downloaded to pyArea.tab/lib/vendor_cpython/ on first run)
import ezdxf  # DXF creation library

# .NET (via pythonnet/clr)
import clr, System
from Autodesk.Revit.DB.ExtensibleStorage import Schema
```

### Script Statistics (Updated Nov 13, 2025)

- **Total Lines:** ~1,820 lines
- **Sections:** 9 clearly marked sections
- **Functions:** 30+ functions (includes helper functions for placeholders, formatting, and coordinate extraction)
- **Error Handlers:** Every function has try/except with meaningful messages
- **Comments:** ~15% of lines are documentation/comments
- **Key Optimizations:**
  - Cached transforms reduce overhead per vertex
  - Dedup helper prevents duplicate vertices
  - Unified curve tessellation simplifies logic

### DXF Configuration

- **DXF Version:** R2010 (AutoCAD 2010 format - widely compatible)
- **Units:** `$INSUNITS = 5` (centimeters)
- **Coordinate System:** Real-world centimeters in modelspace
- **Polylines:** All set to closed with `polyline.closed = True`
- **Text Height:** 10.0 cm (default for all text entities)

### Corrections & Refinements

**Post-Implementation Updates (Nov 7-8, 2025):**
- ✅ Fixed filename convention to use model name and sheet range
  - Single validation function: `get_valid_areaplans_and_uniform_scale()`
  - **Phase 1:** Validates uniform AreaScheme (hard error for missing/mixed)
  - **Phase 2:** Filters valid AreaPlan views (has municipality, has areas, has scale)
  - **Phase 3:** Validates uniform scale across all valid views
  - Returns both scale and pre-validated viewports map
  - Separation: validation at orchestration, processing is pure transformation
  - Benefit: Single source of truth, fail-fast, comprehensive error reporting
- **Placeholder System Enhancements (Nov 9-12, 2025):**
  - Added `<AreaNumber>` placeholder for Tel-Aviv ID field
  - Extracts area number from Revit Area element's Number property
  - Returns empty string if area has no number assigned
  - Total placeholders supported: 12 (Basic: 3, Coordinates: 5, Level-based: 2, Area-specific: 1, Project-level: 2)
- ✅ **Tel-Aviv Municipality Schema Updates (Nov 12, 2025):**
  - Added `BUILDING` field to AreaPlan (default: "1")
  - Added `ID` field to Area with `<AreaNumber>` placeholder support (default: "")
  - Updated `APARTMENT` field to required with default "1"
  - Updated DXF layers: AREA_PLAN_MAIN_FRAME (white), muni_floor (magenta), muni_area (red)
  - Updated string templates for areaplan and area to include new fields
- ✅ **Usage Type "0" Handling (Nov 12, 2025):**
  - Added `format_usage_type()` helper function
  - Converts "0" values to empty strings for all usage type fields
  - Applied to all municipalities: Common (usage_type, usage_type_old), Jerusalem (code, demolition_source_code), Tel-Aviv (code, code_before)
  - Implementation follows naming convention: `format_*` pattern
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
- ✅ **Arc Export Fix (Nov 13, 2025):**
  - Implemented robust arc bulge calculation from Fa.extension reference
  - Uses arc center point + tessellated mid-point for direction detection
  - Tests both CCW/CW directions to match actual arc geometry
  - Fixed polyline creation to use `format='xyseb'` with bulges at creation time
  - Added support for complex curves (splines, ellipses) via tessellation
- ✅ **Performance Optimizations (Nov 13, 2025):**
  - **Cached transforms:** Precomputes `model_to_proj` and `proj_to_sheet` once per viewport
  - **Dedup helper:** `_append_pt()` prevents duplicate consecutive vertices (tolerance 1e-9)
  - **Unified tessellation:** All non-arc curves processed uniformly through single code path
  - **Reduced function calls:** Transform caching eliminates redundant matrix operations
  - Benefits: Faster processing, cleaner code, no duplicate DXF vertices

---

## 13. Constants Reference

```python
# Schema identification (imported from schema_guids.py)
from schema_guids import SCHEMA_GUID, SCHEMA_NAME, FIELD_NAME
# SCHEMA_GUID = "A7B3C9D1-E5F2-4A8B-9C3D-1E2F3A4B5C6D"
# SCHEMA_NAME = "pyArea"
# FIELD_NAME = "Data"

# Coordinate conversion constants
FEET_TO_CM = 30.48          # Revit internal units to centimeters
FEET_TO_METERS = 0.3048     # Revit internal units to meters
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
