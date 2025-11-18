# Calculation Hierarchy Implementation Plan

**Date:** November 14, 2025  
**Status:** DRAFT - Awaiting approval before implementation

---

## Executive Summary

This plan introduces a new **Calculation** level in the data hierarchy between **AreaScheme** and **Sheet**. The Calculation acts as a template/container that holds sheet metadata and default values for views and areas, allowing multiple sheets to share the same calculation settings.

### Current Hierarchy
```
AreaScheme (Municipality, Variant)
└── Sheet (AreaSchemeId, PROJECT, ELEVATION, X, Y, etc. + page_number)
    └── AreaPlan views (FLOOR, BUILDING_NAME, etc.)
        └── Areas (HEIGHT, APARTMENT, etc.)
```

### Proposed Hierarchy
```
AreaScheme (Municipality, Variant)
└── Calculation (AreaSchemeId, PROJECT, ELEVATION, X, Y, + defaults for views/areas)
    └── Sheet (CalculationId, page_number only)
        └── AreaPlan views (inherits/overrides Calculation defaults)
            └── Areas (inherits/overrides Calculation defaults)
```

**Key Benefits:**
1. **Reduced Data Duplication:** Sheet-level fields stored once per Calculation
2. **Consistency:** All sheets in same Calculation share identical metadata
3. **Bulk Operations:** Change settings for multiple sheets at once
4. **Template System:** Calculation provides default values for AreaPlans/Areas

---

## 1. Data Storage Strategy

### Calculation Storage (UPDATED: AreaScheme, not ProjectInformation)

**Why AreaScheme is Better:**
- **Data locality**: Calculations naturally belong to their parent AreaScheme
- **No redundant AreaSchemeId**: Eliminates field that can drift out of sync
- **Cleaner hierarchy**: AreaScheme explicitly owns its Calculations
- **Simpler lookups**: Direct access via AreaScheme element

**Structure (stored on each AreaScheme element):**
```python
# Calculation fields by municipality
CALCULATION_FIELDS = {
    "Common": {
        "CalculationGuid": {
            "type": "string",
            "required": True,
            "description": "System-generated GUID identifier"
        },
        "Name": {
            "type": "string",
            "required": True,
            "description": "User-facing calculation name"
        },
        "Defaults": {
            "type": "dict",
            "required": False,
            "description": "Default values for AreaPlan and Area"
        }
    },
    "Jerusalem": {
        "PROJECT": {
            "type": "string",
            "required": True,
            "description": "Project name"
        },
        "ELEVATION": {
            "type": "string",
            "required": True,
            "description": "Elevation"
        },
        "BUILDING_HEIGHT": {
            "type": "number",
            "required": True,
            "description": "Building height"
        },
        "X": {
            "type": "string",
            "required": True,
            "description": "X coordinate"
        },
        "Y": {
            "type": "string",
            "required": True,
            "description": "Y coordinate"
        },
        "LOT_AREA": {
            "type": "number",
            "required": True,
            "description": "Lot area"
        }
    },
    "Tel-Aviv": {
        "BUILDING": {
            "type": "string",
            "required": True,
            "description": "Building"
        },
        "HEIGHT": {
            "type": "number",
            "required": True,
            "description": "Height"
        },
        "X": {
            "type": "string",
            "required": True,
            "description": "X coordinate"
        },
        "Y": {
            "type": "string",
            "required": True,
            "description": "Y coordinate"
        }
    }
}
```

**Example Calculation Data:**
```json
{
  "Municipality": "Jerusalem",
  "Variant": "Default",
  "Calculations": {
    "a7b3c9d1-e5f2-4a8b-9c3d-1e2f3a4b5c6d": {
      "CalculationGuid": "a7b3c9d1-e5f2-4a8b-9c3d-1e2f3a4b5c6d",
      "Name": "Building A - Standard",
      "PROJECT": "...",
      "ELEVATION": "...",
      "Defaults": {
        "AreaPlan": {...},
        "Area": {...}
      }
    },
    "b8c4d0e2-f6g3-5b9c-0d4e-2f3g4a5b6c7d": {
      "CalculationGuid": "b8c4d0e2-f6g3-5b9c-0d4e-2f3g4a5b6c7d",
      "Name": "Building B - Variant",
      "PROJECT": "...",
      "ELEVATION": "...",
      "Defaults": {...}
    }
  }
}
```

**ProjectInformation Usage (Global Settings Only):**
```json
{
  "SchemaVersion": "2.0",
  "Preferences": {...}
}
```

**Identity Strategy: GUID + Name**
- **CalculationGuid**: System-generated GUID (immutable, collision-free)
- **Name**: User-facing display name (editable, for UI/reports)
- Sheets store `CalculationGuid` to reference parent Calculation

---

## 2. JSON Schema Changes

### 2.1 Calculation Schema (NEW)

**Common Municipality:**
```json
{
  "CalculationGuid": "a7b3c9d1-e5f2-4a8b-9c3d-1e2f3a4b5c6d",
  "Name": "Standard Calculation",
  "Defaults": {
    "AreaPlan": {},
    "Area": {}
  }
}
```

**Jerusalem Municipality:**
```json
{
  "CalculationGuid": "a7b3c9d1-e5f2-4a8b-9c3d-1e2f3a4b5c6d",
  "Name": "Building A",
  "PROJECT": "<Project Name>",
  "ELEVATION": "<SharedElevation@ProjectBasePoint>",
  "BUILDING_HEIGHT": "30.5",
  "X": "<E/W@InternalOrigin>",
  "Y": "<N/S@InternalOrigin>",
  "LOT_AREA": "5000",
  "Defaults": {
    "AreaPlan": {
      "BUILDING_NAME": "1",
      "FLOOR_UNDERGROUND": "no"
    },
    "Area": {
      "HEIGHT": "280"
    }
  }
}
```

**Tel-Aviv Municipality:**
```json
{
  "CalculationGuid": "a7b3c9d1-e5f2-4a8b-9c3d-1e2f3a4b5c6d",
  "Name": "Standard Setup",
  "Defaults": {
    "AreaPlan": {
      "BUILDING": "1",
      "HEIGHT": "280",
      "X": "<E/W@InternalOrigin>",
      "Y": "<N/S@InternalOrigin>"
    },
    "Area": {
      "HETER": "1"
    }
  }
}
```

**Note:** No `AreaSchemeId` field needed - Calculation is stored ON the AreaScheme

### 2.2 Sheet Schema (SIMPLIFIED)

**All Municipalities:**
```json
{
  "CalculationGuid": "a7b3c9d1-e5f2-4a8b-9c3d-1e2f3a4b5c6d"
}
```

**Note:** `page_number` calculated at export time, not stored.

### 2.3 AreaPlan/Area (With Inheritance)

**AreaPlan Example (Jerusalem):**
```json
{
  "BUILDING_NAME": "2",           // Override
  "FLOOR_NAME": "<Level Name>",   // Custom
  "FLOOR_ELEVATION": "<by Project Base Point>",
  "FLOOR_UNDERGROUND": null,      // null = inherit from Calculation default "no"
  "RepresentedViews": []
}
```

**Resolution Order:**
1. Element's explicit value (if not `None`)
2. Calculation's default for element type
3. Field schema default (from municipality_schemas.py)

**Critical: Null-Friendly Validation**
- Current `validate_data()` rejects `None` for typed fields
- Must update to allow `None` as "inherit default" marker
- Without this fix, inheritance pattern won't work

---

## 3. File Changes Overview

### 3.1 `municipality_schemas.py`

**New Additions:**
```python
# New constant
CALCULATION_FIELDS = {
    "Common": {...},
    "Jerusalem": {...},
    "Tel-Aviv": {...}
}

# Updated function
def get_fields_for_element_type(element_type, municipality=None):
    # Now supports "Calculation" as element_type
    field_map = {
        "Calculation": CALCULATION_FIELDS,  # NEW
        "Sheet": SHEET_FIELDS,
        "AreaPlan": AREAPLAN_FIELDS,
        "Area": AREA_FIELDS
    }
```

**Modified:**
```python
# SHEET_FIELDS simplified to just CalculationId
SHEET_FIELDS = {
    "Common": {"CalculationId": {...}},
    "Jerusalem": {"CalculationId": {...}},
    "Tel-Aviv": {"CalculationId": {...}}
}
```

**File Location:** `pyArea.tab/lib/schemas/municipality_schemas.py`

---

### 3.2 `data_manager.py`

**New Functions:**
```python
def get_all_calculations(doc)
def get_calculation(doc, calculation_id)
def set_calculation(doc, calculation_id, calculation_data, municipality)
def delete_calculation(doc, calculation_id)
def get_calculation_from_sheet(doc, sheet)
def resolve_field_value(field_name, element_data, calculation_data, municipality, element_type)
```

**Modified Functions:**
```python
def set_sheet_data(sheet, calculation_id)  # Signature changed
```

**File Location:** `pyArea.tab/lib/data_manager.py`

---

### 3.3 `JSON_TEMPLATES.md`

**Updates:**
- Add new "Calculation" section
- Update Sheet section (simplified to CalculationId only)
- Document inheritance/override pattern for AreaPlan/Area
- Update summary table

**File Location:** `pyArea.tab/lib/JSON_TEMPLATES.md`

---

## 4. ExportDXF_script.py Changes

### 4.1 New Helper Functions

**Data Resolution with Inheritance:**
```python
def resolve_field_value(field_name, element_data, calculation_data, municipality, element_type):
    """Resolve field with Calculation inheritance."""
    # 1. Element's explicit value
    if field_name in element_data and element_data[field_name] is not None:
        return element_data[field_name]
    
    # 2. Calculation default
    if calculation_data:
        defaults = calculation_data.get("Defaults", {}).get(element_type, {})
        if field_name in defaults:
            return defaults[field_name]
    
    # 3. Field schema default
    fields = get_fields_for_element_type(element_type, municipality)
    field_def = fields.get(field_name, {})
    return field_def.get("default")
```

### 4.2 Modified Data Extraction

**get_sheet_data_for_dxf() - BEFORE:**
```python
def get_sheet_data_for_dxf(sheet_elem):
    sheet_json = get_json_data(sheet_elem)
    area_scheme_id = sheet_json.get("AreaSchemeId")
    # ... get municipality from AreaScheme
    return {**sheet_json, "Municipality": municipality}
```

**get_sheet_data_for_dxf() - AFTER:**
```python
def get_sheet_data_for_dxf(sheet_elem):
    """Get combined sheet + calculation data."""
    # Get CalculationId from sheet
    sheet_json = get_json_data(sheet_elem)
    calculation_id = sheet_json.get("CalculationId")
    
    # Get Calculation from ProjectInformation
    proj_info = doc.ProjectInformation
    proj_data = get_json_data(proj_info)
    calculation_data = proj_data.get("Calculations", {}).get(calculation_id)
    
    # Get Municipality from AreaScheme
    area_scheme_id = calculation_data.get("AreaSchemeId")
    area_scheme = doc.GetElement(create_element_id(int(area_scheme_id)))
    municipality = get_json_data(area_scheme).get("Municipality")
    
    # Return combined data
    return {**calculation_data, "Municipality": municipality}
```

### 4.3 Modified AreaPlan/Area Extraction

**BEFORE:**
```python
def get_areaplan_data_for_dxf(view_elem):
    return get_json_data(view_elem)

def get_area_data_for_dxf(area_elem):
    return get_json_data(area_elem)
```

**AFTER:**
```python
def get_areaplan_data_for_dxf(view_elem, calculation_data, municipality):
    """Get AreaPlan data with Calculation defaults applied."""
    areaplan_json = get_json_data(view_elem)
    fields = AREAPLAN_FIELDS.get(municipality, {})
    
    resolved_data = {}
    for field_name in fields.keys():
        resolved_data[field_name] = resolve_field_value(
            field_name, areaplan_json, calculation_data, municipality, "AreaPlan"
        )
    
    return resolved_data

def get_area_data_for_dxf(area_elem, calculation_data, municipality):
    """Get Area data with Calculation defaults applied."""
    area_json = get_json_data(area_elem)
    fields = AREA_FIELDS.get(municipality, {})
    
    resolved_data = {}
    for field_name in fields.keys():
        resolved_data[field_name] = resolve_field_value(
            field_name, area_json, calculation_data, municipality, "Area"
        )
    
    # Add shared parameters
    resolved_data["UsageType"] = get_parameter_value(area_elem, "Usage Type")
    resolved_data["UsageTypePrev"] = get_parameter_value(area_elem, "Usage Type Prev")
    
    return resolved_data
```

### 4.4 Processing Flow Changes

**process_areaplan_viewport() signature:**
```python
# BEFORE
def process_areaplan_viewport(viewport, msp, scale_factor, offset_x, offset_y, municipality, layers):

# AFTER
def process_areaplan_viewport(viewport, msp, scale_factor, offset_x, offset_y, municipality, calculation_data, layers):
```

**process_sheet() updates:**
```python
def process_sheet(sheet_elem, dxf_doc, msp, horizontal_offset, page_number, view_scale, valid_viewports):
    # Get combined sheet + calculation data
    sheet_data = get_sheet_data_for_dxf(sheet_elem)
    municipality = sheet_data.get("Municipality")
    calculation_data = sheet_data  # It IS the calculation data
    
    # Process viewports with calculation_data
    for viewport in valid_viewports:
        process_areaplan_viewport(
            viewport, msp, scale_factor, offset_x, offset_y, 
            municipality, calculation_data, layers  # Pass calculation_data
        )
```

### 4.5 Validation Changes

**get_valid_areaplans_and_uniform_scale():**
```python
# Check sheet has CalculationId
calculation_id = sheet_data.get("CalculationId")
if not calculation_id:
    continue

# Validate Calculation exists
proj_info = doc.ProjectInformation
proj_data = get_json_data(proj_info)
calculation_data = proj_data.get("Calculations", {}).get(calculation_id)
if not calculation_data:
    print("Warning: Calculation {} not found".format(calculation_id))
    continue

# Validate AreaScheme exists
area_scheme_id = calculation_data.get("AreaSchemeId")
area_scheme = doc.GetElement(create_element_id(int(area_scheme_id)))
if not area_scheme:
    continue
```

**File Location:** `pyArea.tab/Export.panel/ExportDXF.pushbutton/ExportDXF_script.py`

---

## 5. CalculationSetup_script.py Changes

### 5.1 New Workflow

**Current:**
1. Select AreaScheme
2. Fill Sheet fields
3. Fill AreaPlan defaults
4. Select Sheets → Save to each Sheet

**Proposed:**
1. Select AreaScheme
2. **Select or Create Calculation** (NEW)
3. Fill Calculation fields (PROJECT, ELEVATION, etc.)
4. Fill AreaPlan/Area defaults
5. Select Sheets to assign to this Calculation
6. Save: Calculation → ProjectInformation, Sheets → CalculationId

### 5.2 UI Changes

**New Controls:**
```python
# Calculation selection
calculation_dropdown = ComboBox()
calculation_dropdown.Items = ["<New Calculation>"] + list(existing_calculation_ids)

# Calculation ID input (for new calculations)
calculation_id_textbox = TextBox()
calculation_id_textbox.Visibility = Collapsed  # Show only if "New" selected

# Sheet assignment section
sheet_checklist = CheckListBox()
sheet_checklist.Items = all_sheets_for_areascheme
```

**Simplified Data Entry:**
- Calculation fields (PROJECT, ELEVATION, etc.) → entered once
- Defaults section for AreaPlan/Area fields → stored in Calculation
- Multiple sheets can reference same Calculation

### 5.3 Implementation Notes

- Will be detailed in separate follow-up plan after core architecture approved
- Requires significant UI/UX redesign
- Should include Calculation management (list, edit, delete existing Calculations)

**File Location:** `pyArea.tab/Define.panel/CalculationSetup.pushbutton/CalculationSetup_script.py`

---

## 6. Migration Strategy

### 6.1 Auto-Migration Script

**Purpose:** Convert existing Sheet data to Calculation hierarchy

**Logic:**
1. Detect old schema (Sheet has PROJECT, ELEVATION, etc.)
2. Group Sheets with identical field values
3. Create Calculation for each unique group
4. Update Sheets to reference CalculationId
5. Mark migration complete (SchemaVersion = "2.0")

**Pseudo-code:**
```python
def migrate_to_calculation_hierarchy(doc):
    # 1. Find all sheets with old schema
    old_sheets = [s for s in all_sheets if "PROJECT" in get_json_data(s)]
    
    # 2. Group by AreaScheme + field values
    groups = group_sheets_by_fields(old_sheets)
    
    # 3. Create Calculations
    for idx, (field_values, sheets) in enumerate(groups):
        calc_id = "CALC_{:03d}".format(idx + 1)
        calculation_data = {
            "CalculationId": calc_id,
            **field_values
        }
        set_calculation(doc, calc_id, calculation_data, municipality)
        
        # 4. Update each Sheet
        for sheet in sheets:
            set_sheet_data(sheet, calc_id)
    
    # 5. Mark migration complete
    proj_info = doc.ProjectInformation
    data = get_json_data(proj_info)
    data["SchemaVersion"] = "2.0"
    set_json_data(proj_info, data)
```

**When to run:**
- Automatically on first script execution in project with old schema
- Or manual "Migrate to v2.0" button in CalculationSetup

### 6.2 Backward Compatibility Detection

```python
def get_schema_version(doc):
    proj_info = doc.ProjectInformation
    data = get_json_data(proj_info)
    return data.get("SchemaVersion", "1.0")

def needs_migration(doc):
    return get_schema_version(doc) == "1.0"
```

---

## 7. Implementation Phases

### Phase 1: Core Architecture (Week 1)
- [ ] Update `municipality_schemas.py` (CALCULATION_FIELDS, simplified SHEET_FIELDS)
- [ ] Update `data_manager.py` (Calculation CRUD methods, resolve_field_value)
- [ ] Update `JSON_TEMPLATES.md` documentation
- [ ] Create migration script (as standalone tool)
- [ ] Test migration on sample projects

### Phase 2: ExportDXF Integration (Week 1-2)
- [ ] Modify data extraction functions
- [ ] Update processing pipeline
- [ ] Update validation logic
- [ ] Test export with migrated data
- [ ] Test inheritance/override patterns

### Phase 3: CalculationSetup UI (Week 2-3)
- [ ] Design new workflow and UI mockup
- [ ] Implement Calculation selection/creation
- [ ] Implement Calculation field editing
- [ ] Implement Sheet assignment
- [ ] Implement Calculation management (list/edit/delete)
- [ ] Test end-to-end workflow

### Phase 4: Testing & Documentation (Week 3-4)
- [ ] Comprehensive testing across all municipalities
- [ ] Test migration edge cases
- [ ] Update user documentation
- [ ] Create training materials
- [ ] Beta testing with users

---

## 8. Open Questions

1. **Calculation Naming:**
   - GUID is auto-generated (system identifier)
   - Name is user-editable (display label)
   - **Decision:** ✅ GUID + Name pattern adopted

2. **Calculation UI:**
   - Separate dialog for Calculation management?
   - Or integrated into CalculationSetup main dialog?
   - **Recommendation:** Integrated with expandable sections

3. **Default Propagation:**
   - Should changing Calculation defaults auto-update existing AreaPlans/Areas?
   - **Recommendation:** No, preserve explicit overrides. Show warning instead.

4. **Multi-AreaScheme Calculations:**
   - Can one Calculation span multiple AreaSchemes?
   - **Recommendation:** No, Calculation is always 1:1 with AreaScheme

5. **Calculation Deletion:**
   - What happens to Sheets when Calculation is deleted?
   - **Recommendation:** Prevent deletion if Sheets reference it, or orphan Sheets

---

## 9. Design Summary

- **Storage location**: Calculations are stored on their parent AreaScheme element.
- **Identity strategy**: Each Calculation has a system GUID (`CalculationGuid`) and a user-editable `Name`.
- **Sheet data**: Sheets store only a `CalculationGuid` reference; PAGE_NO is derived from sheet order at export.
- **Defaults and overrides**: AreaPlans and Areas can override Calculation defaults per field; `None` means "inherit".
- **Validation**: `validate_data()` allows `None` for typed fields to support inheritance.

## 10. Summary

This plan provides a clear path to introducing the Calculation hierarchy while maintaining:
- **Backward compatibility** (auto-migration)
- **Data integrity** (validation at all levels with null-friendly inheritance)
- **Flexibility** (inheritance with overrides where needed)
- **Clean architecture** (AreaScheme-scoped storage, GUID identity)

**Next Steps:**
1. Review and approve this plan
2. Address open questions
3. Begin Phase 1 implementation
4. Create detailed CalculationSetup UI/UX plan

---

**Questions? Feedback?** Please provide comments before proceeding to implementation.
