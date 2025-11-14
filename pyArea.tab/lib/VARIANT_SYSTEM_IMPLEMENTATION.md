# Municipality Variant System Implementation

**Date:** November 12, 2025  
**Purpose:** Enable alternative usage type catalogs per municipality while maintaining single JSON/DXF schema

---

## Overview

The Variant system allows each municipality to have multiple usage type catalogs (e.g., "Common" and "CommonGross") without duplicating JSON field schemas or DXF export configurations.

### Key Design Principles

1. **Separation of Concerns**
   - **Municipality** → Controls JSON field schemas and DXF export layers/templates
   - **Variant** → Controls only which usage type CSV file is loaded

2. **Backward Compatibility**
   - Existing AreaSchemes without `Variant` field default to "Default"
   - All existing functionality continues to work unchanged

3. **Scalability**
   - Easy to add new variants: drop a CSV file and update `MUNICIPALITY_VARIANTS`
   - Jerusalem and Tel-Aviv can add variants in the future

---

## Architecture

### Data Model

```
AreaScheme JSON:
{
  "Municipality": "Common|Jerusalem|Tel-Aviv",   // Required
  "Variant": "Default|Gross|..."                  // Optional, default: "Default"
}
```

### CSV Filename Convention

- **Default variant**: `UsageType_{Municipality}.csv`  
  Example: `UsageType_Common.csv`

- **Other variants**: `UsageType_{Municipality}{Variant}.csv`  
  Example: `UsageType_CommonGross.csv`

### Available Variants

```python
MUNICIPALITY_VARIANTS = {
    "Common": ["Default", "Gross"],
    "Jerusalem": ["Default"],
    "Tel-Aviv": ["Default"]
}
```

---

## Files Modified

### 1. `lib/schemas/municipality_schemas.py`

**Changes:**
- Added `MUNICIPALITY_VARIANTS` dictionary
- Added `get_usage_type_csv_filename(municipality, variant)` helper function
- Added `Variant` field to `AREASCHEME_FIELDS`

**New Code:**
```python
MUNICIPALITY_VARIANTS = {
    "Common": ["Default", "Gross"],
    "Jerusalem": ["Default"],
    "Tel-Aviv": ["Default"]
}

def get_usage_type_csv_filename(municipality, variant="Default"):
    if variant == "Default":
        return "UsageType_{}.csv".format(municipality)
    else:
        return "UsageType_{}{}.csv".format(municipality, variant)

AREASCHEME_FIELDS = {
    "Municipality": {...},
    "Variant": {
        "type": "string",
        "required": False,
        "default": "Default",
        "description": "Usage type catalog variant",
        "hebrew_name": "גרסה"
    }
}
```

### 2. `lib/data_manager.py`

**Changes:**
- Added `get_variant(area_scheme)` method
- Added `set_variant(area_scheme, variant)` method
- Added `get_municipality_and_variant(area_scheme)` method
- Updated `get_municipality_from_view(doc, view)` to return `(municipality, variant)` tuple

**Migration Note:**
Any code calling `get_municipality_from_view()` must now handle the tuple return value.

### 3. `Define.panel/SetAreas.pushbutton/SetAreas_script.py`

**Changes:**
- Updated `load_usage_types_from_csv()` to get variant from view
- Uses `municipality_schemas.get_usage_type_csv_filename(municipality, variant)`

**Before:**
```python
municipality = data_manager.get_municipality_from_view(doc, active_view)
csv_filename = "UsageType_{}.csv".format(municipality)
```

**After:**
```python
municipality, variant = data_manager.get_municipality_from_view(doc, active_view)
csv_filename = municipality_schemas.get_usage_type_csv_filename(municipality, variant)
```

### 4. `Utilities.panel/.../CreateColorScheme.pushbutton/CreateColorScheme_script.py`

**Changes:**
- Added imports: `data_manager`, `municipality_schemas`
- Updated `read_csv_data(municipality, variant)` to accept variant parameter
- Updated `process_color_scheme()` to get variant from area_scheme

**New Code:**
```python
variant = data_manager.get_variant(area_scheme)
csv_data, csv_filename = read_csv_data(municipality, variant)
```

### 5. `Define.panel/CalculationSetup.pushbutton/CalculationSetup_script.py`

**Changes:**
- Added special handling for `Variant` field in UI field builder
- Variant dropdown populates dynamically based on Municipality
- Added `on_municipality_changed()` handler to update Variant dropdown when Municipality changes

**New UI Logic:**
- When AreaScheme is selected, both Municipality and Variant dropdowns appear
- Changing Municipality automatically updates Variant options
- If current variant not available in new municipality, resets to "Default"

### 6. `lib/JSON_TEMPLATES.md`

**Changes:**
- Updated AreaScheme section to document Variant field
- Added field details and examples
- Updated summary table to include Variant

### 7. New File: `lib/UsageType_CommonGross.csv`

**Content:**
Sample CSV demonstrating the Gross variant for Common municipality. Contains same structure as `UsageType_Common.csv` but with "ברוטו" (Gross) suffix on some category names to distinguish it.

---

## Usage

### For Users

1. **Creating an AreaScheme:**
   - In Calculation Setup, select Municipality (e.g., "Common")
   - Select Variant (e.g., "Default" or "Gross")
   - Save

2. **Setting Area Usage Types:**
   - Select areas in an AreaPlan view
   - Run SetAreas command
   - Usage type dropdown loads from the CSV matching the view's municipality + variant

3. **Creating Color Schemes:**
   - Color scheme automatically uses the variant configured on the AreaScheme
   - No user action required

### For Developers Adding New Variants

1. **Create CSV file:**
   - Name: `UsageType_{Municipality}{Variant}.csv`
   - Example: `UsageType_JerusalemNet.csv`

2. **Update `MUNICIPALITY_VARIANTS`:**
   ```python
   MUNICIPALITY_VARIANTS = {
       "Common": ["Default", "Gross"],
       "Jerusalem": ["Default", "Net"],  # ← Add "Net" here
       "Tel-Aviv": ["Default"]
   }
   ```

3. **That's it!** No changes to JSON schemas or DXF export needed.

---

## Testing Checklist

- [ ] Create new AreaScheme with Municipality="Common", Variant="Default"
- [ ] Verify UsageType_Common.csv loads in SetAreas
- [ ] Change Variant to "Gross"
- [ ] Verify UsageType_CommonGross.csv loads in SetAreas
- [ ] Change Municipality from "Common" to "Jerusalem"
- [ ] Verify Variant dropdown updates to only show "Default"
- [ ] Create color scheme and verify correct CSV is used
- [ ] Export DXF and verify no changes to export behavior (uses base municipality)
- [ ] Open existing project with old AreaSchemes (no Variant field)
- [ ] Verify they default to "Default" variant

---

## Technical Notes

### Why Not Add New Municipalities?

Adding `CommonGross` as a separate municipality would require:
- Duplicating `SHEET_FIELDS["Common"]` → `SHEET_FIELDS["CommonGross"]`
- Duplicating `AREAPLAN_FIELDS["Common"]` → `AREAPLAN_FIELDS["CommonGross"]`
- Duplicating `AREA_FIELDS["Common"]` → `AREA_FIELDS["CommonGross"]`
- Duplicating `DXF_CONFIG["Common"]` → `DXF_CONFIG["CommonGross"]`
- Updating validation logic everywhere to recognize "CommonGross"

**Result:** 5x code duplication, maintenance nightmare.

### Why Variant Field Is Better

- Single source of truth for schemas and DXF config (keyed by base municipality)
- Variant only affects usage type catalog (CSV file selection)
- Clean separation: structure vs. content
- Easy to add/remove variants without touching core logic

### Export Behavior

**Important:** DXF export uses `Municipality` field only. The `Variant` field does not affect:
- Layer names
- Layer colors
- String templates
- Field schemas

This is intentional—the variant changes which usage types are available, not how data is structured or exported.

---

## Future Enhancements

### Possible Additions

1. **Variant Description Field:**
   ```python
   MUNICIPALITY_VARIANTS = {
       "Common": [
           {"name": "Default", "description": "Standard usage types"},
           {"name": "Gross", "description": "Gross area calculation"}
       ]
   }
   ```

2. **Validation:**
   - Check CSV file exists before allowing variant selection
   - Warn if switching variants with existing areas

3. **Migration Tool:**
   - Convert old "CommonGross" municipality entries (if any exist) to "Common" + "Gross" variant

---

## Conclusion

The Variant system provides a scalable, maintainable way to support multiple usage type catalogs per municipality without duplicating schema definitions or export logic. The implementation maintains backward compatibility and follows the existing pyArea architecture patterns.
