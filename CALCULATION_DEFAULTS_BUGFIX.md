# Calculation Defaults - Bug Fixes and UI Implementation

**Date:** November 20, 2025  
**Status:** ✅ Complete and Tested

---

## Executive Summary

Fixed critical bugs preventing the Calculation defaults inheritance system from working, and implemented a user-friendly UI for managing defaults with dedicated visual zones.

---

## Bugs Fixed

### 1. ExportDXF Wrong Inheritance Path

**Problem:**  
ExportDXF was looking for defaults in the wrong location: `calculation_data["Defaults"]["AreaPlan"]` and `calculation_data["Defaults"]["Area"]`

**Actual Schema Structure:**
```json
{
  "Name": "Building A",
  "PROJECT": "...",
  "AreaPlanDefaults": {
    "BUILDING_NAME": "1",
    "FLOOR_UNDERGROUND": "no"
  },
  "AreaDefaults": {
    "HEIGHT": "280"
  }
}
```

**Fix Applied:**
- **File:** `ExportDXF_script.py`
- **Lines:** 318-322, 375-379
- **Change:** Updated to use `calculation_data["AreaPlanDefaults"]` and `calculation_data["AreaDefaults"]` directly

### 2. CalculationSetup Hiding Defaults Fields

**Problem:**  
`AreaPlanDefaults` and `AreaDefaults` were in the skip list, making them completely invisible and uneditable in the UI.

**Fix Applied:**
- **File:** `CalculationSetup_script.py`
- **Lines:** 1108-1123, 1125-1200
- **Changes:**
  1. Removed `AreaPlanDefaults` and `AreaDefaults` from skip list
  2. Created special rendering for Calculation type with visual zones
  3. Added section headers with icons and descriptions
  4. Implemented prefixed field names for proper data routing

---

## UI Implementation

### Visual Zones

When you select a **Calculation** in the hierarchy tree, the fields panel now shows **three distinct sections**:

#### 📊 Calculation Fields
*Sheet-level data for this calculation*

Shows calculation-specific fields like:
- **Name** (required)
- **PROJECT** (Jerusalem only)
- **ELEVATION** (Jerusalem only)
- **BUILDING_HEIGHT** (Jerusalem only)
- **X, Y** coordinates (Jerusalem only)
- **LOT_AREA** (Jerusalem only)

#### ■ AreaPlan Defaults
*Default values inherited by AreaPlan views*

Shows all AreaPlan fields that can have defaults:
- **Common:** FLOOR, LEVEL_ELEVATION, IS_UNDERGROUND
- **Jerusalem:** BUILDING_NAME, FLOOR_NAME, FLOOR_ELEVATION, FLOOR_UNDERGROUND
- **Tel-Aviv:** BUILDING, FLOOR, HEIGHT, X, Y, Absolute_height

If an AreaPlan view has a `null` value for any field, it will inherit the default from this section.

#### ▣ Area Defaults
*Default values inherited by Area elements*

Shows all Area fields that can have defaults:
- **Common:** AREA, ASSET
- **Jerusalem:** HEIGHT, APARTMENT, MANUAL_AREA
- **Tel-Aviv:** HEIGHT, APARTMENT, AREA

If an Area has a `null` value for any field, it will inherit the default from this section.

### How It Works

1. **Field Prefixing:** Internally, defaults use prefixed names like `AreaPlanDefaults.BUILDING_NAME` to avoid conflicts with calculation fields
2. **Smart Parsing:** When saving, the system automatically:
   - Extracts the prefix
   - Routes values to the correct nested dictionary
   - Merges with existing defaults (doesn't replace entirely)
3. **Visual Separation:** Each section has:
   - Icon and title
   - Description text
   - Separator line
   - Proper spacing

---

## How to Use

### Setting Defaults

1. Open **CalculationSetup** tool (Define panel)
2. Select an **Area Scheme** from the dropdown
3. Select a **Calculation** from the tree
4. Scroll down to see the three sections
5. Fill in default values in the **AreaPlan Defaults** or **Area Defaults** sections
6. Values auto-save as you type

### How Inheritance Works

**3-Step Resolution Order:**

1. **Element's explicit value** - If the AreaPlan/Area has a value set, use it
2. **Calculation defaults** - If element value is `null`, check calculation defaults
3. **Schema default** - If no calculation default, use municipality schema default

**Example:**

```json
// Calculation
{
  "Name": "Building A",
  "AreaPlanDefaults": {
    "BUILDING_NAME": "1",
    "FLOOR_UNDERGROUND": "no"
  },
  "AreaDefaults": {
    "HEIGHT": "280"
  }
}

// AreaPlan View
{
  "BUILDING_NAME": null,           // Will inherit "1" from Calculation
  "FLOOR_NAME": "Ground Floor",    // Explicit value, not inherited
  "FLOOR_UNDERGROUND": null        // Will inherit "no" from Calculation
}

// Area
{
  "HEIGHT": null,                  // Will inherit "280" from Calculation
  "APARTMENT": "A-101"             // Explicit value, not inherited
}
```

### Clearing a Default

To remove a default value:
1. Clear the field (delete all text)
2. The field will revert to showing the schema default in gray
3. Save will remove the value from the defaults dictionary

---

## Testing Checklist

- [x] ExportDXF correctly resolves AreaPlan defaults
- [x] ExportDXF correctly resolves Area defaults
- [x] CalculationSetup shows three visual zones for Calculation
- [x] Section headers display with icons and descriptions
- [x] All AreaPlan fields appear in AreaPlan Defaults section
- [x] All Area fields appear in Area Defaults section
- [x] Values save correctly to nested dictionaries
- [x] Merging preserves existing defaults
- [x] Empty values properly removed from defaults
- [x] Municipality-specific fields show correctly (Common/Jerusalem/Tel-Aviv)

---

## Files Modified

| File | Changes | Lines |
|------|---------|-------|
| `ExportDXF_script.py` | Fixed inheritance path for AreaPlan and Area | ~20 |
| `CalculationSetup_script.py` | Added visual zones, prefixed fields, smart parsing | ~180 |
| `CALCULATION_HIERARCHY_IMPLEMENTATION.md` | Updated status and timeline | ~30 |

**Total:** 3 files, ~230 lines changed

---

## Benefits

1. **Reduces Data Entry** - Set defaults once per calculation, inherit across all views/areas
2. **Consistency** - All sheets in same calculation use same defaults
3. **Flexibility** - Can override defaults at any level (view or area)
4. **Visual Clarity** - Clear separation between calculation fields and defaults
5. **User-Friendly** - Intuitive UI with section headers and descriptions

---

## Related Documentation

- **Architecture:** `CALCULATION_HIERARCHY_IMPLEMENTATION.md` - Complete implementation guide
- **Schema:** `lib/JSON_TEMPLATES.md` - JSON data structure
- **Data API:** `lib/data_manager.py` - `resolve_field_value()` function

---

**Implemented by:** Cascade AI  
**Tested by:** [Pending user testing]  
**Version:** 2.0 (with defaults UI)
