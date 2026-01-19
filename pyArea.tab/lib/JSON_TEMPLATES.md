# PyArea JSON Data Templates for DXF Export

**Source:** DXF attributes.xlsx  
**Date:** November 2, 2025  
**Last Updated:** November 20, 2025 (Calculation hierarchy + DWFX underlay override field)

This document defines the JSON structure for each element type (AreaScheme, Calculation, Sheet, AreaPlan/View, Area) across all municipalities (Common, Jerusalem, Tel-Aviv).

---

## 1. AreaScheme (Municipality Definition)

**All Municipalities:**
```json
{
  "Municipality": "Common|Jerusalem|Tel-Aviv",
  "Variant": "Default|Gross|..."
}
```

**Field Details:**
- `Municipality`: Base municipality type (required)
- `Variant`: Usage type catalog variant (optional, default: "Default")
  - Controls which `UsageType_{Municipality}{Variant}.csv` file is loaded
  - Available variants per municipality:
    - Common: `Default`, `Gross`
    - Jerusalem: `Default`
    - Tel-Aviv: `Default`

**Notes:**
- Stored once on the AreaScheme element
- All child elements inherit municipality and variant via relationships
- `Variant` only affects usage type CSV selection; JSON field schemas and DXF export config remain based on base `Municipality`
- Backward compatible: existing AreaSchemes without `Variant` default to "Default"
- AreaScheme also contains Calculations (see next section)

---

## 2. Calculation

**Storage:** Calculations are stored ON the AreaScheme element in a `Calculations` dictionary keyed by CalculationGuid.

### Common Municipality
```json
{
  "Name": "Standard Calculation",
  "AreaPlanDefaults": {},
  "AreaDefaults": {}
}
```

### Jerusalem Municipality
```json
{
  "Name": "Building A",
  "PROJECT": "<Project Name>",
  "ELEVATION": "<SharedElevation@ProjectBasePoint>",
  "BUILDING_HEIGHT": "30.5",
  "X": "<E/W@InternalOrigin>",
  "Y": "<N/S@InternalOrigin>",
  "LOT_AREA": "5000",
  "AreaPlanDefaults": {
    "BUILDING_NAME": "1",
    "FLOOR_UNDERGROUND": "no"
  },
  "AreaDefaults": {
    "HEIGHT": "280"
  }
}
```

### Tel-Aviv Municipality
```json
{
  "Name": "Standard Setup",
  "AreaPlanDefaults": {
    "BUILDING": "1",
    "HEIGHT": "280",
    "X": "<E/W@InternalOrigin>",
    "Y": "<N/S@InternalOrigin>"
  },
  "AreaDefaults": {
    "HETER": "1"
  }
}
```

**Field Details:**
- Calculation GUID is the KEY in the Calculations dictionary (not stored as a field)
- `Name`: User-facing display name (editable)
- `AreaPlanDefaults`: Optional default values for AreaPlan elements
  - AreaPlan elements inherit these defaults if their own values are `null`
  - Inheritance order: Element explicit value → AreaPlanDefaults → Schema default
- `AreaDefaults`: Optional default values for Area elements
  - Area elements inherit these defaults if their own values are `null`
  - Inheritance order: Element explicit value → AreaDefaults → Schema default
- Jerusalem municipality: All Sheet-level fields moved to Calculation level
- Tel-Aviv municipality: AreaPlan-level coordinate fields can be set as defaults

**Notes:**
- Multiple Calculations can exist under one AreaScheme
- Multiple Sheets can reference the same Calculation
- All Sheets referencing a Calculation share the same metadata
- `PAGE_NO` is derived from sheet order at export time (not stored)

---

## 3. Sheet

**All Municipalities (v2.0):**
```json
{
  "CalculationGuid": "a7b3c9d1-e5f2-4a8b-9c3d-1e2f3a4b5c6d",
  "DWFx_UnderlayFilename": "MyProject-A101.dwfx" // optional override
}
```

**Notes:**
- Sheets **always** store a `CalculationGuid` reference to their parent Calculation
- `DWFx_UnderlayFilename` is **optional** and used only by ExportDXF to override the DWFX underlay filename for this sheet
  - If empty or missing, ExportDXF falls back to the auto-generated `{ModelName}-{SheetNumber}.dwfx` name
- All Calculation-related metadata (PROJECT, ELEVATION, X, Y, etc.) is stored on the Calculation, not on the Sheet
- `PAGE_NO` is calculated at export time from sheet order (not stored)
- Sheet order determines page numbering in DXF export

**Legacy (Schema v1.0):** Old projects had Sheet-level fields directly on Sheet elements:

### Common Municipality (OLD - v1.0)
```json
{
  "AreaSchemeId": "<Revit ElementId>"
}
```

### Jerusalem Municipality (OLD - v1.0)
```json
{
  "AreaSchemeId": "<Revit ElementId>",
  "PROJECT": "<string>",
  "ELEVATION": <float>,
  "BUILDING_HEIGHT": <float>,
  "X": <float>,
  "Y": <float>,
  "LOT_AREA": <float>
}
```

### Tel-Aviv Municipality (OLD - v1.0)
```json
{
  "AreaSchemeId": "<Revit ElementId>"
}
```

**Migration:** Old projects will be automatically migrated to the new Calculation structure by grouping sheets with identical metadata into shared Calculations.

---

## 4. AreaPlan (View)


### Common Municipality
```json
{
  "BUILDING_NO": "<string>",
  "FLOOR": "<string>",
  "LEVEL_ELEVATION": <float>,
  "IS_UNDERGROUND": <int>,
  "RepresentedViews": ["<Revit ElementId>", "<Revit ElementId>", ...]
}
```

**Field Details:**
- `BUILDING_NO`: Building number. **Default: "1"**
- `FLOOR`: Floor name. **Default: `<View Name>`** (can also use `<Title on Sheet>`)
- `LEVEL_ELEVATION`: Level elevation in meters. **Default: `<by Project Base Point>`** (can also use `<by Shared Coordinates>`)
- `IS_UNDERGROUND`: 0 or 1
- `RepresentedViews`: List of AreaPlan ElementIds that this typical floor represents (empty list if not a typical floor)

**Inheritance:** Any field set to `null` will inherit from AreaPlanDefaults → Schema default

### Jerusalem Municipality
```json
{
  "BUILDING_NAME": "<string>",
  "FLOOR_NAME": "<string>",
  "FLOOR_ELEVATION": <float>,
  "FLOOR_UNDERGROUND": "<string>",
  "RepresentedViews": ["<Revit ElementId>", "<Revit ElementId>", ...]
}
```

**Field Details:**
- `BUILDING_NAME`: Building name (user-entered)
- `FLOOR_NAME`: Floor name. **Default: `<View Name>`** (can also use `<Title on Sheet>`)
- `FLOOR_ELEVATION`: Floor elevation in meters. **Default: `<by Project Base Point>`** (can also use `<by Shared Coordinates>`)
- `FLOOR_UNDERGROUND`: "yes" or "no"
- `RepresentedViews`: List of AreaPlan ElementIds that this typical floor represents (empty list if not a typical floor)

**Inheritance:** Any field set to `null` will inherit from AreaPlanDefaults → Schema default

### Tel-Aviv Municipality
```json
{
  "BUILDING": "<string>",
  "FLOOR": "<string>",
  "HEIGHT": <float>,
  "X": <float>,
  "Y": <float>,
  "Absolute_height": <float>,
  "RepresentedViews": ["<Revit ElementId>", "<Revit ElementId>", ...]
}
```

**Field Details:**
- `BUILDING`: Building name/identifier. **Default: "1"**
- `FLOOR`: Floor name. **Default: `<View Name>`** (can also use `<Title on Sheet>`)
- `HEIGHT`: Floor height in CM. **Default: `<Auto>`** (calculate from difference between levels, or use defined value for topmost)
- `X`: If `<E/W@ProjectBasePoint>`, get shared coordinates X (East/West) of project base point (meters). If `<E/W@InternalOrigin>`, get shared coordinates X (East/West) of internal origin (meters)
- `Y`: If `<N/S@ProjectBasePoint>`, get shared coordinates Y (North/South) of project base point (meters). If `<N/S@InternalOrigin>`, get shared coordinates Y (North/South) of internal origin (meters)
- `Absolute_height`: If `<by Project Base Point>`, use host level height from project base point (meters). If `<by Shared Coordinates>`, use host level height from shared coordinates (meters)
- `RepresentedViews`: List of AreaPlan ElementIds that this typical floor represents (empty list if not a typical floor)

**Inheritance:** Any field set to `null` will inherit from AreaPlanDefaults → Schema default

---

## 5. Area

**Note:** `USAGE_TYPE` and `CODE` are stored in shared parameters "Usage Type" and "Usage Type Prev", NOT in JSON schema.

### Common Municipality
```json
{
  "AREA": <float>,
  "ASSET": "<string>"
}
```

**Field Details:**
- `AREA`: User-entered or calculated area value
- `ASSET`: User-entered asset identifier

**Shared Parameters (not in JSON):**
- `Usage Type`: Current usage type code
- `Usage Type Prev`: Previous usage type code (exported as `USAGE_TYPE_OLD`)

**Inheritance:** Any field set to `null` will inherit from AreaDefaults → Schema default

### Jerusalem Municipality
```json
{
  "AREA": <float>,
  "HEIGHT": <float>,
  "APPARTMENT_NUM": "<string>",
  "HEIGHT2": <float>
}
```

**Field Details:**
- `AREA`: User-entered or calculated area value
- `HEIGHT`: Room/area height
- `APPARTMENT_NUM`: If `<*>` is used, get value from parameter assigned to area element
- `HEIGHT2`: Secondary height value

**Shared Parameters (not in JSON):**
- `Usage Type`: Current usage type code (exported as `CODE`)
- `Usage Type Prev`: Previous usage type code (exported as `DEMOLITION_SOURCE_CODE`)

**Inheritance:** Any field set to `null` will inherit from AreaDefaults → Schema default

### Tel-Aviv Municipality
```json
{
  "ID": "<string>",
  "APARTMENT": "<string>",
  "HETER": "<string>",
  "HEIGHT": <float>
}
```

**Field Details:**
- `ID`: Area identifier/number. **Default: ""** (empty string). Supports `<AreaNumber>` placeholder to use Revit area number
- `APARTMENT`: Apartment identifier. **Default: "1"**
- `HETER`: Permit/variance identifier. **Default: "1"**
- `HEIGHT`: Room/area height

**Shared Parameters (not in JSON):**
- `Usage Type`: Current usage type code (exported as `CODE`)
- `Usage Type Prev`: Previous usage type code (exported as `CODE_BEFORE`)

**Inheritance:** Any field set to `null` will inherit from AreaDefaults → Schema default

---

## Summary Table

| Element Type | Common Fields | Jerusalem Fields | Tel-Aviv Fields |
|--------------|---------------|------------------|-----------------|
| **AreaScheme** | Municipality, Variant, Calculations{} | Municipality, Variant, Calculations{} | Municipality, Variant, Calculations{} |
| **Calculation** | Name, AreaPlanDefaults, AreaDefaults | Name, PROJECT, ELEVATION, BUILDING_HEIGHT, X, Y, LOT_AREA, AreaPlanDefaults, AreaDefaults | Name, AreaPlanDefaults, AreaDefaults |
| **Sheet** | CalculationGuid, DWFx_UnderlayFilename (optional) | CalculationGuid, DWFx_UnderlayFilename (optional) | CalculationGuid, DWFx_UnderlayFilename (optional) |
| **AreaPlan** | BUILDING_NO, FLOOR, LEVEL_ELEVATION, IS_UNDERGROUND, RepresentedViews | BUILDING_NAME, FLOOR_NAME, FLOOR_ELEVATION, FLOOR_UNDERGROUND, RepresentedViews | BUILDING, FLOOR, HEIGHT, X, Y, Absolute_height, RepresentedViews |
| **Area** | AREA, ASSET | AREA, HEIGHT, APPARTMENT_NUM, HEIGHT2 | ID, APARTMENT, HETER, HEIGHT |

**Plus Shared Parameters (all municipalities):**
- `Usage Type` (Text)
- `Usage Type Prev` (Text)

---

## Implementation Notes

1. **Field Naming Convention:**
   - Use ALL CAPS with underscores for JSON keys (e.g., `FLOOR_NAME`, `LOT_AREA`)
   - Match DXF attribute names exactly

2. **Data Types:**
   - `<string>`: Text values
   - `<float>`: Numeric values (elevations, areas, coordinates)
   - `<int>`: Integer values (0/1 for boolean flags)

3. **Special Values (Placeholders):**
   - `<Project Name>`, `<Project Number>`: Get from Revit document properties
   - `<View Name>`: Get view name (always has a value in Revit)
   - `<Title on Sheet>`: Get "Title on Sheet" parameter (falls back to level name if empty)
   - `<Level Name>`: Get associated level name
   - `<by Project Base Point>`: Use project base point coordinate system (for elevations/heights)
   - `<by Shared Coordinates>`: Use shared coordinate system (for elevations/heights)
   - `<E/W@ProjectBasePoint>`: Get East/West (X) shared coordinate from project base point
   - `<E/W@InternalOrigin>`: Get East/West (X) shared coordinate from internal origin
   - `<N/S@ProjectBasePoint>`: Get North/South (Y) shared coordinate from project base point
   - `<N/S@InternalOrigin>`: Get North/South (Y) shared coordinate from internal origin
   - `<SharedElevation@ProjectBasePoint>`: Get elevation at project base point
   - `<AreaNumber>`: Get area number from Revit Area element (Tel-Aviv ID field)
   - `<Auto>`: Calculate automatically
   - `<*>`: Use value from specified parameter

4. **Coordinate Systems:**
   - All coordinates in meters
   - Heights in CM for Tel-Aviv view heights, meters elsewhere
   - Support both Project Base Point and Internal Origin references

5. **Validation:**
   - Required fields vary by municipality
   - Optional fields can be null or omitted
   - Refer to municipality_schemas.py for required/optional flags
