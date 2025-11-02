# PyArea JSON Data Templates for DXF Export

**Source:** DXF attributes.xlsx  
**Date:** November 2, 2025

This document defines the JSON structure for each element type (AreaScheme, Sheet, AreaPlan/View, Area) across all municipalities (Common, Jerusalem, Tel-Aviv).

---

## 1. AreaScheme (Municipality Definition)

**All Municipalities:**
```json
{
  "Municipality": "Common|Jerusalem|Tel-Aviv"
}
```

**Notes:**
- Stored once on the AreaScheme element
- All child elements inherit this value via relationships
- No other data needed on AreaScheme

---

## 2. Sheet

### Common Municipality
```json
{
  "AreaSchemeId": "<Revit ElementId>"
}
```

**Notes:**
- Common has NO DXF export attributes for sheets
- Only stores reference to parent AreaScheme
- `PAGE_NO` is calculated at export time (not stored)

### Jerusalem Municipality
```json
{
  "AreaSchemeId": "<Revit ElementId>",
  "PROJECT": "<string>",
  "ELEVATION": <float>,
  "BUILDING_HEIGHT": <float>,
  "X": <float>,
  "Y": <float>,
  "LOT_AREA": <float>,
  "SCALE": "<string>"
}
```

**Field Details:**
- `PROJECT`: If `<Project Name>` or `<Project Number>`, get from Revit document properties (Project Information)
- `ELEVATION`: If `<Project Base Point>`, get elevation from project base point in shared coordinates (meters)
- `BUILDING_HEIGHT`: User-entered value
- `X`: If `<Project Base Point>`, get X from project base point in shared coordinates (meters). If `<Internal Origin>`, get X from internal origin in shared coordinates (meters)
- `Y`: If `<Project Base Point>`, get Y from project base point in shared coordinates (meters). If `<Internal Origin>`, get Y from internal origin in shared coordinates (meters)
- `LOT_AREA`: User-entered value
- `SCALE`: Get from the scale of the area plans on the sheet

### Tel-Aviv Municipality
```json
{
  "AreaSchemeId": "<Revit ElementId>"
}
```

**Notes:**
- Tel-Aviv has NO DXF export attributes for sheets
- Only stores reference to parent AreaScheme

---

## 3. AreaPlan (View)

### Common Municipality
```json
{
  "FLOOR": "<string>",
  "LEVEL_ELEVATION": <float>,
  "IS_UNDERGROUND": <int>,
  "RepresentedViews": ["<Revit ElementId>", "<Revit ElementId>", ...]
}
```

**Field Details:**
- `FLOOR`: Floor name. **Default: `<View Name>`** (can also use `<Title on Sheet>`)
- `LEVEL_ELEVATION`: Level elevation in meters. **Default: `<Project Base Point>`** (can also use `<Shared Coordinates>`)
- `IS_UNDERGROUND`: 0 or 1
- `RepresentedViews`: List of AreaPlan ElementIds that this typical floor represents (empty list if not a typical floor)

### Jerusalem Municipality
```json
{
  "FLOOR_NAME": "<string>",
  "FLOOR_ELEVATION": <float>,
  "FLOOR_UNDERGROUND": "<string>",
  "RepresentedViews": ["<Revit ElementId>", "<Revit ElementId>", ...]
}
```

**Field Details:**
- `FLOOR_NAME`: Floor name. **Default: `<View Name>`** (can also use `<Title on Sheet>`)
- `FLOOR_ELEVATION`: Floor elevation in meters. **Default: `<Project Base Point>`** (can also use `<Shared Coordinates>`)
- `FLOOR_UNDERGROUND`: "yes" or "no"
- `RepresentedViews`: List of AreaPlan ElementIds that this typical floor represents (empty list if not a typical floor)

### Tel-Aviv Municipality
```json
{
  "FLOOR": "<string>",
  "HEIGHT": <float>,
  "X": <float>,
  "Y": <float>,
  "Absolute_height": <float>,
  "RepresentedViews": ["<Revit ElementId>", "<Revit ElementId>", ...]
}
```

**Field Details:**
- `FLOOR`: Floor name. **Default: `<View Name>`** (can also use `<Title on Sheet>`)
- `HEIGHT`: Floor height in CM. **Default: `<Auto>`** (calculate from difference between levels, or use defined value for topmost)
- `X`: If `<Project Base Point>`, get X from project base point in shared coordinates (meters). If `<Internal Origin>`, get X from internal origin in shared coordinates (meters)
- `Y`: If `<Project Base Point>`, get Y from project base point in shared coordinates (meters). If `<Internal Origin>`, get Y from internal origin in shared coordinates (meters)
- `Absolute_height`: If `<Project Base Point>`, use host level height from project base point (meters). If `<Shared Coordinates>`, use host level height in shared coordinates (meters)
- `RepresentedViews`: List of AreaPlan ElementIds that this typical floor represents (empty list if not a typical floor)

---

## 4. Area

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

### Tel-Aviv Municipality
```json
{
  "APARTMENT": "<string>",
  "HETER": "<string>",
  "HEIGHT": <float>
}
```

**Field Details:**
- `APARTMENT`: If `<*>` is used, get value from parameter assigned to area element
- `HETER`: User-entered permit/variance identifier
- `HEIGHT`: Room/area height

**Shared Parameters (not in JSON):**
- `Usage Type`: Current usage type code (exported as `CODE`)
- `Usage Type Prev`: Previous usage type code (exported as `CODE_BEFORE`)

---

## Summary Table

| Element Type | Common Fields | Jerusalem Fields | Tel-Aviv Fields |
|--------------|---------------|------------------|-----------------|
| **AreaScheme** | Municipality | Municipality | Municipality |
| **Sheet** | AreaSchemeId | AreaSchemeId, PROJECT, ELEVATION, BUILDING_HEIGHT, X, Y, LOT_AREA, SCALE | AreaSchemeId |
| **AreaPlan** | FLOOR, LEVEL_ELEVATION, IS_UNDERGROUND, RepresentedViews | FLOOR_NAME, FLOOR_ELEVATION, FLOOR_UNDERGROUND, RepresentedViews | FLOOR, HEIGHT, X, Y, Absolute_height, RepresentedViews |
| **Area** | AREA, ASSET | AREA, HEIGHT, APPARTMENT_NUM, HEIGHT2 | APARTMENT, HETER, HEIGHT |

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

3. **Special Values:**
   - `<Project Name>`, `<Project Number>`: Get from Revit document properties
   - `<View Name>`, `<Title on Sheet>`: Get from view parameters
   - `<Project Base Point>`: Use project base point in shared coordinates
   - `<Internal Origin>`: Use internal origin in shared coordinates
   - `<Shared Coordinates>`: Use shared coordinate system
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
