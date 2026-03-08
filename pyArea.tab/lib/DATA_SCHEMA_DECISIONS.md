# PyArea Data Schema - Design Decisions

**Date:** November 2, 2025

## Core Decisions

### 1. Storage: Revit Extensible Storage
- Use Extensible Storage API (not parameters, not external DB)
- Data persists with the Revit model
- No UI clutter from excessive parameters

### 2. Schema Structure
```python
Schema Name: "pyArea"
GUID: [Generate once and hardcode - NEVER CHANGE]
Fields: 
  - Data (string): JSON-encoded dictionary
```

**One schema. One field. All element types.**

### 3. Why JSON?
- Revit schemas are **immutable** (can't add/remove fields after creation)
- JSON provides flexibility for varying municipality requirements
- Easy to add new municipalities or fields without schema changes
- No versioning needed - JSON evolves naturally

### 4. Data Hierarchy
```
AreaScheme
  └─ Data: {"Municipality": "Jerusalem"}
      ├─ Sheets → Data: {"AreaSchemeId": "123", "CalculationName": "...", ...}
      ├─ AreaPlans → Data: {"FloorName": "...", "BuildingName": "...", ...}
      └─ Areas → Data: {"Height": 2.8, "Apartment": "A1", ...}
```

- **Municipality stored once** on AreaScheme
- Child elements reference AreaScheme via ID or native Revit relationships
- No redundant municipality field on children

### 5. Shared Parameters (Only 2)
- `Usage Type` (Text) - for color schemes/schedules
- `Usage Type Prev` (Text) - for color schemes/schedules
- **Everything else** goes in Extensible Storage JSON

### 6. Municipality Field Definitions

**Key Insight:** All municipalities have the **same fields** for each element type. The difference is which fields are **required** vs **optional**.

**Example - Area fields:**
- `Height` - Required in Jerusalem, Optional in Common/Tel-Aviv
- `ManualArea` - Required in Jerusalem, Optional in Common/Tel-Aviv
- `Apartment` - Required in Jerusalem, Optional in Common/Tel-Aviv

**Example - Sheet fields:**
- `CalculationName` - Required in Jerusalem, Optional in Common/Tel-Aviv
- `LotArea` - Required in Jerusalem, Optional in Common/Tel-Aviv
- `ProjectCoordX/Y/Z` - Required in Jerusalem, Optional in Common/Tel-Aviv

**Example - AreaPlan fields:**
- `FloorName` - Required in all municipalities
- `Elevation` - Required in all municipalities
- `IsUnderground` - Required in all municipalities
- `BuildingName` - Required in Jerusalem, Optional in Common/Tel-Aviv
- `RepresentedFloors` - Required in Jerusalem, Optional in Common/Tel-Aviv

Stored in: `lib/schemas/municipality_schemas.py` with required/optional flags per municipality

### 7. Implementation Structure
```
pyArea.tab/lib/
  ├── schemas/
  │   ├── schema_guids.py         # SCHEMA_GUID, SCHEMA_NAME constants
  │   ├── schema_manager.py       # get_or_create_schema(), set_data(), get_data()
  │   └── municipality_schemas.py # Field definitions per municipality
  └── data_manager.py             # High-level API (get_municipality, set_sheet_data, etc.)
```

### 8. Key Rules
- ✅ Generate GUID **once** during implementation
- ✅ **Hardcode** GUID in `schema_guids.py`
- ❌ **NEVER** change GUID after deployment
- ❌ No versioning needed (JSON handles evolution)
- ❌ No EntityType field (element type is known from Revit element)
- ❌ No Municipality field on child elements (inherit from AreaScheme)

### 9. API Pattern
```python
# Simple and consistent
set_data(element, {"key": "value"})
data = get_data(element)  # Returns dict
```

---

**End of Summary**
