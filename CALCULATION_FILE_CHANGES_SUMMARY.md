# Calculation Hierarchy - File Changes Summary (v2)

**Date:** November 14, 2025  
**Version:** v2 (AreaScheme storage, CalculationGuid + Name)

This document summarizes the concrete code changes needed to implement the final Calculation design:

- Calculations are stored **on the AreaScheme element**.
- Each Calculation has **CalculationGuid (GUID)** + **Name (user-facing)**.
- Sheets store **only `CalculationGuid`**.
- Defaults for AreaPlan/Area live under `Calculation.Defaults` and are resolved by inheritance.

---

## 1. `pyArea.tab/lib/schemas/municipality_schemas.py`

- **Add `CALCULATION_FIELDS`**
  - Common:
    - `CalculationGuid` (string, required) – system GUID
    - `Name` (string, required) – display label
    - `Defaults` (dict, optional) – holds `{"AreaPlan": {...}, "Area": {...}}`
  - Jerusalem:
    - `CalculationGuid`, `Name`
    - Existing Sheet-level fields (`PROJECT`, `ELEVATION`, `BUILDING_HEIGHT`, `X`, `Y`, `LOT_AREA`), all moved from Sheet to Calculation.
  - Tel-Aviv:
    - `CalculationGuid`, `Name`
    - Any Calculation-level fields you want (often empty, relying on Defaults).

- **Simplify `SHEET_FIELDS`**
  - Replace all per-sheet metadata with a single field:
    - `CalculationGuid` (string, required) – reference to Calculation on AreaScheme.

- **Update `get_fields_for_element_type()`**
  - Add mapping for element type `"Calculation"` → `CALCULATION_FIELDS`.

- **Update `validate_data()`**
  - Allow `None` for typed fields so inheritance works:
    - If value is `None`, **skip type check** (treat as "inherit").

---

## 2. `pyArea.tab/lib/data_manager.py`

Add Calculation helpers that work **per AreaScheme**, not ProjectInformation:

- **New functions**
  - `get_all_calculations(area_scheme)`
    - Reads `schema_manager.get_data(area_scheme)["Calculations"]` (or `{}`).
  - `get_calculation(area_scheme, calculation_guid)`
    - Looks up a single Calculation by GUID from the AreaScheme.
  - `generate_calculation_guid()`
    - Returns a new UUID4 string.
  - `set_calculation(area_scheme, calculation_guid, calculation_data, municipality)`
    - Validates via `municipality_schemas.validate_data("Calculation", ...)`.
    - Ensures `CalculationGuid` and `Name` are set.
    - Writes into `data["Calculations"][calculation_guid]` on the AreaScheme.
  - `delete_calculation(area_scheme, calculation_guid)`
    - Removes entry from `data["Calculations"]` on AreaScheme.
  - `get_calculation_from_sheet(doc, sheet)`
    - Reads `CalculationGuid` from Sheet JSON.
    - Finds first viewport → view → `view.AreaScheme`.
    - Returns `(area_scheme, calculation_data)` or `(None, None)`.
  - `resolve_field_value(field_name, element_data, calculation_data, municipality, element_type)`
    - Resolution order: element explicit → Calculation.Defaults[element_type] → schema default.

- **Change `set_sheet_data`**
  - New signature: `set_sheet_data(sheet, calculation_guid)`.
  - Writes `{ "CalculationGuid": calculation_guid }` onto the Sheet (no validation call needed).

- **Keep schema version helpers**
  - `get_schema_version(doc)` / `set_schema_version(doc, "2.0")` still use ProjectInformation.

---

## 3. `pyArea.tab/lib/JSON_TEMPLATES.md`

- **Calculation section**
  - Document Calculation as a new entity stored on **AreaScheme**:
    - `Calculations: { <CalculationGuid>: { ... } }`.
  - Show examples for each municipality consistent with `CALCULATION_FIELDS`.

- **Sheet section**
  - Replace all old fields with only:
    ```json
    {
      "CalculationGuid": "a7b3c9d1-e5f2-..."
    }
    ```
  - Note that `PAGE_NO` is **not stored**; it is derived from sheet order at export.

- **AreaPlan / Area sections**
  - Describe inheritance semantics:
    - Explicit value → overrides.
    - `null` → inherit from Calculation defaults.
    - Missing → use field schema default.

---

## 4. `pyArea.tab/Export.panel/ExportDXF.pushbutton/ExportDXF_script.py`

Adjust data flow to use `CalculationGuid` + AreaScheme lookup.

- **Helper: `resolve_field_value(...)`**
  - Implement same logic as in `data_manager` (can be duplicated or imported).

- **`get_sheet_data_for_dxf(sheet_elem)`**
  - Read `CalculationGuid` from sheet JSON.
  - From first viewport:
    - Resolve `view.AreaScheme`.
    - Load AreaScheme JSON, get `Calculations[CalculationGuid]`.
    - Read `Municipality` from AreaScheme JSON.
  - Return a dict that is essentially `calculation_data` plus `"Municipality"`.

- **`get_areaplan_data_for_dxf(view_elem, calculation_data, municipality)`**
  - Load original AreaPlan JSON.
  - For each field in `AREAPLAN_FIELDS[municipality]`, call `resolve_field_value(...)`.

- **`get_area_data_for_dxf(area_elem, calculation_data, municipality)`**
  - Load original Area JSON.
  - For each field in `AREA_FIELDS[municipality]`, call `resolve_field_value(...)`.
  - Still add shared parameters: `UsageType`, `UsageTypePrev` from Revit parameters.

- **`process_area` / `process_areaplan_viewport` / `process_sheet`**
  - Thread `calculation_data` + `municipality` through the call stack:
    - `process_sheet` calls `get_sheet_data_for_dxf` once → gets `calculation_data` + `municipality`.
    - For each viewport, call `process_areaplan_viewport(..., municipality, calculation_data, ...)`.
    - Inside, call `get_areaplan_data_for_dxf` and `process_area(..., municipality, calculation_data, ...)`.

- **`get_valid_areaplans_and_uniform_scale`**
  - When validating sheets:
    - Check that `CalculationGuid` exists on sheet.
    - Use first viewport → `view.AreaScheme` to resolve AreaScheme.
    - From AreaScheme JSON, confirm that `Calculations[CalculationGuid]` exists.
    - Use this Calculation/AreaScheme for any further consistency checks.

---

## 5. Migration Script

Create a new pushbutton (e.g. `pyArea.tab/Tools.panel/MigrateToCalculations.pushbutton/`) that:

1. Checks `SchemaVersion` (ProjectInformation) and exits if already `"2.0"`.
2. Collects all sheets using the **old schema** (Sheet has project fields like `PROJECT`).
3. For each sheet:
   - Resolve `AreaScheme` via the first viewport.
   - Determine municipality via existing helpers.
4. Group sheets **per AreaScheme + identical Calculation field values**.
5. For each group under an AreaScheme:
   - Generate a new `CalculationGuid` and `Name`.
   - Build `calculation_data` from the old sheet fields.
   - Call `set_calculation(area_scheme, calculation_guid, calculation_data, municipality)`.
   - For each sheet in the group, call `set_sheet_data(sheet, calculation_guid)`.
6. Set `SchemaVersion` to `"2.0"` using `set_schema_version`.

This logic is almost identical to the original migration plan, but with:

- Calculations written to **AreaScheme** instead of ProjectInformation.
- `CalculationGuid` used instead of `CalculationId`.
- No `AreaSchemeId` stored inside Calculation.

---

## 6. Testing Checklist (v2)

- **Schemas**
  - `CALCULATION_FIELDS` defined with `CalculationGuid` + `Name`.
  - `SHEET_FIELDS` contains only `CalculationGuid`.
  - `validate_data()` accepts `None` values (inheritance).

- **Data manager**
  - `get_all_calculations`, `get_calculation`, `set_calculation`, `delete_calculation` operate on AreaScheme.
  - `get_calculation_from_sheet` correctly resolves AreaScheme via viewport.
  - `set_sheet_data` writes only `CalculationGuid`.

- **Migration**
  - Old projects with Sheet-level metadata migrate to AreaScheme Calculations.
  - Grouping logic correctly groups sheets with identical field values under the same AreaScheme.
  - After migration, all sheets have `CalculationGuid` and no old metadata.

- **ExportDXF**
  - Reads Calculation via AreaScheme + `CalculationGuid`.
  - Inheritance resolution returns the expected values for AreaPlan/Area.
  - Explicit values override Calculation defaults.
  - `null` values correctly inherit from Calculation.
  - Omitted fields fall back to schema defaults.
  - PAGE_NO is derived from sheet order and is **not stored** in JSON.

---

## 7. Summary

This v2 summary reflects the final agreed design:

- Calculations stored **on AreaSchemes**.
- Identity via `CalculationGuid` + `Name`.
- Sheets store only a `CalculationGuid` reference.
- Migration and ExportDXF logic work through AreaScheme, not ProjectInformation.

Use this document together with `CALCULATION_HIERARCHY_PLAN.md` when implementing the changes.
