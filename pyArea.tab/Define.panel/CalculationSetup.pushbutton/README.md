# Calculation Setup Tool

Hierarchical data management tool for AreaSchemes, Sheets, and AreaPlans.

## Interface

### Left Panel - Hierarchy Tree
- **📐 AreaScheme** - Municipality-defined schemes
- **📄 Sheet** - Sheets linked to AreaScheme
- **■ AreaPlan** - Views on sheets (solid square)
- **□ AreaPlan** - Views not on sheets (hollow square)
- **🔗 RepresentedView** - Typical floor references

### Right Panel - Properties
- Element information
- Municipality (auto-detected for Sheets/AreaPlans)
- Data fields (varies by municipality)
- Status messages

## Workflow

1. **Add Scheme** (no selection)
   - Select undefined AreaScheme
   - Set municipality

2. **Add Sheet** (AreaScheme selected)
   - Select sheets to link
   - Pre-checks sheets with AreaPlans

3. **Add AreaPlan** (Sheet selected)
   - Shows views from same AreaScheme
   - ■ = already on sheet, □ = not on sheet

4. **Add Represented View** (AreaPlan selected)
   - Select typical floor views (not on sheets)
   - For representing multiple actual floors

5. **Edit Properties** (any element selected)
   - Fill in fields
   - Click Apply to save

## Buttons

- **➕ Add** - Context-aware (changes text based on selection)
- **🗑 Remove** - Remove data (not Revit elements)
- **🔄 Refresh** - Reload from Revit
- **Load Data** - Load existing data
- **Apply** - Save changes
