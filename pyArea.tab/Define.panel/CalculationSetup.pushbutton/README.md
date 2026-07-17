# Calculation Setup Tool

Hierarchical data management tool for AreaSchemes, Calculations, Sheets, and AreaPlans using a tree-table (DataGrid) interface.

## Interface

### Left Panel - Tree-Table (DataGrid)

**Hierarchy structure:**
```
📐 AreaScheme (selected via dropdown at top)
├── 📊 Calculation (group header, bold, blue background)
│   └── 📄 Sheet (group header, semi-bold, gray background)
│       ├── ■ AreaPlan (on sheet, solid square)
│       │   └── 🔗 RepresentedAreaPlan (typical floor reference)
│       └── ■ AreaPlan ...
└── 📌 Not Placed (group header, italic, amber background)
    └── □ AreaPlan (not on sheet, hollow square)
```

**Dynamic field columns:**
- Columns are built based on municipality-specific `AREAPLAN_FIELDS`
- Cells show `resolved_value ← <placeholder>` for placeholder fields
- Columns where all values are identical are auto-hidden
- Inline editing supported (click cell to edit, applies to all selected rows)

**Tree controls:**
- Expander buttons (▾/▸) to collapse/expand group nodes
- Multi-row selection (Extended mode) for batch editing
- Click empty space to select the active AreaScheme

**Scheme selector row:**
- ComboBox to switch between AreaSchemes
- ✎ Edit Scheme button — selects scheme for property editing

### Right Panel - Properties
- **Title** — element name + type/municipality/variant
- **Fields** — municipality-specific editable fields (auto-save on change)
- **JSON viewer** — raw extensible storage data
- Multi-selection shows merged values with `<Varies>` for differing fields

## Workflow

1. **Select AreaScheme** (dropdown or ✎ button)
   - Set Municipality and Variant

2. **Add Calculation** (no row selected)
   - Creates a new Calculation on the AreaScheme
   - Set PROJECT, ELEVATION, X, Y, and default values

3. **Add Sheet** (Calculation selected)
   - Select sheets to link to this Calculation

4. **Add AreaPlan** (Sheet selected)
   - Shows views from same AreaScheme not yet tracked

5. **Add Represented View** (AreaPlan selected)
   - Select typical floor views (not on sheets)

6. **Set Representing View** (□ AreaPlan or 🔗 selected)
   - Move to a parent AreaPlan on a sheet

7. **Edit Properties** (any element selected)
   - Fields auto-save on change
   - Multi-select applies edits to all selected rows

## Buttons

- **➕ Add** — context-aware (Calculation / Sheet / AreaPlan / Represented / Set Representing View)
- **🗑 Remove** — remove data from selected element(s)
