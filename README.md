# pyArea Extension

pyRevit extension for managing area plans and exporting to DXF/DWFX formats with municipality-specific requirements.

## Features

- **Define Schema** - Hierarchical data management for AreaSchemes, Sheets, and AreaPlans
- **Export DXF** - Export area plans to DXF with municipality-specific attributes
- **Export DWFX** - Export sheets to DWFX format
- **Utilities** - Color schemes and other helper tools

## Supported Municipalities

- Common (default)
- Jerusalem
- Tel-Aviv

## Quick Start

1. **Define AreaScheme**: Click "Define Schema" → Add Scheme → Select municipality
2. **Add Sheets**: Select AreaScheme → Add Sheet → Choose sheets
3. **Define AreaPlans**: Select Sheet → Add AreaPlan → Select views
4. **Set Properties**: Click any element in tree → Edit fields → Apply

## Documentation

- `DATA_SCHEMA_DECISIONS.md` - Architecture and design decisions
- `JSON_TEMPLATES.md` - Data structure and field definitions
- `CalculationSettings.pushbutton/README.md` - Calculation Settings tool documentation