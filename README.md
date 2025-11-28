# pyArea Extension

pyRevit extension for managing area plans and exporting to DXF/DWFX formats with municipality-specific requirements.

## Installation

1. Install [pyRevit](https://github.com/pyrevitlabs/pyRevit/releases)
2. Open Command Prompt (`Win + R` → type `cmd` → press Enter)
3. To install or update pyArea, run:
   ```
   pyrevit extend ui pyArea "https://github.com/blliav/pyArea"
   ```
4. To uninstall pyArea, run:
   ```
   pyrevit extensions delete pyArea
   ```


## Features

- **Define Schema** - Hierarchical data management for AreaSchemes, Sheets, and AreaPlans
- **Export DXF** - Export area plans to DXF with municipality-specific attributes
- **Export DWFX** - Export sheets to DWFX format
- **Utilities** - Color schemes and other helper tools

## Supported Municipalities

- Common (default)
- Common (gross)
- Jerusalem
- Tel-Aviv

## Quick Start

1. **Define AreaScheme**: Click "Define Schema" → Add Scheme → Select municipality
2. **Define Calculations**: In CalculationSetup, create Calculations on the AreaScheme, set project-level fields (PROJECT, ELEVATION, X, Y, etc.), and define default values for AreaPlans/Areas.
3. **Add Sheets**: Assign sheets to Calculations so they share the same metadata and defaults.
4. **Define AreaPlans**: Select Sheet → Add AreaPlan → Select views
5. **Set Properties**: Click any element in tree → Edit fields → Apply
