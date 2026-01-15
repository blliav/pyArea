# pyArea Extension

pyRevit extension for managing area plans and exporting to DXF/DWFX formats with municipality-specific requirements.

> **BETA VERSION** - This extension is currently in beta. Tel-Aviv municipality has not been tested on actual projects.

## Installation

1. Install [pyRevit](https://github.com/pyrevitlabs/pyRevit/releases) latest version
2. Open Command Prompt (`Win + R` → type `cmd` → press Enter)
3. To install pyArea, run:
   ```
   pyrevit extend ui pyArea "https://github.com/blliav/pyArea.git"
   ```

To uninstall pyArea, run:
   ```
   pyrevit extensions delete pyArea
   ```


## Features

- **Define Schema** - Hierarchical data management for AreaSchemes, Sheets, and AreaPlans
- **Set Areas** - Define areas usage types and attributes
- **Export DXF** - Export area plans to DXF with municipality-specific attributes
- **Export DWFx** - Export background sheets in DWFX format. Including floors opaque background removal.
- **Utilities** - Hole filler, color schemes creator, and other helper tools

## Supported Municipalities

- Common (default)
- Common (gross)
- Jerusalem
- Tel-Aviv

## Quick Start

1. **Define AreaScheme**: Click "Define Schema" → Add Scheme → Select municipality
2. **Define Calculations**: In CalculationSetup, create Calculations on the AreaScheme, set sheet-level fields (PROJECT, ELEVATION, X, Y, etc.), and define default values for area plans and areas.
3. **Add Sheets**: Assign sheets to Calculations so they share the same metadata and defaults.
4. **Define AreaPlans**: Select Sheet → Add AreaPlan → Select views
5. **Set Properties**: Click any element in tree → Edit fields → Apply

## Configuration

### User Preferences

Export folder path is stored per-user in:
```
%APPDATA%\pyArea\preferences.json
```

This allows each user to have their own export directory independent of the Revit model.

### Project Preferences

Export settings (DXF/DWFx quality, colors, postprocessing options) are stored in the Revit model's ProjectInformation using extensible storage. This ensures consistent export settings across all team members working on the same model.

### Project Data

Schema data (AreaSchemes, Calculations, Sheets, AreaPlans, Areas) is stored in the Revit model using extensible storage, ensuring project-specific settings travel with the model file.
