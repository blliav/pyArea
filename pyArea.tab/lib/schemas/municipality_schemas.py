# -*- coding: utf-8 -*-
"""Municipality-specific field definitions

Defines which fields are required/optional for each element type
across different municipalities (Common, Jerusalem, Tel-Aviv).

IMPORTANT: This module must remain compatible with BOTH CPython and IronPython.
- Used by IronPython scripts: SetAreas, CalculationSetup, data_manager
- Used by CPython scripts: ExportDXF
- DO NOT import Revit API or pyRevit modules here
- Keep as pure Python with only standard library imports
"""

from collections import OrderedDict

# Municipality options
MUNICIPALITIES = ["Common", "Jerusalem", "Tel-Aviv"]

# Variant configurations by municipality
# Maps municipality -> list of available variants
MUNICIPALITY_VARIANTS = {
    "Common": ["Default", "Gross"],
    "Jerusalem": ["Default"],
    "Tel-Aviv": ["Default"]
}


def get_usage_type_csv_filename(municipality, variant="Default"):
    """Get CSV filename for usage types based on municipality and variant.
    
    Args:
        municipality: Base municipality name
        variant: Variant name (default: "Default")
    
    Returns:
        str: CSV filename (e.g., "UsageType_Common.csv" or "UsageType_CommonGross.csv")
    
    Examples:
        >>> get_usage_type_csv_filename("Common", "Default")
        'UsageType_Common.csv'
        >>> get_usage_type_csv_filename("Common", "Gross")
        'UsageType_CommonGross.csv'
        >>> get_usage_type_csv_filename("Jerusalem")
        'UsageType_Jerusalem.csv'
    """
    if variant == "Default":
        return "UsageType_{}.csv".format(municipality)
    else:
        return "UsageType_{}{}.csv".format(municipality, variant)


# AreaScheme fields (same for all municipalities)
AREASCHEME_FIELDS = {
    "Municipality": {
        "type": "string",
        "required": True,
        "options": MUNICIPALITIES,
        "description": "Municipality type",
        "hebrew_name": "רשות"
    },
    "Variant": {
        "type": "string",
        "required": False,
        "default": "Default",
        "description": "Usage type catalog variant",
        "hebrew_name": "גרסה"
    }
}


# Calculation fields by municipality
# Calculations are stored on AreaScheme elements
# Note: CalculationGuid is the KEY in the Calculations dictionary, not a field
# Field order: Name, calculation-specific fields, AreaPlanDefaults, AreaDefaults
CALCULATION_FIELDS = {
    "Common": OrderedDict([
        ("Name", {
            "type": "string",
            "required": True,
            "description": "User-facing calculation name",
            "hebrew_name": "שם חישוב"
        }),
        ("AreaPlanDefaults", {
            "type": "dict",
            "required": False,
            "description": "Default values for AreaPlan elements"
        }),
        ("AreaDefaults", {
            "type": "dict",
            "required": False,
            "description": "Default values for Area elements"
        })
    ]),
    "Jerusalem": OrderedDict([
        ("Name", {
            "type": "string",
            "required": True,
            "description": "User-facing calculation name",
            "hebrew_name": "שם חישוב"
        }),
        ("PROJECT", {
            "type": "string",
            "required": True,
            "description": "Project name or number",
            "default": "<Project Name>",
            "placeholders": ["<Project Name>", "<Project Number>"],
            "hebrew_name": "פרויקט"
        }),
        ("ELEVATION", {
            "type": "string",
            "required": True,
            "description": "Project base point elevation (meters)",
            "default": "<SharedElevation@ProjectBasePoint>",
            "placeholders": ["<SharedElevation@ProjectBasePoint>"],
            "hebrew_name": "גובה בסיס"
        }),
        ("BUILDING_HEIGHT", {
            "type": "string",
            "required": True,
            "description": "Building height",
            "hebrew_name": "גובה בניין"
        }),
        ("X", {
            "type": "string",
            "required": True,
            "description": "X coordinate (meters)",
            "default": "<E/W@InternalOrigin>",
            "placeholders": ["<E/W@ProjectBasePoint>", "<E/W@InternalOrigin>"]
        }),
        ("Y", {
            "type": "string",
            "required": True,
            "description": "Y coordinate (meters)",
            "default": "<N/S@InternalOrigin>",
            "placeholders": ["<N/S@ProjectBasePoint>", "<N/S@InternalOrigin>"]
        }),
        ("LOT_AREA", {
            "type": "string",
            "required": True,
            "description": "Lot area",
            "hebrew_name": "שטח מגרש"
        }),
        ("AreaPlanDefaults", {
            "type": "dict",
            "required": False,
            "description": "Default values for AreaPlan elements"
        }),
        ("AreaDefaults", {
            "type": "dict",
            "required": False,
            "description": "Default values for Area elements"
        })
    ]),
    "Tel-Aviv": OrderedDict([
        ("Name", {
            "type": "string",
            "required": True,
            "description": "User-facing calculation name",
            "hebrew_name": "שם חישוב"
        }),
        ("AreaPlanDefaults", {
            "type": "dict",
            "required": False,
            "description": "Default values for AreaPlan elements"
        }),
        ("AreaDefaults", {
            "type": "dict",
            "required": False,
            "description": "Default values for Area elements"
        })
    ])
}


# Sheet fields by municipality
# Simplified - Sheets now only store a reference to their parent Calculation
# Exception: DWFx_UnderlayFilename for custom underlay file reference
SHEET_FIELDS = {
    "Common": {
        "CalculationGuid": {
            "type": "string",
            "required": True,
            "description": "Parent Calculation identifier (UUID)"
        },
        "DWFx_UnderlayFilename": {
            "type": "string",
            "required": False,
            "description": "Optional DWFX underlay filename (e.g., MyProject-A101.dwfx)",
            "default": "<same as dxf sheet>",
            "hebrew_name": "קובץ DWFX רקע"
        }
    },
    "Jerusalem": {
        "CalculationGuid": {
            "type": "string",
            "required": True,
            "description": "Parent Calculation identifier (UUID)"
        },
        "DWFx_UnderlayFilename": {
            "type": "string",
            "required": False,
            "description": "Optional DWFX underlay filename (e.g., MyProject-A101.dwfx)",
            "default": "<same as dxf sheet>",
            "hebrew_name": "קובץ DWFX רקע"
        }
    },
    "Tel-Aviv": {
        "CalculationGuid": {
            "type": "string",
            "required": True,
            "description": "Parent Calculation identifier (UUID)"
        },
        "DWFx_UnderlayFilename": {
            "type": "string",
            "required": False,
            "description": "Optional DWFX underlay filename (e.g., MyProject-A101.dwfx)",
            "default": "<same as dxf sheet>",
            "hebrew_name": "קובץ DWFX רקע"
        }
    }
}


# AreaPlan (View) fields by municipality
AREAPLAN_FIELDS = {
    "Common": OrderedDict([
        ("FLOOR", {
            "type": "string",
            "required": True,
            "description": "Floor name from view",
            "default": "<Level Name>",
            "placeholders": ["<Level Name>","<View Name>", "<Title on Sheet>"],
            "hebrew_name": "שם קומה"
        }),
        ("LEVEL_ELEVATION", {
            "type": "string",
            "required": True,
            "description": "Level elevation (meters)",
            "default": "<by Project Base Point>",
            "placeholders": ["<by Project Base Point>", "<by Shared Coordinates>"],
            "hebrew_name": "מפלס קומה"
        }),
        ("IS_UNDERGROUND", {
            "type": "int",
            "required": True,
            "description": "Underground flag (0 or 1)",
            "default": 0,
            "hebrew_name": "תת קרקעי"
        }),
        ("RepresentedViews", {
            "type": "list",
            "required": False,
            "description": "List of represented view ElementIds (for typical floors)",
            "hebrew_name": "קומות מיוצגות"
        })
    ]),
    "Jerusalem": OrderedDict([
        ("BUILDING_NAME", {
            "type": "string",
            "required": True,
            "description": "Building name",
            "default": "1",
            "hebrew_name": "שם בניין"
        }),
        ("FLOOR_NAME", {
            "type": "string",
            "required": True,
            "description": "Floor name from view",
            "default": "<Level Name>",
            "placeholders": ["<Level Name>","<View Name>", "<Title on Sheet>"],
            "hebrew_name": "שם קומה"
        }),
        ("FLOOR_ELEVATION", {
            "type": "string",
            "required": True,
            "description": "Floor elevation (meters)",
            "default": "<by Project Base Point>",
            "placeholders": ["<by Project Base Point>", "<by Shared Coordinates>"],
            "hebrew_name": "מפלס קומה"
        }),
        ("FLOOR_UNDERGROUND", {
            "type": "string",
            "required": True,
            "description": "Underground flag (yes/no)",
            "default": "no",
            "hebrew_name": "תת קרקעי"
        }),
        ("RepresentedViews", {
            "type": "list",
            "required": False,
            "description": "List of represented view ElementIds (for typical floors)",
            "hebrew_name": "קומות מיוצגות"
        })
    ]),
    "Tel-Aviv": OrderedDict([
        ("BUILDING", {
            "type": "string",
            "required": True,
            "description": "Building name",
            "default": "1",
            "hebrew_name": "סימון\מספר בניין"
        }),
        ("FLOOR", {
            "type": "string",
            "required": True,
            "description": "Floor name from view",
            "default": "<Level Name>",
            "placeholders": ["<Level Name>", "<View Name>", "<Title on Sheet>"],
            "hebrew_name": "שם קומה"
        }),
        ("HEIGHT", {
            "type": "string",
            "required": True,
            "description": "Floor height (CM)",
            "hebrew_name": "גובה"
        }),
        ("X", {
            "type": "string",
            "required": True,
            "description": "X coordinate (meters)",
            "default": "<E/W@InternalOrigin>",
            "placeholders": ["<E/W@ProjectBasePoint>", "<E/W@InternalOrigin>"]
        }),
        ("Y", {
            "type": "string",
            "required": True,
            "description": "Y coordinate (meters)",
            "default": "<N/S@InternalOrigin>",
            "placeholders": ["<N/S@ProjectBasePoint>", "<N/S@InternalOrigin>"]
        }),
        ("Absolute_height", {
            "type": "string",
            "required": True,
            "description": "Absolute height (meters)",
            "default": "<by Shared Coordinates>",
            "placeholders": ["<by Shared Coordinates>"],
            "hebrew_name": "גובה אבסולוטי"
        }),
        ("RepresentedViews", {
            "type": "list",
            "required": False,
            "description": "List of represented view ElementIds (for typical floors)",
            "hebrew_name": "קומות מיוצגות"
        })
    ])
}


# Area fields by municipality
AREA_FIELDS = {
    "Common": OrderedDict([
        ("AREA", {
            "type": "string",
            "required": False,
            "description": "Manual area",
            "hebrew_name": "שטח ידני"
        }),
        ("ASSET", {
            "type": "string",
            "required": False,
            "description": "Asset identifier",
            "hebrew_name": ""
        })
    ]),
    "Jerusalem": OrderedDict([
        ("HEIGHT", {
            "type": "string",
            "required": True,
            "description": "Room height",
            "default": "<by Level Above>",
            "placeholders": ["<by Level Above>"],
            "hebrew_name": "גובה"
        }),
        ("AREA", {
            "type": "string",
            "required": False,
            "description": "Manual area",
            "hebrew_name": "שטח ידני"
        }),
        ("APPARTMENT_NUM", {
            "type": "string",
            "required": False,
            "description": "Apartment number",
            "hebrew_name": "מס' דירה"
        }),
        ("HEIGHT2", {
            "type": "string",
            "required": False,
            "description": "Secondary height",
            "hebrew_name": "גובה 2"
        })
    ]),
    "Tel-Aviv": OrderedDict([
        ("ID", {
            "type": "string",
            "required": False,
            "description": "Area identifier/number",
            "default": "",
            "placeholders": ["<AreaNumber>"],
            "hebrew_name": "מזהה יחודי לתא שטח בקומה"
        }),
        ("APARTMENT", {
            "type": "string",
            "required": True,
            "default": "1",
            "description": "Apartment identifier",
            "hebrew_name": "דירה"
        }),
        ("HETER", {
            "type": "string",
            "required": True,
            "description": "Permit/variance identifier",
            "default": "1",
            "hebrew_name": "היתר"
        }),
        ("HEIGHT", {
            "type": "string",
            "required": False,
            "description": "Room height",
            "hebrew_name": "גובה"
        })
    ])
}


def get_fields_for_element_type(element_type, municipality=None):
    """Get field definitions for a specific element type and municipality.
    
    Args:
        element_type: "AreaScheme", "Calculation", "Sheet", "AreaPlan", or "Area"
        municipality: Municipality name (required for Calculation, Sheet, AreaPlan, Area)
        
    Returns:
        dict: Field definitions for the element type
    """
    if element_type == "AreaScheme":
        return AREASCHEME_FIELDS
    
    if not municipality:
        raise ValueError("Municipality required for {}".format(element_type))
    
    if municipality not in MUNICIPALITIES:
        raise ValueError("Invalid municipality: {}".format(municipality))
    
    field_map = {
        "Calculation": CALCULATION_FIELDS,
        "Sheet": SHEET_FIELDS,
        "AreaPlan": AREAPLAN_FIELDS,
        "Area": AREA_FIELDS
    }
    
    if element_type not in field_map:
        raise ValueError("Invalid element type: {}".format(element_type))
    
    return field_map[element_type].get(municipality, {})


def get_required_fields(element_type, municipality=None):
    """Get list of required field names for element type and municipality.
    
    Args:
        element_type: "AreaScheme", "Calculation", "Sheet", "AreaPlan", or "Area"
        municipality: Municipality name (required for Calculation, Sheet, AreaPlan, Area)
        
    Returns:
        list: List of required field names
    """
    fields = get_fields_for_element_type(element_type, municipality)
    return [name for name, props in fields.items() if props.get("required", False)]


def validate_data(element_type, data_dict, municipality=None):
    """Validate data dictionary against field definitions.
    
    Args:
        element_type: "AreaScheme", "Calculation", "Sheet", "AreaPlan", or "Area"
        data_dict: Data dictionary to validate
        municipality: Municipality name (required for Calculation, Sheet, AreaPlan, Area)
        
    Returns:
        tuple: (is_valid, error_messages)
    
    Note:
        None values are allowed for any field to support inheritance.
        Type checking is skipped for None values.
    """
    errors = []
    
    try:
        fields = get_fields_for_element_type(element_type, municipality)
    except ValueError as e:
        return False, [str(e)]
    
    # Check required fields
    # Note: Required fields with default values can be missing/None
    # since the default will be used during export
    required = get_required_fields(element_type, municipality)
    for field_name in required:
        if field_name not in data_dict or data_dict[field_name] is None:
            # Check if field has a default value
            field_def = fields.get(field_name, {})
            has_default = "default" in field_def or "placeholders" in field_def
            
            # Only error if required field has no default
            if not has_default:
                errors.append("Missing required field: {}".format(field_name))
    
    # Check data types
    for field_name, value in data_dict.items():
        if field_name not in fields:
            continue  # Allow extra fields
        
        # Skip type check for None values (supports inheritance)
        if value is None:
            continue
        
        field_type = fields[field_name].get("type")
        if field_type == "string" and not isinstance(value, (str, unicode)):
            errors.append("{} must be a string".format(field_name))
        elif field_type == "float" and not isinstance(value, (int, float)):
            errors.append("{} must be a number".format(field_name))
        elif field_type == "int" and not isinstance(value, int):
            errors.append("{} must be an integer".format(field_name))
        elif field_type == "list" and not isinstance(value, list):
            errors.append("{} must be a list".format(field_name))
    
    return len(errors) == 0, errors


# ============================================================================
# DXF EXPORT CONFIGURATION
# ============================================================================
# Single source of truth for DXF export settings (layers, colors, string templates)
# Used by ExportDXF.pushbutton (CPython)

DXF_CONFIG = {
    "Common": {
        "layers": {
            'sheet_frame': 'RZ_FRAME',
            'sheet_text': 'RZ_FRAME',
            'areaplan_frame': 'RZ_FLOOR',
            'areaplan_text': 'RZ_FLOOR',
            'area_boundary': 'RZ_AREA',
            'area_text': 'RZ_AREA'
        },
        "layer_colors": {
            'RZ_FRAME': 7,   # White
            'RZ_FLOOR': 7,   # White
            'RZ_AREA': 1     # Red
        },
        "string_templates": {
            "sheet": "PAGE_NO={page_number}",
            "areaplan": "BUILDING_NO={building_no}&&&FLOOR={floor}&&&LEVEL_ELEVATION={level_elevation}&&&IS_UNDERGROUND={is_underground}",
            "area": "USAGE_TYPE={usage_type}&&&USAGE_TYPE_OLD={usage_type_old}&&&AREA={area}&&&ASSET={asset}"
        }
    },
    "Jerusalem": {
        "layers": {
            'sheet_frame': 'AREA_PLAN_MAIN_FRAME',
            'sheet_text': 'AREA_PLAN_MAIN_FRAME',
            'areaplan_frame': 'AREA_PLAN_FLOOR_FRAME',
            'areaplan_text': 'AREA_PLAN_FLOOR_TABLE',
            'area_boundary': 'AREA_PLAN_BORDER',
            'area_text': 'AREA_PLAN_SYMBOL'
        },
        "layer_colors": {
            'AREA_PLAN_MAIN_FRAME': 7,    # White
            'AREA_PLAN_FLOOR_FRAME': 7,   # White
            'AREA_PLAN_FLOOR_TABLE': 3,   # Green
            'AREA_PLAN_BORDER': 1,        # Red
            'AREA_PLAN_SYMBOL': 1         # Red
        },
        "string_templates": {
            "sheet": "PROJECT={project}&&&ELEVATION={elevation}&&&BUILDING_HEIGHT={building_height}&&&X={x}&&&Y={y}&&&LOT_AREA={lot_area}&&&SCALE={scale}",
            "areaplan": "BUILDING_NAME={building_name}&&&FLOOR_NAME={floor_name}&&&FLOOR_ELEVATION={floor_elevation}&&&FLOOR_UNDERGROUND={floor_underground}",
            "area": "CODE={code}&&&DEMOLITION_SOURCE_CODE={demolition_source_code}&&&AREA={area}&&&HEIGHT1={height1}&&&APPARTMENT_NUM={appartment_num}&&&HEIGHT2={height2}"
        }
    },
    "Tel-Aviv": {
        "layers": {
            'sheet_frame': 'AREA_PLAN_MAIN_FRAME',
            'sheet_text': 'AREA_PLAN_MAIN_FRAME',
            'areaplan_frame': 'muni_floor',
            'areaplan_text': 'muni_floor',
            'area_boundary': 'muni_area',
            'area_text': 'muni_area'
        },
        "layer_colors": {
            'AREA_PLAN_MAIN_FRAME': 7,   # White
            'muni_floor': 6,   # Magenta
            'muni_area': 1     # Red
        },
        "string_templates": {
            "sheet": "PAGE_NO={page_number}",
            "areaplan": "BUILDING={building}&&&FLOOR={floor}&&&HEIGHT={height}&&&X={x}&&&Y={y}&&&ABSOLUTE_HEIGHT={absolute_height}",
            "area": "CODE={code}&&&CODE_BEFORE={code_before}&&&ID={id}&&&APARTMENT={apartment}&&&HETER={heter}&&&HEIGHT={height}"
        }
    }
}
