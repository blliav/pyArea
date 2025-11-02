# -*- coding: utf-8 -*-
"""Municipality-specific field definitions

Defines which fields are required/optional for each element type
across different municipalities (Common, Jerusalem, Tel-Aviv).
"""

# Municipality options
MUNICIPALITIES = ["Common", "Jerusalem", "Tel-Aviv"]


# AreaScheme fields (same for all municipalities)
AREASCHEME_FIELDS = {
    "Municipality": {
        "type": "string",
        "required": True,
        "options": MUNICIPALITIES,
        "description": "Municipality type"
    }
}


# Sheet fields by municipality
SHEET_FIELDS = {
    "Common": {
        "AreaSchemeId": {
            "type": "string",
            "required": True,
            "description": "Parent AreaScheme ElementId"
        }
    },
    "Jerusalem": {
        "AreaSchemeId": {
            "type": "string",
            "required": True,
            "description": "Parent AreaScheme ElementId"
        },
        "PROJECT": {
            "type": "string",
            "required": True,
            "description": "Project name or number"
        },
        "ELEVATION": {
            "type": "string",
            "required": True,
            "description": "Project base point elevation (meters)"
        },
        "BUILDING_HEIGHT": {
            "type": "string",
            "required": True,
            "description": "Building height"
        },
        "X": {
            "type": "string",
            "required": True,
            "description": "X coordinate (meters)"
        },
        "Y": {
            "type": "string",
            "required": True,
            "description": "Y coordinate (meters)"
        },
        "LOT_AREA": {
            "type": "string",
            "required": True,
            "description": "Lot area"
        },
        "SCALE": {
            "type": "string",
            "required": True,
            "description": "Drawing scale"
        }
    },
    "Tel-Aviv": {
        "AreaSchemeId": {
            "type": "string",
            "required": True,
            "description": "Parent AreaScheme ElementId"
        }
    }
}


# AreaPlan (View) fields by municipality
AREAPLAN_FIELDS = {
    "Common": {
        "FLOOR": {
            "type": "string",
            "required": True,
            "description": "Floor name from view",
            "default": "<View Name>"
        },
        "LEVEL_ELEVATION": {
            "type": "string",
            "required": True,
            "description": "Level elevation (meters)",
            "default": "<Project Base Point>"
        },
        "IS_UNDERGROUND": {
            "type": "int",
            "required": True,
            "description": "Underground flag (0 or 1)"
        },
        "RepresentedViews": {
            "type": "list",
            "required": False,
            "description": "List of represented view ElementIds (for typical floors)"
        }
    },
    "Jerusalem": {
        "FLOOR_NAME": {
            "type": "string",
            "required": True,
            "description": "Floor name from view",
            "default": "<View Name>"
        },
        "FLOOR_ELEVATION": {
            "type": "string",
            "required": True,
            "description": "Floor elevation (meters)",
            "default": "<Project Base Point>"
        },
        "FLOOR_UNDERGROUND": {
            "type": "string",
            "required": True,
            "description": "Underground flag (yes/no)"
        },
        "RepresentedViews": {
            "type": "list",
            "required": False,
            "description": "List of represented view ElementIds (for typical floors)"
        }
    },
    "Tel-Aviv": {
        "FLOOR": {
            "type": "string",
            "required": True,
            "description": "Floor name from view",
            "default": "<View Name>"
        },
        "HEIGHT": {
            "type": "string",
            "required": True,
            "description": "Floor height (CM)"
        },
        "X": {
            "type": "string",
            "required": True,
            "description": "X coordinate (meters)"
        },
        "Y": {
            "type": "string",
            "required": True,
            "description": "Y coordinate (meters)"
        },
        "Absolute_height": {
            "type": "string",
            "required": True,
            "description": "Absolute height (meters)"
        },
        "RepresentedViews": {
            "type": "list",
            "required": False,
            "description": "List of represented view ElementIds (for typical floors)"
        }
    }
}


# Area fields by municipality
AREA_FIELDS = {
    "Common": {
        "AREA": {
            "type": "string",
            "required": True,
            "description": "Area value"
        },
        "ASSET": {
            "type": "string",
            "required": False,
            "description": "Asset identifier"
        }
    },
    "Jerusalem": {
        "AREA": {
            "type": "string",
            "required": True,
            "description": "Area value"
        },
        "HEIGHT": {
            "type": "string",
            "required": True,
            "description": "Room height"
        },
        "APPARTMENT_NUM": {
            "type": "string",
            "required": True,
            "description": "Apartment number"
        },
        "HEIGHT2": {
            "type": "string",
            "required": False,
            "description": "Secondary height"
        }
    },
    "Tel-Aviv": {
        "APARTMENT": {
            "type": "string",
            "required": False,
            "description": "Apartment identifier"
        },
        "HETER": {
            "type": "string",
            "required": False,
            "description": "Permit/variance identifier"
        },
        "HEIGHT": {
            "type": "string",
            "required": True,
            "description": "Room height"
        }
    }
}


def get_fields_for_element_type(element_type, municipality=None):
    """Get field definitions for a specific element type and municipality.
    
    Args:
        element_type: "AreaScheme", "Sheet", "AreaPlan", or "Area"
        municipality: Municipality name (required for Sheet, AreaPlan, Area)
        
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
        element_type: "AreaScheme", "Sheet", "AreaPlan", or "Area"
        municipality: Municipality name (required for Sheet, AreaPlan, Area)
        
    Returns:
        list: List of required field names
    """
    fields = get_fields_for_element_type(element_type, municipality)
    return [name for name, props in fields.items() if props.get("required", False)]


def validate_data(element_type, data_dict, municipality=None):
    """Validate data dictionary against field definitions.
    
    Args:
        element_type: "AreaScheme", "Sheet", "AreaPlan", or "Area"
        data_dict: Data dictionary to validate
        municipality: Municipality name (required for Sheet, AreaPlan, Area)
        
    Returns:
        tuple: (is_valid, error_messages)
    """
    errors = []
    
    try:
        fields = get_fields_for_element_type(element_type, municipality)
    except ValueError as e:
        return False, [str(e)]
    
    # Check required fields
    required = get_required_fields(element_type, municipality)
    for field_name in required:
        if field_name not in data_dict or data_dict[field_name] is None:
            errors.append("Missing required field: {}".format(field_name))
    
    # Check data types
    for field_name, value in data_dict.items():
        if field_name not in fields:
            continue  # Allow extra fields
        
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
