# -*- coding: utf-8 -*-
"""Low-level Extensible Storage Schema Manager

Handles creation and access to Revit Extensible Storage schema.
Stores data as JSON strings for maximum flexibility.
"""

import json
import System
from pyrevit import DB
from schema_guids import SCHEMA_GUID, SCHEMA_NAME, FIELD_NAME


def get_or_create_schema():
    """Get existing schema or create new one if it doesn't exist.
    
    Returns:
        DB.ExtensibleStorage.Schema: The pyArea schema
    """
    # Try to find existing schema
    schema_guid = System.Guid(SCHEMA_GUID)
    schema = DB.ExtensibleStorage.Schema.Lookup(schema_guid)
    
    if schema is not None:
        return schema
    
    # Create new schema
    schema_builder = DB.ExtensibleStorage.SchemaBuilder(schema_guid)
    schema_builder.SetSchemaName(SCHEMA_NAME)
    
    # Add single Data field (JSON string)
    schema_builder.AddSimpleField(FIELD_NAME, str)
    
    # Set read/write access
    schema_builder.SetReadAccessLevel(DB.ExtensibleStorage.AccessLevel.Public)
    schema_builder.SetWriteAccessLevel(DB.ExtensibleStorage.AccessLevel.Public)
    
    return schema_builder.Finish()


def set_data(element, data_dict):
    """Store data dictionary as JSON in element's extensible storage.
    
    Args:
        element: Revit element to store data on
        data_dict: Dictionary to store (will be JSON-encoded)
        
    Returns:
        bool: True if successful, False otherwise
    """
    if not element or not isinstance(data_dict, dict):
        return False
    
    try:
        schema = get_or_create_schema()
        entity = DB.ExtensibleStorage.Entity(schema)
        
        # Convert dict to JSON string
        json_string = json.dumps(data_dict, ensure_ascii=False)
        
        # Store in entity
        entity.Set[str](FIELD_NAME, json_string)
        
        # Save to element
        element.SetEntity(entity)
        
        return True
        
    except Exception as e:
        print("Error setting data: {}".format(e))
        return False


def get_data(element):
    """Retrieve data dictionary from element's extensible storage.
    
    Args:
        element: Revit element to retrieve data from
        
    Returns:
        dict: Stored data dictionary, or empty dict if no data found
    """
    if not element:
        return {}
    
    try:
        schema = get_or_create_schema()
        entity = element.GetEntity(schema)
        
        if not entity.IsValid():
            return {}
        
        # Get JSON string
        json_string = entity.Get[str](FIELD_NAME)
        
        if not json_string:
            return {}
        
        # Parse JSON to dict
        return json.loads(json_string)
        
    except Exception as e:
        print("Error getting data: {}".format(e))
        return {}


def delete_data(element):
    """Delete extensible storage data from element.
    
    Args:
        element: Revit element to delete data from
        
    Returns:
        bool: True if successful, False otherwise
    """
    if not element:
        return False
    
    try:
        schema = get_or_create_schema()
        element.DeleteEntity(schema)
        return True
        
    except Exception as e:
        print("Error deleting data: {}".format(e))
        return False


def has_data(element):
    """Check if element has pyArea extensible storage data.
    
    Args:
        element: Revit element to check
        
    Returns:
        bool: True if element has data, False otherwise
    """
    if not element:
        return False
    
    try:
        schema = get_or_create_schema()
        entity = element.GetEntity(schema)
        return entity.IsValid()
        
    except Exception as e:
        print("Error checking data: {}".format(e))
        return False
