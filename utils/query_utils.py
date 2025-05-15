# utils/query_utils.py

from typing import Dict, Any
import math

def build_query_filters(filters: Dict[str, Any]) -> Dict[str, Any]:
    """
    Convert user-friendly filter parameters to MongoDB query syntax.
    
    Args:
        filters: Dictionary of filter parameters
        
    Returns:
        Dictionary with MongoDB query operators
    """
    mongo_query = {}
    
    # Handle text search separately if present
    if "text_search" in filters and filters["text_search"]:
        mongo_query["$text"] = {"$search": filters["text_search"]}
        filters.pop("text_search")
    
    # Process the rest of the filters
    for key, value in filters.items():
        # Skip None values and NaN values
        if value is None or (isinstance(value, float) and (math.isnan(value) or math.isinf(value))):
            continue
            
        if key.endswith("_contains"):
            # Handle contains operator (case insensitive substring match)
            field_name = key.replace("_contains", "")
            mongo_query[field_name] = {"$regex": value, "$options": "i"}
        
        elif key.endswith("_gte"):
            # Handle greater than or equal operator
            field_name = key.replace("_gte", "")
            if field_name in mongo_query:
                mongo_query[field_name]["$gte"] = value
            else:
                mongo_query[field_name] = {"$gte": value}
        
        elif key.endswith("_lte"):
            # Handle less than or equal operator
            field_name = key.replace("_lte", "")
            if field_name in mongo_query:
                mongo_query[field_name]["$lte"] = value
            else:
                mongo_query[field_name] = {"$lte": value}
        
        else:
            # Handle exact match
            mongo_query[key] = value
    
    return mongo_query

def try_convert_numeric(value: Any) -> Any:
    """
    Try to convert a value to int or float if possible.
    """
    if isinstance(value, (int, float)):
        return value
        
    if isinstance(value, str):
        # Try int first
        try:
            return int(value)
        except ValueError:
            # Then try float
            try:
                return float(value)
            except ValueError:
                pass
                
    return value

