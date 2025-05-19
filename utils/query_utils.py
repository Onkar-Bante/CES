# utils/query_utils.py

from typing import Dict, Any
import math

def build_query_filters(filters):
    """
    Builds MongoDB query filters from filter parameters.
    Handles special cases for text search and comparison operators.
    """
    mongo_filters = {}
    
    # Handle full-text search (applied across multiple fields)
    if "text_search" in filters and filters["text_search"]:
        # MongoDB $text search requires a text index
        # As an alternative, we'll implement a simple regex search across common fields
        search_term = filters.pop("text_search")
        search_regex = {"$regex": search_term, "$options": "i"}
        
        mongo_filters["$or"] = [
            {"name_of_employees": search_regex},
            {"email": search_regex},
            {"emp_id": search_regex},
            {"designation": search_regex},
            {"name_of_site": search_regex}
        ]
    
    # Process the remaining filters
    for key, value in filters.items():
        # Skip empty filters
        if value is None:
            continue
        
        # Handle contains operator (case-insensitive)
        if key.endswith("_contains"):
            base_field = key.replace("_contains", "")
            mongo_filters[base_field] = {"$regex": value, "$options": "i"}
            
        # Handle greater than or equal
        elif key.endswith("_gte"):
            base_field = key.replace("_gte", "")
            if base_field not in mongo_filters:
                mongo_filters[base_field] = {}
            mongo_filters[base_field]["$gte"] = value
            
        # Handle less than or equal
        elif key.endswith("_lte"):
            base_field = key.replace("_lte", "")
            if base_field not in mongo_filters:
                mongo_filters[base_field] = {}
            mongo_filters[base_field]["$lte"] = value
            
        # Handle exact match (case-sensitive)
        else:
            mongo_filters[key] = value
    
    return mongo_filters

def build_attendance_query_filters(filters):
    """
    Builds MongoDB query filters specifically for attendance records.
    Handles date ranges and attendance statuses.
    """
    mongo_filters = {}
    
    # Process company_id and employee_id directly
    if "company_id" in filters:
        mongo_filters["company_id"] = filters["company_id"]
    
    if "employee_id" in filters:
        mongo_filters["employee_id"] = filters["employee_id"]
    
    # Handle date range filters
    if "start_date" in filters and filters["start_date"]:
        start_date = filters["start_date"].isoformat()
        if "date" not in mongo_filters:
            mongo_filters["date"] = {}
        mongo_filters["date"]["$gte"] = start_date
    
    if "end_date" in filters and filters["end_date"]:
        end_date = filters["end_date"].isoformat()
        if "date" not in mongo_filters:
            mongo_filters["date"] = {}
        mongo_filters["date"]["$lte"] = end_date
    
    # Handle status filter (exact match)
    if "status" in filters and filters["status"]:
        mongo_filters["status"] = filters["status"]
    
    return mongo_filters


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

