# models/employee.py
from pydantic import BaseModel
from typing import Dict, Any, List, Optional

class EmployeeUploadRequest(BaseModel):
    file: str  # just placeholder; handled via UploadFile

class EmployeeCreate(BaseModel):
    # This is a dynamic model that will accept any fields
    company_id: str
    data: Dict[str, Any]

class EmployeeUpdate(BaseModel):
    data: Dict[str, Any]

class EmployeeFilterParams(BaseModel):
    company_id: str
    filters: Optional[Dict[str, Any]] = None