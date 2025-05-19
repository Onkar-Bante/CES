# models/attendance.py
from pydantic import BaseModel
from typing import Dict, Any, List, Optional
from datetime import date

class AttendanceCreate(BaseModel):
    company_id: str
    employee_id: str
    date: date
    status: str  # "present", "absent", "half-day", "leave", etc.
    notes: Optional[str] = None

class AttendanceBulkCreate(BaseModel):
    company_id: str
    records: List[Dict[str, Any]]  # List of attendance records

class AttendanceUpdate(BaseModel):
    status: str
    notes: Optional[str] = None

class AttendanceFilterParams(BaseModel):
    company_id: str
    employee_id: Optional[str] = None
    start_date: Optional[date] = None
    end_date: Optional[date] = None
    status: Optional[str] = None

class AttendanceSummary(BaseModel):
    total_days: int
    present_days: int
    absent_days: int
    half_days: int
    leaves: int
