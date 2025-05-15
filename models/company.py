# models/company.py
from pydantic import BaseModel
from typing import List, Optional

class CompanyCreate(BaseModel):
    name: str
    gstn: str
    location: str
    holidays: List[str]
    working_days: List[str]
    salary_sheet_columns: Optional[List[str]] = None  # Now optional

class CompanyUpdateSalaryFormat(BaseModel):
    salary_sheet_columns: List[str]