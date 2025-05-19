# routers/employee_router.py
from fastapi import APIRouter, UploadFile, File, HTTPException, Body, Query
from services.employee_service import (
    upload_employees, 
    add_employee, 
    get_employees, 
    get_employee, 
    update_employee, 
    delete_employee,
    export_employees
)
from models.employee import EmployeeCreate, EmployeeUpdate
from typing import Dict, Any, List, Optional
from fastapi.responses import StreamingResponse
from datetime import datetime

router = APIRouter()

@router.post("/upload_employees/{company_id}")
async def api_upload_employees(company_id: str, file: UploadFile = File(...)):
    return await upload_employees(company_id, file)

@router.post("/employees")
async def api_add_employee(employee: EmployeeCreate):
    return await add_employee(employee)

@router.get("/employees/{company_id}")
async def api_get_employees(
    company_id: str, 
    skip: int = 0, 
    limit: int = 100,
    text_search: Optional[str] = None,
    emp_id: Optional[str] = None,
    name_contains: Optional[str] = None,
    email_contains: Optional[str] = None,
    designation: Optional[str] = None,
    site: Optional[str] = None,
    basic_pay_gte: Optional[float] = None,
    basic_pay_lte: Optional[float] = None,
    net_amt_gte: Optional[float] = None,
    net_amt_lte: Optional[float] = None
):
    # Convert query params to a filters dict
    filters = {
        "text_search": text_search,
        "EMP ID": emp_id,
        "Name of Employees_contains": name_contains,
        "Email_contains": email_contains,
        "Designation": designation,
        "Name of Site": site,
        "Basic Pay_gte": basic_pay_gte,
        "Basic Pay_lte": basic_pay_lte,
        "Net Amt_gte": net_amt_gte,
        "Net Amt_lte": net_amt_lte
    }
    
    # Remove None values
    filters = {k: v for k, v in filters.items() if v is not None}
    
    return await get_employees(company_id, skip, limit, filters)

@router.get("/employees/{company_id}/{employee_id}")
async def api_get_employee(company_id: str, employee_id: str):
    return await get_employee(company_id, employee_id)

@router.put("/employees/{company_id}/{employee_id}")
async def api_update_employee(company_id: str, employee_id: str, employee_data: EmployeeUpdate):
    return await update_employee(company_id, employee_id, employee_data)

@router.delete("/employees/{company_id}/{employee_id}")
async def api_delete_employee(company_id: str, employee_id: str):
    return await delete_employee(company_id, employee_id)

@router.get("/export_employees/{company_id}")
async def api_export_employees(
    company_id: str,
    text_search: Optional[str] = None,
    emp_id: Optional[str] = None,
    name_contains: Optional[str] = None,
    email_contains: Optional[str] = None,
    designation: Optional[str] = None,
    site: Optional[str] = None,
    basic_pay_gte: Optional[float] = None,
    basic_pay_lte: Optional[float] = None,
    net_amt_gte: Optional[float] = None,
    net_amt_lte: Optional[float] = None,
    year: Optional[int] = None,
    month: Optional[int] = None
):
    # Convert query params to a filters dict
    filters = {
        "text_search": text_search,
        "EMP ID": emp_id,
        "Name of Employees_contains": name_contains,
        "Email_contains": email_contains,
        "Designation": designation,
        "Name of Site": site,
        "Basic Pay_gte": basic_pay_gte,
        "Basic Pay_lte": basic_pay_lte,
        "Net Amt_gte": net_amt_gte,
        "Net Amt_lte": net_amt_lte
    }
    
    # Remove None values
    filters = {k: v for k, v in filters.items() if v is not None}
    
    # If year and month not provided, use current date
    if year is None or month is None:
        current_date = datetime.now()
        year = year or current_date.year
        month = month or current_date.month
    
    file_stream = await export_employees(company_id, filters, year, month)
    
    # Include month and year in the filename
    month_name = datetime(year, month, 1).strftime("%B")
    
    return StreamingResponse(
        file_stream,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment; filename=employees_{company_id}_{month_name}_{year}.xlsx"}
    )

@router.get("/download_sample_template/{company_id}")
async def download_sample_template(company_id: str):
    """
    Download a pre-filled sample template with example data to show the correct format.
    """
    from database import get_company_collection
    from bson.objectid import ObjectId
    from utils.excel_utils import generate_sample_template
    
    company = await get_company_collection().find_one({"_id": ObjectId(company_id)})
    if not company:
        raise HTTPException(status_code=404, detail="Company not found")
        
    columns = company["salary_sheet_columns"]
    file_stream = generate_sample_template(columns, company["name"])
    
    return StreamingResponse(
        file_stream,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment; filename={company_id}_sample_template.xlsx"}
    )