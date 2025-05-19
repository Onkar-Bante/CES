# routers/attendance_router.py
from fastapi import APIRouter, HTTPException, Body, Query
from services.attendance_service import (
    add_attendance,
    bulk_add_attendance,
    get_attendance_records,
    get_attendance_summary,
    update_attendance,
    delete_attendance
)
from models.attendance_model import AttendanceCreate, AttendanceBulkCreate, AttendanceUpdate, AttendanceFilterParams
from typing import Optional
from datetime import date

router = APIRouter()

@router.post("/attendance")
async def api_add_attendance(attendance: AttendanceCreate):
    return await add_attendance(attendance)

@router.post("/attendance/bulk")
async def api_bulk_add_attendance(bulk_data: AttendanceBulkCreate):
    return await bulk_add_attendance(bulk_data)

@router.get("/attendance/{company_id}")
async def api_get_attendance_records(
    company_id: str,
    employee_id: Optional[str] = None,
    start_date: Optional[date] = None,
    end_date: Optional[date] = None,
    status: Optional[str] = None,
    skip: int = 0,
    limit: int = 100
):
    return await get_attendance_records(
        company_id, employee_id, start_date, end_date, status, skip, limit
    )

@router.get("/attendance/{company_id}/{employee_id}/summary")
async def api_get_attendance_summary(
    company_id: str,
    employee_id: str,
    year: int,
    month: int
):
    return await get_attendance_summary(company_id, employee_id, year, month)

@router.put("/attendance/{record_id}")
async def api_update_attendance(record_id: str, update_data: AttendanceUpdate):
    return await update_attendance(record_id, update_data)

@router.delete("/attendance/{record_id}")
async def api_delete_attendance(record_id: str):
    return await delete_attendance(record_id)
