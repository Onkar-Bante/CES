# services/attendance_service.py
from database import get_company_collection, get_employee_collection
from bson.objectid import ObjectId
from fastapi import HTTPException
from datetime import datetime, date, timedelta
import calendar
from typing import Dict, Any, List, Optional
from models.attendance_model import AttendanceCreate, AttendanceUpdate, AttendanceSummary

def get_attendance_collection():
    from database import db
    return db["attendance"]

async def add_attendance(attendance_data: AttendanceCreate):
    """Add a single attendance record"""
    # Validate company exists
    company = await get_company_collection().find_one({"_id": ObjectId(attendance_data.company_id)})
    if not company:
        raise HTTPException(status_code=404, detail="Company not found")
    
    # Validate employee exists
    employee = await get_employee_collection().find_one({
        "_id": ObjectId(attendance_data.employee_id),
        "company_id": attendance_data.company_id
    })
    if not employee:
        raise HTTPException(status_code=404, detail="Employee not found")
    
    # Check if attendance already exists for this date
    existing = await get_attendance_collection().find_one({
        "company_id": attendance_data.company_id,
        "employee_id": attendance_data.employee_id,
        "date": attendance_data.date.isoformat()
    })
    
    if existing:
        # Update existing record
        await get_attendance_collection().update_one(
            {"_id": existing["_id"]},
            {"$set": {
                "status": attendance_data.status,
                "notes": attendance_data.notes,
                "updated_at": datetime.now()
            }}
        )
        return {"message": "Attendance record updated", "record_id": str(existing["_id"])}
    else:
        # Create new record
        attendance_doc = {
            "company_id": attendance_data.company_id,
            "employee_id": attendance_data.employee_id,
            "date": attendance_data.date.isoformat(),
            "status": attendance_data.status,
            "notes": attendance_data.notes,
            "created_at": datetime.now(),
            "updated_at": datetime.now()
        }
        
        result = await get_attendance_collection().insert_one(attendance_doc)
        return {
            "message": "Attendance record added successfully",
            "record_id": str(result.inserted_id)
        }

async def bulk_add_attendance(bulk_data):
    """Add multiple attendance records at once"""
    company_id = bulk_data.company_id
    
    # Validate company exists
    company = await get_company_collection().find_one({"_id": ObjectId(company_id)})
    if not company:
        raise HTTPException(status_code=404, detail="Company not found")
    
    inserted = 0
    updated = 0
    errors = []
    
    for record in bulk_data.records:
        try:
            # Validate employee exists
            employee = await get_employee_collection().find_one({
                "_id": ObjectId(record["employee_id"]),
                "company_id": company_id
            })
            if not employee:
                errors.append(f"Employee {record['employee_id']} not found")
                continue
            
            # Convert date string to date object if needed
            if isinstance(record["date"], str):
                try:
                    record_date = date.fromisoformat(record["date"])
                except ValueError:
                    record_date = datetime.strptime(record["date"], "%Y-%m-%d").date()
            else:
                record_date = record["date"]
                
            # Check if attendance already exists for this date
            existing = await get_attendance_collection().find_one({
                "company_id": company_id,
                "employee_id": record["employee_id"],
                "date": record_date.isoformat()
            })
            
            if existing:
                # Update existing record
                await get_attendance_collection().update_one(
                    {"_id": existing["_id"]},
                    {"$set": {
                        "status": record["status"],
                        "notes": record.get("notes"),
                        "updated_at": datetime.now()
                    }}
                )
                updated += 1
            else:
                # Create new record
                attendance_doc = {
                    "company_id": company_id,
                    "employee_id": record["employee_id"],
                    "date": record_date.isoformat(),
                    "status": record["status"],
                    "notes": record.get("notes"),
                    "created_at": datetime.now(),
                    "updated_at": datetime.now()
                }
                
                await get_attendance_collection().insert_one(attendance_doc)
                inserted += 1
                
        except Exception as e:
            errors.append(f"Error processing record for employee {record.get('employee_id')}: {str(e)}")
    
    return {
        "message": "Bulk attendance update completed",
        "inserted": inserted,
        "updated": updated,
        "errors": errors
    }

async def get_attendance_records(
    company_id: str,
    employee_id: Optional[str] = None,
    start_date: Optional[date] = None,
    end_date: Optional[date] = None,
    status: Optional[str] = None,
    skip: int = 0,
    limit: int = 100
):
    """Get attendance records with optional filtering"""
    # Validate company exists
    company = await get_company_collection().find_one({"_id": ObjectId(company_id)})
    if not company:
        raise HTTPException(status_code=404, detail="Company not found")
    
    # Build query
    query = {"company_id": company_id}
    
    if employee_id:
        query["employee_id"] = employee_id
        
    if start_date and end_date:
        query["date"] = {
            "$gte": start_date.isoformat(),
            "$lte": end_date.isoformat()
        }
    elif start_date:
        query["date"] = {"$gte": start_date.isoformat()}
    elif end_date:
        query["date"] = {"$lte": end_date.isoformat()}
        
    if status:
        query["status"] = status
    
    # Execute query
    cursor = get_attendance_collection().find(query).sort("date", -1).skip(skip).limit(limit)
    records = await cursor.to_list(length=limit)
    
    # Get total count for pagination
    total_count = await get_attendance_collection().count_documents(query)
    
    # Convert ObjectId to string
    for record in records:
        record["_id"] = str(record["_id"])
    
    return {"records": records, "total": total_count}

async def get_attendance_summary(
    company_id: str,
    employee_id: str,
    year: int,
    month: int
):
    """Get monthly attendance summary for an employee"""
    # Validate company exists
    company = await get_company_collection().find_one({"_id": ObjectId(company_id)})
    if not company:
        raise HTTPException(status_code=404, detail="Company not found")
    
    # Validate employee exists
    employee = await get_employee_collection().find_one({
        "_id": ObjectId(employee_id),
        "company_id": company_id
    })
    if not employee:
        raise HTTPException(status_code=404, detail="Employee not found")
    
    # Calculate month start and end dates
    days_in_month = calendar.monthrange(year, month)[1]
    start_date = date(year, month, 1)
    end_date = date(year, month, days_in_month)
    
    # Get all attendance records for the month
    records = await get_attendance_collection().find({
        "company_id": company_id,
        "employee_id": employee_id,
        "date": {
            "$gte": start_date.isoformat(),
            "$lte": end_date.isoformat()
        }
    }).to_list(length=None)
    
    # Calculate summary
    present_days = sum(1 for r in records if r["status"] == "present")
    absent_days = sum(1 for r in records if r["status"] == "absent")
    half_days = sum(1 for r in records if r["status"] == "half-day")
    leaves = sum(1 for r in records if r["status"] == "leave")
    
    # Handle days without records as absent
    dates_with_records = set(date.fromisoformat(r["date"]) for r in records)
    missing_days = 0
    
    for day in range(1, days_in_month + 1):
        current_date = date(year, month, day)
        if current_date not in dates_with_records:
            missing_days += 1
    
    # Add missing days to absent total
    absent_days += missing_days
    
    summary = {
        "total_days": days_in_month,
        "present_days": present_days,
        "absent_days": absent_days,
        "half_days": half_days,
        "leaves": leaves,
        "missing_records": missing_days
    }
    
    return summary

async def update_attendance(record_id: str, update_data: AttendanceUpdate):
    """Update an existing attendance record"""
    # Check if record exists
    record = await get_attendance_collection().find_one({"_id": ObjectId(record_id)})
    if not record:
        raise HTTPException(status_code=404, detail="Attendance record not found")
    
    try:
        await get_attendance_collection().update_one(
            {"_id": ObjectId(record_id)},
            {"$set": {
                "status": update_data.status,
                "notes": update_data.notes,
                "updated_at": datetime.now()
            }}
        )
        return {"message": "Attendance record updated successfully"}
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

async def delete_attendance(record_id: str):
    """Delete an attendance record"""
    # Check if record exists
    record = await get_attendance_collection().find_one({"_id": ObjectId(record_id)})
    if not record:
        raise HTTPException(status_code=404, detail="Attendance record not found")
    
    try:
        await get_attendance_collection().delete_one({"_id": ObjectId(record_id)})
        return {"message": "Attendance record deleted successfully"}
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

async def get_employee_attendance_for_export(
    company_id: str,
    employee_id: str,
    year: int,
    month: int
):
    """Get attendance data for a specific employee and month for export"""
    summary = await get_attendance_summary(company_id, employee_id, year, month)
    
    # Add month name for better readability
    month_name = datetime(year, month, 1).strftime("%B")
    
    return {
        "month": month_name,
        "year": year,
        "total_days": summary["total_days"],
        "present_days": summary["present_days"],
        "absent_days": summary["absent_days"],
        "half_days": summary["half_days"],
        "leaves": summary["leaves"]
    }
