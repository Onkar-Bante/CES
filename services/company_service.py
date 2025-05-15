# services/company_service.py
from database import get_company_collection
from models.company import CompanyCreate, CompanyUpdateSalaryFormat
from bson.objectid import ObjectId
from fastapi import HTTPException, UploadFile
from utils.excel_extraction import extract_columns_from_excel, extract_formulas_from_excel

async def create_company(company_data: CompanyCreate):
    collection = get_company_collection()
    result = await collection.insert_one(company_data.dict())
    return {"message": "Company created", "company_id": str(result.inserted_id)}

async def update_salary_format(company_id: str, data: CompanyUpdateSalaryFormat):
    collection = get_company_collection()
    company = await collection.find_one({"_id": ObjectId(company_id)})
    if not company:
        raise HTTPException(status_code=404, detail="Company not found")

    await collection.update_one(
        {"_id": ObjectId(company_id)},
        {"$set": {"salary_sheet_columns": data.salary_sheet_columns}}
    )
    return {"message": "Salary format updated"}

async def extract_and_update_salary_format(company_id: str, file: UploadFile):
    """
    Extract column names from an Excel template and update company's salary format.
    
    Args:
        company_id: ID of the company
        file: Uploaded Excel template file
        
    Returns:
        Dictionary with success message and list of extracted columns
    """
    collection = get_company_collection()
    company = await collection.find_one({"_id": ObjectId(company_id)})
    if not company:
        raise HTTPException(status_code=404, detail="Company not found")
    
    # Extract columns from the Excel file
    columns = await extract_columns_from_excel(file)
    
    # Extract formulas
    formulas = await extract_formulas_from_excel(file)
    
    # Update company document
    await get_company_collection().update_one(
        {"_id": ObjectId(company_id)},
        {"$set": {
            "salary_sheet_columns": columns,
            "salary_sheet_formulas": formulas
        }}
    )
    
    return {
        "message": "Salary format updated from Excel template",
        "columns": columns
    }