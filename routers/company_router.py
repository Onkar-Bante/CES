# routers/company_router.py
from fastapi import APIRouter, Depends, HTTPException, UploadFile, File
from models.company import CompanyCreate, CompanyUpdateSalaryFormat
from services.company_service import create_company, update_salary_format, extract_and_update_salary_format
from utils.excel_utils import generate_excel_template
from fastapi.responses import StreamingResponse
from io import BytesIO
from bson.objectid import ObjectId
from database import get_company_collection

router = APIRouter()

@router.post("/create_company")
async def api_create_company(company: CompanyCreate):
    return await create_company(company)

@router.put("/update_salary_format/{company_id}")
async def api_update_salary_format(company_id: str, data: CompanyUpdateSalaryFormat):
    return await update_salary_format(company_id, data)

@router.post("/extract_salary_columns/{company_id}")
async def api_extract_salary_columns(company_id: str, file: UploadFile = File(...)):
    """
    Extract salary sheet columns from an uploaded Excel template
    and update the company's salary format
    """
    return await extract_and_update_salary_format(company_id, file)

@router.post("/upload_salary_template/{company_id}")
async def import_salary_sheet_with_formulas(company_id: str, file: UploadFile, expected_columns: list):
    """
    Import a salary sheet Excel file with formulas, dealing with potentially different headers.
    
    Args:
        company_id: ID of the company
        file: Uploaded Excel file
        expected_columns: List of expected column names
        
    Returns:
        Dict with success message
    """
    try:
        # Read the file
        contents = await file.read()
        file.file.seek(0)  # Reset file pointer
        
        # Use openpyxl to read the workbook with formulas
        from openpyxl import load_workbook
        
        excel_file = BytesIO(contents)
        workbook = load_workbook(excel_file, data_only=False)
        sheet = workbook.active
        
        # Try to find the header row
        header_row_idx = None
        
        # Search first 10 rows
        for i in range(1, 11):
            row_values = [cell.value for cell in sheet[i]]
            # Filter out None and empty strings
            row_values = [str(val).strip() for val in row_values if val is not None and str(val).strip()]
            
            # Check if this row might be headers (has enough non-empty cells)
            if len(row_values) >= 5:
                # Check if any common salary terms appear
                row_text = " ".join(row_values).lower()
                if any(term in row_text for term in ["salary", "basic", "hra", "gross", "deduction", "net"]):
                    header_row_idx = i
                    break
        
        if header_row_idx is None:
            # Default to row 3 if headers not found
            header_row_idx = 3
        
        # Map actual column indices to expected column names
        column_mapping = {}
        
        # Get actual headers from the identified row
        actual_headers = {}
        for col_idx in range(1, sheet.max_column + 1):
            cell = sheet.cell(row=header_row_idx, column=col_idx)
            if cell.value:
                actual_headers[col_idx] = str(cell.value).strip()
        
        # Try to map actual headers to expected columns
        for col_idx, header in actual_headers.items():
            header_lower = header.lower()
            
            # First try exact match
            exact_matches = [col for col in expected_columns if col.lower() == header_lower]
            if exact_matches:
                column_mapping[col_idx] = exact_matches[0]
                continue
                
            # Try partial matches
            partial_matches = []
            for expected_col in expected_columns:
                expected_lower = expected_col.lower()
                # Check if either contains the other
                if header_lower in expected_lower or expected_lower in header_lower:
                    partial_matches.append(expected_col)
            
            if len(partial_matches) == 1:
                column_mapping[col_idx] = partial_matches[0]
        
        # Extract formulas from first data row
        formulas = {}
        data_row_idx = header_row_idx + 1
        
        for col_idx, expected_col in column_mapping.items():
            cell = sheet.cell(row=data_row_idx, column=col_idx)
            if cell.data_type == 'f' and cell.value and str(cell.value).startswith('='):
                formula = str(cell.value)
                # Replace the specific row number with {row} template
                formula = formula.replace(str(data_row_idx), "{row}")
                formulas[expected_col] = formula
        
        # Update company document
        await get_company_collection().update_one(
            {"_id": ObjectId(company_id)},
            {"$set": {
                "salary_sheet_formulas": formulas
            }}
        )
        
        return {"message": "Successfully imported salary sheet formulas", "formulas_count": len(formulas)}
        
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"Error importing salary sheet with formulas: {str(e)}")

@router.get("/generate_template/{company_id}")
async def generate_template(company_id: str):
    from database import get_company_collection
    from bson.objectid import ObjectId

    company = await get_company_collection().find_one({"_id": ObjectId(company_id)})
    if not company:
        raise HTTPException(status_code=404, detail="Company not found")

    columns = company["salary_sheet_columns"]
    file_stream = generate_excel_template(columns)

    return StreamingResponse(
        file_stream,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment; filename={company_id}_template.xlsx"}
    )