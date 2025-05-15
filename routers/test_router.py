# routers/test_router.py
from fastapi import APIRouter, UploadFile, File
from utils.excel_extraction import extract_columns_from_excel

router = APIRouter(tags=["testing"])

@router.post("/test_extract_columns")
async def test_extract_columns(file: UploadFile = File(...)):
    """
    Test endpoint to extract columns from an Excel file without saving them.
    This helps to verify how the extraction works on different file formats.
    """
    columns = await extract_columns_from_excel(file)
    return {"extracted_columns": columns}