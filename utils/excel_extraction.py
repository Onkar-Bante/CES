# utils/excel_extraction.py
import pandas as pd
from io import BytesIO
from fastapi import UploadFile, HTTPException

async def extract_columns_from_excel(file: UploadFile, header_row_index: int = None) -> dict:
    try:
        contents = await file.read()
        file.file.seek(0)
        excel_file = BytesIO(contents)

        indicators = ["sr", "emp", "id", "name", "email", "basic", "hra", "gross", "net", "tax", "deduction"]
        
        if header_row_index is not None:
            df = pd.read_excel(excel_file, header=header_row_index)
            cleaned_columns = clean_columns(df.columns)
            return {"columns": cleaned_columns, "header_row_index": header_row_index}

        preview_df = pd.read_excel(excel_file, header=None, nrows=10)
        candidates = []

        for i in range(len(preview_df)):
            row = preview_df.iloc[i]
            row_strs = row.fillna("").astype(str).str.lower()
            keyword_matches = sum(1 for cell in row_strs if any(ind in cell for ind in indicators))
            non_empty_cells = sum(1 for cell in row_strs if cell.strip() and "unnamed" not in cell)

            if keyword_matches >= 3 and non_empty_cells >= len(row) * 0.5:
                candidates.append((i, keyword_matches))

        if candidates:
            best_row = max(candidates, key=lambda x: x[1])[0]
            file.file.seek(0)
            df = pd.read_excel(file.file, header=best_row)
            cleaned_columns = clean_columns(df.columns)
            return {"columns": cleaned_columns, "header_row_index": best_row}

        # Fallback to row 3
        file.file.seek(0)
        df = pd.read_excel(file.file, header=2)
        cleaned_columns = clean_columns(df.columns)
        return {"columns": cleaned_columns, "header_row_index": 2}

    except Exception as e:
        raise HTTPException(status_code=400, detail=f"Failed to extract headers: {str(e)}")

def clean_columns(columns):
    cleaned = []
    for i, col in enumerate(columns):
        if pd.isna(col) or "unnamed" in str(col).lower():
            cleaned.append(f"Column_{i+1}")
        else:
            cleaned.append(str(col).strip())  # Already strips leading/trailing spaces
    return cleaned

    
async def extract_formulas_from_excel(file: UploadFile, header_row_index=2) -> dict:
    """
    Extract Excel formulas from an uploaded file.
    
    Args:
        file: Uploaded Excel file
        header_row_index: Index of the header row (default is 2 for row 3)
        
    Returns:
        Dictionary mapping column names to formula templates
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
        
        # Get headers (assumed to be at header_row_index, typically row 3)
        headers = {}
        for col_idx in range(1, sheet.max_column + 1):
            cell = sheet.cell(row=header_row_index + 1, column=col_idx)
            if cell.value:
                headers[col_idx] = str(cell.value).strip()
        
        # Check for formulas in the first data row (header_row_index + 1)
        formulas = {}
        first_data_row = header_row_index + 2  # Usually row 4
        
        for col_idx in range(1, sheet.max_column + 1):
            if col_idx in headers:
                cell = sheet.cell(row=first_data_row, column=col_idx)
                if cell.data_type == 'f' and cell.value and cell.value.startswith('='):
                    formula = cell.value
                    # Replace the specific row number with {row} template
                    formula = formula.replace(str(first_data_row), "{row}")
                    formulas[headers[col_idx]] = formula
        
        return formulas
        
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"Error extracting formulas: {str(e)}")