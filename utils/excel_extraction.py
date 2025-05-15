# utils/excel_extraction.py
import pandas as pd
from io import BytesIO
from fastapi import UploadFile, HTTPException

async def extract_columns_from_excel(file: UploadFile) -> list:
    """
    Extract column names from an Excel file, with improved header detection.
    
    Args:
        file: Uploaded Excel file
        
    Returns:
        List of column names
    """
    try:
        # Read the excel file
        contents = await file.read()
        file.file.seek(0)  # Reset file pointer for potential further use
        
        # Create a BytesIO object
        excel_file = BytesIO(contents)
        
        # Read first few rows to inspect content
        df_preview = pd.read_excel(excel_file, header=None, nrows=10)
        excel_file.seek(0)
        
        # Look for a row with standard salary sheet headers
        # Common indicators in salary sheet headers
        indicators = ["sr", "emp", "id", "name", "email", "basic", "hra", "gross", "net", "tax", "deduction"]
        
        header_row_idx = None
        best_match_count = 0
        
        # Check each row for potential headers
        for i in range(min(10, len(df_preview))):
            row = df_preview.iloc[i].astype(str)
            # Convert to string and lowercase for comparison
            row_text = " ".join(row.str.lower())
            
            # Count how many indicators appear in this row
            match_count = sum(1 for ind in indicators if ind in row_text)
            
            if match_count > best_match_count:
                best_match_count = match_count
                header_row_idx = i
        
        # If we found a good header row
        if header_row_idx is not None and best_match_count >= 3:
            # Use this row as header
            df = pd.read_excel(excel_file, header=header_row_idx)
            
            # Clean up column names
            columns = []
            for col in df.columns:
                if pd.isna(col) or "unnamed" in str(col).lower():
                    columns.append(f"Column_{len(columns)+1}")
                else:
                    columns.append(str(col).strip())
            
            return columns
        
        # If no good header row found, try the most common locations (row 3)
        excel_file.seek(0)
        try:
            df = pd.read_excel(excel_file, header=2)  # Row 3 (index 2) is common for headers
            columns = []
            for col in df.columns:
                if pd.isna(col) or "unnamed" in str(col).lower():
                    columns.append(f"Column_{len(columns)+1}")
                else:
                    columns.append(str(col).strip())
            
            if any(ind in " ".join(col.lower() for col in columns) for ind in indicators):
                return columns
        except:
            pass
        
        # Last resort - try reading with openpyxl to get more detailed cell information
        from openpyxl import load_workbook
        
        excel_file.seek(0)
        wb = load_workbook(filename=excel_file)
        ws = wb.active
        
        # Try row 3 (common for salary sheets with titles)
        header_row = []
        for cell in ws[3]:
            val = cell.value
            if val is not None:
                header_row.append(str(val).strip())
            else:
                header_row.append(f"Column_{len(header_row)+1}")
        
        if header_row and len(header_row) > 5:  # At least 5 columns
            return header_row
            
        # If all else fails
        raise HTTPException(
            status_code=400, 
            detail="Could not identify proper column headers. Please ensure your Excel file has headers in one of the first 10 rows."
        )
    
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"Error processing Excel file: {str(e)}")
    

    
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