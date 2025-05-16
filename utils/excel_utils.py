# utils/excel_utils.py
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill
from io import BytesIO
from datetime import datetime
import calendar

def validate_excel_columns(actual_cols, expected_cols):
    def norm(c): return c.strip().lower()
    actual_norm = [norm(c) for c in actual_cols]
    expected_norm = [norm(c) for c in expected_cols]
    return set(actual_norm) == set(expected_norm)


def generate_excel_template(columns):
    df = pd.DataFrame(columns=columns)
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False)
        
        # Get the worksheet
        worksheet = writer.sheets['Sheet1']
        
        # Format header row
        for idx, col in enumerate(columns, 1):
            cell = worksheet.cell(row=1, column=idx)
            cell.font = Font(bold=True)
            cell.alignment = Alignment(horizontal='center')
            cell.fill = PatternFill(start_color="D3D3D3", end_color="D3D3D3", fill_type="solid")
            
            # Auto-adjust column width based on content
            worksheet.column_dimensions[cell.column_letter].width = max(15, len(col) + 2)
    
    output.seek(0)
    return output

def create_excel_from_employees_with_formulas(employees, columns, company_name, formula_mapping=None):
    """
    Create an Excel file from employees data that includes formulas for calculated fields.
    
    Args:
        employees: List of employee documents
        columns: List of column names
        company_name: Name of the company
        formula_mapping: Dictionary mapping column names to Excel formula templates
        
    Returns:
        BytesIO containing the Excel file
    """
    import pandas as pd
    from openpyxl import Workbook
    from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
    from io import BytesIO
    from datetime import datetime
    import calendar
    
    # Create a DataFrame with only the specified columns in the correct order
    data = []
    for emp in employees:
        row = {}
        for col in columns:
            # Case-insensitive key lookup
            col_lower = col.lower()
            matching_keys = [k for k in emp.keys() if k.lower() == col_lower]
            row[col] = emp[matching_keys[0]] if matching_keys else None
        data.append(row)
    
    df = pd.DataFrame(data, columns=columns)
    
    # Create Excel file
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # Add title rows
        current_date = datetime.now()
        month_year = f"{calendar.month_name[current_date.month]}- {str(current_date.year)[2:]}"
        
        # Write to excel but starting from row 3 to leave space for the header
        df.to_excel(writer, index=False, startrow=2)
        
        # Get the worksheet
        worksheet = writer.sheets['Sheet1']
        
        # Add company name
        worksheet.merge_cells('A1:D1')
        cell = worksheet.cell(row=1, column=1)
        cell.value = f"{company_name} SALARY SHEET"
        cell.font = Font(bold=True, size=14)
        cell.alignment = Alignment(horizontal='left')
        
        # Add month-year
        worksheet.merge_cells('E1:G1')
        cell = worksheet.cell(row=1, column=5)
        cell.value = month_year
        cell.font = Font(bold=True, size=12)
        cell.alignment = Alignment(horizontal='left')
        
        # Style header row (row 3)
        for idx, col in enumerate(columns, 1):
            cell = worksheet.cell(row=3, column=idx)
            cell.font = Font(bold=True)
            cell.alignment = Alignment(horizontal='center')
            cell.fill = PatternFill(start_color="D3D3D3", end_color="D3D3D3", fill_type="solid")
            
            # Add a border
            thin_border = Border(
                left=Side(style='thin'), 
                right=Side(style='thin'), 
                top=Side(style='thin'), 
                bottom=Side(style='thin')
            )
            cell.border = thin_border
            
            # Auto-adjust column width based on content
            max_length = len(col)
            for row_idx in range(4, len(df) + 4):  # +4 because data starts at row 4
                cell_value = worksheet.cell(row=row_idx, column=idx).value
                if cell_value:
                    max_length = max(max_length, len(str(cell_value)))
            
            worksheet.column_dimensions[cell.column_letter].width = max(15, max_length + 2)
        
        # Apply formulas if formula_mapping is provided
        if formula_mapping:
            for row_idx in range(4, len(df) + 4):  # +4 because data starts at row 4
                for col_name, formula_template in formula_mapping.items():
                    # Find the column index for this formula
                    col_idx = None
                    for i, col in enumerate(columns, 1):
                        if col.lower() == col_name.lower():
                            col_idx = i
                            break
                    
                    if col_idx:
                        # Replace {row} in formula with current row
                        formula = formula_template.replace("{row}", str(row_idx))
                        
                        # Apply the formula to the cell
                        cell = worksheet.cell(row=row_idx, column=col_idx)
                        cell.value = formula
                        cell.data_type = 'f'  # Set as formula
    
    output.seek(0)
    return output

def generate_sample_template(columns, company_name):
    """
    Generate a sample Excel template with example data to demonstrate the correct format.
    
    Args:
        columns: List of column names
        company_name: Name of the company
        
    Returns:
        BytesIO stream containing the Excel file
    """
    import pandas as pd
    from openpyxl import Workbook
    from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
    from io import BytesIO
    from datetime import datetime
    import calendar
    
    # Create sample data for the template
    sample_data = []
    
    # Define some sample values for common columns
    sample_values = {
        "sr. no.": [1, 2, 3],
        "emp id": ["EMP001", "EMP002", "EMP003"],
        "name of employees": ["John Doe", "Jane Smith", "Alex Johnson"],
        "email": ["john@example.com", "jane@example.com", "alex@example.com"],
        "designation": ["Software Engineer", "HR Manager", "Project Manager"],
        "name of site": ["Main Office", "Branch Office", "Remote"],
        "no. of days in a month": [30, 30, 30],
        "no. of days present": [22, 21, 20],
        "basic pay": [30000, 40000, 50000],
        "hra": [15000, 20000, 25000],
        "conv allow/ vehicle reimb": [3000, 3500, 4000],
        "child education allow": [2000, 2000, 2000],
        "lta monthly": [2500, 3000, 3500],
        "medical monthly": [1250, 1500, 1750],
        "attire reimb": [1000, 1000, 1000],
        "sp.all": [2000, 2500, 3000],
        "other allow": [1000, 1200, 1500],
        "gross amount": [57750, 74700, 91750],
        "prof. tax": [200, 200, 200],
        "esic": [0, 0, 0],
        "p. f. cont.": [1800, 1800, 1800],
        "tds": [4000, 6000, 8000],
        "advance/ other deductions": [0, 0, 0],
        "total ded": [6000, 8000, 10000],
        "net amt": [51750, 66700, 81750],
        "other reimbursements": [0, 0, 0],
        "pf on arrears": [0, 0, 0],
        "bonus": [0, 0, 0],
        "other": [0, 0, 0],
        "payable": [51750, 66700, 81750]
    }
    
    # Create 3 sample records
    for i in range(3):
        record = {}
        for col in columns:
            col_lower = col.lower()
            # Try to find a matching sample value
            found = False
            for key, values in sample_values.items():
                if key in col_lower or col_lower in key:
                    record[col] = values[i]
                    found = True
                    break
            
            # If no match found, use a default value
            if not found:
                if "amount" in col_lower or "pay" in col_lower or "amt" in col_lower:
                    record[col] = 1000 * (i + 1)
                elif "date" in col_lower:
                    record[col] = f"2023-05-{15 + i}"
                elif "name" in col_lower:
                    record[col] = f"Sample Name {i+1}"
                elif "id" in col_lower:
                    record[col] = f"ID{100 + i}"
                else:
                    record[col] = f"Sample {i+1}"
        
        sample_data.append(record)
    
    # Create DataFrame with sample data
    df = pd.DataFrame(sample_data)
    
    # Create Excel file
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # Add current date
        current_date = datetime.now()
        month_year = f"{calendar.month_name[current_date.month]}- {str(current_date.year)[2:]}"
        
        # Create the Excel workbook
        workbook = writer.book
        
        # Add a sheet for instructions
        instructions = workbook.create_sheet("Instructions", 0)
        instructions.column_dimensions['A'].width = 100
        
        # Add instructions
        instructions['A1'] = "HOW TO USE THIS TEMPLATE"
        instructions['A1'].font = Font(bold=True, size=14)
        
        instructions['A3'] = "1. This is a SAMPLE template showing the required format for your salary sheet."
        instructions['A4'] = "2. The sheet 'Sample Data' shows what your data should look like."
        instructions['A5'] = "3. Use the 'Template' sheet to enter your actual employee data."
        instructions['A6'] = "4. IMPORTANT: Do not change the column headers or their order."
        instructions['A7'] = "5. The first row must contain the company name and month."
        instructions['A8'] = "6. The third row must contain the column headers exactly as shown."
        instructions['A9'] = "7. Data should start from the fourth row."
        instructions['A10'] = "8. When uploading, ensure all required columns have data."
        
        # Create a sample data sheet
        df.to_excel(writer, sheet_name='Sample Data', index=False, startrow=2)
        sample_sheet = writer.sheets['Sample Data']
        
        # Add company name and month in the sample sheet
        sample_sheet.merge_cells('A1:D1')
        cell = sample_sheet.cell(row=1, column=1)
        cell.value = f"{company_name} SALARY SHEET"
        cell.font = Font(bold=True, size=14)
        cell.alignment = Alignment(horizontal='left')
        
        # Add month-year
        sample_sheet.merge_cells('E1:G1')
        cell = sample_sheet.cell(row=1, column=5)
        cell.value = month_year
        cell.font = Font(bold=True, size=12)
        cell.alignment = Alignment(horizontal='left')
        
        # Style the header row
        for idx, col in enumerate(columns, 1):
            cell = sample_sheet.cell(row=3, column=idx)
            cell.font = Font(bold=True)
            cell.alignment = Alignment(horizontal='center')
            cell.fill = PatternFill(start_color="D3D3D3", end_color="D3D3D3", fill_type="solid")
            
            # Add a border
            thin_border = Border(
                left=Side(style='thin'), 
                right=Side(style='thin'), 
                top=Side(style='thin'), 
                bottom=Side(style='thin')
            )
            cell.border = thin_border
            
            # Set column width
            sample_sheet.column_dimensions[cell.column_letter].width = max(15, len(col) + 2)
        
        # Add an empty template sheet
        empty_df = pd.DataFrame(columns=columns)
        empty_df.to_excel(writer, sheet_name='Template', index=False, startrow=2)
        template_sheet = writer.sheets['Template']
        
        # Add company name and month in the template sheet
        template_sheet.merge_cells('A1:D1')
        cell = template_sheet.cell(row=1, column=1)
        cell.value = f"{company_name} SALARY SHEET"
        cell.font = Font(bold=True, size=14)
        cell.alignment = Alignment(horizontal='left')
        
        # Add month-year
        template_sheet.merge_cells('E1:G1')
        cell = template_sheet.cell(row=1, column=5)
        cell.value = month_year
        cell.font = Font(bold=True, size=12)
        cell.alignment = Alignment(horizontal='left')
        
        # Style the header row in template sheet
        for idx, col in enumerate(columns, 1):
            cell = template_sheet.cell(row=3, column=idx)
            cell.font = Font(bold=True)
            cell.alignment = Alignment(horizontal='center')
            cell.fill = PatternFill(start_color="D3D3D3", end_color="D3D3D3", fill_type="solid")
            
            # Add a border
            thin_border = Border(
                left=Side(style='thin'), 
                right=Side(style='thin'), 
                top=Side(style='thin'), 
                bottom=Side(style='thin')
            )
            cell.border = thin_border
            
            # Set column width
            template_sheet.column_dimensions[cell.column_letter].width = max(15, len(col) + 2)
    
    output.seek(0)
    return output