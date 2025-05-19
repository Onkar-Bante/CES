# utils/excel_utils.py

import pandas as pd
from io import BytesIO
import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
import calendar
from datetime import datetime

def validate_excel_columns(uploaded_columns, expected_columns):
    """
    Validates that the uploaded Excel has the expected columns.
    Allows for case insensitivity and extra columns.
    """
    uploaded_lower = [col.strip().lower() for col in uploaded_columns]
    expected_lower = [col.strip().lower() for col in expected_columns]
    
    # Check if all expected columns are present (case insensitive)
    for expected in expected_lower:
        if expected not in uploaded_lower:
            return False
    
    return True

def create_excel_from_employees_with_formulas(employees, columns, company_name, formula_mapping=None):
    """
    Creates an Excel file from employee data with formulas preserved.
    Now includes attendance data columns.
    
    Args:
        employees: List of employee dictionaries
        columns: List of column names
        company_name: Name of the company
        formula_mapping: Dictionary mapping column names to Excel formulas
        
    Returns:
        BytesIO stream containing the Excel file
    """
    # Create a new workbook
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Employees"
    
    # Add title row
    ws.merge_cells('A1:H1')
    title_cell = ws['A1']
    title_cell.value = f"{company_name} - Employee Data"
    title_cell.font = Font(bold=True, size=14)
    title_cell.alignment = Alignment(horizontal='center', vertical='center')
    
    # Add date row
    ws.merge_cells('A2:H2')
    date_cell = ws['A2']
    date_cell.value = f"Generated on: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
    date_cell.font = Font(italic=True)
    date_cell.alignment = Alignment(horizontal='center', vertical='center')
    
    # Define header styles
    header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True)
    border = Border(
        left=Side(border_style="thin"),
        right=Side(border_style="thin"),
        top=Side(border_style="thin"),
        bottom=Side(border_style="thin")
    )
    
    # Write header row
    for col_idx, col_name in enumerate(columns, 1):
        cell = ws.cell(row=4, column=col_idx)
        cell.value = col_name
        cell.fill = header_fill
        cell.font = header_font
        cell.border = border
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    
    # Auto-adjust column widths based on content
    for col_idx, column in enumerate(columns, 1):
        ws.column_dimensions[get_column_letter(col_idx)].width = max(12, min(30, len(column) + 2))
    
    # Process employees
    row_idx = 5
    for employee in employees:
        for col_idx, col_name in enumerate(columns, 1):
            cell = ws.cell(row=row_idx, column=col_idx)
            
            # Check if this column has a formula
            if formula_mapping and col_name in formula_mapping:
                formula = formula_mapping[col_name].format(row=row_idx)
                cell.value = formula
            else:
                # Regular cell value
                if col_name in employee:
                    cell.value = employee[col_name]
                elif col_name.lower() in [k.lower() for k in employee.keys()]:
                    # Case-insensitive match
                    key = next(k for k in employee.keys() if k.lower() == col_name.lower())
                    cell.value = employee[key]
            
            # Apply borders
            cell.border = border
            
            # Center-align certain columns
            if any(term in col_name.lower() for term in ["id", "total days", "days present", "days absent", "half days", "leaves"]):
                cell.alignment = Alignment(horizontal='center')
            
            # Right-align numeric columns
            if any(term in col_name.lower() for term in ["pay", "salary", "amount", "hra", "allowance", "deduction", "pf", "esi", "tax", "net"]):
                try:
                    if isinstance(cell.value, (int, float)) or (isinstance(cell.value, str) and cell.value.replace('.', '', 1).isdigit()):
                        cell.alignment = Alignment(horizontal='right')
                except:
                    pass
        
        row_idx += 1
    
    # Create special section for attendance information
    # Check if we have attendance columns
    attendance_cols = [col for col in columns if "attendance" in col.lower() or "days" in col.lower() or "present" in col.lower()]
    
    if attendance_cols:
        # Add attendance summary section
        summary_row = row_idx + 2
        ws.merge_cells(f'A{summary_row}:D{summary_row}')
        summary_cell = ws[f'A{summary_row}']
        summary_cell.value = "Attendance Summary"
        summary_cell.font = Font(bold=True, size=12)
        
        # Calculate totals
        total_days_col = next((i for i, col in enumerate(columns, 1) if "total days" in col.lower()), None)
        present_days_col = next((i for i, col in enumerate(columns, 1) if "days present" in col.lower()), None)
        absent_days_col = next((i for i, col in enumerate(columns, 1) if "days absent" in col.lower()), None)
        
        if all([total_days_col, present_days_col, absent_days_col]):
            # Add formulas to calculate totals
            total_row = summary_row + 2
            
            ws.cell(row=total_row, column=1).value = "Total Working Days:"
            ws.cell(row=total_row, column=1).font = Font(bold=True)
            
            # Use AVERAGE since all employees should have the same total days in a month
            ws.cell(row=total_row, column=2).value = f"=AVERAGE({get_column_letter(total_days_col)}5:{get_column_letter(total_days_col)}{row_idx-1})"
            
            ws.cell(row=total_row+1, column=1).value = "Total Present Days:"
            ws.cell(row=total_row+1, column=1).font = Font(bold=True)
            ws.cell(row=total_row+1, column=2).value = f"=SUM({get_column_letter(present_days_col)}5:{get_column_letter(present_days_col)}{row_idx-1})"
            
            ws.cell(row=total_row+2, column=1).value = "Total Absent Days:"
            ws.cell(row=total_row+2, column=1).font = Font(bold=True)
            ws.cell(row=total_row+2, column=2).value = f"=SUM({get_column_letter(absent_days_col)}5:{get_column_letter(absent_days_col)}{row_idx-1})"
            
            ws.cell(row=total_row+3, column=1).value = "Attendance Rate:"
            ws.cell(row=total_row+3, column=1).font = Font(bold=True)
            ws.cell(row=total_row+3, column=2).value = f"=B{total_row+1}/(B{total_row}*{row_idx-5})"
            ws.cell(row=total_row+3, column=2).number_format = "0.00%"
    
    # Save to BytesIO stream
    excel_stream = BytesIO()
    wb.save(excel_stream)
    excel_stream.seek(0)
    
    return excel_stream

def generate_sample_template(columns, company_name):
    """
    Generates a sample Excel template with the specified columns.
    Includes example data for reference.
    
    Args:
        columns: List of column names
        company_name: Name of the company
        
    Returns:
        BytesIO stream containing the Excel file
    """
    # Create a new workbook
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Template"
    
    # Add title row
    ws.merge_cells('A1:H1')
    title_cell = ws['A1']
    title_cell.value = f"{company_name} - Employee Upload Template"
    title_cell.font = Font(bold=True, size=14)
    title_cell.alignment = Alignment(horizontal='center', vertical='center')
    
    # Add instruction row
    ws.merge_cells('A2:H2')
    instruction_cell = ws['A2']
    instruction_cell.value = "Please fill in your employee data in the format shown below."
    instruction_cell.font = Font(italic=True)
    instruction_cell.alignment = Alignment(horizontal='center', vertical='center')
    
    # Define header styles
    header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True)
    border = Border(
        left=Side(border_style="thin"),
        right=Side(border_style="thin"),
        top=Side(border_style="thin"),
        bottom=Side(border_style="thin")
    )
    
    # Write header row
    for col_idx, col_name in enumerate(columns, 1):
        cell = ws.cell(row=4, column=col_idx)
        cell.value = col_name
        cell.fill = header_fill
        cell.font = header_font
        cell.border = border
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    
    # Auto-adjust column widths based on content
    for col_idx, column in enumerate(columns, 1):
        ws.column_dimensions[get_column_letter(col_idx)].width = max(12, min(30, len(column) + 2))
    
    # Add sample rows with example data
    sample_data = [
        {
            "EMP ID": "EMP001",
            "Name of Employees": "John Smith",
            "Email": "john.smith@example.com",
            "Designation": "Software Engineer",
            "Name of Site": "Headquarters",
            "Basic Pay": 5000,
            "HRA": 2000,
            "Conveyance Allowance": 800,
            "Special Allowance": 1200,
            "Gross Amount": "=SUM(F5:I5)",
            "EPF": 600,
            "ESI": 100,
            "PT": 200,
            "IT": 500,
            "Total Deduction": "=SUM(K5:N5)",
            "Net Amt": "=J5-O5"
        },
        {
            "EMP ID": "EMP002",
            "Name of Employees": "Jane Doe",
            "Email": "jane.doe@example.com",
            "Designation": "Project Manager",
            "Name of Site": "Branch Office",
            "Basic Pay": 7000,
            "HRA": 2800,
            "Conveyance Allowance": 1000,
            "Special Allowance": 1500,
            "Gross Amount": "=SUM(F6:I6)",
            "EPF": 840,
            "ESI": 140,
            "PT": 200,
            "IT": 800,
            "Total Deduction": "=SUM(K6:N6)",
            "Net Amt": "=J6-O6"
        }
    ]
    
    # Map sample data to configured columns
    row_idx = 5
    for sample in sample_data:
        for col_idx, col_name in enumerate(columns, 1):
            cell = ws.cell(row=row_idx, column=col_idx)
            
            # Look for a match in the sample data (case-insensitive)
            value = None
            for sample_key, sample_value in sample.items():
                if sample_key.lower() == col_name.lower() or sample_key.lower() in col_name.lower() or col_name.lower() in sample_key.lower():
                    value = sample_value
                    break
            
            cell.value = value
            cell.border = border
        
        row_idx += 1
    
    # Add note about attendance
    note_row = row_idx + 2
    ws.merge_cells(f'A{note_row}:H{note_row}')
    note_cell = ws[f'A{note_row}']
    note_cell.value = "Note: Attendance data should be managed through the attendance management system."
    note_cell.font = Font(italic=True)
    
    # Save to BytesIO stream
    excel_stream = BytesIO()
    wb.save(excel_stream)
    excel_stream.seek(0)
    
    return excel_stream