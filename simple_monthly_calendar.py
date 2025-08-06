#!/usr/bin/env python3
"""
Simple Monthly Leave Calendar Excel Template Generator
Creates an Excel template with monthly calendar views for tracking team leaves
"""

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter
try:
    from openpyxl.worksheet.data_validation import DataValidation
except ImportError:
    from openpyxl.worksheet.datavalidation import DataValidation
from datetime import datetime, date
import calendar

def create_simple_monthly_calendar():
    """Create a simple monthly leave calendar Excel template"""
    
    wb = Workbook()
    wb.remove(wb.active)
    
    current_year = datetime.now().year
    month_names = [calendar.month_name[i] for i in range(1, 13)]
    
    # Create sheets
    overview_sheet = wb.create_sheet("Overview", 0)
    
    # Create monthly sheets
    for month_num, month_name in enumerate(month_names, 1):
        sheet_name = f"{month_name[:3]} {current_year}"
        wb.create_sheet(sheet_name, month_num)
    
    employee_sheet = wb.create_sheet("Employees", 13)
    
    # Define styles
    header_font = Font(bold=True, color="FFFFFF", size=11)
    header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    weekend_fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
    today_fill = PatternFill(start_color="FFC000", end_color="FFC000", fill_type="solid")
    
    # Leave type colors
    leave_colors = {
        "AL": PatternFill(start_color="70AD47", end_color="70AD47", fill_type="solid"),  # Green
        "SL": PatternFill(start_color="FF6B6B", end_color="FF6B6B", fill_type="solid"),  # Red
        "PL": PatternFill(start_color="4ECDC4", end_color="4ECDC4", fill_type="solid"),  # Teal
        "ML": PatternFill(start_color="9B59B6", end_color="9B59B6", fill_type="solid"),  # Purple
        "EL": PatternFill(start_color="F39C12", end_color="F39C12", fill_type="solid"),  # Orange
        "TL": PatternFill(start_color="3498DB", end_color="3498DB", fill_type="solid"),  # Blue
    }
    
    border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    
    center_alignment = Alignment(horizontal='center', vertical='center')
    
    # ===== EMPLOYEE SHEET =====
    emp_headers = ["Employee Code", "Full Name", "Department", "Manager"]
    
    for col, header in enumerate(emp_headers, 1):
        cell = employee_sheet.cell(row=1, column=col, value=header)
        cell.font = header_font
        cell.fill = header_fill
        cell.border = border
        cell.alignment = center_alignment
    
    # Set column widths
    emp_widths = [15, 25, 20, 25]
    for col, width in enumerate(emp_widths, 1):
        employee_sheet.column_dimensions[get_column_letter(col)].width = width
    
    # Sample employees
    sample_employees = [
        ["JS", "John Smith", "Engineering", "Jane Manager"],
        ["SJ", "Sarah Johnson", "Marketing", "Mike Lead"],
        ["MD", "Mike Davis", "HR", "Lisa Director"],
        ["LW", "Lisa Wilson", "Finance", "Tom CFO"],
        ["TB", "Tom Brown", "Engineering", "Jane Manager"],
        ["ED", "Emma Davis", "Marketing", "Mike Lead"],
        ["JW", "James Wilson", "Operations", "Sam Director"],
        ["AM", "Anna Miller", "Finance", "Tom CFO"],
    ]
    
    for row, emp_data in enumerate(sample_employees, 2):
        for col, value in enumerate(emp_data, 1):
            cell = employee_sheet.cell(row=row, column=col, value=value)
            cell.border = border
    
    # ===== OVERVIEW SHEET =====
    overview_sheet.cell(row=1, column=1, value="MONTHLY LEAVE CALENDAR OVERVIEW").font = Font(bold=True, size=16)
    
    overview_info = [
        ["", ""],
        ["How to Use:", ""],
        ["1. Enter leaves using: [Employee Code]-[Leave Type]", ""],
        ["2. Example: JS-AL = John Smith Annual Leave", ""],
        ["", ""],
        ["Leave Types:", "Colors:"],
        ["AL = Annual Leave", "Green"],
        ["SL = Sick Leave", "Red"],
        ["PL = Personal Leave", "Teal"],
        ["ML = Maternity Leave", "Purple"],
        ["EL = Emergency Leave", "Orange"],
        ["TL = Training Leave", "Blue"],
        ["", ""],
        ["Navigation:", ""],
        ["• Click on month tabs at bottom", ""],
        ["• Check 'Employees' sheet for codes", ""],
        ["• Weekends are highlighted in gray", ""],
        ["• Today is highlighted in yellow", ""],
    ]
    
    for row, (col1, col2) in enumerate(overview_info, 3):
        overview_sheet.cell(row=row, column=1, value=col1)
        overview_sheet.cell(row=row, column=2, value=col2)
        if col1.endswith(":"):
            overview_sheet.cell(row=row, column=1).font = Font(bold=True)
    
    overview_sheet.column_dimensions['A'].width = 35
    overview_sheet.column_dimensions['B'].width = 20
    
    # ===== CREATE MONTHLY SHEETS =====
    def create_monthly_sheet(sheet, year, month):
        cal = calendar.monthcalendar(year, month)
        month_name = calendar.month_name[month]
        
        # Title
        title_cell = sheet.cell(row=1, column=1, value=f"{month_name} {year}")
        title_cell.font = Font(bold=True, size=16)
        
        # Day headers
        day_names = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun']
        for col, day in enumerate(day_names, 1):
            cell = sheet.cell(row=3, column=col, value=day)
            cell.font = header_font
            cell.fill = header_fill
            cell.border = border
            cell.alignment = center_alignment
        
        # Set column widths
        for col in range(1, 8):
            sheet.column_dimensions[get_column_letter(col)].width = 20
        
        # Create calendar
        current_row = 4
        for week in cal:
            # Date row
            for day_idx, day in enumerate(week, 1):
                cell = sheet.cell(row=current_row, column=day_idx)
                if day == 0:
                    cell.value = ""
                    cell.fill = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
                else:
                    cell.value = f"Day {day}"
                    cell.font = Font(bold=True, size=10)
                    
                    # Weekend check
                    if day_idx in [6, 7]:  # Saturday, Sunday
                        cell.fill = weekend_fill
                    
                    # Today check
                    if (year == datetime.now().year and 
                        month == datetime.now().month and 
                        day == datetime.now().day):
                        cell.fill = today_fill
                
                cell.border = border
                cell.alignment = center_alignment
            
            current_row += 1
            
            # Leave entry rows (3 rows for employees per week)
            for emp_row in range(3):
                for day_idx in range(1, 8):
                    cell = sheet.cell(row=current_row, column=day_idx)
                    if week[day_idx-1] == 0:
                        cell.fill = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
                    else:
                        if day_idx in [6, 7]:  # Weekend styling
                            cell.fill = weekend_fill
                    
                    cell.border = border
                    cell.alignment = center_alignment
                
                current_row += 1
            
            current_row += 1  # Space between weeks
        
        # Add sample data for January
        if month == 1:
            sheet.cell(row=5, column=2, value="JS-AL").fill = leave_colors["AL"]
            sheet.cell(row=6, column=3, value="SJ-PL").fill = leave_colors["PL"]
            sheet.cell(row=11, column=1, value="MD-SL").fill = leave_colors["SL"]
            sheet.cell(row=18, column=4, value="LW-AL").fill = leave_colors["AL"]
        
        # Add data validation
        validation_options = []
        for emp_code in ["JS", "SJ", "MD", "LW", "TB", "ED", "JW", "AM"]:
            for leave_type in ["AL", "SL", "PL", "ML", "EL", "TL"]:
                validation_options.append(f"{emp_code}-{leave_type}")
        
        validation_formula = f'"{",".join(validation_options)}"'
        
        if len(validation_formula) < 255:  # Excel limit
            leave_validation = DataValidation(
                type="list",
                formula1=validation_formula,
                showDropDown=True
            )
            sheet.add_data_validation(leave_validation)
            
            # Apply to employee cells (rows 5 onwards, excluding date rows)
            for row in range(5, current_row):
                if not any(sheet.cell(row=row, column=1).value and 
                          str(sheet.cell(row=row, column=1).value).startswith("Day")):
                    for col in range(1, 8):
                        cell_ref = f"{get_column_letter(col)}{row}"
                        leave_validation.add(cell_ref)
    
    # Create all monthly sheets
    for month_num, month_name in enumerate(month_names, 1):
        sheet = wb[f"{month_name[:3]} {current_year}"]
        create_monthly_sheet(sheet, current_year, month_num)
    
    # Save workbook
    filename = "/workspace/Monthly_Leave_Calendar.xlsx"
    wb.save(filename)
    return filename

def main():
    try:
        print("Creating Monthly Leave Calendar...")
        filename = create_simple_monthly_calendar()
        print(f"✅ Calendar created successfully: {filename}")
        print("\nFeatures:")
        print("• 12 monthly calendar sheets")
        print("• Visual calendar layout")
        print("• Color-coded leave types")
        print("• Weekend highlighting")
        print("• Employee database")
        print("• Sample data included")
        print("• Data validation dropdowns")
        
    except Exception as e:
        print(f"❌ Error: {str(e)}")
        return False
    
    return True

if __name__ == "__main__":
    main()