#!/usr/bin/env python3
"""
Monthly Leave Calendar Excel Template Generator
Creates an Excel template with monthly calendar views for tracking team leaves
"""

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter
try:
    from openpyxl.worksheet.data_validation import DataValidation
except ImportError:
    from openpyxl.worksheet.datavalidation import DataValidation
from datetime import datetime, timedelta, date
import calendar
import os

def create_monthly_leave_calendar():
    """Create a comprehensive monthly leave calendar Excel template"""
    
    # Create workbook
    wb = Workbook()
    wb.remove(wb.active)
    
    # Get current year
    current_year = datetime.now().year
    
    # Create sheets for each month
    month_names = [calendar.month_name[i] for i in range(1, 13)]
    
    # Create overview sheet first
    overview_sheet = wb.create_sheet("Leave Overview", 0)
    
    # Create monthly sheets
    for month_num, month_name in enumerate(month_names, 1):
        sheet_name = f"{month_name} {current_year}"
        wb.create_sheet(sheet_name, month_num)
    
    # Add employee setup and legend sheets
    employee_sheet = wb.create_sheet("Employee Setup", 13)
    legend_sheet = wb.create_sheet("Legend & Settings", 14)
    
    # Define styles
    header_font = Font(bold=True, color="FFFFFF", size=12)
    header_fill = PatternFill(start_color="2F4F4F", end_color="2F4F4F", fill_type="solid")
    
    weekend_fill = PatternFill(start_color="F0F0F0", end_color="F0F0F0", fill_type="solid")
    today_fill = PatternFill(start_color="FFE4B5", end_color="FFE4B5", fill_type="solid")
    
    # Leave type colors
    leave_colors = {
        "AL": PatternFill(start_color="90EE90", end_color="90EE90", fill_type="solid"),  # Light Green - Annual Leave
        "SL": PatternFill(start_color="FFB6C1", end_color="FFB6C1", fill_type="solid"),  # Light Pink - Sick Leave
        "PL": PatternFill(start_color="87CEEB", end_color="87CEEB", fill_type="solid"),  # Sky Blue - Personal Leave
        "ML": PatternFill(start_color="DDA0DD", end_color="DDA0DD", fill_type="solid"),  # Plum - Maternity Leave
        "EL": PatternFill(start_color="F0E68C", end_color="F0E68C", fill_type="solid"),  # Khaki - Emergency Leave
        "TL": PatternFill(start_color="98FB98", end_color="98FB98", fill_type="solid"),  # Pale Green - Training Leave
    }
    
    border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    
    center_alignment = Alignment(horizontal='center', vertical='center')
    
    # ===== EMPLOYEE SETUP SHEET =====
    emp_headers = ["Employee ID", "Full Name", "Department", "Position", "Manager", "Employee Code"]
    
    for col, header in enumerate(emp_headers, 1):
        cell = employee_sheet.cell(row=1, column=col, value=header)
        cell.font = header_font
        cell.fill = header_fill
        cell.border = border
        cell.alignment = center_alignment
    
    # Set column widths for employee sheet
    emp_widths = [15, 25, 20, 25, 25, 15]
    for col, width in enumerate(emp_widths, 1):
        employee_sheet.column_dimensions[get_column_letter(col)].width = width
    
    # Add sample employee data
    sample_employees = [
        ["E001", "John Smith", "Engineering", "Senior Developer", "Jane Manager", "JS"],
        ["E002", "Sarah Johnson", "Marketing", "Marketing Specialist", "Mike Lead", "SJ"],
        ["E003", "Mike Davis", "HR", "HR Coordinator", "Lisa Director", "MD"],
        ["E004", "Lisa Wilson", "Finance", "Financial Analyst", "Tom CFO", "LW"],
        ["E005", "Tom Brown", "Engineering", "Software Engineer", "Jane Manager", "TB"],
        ["E006", "Emma Davis", "Marketing", "Marketing Manager", "Mike Lead", "ED"],
        ["E007", "James Wilson", "Operations", "Operations Coordinator", "Sam Director", "JW"],
        ["E008", "Anna Miller", "Finance", "Accountant", "Tom CFO", "AM"],
    ]
    
    for row, emp_data in enumerate(sample_employees, 2):
        for col, value in enumerate(emp_data, 1):
            cell = employee_sheet.cell(row=row, column=col, value=value)
            cell.border = border
    
    # ===== LEGEND & SETTINGS SHEET =====
    title_cell = legend_sheet.cell(row=1, column=1, value="LEAVE CALENDAR LEGEND & SETTINGS")
    title_cell.font = Font(bold=True, size=16, color="FFFFFF")
    title_cell.fill = header_fill
    title_cell.alignment = center_alignment
    legend_sheet.merge_cells('A1:F1')
    
    # Leave type legend
    legend_data = [
        ["", "", "", "", "", ""],
        ["Leave Type", "Code", "Color", "Description", "", ""],
        ["Annual Leave", "AL", "", "Paid vacation time", "", ""],
        ["Sick Leave", "SL", "", "Medical leave", "", ""],
        ["Personal Leave", "PL", "", "Personal time off", "", ""],
        ["Maternity Leave", "ML", "", "Maternity/Paternity leave", "", ""],
        ["Emergency Leave", "EL", "", "Emergency situations", "", ""],
        ["Training Leave", "TL", "", "Training/Development", "", ""],
        ["", "", "", "", "", ""],
        ["Instructions:", "", "", "", "", ""],
        ["1. Setup employees in 'Employee Setup' sheet", "", "", "", "", ""],
        ["2. Use employee codes (JS, SJ, etc.) in calendar", "", "", "", "", ""],
        ["3. Add leave type code after employee code", "", "", "", "", ""],
        ["4. Example: 'JS-AL' = John Smith on Annual Leave", "", "", "", "", ""],
        ["5. Weekend days are automatically highlighted", "", "", "", "", ""],
        ["6. Use dropdown validation for consistent entries", "", "", "", "", ""],
    ]
    
    for row, data in enumerate(legend_data, 1):
        for col, value in enumerate(data, 1):
            cell = legend_sheet.cell(row=row, column=col, value=value)
            if row == 2:  # Header row
                cell.font = header_font
                cell.fill = header_fill
            cell.border = border
            if col == 3 and row >= 3 and row <= 8:  # Color column
                leave_type = data[1]
                if leave_type in leave_colors:
                    cell.fill = leave_colors[leave_type]
    
    # Set column widths for legend sheet
    legend_widths = [20, 10, 10, 30, 20, 20]
    for col, width in enumerate(legend_widths, 1):
        legend_sheet.column_dimensions[get_column_letter(col)].width = width
    
    # ===== CREATE MONTHLY CALENDAR SHEETS =====
    def create_monthly_sheet(sheet, year, month):
        # Get calendar data
        cal = calendar.monthcalendar(year, month)
        month_name = calendar.month_name[month]
        
        # Title
        title_cell = sheet.cell(row=1, column=1, value=f"{month_name} {year} - Leave Calendar")
        title_cell.font = Font(bold=True, size=16, color="FFFFFF")
        title_cell.fill = header_fill
        title_cell.alignment = center_alignment
        sheet.merge_cells('A1:H1')
        
        # Day headers
        day_names = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday']
        for col, day in enumerate(day_names, 1):
            cell = sheet.cell(row=3, column=col, value=day)
            cell.font = header_font
            cell.fill = header_fill
            cell.border = border
            cell.alignment = center_alignment
        
        # Set column widths
        for col in range(1, 8):
            sheet.column_dimensions[get_column_letter(col)].width = 18
        
        # Create calendar grid
        row_start = 4
        for week_num, week in enumerate(cal):
            # Date row
            date_row = row_start + (week_num * 6)  # 6 rows per week (1 for dates, 5 for employees)
            
            for day_num, day in enumerate(week, 1):
                cell = sheet.cell(row=date_row, column=day_num)
                if day == 0:
                    cell.value = ""
                    cell.fill = PatternFill(start_color="E0E0E0", end_color="E0E0E0", fill_type="solid")
                else:
                    cell.value = day
                    cell.font = Font(bold=True, size=12)
                    
                    # Check if weekend
                    day_of_week = date(year, month, day).weekday()
                    if day_of_week >= 5:  # Saturday=5, Sunday=6
                        cell.fill = weekend_fill
                    
                    # Check if today
                    if (year == datetime.now().year and 
                        month == datetime.now().month and 
                        day == datetime.now().day):
                        cell.fill = today_fill
                
                cell.border = border
                cell.alignment = center_alignment
            
            # Employee rows (5 rows for employee entries)
            for emp_row in range(1, 6):
                actual_row = date_row + emp_row
                for day_num in range(1, 8):
                    cell = sheet.cell(row=actual_row, column=day_num)
                    if week[day_num-1] == 0:  # No date in this cell
                        cell.fill = PatternFill(start_color="E0E0E0", end_color="E0E0E0", fill_type="solid")
                    else:
                        # Check if weekend for styling
                        if week[day_num-1] != 0:
                            day_of_week = date(year, month, week[day_num-1]).weekday()
                            if day_of_week >= 5:
                                cell.fill = weekend_fill
                    
                    cell.border = border
                    cell.alignment = center_alignment
        
        # Add sample data for first month
        if month == 1:
            # Add some sample leave entries
            sample_leaves = [
                (1, 2, "JS-AL"),  # Week 1, Day 2, John Smith Annual Leave
                (1, 3, "SJ-PL"),  # Week 1, Day 3, Sarah Johnson Personal Leave
                (2, 1, "MD-SL"),  # Week 2, Day 1, Mike Davis Sick Leave
                (3, 4, "LW-AL"),  # Week 3, Day 4, Lisa Wilson Annual Leave
                (3, 5, "TB-TL"),  # Week 3, Day 5, Tom Brown Training Leave
            ]
            
            for week_idx, day_idx, leave_code in sample_leaves:
                if week_idx <= len(cal) and day_idx <= len(cal[week_idx-1]):
                    if cal[week_idx-1][day_idx-1] != 0:  # Valid date
                        # Find first empty employee row
                        date_row = row_start + ((week_idx-1) * 6)
                        for emp_row in range(1, 6):
                            cell = sheet.cell(row=date_row + emp_row, column=day_idx)
                            if cell.value is None:
                                cell.value = leave_code
                                # Apply color based on leave type
                                leave_type = leave_code.split('-')[-1]
                                if leave_type in leave_colors:
                                    cell.fill = leave_colors[leave_type]
                                break
        
        # Add validation for leave codes
        validation_formula = '"JS-AL,JS-SL,JS-PL,JS-ML,JS-EL,JS-TL,SJ-AL,SJ-SL,SJ-PL,SJ-ML,SJ-EL,SJ-TL,MD-AL,MD-SL,MD-PL,MD-ML,MD-EL,MD-TL,LW-AL,LW-SL,LW-PL,LW-ML,LW-EL,LW-TL,TB-AL,TB-SL,TB-PL,TB-ML,TB-EL,TB-TL,ED-AL,ED-SL,ED-PL,ED-ML,ED-EL,ED-TL,JW-AL,JW-SL,JW-PL,JW-ML,JW-EL,JW-TL,AM-AL,AM-SL,AM-PL,AM-ML,AM-EL,AM-TL"'
        
        leave_validation = DataValidation(
            type="list",
            formula1=validation_formula,
            showDropDown=True
        )
        sheet.add_data_validation(leave_validation)
        
        # Apply validation to all employee cells
        for week_num in range(len(cal)):
            date_row = row_start + (week_num * 6)
            for emp_row in range(1, 6):
                for day_col in range(1, 8):
                    cell_ref = f"{get_column_letter(day_col)}{date_row + emp_row}"
                    leave_validation.add(cell_ref)
        
        # Freeze panes
        sheet.freeze_panes = 'A4'
    
    # Create all monthly sheets
    for month_num, month_name in enumerate(month_names, 1):
        sheet = wb[f"{month_name} {current_year}"]
        create_monthly_sheet(sheet, current_year, month_num)
    
    # ===== OVERVIEW SHEET =====
    overview_title = overview_sheet.cell(row=1, column=1, value=f"TEAM LEAVE OVERVIEW - {current_year}")
    overview_title.font = Font(bold=True, size=16, color="FFFFFF")
    overview_title.fill = header_fill
    overview_title.alignment = center_alignment
    overview_sheet.merge_cells('A1:L1')
    
    # Monthly navigation
    nav_headers = ["Month", "Quick Navigation", "Total Leave Days", "Peak Leave Period"]
    for col, header in enumerate(nav_headers, 1):
        cell = overview_sheet.cell(row=3, column=col, value=header)
        cell.font = header_font
        cell.fill = header_fill
        cell.border = border
        cell.alignment = center_alignment
    
    # Add month navigation
    for month_num, month_name in enumerate(month_names, 1):
        row = 3 + month_num
        overview_sheet.cell(row=row, column=1, value=month_name).border = border
        overview_sheet.cell(row=row, column=2, value=f"Go to {month_name} sheet").border = border
        overview_sheet.cell(row=row, column=3, value="=COUNTIF(...)")  # Placeholder for formula
        overview_sheet.cell(row=row, column=4, value="TBD").border = border
    
    # Set column widths for overview
    overview_widths = [15, 25, 20, 25]
    for col, width in enumerate(overview_widths, 1):
        overview_sheet.column_dimensions[get_column_letter(col)].width = width
    
    # Add instructions
    instructions = [
        "",
        "HOW TO USE THE MONTHLY LEAVE CALENDAR:",
        "",
        "1. SETUP:",
        "   • Add your team members in the 'Employee Setup' sheet",
        "   • Note their employee codes (JS, SJ, etc.)",
        "",
        "2. ENTER LEAVES:",
        "   • Go to the specific month sheet",
        "   • Click on any cell under a date",
        "   • Enter: [Employee Code]-[Leave Type]",
        "   • Example: 'JS-AL' for John Smith Annual Leave",
        "",
        "3. LEAVE CODES:",
        "   • AL = Annual Leave (Green)",
        "   • SL = Sick Leave (Pink)", 
        "   • PL = Personal Leave (Blue)",
        "   • ML = Maternity Leave (Purple)",
        "   • EL = Emergency Leave (Yellow)",
        "   • TL = Training Leave (Light Green)",
        "",
        "4. VISUAL FEATURES:",
        "   • Weekends are gray shaded",
        "   • Today is highlighted in orange",
        "   • Leave types are color-coded",
        "   • Dropdown validation available",
    ]
    
    for i, instruction in enumerate(instructions, 17):
        if instruction:  # Only process non-empty instructions
            cell = overview_sheet.cell(row=i, column=1, value=instruction)
            if instruction.startswith(("HOW TO USE", "1. SETUP:", "2. ENTER", "3. LEAVE", "4. VISUAL")):
                cell.font = Font(bold=True, color="2F4F4F")
            overview_sheet.merge_cells(f'A{i}:L{i}')
    
    # Save the workbook
    filename = "/workspace/Monthly_Leave_Calendar.xlsx"
    wb.save(filename)
    
    return filename

def main():
    """Main function to create the monthly leave calendar template"""
    try:
        print("Creating Monthly Leave Calendar Excel Template...")
        filename = create_monthly_leave_calendar()
        print(f"✅ Calendar template created successfully: {filename}")
        print(f"\nTemplate includes:")
        print(f"• 12 monthly calendar sheets for {datetime.now().year}")
        print("• Visual calendar grid with date squares")
        print("• Employee setup and legend sheets")
        print("• Color-coded leave types")
        print("• Weekend and today highlighting")
        print("• Dropdown validation for leave entries")
        print("• Sample leave data for January")
        print("\nFeatures:")
        print("• Professional calendar layout")
        print("• Easy visual leave tracking")
        print("• Multiple employees per day support")
        print("• Automated weekend detection")
        print("• Color-coded leave types")
        print("• Comprehensive instructions included")
        
    except Exception as e:
        print(f"❌ Error creating calendar template: {str(e)}")
        return False
    
    return True

if __name__ == "__main__":
    main()