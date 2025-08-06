#!/usr/bin/env python3
"""
Team Leave Tracker Excel Template Generator
Creates a comprehensive Excel template for tracking employee leaves
"""

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils.dataframe import dataframe_to_rows
try:
    from openpyxl.worksheet.data_validation import DataValidation
except ImportError:
    from openpyxl.worksheet.datavalidation import DataValidation
from datetime import datetime, timedelta
import os

def create_leave_tracker_template():
    """Create a comprehensive leave tracking Excel template"""
    
    # Create workbook with multiple sheets
    wb = Workbook()
    
    # Remove default sheet
    wb.remove(wb.active)
    
    # Create sheets
    leaves_sheet = wb.create_sheet("Leave Requests", 0)
    summary_sheet = wb.create_sheet("Leave Summary", 1)
    employees_sheet = wb.create_sheet("Employee List", 2)
    leave_types_sheet = wb.create_sheet("Leave Types", 3)
    
    # Define styles
    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
    border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    center_alignment = Alignment(horizontal='center', vertical='center')
    
    # ===== LEAVE REQUESTS SHEET =====
    leaves_headers = [
        "Request ID", "Employee ID", "Employee Name", "Department", 
        "Leave Type", "Start Date", "End Date", "Days Requested",
        "Reason", "Status", "Applied Date", "Approved By", "Approved Date", "Comments"
    ]
    
    # Add headers
    for col, header in enumerate(leaves_headers, 1):
        cell = leaves_sheet.cell(row=1, column=col, value=header)
        cell.font = header_font
        cell.fill = header_fill
        cell.border = border
        cell.alignment = center_alignment
    
    # Set column widths
    column_widths = [12, 12, 20, 15, 15, 12, 12, 12, 25, 12, 12, 15, 12, 25]
    for col, width in enumerate(column_widths, 1):
        leaves_sheet.column_dimensions[leaves_sheet.cell(row=1, column=col).column_letter].width = width
    
    # Add sample data
    sample_data = [
        ["REQ001", "EMP001", "John Smith", "Engineering", "Annual Leave", "2024-03-15", "2024-03-20", 6, "Family vacation", "Approved", "2024-03-01", "Jane Manager", "2024-03-02", "Enjoy your vacation!"],
        ["REQ002", "EMP002", "Sarah Johnson", "Marketing", "Sick Leave", "2024-03-10", "2024-03-12", 3, "Flu symptoms", "Approved", "2024-03-10", "Mike Lead", "2024-03-10", "Get well soon"],
        ["REQ003", "EMP003", "Mike Davis", "HR", "Personal Leave", "2024-03-25", "2024-03-25", 1, "Personal appointment", "Pending", "2024-03-20", "", "", ""],
    ]
    
    for row, data in enumerate(sample_data, 2):
        for col, value in enumerate(data, 1):
            cell = leaves_sheet.cell(row=row, column=col, value=value)
            cell.border = border
            if col in [6, 7, 11, 13]:  # Date columns
                cell.number_format = 'DD/MM/YYYY'
    
    # Add data validation for Status column (column 10)
    status_validation = DataValidation(
        type="list",
        formula1='"Pending,Approved,Rejected,Cancelled"',
        showDropDown=True
    )
    leaves_sheet.add_data_validation(status_validation)
    status_validation.add(f"J2:J1000")
    
    # ===== EMPLOYEE LIST SHEET =====
    emp_headers = ["Employee ID", "Full Name", "Department", "Position", "Manager", "Email", "Start Date", "Annual Leave Entitlement"]
    
    for col, header in enumerate(emp_headers, 1):
        cell = employees_sheet.cell(row=1, column=col, value=header)
        cell.font = header_font
        cell.fill = header_fill
        cell.border = border
        cell.alignment = center_alignment
    
    # Set column widths for employee sheet
    emp_widths = [12, 20, 15, 20, 20, 25, 12, 20]
    for col, width in enumerate(emp_widths, 1):
        employees_sheet.column_dimensions[employees_sheet.cell(row=1, column=col).column_letter].width = width
    
    # Add sample employee data
    emp_data = [
        ["EMP001", "John Smith", "Engineering", "Senior Developer", "Jane Manager", "john.smith@company.com", "2023-01-15", 25],
        ["EMP002", "Sarah Johnson", "Marketing", "Marketing Specialist", "Mike Lead", "sarah.johnson@company.com", "2023-03-01", 20],
        ["EMP003", "Mike Davis", "HR", "HR Coordinator", "Lisa Director", "mike.davis@company.com", "2022-06-10", 22],
        ["EMP004", "Lisa Wilson", "Finance", "Financial Analyst", "Tom CFO", "lisa.wilson@company.com", "2023-02-20", 20],
    ]
    
    for row, data in enumerate(emp_data, 2):
        for col, value in enumerate(data, 1):
            cell = employees_sheet.cell(row=row, column=col, value=value)
            cell.border = border
            if col == 7:  # Start Date column
                cell.number_format = 'DD/MM/YYYY'
    
    # ===== LEAVE TYPES SHEET =====
    leave_type_headers = ["Leave Type", "Description", "Max Days Per Year", "Carry Forward", "Requires Approval"]
    
    for col, header in enumerate(leave_type_headers, 1):
        cell = leave_types_sheet.cell(row=1, column=col, value=header)
        cell.font = header_font
        cell.fill = header_fill
        cell.border = border
        cell.alignment = center_alignment
    
    # Set column widths for leave types sheet
    leave_type_widths = [15, 30, 18, 15, 18]
    for col, width in enumerate(leave_type_widths, 1):
        leave_types_sheet.column_dimensions[leave_types_sheet.cell(row=1, column=col).column_letter].width = width
    
    # Add leave type data
    leave_type_data = [
        ["Annual Leave", "Paid vacation time", 25, "Yes", "Yes"],
        ["Sick Leave", "Medical leave for illness", 10, "No", "No"],
        ["Personal Leave", "Personal time off", 5, "No", "Yes"],
        ["Maternity Leave", "Maternity/Paternity leave", 90, "No", "Yes"],
        ["Emergency Leave", "Unexpected emergency situations", 3, "No", "Yes"],
        ["Study Leave", "Educational purposes", 10, "No", "Yes"],
    ]
    
    for row, data in enumerate(leave_type_data, 2):
        for col, value in enumerate(data, 1):
            cell = leave_types_sheet.cell(row=row, column=col, value=value)
            cell.border = border
    
    # ===== LEAVE SUMMARY SHEET =====
    summary_sheet.merge_cells('A1:E1')
    title_cell = summary_sheet.cell(row=1, column=1, value="TEAM LEAVE SUMMARY DASHBOARD")
    title_cell.font = Font(bold=True, size=16, color="FFFFFF")
    title_cell.fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
    title_cell.alignment = center_alignment
    
    # Add summary sections
    summary_sections = [
        ("Employee", "Total Days Taken", "Remaining Annual Leave", "Pending Requests", "Last Leave Date"),
    ]
    
    # Add headers for summary
    for col, header in enumerate(summary_sections[0], 1):
        cell = summary_sheet.cell(row=3, column=col, value=header)
        cell.font = header_font
        cell.fill = header_fill
        cell.border = border
        cell.alignment = center_alignment
    
    # Set column widths for summary sheet
    summary_widths = [20, 18, 22, 18, 16]
    for col, width in enumerate(summary_widths, 1):
        summary_sheet.column_dimensions[summary_sheet.cell(row=3, column=col).column_letter].width = width
    
    # Add instructions
    instructions = [
        "",
        "INSTRUCTIONS FOR USE:",
        "1. Enter employee information in 'Employee List' sheet",
        "2. Configure leave types in 'Leave Types' sheet",
        "3. Record leave requests in 'Leave Requests' sheet",
        "4. Monitor team leave status in 'Leave Summary' sheet",
        "",
        "FEATURES:",
        "• Dropdown validation for leave status",
        "• Automatic date formatting",
        "• Multiple leave types supported",
        "• Employee database integration",
        "• Summary dashboard for managers",
    ]
    
    for i, instruction in enumerate(instructions, 6):
        cell = summary_sheet.cell(row=i, column=1, value=instruction)
        if instruction.startswith(("INSTRUCTIONS", "FEATURES")):
            cell.font = Font(bold=True, color="366092")
        summary_sheet.merge_cells(f'A{i}:E{i}')
    
    # Freeze panes on main sheets
    leaves_sheet.freeze_panes = 'A2'
    employees_sheet.freeze_panes = 'A2'
    leave_types_sheet.freeze_panes = 'A2'
    
    # Save the workbook
    filename = "/workspace/Team_Leave_Tracker_Template.xlsx"
    wb.save(filename)
    
    return filename

def main():
    """Main function to create the leave tracker template"""
    try:
        print("Creating Team Leave Tracker Excel Template...")
        filename = create_leave_tracker_template()
        print(f"✅ Template created successfully: {filename}")
        print("\nTemplate includes:")
        print("• Leave Requests - Main tracking sheet")
        print("• Employee List - Employee database")
        print("• Leave Types - Leave category definitions")
        print("• Leave Summary - Dashboard for managers")
        print("\nFeatures:")
        print("• Pre-formatted with professional styling")
        print("• Data validation for consistent entries")
        print("• Sample data for immediate use")
        print("• Multiple leave types supported")
        print("• Approval workflow tracking")
        
    except Exception as e:
        print(f"❌ Error creating template: {str(e)}")
        return False
    
    return True

if __name__ == "__main__":
    main()