# Team Leave Tracker Excel Template

A comprehensive Excel template for tracking employee leave requests with professional formatting and automated features.

## 🚀 Quick Start

### Option 1: Use the Pre-generated Template
Simply open `Team_Leave_Tracker_Template.xlsx` in Excel and start using it immediately.

### Option 2: Generate a New Template
```bash
# Install dependencies
pip install -r requirements.txt

# Generate the template
python3 team_leave_tracker.py
```

## 📋 Template Features

### 🔹 Leave Requests Sheet
- **Request ID**: Unique identifier for each leave request
- **Employee Information**: ID, Name, Department
- **Leave Details**: Type, Start/End dates, Days requested
- **Approval Workflow**: Status tracking with dropdown validation
- **Comments**: Space for additional notes and feedback

### 🔹 Employee List Sheet
- Complete employee database
- Department and position information
- Manager assignments
- Annual leave entitlements
- Contact information

### 🔹 Leave Types Sheet
- Configurable leave categories
- Maximum days per year settings
- Carry-forward policies
- Approval requirements

### 🔹 Leave Summary Dashboard
- Team overview and statistics
- Instructions for use
- Quick reference guide

## 🎨 Template Highlights

- **Professional Styling**: Clean, corporate appearance with branded colors
- **Data Validation**: Dropdown menus for consistent data entry
- **Sample Data**: Pre-filled examples to get started quickly
- **Frozen Headers**: Easy navigation in large datasets
- **Automatic Formatting**: Date fields and proper column sizing
- **Multiple Leave Types**: Annual, Sick, Personal, Maternity, Emergency, Study

## 📊 Leave Types Included

| Leave Type | Max Days/Year | Carry Forward | Approval Required |
|------------|---------------|---------------|-------------------|
| Annual Leave | 25 | Yes | Yes |
| Sick Leave | 10 | No | No |
| Personal Leave | 5 | No | Yes |
| Maternity Leave | 90 | No | Yes |
| Emergency Leave | 3 | No | Yes |
| Study Leave | 10 | No | Yes |

## 🔧 Customization

The template is fully customizable:
- Add/modify leave types in the "Leave Types" sheet
- Update employee information in the "Employee List" sheet
- Adjust approval workflows as needed
- Customize departments and positions

## 📝 Usage Instructions

1. **Setup**: Enter your team's information in the "Employee List" sheet
2. **Configure**: Adjust leave types and policies in the "Leave Types" sheet
3. **Track**: Record leave requests in the "Leave Requests" sheet
4. **Monitor**: Use the "Leave Summary" sheet for team overview

## 🛠️ Technical Requirements

- Microsoft Excel 2016 or later
- Python 3.7+ (for template generation)
- openpyxl and pandas libraries

## 📄 File Structure

```
/workspace/
├── Team_Leave_Tracker_Template.xlsx    # Ready-to-use Excel template
├── team_leave_tracker.py               # Template generator script
├── requirements.txt                     # Python dependencies
└── README.md                           # This documentation
```

## 🤝 Support

For questions or customization requests, please refer to the instructions within the Excel template or modify the Python generator script to meet your specific needs.

---

**Created with ❤️ for efficient team leave management**