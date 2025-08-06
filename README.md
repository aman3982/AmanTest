# Team Leave Management Excel Templates

A comprehensive set of Excel templates for tracking employee leaves with both detailed tracking and visual calendar views.

## 📁 Available Templates

### 1. Team Leave Tracker (`Team_Leave_Tracker_Template.xlsx`)
Traditional tabular leave tracking with detailed approval workflow.

### 2. Monthly Leave Calendar (`Monthly_Leave_Calendar.xlsx`)
Visual monthly calendar view for easy leave visualization and planning.

## 🚀 Quick Start

### Option 1: Use Pre-generated Templates
Simply open either Excel file and start using immediately.

### Option 2: Generate New Templates
```bash
# Install dependencies
pip install -r requirements.txt

# Generate detailed leave tracker
python3 team_leave_tracker.py

# Generate monthly calendar view
python3 simple_monthly_calendar.py
```

## 📋 Template 1: Detailed Leave Tracker Features

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

## 📅 Template 2: Monthly Calendar Features

### 🔹 Monthly Calendar Sheets
- **Visual Calendar Layout**: Traditional calendar grid for each month
- **Daily Leave Tracking**: Multiple employees per day support
- **Color-Coded Leaves**: Instant visual identification of leave types
- **Weekend Highlighting**: Automatic weekend detection and styling
- **Today Indicator**: Current date highlighting

### 🔹 Employee Database
- Simple employee codes (JS, SJ, etc.)
- Full names and departments
- Manager information

### 🔹 Overview & Instructions
- Complete usage guide
- Leave type legend with colors
- Navigation instructions

## 🎨 Leave Type Color Coding

| Leave Type | Code | Color | Usage |
|------------|------|-------|-------|
| Annual Leave | AL | Green | Paid vacation time |
| Sick Leave | SL | Red | Medical leave |
| Personal Leave | PL | Teal | Personal time off |
| Maternity Leave | ML | Purple | Maternity/Paternity leave |
| Emergency Leave | EL | Orange | Emergency situations |
| Training Leave | TL | Blue | Training/Development |

## 📝 How to Use Monthly Calendar

1. **Setup**: Check employee codes in 'Employees' sheet
2. **Enter Leaves**: Click on calendar cells and enter `[Code]-[Type]`
   - Example: `JS-AL` = John Smith on Annual Leave
3. **Visual Tracking**: Leaves automatically color-coded
4. **Navigation**: Use sheet tabs for different months

## 📊 Template Comparison

| Feature | Leave Tracker | Monthly Calendar |
|---------|---------------|------------------|
| **View Type** | Table/List | Visual Calendar |
| **Best For** | Detailed tracking | Quick visualization |
| **Approval Workflow** | ✅ Full workflow | ❌ Basic only |
| **Visual Planning** | ❌ Limited | ✅ Excellent |
| **Data Analysis** | ✅ Comprehensive | ❌ Basic |
| **Team Overview** | ✅ Dashboard | ✅ Monthly view |
| **Sample Data** | ✅ Included | ✅ Included |

## 🛠️ Technical Requirements

- Microsoft Excel 2016 or later
- Python 3.7+ (for template generation)
- openpyxl and pandas libraries

## 📄 File Structure

```
/workspace/
├── Team_Leave_Tracker_Template.xlsx     # Detailed tracking template
├── Monthly_Leave_Calendar.xlsx          # Visual calendar template
├── team_leave_tracker.py               # Detailed tracker generator
├── simple_monthly_calendar.py          # Calendar generator
├── requirements.txt                     # Python dependencies
└── README.md                           # This documentation
```

## 🎯 Usage Recommendations

- **Use Leave Tracker for**: HR departments, formal approval processes, detailed reporting
- **Use Monthly Calendar for**: Team managers, visual planning, quick leave overview
- **Use Both**: Combine for comprehensive leave management system

## 🤝 Support

Both templates include built-in instructions and sample data. Modify the Python generator scripts to customize for your specific organizational needs.

---

**Created with ❤️ for efficient team leave management**