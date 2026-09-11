# 📊 Excel VBA Salary Slip Generator & PDF Automation

An Excel VBA automation project that generates individual employee salary slips and exports them as PDF files with a single click.

The project demonstrates how Excel VBA can automate repetitive payroll-related document generation using employee master data and a reusable salary slip template.

> **Note:** This repository uses sample/dummy employee information for demonstration purposes. No confidential, personal, client, or company data is included.

---

## 🚀 Project Overview

Preparing salary slips manually for a large number of employees can involve repetitive data entry, copy-paste operations, formatting, and PDF exporting.

This project automates that workflow.

The VBA macro:

1. Reads employee information from the `Employee_Master` worksheet.
2. Updates the `Salary_Slip` template.
3. Calculates salary components.
4. Populates employee details.
5. Exports the completed salary slip as a PDF.
6. Creates an output folder automatically.
7. Repeats the process for all employees.

---

## 🔄 Automation Workflow

```text
Employee Master Data
        ↓
Excel VBA Macro
        ↓
Read Employee Record
        ↓
Populate Salary Slip
        ↓
Calculate Salary
        ↓
Export as PDF
        ↓
Save in Salary Slips Folder
        ↓
Next Employee
```

---

## ✨ Features

* One-click salary slip generation
* Automatic PDF export
* Bulk processing of employees
* Excel VBA automation
* Dynamic employee information
* Salary calculation
* Automatic output folder creation
* Reusable salary slip template
* Progress status in Excel
* Sample employee data
* Modular VBA source code

---

## 🛠️ Technologies Used

| Technology      | Purpose                               |
| --------------- | ------------------------------------- |
| Microsoft Excel | Data management and template          |
| VBA             | Process automation                    |
| Excel Formulas  | Salary calculations                   |
| PDF Export      | Salary slip document generation       |
| Git             | Version control                       |
| GitHub          | Source code and project documentation |

---

## 📁 Project Structure

```text
excel-vba-salary-slip-automation/
│
├── README.md
│
├── Excel/
│   └── Salary_Slip_Automation.xlsm
│
├── VBA/
│   ├── GenerateSalarySlips.bas
│   ├── NumberToWords.bas
│   └── Utilities.bas
│
├── Screenshots/
│   ├── employee-master.png
│   ├── salary-slip.png
│   ├── dashboard.png
│   ├── vba-code.png
│   └── generated-pdfs.png
│
├── Sample_Data/
│   └── employee_master_sample.xlsx
│
├── .gitignore
└── LICENSE
```

---

## 📋 Workbook Structure

### 1. Employee_Master

Contains sample employee information such as:

* Employee Name
* Employee ID
* Department
* Designation
* Basic Salary
* Allowances
* Bonus
* Deductions
* Other payroll fields

Each employee is stored as one row.

---

### 2. Salary_Slip

This worksheet acts as the salary slip template.

Employee information is populated dynamically by the VBA macro.

The template can be customized according to the organization's requirements.

---

### 3. Dashboard

The dashboard provides a simple overview of the automation process, including:

* Total employees
* Generated salary slips
* Successful records
* Failed records
* Last execution status

---

### 4. Instructions

Contains instructions explaining how to use the workbook and generate salary slips.

---

## ⚙️ How to Run the Project

### Step 1 — Download the Repository

Clone the repository:

```bash
git clone YOUR_GITHUB_REPOSITORY_URL
```

---

### Step 2 — Open Excel

Open:

```text
Excel/Salary_Slip_Automation.xlsm
```

---

### Step 3 — Enable Macros

When Excel displays the security warning, enable macros/content if you trust the workbook.

---

### Step 4 — Add Employee Data

Open:

```text
Employee_Master
```

Enter or replace the sample employee records with your own test data.

For public GitHub demonstrations, use dummy data only.

---

### Step 5 — Generate Salary Slips

Click:

```text
GENERATE SALARY SLIPS
```

The VBA automation will process the employee records.

---

### Step 6 — Check Output

Generated PDF files will be stored in:

```text
Salary Slips/
```

Example:

```text
Salary Slips/
├── EMP001_Rahul Sharma.pdf
├── EMP002_Priya Patil.pdf
└── EMP003_Amit Joshi.pdf
```

---

## 💻 VBA Modules

The project is organized into separate VBA modules.

### GenerateSalarySlips.bas

Responsible for:

* Reading employee records
* Updating the salary slip
* Processing employees
* Exporting PDFs
* Managing the output folder

### NumberToWords.bas

Converts numeric salary amounts into words.

### Utilities.bas

Contains reusable helper procedures/functions used by the automation.

---

## 📸 Screenshots

### Employee Master

Add screenshot here:

```text




```

### Salary Slip Template

Add screenshot here:

```text

```

### Dashboard

Add screenshot here:

```text

```

### VBA Automation

Add screenshot here:

```text

```

### Generated PDFs

Add screenshot here:

```text

```

---

## 📈 Example Result

For 100 sample employees:

```text
Total Employees:       100
Salary Slips Generated: 100
Output Format:          PDF
Manual Export Required: No
```

The actual results depend on the employee data provided to the workbook.

---

## 🎯 Skills Demonstrated

This project demonstrates practical experience with:

* Excel
* VBA
* Excel Automation
* Data Processing
* Data Validation
* Formula-Based Calculations
* PDF Automation
* File Handling
* Reporting
* Dashboard Development
* Process Improvement
* Repetitive Task Automation
* Git
* GitHub

---

## 🔐 Data Privacy

This repository contains only dummy/sample information.

Do not upload:

* Real employee salary information
* PAN numbers
* Bank account details
* Aadhaar information
* Personal contact information
* Client information
* Company confidential information
* Internal business/process data

---

## 🔮 Future Improvements

Possible future enhancements include:

* UserForm for employee input
* Employee search functionality
* PDF email automation
* Password-protected worksheets
* Error logging
* Duplicate PDF detection
* Progress bar
* Monthly payroll selection
* Automated folder selection
* Dashboard charts
* CSV import
* Excel table-based data processing
* Outlook integration

---

## 👨‍💻 Author

**Rushikesh Avhale**

Interested in:

* Excel Automation
* VBA
* Data Analytics
* Power BI
* SQL
* Python
* Business Process Automation

---

## ⭐ Project Purpose

This project was created as a practical demonstration of Excel VBA automation and process improvement.

If you find the project useful, consider giving the repository a ⭐.
