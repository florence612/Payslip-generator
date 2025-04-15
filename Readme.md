# Payslip Generator and Email Sender

This Python script reads employee salary data from an Excel file, calculates net salaries, generates professional-looking PDF payslips for each employee, and sends them via email.

---

## Features

- Reads employee data from `employees.xlsx`
- Automatically calculates Net Salary = Basic Salary + Allowance - Deductions
- Generates PDF payslips branded with **Tank Logistics**
- Payslips are signed by **Tatenda Karindi**
- Sends payslips to employee emails using Gmail

---

## Requirements

Make sure Python is installed, then install the required libraries:

```bash
pip install pandas reportlab yagmail openpyxl
