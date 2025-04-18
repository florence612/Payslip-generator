
import pandas as pd
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib import colors
from reportlab.lib.units import inch
import os

# === Step 1: Load Excel and Process Data ===
file_path = 'employees.xlsx'

# Read Excel data
df = pd.read_excel(file_path, sheet_name='Sheet1')

# Clean column headers
df.columns = ['Employee ID', 'Employee Name', 'Department', 'Email', 'Basic Salary', 'Allowance', 'Deductions']

# Calculate Net Salary
df['Net Salary'] = df['Basic Salary'] + df['Allowance'] - df['Deductions']

# === Step 2: Payslip Generation ===
output_folder = 'payslips'
os.makedirs(output_folder, exist_ok=True)

def generate_payslip(row):
    emp_id = row['Employee ID']
    emp_name = row['Employee Name']
    filename = f"{output_folder}/{emp_id}.pdf"
    c = canvas.Canvas(filename, pagesize=A4)
    width, height = A4

    # Header
    c.setFont("Helvetica-Bold", 18)
    c.drawCentredString(width / 2, height - 60, "Tank Logistics")
    c.setFont("Helvetica-Bold", 14)
    c.drawCentredString(width / 2, height - 85, "Employee Payslip")

    # Employee Info
    c.setFont("Helvetica", 12)
    y = height - 130
    c.drawString(50, y, f"Employee Name: {emp_name}")
    c.drawString(330, y, f"Employee ID: {emp_id}")
    y -= 30
    c.drawString(50, y, f"Department: {row['Department']}")
    y -= 30
    c.drawString(50, y, f"Email: {row['Email']}")
    y -= 40

    # Draw salary breakdown table
    c.setFont("Helvetica-Bold", 12)
    c.drawString(50, y, "Earnings & Deductions")

    y -= 20
    c.setFont("Helvetica", 12)
    c.line(50, y, 500, y)
    y -= 25

    c.drawString(60, y, "Basic Salary")
    c.drawRightString(500, y, f"${row['Basic Salary']:.2f}")
    y -= 20
    c.drawString(60, y, "Allowances")
    c.drawRightString(500, y, f"${row['Allowance']:.2f}")
    y -= 20
    c.drawString(60, y, "Deductions")
    c.drawRightString(500, y, f"-${row['Deductions']:.2f}")
    y -= 20

    c.line(50, y, 500, y)
    y -= 25

    # Net Salary
    c.setFont("Helvetica-Bold", 12)
    c.drawString(60, y, "Net Salary")
    c.drawRightString(500, y, f"${row['Net Salary']:.2f}")
    y -= 50

    # Signature
    c.setFont("Helvetica-Oblique", 11)
    c.drawString(50, y, "Authorized by: Tatenda Karindi")

    # Footer
    c.setFont("Helvetica", 9)
    c.drawCentredString(width / 2, 40, "This is a system-generated payslip. For queries, contact HR.")

    c.save()
    print(f"✔ Payslip saved: {filename}")

# === Step 3: Loop through all employees and generate payslips ===
for _, row in df.iterrows():
    generate_payslip(row)

import yagmail

# === Step 4: Emailing Payslips ===

# Set your sender email and password/app-password
sender_email = 'example@gmail.com'
app_password = 'app_password'  # For Gmail, generate an App Password

# Initialize yagmail client
yag = yagmail.SMTP(user=sender_email, password=app_password)

# Loop through each employee and send email
for _, row in df.iterrows():
    emp_id = row['Employee ID']
    emp_name = row['Employee Name']
    recipient = row['Email']
    filename = f"{output_folder}/{emp_id}.pdf"

    subject = "Your Payslip for This Month"
    body = f"""
Dear {emp_name},

Attached is your payslip for this month.

Please review the details, and if you have any questions, feel free to contact HR.

Best regards,  
Tank Logistics Payroll Team
"""

    try:
        yag.send(
            to=recipient,
            subject=subject,
            contents=body,
            attachments=filename
        )
        print(f" Email sent to {emp_name} ({recipient})")
    except Exception as e:
        print(f" Failed to send email to {emp_name}: {e}")


