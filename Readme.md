Tank Logistics: Automated Payslip Generator & Email Sender
This Python script automates the process of:

Reading employee salary data from an Excel file

Generating professional-looking PDF payslips

Emailing the payslips to each employee securely

 Project Structure
bash
Copy
Edit
.
├── employees.xlsx              # Your input Excel file
├── payslips/                   # Output folder for generated PDF payslips
├── main.py                     # Main script to run the process
└── README.md                   # This documentation
 Features
Automatically calculates Net Salary

Beautiful payslip layout with company branding

Emails each employee their payslip using Gmail

Logs all processing and email activity

Fully customizable

 Requirements
Install the following Python libraries:

bash
Copy
Edit
pip install pandas reportlab yagmail openpyxl
 Input Excel Format
Your employees.xlsx should have the following columns in Sheet1:


Employee ID	Employee Name	Department	Email	Basic Salary	Allowance	Deductions
1001	John Doe	Finance	johndoe@email.com	1000	200	50
 How the Script Works
1. Load & Process Excel File
Reads the Excel file

Renames columns

Calculates Net Salary = Basic Salary + Allowance - Deductions

2. Generate PDF Payslips
For each employee, a personalized PDF is created under the payslips/ folder

Payslip includes:

Employee details

Salary breakdown

Company name (Tank Logistics)

Authorized signature by Tatenda Karindi

3. Send Payslips by Email
Uses yagmail to send emails

Attaches PDF payslip for each employee

Sends to their email from the Excel sheet

 How to Run the Script
Step 1: Prepare your Excel
Edit or create employees.xlsx using the format above.

Step 2: Enable Gmail App Passwords
Go to Google Account Security

Enable 2-Step Verification (if not already enabled)

Create an App Password for "Mail"

Copy that app password

 Never store passwords directly in the script. Use environment variables or .env files in production.

Step 3: Add Your Email to the Script
Edit these lines with your actual email and app password:

python
Copy
Edit
sender_email = 'your-email@gmail.com'
app_password = 'your-app-password'
Step 4: Run the Script
bash
Copy
Edit
python main.py
 Security Warning
This line stores your password directly in the code:

python
Copy
Edit
app_password = 'vxvs rnti xymw vern'
Instead, you should:

python
Copy
Edit
import os

sender_email = os.getenv("EMAIL_USER")
app_password = os.getenv("EMAIL_PASS")
Then use a .env file or set them in your terminal session.

 To-Do / Improvements
 Secure credentials with environment variables

 Zip all payslips before sending (optional)

 Add scheduling with schedule module

 Add unit tests for each function
