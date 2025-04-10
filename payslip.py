import pandas as pd
from fpdf import FPDF
import yagmail
import os
from dotenv import load_dotenv

# Load environment variables from .env file
load_dotenv()
SENDER_EMAIL = os.getenv("SENDER_EMAIL")
SENDER_PASSWORD = os.getenv("SENDER_PASSWORD")
SMTP_SERVER = os.getenv("SMTP_SERVER", "smtp.gmail.com")
SMTP_PORT = int(os.getenv("SMTP_PORT", 465))

# Read employee data from Excel
def read_employee_data(filename):
    df = pd.read_excel(filename)
    df.columns = df.columns.str.strip()  # Remove leading/trailing whitespace
    df['Basic Salary'] = pd.to_numeric(df['Basic Salary'], errors='coerce')
    df['Allowances'] = pd.to_numeric(df['Allowances'], errors='coerce')
    df['Deductions'] = pd.to_numeric(df['Deductions'], errors='coerce')
    return df

# Calculate Net Salary
def calculate_net_salary(basic_salary, allowances, deductions):
    return basic_salary + allowances - deductions

# Generate Payslip PDF
def generate_payslip(employee):
    pdf = FPDF()
    pdf.add_page()

    # Header
    pdf.set_font("Arial", 'B', 16)
    pdf.cell(200, 10, f"Payslip for {employee['Name']}", ln=True, align='C')
    pdf.ln(10)

    # Employee Info
    emp_id = employee.get('Employee ID', 'N/A')
    name = employee.get('Name', 'N/A')
    email = employee.get('Email', 'N/A')

    pdf.set_font("Arial", '', 12)
    pdf.cell(100, 10, f"Employee ID: {emp_id}", ln=True)
    pdf.cell(100, 10, f"Name: {name}", ln=True)
    pdf.cell(100, 10, f"Email: {email}", ln=True)
    pdf.ln(10)

    # Salary Info
    pdf.cell(100, 10, f"Basic Salary: {employee['Basic Salary']}", ln=True)
    pdf.cell(100, 10, f"Allowances: {employee['Allowances']}", ln=True)
    pdf.cell(100, 10, f"Deductions: {employee['Deductions']}", ln=True)
    pdf.cell(100, 10, f"Net Salary: {employee['Net Salary']}", ln=True)

    # Save PDF
    payslip_dir = "payslips"
    os.makedirs(payslip_dir, exist_ok=True)
    file_path = f"{payslip_dir}/{name.replace(' ', '_')}.pdf"
    pdf.output(file_path)
    return file_path

# Send Payslip Email
def send_email(employee, payslip_file):
    subject = "Your Payslip for This Month"
    body = f"""
    Dear {employee['Name']},

    Please find attached your payslip for this month.

    Best Regards,
    Your Company
    """
    try:
        yag = yagmail.SMTP(user=SENDER_EMAIL, password=SENDER_PASSWORD)
        yag.send(
            to=employee['Email'],
            subject=subject,
            contents=body,
            attachments=payslip_file
        )
        print(f"Payslip for {employee['Name']} sent successfully.")
    except Exception as e:
        print(f"Error sending email to {employee['Email']}: {e}")

# Main Execution
def main():
    filename = "employees.xlsx"
    employees = read_employee_data(filename)

    for _, employee in employees.iterrows():
        # Check for missing/invalid numeric data
        if pd.isna(employee['Basic Salary']) or pd.isna(employee['Allowances']) or pd.isna(employee['Deductions']):
            print(f"Warning: Missing or invalid data for Employee ID {employee.get('Employee ID', 'Unknown')}")
            continue

        # Calculate Net Salary
        employee['Net Salary'] = calculate_net_salary(
            employee['Basic Salary'],
            employee['Allowances'],
            employee['Deductions']
        )

        # Generate and Send Payslip
        payslip_path = generate_payslip(employee)
        send_email(employee, payslip_path)

if __name__ == "__main__":
    main()
