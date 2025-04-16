import os
import pandas as pd
from fpdf import FPDF
import yagmail

# --- Configuration ---
EMAIL_USER = os.getenv("couragemwadziwana@gmail.com")         # e.g. your Gmail address
EMAIL_PASSWORD = os.getenv("fzep dmgi jngc cmex") # e.g. Gmail app password
    
EXCEL_FILE = "employees.xlsx"
# print(df.columns.tolist())
OUTPUT_DIR = "payslips"

# --- Ensure Output Directory Exists ---
os.makedirs(OUTPUT_DIR, exist_ok=True)

# --- Load Employee Data ---
try:
    df = pd.read_excel(EXCEL_FILE)
    print(df.columns.tolist())
except Exception as e:
    print(f"Error reading Excel file: {e}")
    exit()

# --- Calculate Net Salary ---
df['Net Salary'] = df['Basic Salary'] + df['Allowances'] - df['Deductions']

# --- Initialize Email Client ---
try:
    yag = yagmail.SMTP("couragemwadziwana@gmail.com", "fzep dmgi jngc cmex")
except Exception as e:
    print(f"Error logging into email: {e}")
    exit()

# # --- Generate Payslips and Send Emails ---
for _, row in df.iterrows():
    emp_id = row['Employee ID']
    name = row['NAME']
    email = row['Email']
    basic = row['Basic Salary']
    allow = row['Allowances']
    deduct = row['Deductions']
    net = row['Net Salary']

    pdf_path = f"{OUTPUT_DIR}/{emp_id}.pdf"

#     # Generate PDF
    try:
        pdf = FPDF()
        pdf.add_page()
        pdf.set_font("Arial", size=12)

        pdf.cell(200, 10, txt="Payslip", ln=True, align='C')
        pdf.ln(10)
        pdf.cell(200, 10, txt=f"Employee Name: {name}", ln=True)
        pdf.cell(200, 10, txt=f"Employee ID: {emp_id}", ln=True)
        pdf.cell(200, 10, txt=f"Basic Salary: ${basic:.2f}", ln=True)
        pdf.cell(200, 10, txt=f"Allowances: ${allow:.2f}", ln=True)
        pdf.cell(200, 10, txt=f"Deductions: ${deduct:.2f}", ln=True)
        pdf.cell(200, 10, txt=f"Net Salary: ${net:.2f}", ln=True)
 
        pdf.output(pdf_path)
        print(f"[✔] PDF generated for {name}: {pdf_path}")
    except Exception as e:
        print(f"[✘] Failed to generate PDF for {name}: {e}")
        continue

#     # Send Email
    try:
        subject = "Your Payslip for This Month"
        body = f"Hello {name},\n\nPlease find attached your payslip for this month.\n\nRegards,\nHR Team"
        yag.send(to=email, subject=subject, contents=body, attachments=pdf_path)
        print(f"[📧] Email sent to {email}")
    except Exception as e:
        print(f"[✘] Failed to send email to {email}: {e}")

print("\nDone!")
