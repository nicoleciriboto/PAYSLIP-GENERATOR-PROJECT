#INSTALL AND IMPORT LIBRARIES
import os
import json
import pandas as pd 
import yagmail
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib import colors
from reportlab.lib.units import inch
from reportlab.platypus import Table, TableStyle

#LOAD EMAIL CREDENTIALS
with open('config.json') as f:
    config = json.load(f)

SENDER_EMAIL= config['sender email']
APP_PASSWORD= config['app password']

#READ THE EXCEL FILE
df = pd.read_excel('employees.xlsx')

#ENSURE PAYSLIPS DIRECTORY EXISTS
os.makedirs('payslips', exist_ok=True)

#FUNCTION TO GENERATE PDF PAYSLIP
def generate_payslip(employee):
    employee_id = employee['Employee ID']
    employee_name = employee['Name']
    basic_salary = employee['Basic Salary']
    allowance = employee['Allowances']
    deductions = employee['Deductions']
    net_salary = basic_salary + allowance - deductions

# Create a PDF file for the payslip
    pdf_filename = f'payslips/{employee_id}.pdf'
    c = canvas.Canvas(pdf_filename, pagesize=A4)
    width, height = A4

#HEADER SECTION
#LOGO AND COMPANY NAME
    logo_path = "BeautyLogo.jpeg"
    if os.path.exists(logo_path):
        c.drawImage(logo_path, 50, height - 80, width=80, height=50)

    c.setFont("Times-Bold", 20)
    c.drawCentredString(width / 2, height - 60, "The Beauty Clinic")

    c.setFont("Times-Bold", 10)
    c.drawCentredString(width / 2, height - 90, "Margarett Powell House  21 Cork Road  Avondale Outlets")
    c.line(30, height-110, width-40, height-110)

#TABLE
    data = [
    ["Description", "Amount (USD)"],
    ["Basic Salary", f"${basic_salary:.2f}"],
    ["Allowances", f"${allowance:.2f}"],
    ["Deductions", f"${deductions:.2f}"],
    ["Net Salary", f"${net_salary}"]
]
    table= Table(data, colWidths=[200, 200])
    table.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1, 0), colors.HexColor("#003366")),
        ('TEXTCOLOR', (0,0), (-1, 0), colors.white),
        ('ALIGN', (1,1), (-1, -1), 'RIGHT'),
        ('FONTNAME', (0,0), (-1,0), 'Times-Bold'),
        ('BOTTOMPADDING', (0,0), (-1,0), 12),
        ('GRID', (0,0), (-1, -1), 0.5, colors.grey),
    ]))
    table.wrapOn(c, width, height)
    table.drawOn(c, 50, height-300)


    c.setFont("Times-Roman", 12)
    c.setFillColor(colors.black)
    c.drawString(50, 450, "Thank you for your hard work! It is greatly appreciated.")
    c.save()
    
    return pdf_filename

#INITIALIZE EMAIL CLIENT
yag= yagmail.SMTP(SENDER_EMAIL, APP_PASSWORD)

#GENERATE PAYSLIPS AND SEND EMAILS
for _, row in df.iterrows():
    pdf_path = generate_payslip(row)
    recipient_email = row['Email']
    subject = "Your Payslip for this month"
    body = f"""Dear {row['Name']},
      Please find your attached payslip for this month. 
      If you have any questions, feel free to reach out."""
    
    #SEND EMAIL WITH ATTACHMENT
    try:
        yag.send(to=recipient_email, subject=subject, contents=body, attachments=pdf_path)
        print(f"Payslip sent to {row['Name']} at {recipient_email}")
    except Exception as e:
         print(f"Failed to send payslip to {row['Name']} at {recipient_email}. Error: {e}")









