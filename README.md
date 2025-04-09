# PAYSLIP GENERATOR PYTHON PROJECT
## OVERVIEW 

This Python script automates the process of generating payslips in PDF format and emailing them to employees using using the data from the Excel file within the folder.

## FEATURES

- Processes an Excel file containing employee information.
- Uses pandas to read the Excel file.
- Calculate the net salary using the formula indicated.
- Generate individual payslips as PDFs.
- Each PDF is to emailed to the corresponding employee using yagmail.
- Configure email settings using JSON.

## REQUIREMENTS
- Libraries:
  -pandas
  -reportlab
  -yagmail
  -openpyxl

INSTALL REQUIREMENTS USING
-pip install pandas reportlab yagmail openpyxl 