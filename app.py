from flask import Flask, request, render_template
import openpyxl
import os

app = Flask(__name__)

# Define the Excel file path
EXCEL_FILE = "contact_form_data.xlsx"

# Create an Excel file with headers if it doesn't exist
if not os.path.exists(EXCEL_FILE):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["Full Name", "Email", "Message"])  # Column headers
    wb.save(EXCEL_FILE)

@app.route("/submit-form", methods=["POST"])
def submit_form():
    fullname = request.form.get("fullname")
    email = request.form.get("email")
    message = request.form.get("message")

    # Load the workbook and get the active sheet
    wb = openpyxl.load_workbook(EXCEL_FILE)
    ws = wb.active

    # Append new data
    ws.append([fullname, email, message])

    # Save the workbook
    wb.save(EXCEL_FILE)

    return "Form submitted successfully!"
print("Excel file saved at:", os.path.abspath(EXCEL_FILE))

if __name__ == "__main__":
    app.run(debug=True)
