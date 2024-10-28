from flask import Flask, render_template, request, flash, redirect, url_for
import pandas as pd
from docxtpl import DocxTemplate
from docx2pdf import convert
import os
import datetime
import secrets

app = Flask(__name__)
app.secret_key = secrets.token_hex(16)
# paths for template and output folders
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
template = os.path.join(BASE_DIR, "admission.docx")
filled_forms_folder = os.path.join(BASE_DIR, "filled_forms")
uploads_folder = os.path.join(BASE_DIR, "uploads")


os.makedirs(filled_forms_folder, exist_ok=True)
os.makedirs(uploads_folder, exist_ok=True)


def format_date(date_value):
    if pd.notna(date_value) and isinstance(date_value, (pd.Timestamp, datetime.date)):
        return date_value.strftime("%Y-%m-%d")
    return date_value


def generate_form(data_row):
    # Load DOCX template
    doc = DocxTemplate(template)

    # Prepare the context dictionary with data from the row
    context = {
        "Name": data_row.get("Name", ""),
        "Fathers_Name": data_row.get("Father's Name", ""),
        "Nationality": data_row.get("Nationality", ""),
        "Gender": data_row.get("Gender", ""),
        "CNIC": data_row.get("CNIC", ""),
        "Present_Address": data_row.get("Present Address", ""),
        "Permanent_Address": data_row.get("Permanent Address", ""),
        "Last_Exam_Passed": data_row.get("Last Exam Passed", ""),
        "University_Board": data_row.get("University/Board", ""),
        "Passing_Year": data_row.get("Passing Year", ""),
        "Division_Class": data_row.get("Division/Class", ""),
        "Eligibility_Certificate": data_row.get("Eligibility Certificate", ""),
        "Cert_No": data_row.get("Cert No", ""),
        "Cert_Date": format_date(data_row.get("Cert Date", "")),
        "DOB": format_date(data_row.get("DOB", "")),
        "Matric_Cert_Status": data_row.get("Matric Cert Status", ""),
        "Batch": data_row.get("Batch", ""),
        "Department": data_row.get("Department", ""),
        "Roll_No": data_row.get("Roll No", ""),
        "Challan_No": data_row.get("Challan No", ""),
        "Challan_Date": format_date(data_row.get("Challan Date", "")),
    }

    # Render and save the DOCX form
    doc.render(context)
    docx_path = os.path.join(filled_forms_folder, f"Form_{data_row['index']}.docx")
    doc.save(docx_path)

    # Convert DOCX to PDF
    pdf_path = docx_path.replace(".docx", ".pdf")
    convert(docx_path, pdf_path)

    return pdf_path


@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":

        excel_file = request.files.get("excel_file")
        if not excel_file or not excel_file.filename.endswith((".xls", ".xlsx")):
            flash("Please upload a valid Excel file.", "danger")
            return redirect(url_for("index"))

        # Save Excel file to uploads directory
        excel_path = os.path.join(uploads_folder, excel_file.filename)
        excel_file.save(excel_path)

        # Load data and generate forms
        try:
            data = pd.read_excel(excel_path)
            for index, row in data.iterrows():
                row["index"] = index + 1
                generate_form(row)

            flash(
                "All forms have been successfully saved in the filled_forms folder.",
                "success",
            )

        except Exception as e:
            flash(f"Error processing the Excel file: {e}", "danger")

        # Redirect to avoid form resubmission
        return redirect(url_for("index"))

    return render_template("index.html")


if __name__ == "__main__":
    app.run(debug=True, host="0.0.0.0")
