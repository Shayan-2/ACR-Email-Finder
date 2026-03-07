import os
import imaplib
from pathlib import Path

from dotenv import load_dotenv
from flask import Flask, abort, render_template, request, send_file

from acr_processor.email_client import download_acr_pdfs
from acr_processor.excel_writer import write_students_to_excel
from acr_processor.models import StudentRecord
from acr_processor.pdf_parser import extract_text_from_pdf, parse_student_fields

load_dotenv(override=True)

BASE_DIR = Path(__file__).parent
DOWNLOAD_DIR = BASE_DIR / "downloads"
DEFAULT_EXCEL_OUTPUT = BASE_DIR / "outputs" / "acr_students.xlsx"

app = Flask(__name__)


def process_acr_emails() -> dict:
    keywords = [
        keyword.strip()
        for keyword in os.getenv("ACR_KEYWORDS", "Academic Consideration Request,ACR").split(",")
        if keyword.strip()
    ]

    saved_pdfs = download_acr_pdfs(DOWNLOAD_DIR, keywords)

    students = []
    for file_data in saved_pdfs:
        pdf_path = Path(file_data["pdf_path"])
        text = extract_text_from_pdf(pdf_path)
        fields = parse_student_fields(text)

        student_name = fields.get("student_name", "").strip()
        student_email = fields.get("student_email", "").strip().lower()

        students.append(
            StudentRecord(
                student_name=student_name,
                student_number=fields.get("student_number", ""),
                course_code=fields.get("course_code", ""),
                request_date=fields.get("request_date", ""),
                email_subject=file_data["email_subject"],
                student_email=student_email,
                pdf_filename=file_data["pdf_filename"],
                pdf_path=file_data["pdf_path"],
            )
        )

    excel_output = Path(os.getenv("EXCEL_OUTPUT", str(DEFAULT_EXCEL_OUTPUT)))
    excel_path = write_students_to_excel(students, excel_output)
    excel_path_obj = Path(excel_path)

    return {
        "students": students,
        "excel_path": excel_path,
        "excel_filename": excel_path_obj.name,
        "download_count": len(saved_pdfs),
    }


@app.route("/", methods=["GET", "POST"])
def index():
    result = None
    error = None

    if request.method == "POST":
        try:
            result = process_acr_emails()
        except imaplib.IMAP4.error:
            imap_host = os.getenv("IMAP_HOST", "")
            if "gmail" in imap_host.lower():
                error = (
                    "IMAP login failed for Gmail. Use IMAP_HOST=imap.gmail.com, your full Gmail "
                    "address in IMAP_USERNAME, and a 16-character Google App Password in IMAP_PASSWORD "
                    "(normal account passwords usually fail)."
                )
            elif "office365" in imap_host.lower() or "outlook" in imap_host.lower():
                error = (
                    "IMAP login failed for Outlook. Verify IMAP_HOST=outlook.office365.com, use your full "
                    "email in IMAP_USERNAME, and use an app password if MFA is enabled. Some school/work "
                    "Microsoft accounts block basic IMAP login entirely."
                )
            else:
                error = (
                    "IMAP login failed. Verify IMAP_HOST, IMAP_USERNAME, IMAP_PASSWORD, and provider IMAP "
                    "policy (app passwords may be required)."
                )
        except Exception as exc:
            error = str(exc)

    return render_template("index.html", result=result, error=error)


@app.route("/download-excel")
def download_excel():
    excel_output = Path(os.getenv("EXCEL_OUTPUT", str(DEFAULT_EXCEL_OUTPUT)))
    if not excel_output.exists():
        abort(404, "Excel file not found. Run collection first.")

    return send_file(excel_output, as_attachment=True, download_name=excel_output.name)


if __name__ == "__main__":
    app.run(debug=True)
