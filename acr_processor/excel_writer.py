from pathlib import Path
from typing import List

from openpyxl import Workbook

from acr_processor.models import StudentRecord


def write_students_to_excel(students: List[StudentRecord], output_path: Path) -> str:
    output_path.parent.mkdir(parents=True, exist_ok=True)

    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "ACR Students"

    headers = [
        "Student Name",
        "Student Number",
        "Course Code",
        "Request Date",
        "Email Subject",
        "Student Email",
        "PDF Filename",
        "PDF Path",
    ]
    sheet.append(headers)

    for student in students:
        sheet.append(
            [
                student.student_name,
                student.student_number,
                student.course_code,
                student.request_date,
                student.email_subject,
                student.student_email,
                student.pdf_filename,
                student.pdf_path,
            ]
        )

    workbook.save(output_path)
    return str(output_path)
