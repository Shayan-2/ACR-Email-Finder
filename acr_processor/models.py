from dataclasses import dataclass


@dataclass
class StudentRecord:
    student_name: str
    student_number: str
    course_code: str
    request_date: str
    email_subject: str
    student_email: str
    pdf_filename: str
    pdf_path: str
