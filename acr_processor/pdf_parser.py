import re
from datetime import datetime
from pathlib import Path
from typing import Dict

from pypdf import PdfReader


def extract_text_from_pdf(pdf_path: Path) -> str:
    reader = PdfReader(str(pdf_path))
    pages_text = []
    for page in reader.pages:
        pages_text.append(page.extract_text() or "")
    return "\n".join(pages_text)


def _find_first(pattern: str, text: str, group: int = 1) -> str:
    match = re.search(pattern, text, flags=re.IGNORECASE)
    if not match:
        return ""
    return match.group(group).strip()


def _lines(text: str) -> list[str]:
    return [re.sub(r"\s+", " ", line).strip() for line in text.splitlines()]


def _extract_labeled_value(text: str, labels: list[str]) -> str:
    entries = _lines(text)
    if not entries:
        return ""

    normalized_labels = [label.lower() for label in labels]
    for index, line in enumerate(entries):
        if not line:
            continue
        lower_line = line.lower()
        for label in normalized_labels:
            if label not in lower_line:
                continue

            # Case 1: label and value are on the same line.
            delimiter_match = re.search(r"[:\-]\s*(.+)$", line)
            if delimiter_match:
                return delimiter_match.group(1).strip()

            # Case 2: value appears right after the label on the same line without a delimiter.
            inline = re.sub(label, "", lower_line, count=1).strip()
            if inline:
                original_inline = line[len(line) - len(inline):].strip()
                if original_inline:
                    return original_inline

            # Case 3: value is on the next non-empty line.
            for next_index in range(index + 1, min(index + 4, len(entries))):
                candidate = entries[next_index].strip()
                if not candidate:
                    continue
                if re.search(r"\b(name|email|student|course|date|reason|instructor|number|id)\b", candidate, re.IGNORECASE):
                    break
                return candidate

    return ""


def _collect_name_lines(entries: list[str], start_index: int) -> str:
    parts: list[str] = []
    stop_re = re.compile(
        r"\b(student|id|number|course|class|date|email|reason|instructor|signature|term|semester)\b",
        re.IGNORECASE,
    )

    for i in range(start_index, min(start_index + 6, len(entries))):
        candidate = entries[i].strip(" :-\t\r\n")
        if not candidate:
            continue
        if stop_re.search(candidate):
            break
        if re.search(r"@", candidate):
            break
        if not re.search(r"[A-Za-z]", candidate):
            continue
        parts.append(candidate)

    return _clean_name(" ".join(parts))


def _clean_name(name: str) -> str:
    cleaned = re.sub(r"\s+", " ", (name or "")).strip(" :-\t\r\n")
    # Avoid accidental captures of emails or long lines of text.
    if "@" in cleaned or len(cleaned) > 80:
        return ""
    if not re.search(r"[A-Za-z]", cleaned):
        return ""
    return cleaned


def _extract_student_name(text: str) -> str:
    entries = _lines(text)

    # Pattern: line contains only "Name:" then first/last names on next lines.
    for idx, line in enumerate(entries):
        if re.fullmatch(r"(?:student\s*)?name\s*:?", line, flags=re.IGNORECASE):
            from_next = _collect_name_lines(entries, idx + 1)
            if from_next:
                return from_next

    # Pattern: name appears at top of the document before student-id labels.
    if entries:
        top_name = _collect_name_lines(entries, 0)
        if top_name:
            return top_name

    # Prefer clear labels first.
    labeled = _extract_labeled_value(
        text,
        ["Student Name", "Name of Student", "Applicant Name", "Full Name", "Name"],
    )
    cleaned_labeled = _clean_name(labeled)
    if cleaned_labeled:
        return cleaned_labeled

    # Common pattern on forms: First Name / Last Name.
    first_name = _clean_name(_extract_labeled_value(text, ["First Name", "Given Name"]))
    last_name = _clean_name(_extract_labeled_value(text, ["Last Name", "Surname", "Family Name"]))
    if first_name and last_name:
        return f"{first_name} {last_name}".strip()

    name_patterns = [
        r"Student\s*Name\s*[:\-]\s*([^\n\r]+)",
        r"Name\s*of\s*Student\s*[:\-]\s*([^\n\r]+)",
        r"Applicant\s*Name\s*[:\-]\s*([^\n\r]+)",
    ]
    for pattern in name_patterns:
        value = _find_first(pattern, text)
        cleaned = _clean_name(value)
        if cleaned:
            return cleaned
    return ""


def _extract_student_email(text: str) -> str:
    labeled = _extract_labeled_value(
        text,
        ["Student Email", "TMU Email", "Ryerson Email", "Email Address", "Email"],
    )
    if labeled:
        match = re.search(r"[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}", labeled)
        if match:
            return match.group(0).strip().lower()

    # Fallback: first email-like token in the document.
    match = re.search(r"[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}", text)
    if match:
        return match.group(0).strip().lower()
    return ""


def _normalize_date(raw_date: str) -> str:
    if not raw_date:
        return ""

    value = raw_date.strip()
    value = re.sub(r"(\d{1,2})(st|nd|rd|th)", r"\1", value, flags=re.IGNORECASE)
    value = re.sub(r"\s+", " ", value)

    formats = [
        "%d %B %Y",
        "%d %b %Y",
        "%B %d %Y",
        "%b %d %Y",
        "%B %d, %Y",
        "%b %d, %Y",
        "%d/%m/%Y",
        "%m/%d/%Y",
        "%Y-%m-%d",
    ]

    for fmt in formats:
        try:
            parsed = datetime.strptime(value, fmt)
            return parsed.strftime("%d.%b.%Y")
        except ValueError:
            continue

    # Last attempt: pick the first recognized date-like token sequence in the text.
    candidates = [
        r"\b\d{1,2}(?:st|nd|rd|th)?\s+[A-Za-z]+\s+\d{4}\b",
        r"\b[A-Za-z]+\s+\d{1,2}(?:st|nd|rd|th)?(?:,)?\s+\d{4}\b",
        r"\b\d{1,2}/\d{1,2}/\d{4}\b",
        r"\b\d{4}-\d{1,2}-\d{1,2}\b",
    ]
    for pattern in candidates:
        match = re.search(pattern, raw_date, flags=re.IGNORECASE)
        if not match:
            continue
        token = re.sub(r"(\d{1,2})(st|nd|rd|th)", r"\1", match.group(0), flags=re.IGNORECASE)
        token = token.replace("  ", " ").strip()
        for fmt in formats:
            try:
                parsed = datetime.strptime(token, fmt)
                return parsed.strftime("%d.%b.%Y")
            except ValueError:
                continue

    return ""


def _extract_request_date(text: str) -> str:
    labeled_date = _find_first(r"(?:Request\s*Date|Date)\s*[:\-]\s*([^\n\r]+)", text)
    normalized = _normalize_date(labeled_date)
    if normalized:
        return normalized

    # Fallback to first date-like occurrence in the full text.
    fallback_patterns = [
        r"\b\d{1,2}(?:st|nd|rd|th)?\s+[A-Za-z]+\s+\d{4}\b",
        r"\b[A-Za-z]+\s+\d{1,2}(?:st|nd|rd|th)?(?:,)?\s+\d{4}\b",
        r"\b\d{1,2}/\d{1,2}/\d{4}\b",
        r"\b\d{4}-\d{1,2}-\d{1,2}\b",
    ]
    for pattern in fallback_patterns:
        match = re.search(pattern, text, flags=re.IGNORECASE)
        if not match:
            continue
        normalized = _normalize_date(match.group(0))
        if normalized:
            return normalized

    return ""


def parse_student_fields(text: str) -> Dict[str, str]:
    # These patterns are intentionally flexible because ACR forms can vary in formatting.
    return {
        "student_name": _extract_student_name(text),
        "student_email": _extract_student_email(text),
        "student_number": _find_first(r"Student\s*(?:Number|ID)\s*[:\-]\s*([A-Za-z0-9-]+)", text),
        "course_code": _find_first(r"(?:Course\s*(?:Code)?|Class)\s*[:\-]\s*([A-Za-z]{2,}\s*\d{3,})", text),
        "request_date": _extract_request_date(text),
    }
