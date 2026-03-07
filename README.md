# ACR-Email-Finder

This project scans an email inbox for Academic Consideration Request (ACR) emails, downloads attached PDF files, extracts student details from each PDF, and creates an Excel file for easy review.

## What this app does

- Reads emails from your IMAP mailbox
- Filters emails by ACR-related keywords
- Downloads only PDF attachments from matching emails
- Parses basic student fields from the PDFs
- Writes all records into an Excel file
- Shows run results in a web frontend

## Project structure

```text
acr_project/
  app.py
  .env.example
  requirements.txt
  downloads/
  outputs/
  templates/
    index.html
  static/
    styles.css
  acr_processor/
    email_client.py
    pdf_parser.py
    excel_writer.py
    models.py
```

## 1. Create and activate a virtual environment (PowerShell)

```powershell
cd "c:\Users\shaya\OneDrive\Documents\Toronto Metropolitan University\Winter 2026\CPS 506\BigOF Project\acr_project"
python -m venv .venv
.\.venv\Scripts\Activate.ps1
```

## 2. Install dependencies

```powershell
pip install -r requirements.txt
```

## 3. Configure environment

1. Duplicate `.env.example` as `.env`.
2. Fill in your email account settings:
   - `IMAP_HOST`
   - `IMAP_PORT`
   - `IMAP_USERNAME`
   - `IMAP_PASSWORD`

### Important mail provider notes

- Gmail: use an **App Password** (not your normal password).
- Outlook: use IMAP-enabled account credentials.
- School email: check your university IMAP settings.

## 4. Run the web app

```powershell
python app.py
```

Open the URL shown in terminal (usually `http://127.0.0.1:5000`) and click **Run Collection**.

## Output

- PDFs are saved in `downloads/`
- Excel file is saved in `outputs/acr_students.xlsx` (or your custom `EXCEL_OUTPUT` path)

## Speed tuning (important)

If your mailbox is large, scanning can take time. Use these optional `.env` values:

- `IMAP_DAYS_BACK=14` to scan only recent emails (set `0` to scan all)
- `IMAP_MAX_EMAILS=200` to process only the newest N emails per run (set `0` for no cap)
- `IMAP_UNSEEN_ONLY=true` to scan only unread emails

Recommended quick setup for student inboxes:

- `IMAP_DAYS_BACK=14`
- `IMAP_UNSEEN_ONLY=true`

## Notes on PDF parsing

PDF forms can have different layouts. If some fields appear blank in Excel, update regex patterns in `acr_processor/pdf_parser.py`.

## Security tips

- Never commit `.env` to Git.
- Keep app passwords private.
- If possible, use a test mailbox while developing.
