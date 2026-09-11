"""
run_cloud_daily.py — Vietnam ESL Daily Automation
===================================================
Orchestrates the full pipeline:
  1. Scrape job listings (web + Facebook)
  2. Generate Excel, PDF cover letters, JSON dump
  3. Upload all files to a dated Google Drive folder
  4. Send HTML email notification to all recipients
"""

import os
import json
import smtplib
from datetime import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from zoneinfo import ZoneInfo

from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, HRFlowable
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_JUSTIFY

from scraper import scrape_all

# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------

PARENT_FOLDER_ID = os.getenv("PARENT_FOLDER_ID", "13H4RSyyrwza7oml5OYjlCki3QH8oM_25")
RECIPIENTS = [
    "klachhab@icloud.com",
    "wavesofmorocco@gmail.com",
    "tahamahdaoui07@outlook.com",
]
GMAIL_SENDER = os.getenv("GMAIL_SENDER")
GMAIL_APP_PASSWORD = os.getenv("GMAIL_APP_PASSWORD")
PORTFOLIO_LINK = "https://drive.google.com/drive/folders/1_NzhcAhTIkWYhS1dDikIed0TOhlGt2X4"


# ---------------------------------------------------------------------------
# Google Drive helpers
# ---------------------------------------------------------------------------

def get_drive_service():
    sa_json = os.getenv("GDRIVE_SERVICE_ACCOUNT_JSON")
    sa_info = json.loads(sa_json)
    creds = Credentials.from_service_account_info(
        sa_info,
        scopes=["https://www.googleapis.com/auth/drive"],
    )
    # cache_discovery=False prevents FileNotFoundError on ephemeral CI runners
    return build("drive", "v3", credentials=creds, cache_discovery=False)


def create_daily_drive_folder(service, folder_name: str, parent_id: str):
    meta = {
        "name": folder_name,
        "mimeType": "application/vnd.google-apps.folder",
        "parents": [parent_id],
    }
    folder = service.files().create(body=meta, fields="id, webViewLink").execute()
    return folder["id"], folder["webViewLink"]


def upload_to_drive(service, file_path: str, folder_id: str) -> str:
    file_name = os.path.basename(file_path)
    meta = {"name": file_name, "parents": [folder_id]}
    media = MediaFileUpload(file_path, resumable=True)
    uploaded = service.files().create(
        body=meta, media_body=media, fields="id, webViewLink"
    ).execute()
    return uploaded["webViewLink"]


# ---------------------------------------------------------------------------
# Document builders
# ---------------------------------------------------------------------------

def build_pdf_letter(output_path: str, name: str, qualifications: list,
                     salutation: str, target_role: str):
    doc = SimpleDocTemplate(
        output_path, pagesize=A4,
        rightMargin=45, leftMargin=45, topMargin=45, bottomMargin=45
    )
    styles = getSampleStyleSheet()
    story = []

    body_style = ParagraphStyle(
        "Body", parent=styles["Normal"],
        fontSize=10.5, leading=15,
        alignment=TA_JUSTIFY,
        textColor=colors.HexColor("#222222"),
    )
    title_style = ParagraphStyle(
        "Title",
        fontName="Helvetica-Bold",
        fontSize=18, leading=22,
        textColor=colors.HexColor("#0B2545"),
    )

    story.append(Paragraph(name, title_style))
    story.append(Spacer(1, 4))
    story.append(Paragraph(
        f"Application for: {target_role} | Certified C1 CEFR Teacher",
        styles["Normal"]
    ))
    story.append(Spacer(1, 8))
    story.append(HRFlowable(
        width="100%", thickness=1.5,
        color=colors.HexColor("#0B2545"), spaceAfter=14
    ))
    story.append(Paragraph(salutation, ParagraphStyle(
        "Salutation", fontName="Helvetica-Bold", fontSize=11, leading=15
    )))
    story.append(Spacer(1, 10))

    p1 = (
        f"I am writing to express my strong interest in English language teaching positions in Vietnam. "
        f"I hold accredited credentials including {qualifications[0]} and {qualifications[1]}. "
        f"My background enables me to deliver communicative, student-centered lessons tailored to "
        f"learners' developmental stages."
    )
    story.append(Paragraph(p1, body_style))
    story.append(Spacer(1, 10))

    p2 = (
        f"In my previous instructional assignments, including {qualifications[3]}, I developed engaging "
        f"classroom activities and structured evaluations. Fluent across English and multilingual "
        f"environments, I prioritize clear pronunciation modeling, positive reinforcement, and active "
        f"student participation."
    )
    story.append(Paragraph(p2, body_style))
    story.append(Spacer(1, 10))

    p3 = (
        f"My complete documentation, degrees, background checks, and demo videos are available for "
        f"review via my <a href='{PORTFOLIO_LINK}'><u>Google Drive Credentials Portfolio</u></a>. "
        f"I welcome the opportunity to discuss how my qualifications align with your academic goals."
    )
    story.append(Paragraph(p3, body_style))
    story.append(Spacer(1, 14))
    story.append(Paragraph("Sincerely,<br/>" + name, body_style))

    doc.build(story)


def build_excel(output_path: str, data: dict):
    wb = openpyxl.Workbook()
    ws_web = wb.active
    ws_web.title = "Web_Jobs"
    ws_fb = wb.create_sheet(title="Facebook_Jobs")

    headers = ["Job Title", "Institution / Company", "Location",
               "Salary Range", "Key Requirements", "Email", "URL"]

    for ws, key in [(ws_web, "jobs"), (ws_fb, "jobs_on_facebook")]:
        ws.append(headers)
        for col_num in range(1, len(headers) + 1):
            cell = ws.cell(row=1, column=col_num)
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
            cell.alignment = Alignment(horizontal="center", vertical="center")

        for item in data.get(key, []):
            reqs = item.get("key_requirements", [])
            reqs_str = " • ".join(reqs) if isinstance(reqs, list) else reqs
            ws.append([
                item.get("title", ""),
                item.get("company", ""),
                item.get("location", ""),
                item.get("salary_range", "Not Specified"),
                reqs_str,
                item.get("email", "N/A"),
                item.get("url", ""),
            ])

        for col in ws.columns:
            max_len = max(len(str(cell.value or "")) for cell in col)
            ws.column_dimensions[get_column_letter(col[0].column)].width = max(max_len + 3, 14)

    wb.save(output_path)


# ---------------------------------------------------------------------------
# Email notification
# ---------------------------------------------------------------------------

def send_notification(folder_link: str, file_links: dict, date_str: str):
    msg = MIMEMultipart("alternative")
    msg["Subject"] = f"Daily English Teaching Job Digest — Vietnam — {date_str}"
    msg["From"] = GMAIL_SENDER
    msg["To"] = ", ".join(RECIPIENTS)

    rows = "".join(
        f'<li><a href="{link}">{fname}</a></li>'
        for fname, link in file_links.items()
    )

    html = f"""
    <html><body>
    <h2 style="color:#0B2545;">Daily Vietnam ESL Dispatch: {date_str}</h2>
    <p>The daily search and document generation cycle completed successfully.</p>
    <p><b>Today's Google Drive Folder:</b>
       <a href="{folder_link}">{date_str} Folder</a></p>
    <h3>Generated Artifacts:</h3>
    <ul>{rows}</ul>
    <hr/>
    <p style="color:#888;font-size:12px;">
      Automated by GitHub Actions · Vietnam ESL Pipeline
    </p>
    </body></html>
    """
    msg.attach(MIMEText(html, "html"))

    with smtplib.SMTP("smtp.gmail.com", 587) as server:
        server.starttls()
        server.login(GMAIL_SENDER, GMAIL_APP_PASSWORD)
        server.sendmail(GMAIL_SENDER, RECIPIENTS, msg.as_string())

    print(f"[email] Notification sent to {RECIPIENTS}")


# ---------------------------------------------------------------------------
# Main orchestrator
# ---------------------------------------------------------------------------

def main():
    casablanca = ZoneInfo("Africa/Casablanca")
    today_str = datetime.now(casablanca).strftime("%d-%m-%Y")
    os.makedirs(today_str, exist_ok=True)
    print(f"[pipeline] Starting run for {today_str}")

    # 1. Scrape
    print("[pipeline] Step 1: Scraping job data...")
    pipeline_data = scrape_all()

    # 2. Generate files
    print("[pipeline] Step 2: Generating output files...")
    json_path   = os.path.join(today_str, "pipeline_data.json")
    excel_path  = os.path.join(today_str, "Vietnam_ESL.xlsx")
    pdf_khalid  = os.path.join(today_str, "Cover_Letter_Khalid_Lachhab.pdf")
    pdf_taha    = os.path.join(today_str, "Cover_Letter_Taha_Mahdaoui.pdf")
    script_path = os.path.join(today_str, "vietnam_esl_automation.py")

    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(pipeline_data, f, indent=2, ensure_ascii=False)

    build_excel(excel_path, pipeline_data)

    build_pdf_letter(
        pdf_khalid, "Khalid Lachhab",
        [
            "120-hour accredited face-to-face TESOL",
            "Master's degree from SupMIT",
            "C1 CEFR proficiency",
            "teaching placement in Antalya, Turkey",
        ],
        "Dear Hiring Committee,",
        "English Language Teacher (Primary / Secondary)",
    )

    build_pdf_letter(
        pdf_taha, "Taha Mahdaoui",
        [
            "120+ hour TESOL + ICT in ESL certificate",
            "Bachelor's degree in Mechanical Engineering",
            "C1 CEFR multilingual proficiency",
            "primary classroom instruction at Alwafae School",
        ],
        "Dear Academic Director,",
        "English Language Teacher (Primary & Bilingual Programs)",
    )

    # Copy this script into the dated folder for archival
    with open(__file__, "r", encoding="utf-8") as src, \
         open(script_path, "w", encoding="utf-8") as dst:
        dst.write(src.read())

    # 3. Upload to Google Drive
    print("[pipeline] Step 3: Uploading to Google Drive...")
    service = get_drive_service()
    folder_id, folder_link = create_daily_drive_folder(service, today_str, PARENT_FOLDER_ID)

    file_links = {}
    for local_file in [json_path, excel_path, pdf_khalid, pdf_taha, script_path]:
        fname = os.path.basename(local_file)
        link = upload_to_drive(service, local_file, folder_id)
        file_links[fname] = link
        print(f"  ✓ Uploaded {fname}")

    # 4. Send email
    print("[pipeline] Step 4: Sending notification email...")
    send_notification(folder_link, file_links, today_str)

    print(f"[pipeline] ✅ Done! Folder: {folder_link}")


if __name__ == "__main__":
    main()
