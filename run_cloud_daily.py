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

from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request
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
    """Build Drive service using OAuth2 user credentials (refresh token).
    Uploads run as the real user, so Drive quota is never an issue.
    """
    creds = Credentials(
        token=None,
        refresh_token=os.environ["GDRIVE_REFRESH_TOKEN"],
        client_id=os.environ["GDRIVE_CLIENT_ID"],
        client_secret=os.environ["GDRIVE_CLIENT_SECRET"],
        token_uri="https://oauth2.googleapis.com/token",
    )
    # Refresh the access token immediately
    creds.refresh(Request())
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


HISTORY_FILE = "history_jobs.json"

def load_job_history() -> dict:
    if os.path.exists(HISTORY_FILE):
        try:
            with open(HISTORY_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception as e:
            print(f"[history] Warning reading {HISTORY_FILE}: {e}")
    return {}


def save_job_history(history: dict):
    try:
        with open(HISTORY_FILE, "w", encoding="utf-8") as f:
            json.dump(history, f, indent=2, ensure_ascii=False)
        print(f"[history] Saved {len(history)} total jobs in {HISTORY_FILE}")
    except Exception as e:
        print(f"[history] Error saving {HISTORY_FILE}: {e}")


def build_excel(output_path: str, data: dict, new_jobs: list):
    wb = openpyxl.Workbook()
    ws_new = wb.active
    ws_new.title = "New_Offers_Today"
    ws_web = wb.create_sheet(title="All_Active_Web_Jobs")
    ws_fb = wb.create_sheet(title="Facebook_Jobs")

    # Sheet 1: New offers
    new_headers = ["Job Title", "Source", "Institution / Company", "Location",
                   "Salary Range", "Key Requirements", "Email", "URL"]
    ws_new.append(new_headers)
    for col_num in range(1, len(new_headers) + 1):
        cell = ws_new.cell(row=1, column=col_num)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
        cell.alignment = Alignment(horizontal="center", vertical="center")

    for item in new_jobs:
        reqs = item.get("key_requirements", [])
        reqs_str = " • ".join(reqs) if isinstance(reqs, list) else str(reqs)
        ws_new.append([
            item.get("title", ""),
            item.get("source", "Web"),
            item.get("company", ""),
            item.get("location", ""),
            item.get("salary_range", "Not Specified"),
            reqs_str,
            item.get("email", "N/A"),
            item.get("url", ""),
        ])

    # Sheet 2 & 3: Web & Facebook jobs
    general_headers = ["Job Title", "Institution / Company", "Location",
                       "Salary Range", "Key Requirements", "Email", "URL"]

    for ws, key in [(ws_web, "jobs"), (ws_fb, "jobs_on_facebook")]:
        ws.append(general_headers)
        for col_num in range(1, len(general_headers) + 1):
            cell = ws.cell(row=1, column=col_num)
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill(start_color="203764", end_color="203764", fill_type="solid")
            cell.alignment = Alignment(horizontal="center", vertical="center")

        for item in data.get(key, []):
            reqs = item.get("key_requirements", [])
            reqs_str = " • ".join(reqs) if isinstance(reqs, list) else str(reqs)
            ws.append([
                item.get("title", ""),
                item.get("company", ""),
                item.get("location", ""),
                item.get("salary_range", "Not Specified"),
                reqs_str,
                item.get("email", "N/A"),
                item.get("url", ""),
            ])

    for ws in [ws_new, ws_web, ws_fb]:
        for col in ws.columns:
            max_len = max(len(str(cell.value or "")) for cell in col)
            ws.column_dimensions[get_column_letter(col[0].column)].width = min(max(max_len + 3, 14), 50)

    wb.save(output_path)


# ---------------------------------------------------------------------------
# Email notification
# ---------------------------------------------------------------------------

def send_no_offers_notification(date_str: str, scanned_stats: dict):
    """
    Sends a clear notification when no new ESL offers were found today.
    """
    if not GMAIL_SENDER or not GMAIL_APP_PASSWORD:
        print("[email] Skipping email (GMAIL_SENDER or GMAIL_APP_PASSWORD not set)")
        return

    msg = MIMEMultipart("alternative")
    msg["Subject"] = f"Daily Vietnam ESL Digest — {date_str} — No New Offers Found"
    msg["From"] = GMAIL_SENDER
    msg["To"] = ", ".join(RECIPIENTS)

    web_total = scanned_stats.get("web_total", 0)
    fb_total = scanned_stats.get("fb_total", 0)

    html = f"""
    <html>
    <body style="font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif; color: #1e293b; line-height: 1.6; background-color: #f8fafc; padding: 20px;">
      <div style="max-width: 600px; margin: auto; background: #ffffff; border: 1px solid #e2e8f0; border-radius: 10px; overflow: hidden; box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05);">
        <div style="background-color: #0B2545; color: #ffffff; padding: 24px; text-align: center;">
          <h2 style="margin: 0; font-size: 20px;">Daily Vietnam ESL Job Dispatch</h2>
          <p style="margin: 6px 0 0; opacity: 0.85; font-size: 14px;">Date: {date_str}</p>
        </div>

        <div style="padding: 24px;">
          <div style="background-color: #f1f5f9; border-left: 4px solid #64748b; padding: 16px; border-radius: 4px; margin-bottom: 20px;">
            <h3 style="margin: 0 0 6px 0; color: #0f172a; font-size: 16px;">ℹ️ No New Offers Found Today</h3>
            <p style="margin: 0; font-size: 14px; color: #475569;">
              Our automated daily search completed across the web and Facebook groups, but <b>no brand-new English teaching positions were posted in the last 24 hours</b>.
            </p>
          </div>

          <h4 style="margin: 16px 0 8px 0; color: #0B2545; font-size: 14px; text-transform: uppercase; letter-spacing: 0.5px;">Search Summary:</h4>
          <table style="width: 100%; border-collapse: collapse; font-size: 13px; margin-bottom: 20px;">
            <tr style="border-bottom: 1px solid #e2e8f0;">
              <td style="padding: 8px 0; color: #64748b;">Locations Monitored:</td>
              <td style="padding: 8px 0; font-weight: 600; color: #1e293b;">Hà Nội, TP. Hồ Chí Minh, Hải Phòng, Đà Nẵng, Cần Thơ, Khánh Hòa, etc.</td>
            </tr>
            <tr style="border-bottom: 1px solid #e2e8f0;">
              <td style="padding: 8px 0; color: #64748b;">Vietnam Teaching Jobs:</td>
              <td style="padding: 8px 0; color: #1e293b;">Checked ({web_total} active listings scanned, none new today)</td>
            </tr>
            <tr>
              <td style="padding: 8px 0; color: #64748b;">Facebook ESL Groups:</td>
              <td style="padding: 8px 0; color: #1e293b;">Checked ({fb_total} community posts scanned, none new today)</td>
            </tr>
          </table>

          <p style="font-size: 13px; color: #64748b; margin: 0;">
            The automation is active and will run again tomorrow night at <b>23:30 (Casablanca time)</b>. You will be alerted immediately when fresh opportunities appear.
          </p>
        </div>

        <div style="background-color: #f8fafc; padding: 14px 24px; font-size: 11px; color: #94a3b8; text-align: center; border-top: 1px solid #e2e8f0;">
          Automated by GitHub Actions · Vietnam ESL Job Pipeline
        </div>
      </div>
    </body>
    </html>
    """
    msg.attach(MIMEText(html, "html"))

    with smtplib.SMTP("smtp.gmail.com", 587) as server:
        server.starttls()
        server.login(GMAIL_SENDER, GMAIL_APP_PASSWORD)
        server.sendmail(GMAIL_SENDER, RECIPIENTS, msg.as_string())

    print(f"[email] 'No new offers' notification sent to {RECIPIENTS}")


def send_notification(folder_link: str, file_links: dict, date_str: str, new_jobs: list):
    """
    Sends a digest email showcasing newly found offers with Drive links.
    """
    if not GMAIL_SENDER or not GMAIL_APP_PASSWORD:
        print("[email] Skipping email (GMAIL_SENDER or GMAIL_APP_PASSWORD not set)")
        return

    msg = MIMEMultipart("alternative")
    msg["Subject"] = f"Daily Vietnam ESL Digest — {date_str} — {len(new_jobs)} New Offers Found"
    msg["From"] = GMAIL_SENDER
    msg["To"] = ", ".join(RECIPIENTS)

    # Distinct locations found
    locations = sorted(list(set(j.get("location", "Vietnam") for j in new_jobs)))
    loc_summary = ", ".join(locations[:6])
    if len(locations) > 6:
        loc_summary += f" +{len(locations)-6} more"

    # Artifact links
    artifact_items = "".join(
        f'<li style="margin-bottom: 4px;"><a href="{link}" style="color:#1F4E79;text-decoration:none;font-weight:600;">{fname}</a></li>'
        for fname, link in file_links.items()
    )

    # Top offers table (up to 8)
    table_rows = ""
    for j in new_jobs[:8]:
        url = j.get("url", "#")
        title = j.get("title", "ESL Teacher")
        company = j.get("company", "Educational Institution")
        loc = j.get("location", "Vietnam")
        sal = j.get("salary_range", "Competitive")
        source = j.get("source", "Web")
        table_rows += f"""
        <tr style="border-bottom: 1px solid #f1f5f9;">
          <td style="padding: 10px 8px; font-size: 13px;">
            <a href="{url}" style="color: #0B2545; font-weight: 600; text-decoration: none;">{title}</a>
            <div style="font-size: 11px; color: #64748b;">{company} • <span style="background:#e0f2fe;color:#0369a1;padding:1px 5px;border-radius:3px;">{source}</span></div>
          </td>
          <td style="padding: 10px 8px; font-size: 12px; color: #334155;">{loc}</td>
          <td style="padding: 10px 8px; font-size: 12px; color: #166534; font-weight: 600;">{sal}</td>
        </tr>
        """

    html = f"""
    <html>
    <body style="font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif; color: #1e293b; line-height: 1.6; background-color: #f8fafc; padding: 20px;">
      <div style="max-width: 640px; margin: auto; background: #ffffff; border: 1px solid #e2e8f0; border-radius: 10px; overflow: hidden; box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05);">
        <div style="background-color: #0B2545; color: #ffffff; padding: 24px;">
          <h2 style="margin: 0; font-size: 20px;">Daily Vietnam ESL Job Dispatch</h2>
          <p style="margin: 6px 0 0; opacity: 0.85; font-size: 14px;">Date: {date_str} • <b>{len(new_jobs)} New Offers Discovered</b></p>
        </div>

        <div style="padding: 24px;">
          <div style="background-color: #eff6ff; border-left: 4px solid #2563eb; padding: 14px 16px; border-radius: 4px; margin-bottom: 20px;">
            <div style="font-weight: 600; color: #1e40af; font-size: 14px;">🎯 Locations with new vacancies today:</div>
            <div style="color: #1e3a8a; font-size: 13px; margin-top: 2px;">{loc_summary}</div>
          </div>

          <div style="margin-bottom: 20px; text-align: center;">
            <a href="{folder_link}" style="display: inline-block; background-color: #1F4E79; color: #ffffff; font-weight: 600; font-size: 14px; padding: 10px 22px; text-decoration: none; border-radius: 6px;">
              📁 Open Today's Google Drive Folder
            </a>
          </div>

          <h4 style="margin: 20px 0 8px 0; color: #0B2545; font-size: 14px; text-transform: uppercase; letter-spacing: 0.5px;">Featured New Offers:</h4>
          <table style="width: 100%; border-collapse: collapse; margin-bottom: 20px;">
            <thead>
              <tr style="background-color: #f8fafc; text-align: left; font-size: 11px; color: #64748b; text-transform: uppercase;">
                <th style="padding: 8px;">Job & Institution</th>
                <th style="padding: 8px;">Location</th>
                <th style="padding: 8px;">Salary</th>
              </tr>
            </thead>
            <tbody>
              {table_rows}
            </tbody>
          </table>

          <h4 style="margin: 20px 0 8px 0; color: #0B2545; font-size: 14px; text-transform: uppercase; letter-spacing: 0.5px;">Generated Files in Drive:</h4>
          <ul style="font-size: 13px; color: #475569; padding-left: 20px; margin-top: 4px;">
            {artifact_items}
          </ul>
        </div>

        <div style="background-color: #f8fafc; padding: 14px 24px; font-size: 11px; color: #94a3b8; text-align: center; border-top: 1px solid #e2e8f0;">
          Automated by GitHub Actions · Vietnam ESL Job Pipeline
        </div>
      </div>
    </body>
    </html>
    """
    msg.attach(MIMEText(html, "html"))

    with smtplib.SMTP("smtp.gmail.com", 587) as server:
        server.starttls()
        server.login(GMAIL_SENDER, GMAIL_APP_PASSWORD)
        server.sendmail(GMAIL_SENDER, RECIPIENTS, msg.as_string())

    print(f"[email] New offers notification sent to {RECIPIENTS}")


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
    web_jobs = pipeline_data.get("jobs", [])
    fb_jobs = pipeline_data.get("jobs_on_facebook", [])

    # 2. History & Deduplication check
    print("[pipeline] Step 2: Checking for new offers...")
    history = load_job_history()
    is_initial_history = len(history) == 0

    new_web_jobs = [j for j in web_jobs if j.get("url") not in history]
    new_fb_jobs = [j for j in fb_jobs if j.get("url") not in history]
    new_jobs = new_web_jobs + new_fb_jobs

    total_new = len(new_jobs)
    print(f"[pipeline] Discovered {total_new} new offers ({len(new_web_jobs)} web, {len(new_fb_jobs)} Facebook)")

    # If NO NEW OFFERS found
    if total_new == 0:
        print("[pipeline] ℹ️ No new offers found today. Sending status notification email...")
        scanned_stats = {
            "web_total": len(web_jobs),
            "fb_total": len(fb_jobs),
        }
        send_no_offers_notification(today_str, scanned_stats)
        print(f"[pipeline] ✅ Done! No new offers email sent.")
        return

    # 3. If NEW OFFERS found: Generate files
    print("[pipeline] Step 3: Generating output files...")
    json_path   = os.path.join(today_str, "pipeline_data.json")
    excel_path  = os.path.join(today_str, "Vietnam_ESL.xlsx")
    pdf_khalid  = os.path.join(today_str, "Cover_Letter_Khalid_Lachhab.pdf")
    pdf_taha    = os.path.join(today_str, "Cover_Letter_Taha_Mahdaoui.pdf")
    script_path = os.path.join(today_str, "vietnam_esl_automation.py")

    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(pipeline_data, f, indent=2, ensure_ascii=False)

    build_excel(excel_path, pipeline_data, new_jobs)

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

    with open(__file__, "r", encoding="utf-8") as src, \
         open(script_path, "w", encoding="utf-8") as dst:
        dst.write(src.read())

    # 4. Upload to Google Drive
    print("[pipeline] Step 4: Uploading to Google Drive...")
    service = get_drive_service()
    folder_id, folder_link = create_daily_drive_folder(service, today_str, PARENT_FOLDER_ID)

    file_links = {}
    for local_file in [json_path, excel_path, pdf_khalid, pdf_taha, script_path]:
        fname = os.path.basename(local_file)
        link = upload_to_drive(service, local_file, folder_id)
        file_links[fname] = link
        print(f"  ✓ Uploaded {fname}")

    # 5. Send digest email
    print("[pipeline] Step 5: Sending digest email...")
    send_notification(folder_link, file_links, today_str, new_jobs)

    # 6. Update history
    for j in new_jobs:
        url = j.get("url")
        if url:
            history[url] = {
                "first_seen": today_str,
                "title": j.get("title", ""),
                "company": j.get("company", ""),
                "location": j.get("location", ""),
            }
    save_job_history(history)

    print(f"[pipeline] ✅ Done! Folder: {folder_link}")


if __name__ == "__main__":
    main()

