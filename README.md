# Vietnam ESL Daily Automation 🇻🇳

Automated pipeline that runs every night at **23:30 Casablanca time** via GitHub Actions.

## What it does

1. **Scrapes** English teaching job listings from VietnamWorks, TopCV, and Facebook groups
2. **Generates** 5 files per day:
   - `Vietnam_ESL.xlsx` — formatted job spreadsheet (Web + Facebook tabs)
   - `Cover_Letter_Khalid_Lachhab.pdf`
   - `Cover_Letter_Taha_Mahdaoui.pdf`
   - `pipeline_data.json` — raw scraped data
   - `vietnam_esl_automation.py` — archived copy of this script
3. **Creates** a dated folder (`DD-MM-YYYY`) in Google Drive under the Job Applications parent
4. **Uploads** all 5 files to that folder
5. **Emails** all 3 recipients with Drive links

## Required GitHub Secrets

Go to `Settings → Secrets and variables → Actions` and add:

| Secret | Value |
|--------|-------|
| `GMAIL_SENDER` | Your Gmail address (e.g. `klachhab@gmail.com`) |
| `GMAIL_APP_PASSWORD` | 16-char Gmail App Password (not your login password) |
| `GDRIVE_SERVICE_ACCOUNT_JSON` | Full contents of the GCP Service Account JSON key file |

## Manual trigger

Go to **Actions → Daily Vietnam ESL Pipeline → Run workflow** anytime to test.

## File structure

```
.
├── .github/workflows/daily_automation.yml   # Cron schedule + CI steps
├── run_cloud_daily.py                        # Main orchestrator
├── scraper.py                                # VietnamWorks + TopCV + Facebook scraper
└── README.md
```
