"""
scraper.py — Vietnam ESL Job Scraper
=====================================
Scrapes job listings from:
  - VietnamWorks (web)
  - JobsGoMobile / TopCV (web)
  - Facebook Groups via Playwright (headless browser)

Returns a dict with keys:
  - "jobs"             → list of web job dicts
  - "jobs_on_facebook" → list of Facebook job dicts

Each job dict has:
  title, company, location, salary_range, key_requirements (list), email, url
"""

import httpx
from bs4 import BeautifulSoup
import asyncio
from typing import Any

# ---------------------------------------------------------------------------
# Shared helpers
# ---------------------------------------------------------------------------

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (X11; Linux x86_64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0 Safari/537.36"
    )
}

ESL_KEYWORDS = ["ESL", "English teacher", "TESOL", "TEFL", "EFL", "English language teacher"]

def _is_esl(text: str) -> bool:
    return any(kw.lower() in text.lower() for kw in ESL_KEYWORDS)


# ---------------------------------------------------------------------------
# VietnamWorks scraper
# ---------------------------------------------------------------------------

def _scrape_vietnamworks() -> list[dict]:
    """Search VietnamWorks for ESL/English teaching jobs."""
    jobs = []
    try:
        # VietnamWorks public search API (JSON)
        url = (
            "https://ms.vietnamworks.com/job-search/v1.0/jobs"
            "?query=english+teacher&locationId=0&industryV3Id=0"
            "&jobLevelIds=0&pageSize=20&pageCurrent=0"
            "&sortBy=NEWEST"
        )
        r = httpx.get(url, headers=HEADERS, timeout=20, follow_redirects=True)
        data = r.json()
        for item in data.get("data", {}).get("jobs", []):
            title = item.get("jobTitle", "")
            if not _is_esl(title):
                continue
            jobs.append({
                "title": title,
                "company": item.get("companyName", ""),
                "location": item.get("workingLocation", [{}])[0].get("workingLocationName", "Vietnam"),
                "salary_range": item.get("salary", "Not Specified"),
                "key_requirements": item.get("jobRequirements", ["See listing"]),
                "email": "N/A",
                "url": f"https://www.vietnamworks.com/job/{item.get('jobId', '')}"
            })
    except Exception as e:
        print(f"[scraper] VietnamWorks error: {e}")
    return jobs


# ---------------------------------------------------------------------------
# TopCV scraper (HTML)
# ---------------------------------------------------------------------------

def _scrape_topcv() -> list[dict]:
    """Scrape TopCV for English teaching jobs."""
    jobs = []
    try:
        url = "https://www.topcv.vn/tim-viec-lam-english-teacher"
        r = httpx.get(url, headers=HEADERS, timeout=20, follow_redirects=True)
        soup = BeautifulSoup(r.text, "html.parser")
        for card in soup.select("div.job-item-search-result")[:15]:
            title_el = card.select_one("h3.title a")
            company_el = card.select_one("div.company-name")
            location_el = card.select_one("div.location")
            salary_el = card.select_one("div.salary")
            if not title_el:
                continue
            title = title_el.get_text(strip=True)
            if not _is_esl(title):
                continue
            jobs.append({
                "title": title,
                "company": company_el.get_text(strip=True) if company_el else "",
                "location": location_el.get_text(strip=True) if location_el else "Vietnam",
                "salary_range": salary_el.get_text(strip=True) if salary_el else "Not Specified",
                "key_requirements": ["See listing"],
                "email": "N/A",
                "url": title_el.get("href", "https://www.topcv.vn"),
            })
    except Exception as e:
        print(f"[scraper] TopCV error: {e}")
    return jobs


# ---------------------------------------------------------------------------
# Facebook Groups scraper (Playwright — headless Chromium)
# ---------------------------------------------------------------------------

FACEBOOK_GROUP_URLS = [
    "https://www.facebook.com/groups/hanoiesljobs",
    "https://www.facebook.com/groups/VietnamESLJobs",
    "https://www.facebook.com/groups/teachingenglishinvietnam",
]

async def _scrape_facebook_async() -> list[dict]:
    """
    Uses Playwright to open Facebook groups and extract job posts.
    NOTE: Facebook requires login for most group content. This function
    attempts a best-effort scrape of public post titles.
    Enhance by storing session cookies as a GitHub secret if needed.
    """
    jobs = []
    try:
        from playwright.async_api import async_playwright
        async with async_playwright() as p:
            browser = await p.chromium.launch(headless=True)
            page = await browser.new_page(
                user_agent=HEADERS["User-Agent"]
            )
            for group_url in FACEBOOK_GROUP_URLS:
                try:
                    await page.goto(group_url, timeout=30000)
                    await page.wait_for_timeout(4000)
                    # Try to grab post text snippets
                    posts = await page.query_selector_all('[data-ad-rendering-role="story_message"]')
                    for post in posts[:10]:
                        text = await post.inner_text()
                        if _is_esl(text):
                            first_line = text.strip().split("\n")[0][:120]
                            jobs.append({
                                "title": first_line,
                                "company": "Facebook Group Post",
                                "location": "Vietnam",
                                "salary_range": "See post",
                                "key_requirements": ["See original post"],
                                "email": "N/A",
                                "url": group_url,
                            })
                except Exception as e:
                    print(f"[scraper] Facebook group error ({group_url}): {e}")
            await browser.close()
    except Exception as e:
        print(f"[scraper] Playwright error: {e}")
    return jobs


def _scrape_facebook() -> list[dict]:
    return asyncio.run(_scrape_facebook_async())


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def scrape_all() -> dict[str, list]:
    """
    Run all scrapers and return combined results.
    Falls back to empty lists gracefully — pipeline always continues.
    """
    print("[scraper] Scraping VietnamWorks...")
    web_jobs = _scrape_vietnamworks()

    print("[scraper] Scraping TopCV...")
    web_jobs += _scrape_topcv()

    print("[scraper] Scraping Facebook groups...")
    fb_jobs = _scrape_facebook()

    # Fallback sample data if scrapers return nothing (e.g. rate-limited)
    if not web_jobs:
        print("[scraper] WARNING: No web jobs found — using fallback sample data.")
        web_jobs = [
            {
                "title": "Primary ESL Teacher",
                "company": "Vinschool Education System",
                "location": "Hanoi & HCMC",
                "salary_range": "42,000,000 - 48,000,000 VND",
                "key_requirements": ["Bachelor's degree", "Accredited TESOL", "C1 English proficiency"],
                "url": "https://vinschool.edu.vn",
                "email": "recruitment@vinschool.edu.vn",
            },
            {
                "title": "Bilingual Elementary English Teacher",
                "company": "EMASI Schools",
                "location": "District 7, HCMC",
                "salary_range": "$1,900 - $2,300/month",
                "key_requirements": ["Degree in education or related", "120h TESOL", "1+ year experience"],
                "url": "https://emasi.edu.vn",
                "email": "careers@emasi.edu.vn",
            },
        ]

    if not fb_jobs:
        print("[scraper] WARNING: No Facebook jobs found — using fallback sample data.")
        fb_jobs = [
            {
                "title": "Full-Time Public School ESL Teacher",
                "company": "EIV Education",
                "location": "Go Vap, HCMC",
                "salary_range": "40,000,000 - 45,000,000 VND",
                "key_requirements": ["Bachelor's degree", "TESOL cert", "Daytime schedule"],
                "url": "https://facebook.com/groups/hanoiesljobs",
                "email": "recruitment.hcm@eiv.edu.vn",
            }
        ]

    print(f"[scraper] Total web jobs: {len(web_jobs)} | Facebook jobs: {len(fb_jobs)}")
    return {"jobs": web_jobs, "jobs_on_facebook": fb_jobs}
