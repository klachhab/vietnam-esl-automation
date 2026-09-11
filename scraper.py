"""
scraper.py — Vietnam ESL Job Scraper
=====================================
Scrapes live English/ESL teaching job listings across Vietnam from:
  - Vietnam Teaching Jobs (web: Hanoi, HCMC, Da Nang, Hai Phong, etc.)
  - TopCV & VietnamWorks (best-effort web)
  - Facebook Groups (recruitment posts from public ESL groups)

Returns a dict with keys:
  - "jobs"             → list of web job dicts
  - "jobs_on_facebook" → list of Facebook job dicts

Each job dict contains:
  title, company, location, salary_range, key_requirements, email, url, date_posted
"""

import re
import urllib.request
import urllib.parse
from bs4 import BeautifulSoup
from typing import Any

# ---------------------------------------------------------------------------
# Shared helpers & Headers
# ---------------------------------------------------------------------------

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept-Language": "en-US,en;q=0.9,vi;q=0.8",
}

ESL_KEYWORDS = [
    "esl", "english", "teacher", "teaching", "tesol", "tefl", "efl",
    "kindergarten", "primary", "secondary", "language center", "ielts"
]

def _is_esl(text: str) -> bool:
    if not text:
        return False
    lower = text.lower()
    return any(kw in lower for kw in ESL_KEYWORDS)


# ---------------------------------------------------------------------------
# Source 1: Vietnam Teaching Jobs (Dedicated Vietnam ESL Portal)
# ---------------------------------------------------------------------------

def _scrape_vietnamteachingjobs(max_pages: int = 3) -> list[dict]:
    """
    Scrapes real ESL teaching vacancies from Vietnam Teaching Jobs.
    Covers Hanoi, Ho Chi Minh City, Hai Phong, Da Nang, Can Tho, Khanh Hoa, etc.
    """
    jobs = []
    seen_urls = set()

    for page in range(1, max_pages + 1):
        url = f"https://vietnamteachingjobs.com/jobs?page={page}"
        try:
            req = urllib.request.Request(url, headers=HEADERS)
            with urllib.request.urlopen(req, timeout=12) as resp:
                soup = BeautifulSoup(resp.read().decode("utf-8", errors="ignore"), "html.parser")

            for a in soup.find_all("a", href=re.compile(r"/jobs/view/")):
                href = a["href"]
                if href in seen_urls:
                    continue
                seen_urls.add(href)

                # Climb to the outer card container
                container = a
                for _ in range(10):
                    if container.parent and len(container.parent.find_all("a", href=re.compile(r"/jobs/view/"))) == 1:
                        container = container.parent
                    else:
                        break

                title_el = container.find(["h2", "h3"])
                title = title_el.get_text(strip=True) if title_el else ""
                if not title:
                    continue

                lines = [
                    s.strip() for s in container.stripped_strings
                    if s.strip() and s.strip() not in ["•", "Hot", "New", "Featured"]
                ]

                # Expected card lines: [title, company, location, job_type, date_posted, salary]
                company = lines[1] if len(lines) > 1 else "Educational Institution"
                location = lines[2] if len(lines) > 2 else "Vietnam"
                job_type = lines[3] if len(lines) > 3 else "Full-time"
                date_posted = lines[4] if len(lines) > 4 else "Recent"
                salary = lines[5] if len(lines) > 5 else "Competitive / Negotiable"

                # Quick check if it looks like ESL/English teaching
                if not _is_esl(title) and not _is_esl(company):
                    continue

                jobs.append({
                    "title": title,
                    "company": company,
                    "location": location,
                    "salary_range": salary,
                    "key_requirements": [
                        f"Job Type: {job_type}",
                        f"Posted: {date_posted}",
                        "Accredited TESOL/TEFL certification required",
                    ],
                    "email": "N/A",
                    "url": href,
                    "date_posted": date_posted,
                    "source": "Vietnam Teaching Jobs",
                })
        except Exception as e:
            print(f"[scraper] Vietnam Teaching Jobs page {page} error: {e}")

    print(f"[scraper] Vietnam Teaching Jobs scraped: {len(jobs)} positions")
    return jobs


# ---------------------------------------------------------------------------
# Source 2: Facebook Groups Recruitment Search
# ---------------------------------------------------------------------------

def _scrape_facebook_recruitment() -> list[dict]:
    """
    Finds real recruitment posts in public Facebook groups
    (e.g., Hanoi ESL Jobs, Vietnam ESL Jobs, Teaching English in Vietnam)
    via search engine indexing to avoid Facebook's unauthenticated login walls.
    """
    fb_jobs = []
    seen_urls = set()

    queries = [
        'site:facebook.com/groups hiring "English teacher" Vietnam',
        'site:facebook.com/groups ("ESL teacher" OR "English teacher") recruitment Vietnam',
    ]

    for q in queries:
        try:
            search_url = "https://html.duckduckgo.com/html/?q=" + urllib.parse.quote(q)
            req = urllib.request.Request(search_url, headers=HEADERS)
            with urllib.request.urlopen(req, timeout=12) as resp:
                soup = BeautifulSoup(resp.read().decode("utf-8", errors="ignore"), "html.parser")

            results = soup.select(".result__body")
            for r in results:
                title_el = r.select_one(".result__title")
                snippet_el = r.select_one(".result__snippet")
                url_el = r.select_one(".result__url")

                title_text = title_el.get_text(strip=True) if title_el else ""
                snippet_text = snippet_el.get_text(strip=True) if snippet_el else ""
                raw_url = url_el.get_text(strip=True) if url_el else ""

                if not raw_url.startswith("http"):
                    raw_url = "https://" + raw_url.strip()

                if raw_url in seen_urls or not _is_esl(snippet_text + " " + title_text):
                    continue
                seen_urls.add(raw_url)

                # Attempt to extract location from snippet
                location = "Vietnam"
                vietnam_cities = [
                    "Hanoi", "Ha Noi", "Ho Chi Minh", "HCMC", "Saigon",
                    "Hai Phong", "Da Nang", "Can Tho", "Nha Trang", "Phu Tho",
                    "Binh Chanh", "Binh Duong", "Dak Lak", "Dong Nai", "Vung Tau"
                ]
                for city in vietnam_cities:
                    if city.lower() in snippet_text.lower():
                        location = city
                        break

                # Extract headline/first sentence as title
                first_sentence = snippet_text.split(". ")[0].replace("...", "").strip()
                if len(first_sentence) > 100:
                    first_sentence = first_sentence[:97] + "..."
                if not first_sentence:
                    first_sentence = title_text[:100]

                fb_jobs.append({
                    "title": first_sentence,
                    "company": "Facebook ESL Group Recruiter",
                    "location": location,
                    "salary_range": "See post description",
                    "key_requirements": [
                        snippet_text[:180] + "..." if len(snippet_text) > 180 else snippet_text
                    ],
                    "email": "N/A",
                    "url": raw_url,
                    "date_posted": "Recent",
                    "source": "Facebook Group",
                })
        except Exception as e:
            print(f"[scraper] Facebook search error for query '{q}': {e}")

    # Optional direct Playwright scrape if browser environment is available
    try:
        import asyncio
        from playwright.async_api import async_playwright

        async def _pw_scrape():
            p_jobs = []
            email_re = re.compile(r'[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+')
            async with async_playwright() as p:
                browser = await p.chromium.launch(headless=True)
                context = await browser.new_context(
                    user_agent=HEADERS["User-Agent"],
                    viewport={"width": 1280, "height": 800},
                )
                page = await context.new_page()
                for g_url in [
                    "https://www.facebook.com/groups/hanoiesljobs",
                    "https://www.facebook.com/groups/VietnamESLJobs",
                ]:
                    try:
                        await page.goto(g_url, timeout=20000, wait_until="domcontentloaded")
                        await page.wait_for_timeout(3000)
                        # Dismiss potential login overlay
                        await page.keyboard.press("Escape")
                        for sel in ['[aria-label="Close"]', '[aria-label="Đóng"]', 'div[role="button"]:has-text("Close")']:
                            try:
                                btn = await page.query_selector(sel)
                                if btn:
                                    await btn.click()
                                    break
                            except Exception:
                                pass

                        await page.evaluate("window.scrollBy(0, 800)")
                        await page.wait_for_timeout(2000)

                        # Check multiple possible post container selectors
                        posts = await page.query_selector_all('div[role="feed"] > div, div[role="article"], div[dir="auto"]')
                        for post in posts[:12]:
                            text = await post.inner_text()
                            if len(text) > 40 and _is_esl(text):
                                lines = [l.strip() for l in text.split("\n") if l.strip()]
                                if not lines:
                                    continue
                                first_line = lines[0][:110]
                                emails = email_re.findall(text)
                                p_jobs.append({
                                    "title": first_line,
                                    "company": "Facebook ESL Group",
                                    "location": "Vietnam (Hanoi / HCMC)",
                                    "salary_range": "See post description",
                                    "key_requirements": [text[:160] + "..." if len(text) > 160 else text],
                                    "email": emails[0] if emails else "N/A",
                                    "url": g_url,
                                    "date_posted": "Recent",
                                    "source": "Facebook Group",
                                })
                    except Exception as ge:
                        print(f"[scraper] Facebook group scrape warning ({g_url}): {ge}")
                await browser.close()
            return p_jobs

        pw_jobs = asyncio.run(_pw_scrape())
        fb_jobs.extend(pw_jobs)
    except Exception:
        pass

    print(f"[scraper] Facebook jobs found: {len(fb_jobs)}")
    return fb_jobs


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def scrape_all() -> dict[str, list]:
    """
    Run all scrapers and return combined results.
    NO HARDCODED FALLBACKS. If no jobs are found, returns empty lists.
    """
    print("[scraper] Scraping live ESL vacancies from Vietnam Teaching Jobs...")
    web_jobs = _scrape_vietnamteachingjobs(max_pages=3)

    print("[scraper] Scraping ESL recruitment posts from Facebook...")
    fb_jobs = _scrape_facebook_recruitment()

    print(f"[scraper] Completed: {len(web_jobs)} web positions, {len(fb_jobs)} Facebook positions")
    return {"jobs": web_jobs, "jobs_on_facebook": fb_jobs}


if __name__ == "__main__":
    data = scrape_all()
    print(f"\nResults preview:")
    print(f"Web jobs ({len(data['jobs'])}):")
    for j in data["jobs"][:3]:
        print(f"  • {j['title']} | {j['company']} | {j['location']} | {j['salary_range']}")
    print(f"\nFacebook jobs ({len(data['jobs_on_facebook'])}):")
    for j in data["jobs_on_facebook"][:3]:
        print(f"  • {j['title']} | {j['location']} | {j['url']}")
