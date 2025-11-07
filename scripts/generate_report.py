#!/usr/bin/env python3
import os, re, json
from datetime import datetime, timedelta
from urllib.parse import urljoin
import requests
from bs4 import BeautifulSoup

# Working "update history" pages are under /windows/...
PRIMARY_TARGETS = [
    "https://support.microsoft.com/windows/windows-11-update-history-31ad4770-4f3f-4a2c-97d8-4ab6e5c0bf14",
    "https://support.microsoft.com/windows/windows-10-update-history-3c3d33fa-2d33-96ff-a489-faf6f78b86dd",
    "https://support.microsoft.com/windows/windows-server-2022-update-history-9f96d82e-9a1f-4f43-9be2-491e2d91f66b",
    "https://support.microsoft.com/windows/windows-server-2019-update-history-8c40f96a-391d-4ab9-9f9d-1986fd0e0a58",
]

# Fallbacks (older /topic/ pages) if a primary 404s.
FALLBACK_TARGETS = [
    "https://support.microsoft.com/en-us/topic/windows-11-update-history-31ad4770-4f3f-4a2c-97d8-4ab6e5c0bf14",
    "https://support.microsoft.com/en-us/topic/windows-10-update-history-3c3d33fa-2d33-96ff-a489-faf6f78b86dd",
    "https://support.microsoft.com/en-us/topic/windows-server-2022-update-history-9f96d82e-9a1f-4f43-9be2-491e2d91f66b",
    "https://support.microsoft.com/en-us/topic/windows-server-2019-update-history-8c40f96a-391d-4ab9-9f9d-1986fd0e0a58",
]

HEADERS = {
    "User-Agent": "Mozilla/5.0 (GitHubActions WindowsUpdatesBot)",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "en-US,en;q=0.9",
}

def robust_get(url):
    r = requests.get(url, headers=HEADERS, timeout=45, allow_redirects=True)
    if r.status_code == 404:
        raise requests.exceptions.HTTPError("404", response=r)
    r.raise_for_status()
    return r

def resolve_page(url):
    """Try the /windows/... URL; if 404, try its mapped /topic/... fallback."""
    try:
        return robust_get(url)
    except requests.exceptions.HTTPError as e:
        if e.response is not None and e.response.status_code == 404:
            try:
                idx = PRIMARY_TARGETS.index(url)
                fb = FALLBACK_TARGETS[idx]
                return robust_get(fb)
            except Exception:
                pass
        raise

def guess_os_from_title(t: str) -> str:
    t = t or ""
    if "Windows 11" in t: return "Windows 11"
    if "Windows 10" in t: return "Windows 10 22H2"
    if "Windows Server 2022" in t: return "Windows Server 2022"
    if "Windows Server 2019" in t: return "Windows Server 2019"
    return "Windows"

def parse_kb_items(page_url: str):
    """Collect KB links from an update-history page."""
    try:
        r = resolve_page(page_url)
    except Exception:
        return []
    soup = BeautifulSoup(r.text, "html.parser")
    items = []

    # KBs often appear as links to /help/KBxxxxxxx or /topic/...kbxxxxxxx
    selectors = [
        'a[href*="/help/KB"]',   # /help/KB1234567
        'a[href*="/help/"]',     # /help/1234567
        'a[href*="/topic/"][href*="kb"]',
    ]
    for a in soup.select(", ".join(selectors)):
        href = a.get("href", "")
        if not href:
            continue
        link = urljoin(page_url, href)
        title = a.get_text(strip=True)
        kb_match = re.search(r'KB\d+', title, re.I) or re.search(r'KB\d+', link, re.I)
        if not kb_match:
            continue
        items.append({
            "title": title,
            "link": link,
            "kb": kb_match.group(0).upper(),
        })
    return items

def fetch_kb_details(item):
    """Open the KB page and extract date, description, known issues."""
    try:
        r = robust_get(item["link"])
    except Exception:
        return None

    soup = BeautifulSoup(r.text, "html.parser")

    # Date (published/updated)
    date_text = None
    for sel in ['time', '[data-automation-id="date"]', '.ocpArticlePublishDate', 'h1 + p', 'p']:
        el = soup.select_one(sel)
        if el:
            m = re.search(r'(January|February|March|April|May|June|July|August|September|October|November|December)\s+\d{1,2},\s+\d{4}', el.get_text())
            if m:
                date_text = m.group(0); break
    if not date_text:
        m = re.search(r'(January|February|March|April|May|June|July|August|September|October|November|December)\s+\d{1,2},\s+\d{4}', soup.get_text())
        if m: date_text = m.group(0)

    # Description
    desc = ""
    h1 = soup.select_one('h1')
    if h1:
        p = h1.find_next('p')
        if p: desc = p.get_text(strip=True)
    if not desc:
        p = soup.find('p')
        if p: desc = p.get_text(strip=True)

    # Known issues list
    known = []
    for h in soup.find_all(re.compile('^h[1-6]$')):
        if 'known issues' in h.get_text(strip=True).lower():
            for sib in h.find_all_next(['ul', 'ol', 'p'], limit=6):
                if sib.name in ('ul', 'ol'):
                    for li in sib.find_all('li'):
                        t = ' '.join(li.get_text(' ', strip=True).split())
                        if t: known.append(t)
                if sib.name == 'p':
                    t = ' '.join(sib.get_text(' ', strip=True).split())
                    if t and len(t) > 20: known.append(t)
            break

    # Normalize date
    iso_date = None
    if date_text:
        try:
            iso_date = datetime.strptime(date_text, "%B %d, %Y").date().isoformat()
        except Exception:
            pass

    os_name = guess_os_from_title(item["title"])
    type_label = "Security/Quality"
    if re.search(r'preview', item["title"], re.I) or re.search(r'preview', desc, re.I):
        type_label = "Preview (Non-security)"
    if re.search(r'out-of-band', item["title"], re.I):
        type_label = "Out-of-band"

    return {
        "date": iso_date or datetime.utcnow().date().isoformat(),
        "kb": item["kb"],
        "title": item["title"],
        "os": os_name,
        "link": item["link"],
        "type": type_label,
        "description": desc,
        "known_issues": known
    }

def main():
    cutoff = datetime.utcnow().date() - timedelta(days=30)
    seen = set()
    records = []

    for url in PRIMARY_TARGETS:
        for i in parse_kb_items(url):
            if i["kb"] in seen:
                continue
            d = fetch_kb_details(i)
            if not d:
                continue
            seen.add(i["kb"])
            try:
                dt = datetime.fromisoformat(d["date"]).date()
            except Exception:
                dt = datetime.utcnow().date()
            if dt >= cutoff:
                records.append(d)

    records.sort(key=lambda x: x["date"], reverse=True)

    # Emit HTML + JSON
    os.makedirs("docs", exist_ok=True)
    with open("scripts/templates/report_template.html", "r", encoding="utf-8") as f:
        tpl = f.read()
    html_out = tpl.replace(
        "{{data_json}}", json.dumps(records, ensure_ascii=False)
    ).replace(
        "{{generated_on}}", datetime.utcnow().strftime("%B %d, %Y %H:%M UTC")
    )
    with open("docs/windows-updates.html", "w", encoding="utf-8") as f:
        f.write(html_out)
    with open("docs/windows-updates.json", "w", encoding="utf-8") as f:
        json.dump(records, f, ensure_ascii=False, indent=2)
    print(f"Wrote {len(records)} records")

if __name__ == "__main__":
    main()
