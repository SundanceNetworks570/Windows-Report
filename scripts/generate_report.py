#!/usr/bin/env python3
import os, re, json, time
from datetime import datetime, timedelta
from urllib.parse import urljoin
import requests
from bs4 import BeautifulSoup

TARGETS = [
    "https://support.microsoft.com/en-us/topic/windows-11-update-history-31ad4770-4f3f-4a2c-97d8-4ab6e5c0bf14",
    "https://support.microsoft.com/en-us/topic/windows-10-update-history-3c3d33fa-2d33-96ff-a489-faf6f78b86dd",
    "https://support.microsoft.com/en-us/topic/windows-server-2022-update-history-9f96d82e-9a1f-4f43-9be2-491e2d91f66b",
    "https://support.microsoft.com/en-us/topic/windows-server-2019-update-history-8c40f96a-391d-4ab9-9f9d-1986fd0e0a58"
]
HEADERS = {"User-Agent": "Mozilla/5.0 (GitHubAction WindowsUpdatesBot)"}

def guess_os_from_title(t):
    if not t: return "Windows"
    if "Windows 11" in t: return "Windows 11"
    if "Windows 10" in t: return "Windows 10 22H2"
    if "Windows Server 2022" in t: return "Windows Server 2022"
    if "Windows Server 2019" in t: return "Windows Server 2019"
    return "Windows"

def parse_kb_items(page_url):
    r = requests.get(page_url, headers=HEADERS, timeout=45)
    r.raise_for_status()
    soup = BeautifulSoup(r.text, "html.parser")
    items = []
    for a in soup.select('a[href*="kb"]'):
        href = a.get("href", "")
        if not href or not re.search(r'/kb\\d+', href, re.I):
            continue
        link = urljoin(page_url, href)
        title = a.get_text(strip=True)
        kb_match = re.search(r'KB\\d+', title) or re.search(r'KB\\d+', link)
        if not kb_match:
            continue
        items.append({"title": title, "link": link, "kb": kb_match.group(0)})
    return items

def fetch_kb_details(item):
    try:
        r = requests.get(item["link"], headers=HEADERS, timeout=45)
        r.raise_for_status()
        soup = BeautifulSoup(r.text, "html.parser")
        # Date
        date_text = None
        for sel in ['time', '[data-automation-id="date"]', '.ocpArticlePublishDate', 'h1 + p']:
            el = soup.select_one(sel)
            if el:
                m = re.search(r'(January|February|March|April|May|June|July|August|September|October|November|December)\\s+\\d{1,2},\\s+\\d{4}', el.get_text())
                if m:
                    date_text = m.group(0); break
        # Fallback
        if not date_text:
            m = re.search(r'(January|February|March|April|May|June|July|August|September|October|November|December)\\s+\\d{1,2},\\s+\\d{4}', soup.get_text())
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
        # Known issues
        known = []
        for h in soup.find_all(re.compile('^h[1-6]$')):
            if 'known issues' in h.get_text(strip=True).lower():
                for sib in h.find_all_next(['ul','ol','p'], limit=6):
                    if sib.name in ('ul','ol'):
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
    except Exception:
        return None

def main():
    cutoff = datetime.utcnow().date() - timedelta(days=30)
    seen = set()
    records = []
    for url in TARGETS:
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
    # Sort and emit HTML
    records.sort(key=lambda x: x["date"], reverse=True)
    os.makedirs("docs", exist_ok=True)
    with open("scripts/templates/report_template.html","r",encoding="utf-8") as f:
        tpl = f.read()
    html_out = tpl.replace("{{data_json}}", json.dumps(records, ensure_ascii=False)).replace("{{generated_on}}", datetime.utcnow().strftime("%B %d, %Y %H:%M UTC"))
    with open("docs/windows-updates.html","w",encoding="utf-8") as f:
        f.write(html_out)
    with open("docs/windows-updates.json","w",encoding="utf-8") as f:
        json.dump(records, f, ensure_ascii=False, indent=2)
    print(f"Wrote {len(records)} records")

if __name__ == "__main__":
    main()
