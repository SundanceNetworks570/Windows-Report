#!/usr/bin/env python3
import os, re, json
from datetime import datetime, timedelta
from urllib.parse import urljoin
import requests
from bs4 import BeautifulSoup

# === Official Microsoft update-history pages (with en-us prefix) ===
TARGETS = {
    "Windows 11": "https://support.microsoft.com/en-us/windows/windows-11-update-history-31ad4770-4f3f-4a2c-97d8-4ab6e5c0bf14",
    "Windows 10 22H2": "https://support.microsoft.com/en-us/windows/windows-10-update-history-3c3d33fa-2d33-96ff-a489-faf6f78b86dd",
    "Windows Server 2022": "https://support.microsoft.com/en-us/windows/windows-server-2022-update-history-9f96d82e-9a1f-4f43-9be2-491e2d91f66b",
    "Windows Server 2019": "https://support.microsoft.com/en-us/windows/windows-server-2019-update-history-8c40f96a-391d-4ab9-9f9d-1986fd0e0a58",
}

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) GitHubActionsBot",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "en-US,en;q=0.9",
}

def fetch(url):
    """Fetch a page; returns None on failure (never throws)."""
    try:
        r = requests.get(url, headers=HEADERS, timeout=45)
        r.raise_for_status()
        return r
    except Exception as e:
        print(f"⚠️ Failed to fetch {url} -> {e}")
        return None

def parse_kb_links(html_text, base_url):
    """Find KB links on an update-history page."""
    soup = BeautifulSoup(html_text, "html.parser")
    items = []
    for a in soup.find_all("a", href=True):
        href = a["href"]
        text = a.get_text(strip=True)
        if not href:
            continue
        # KBs usually appear as KB####### or /help/#######
        if "KB" in text or "/help/" in href:
            m = re.search(r"KB\d+", text) or re.search(r"KB\d+", href)
            if not m:
                continue
            items.append({
                "title": text,
                "link": urljoin(base_url, href),
                "kb": m.group(0).upper()
            })
    return items

def parse_kb_page(item):
    """Extract date, description, known issues from a KB page."""
    r = fetch(item["link"])
    if not r:
        return None
    soup = BeautifulSoup(r.text, "html.parser")

    # Date
    m = re.search(
        r"(January|February|March|April|May|June|July|August|September|October|November|December)\s+\d{1,2},\s+\d{4}",
        soup.get_text(" ", strip=True),
    )
    date_iso = datetime.utcnow().date().isoformat()
    if m:
        try:
            date_iso = datetime.strptime(m.group(0), "%B %d, %Y").date().isoformat()
        except Exception:
            pass

    # Description
    desc = ""
    p = soup.find("p")
    if p:
        desc = p.get_text(" ", strip=True)

    # Known issues
    known = []
    for h in soup.find_all(re.compile("^h[1-6]$")):
        if "known issues" in h.get_text(strip=True).lower():
            for li in h.find_all_next("li"):
                t = li.get_text(" ", strip=True)
                if t:
                    known.append(t)
            break

    item.update({
        "date": date_iso,
        "description": desc,
        "known_issues": known,
        "type": "Security/Quality",
    })
    return item

def main():
    cutoff = datetime.utcnow().date() - timedelta(days=30)
    updates = []

    for os_name, url in TARGETS.items():
        print(f"Processing {os_name} ...")
        page = fetch(url)
        if not page:
            continue
        kb_items = parse_kb_links(page.text, url)
        for kb in kb_items:
            d = parse_kb_page(kb)
            if not d:
                continue
            d["os"] = os_name
            try:
                when = datetime.fromisoformat(d["date"]).date()
            except Exception:
                when = datetime.utcnow().date()
            if when >= cutoff:
                updates.append(d)

    updates.sort(key=lambda x: x["date"], reverse=True)

    # Always emit HTML/JSON (even if 0 updates) so Pages has a file
    os.makedirs("docs", exist_ok=True)
    with open("scripts/templates/report_template.html", "r", encoding="utf-8") as f:
        tpl = f.read()
    html_out = tpl.replace("{{data_json}}", json.dumps(updates, ensure_ascii=False)) \
                  .replace("{{generated_on}}", datetime.utcnow().strftime("%B %d, %Y %H:%M UTC"))
    with open("docs/windows-updates.html", "w", encoding="utf-8") as f:
        f.write(html_out)
    with open("docs/windows-updates.json", "w", encoding="utf-8") as f:
        json.dump(updates, f, indent=2, ensure_ascii=False)

    print(f"✅ Collected {len(updates)} updates")

if __name__ == "__main__":
    main()
