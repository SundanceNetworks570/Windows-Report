#!/usr/bin/env python3
import os, re, json
from datetime import datetime, timedelta
from urllib.parse import urljoin
import requests
from bs4 import BeautifulSoup

# --- Update-history pages ---
TARGETS = {
    "Windows 11": "https://support.microsoft.com/windows/windows-11-update-history-31ad4770-4f3f-4a2c-97d8-4ab6e5c0bf14",
    "Windows 10 22H2": "https://support.microsoft.com/windows/windows-10-update-history-3c3d33fa-2d33-96ff-a489-faf6f78b86dd",
    "Windows Server 2022": "https://support.microsoft.com/windows/windows-server-2022-update-history-9f96d82e-9a1f-4f43-9be2-491e2d91f66b",
    "Windows Server 2019": "https://support.microsoft.com/windows/windows-server-2019-update-history-8c40f96a-391d-4ab9-9f9d-1986fd0e0a58",
}

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) GitHubActionsBot",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
}

def fetch(url):
    r = requests.get(url, headers=HEADERS, timeout=45)
    r.raise_for_status()
    return r

def parse_kb_links(page_html, base_url):
    soup = BeautifulSoup(page_html, "html.parser")
    items = []
    for a in soup.find_all("a", href=True):
        href = a["href"]
        if "KB" in a.text or "/help/" in href:
            kb_match = re.search(r'KB\d+', a.text) or re.search(r'KB\d+', href)
            if not kb_match:
                continue
            items.append({
                "title": a.get_text(strip=True),
                "link": urljoin(base_url, href),
                "kb": kb_match.group(0)
            })
    return items

def parse_kb_page(item):
    """Fetch KB detail: date, description, known issues."""
    try:
        r = fetch(item["link"])
    except Exception:
        return None
    soup = BeautifulSoup(r.text, "html.parser")

    # date
    m = re.search(r'(January|February|March|April|May|June|July|August|September|October|November|December)\s+\d{1,2},\s+\d{4}', soup.get_text())
    date_str = m.group(0) if m else None
    if date_str:
        try:
            date_iso = datetime.strptime(date_str, "%B %d, %Y").date().isoformat()
        except Exception:
            date_iso = datetime.utcnow().date().isoformat()
    else:
        date_iso = datetime.utcnow().date().isoformat()

    # description
    desc = ""
    first_p = soup.find("p")
    if first_p:
        desc = first_p.get_text(" ", strip=True)

    # known issues
    known = []
    for h in soup.find_all(re.compile("^h[1-6]$")):
        if "known issues" in h.get_text(strip=True).lower():
            nxt = h.find_next(["ul", "ol", "p"])
            if nxt:
                for li in nxt.find_all("li"):
                    txt = li.get_text(" ", strip=True)
                    if txt:
                        known.append(txt)
            break

    item.update({
        "date": date_iso,
        "description": desc,
        "known_issues": known,
        "type": "Security/Quality"
    })
    return item

def main():
    all_updates = []
    cutoff = datetime.utcnow().date() - timedelta(days=30)

    for os_name, url in TARGETS.items():
        print(f"Processing {os_name} ...")
        try:
            page = fetch(url)
        except Exception as e:
            print(f"Failed {os_name}: {e}")
            continue
        kb_items = parse_kb_links(page.text, url)
        for kb in kb_items:
            detail = parse_kb_page(kb)
            if not detail:
                continue
            detail["os"] = os_name
            try:
                if datetime.fromisoformat(detail["date"]).date() >= cutoff:
                    all_updates.append(detail)
            except Exception:
                all_updates.append(detail)

    all_updates.sort(key=lambda x: x["date"], reverse=True)

    os.makedirs("docs", exist_ok=True)
    with open("scripts/templates/report_template.html", "r", encoding="utf-8") as f:
        tpl = f.read()

    html_out = tpl.replace(
        "{{data_json}}", json.dumps(all_updates, ensure_ascii=False)
    ).replace(
        "{{generated_on}}", datetime.utcnow().strftime("%B %d, %Y %H:%M UTC")
    )

    with open("docs/windows-updates.html", "w", encoding="utf-8") as f:
        f.write(html_out)
    with open("docs/windows-updates.json", "w", encoding="utf-8") as f:
        json.dump(all_updates, f, indent=2, ensure_ascii=False)

    print(f"Collected {len(all_updates)} updates")

if __name__ == "__main__":
    main()
