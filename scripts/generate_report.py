#!/usr/bin/env python3
import os, re, json
from datetime import datetime, timedelta
from urllib.parse import urljoin
import requests
from bs4 import BeautifulSoup

# === Microsoft Update History URLs ===
TARGETS = {
    "Windows 11": "https://support.microsoft.com/windows/windows-11-update-history-31ad4770-4f3f-4a2c-97d8-4ab6e5c0bf14",
    "Windows 10 22H2": "https://support.microsoft.com/windows/windows-10-update-history-3c3d33fa-2d33-96ff-a489-faf6f78b86dd",
    "Windows Server 2022": "https://support.microsoft.com/windows/windows-server-2022-update-history-9f96d82e-9a1f-4f43-9be2-491e2d91f66b",
    "Windows Server 2019": "https://support.microsoft.com/windows/windows-server-2019-update-history-8c40f96a-391d-4ab9-9f9d-1986fd0e0a58"
}

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) GitHubActionsBot",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8"
}


# === Helper Functions ===
def fetch(url):
    """Download a page safely."""
    r = requests.get(url, headers=HEADERS, timeout=45)
    if r.status_code == 404:
        print(f"⚠️ 404 Not Found: {url}")
        return None
    r.raise_for_status()
    return r


def parse_kb_links(html_text, base_url):
    """Extract KB articles from a given update history page."""
    soup = BeautifulSoup(html_text, "html.parser")
    items = []
    for a in soup.find_all("a", href=True):
        href = a["href"]
        text = a.get_text(strip=True)
        if not href:
            continue
        if "KB" in text or "/help/" in href:
            kb_match = re.search(r"KB\d+", text) or re.search(r"KB\d+", href)
            if kb_match:
                items.append({
                    "title": text,
                    "link": urljoin(base_url, href),
                    "kb": kb_match.group(0)
                })
    return items


def parse_kb_page(item):
    """Parse individual KB page for details."""
    try:
        r = fetch(item["link"])
        if not r:
            return None
    except Exception:
        return None

    soup = BeautifulSoup(r.text, "html.parser")

    # Extract publication date
    m = re.search(r"(January|February|March|April|May|June|July|August|September|October|November|December)\s+\d{1,2},\s+\d{4}", soup.get_text())
    date_str = m.group(0) if m else None
    if date_str:
        try:
            date_iso = datetime.strptime(date_str, "%B %d, %Y").date().isoformat()
        except Exception:
            date_iso = datetime.utcnow().date().isoformat()
    else:
        date_iso = datetime.utcnow().date().isoformat()

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


# === Main Logic ===
def main():
    cutoff = datetime.utcnow().date() - timedelta(days=30)
    results = []

    for os_name, url in TARGETS.items():
        print(f"Processing {os_name} ...")
        page = fetch(url)
        if not page:
            print(f"Failed to fetch {url}")
            continue

        kb_list = parse_kb_links(page.text, url)
        for kb in kb_list:
            detail = parse_kb_page(kb)
            if not detail:
                continue
            detail["os"] = os_name

            try:
                kb_date = datetime.fromisoformat(detail["date"]).date()
            except Exception:
                kb_date = datetime.utcnow().date()

            if kb_date >= cutoff:
                results.append(detail)

    results.sort(key=lambda x: x["date"], reverse=True)

    os.makedirs("docs", exist_ok=True)
    with open("scripts/templates/report_template.html", "r", encoding="utf-8") as f:
        template = f.read()

    html_out = template.replace(
        "{{data_json}}", json.dumps(results, ensure_ascii=False)
    ).replace(
        "{{generated_on}}", datetime.utcnow().strftime("%B %d, %Y %H:%M UTC")
    )

    with open("docs/windows-updates.html", "w", encoding="utf-8") as f:
        f.write(html_out)

    with open("docs/windows-updates.json", "w", encoding="utf-8") as f:
        json.dump(results, f, indent=2, ensure_ascii=False)

    print(f"✅ Collected {len(results)} updates")


if __name__ == "__main__":
    main()
