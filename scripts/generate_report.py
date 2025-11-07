#!/usr/bin/env python3
import os, re, json, xml.etree.ElementTree as ET
from datetime import datetime, timedelta
from urllib.parse import urljoin
import requests
from bs4 import BeautifulSoup

HEADERS = {"User-Agent": "Mozilla/5.0 (GitHubActions WindowsUpdatesBot)"}

# RSS feeds by OS family
FEEDS = {
    "Windows 11": "https://support.microsoft.com/en-us/feed/subject/windows-11-update-history",
    "Windows 10 22H2": "https://support.microsoft.com/en-us/feed/subject/windows-10-update-history",
    "Windows Server 2022": "https://support.microsoft.com/en-us/feed/subject/windows-server-2022-update-history",
    "Windows Server 2019": "https://support.microsoft.com/en-us/feed/subject/windows-server-2019-update-history",
}

def parse_feed(os_name, url):
    """Return simplified list of KBs from Microsoft support RSS feed."""
    out = []
    try:
        r = requests.get(url, headers=HEADERS, timeout=30)
        r.raise_for_status()
        root = ET.fromstring(r.text)
        for item in root.findall(".//item"):
            title = item.findtext("title") or ""
            link = item.findtext("link") or ""
            desc = (item.findtext("description") or "").strip()
            pub = item.findtext("pubDate") or ""
            if not pub:
                continue
            date = datetime.strptime(pub[:16], "%a, %d %b %Y").date().isoformat()
            kb_match = re.search(r"KB\d+", title) or re.search(r"KB\d+", link)
            kb = kb_match.group(0) if kb_match else "(n/a)"
            type_label = "Security/Quality"
            if re.search("Preview", title, re.I):
                type_label = "Preview (Non-security)"
            out.append({
                "date": date,
                "kb": kb,
                "title": title,
                "os": os_name,
                "link": link,
                "type": type_label,
                "description": desc,
                "known_issues": []
            })
    except Exception as e:
        print(f"Failed {os_name}: {e}")
    return out

def main():
    cutoff = datetime.utcnow().date() - timedelta(days=30)
    records = []
    for os_name, feed in FEEDS.items():
        for d in parse_feed(os_name, feed):
            try:
                dt = datetime.fromisoformat(d["date"]).date()
                if dt >= cutoff:
                    records.append(d)
            except Exception:
                pass
    records.sort(key=lambda x: x["date"], reverse=True)
    print(f"Collected {len(records)} updates")

    os.makedirs("docs", exist_ok=True)
    tpl = open("scripts/templates/report_template.html", encoding="utf-8").read()
    html_out = tpl.replace(
        "{{data_json}}", json.dumps(records, ensure_ascii=False)
    ).replace(
        "{{generated_on}}", datetime.utcnow().strftime("%B %d, %Y %H:%M UTC")
    )
    with open("docs/windows-updates.html", "w", encoding="utf-8") as f:
        f.write(html_out)
    with open("docs/windows-updates.json", "w", encoding="utf-8") as f:
        json.dump(records, f, ensure_ascii=False, indent=2)

if __name__ == "__main__":
    main()
