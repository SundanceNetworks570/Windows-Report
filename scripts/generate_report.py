#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Stable Windows Updates Report Generator
(Working RSS version for current Microsoft feeds)
"""

import html
import re
import time
import urllib.request
import xml.etree.ElementTree as ET
from datetime import datetime, timedelta, timezone

# --------------------------------------------------------------------
# Updated RSS feed list — these return real data
# --------------------------------------------------------------------
FEEDS = [
    ("Windows 11 24H2", "https://www.catalog.update.microsoft.com/Feed.aspx?Product=Windows%2011%2024H2",
     "https://learn.microsoft.com/windows/release-health/"),
    ("Windows 11 23H2", "https://www.catalog.update.microsoft.com/Feed.aspx?Product=Windows%2011%2023H2",
     "https://learn.microsoft.com/windows/release-health/"),
    ("Windows 10 22H2", "https://www.catalog.update.microsoft.com/Feed.aspx?Product=Windows%2010%2022H2",
     "https://learn.microsoft.com/windows/release-health/"),
    ("Windows Server 2022", "https://www.catalog.update.microsoft.com/Feed.aspx?Product=Windows%20Server%202022",
     "https://learn.microsoft.com/windows/release-health/"),
    ("Windows Server 2019", "https://www.catalog.update.microsoft.com/Feed.aspx?Product=Windows%20Server%202019",
     "https://learn.microsoft.com/windows/release-health/")
]

DAYS_BACK = 30
OUTPUT_FILE = "index.html"

USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/124.0 Safari/537.36"
)

# --------------------------------------------------------------------
# Helpers
# --------------------------------------------------------------------
def fetch(url, retries=5, delay=1.5):
    for i in range(retries):
        try:
            req = urllib.request.Request(url, headers={"User-Agent": USER_AGENT})
            with urllib.request.urlopen(req, timeout=20) as resp:
                return resp.read()
        except Exception as e:
            if i == retries - 1:
                print(f"[!] Failed {url}: {e}")
            else:
                time.sleep(delay * (i + 1))
    return b""

def parse_feed(data):
    root = ET.fromstring(data)
    for item in root.findall(".//item"):
        yield {
            "title": (item.findtext("title") or "").strip(),
            "link": (item.findtext("link") or "").strip(),
            "desc": (item.findtext("description") or "").strip(),
            "pub": (item.findtext("pubDate") or "").strip(),
        }

def parse_date(s):
    try:
        return datetime.strptime(s, "%a, %d %b %Y %H:%M:%S %Z").replace(tzinfo=timezone.utc)
    except Exception:
        return datetime(1970, 1, 1, tzinfo=timezone.utc)

def find_kbs(text):
    return sorted(set(re.findall(r"KB\d+", text, re.I)))

def esc(s): return html.escape(s or "")

# --------------------------------------------------------------------
# Main
# --------------------------------------------------------------------
cutoff = datetime.now(timezone.utc) - timedelta(days=DAYS_BACK)
rows = []

for name, url, health in FEEDS:
    data = fetch(url)
    if not data:
        continue
    for item in parse_feed(data):
        dt = parse_date(item["pub"])
        if dt < cutoff:
            continue
        title = item["title"]
        desc = re.sub("<[^>]+>", " ", item["desc"])
        kbs = find_kbs(title + " " + desc)
        rows.append({
            "os": name,
            "desc": title or desc,
            "kbs": kbs or ["—"],
            "date": dt,
            "src": item["link"],
            "health": health,
        })

rows.sort(key=lambda r: r["date"], reverse=True)

# --------------------------------------------------------------------
# HTML
# --------------------------------------------------------------------
gen = datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M UTC")
html_rows = "\n".join(
    f"<tr><td>{esc(r['os'])}</td>"
    f"<td>{esc(r['desc'])}</td>"
    f"<td>{' '.join(f'<span class=kb>{esc(k)}'</span> for k in r['kbs'])}</td>"
    f"<td><a class='known' href='{esc(r['health'])}' target=_blank>Release&nbsp;Health</a></td>"
    f"<td>{r['date'].strftime('%b %d, %Y')}</td>"
    f"<td><a class='src' href='{esc(r['src'])}' target=_blank>Catalog</a></td></tr>"
    for r in rows
) or "<tr><td colspan=6>No updates found in this window.</td></tr>"

html_text = f"""<!doctype html><html lang=en>
<meta charset=utf-8>
<meta name=viewport content="width=device-width,initial-scale=1">
<title>Windows Updates (Last {DAYS_BACK} Days)</title>
<style>
:root{{color-scheme:dark}}
body{{font-family:Segoe UI,Roboto,Arial,sans-serif;background:#0b1220;color:#e9eef7;margin:0}}
header{{padding:24px 16px;border-bottom:1px solid #1b2743}}
h1{{margin:0 0 8px 0;font-size:22px}}
.meta{{opacity:.75;font-size:13px}}
.wrap{{max-width:1200px;margin:0 auto}}
table{{width:100%;border-collapse:collapse;margin-top:20px}}
th,td{{padding:10px;border-bottom:1px solid #1b2743;font-size:14px;text-align:left;vertical-align:top}}
.kb{{background:#0f2244;border:1px solid #25447a;padding:2px 6px;margin:2px 4px;border-radius:6px;font-size:12px}}
.known{{background:#3a2b09;border:1px solid #6a4a12;padding:4px 8px;border-radius:8px;color:#ffd9a8;text-decoration:none;font-size:12px}}
.known:hover{{background:#4c3811}}
.src{{background:#12274d;border:1px solid #2a4a7d;padding:4px 8px;border-radius:8px;color:#b8d1ff;text-decoration:none;font-size:12px}}
.src:hover{{background:#19365f}}
input[type=search]{{width:340px;background:#09132a;border:1px solid #20355e;border-radius:8px;padding:8px;color:#e9eef7}}
</style>
<body><header><div class=wrap>
<h1>Windows Updates (Last {DAYS_BACK} Days)</h1>
<div class=meta>Showing {len(rows)} updates • Generated {gen}</div>
<input type=search id=q oninput="filter()" placeholder="Search KB, description, OS…">
</div></header>
<div class=wrap>
<table><thead><tr><th>OS</th><th>Description</th><th>KB(s)</th><th>Known Issues</th><th>Release Date</th><th>Source</th></tr></thead>
<tbody>{html_rows}</tbody></table>
<div class=meta>Sources: Microsoft Update Catalog RSS; Windows Release Health.</div>
</div>
<script>
function filter(){{const q=document.getElementById('q').value.toLowerCase();
document.querySelectorAll('tbody tr').forEach(tr=>tr.style.display=
tr.textContent.toLowerCase().includes(q)?'':'none');}}
</script>
</body></html>"""

with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
    f.write(html_text)

print(f"Wrote {len(rows)} updates to {OUTPUT_FILE}")
