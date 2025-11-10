#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Windows Updates HTML Report – Stable FeedParser Version
Tested: Python 3.11 (GitHub Actions)
"""

import feedparser, html, re
from datetime import datetime, timedelta, timezone

DAYS_BACK = 90
OUTPUT_FILE = "index.html"

FEEDS = [
    ("Windows 11", "https://learn.microsoft.com/en-us/feed/windows/release-health-windows11"),
    ("Windows 10", "https://learn.microsoft.com/en-us/feed/windows/release-health-windows10"),
    ("Windows Server 2022", "https://learn.microsoft.com/en-us/feed/windows/release-health-windows-server-2022"),
    ("Windows Server 2019", "https://learn.microsoft.com/en-us/feed/windows/release-health-windows-server-2019"),
]

def extract_kbs(text):
    return sorted(set(re.findall(r"KB\d+", text or "", re.I)))

def esc(s): return html.escape(s or "", quote=True)

def collect_updates():
    cutoff = datetime.now(timezone.utc) - timedelta(days=DAYS_BACK)
    rows = []
    for name, url in FEEDS:
        feed = feedparser.parse(url)
        for e in feed.entries:
            pub = getattr(e, "published_parsed", None)
            if not pub: continue
            dt = datetime(*pub[:6], tzinfo=timezone.utc)
            if dt < cutoff: continue
            title = e.title
            summary = getattr(e, "summary", "")
            kbs = extract_kbs(title + " " + summary)
            rows.append({
                "os": name,
                "desc": title,
                "kbs": kbs or ["—"],
                "date": dt,
                "src": e.link,
            })
    rows.sort(key=lambda r: r["date"], reverse=True)
    return rows

def build_html(rows):
    gen = datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M UTC")
    head = f"""<!doctype html><html lang=en><meta charset=utf-8>
<title>Windows Updates (Last {DAYS_BACK} Days)</title>
<style>
:root{{color-scheme:dark}}
body{{font-family:Segoe UI,Roboto,Arial,sans-serif;background:#0b1220;color:#e9eef7;margin:0}}
header{{padding:24px 16px;border-bottom:1px solid #1b2743}}
.wrap{{max-width:1200px;margin:0 auto}}
h1{{margin:0 0 8px 0;font-size:22px}}
.meta{{opacity:.75;font-size:13px}}
input{{width:340px;background:#09132a;border:1px solid #20355e;border-radius:8px;padding:8px;color:#e9eef7}}
table{{width:100%;border-collapse:collapse;margin-top:20px}}
th,td{{padding:10px;border-bottom:1px solid #1b2743;font-size:14px;text-align:left;vertical-align:top}}
.kb{{background:#0f2244;border:1px solid #25447a;padding:2px 6px;margin:2px 4px;border-radius:6px;font-size:12px}}
.src{{background:#12274d;border:1px solid #2a4a7d;padding:4px 8px;border-radius:8px;color:#b8d1ff;text-decoration:none;font-size:12px}}
</style>
<body><header><div class=wrap>
<h1>Windows Updates (Last {DAYS_BACK} Days)</h1>
<div class=meta>Generated {gen} — {len(rows)} updates</div>
<input id=q oninput="f()" placeholder="Search KB, OS, description…">
</div></header><div class=wrap>
<table><thead><tr><th>OS</th><th>Description</th><th>KB(s)</th><th>Release Date</th><th>Source</th></tr></thead><tbody>
"""
    body = ""
    for r in rows:
        kb_html = " ".join(f'<span class="kb">{esc(k)}</span>' for k in r["kbs"])
        body += (
            "<tr>"
            f"<td>{esc(r['os'])}</td>"
            f"<td>{esc(r['desc'])}</td>"
            f"<td>{kb_html}</td>"
            f"<td>{r['date'].strftime('%b %d, %Y')}</td>"
            f"<td><a class=\"src\" href=\"{esc(r['src'])}\" target=\"_blank\">Link</a></td>"
            "</tr>\n"
        )
    if not body:
        body = "<tr><td colspan=5>No updates found.</td></tr>"
    tail = """</tbody></table></div>
<script>
function f(){
 const q=document.getElementById('q').value.toLowerCase();
 document.querySelectorAll('tbody tr').forEach(tr=>{
  tr.style.display=tr.textContent.toLowerCase().includes(q)?'':'none';
 });
}
</script></body></html>"""
    return head + body + tail

if __name__ == "__main__":
    rows = collect_updates()
    html_out = build_html(rows)
    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        f.write(html_out)
    print(f"Wrote {len(rows)} updates to {OUTPUT_FILE}")
