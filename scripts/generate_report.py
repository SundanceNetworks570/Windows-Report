#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Windows Update Report Generator (Permanent RSS Fix)
---------------------------------------------------
This version uses official Windows release health RSS feeds from Microsoft,
which are actively maintained and contain KB numbers, release dates, and issue notes.
"""

import feedparser
from datetime import datetime, timedelta, timezone
import re, html

# ---------------- Configuration ---------------- #

DAYS_BACK = 90
OUTPUT_FILE = "index.html"

# Official release health RSS feeds
FEEDS = [
    ("Windows 11", "https://learn.microsoft.com/en-us/feed/windows/release-health-windows11"),
    ("Windows 10", "https://learn.microsoft.com/en-us/feed/windows/release-health-windows10"),
    ("Windows Server 2022", "https://learn.microsoft.com/en-us/feed/windows/release-health-windows-server-2022"),
    ("Windows Server 2019", "https://learn.microsoft.com/en-us/feed/windows/release-health-windows-server-2019"),
]

# ---------------- Helpers ---------------- #

def extract_kbs(text):
    return sorted(set(re.findall(r"KB\d+", text or "", re.IGNORECASE)))

def html_escape(s):
    return html.escape(s or "", quote=True)

def fmt_date(dt):
    return dt.strftime("%b %d, %Y")

# ---------------- Core ---------------- #

def collect_updates():
    cutoff = datetime.now(timezone.utc) - timedelta(days=DAYS_BACK)
    all_rows = []

    for product, feed_url in FEEDS:
        feed = feedparser.parse(feed_url)
        for e in feed.entries:
            # Try multiple date sources
            pub_date = getattr(e, "published", getattr(e, "updated", None))
            if not pub_date:
                continue
            try:
                dt = datetime(*e.published_parsed[:6], tzinfo=timezone.utc)
            except Exception:
                continue
            if dt < cutoff:
                continue

            title = e.title
            link = e.link
            summary = getattr(e, "summary", "")

            kbs = extract_kbs(title + " " + summary)
            desc = html_escape(title or summary)
            all_rows.append({
                "product": product,
                "description": desc,
                "kbs": kbs or ["—"],
                "date": dt,
                "source": "Release Health RSS",
                "source_link": link,
                "health_link": "https://learn.microsoft.com/windows/release-health/",
            })

    # Sort newest first
    all_rows.sort(key=lambda x: x["date"], reverse=True)
    return all_rows

# ---------------- HTML Output ---------------- #

def render_html(rows):
    count = len(rows)
    gen_ts = datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M UTC")

    css = """
    :root{color-scheme:dark}
    body{font-family:Inter,Segoe UI,Roboto,Arial,sans-serif;margin:0;background:#0b1220;color:#e9eef7}
    header{padding:28px 20px;border-bottom:1px solid #1b2743}
    h1{margin:0 0 8px 0;font-size:22px}
    .meta{opacity:.75;font-size:13px}
    .wrap{max-width:1200px;margin:0 auto}
    .controls{display:flex;gap:10px;margin:16px 0}
    .btn{padding:8px 12px;border-radius:8px;background:#1a2a44;border:1px solid #243b63;color:#dfe8ff;text-decoration:none;font-size:13px}
    .btn:hover{background:#21365b}
    table{width:100%;border-collapse:collapse;margin:12px 0 40px 0}
    th,td{padding:12px;border-bottom:1px solid #1b2743;vertical-align:top;font-size:14px}
    th{opacity:.9;text-align:left}
    .kb{display:inline-block;background:#0f2244;border:1px solid #25447a;color:#b8d1ff;border-radius:6px;padding:2px 8px;margin:2px 6px 2px 0;font-size:12px}
    .src{display:inline-block;background:#0d1c33;border:1px solid #223b6e;color:#a9c5ff;border-radius:999px;padding:4px 10px;font-size:12px;text-decoration:none}
    .src:hover{background:#12274d}
    .known{display:inline-block;background:#38210a;border:1px solid #6a3c12;color:#ffd9a8;border-radius:999px;padding:4px 10px;font-size:12px;text-decoration:none}
    .known:hover{background:#4a2b0f}
    input[type="search"]{width:360px;background:#09132a;border:1px solid #20355e;border-radius:8px;padding:8px 10px;color:#e9eef7}
    """

    js = """
    function filterRows(){
      const q = document.getElementById('q').value.toLowerCase();
      const rows = document.querySelectorAll('tbody tr');
      rows.forEach(tr=>{
        const t = tr.textContent.toLowerCase();
        tr.style.display = (t.indexOf(q) !== -1) ? '' : 'none';
      });
    }
    """

    rows_html = ""
    for r in rows:
        kb_html = " ".join(f'<span class="kb">{html_escape(kb)}</span>' for kb in r["kbs"])
        rows_html += f"""
        <tr>
          <td>{html_escape(r["product"])}</td>
          <td>{html_escape(r["description"])}</td>
          <td>{kb_html}</td>
          <td><a class="known" href="{html_escape(r["health_link"])}" target="_blank">Release Health</a></td>
          <td>{fmt_date(r["date"])}</td>
          <td><a class="src" href="{html_escape(r["source_link"])}" target="_blank">{r["source"]}</a></td>
        </tr>
        """

    html_out = f"""<!doctype html>
<html lang="en">
<meta charset="utf-8">
<title>Windows Updates (Last {DAYS_BACK} Days)</title>
<style>{css}</style>
<body>
<header>
  <div class="wrap">
    <h1>Windows Updates (Last {DAYS_BACK} Days)</h1>
    <div class="meta">Showing {count} updates • Generated {gen_ts}</div>
    <div class="controls">
      <input id="q" type="search" placeholder="Search KB, OS, or description..." oninput="filterRows()">
      <a class="btn" href="#" onclick="window.print();return false;">Print / PDF</a>
    </div>
  </div>
</header>
<div class="wrap">
<table>
  <thead>
    <tr><th>OS</th><th>Description</th><th>KB(s)</th><th>Known Issues</th><th>Release Date</th><th>Source</th></tr>
  </thead>
  <tbody>
    {rows_html or '<tr><td colspan="6">No updates found in the last 30 days.</td></tr>'}
  </tbody>
</table>
</div>
<script>{js}</script>
</body>
</html>"""

    return html_out


def main():
    rows = collect_updates()
    html_text = render_html(rows)
    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        f.write(html_text)
    print(f"Wrote {len(rows)} updates to {OUTPUT_FILE}")


if __name__ == "__main__":
    main()
