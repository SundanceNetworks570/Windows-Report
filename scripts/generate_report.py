#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Stable Windows Updates report generator.

- Uses Microsoft Update Catalog RSS feeds (reliable + include KB numbers)
- Writes a single index.html at repo root
- Adds OS-specific "Release Health" known-issues links (stable landing pages)
- Filters to last 30 days
- Retries with backoff, browser UA, and timeouts

This avoids the fragile support.microsoft.com "topic" pages that 404.
"""

import sys
import re
import time
import math
import html
from datetime import datetime, timedelta, timezone
from typing import List, Dict, Any
import xml.etree.ElementTree as ET
import urllib.request

# -----------------------------
# Configuration
# -----------------------------

# Microsoft Update Catalog RSS (use HTTP — this endpoint is plain RSS and very stable)
FEEDS = [
    # Display Name, RSS URL, Release Health landing (for Known Issues button)
    ("Windows 11", "http://www.catalog.update.microsoft.com/Feed.aspx?Product=Windows%2011",
     "https://learn.microsoft.com/windows/release-health/"),
    ("Windows 10", "http://www.catalog.update.microsoft.com/Feed.aspx?Product=Windows%2010",
     "https://learn.microsoft.com/windows/release-health/"),
    ("Windows Server 2022", "http://www.catalog.update.microsoft.com/Feed.aspx?Product=Windows%20Server%202022",
     "https://learn.microsoft.com/windows/release-health/"),
    ("Windows Server 2019", "http://www.catalog.update.microsoft.com/Feed.aspx?Product=Windows%20Server%202019",
     "https://learn.microsoft.com/windows/release-health/"),
    ("Windows Server 2016", "http://www.catalog.update.microsoft.com/Feed.aspx?Product=Windows%20Server%202016",
     "https://learn.microsoft.com/windows/release-health/"),
]

DAYS_BACK = 30
OUTPUT_FILE = "index.html"

# Browser-y UA so catalog doesn’t act funny
USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/124.0 Safari/537.36"
)

# -----------------------------
# Utilities
# -----------------------------

def http_get(url: str, timeout: int = 20, attempts: int = 5, base_delay: float = 0.75) -> bytes:
    """GET with retry/backoff and browser UA."""
    last = None
    for i in range(attempts):
        try:
            req = urllib.request.Request(url, headers={"User-Agent": USER_AGENT})
            with urllib.request.urlopen(req, timeout=timeout) as resp:
                return resp.read()
        except Exception as e:
            last = e
            # backoff
            sleep_s = base_delay * (2 ** i) + (i * 0.1)
            time.sleep(min(8.0, sleep_s))
    raise RuntimeError(f"Failed to fetch {url}: {last}")

def parse_rss(xml_bytes: bytes) -> List[Dict[str, Any]]:
    """Parse a simple RSS feed from Microsoft Update Catalog."""
    items = []
    # The feed is standard RSS 2.0
    try:
        root = ET.fromstring(xml_bytes)
    except ET.ParseError:
        return items

    # Find all <item>
    for item in root.findall(".//item"):
        title = (item.findtext("title") or "").strip()
        link = (item.findtext("link") or "").strip()
        pub_date = (item.findtext("pubDate") or "").strip()
        description = (item.findtext("description") or "").strip()

        items.append({
            "title": title,
            "link": link,
            "pubDate": pub_date,
            "description": description,
        })
    return items

KB_RE = re.compile(r"\bKB\d+\b", re.IGNORECASE)

def extract_kbs(text: str) -> List[str]:
    """Extract KB numbers from text."""
    return sorted(set(m.group(0).upper() for m in KB_RE.finditer(text or "")))

def try_parse_date_rss(s: str) -> datetime:
    """
    Parse RFC822-ish dates from RSS.
    Example: Tue, 05 Nov 2025 00:00:00 GMT
    """
    fmts = [
        "%a, %d %b %Y %H:%M:%S %Z",
        "%a, %d %b %Y %H:%M:%S GMT",
        "%a, %d %b %Y %H:%M:%S +0000",
        "%d %b %Y %H:%M:%S %Z",
    ]
    for f in fmts:
        try:
            return datetime.strptime(s, f).replace(tzinfo=timezone.utc)
        except Exception:
            pass
    # If unknown, return epoch so it likely filters out
    return datetime(1970, 1, 1, tzinfo=timezone.utc)

def html_escape(s: str) -> str:
    return html.escape(s or "", quote=True)

# -----------------------------
# Core
# -----------------------------

def collect_updates() -> List[Dict[str, Any]]:
    cutoff = datetime.now(timezone.utc) - timedelta(days=DAYS_BACK)
    rows: List[Dict[str, Any]] = []

    for product_name, feed_url, health_url in FEEDS:
        try:
            data = http_get(feed_url)
            items = parse_rss(data)
        except Exception as e:
            # If a feed fails, continue; we still want others
            items = []

        for it in items:
            title = it["title"]
            desc = it["description"]
            link = it["link"]
            dt = try_parse_date_rss(it["pubDate"])
            if dt < cutoff:
                continue

            kbs = extract_kbs(title + " " + desc)
            # Simple, readable description: prefer title; description can be verbose HTML
            clean_desc = title
            # Fallback to description if title is empty
            if not clean_desc and desc:
                clean_desc = re.sub("<[^>]+>", " ", desc)
                clean_desc = re.sub(r"\s+", " ", clean_desc).strip()

            rows.append({
                "product": product_name,
                "description": clean_desc,
                "kbs": kbs or ["—"],
                "date": dt,
                "source": "Update Catalog RSS",
                "source_link": link,
                "health_link": health_url,
            })

    # Dedupe by (product, description, kb set)
    seen = set()
    unique = []
    for r in rows:
        key = (r["product"], r["description"], tuple(r["kbs"]))
        if key in seen:
            continue
        seen.add(key)
        unique.append(r)

    # Sort: newest first
    unique.sort(key=lambda r: r["date"], reverse=True)
    return unique

def render_html(rows: List[Dict[str, Any]]) -> str:
    count = len(rows)
    gen_ts = datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M UTC")

    def fmt_date(dt: datetime) -> str:
        return dt.strftime("%b %d, %Y")

    # Simple, self-contained styles
    css = """
    :root{color-scheme:dark}
    body{font-family:Inter,system-ui,Segoe UI,Roboto,Arial,sans-serif;margin:0;background:#0b1220;color:#e9eef7}
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
    .pill{display:inline-block;background:#15253f;border:1px solid #25447a;color:#cfe0ff;border-radius:999px;padding:4px 10px;font-size:12px}
    .src{display:inline-block;background:#0d1c33;border:1px solid #223b6e;color:#a9c5ff;border-radius:999px;padding:4px 10px;font-size:12px;text-decoration:none}
    .src:hover{background:#12274d}
    .known{display:inline-block;background:#38210a;border:1px solid #6a3c12;color:#ffd9a8;border-radius:999px;padding:4px 10px;font-size:12px;text-decoration:none}
    .known:hover{background:#4a2b0f}
    .muted{opacity:.7}
    input[type="search"]{width:360px;background:#09132a;border:1px solid #20355e;border-radius:8px;padding:8px 10px;color:#e9eef7}
    select{background:#09132a;border:1px solid #20355e;border-radius:8px;padding:8px 10px;color:#e9eef7}
    """

    # Minimal client-side filter by text
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

    # Build rows
    body_rows = []
    for r in rows:
        kb_html = " ".join(f'<span class="kb">{html_escape(kb)}</span>' for kb in r["kbs"])
        body_rows.append(f"""
        <tr>
          <td>{html_escape(r["product"])}</td>
          <td>{html_escape(r["description"])}</td>
          <td>{kb_html or "—"}</td>
          <td><a class="known" href="{html_escape(r["health_link"])}" target="_blank" rel="noopener">Release Health</a></td>
          <td>{fmt_date(r["date"])}</td>
          <td><a class="src" href="{html_escape(r["source_link"])}" target="_blank" rel="noopener">{html_escape(r["source"])}</a></td>
        </tr>
        """)

    html_out = f"""<!doctype html>
<html lang="en">
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Windows Updates (Last {DAYS_BACK} Days)</title>
<style>{css}</style>
<body>
  <header>
    <div class="wrap">
      <h1>Windows Updates (Last {DAYS_BACK} Days)</h1>
      <div class="meta">Showing {count} updates • Generated {gen_ts}</div>
      <div class="controls">
        <input id="q" type="search" placeholder="Search CVE, KB, description, OS…" oninput="filterRows()">
        <a class="btn" href="#" onclick="window.print();return false;">Print / PDF</a>
      </div>
    </div>
  </header>
  <div class="wrap">
    <table>
      <thead>
        <tr>
          <th>OS</th>
          <th>Description</th>
          <th>KB(s)</th>
          <th>Known Issues</th>
          <th>Release Date</th>
          <th>Source</th>
        </tr>
      </thead>
      <tbody>
        {''.join(body_rows) if body_rows else '<tr><td colspan="6" class="muted">No updates in the selected window.</td></tr>'}
      </tbody>
    </table>
    <div class="muted">Sources: Microsoft Update Catalog RSS; Windows Release Health landing pages.</div>
  </div>
<script>{js}</script>
</body>
</html>"""
    return html_out

def main():
    try:
        rows = collect_updates()
        html_text = render_html(rows)
        with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
            f.write(html_text)
        print(f"Wrote {len(rows)} updates to {OUTPUT_FILE}")
    except Exception as e:
        print(f"ERROR: {e}", file=sys.stderr)
        sys.exit(1)

if __name__ == "__main__":
    main()
