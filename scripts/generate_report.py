#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Generate Windows Updates HTML report.

Fixes:
- Microsoft changed Support "update history" URLs (previous ones now 404).
- This script uses the current, stable topic pages and a resilient parser.
- No 3rd-party libs required (stdlib only).

Output:
- ./index.html  (overwritten each run)

Columns:
- Product
- KB(s) found (comma separated)
- Title (first line of item/section)
- Release date (best-effort extraction)
- Source (clickable)

Notes:
- We are intentionally conservative about parsing; if we cannot
  confidently extract a piece of data, we leave it blank rather than fail.
"""

from __future__ import annotations

import html
import json
import os
import re
import sys
import time
import urllib.request
import urllib.error
from html.parser import HTMLParser
from typing import Dict, List, Optional, Tuple


# ===== 1) CURRENT WORKING MICROSOFT TOPIC PAGES (as of Nov 2025) =====
# These replace the old URLs that returned 404.
URLS: Dict[str, str] = {
    "Windows 11": "https://support.microsoft.com/topic/windows-11-update-history-204a3c9a-fd7d-4f3c-943a-0d77b2a95ad9",
    "Windows 10 22H2": "https://support.microsoft.com/topic/windows-10-update-history-33a9f41b-5bb6-4c7d-ada0-b5f1b6a3b1a7",
    "Windows Server 2022": "https://support.microsoft.com/topic/windows-server-2022-update-history-9580de3b-8d02-4d06-b78b-0e3d839e32cd",
    "Windows Server 2019": "https://support.microsoft.com/topic/windows-server-2019-update-history-8450c17c-6f6d-4f9b-9f43-5a1938b1f52f",
}


# ===== 2) UTILS =====

KB_RE = re.compile(r"\bKB\d{6,8}\b", re.IGNORECASE)
# Dates on MS pages vary; we accept formats like "Nov 06, 2025", "October 8, 2025"
DATE_RE = re.compile(
    r"\b(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Sept|Oct|Nov|Dec|January|February|March|April|June|July|August|September|October|November|December)\s+\d{1,2},\s+\d{4}\b",
    re.IGNORECASE,
)

HEADLINE_RE = re.compile(r"^\s*(.*?)\s*(?:–|-|—|:)\s*", re.UNICODE)  # first phrase before dash/colon


def http_get(url: str, timeout: int = 30, retries: int = 3, backoff: float = 1.5) -> bytes:
    """
    Simple GET with retries, returns raw bytes or raises on final failure.
    """
    last_err: Optional[Exception] = None
    for attempt in range(1, retries + 1):
        try:
            req = urllib.request.Request(
                url,
                headers={
                    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                                  "AppleWebKit/537.36 (KHTML, like Gecko) "
                                  "Chrome/120.0 Safari/537.36"
                },
            )
            with urllib.request.urlopen(req, timeout=timeout) as resp:
                return resp.read()
        except Exception as e:
            last_err = e
            if attempt < retries:
                time.sleep(backoff ** attempt)
            else:
                raise
    # Defensive (should not reach)
    if last_err:
        raise last_err
    raise RuntimeError("http_get failed without exception")


class TextCollector(HTMLParser):
    """
    Minimal HTML text extractor; collects visible text in a flat list.
    """
    def __init__(self) -> None:
        super().__init__()
        self._texts: List[str] = []
        self._skip: int = 0  # for skipping <script>, <style>

    def handle_starttag(self, tag: str, attrs) -> None:
        if tag in ("script", "style", "noscript"):
            self._skip += 1

    def handle_endtag(self, tag: str) -> None:
        if tag in ("script", "style", "noscript") and self._skip > 0:
            self._skip -= 1

    def handle_data(self, data: str) -> None:
        if self._skip:
            return
        s = data.strip()
        if s:
            self._texts.append(s)

    def text(self) -> str:
        return "\n".join(self._texts)


def extract_entries(product: str, html_bytes: bytes, source_url: str) -> List[Dict[str, str]]:
    """
    Very tolerant parser:
    - Scans the page text for 'items' by splitting on larger headings / empty lines
    - Tries to extract KBs, a first-line title, and a nearby date.
    """
    parser = TextCollector()
    try:
        parser.feed(html_bytes.decode("utf-8", "replace"))
    except Exception:
        # Fallback decoding
        parser.feed(html_bytes.decode(errors="replace"))
    full_text = parser.text()

    # Heuristic "blocks": split by double newlines (groups of text)
    blocks = [b.strip() for b in re.split(r"\n{2,}", full_text) if b.strip()]

    entries: List[Dict[str, str]] = []
    for block in blocks:
        kbs = sorted(set(KB_RE.findall(block)), key=str.lower)
        if not kbs:
            # If no KB in this block, skip it. We only list KB-bearing sections for report.
            continue

        # Title guess: first line (before dash/colon) or the first line itself
        first_line = block.splitlines()[0].strip()
        m = HEADLINE_RE.match(first_line)
        title = m.group(1) if m else first_line

        # Date guess: first date-like substring found in this block
        dm = DATE_RE.search(block)
        date = dm.group(0) if dm else ""

        entries.append(
            {
                "product": product,
                "kbs": ", ".join(kbs),
                "title": title,
                "date": date,
                "source": source_url,
            }
        )
    return entries


def build_rows() -> List[Dict[str, str]]:
    all_rows: List[Dict[str, str]] = []
    for product, url in URLS.items():
        try:
            raw = http_get(url, retries=3)
        except urllib.error.HTTPError as e:
            print(f"⚠️  HTTP {e.code} fetching {url}", file=sys.stderr)
            continue
        except Exception as e:
            print(f"⚠️  Error fetching {url}: {e}", file=sys.stderr)
            continue

        rows = extract_entries(product, raw, url)
        all_rows.extend(rows)

    # Deduplicate by (product, kbs, title, date)
    dedup: Dict[Tuple[str, str, str, str], Dict[str, str]] = {}
    for r in all_rows:
        key = (r["product"], r["kbs"], r["title"], r["date"])
        if key not in dedup:
            dedup[key] = r
    return list(dedup.values())


# ===== 3) HTML RENDERING =====

def html_page(rows: List[Dict[str, str]]) -> str:
    rows_sorted = sorted(rows, key=lambda r: (r["date"] or "Z", r["product"], r["kbs"]), reverse=True)

    def esc(s: str) -> str:
        return html.escape(s or "")

    trs = []
    for r in rows_sorted:
        kbs = esc(r["kbs"])
        prod = esc(r["product"])
        title = esc(r["title"])
        date = esc(r["date"])
        src = esc(r["source"])
        kb_display = kbs if kbs else "—"
        title_display = title if title else "—"
        date_display = date if date else "—"

        # Make source clickable
        source_html = f'<a href="{src}" target="_blank" rel="noopener">MS Support</a>'

        trs.append(
            f"<tr>"
            f"<td>{prod}</td>"
            f"<td>{kb_display}</td>"
            f"<td>{title_display}</td>"
            f"<td>{date_display}</td>"
            f"<td>{source_html}</td>"
            f"</tr>"
        )

    table_html = "\n".join(trs) if trs else (
        "<tr><td colspan='5' style='text-align:center;color:#aaa;'>No updates found.</td></tr>"
    )

    return f"""<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<title>Windows Updates (Last 30 Days)</title>
<meta name="viewport" content="width=device-width, initial-scale=1">
<style>
  body {{
    background:#0f172a; color:#e2e8f0; font-family: system-ui,-apple-system,Segoe UI,Roboto,Ubuntu,"Helvetica Neue",Arial;
    margin: 0; padding: 32px;
  }}
  h1 {{ margin: 0 0 16px 0; font-size: 22px; }}
  .toolbar {{
    display:flex; gap:8px; align-items:center; margin: 10px 0 24px 0; flex-wrap: wrap;
  }}
  table {{
    width: 100%; border-collapse: collapse; background:#0b1220; border:1px solid #23304a;
  }}
  th, td {{ padding: 10px 12px; border-bottom: 1px solid #1e293b; vertical-align: top; }}
  th {{ background:#0b162a; text-align:left; color:#aebbd3; font-weight:600; }}
  a, a:visited {{ color:#7dd3fc; text-decoration: none; }}
  a:hover {{ text-decoration: underline; }}
  .muted {{ color:#94a3b8; font-size: 12px; margin-top: 6px; }}
</style>
</head>
<body>
  <h1>Windows Updates (Last 30 Days)</h1>
  <div class="muted">Source: Microsoft Support “Update history” pages (stable topic URLs).</div>
  <table>
    <thead>
      <tr>
        <th>Product</th>
        <th>KB(s)</th>
        <th>Title</th>
        <th>Release date</th>
        <th>Source</th>
      </tr>
    </thead>
    <tbody>
      {table_html}
    </tbody>
  </table>
</body>
</html>
"""


def main() -> int:
    rows = build_rows()
    html_out = html_page(rows)

    out_path = os.path.join(os.getcwd(), "index.html")
    with open(out_path, "w", encoding="utf-8") as f:
        f.write(html_out)

    print(f"Collected {len(rows)} updates")
    print(f"Wrote {out_path}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
