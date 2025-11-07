# Windows Updates Report → GitHub Pages (Auto)

This template uses **GitHub Actions** to build and commit an HTML report of Windows updates from the last 30 days every Tuesday.

## How to use
1. Create a new repo and upload the contents of `windows-updates-report.zip`.
2. Enable **GitHub Pages** (Settings → Pages): Source: *Deploy from a branch*, Branch: `main`, Folder: `/docs`.
3. Go to **Actions** and run **Build Windows Updates HTML** once to generate the first report. Thereafter it runs weekly (Tuesday 14:00 UTC).
4. Your report will be at: `https://<you>.github.io/<repo>/windows-updates.html`

### Customize
- Add/remove OS families by editing `TARGETS` in `scripts/generate_report.py`.
- Adjust schedule in `.github/workflows/windows-updates-report.yml` (cron is UTC).
- The HTML lives in `docs/windows-updates.html`.

> The scraper is best-effort. If Microsoft changes page layouts, tweak the selectors in the script.
