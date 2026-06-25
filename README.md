# AHD Group — Daily Marketing Pulse

Automated daily email (09:00 Cairo) to the marketing team. Pulls **Meta ads** data
(spend yesterday + month-to-date, ad winners/losers, top organic posts to boost) via
the Graph API and merges it with the **lead + sales pipeline** from the 3 Google Sheets,
then emails an action-oriented HTML report. Runs in the cloud on **GitHub Actions** —
no laptop or app needs to be open.

## Files
- `meta_pull.py` — Meta Graph API: per-account spend/leads/CPL/CTR (yesterday + MTD),
  ad-level winners/losers (7d), top organic posts per page (14d).
- `sheets_pull.py` — the 3 Google Sheets (funnel, pipeline value, over-budget /
  no-answer / contracted-no-order lists, lead dispositions) **plus the live "cash in
  to date" per company** from the Looker-Studio-backed treasury sheet.
- `daily_report.py` — merges both, computes marketing→sales linkage, renders the HTML
  email, sends via Gmail SMTP. Guarded to send only at 09:00 Cairo (DST-safe).
- `.github/workflows/daily-marketing-email.yml` — the daily cloud schedule.

## Live "cash in by company"
The report shows **cash in by company** (Total / AHD / Designy, latest month + YTD) pulled
live from the treasury sheet's **"Cash In by Company"** tab — so it auto-updates. Point the
job at that sheet with env vars / repo secrets:
- `CASHIN_SHEET_ID` — the sheet ID (must be readable: File → Share → "Anyone with the link –
  Viewer", or Publish to web).
- `CASHIN_TAB` — (optional) the tab name to read; defaults to `Cash In by Company`. Read by
  name via the gviz endpoint, so it survives gid changes.
- `CASHIN_GID` — (optional) read a specific tab gid instead of by name.

Expected layout: a monthly table with a header row containing `AHD`, `Designy`, … and
`Total EGP`, one row per month (a `Year` + `Month` column). The latest row with a non-zero
`Total EGP` is treated as the most recent actuals; YTD sums all rows of that year. The
section is hidden if `CASHIN_SHEET_ID` is unset.

## Secrets (GitHub repo → Settings → Secrets and variables → Actions)
| Secret | What |
|---|---|
| `META_TOKEN` | Meta **System User** token (non-expiring) with `ads_read`, `read_insights`, `pages_read_engagement`, `pages_read_user_content`, `business_management`. Assign all 4 ad accounts + 3 pages. |
| `MAIL_USER` | Sender address (e.g. `projects@designy-egypt.com`). Host is auto-picked from the domain. |
| `MAIL_PASSWORD` | App password for that mailbox (Outlook/Microsoft 365 or Gmail). |
| `SMTP_HOST` | (optional) override SMTP host, e.g. `smtp.office365.com`. Inferred from the address if unset. |
| `SMTP_PORT` | (optional) override SMTP port (`587` STARTTLS or `465` SSL). |
| `RECIPIENTS` | (optional) comma-separated override of the recipient list. |
| `CASHIN_SHEET_ID` | (optional) Google Sheet ID feeding the Looker Studio treasury dashboard, for the live "cash in to date" line. Must be CSV-published. |
| `CASHIN_GID` | (optional) specific tab gid within `CASHIN_SHEET_ID`. |

## Run locally
```bash
export META_TOKEN=...            # optional; sheets still work without it
python3 daily_report.py --dry-run    # builds report.html, prints summary, no send
# add MAIL_USER + MAIL_PASSWORD (Outlook/365 or Gmail) then:
python3 daily_report.py --force      # build + send now
```

## Schedule / DST
GitHub cron is UTC. The workflow fires at 06:07 and 07:07 UTC; the script only *sends*
when the Cairo local hour is 09, so it stays at 9 AM Cairo across DST changes.
