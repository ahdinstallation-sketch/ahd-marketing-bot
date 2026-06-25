# AHD Group — Daily Marketing Pulse

Automated daily email (09:00 Cairo) to the marketing team. Pulls **Meta ads** data
(spend yesterday + month-to-date, ad winners/losers, top organic posts to boost) via
the Graph API and merges it with the **lead + sales pipeline** from the 3 Google Sheets,
then emails an action-oriented HTML report. Runs in the cloud on **GitHub Actions** —
no laptop or app needs to be open.

## Files
- `meta_pull.py` — Meta Graph API: per-account spend/leads/CPL/CTR (yesterday + MTD),
  ad-level winners/losers (7d), top organic posts per page (14d).
- `sheets_pull.py` — the 3 Google Sheets: funnel, pipeline value, over-budget /
  no-answer / contracted-no-order lists, lead dispositions.
- `daily_report.py` — merges both, computes marketing→sales linkage, renders the HTML
  email, sends via Gmail SMTP. Guarded to send only at 09:00 Cairo (DST-safe).
- `.github/workflows/daily-marketing-email.yml` — the daily cloud schedule.

## Secrets (GitHub repo → Settings → Secrets and variables → Actions)
| Secret | What |
|---|---|
| `META_TOKEN` | Meta **System User** token (non-expiring) with `ads_read`, `read_insights`, `pages_read_engagement`, `pages_read_user_content`, `business_management`. Assign all 4 ad accounts + 3 pages. |
| `GMAIL_USER` | Sender Gmail/Workspace address. |
| `GMAIL_APP_PASSWORD` | 16-char Gmail App Password (account needs 2FA on). |
| `RECIPIENTS` | (optional) comma-separated override of the recipient list. |

## Run locally
```bash
export META_TOKEN=...            # optional; sheets still work without it
python3 daily_report.py --dry-run    # builds report.html, prints summary, no send
# add GMAIL_USER + GMAIL_APP_PASSWORD then:
python3 daily_report.py --force      # build + send now
```

## Schedule / DST
GitHub cron is UTC. The workflow fires at 06:07 and 07:07 UTC; the script only *sends*
when the Cairo local hour is 09, so it stays at 9 AM Cairo across DST changes.
