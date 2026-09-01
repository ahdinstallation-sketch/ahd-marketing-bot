#!/usr/bin/env python3
"""Last line of defence: catch a daily report that never went out at all.

Why this exists
---------------
Each workflow has an `if: failure()` alert step, but that can only fire if the
job actually starts. GitHub's scheduled crons are explicitly best-effort — they
fire late and can be dropped entirely — so "the workflow never ran" is a real
failure mode that is invisible from inside the workflow. Nothing else in this
repo can detect it.

This runs late in the Cairo day, after every scheduled slot for all three
reports has passed. For each report it checks the once-per-day send marker; if
one is missing on a working day, it sends that report (recovery) and tells
Ahmed which one had to be rescued.

Safe to run repeatedly: a report whose marker is present is left alone, and by
the time this fires no scheduled slot remains, so it cannot race a normal run.

  python3 watchdog.py             # check, recover what's missing, alert
  python3 watchdog.py --dry-run   # report what it WOULD do, send nothing
"""
import os
import subprocess
import sys
import html

from daily_report import cairo_now, send_email
from alert import DEFAULT_ALERT_RECIPIENTS

HERE = os.path.dirname(os.path.abspath(__file__))

# (label, script, marker filename) — marker names match the workflows' cache paths.
REPORTS = [
    ("AHD Cash Position",   "cash_report.py",     ".cash_sent_marker"),
    ("AHD Client Follow-up", "followup_report.py", ".followup_sent_marker"),
    ("AHD Marketing Pulse", "daily_report.py",    ".sent_marker"),
]


def _marker_is_today(fname):
    try:
        with open(os.path.join(HERE, fname), encoding="utf-8") as fh:
            return fh.read().strip() == cairo_now().date().isoformat()
    except OSError:
        return False


def _summary_body(rescued, failed):
    def rows(items, colour):
        return "".join(
            f'<li style="margin:4px 0;color:{colour}">{html.escape(x)}</li>' for x in items)

    parts = []
    if rescued:
        parts.append('<p style="font-size:14px;margin:0 0 6px"><b>Sent late by the watchdog:</b></p>'
                     f'<ul style="margin:0 0 16px;padding-left:20px">{rows(rescued, "#1e7a3c")}</ul>')
    if failed:
        parts.append('<p style="font-size:14px;margin:0 0 6px"><b>Still not sent — needs a look:</b></p>'
                     f'<ul style="margin:0 0 16px;padding-left:20px">{rows(failed, "#c0392b")}</ul>')
    return (
        '<div style="font-family:Arial,Helvetica,sans-serif;background:#faf8f4;'
        'padding:24px;color:#2b2b2b">'
        '<div style="max-width:560px;margin:0 auto;background:#fff;'
        'border:1px solid #e6ded0;border-left:5px solid #e8b84b;border-radius:8px;'
        'overflow:hidden">'
        '<div style="background:#1c1c1c;color:#e8b84b;padding:16px 20px;'
        'font-size:17px;font-weight:bold">Watchdog: a daily email was missing</div>'
        f'<div style="padding:20px"><p style="font-size:15px;margin:0 0 16px">'
        f'As of {cairo_now():%a %d %b %Y %H:%M} Cairo, these reports had not gone out '
        f'through their normal schedule.</p>{"".join(parts)}'
        '<p style="font-size:12px;color:#7a7a7a;margin:8px 0 0">'
        'Most likely GitHub dropped or badly delayed the scheduled run. If this '
        'repeats, the schedule needs to move off GitHub Actions.</p>'
        '</div></div></div>'
    )


def main():
    dry = "--dry-run" in sys.argv
    now = cairo_now()

    if now.weekday() == 4:  # Friday — none of the reports send, nothing to rescue
        print("Cairo day is Friday — no reports are due. Nothing to check.")
        return 0

    rescued, failed, ok = [], [], []
    for label, script, marker in REPORTS:
        if _marker_is_today(marker):
            ok.append(label)
            print(f"ok       {label} — already sent today.")
            continue
        print(f"MISSING  {label} — no send marker for {now.date()}.")
        if dry:
            failed.append(f"{label} (dry run — not sent)")
            continue
        # --force bypasses the window/dedupe: by now the window has closed, but a
        # missing report is worth more late than never.
        proc = subprocess.run([sys.executable, os.path.join(HERE, script), "--force"],
                              capture_output=True, text=True)
        sys.stdout.write(proc.stdout)
        sys.stderr.write(proc.stderr)
        if proc.returncode == 0:
            rescued.append(label)
            print(f"  -> recovered: {label} sent late.")
        else:
            tail = (proc.stderr or proc.stdout or "").strip().splitlines()
            failed.append(f"{label} — {tail[-1] if tail else f'exit {proc.returncode}'}")
            print(f"  -> STILL FAILED: {label}")

    print(f"\n{len(ok)} sent normally, {len(rescued)} rescued, {len(failed)} still failing.")

    if not rescued and not failed:
        print("All three went out on schedule — no alert needed.")
        return 0
    if dry:
        print("DRY RUN — not alerting.")
        return 0

    recipients = [r.strip() for r in os.environ.get("ALERT_RECIPIENTS", "").split(",")
                  if r.strip()] or DEFAULT_ALERT_RECIPIENTS
    send_email(f"[WATCHDOG] Missing daily email — {now:%a %d %b}",
               _summary_body(rescued, failed), recipients)
    # Go red only if something is still not sent; a successful rescue is a win.
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main())
