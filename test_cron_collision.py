#!/usr/bin/env python3
"""Guard the daily emails against the failure that hit them on 1 Sep 2026.

Two of the three daily reports (Client Follow-up and Marketing Pulse) were
scheduled on the IDENTICAL nine cron minutes. They fired simultaneously on every
slot and logged into the SAME Gmail mailbox at the same instant. On 1 Sep both
went missing while the Cash Position report — the only one that never collides —
arrived normally.

The 30 Aug commit tried to "de-congest" these schedules and de-conflicted Cash
but left the other two on top of each other, so this has already regressed once.
These tests make it fail loudly instead of silently.

Offline: reads the workflow YAML and the report scripts, touches no network.

  python3 test_cron_collision.py
"""
from __future__ import annotations

import os
import re
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
WF = os.path.join(HERE, ".github", "workflows")

SENDERS = {
    "cash":      "daily-cash-email.yml",
    "followup":  "daily-followup-email.yml",
    "marketing": "daily-marketing-email.yml",
}

fails = []


def check(label, ok, detail=""):
    print(f"  {'ok  ' if ok else 'FAIL'}  {label}{('  ' + detail) if detail else ''}")
    if not ok:
        fails.append(label)


def crons(fname):
    """Every `- cron: "M H * * *"` in a workflow, as (minute, hour) ints."""
    with open(os.path.join(WF, fname), encoding="utf-8") as fh:
        text = fh.read()
    out = []
    for m in re.finditer(r'^\s*-\s*cron:\s*["\']([^"\']+)["\']', text, re.M):
        parts = m.group(1).split()
        out.append((int(parts[0]), int(parts[1])))
    return out


def main():
    print("1. every sender still has a schedule")
    sched = {}
    for name, fname in SENDERS.items():
        c = crons(fname)
        sched[name] = c
        check(f"{name} has cron entries", len(c) > 0, f"({len(c)} slots)")
    if fails:
        print("\nFAILED early — cannot compare schedules.")
        return 1

    # ---- the actual regression: two reports firing at the same instant -------
    print("\n2. no two senders fire at the same minute (the 1 Sep failure)")
    names = sorted(sched)
    for i, a in enumerate(names):
        for b in names[i + 1:]:
            overlap = sorted(set(sched[a]) & set(sched[b]))
            check(f"{a} vs {b} never collide", not overlap,
                  f"({len(overlap)} shared slots: "
                  f"{', '.join(f'{h:02d}:{m:02d}Z' for m, h in overlap)})" if overlap else "")

    # A collision is worst back-to-back, but two sends seconds apart into one
    # mailbox is still a rate-limit risk. Require real separation.
    print("\n3. senders are at least 5 minutes apart on any shared hour")
    for i, a in enumerate(names):
        for b in names[i + 1:]:
            tight = []
            for m1, h1 in sched[a]:
                for m2, h2 in sched[b]:
                    gap = abs((h1 * 60 + m1) - (h2 * 60 + m2))
                    if 0 < gap < 5:
                        tight.append(f"{h1:02d}:{m1:02d}Z/{h2:02d}:{m2:02d}Z")
            check(f"{a} vs {b} keep >=5 min clearance", not tight,
                  f"({', '.join(tight)})" if tight else "")

    # ---- the watchdog must outlast every scheduled slot ---------------------
    print("\n4. the watchdog runs after every sender's last slot")
    wd = crons("daily-watchdog.yml")
    check("watchdog has a schedule", len(wd) == 1, f"({len(wd)} slots)")
    if wd:
        wd_min = wd[0][1] * 60 + wd[0][0]
        latest, who = -1, ""
        for name, c in sched.items():
            for m, h in c:
                if h * 60 + m > latest:
                    latest, who = h * 60 + m, name
        check("watchdog fires after the last report slot", wd_min > latest,
              f"(watchdog {wd[0][1]:02d}:{wd[0][0]:02d}Z vs latest "
              f"{latest // 60:02d}:{latest % 60:02d}Z from {who})")

    # ---- the send-gate defects fixed alongside the collision ----------------
    print("\n5. a failed send never marks the day as done")
    for script in ("daily_report.py", "cash_report.py", "followup_report.py"):
        with open(os.path.join(HERE, script), encoding="utf-8") as fh:
            src = fh.read()
        # The old marketing bug: send_email(...) on its own line, marker written
        # regardless of the result, burning the day's remaining retry slots.
        bad = re.search(r'^\s*send_email\(subject, html_body, recipients\)\s*$', src, re.M)
        check(f"{script} gates the marker on send success", not bad)
        check(f"{script} exits non-zero on a failed send",
              "return 1" in src.split("def main(")[-1])

    print("\n6. all three senders skip Friday")
    for script in ("daily_report.py", "cash_report.py", "followup_report.py"):
        with open(os.path.join(HERE, script), encoding="utf-8") as fh:
            src = fh.read()
        check(f"{script} has the Friday skip", "weekday() == 4" in src)

    print("\n7. every sender alerts on failure")
    for name, fname in SENDERS.items():
        with open(os.path.join(WF, fname), encoding="utf-8") as fh:
            text = fh.read()
        check(f"{name} has an if: failure() alert step",
              "if: failure()" in text and "alert.py" in text)
        check(f"{name} serialises its own runs", "concurrency:" in text)

    print("\n" + ("FAILED: " + "; ".join(fails) if fails else
                  "all checks passed"))
    return 1 if fails else 0


if __name__ == "__main__":
    sys.exit(main())
