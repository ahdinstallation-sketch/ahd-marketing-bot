#!/usr/bin/env python3
"""AHD Group — Daily Lead / Client Follow-up.

A daily sales-pipeline "who's missing what" email built from the AHD sales
tracker (the same sheet the reception/sales team keep). For every open client it
works out the next missing step (needs an offer, a presentation, a signature, a
production order, a re-engagement…), groups them into action buckets sorted by
money at risk, and recommends where the team should focus today.

Source: the ahd_tracker sheet (public gviz/CSV). No auth needed.

Env vars:
  MAIL_USER / MAIL_PASSWORD   sender + app password (reused from the marketing bot)
  FOLLOWUP_RECIPIENTS         comma-separated recipient list (overrides default)

Usage:
  python3 followup_report.py --dry-run   # build followup_report.html, DO NOT send
  python3 followup_report.py --force     # build + send now regardless of the hour
  python3 followup_report.py             # build + send (guarded to the morning window)
"""
import os, sys, json, html, datetime

from sheets_pull import fetch, analyze_tracker, SHEETS, TRACKER_GID
from daily_report import cairo_now, send_email

HERE = os.path.dirname(os.path.abspath(__file__))

# NOTE: the contract → factory handover checklist is NOT handled here. It is a
# SEPARATE, instant email fired the moment a client is ticked CONTRACTED, run by a
# Google Apps Script bound to the tracker sheet (see handover_apps_script.gs). This
# follow-up email keeps only the daily pipeline + the "clients this month" list.

DEFAULT_RECIPIENTS = [
    "ahmed.helmy@amrhelmydesigns.com",     # Ahmed (me)
    "ahdh@amrhelmydesigns.com",            # Ezz (AHDH)
    "helmymalak@gmail.com",                # Malak Helmy
    "orders@amrhelmydesigns.com",          # Orders
    "crm@amrhelmydesigns.com",             # CRM
    "mohamed.fahmy@amrhelmydesigns.com",   # Mohamed Fahmy
]

SEND_HOUR_CAIRO = 9
# Widened 21→23: GitHub cron can fire hours late, so accept any run 09:00–23:00 Cairo
# (dedupe keeps it once/day) rather than dropping a late fire and missing the day.
SEND_WINDOW_END_CAIRO = 23
SENT_MARKER = os.environ.get("FOLLOWUP_SENT_MARKER") or os.path.join(HERE, ".followup_sent_marker")

STALE_DAYS = 90    # a deal untouched longer than this is flagged as ageing
RECENT_DAYS = 120  # "recent" cutoff for the focus list
HIGH_TICKET_EGP = float(os.environ.get("HIGH_TICKET_EGP", "2000000") or 2000000)  # "chase first" cutoff
MID_TICKET_EGP = float(os.environ.get("MID_TICKET_EGP", "1000000") or 1000000)    # second-tier "also chase" floor (1M–HIGH band)


# ---------- formatting ----------

def esc(s):
    return html.escape(str(s if s is not None else ""))


def fmt(n):
    try:
        return f"{round(float(n)):,}"
    except (TypeError, ValueError):
        return "0"


def fmt_m(n):
    try:
        return f"{float(n) / 1e6:.1f}M"
    except (TypeError, ValueError):
        return "0"


def _clean_name(s):
    return (s or "").replace("\xa0", " ").strip()


# ---------- bucketing: what is each open client missing? ----------
# Ordered by sales priority. The first matching rule wins. "Won" stages (order in
# production) are summarised, not chased. Over-budget / no-answer clients are
# intentionally excluded — this email is the forward-moving follow-up list.
_EXCLUDE_STATUS = {"OVER BUDGET", "NO ANSWER AFTER OFFER"}


def _bucket(rec):
    """Return (key, label, why) for a client, or None if won / excluded."""
    stage = (rec.get("stage") or "").upper()
    status = (rec.get("status") or "").upper()

    if status in _EXCLUDE_STATUS:
        return None  # over budget / silent — handled separately, not here
    if stage == "ORDER":
        return None  # in production — done
    if stage == "CONTRACTED" or status == "SIGNED CONTRACT":
        return None  # signed / awaiting production order — closed for sales follow-up
    if stage == "FINAL PRESENTATION AFTER CLIENT COMMENTS":
        return ("presented_no_contract", "Presented — awaiting signature",
                "Full presentation done, revisions addressed. This is the closing step.")
    if stage in ("INITIAL PRESENTATION", "OFFER"):
        return ("needs_presentation", "Needs a presentation",
                "Quote / initial design is out — book the design presentation.")
    if stage == "SESSION":
        return ("needs_offer", "Needs an offer / quote",
                "Session done — send the pricing/offer.")
    return ("new", "New — needs a first session",
            "No stage logged yet. Book the discovery/measurement session.")


# Display order + short titles for the action sections.
_BUCKET_ORDER = [
    ("presented_no_contract", "🖊️ Presented — awaiting signature"),
    ("needs_presentation", "📐 Needs a presentation"),
    ("needs_offer", "🧾 Needs an offer / quote"),
    ("new", "🆕 New — needs first session"),
]


def build(t):
    """Turn analyze_tracker output into follow-up buckets + recommendations."""
    buckets = {k: [] for k, _ in _BUCKET_ORDER}
    won_orders = won_contracted = 0
    for rec in t.get("clients_detail", []):
        stage = (rec.get("stage") or "").upper()
        if stage == "ORDER":
            won_orders += 1
            continue
        b = _bucket(rec)
        if not b:
            continue
        key, label, why = b
        buckets.setdefault(key, []).append({
            "client": _clean_name(rec.get("client")),
            "rep": _clean_name(rec.get("rep")),
            "amount": rec.get("amount_num") or 0,
            "usd": bool(rec.get("amount_usd_converted")),
            "days_ago": rec.get("days_ago"),
            "note": _clean_name(rec.get("note")),
            "why": why, "label": label,
        })
    for k in buckets:
        buckets[k].sort(key=lambda d: -(d["amount"] or 0))

    # Money at risk per bucket (the recommendation engine ranks on this).
    ranked = []
    for key, title in _BUCKET_ORDER:
        rows = buckets.get(key, [])
        if not rows:
            continue
        total = sum(r["amount"] or 0 for r in rows)
        ranked.append({"key": key, "title": title, "rows": rows,
                       "count": len(rows), "value": total})
    ranked_by_value = sorted(ranked, key=lambda b: -b["value"])

    # --- Focus list: customer-oriented, high-ticket AND recent ---
    everyone = [c for b in ranked for c in b["rows"] if (c["amount"] or 0) > 0]
    recent = [c for c in everyone
              if c["days_ago"] is not None and c["days_ago"] <= RECENT_DAYS]
    pool = recent if recent else everyone
    pool = sorted(pool, key=lambda c: -(c["amount"] or 0))
    focus = pool[:5]

    # --- High-ticket "chase first": every open deal at/above the threshold, biggest first.
    high_ticket = sorted(
        [c for c in everyone if (c["amount"] or 0) >= HIGH_TICKET_EGP],
        key=lambda c: -(c["amount"] or 0))
    # --- Mid-ticket "also chase": the 1M–HIGH band (no overlap with high_ticket).
    mid_ticket = sorted(
        [c for c in everyone if MID_TICKET_EGP <= (c["amount"] or 0) < HIGH_TICKET_EGP],
        key=lambda c: -(c["amount"] or 0))
    # --- The chase list: EVERY open deal ≥ MID_TICKET_EGP (1M and up), biggest first —
    # one ranked table shown in full on a single page.
    chase = sorted(
        [c for c in everyone if (c["amount"] or 0) >= MID_TICKET_EGP],
        key=lambda c: -(c["amount"] or 0))

    # --- Recent clients with NO deal value logged (can't be ranked → easy to miss).
    missing_amount = [c for b in ranked for c in b["rows"]
                      if (c["amount"] or 0) == 0
                      and c["days_ago"] is not None and c["days_ago"] <= RECENT_DAYS]
    missing_amount.sort(key=lambda c: c["days_ago"])

    return {
        "buckets": ranked, "ranked_by_value": ranked_by_value,
        "focus": focus, "focus_recent": bool(recent),
        "high_ticket": high_ticket,
        "high_ticket_value": sum(c["amount"] or 0 for c in high_ticket),
        "mid_ticket": mid_ticket,
        "mid_ticket_value": sum(c["amount"] or 0 for c in mid_ticket),
        "chase": chase,
        "chase_value": sum(c["amount"] or 0 for c in chase),
        "missing_amount": missing_amount,
        "won_orders": won_orders,
        "open_count": sum(b["count"] for b in ranked),
        "open_value": sum(b["value"] for b in ranked),
    }


# ---------- this-month builder ----------

def this_month(t):
    """Clients whose latest tracked activity falls in the current Cairo month —
    a quick 'who came in this month' list at the top of the email."""
    now = cairo_now()
    ym = (now.year, now.month)
    all_dated = []
    for rec in t.get("clients_detail", []):
        ds = rec.get("date")
        if not ds:
            continue
        try:
            d = datetime.date.fromisoformat(ds)
        except (TypeError, ValueError):
            continue
        all_dated.append({
            "client": _clean_name(rec.get("client")),
            "rep": _clean_name(rec.get("rep")),
            "amount": rec.get("amount_num") or 0,
            "usd": bool(rec.get("amount_usd_converted")),
            "stage": (rec.get("stage") or "").title() or "—",
            "status": _clean_name(rec.get("status")),
            "date": d,
        })
    all_dated.sort(key=lambda r: (r["date"], r["amount"] or 0), reverse=True)
    rows = [r for r in all_dated if (r["date"].year, r["date"].month) == ym]
    # The team's date columns lag, so the current month can legitimately be empty
    # early on. Rather than show a dead section, fall back to the 6 most recently
    # dated clients (clearly flagged) so there's always something actionable.
    fallback = not rows
    shown = rows if rows else all_dated[:6]
    return {"rows": shown, "value": sum(r["amount"] or 0 for r in shown),
            "label": f"{now:%B %Y}", "fallback": fallback,
            "month_count": len(rows)}


# ---------- render ----------

CSS = """
@page{size:A4;margin:8mm}
body{margin:0;background:#f4f1ea;font-family:-apple-system,Segoe UI,Roboto,Helvetica,Arial,sans-serif;color:#2b2b2b}
.wrap{max-width:700px;margin:0 auto;padding:10px 16px 8px}
.head{text-align:center;padding:2px 0 8px}
.head h1{margin:0;font-size:19px;letter-spacing:.5px;color:#1c1c1c}
.head .sub{font-size:12px;color:#8a7a52;margin-top:3px}
.card{background:#fff;border:1px solid #e7e0cf;border-radius:10px;padding:8px 14px;margin:5px 0;box-shadow:0 1px 2px rgba(0,0,0,.03)}
.kpis{display:flex;gap:8px;flex-wrap:wrap}
.kpi{flex:1;min-width:120px;background:#fbf9f3;border:1px solid #ece4d2;border-radius:9px;padding:9px 12px}
.kpi .lbl{font-size:10px;text-transform:uppercase;letter-spacing:.5px;color:#9a8c66}
.kpi .val{font-size:20px;font-weight:700;margin-top:2px}
.kpi .sm{font-size:11px;color:#8a8a8a;margin-top:1px}
h3{margin:2px 0 7px;font-size:14px;color:#1c1c1c}
.bkt{border:1px solid #ece4d2;border-radius:9px;margin:8px 0;overflow:hidden}
.bkt .bh{background:#fbf7ee;padding:7px 12px;display:flex;justify-content:space-between;align-items:center;border-bottom:1px solid #ece4d2}
.bkt .bt{font-size:12.5px;font-weight:700}
.bkt .bv{font-size:11.5px;color:#8a7a52}
.bkt .why{font-size:11px;color:#8a8a8a;padding:5px 12px 0}
table{width:100%;border-collapse:collapse;font-size:12px}
th,td{padding:4px 12px;text-align:right;border-bottom:1px solid #f2ede2}
th:first-child,td:first-child{text-align:left}
thead th{font-size:9.5px;text-transform:uppercase;letter-spacing:.4px;color:#9a8c66}
tr:last-child td{border-bottom:none}
table.chase{font-size:10px}
table.chase th,table.chase td{padding:0.5px 10px}
table.chase thead th{font-size:8.5px}
.stale{color:#b06a12;font-size:10px;font-weight:700}
.usd{color:#2f6f8f;font-size:9.5px;font-weight:700;background:#eaf2f6;border:1px solid #cfe0e8;border-radius:4px;padding:0 4px;margin-left:4px}
.rec{background:#fbf7ee;border:1px solid #ece4d2;border-left:3px solid #c8a24a;border-radius:7px;padding:6px 11px;margin:5px 0;font-size:12.5px}
.rec b{color:#8a6a12}
.note{font-size:11px;color:#9a9a9a;margin-top:6px}
.more{font-size:11px;color:#9a8c66;padding:5px 12px}
.chips{font-size:12px;color:#555;margin-top:8px}
.chips b{color:#1c1c1c}
.foot{text-align:center;font-size:10px;color:#a9a190;margin-top:8px}
.hv-hint{font-size:11px;color:#8a8a8a;margin:0 0 6px}
.hv-hint a{color:#8a6a12;font-weight:600;text-decoration:none}
"""

def _days_tag(da):
    if da is None:
        return ""
    if da > STALE_DAYS:
        return f' <span class="stale">{da}d</span>'
    return f' <span style="color:#9a9a9a;font-size:10.5px">{da}d</span>'


def _usd_tag(is_usd):
    return ' <span class="usd">$→EGP</span>' if is_usd else ""


def _chase_all(b):
    """Every open deal ≥ MID_TICKET_EGP (1M and up), biggest first, shown in full on one
    page. Deals ≥ HIGH_TICKET_EGP get a 💎 marker so the top tier still stands out."""
    rows = b.get("chase", [])
    if not rows:
        return ('<div class="note">No open deals at/above '
                f'{fmt(MID_TICKET_EGP)} EGP right now.</div>')
    body = ""
    for i, c in enumerate(rows, 1):
        amt = fmt(c["amount"]) if c["amount"] else "—"
        gem = "💎 " if (c["amount"] or 0) >= HIGH_TICKET_EGP else ""
        body += (f'<tr><td><b>{i}.</b> {gem}{esc(c["client"]) or "—"}{_days_tag(c["days_ago"])}</td>'
                 f'<td style="color:#8a7a52">{esc(c["label"])}</td>'
                 f'<td><b>{amt}</b>{_usd_tag(c.get("usd"))}</td></tr>')
    return (f'<table class="chase"><thead><tr><th>Client</th><th>Next step</th>'
            f'<th>Amount (EGP)</th></tr></thead><tbody>{body}</tbody></table>')


def _missing(b):
    """Recent clients with no amount entered — flag them so a value gets added and
    they stop falling out of the ranked focus list."""
    rows = b.get("missing_amount", [])
    if not rows:
        return '<div class="note">All recent clients have a value logged. ✅</div>'
    items = []
    for c in rows[:12]:
        da = c.get("days_ago")
        rec = ("dated ahead" if (da is not None and da < 0)
               else (f"{da}d ago" if da is not None else "no date"))
        lab = f' · {esc(c["label"])}' if c.get("label") else ""
        items.append(f'<b>{esc(c["client"]) or "—"}</b>'
                     f'<span style="color:#9a9a9a;font-size:11px"> ({rec}{lab})</span>')
    more = f' &nbsp;·&nbsp; +{len(rows) - 12} more' if len(rows) > 12 else ""
    return f'<div class="chips">{" &nbsp;·&nbsp; ".join(items)}{more}</div>'


def _month_list(m):
    """Compact list of clients seen this month, biggest/most-recent first."""
    rows = m.get("rows", [])
    if not rows:
        return '<div class="note">No clients logged with a date in this month yet.</div>'
    note = ""
    if m.get("fallback"):
        note = (f'<div class="hv-hint">No clients dated in {esc(m.get("label",""))} yet '
                f'(the tracker\'s date columns lag) — showing the most recent instead.</div>')
    body = ""
    for i, c in enumerate(rows, 1):
        amt = fmt(c["amount"]) if c["amount"] else "—"
        extra = f' · {esc(c["status"])}' if c.get("status") else ""
        body += (f'<tr><td><b>{i}.</b> {esc(c["client"]) or "—"}'
                 f'<span style="color:#9a9a9a;font-size:10px"> · {esc(c["date"].strftime("%d %b"))}</span></td>'
                 f'<td style="color:#8a7a52">{esc(c["stage"])}{extra}</td>'
                 f'<td><b>{amt}</b>{_usd_tag(c.get("usd"))}</td></tr>')
    return (note + f'<table class="chase"><thead><tr><th>Client</th><th>Stage</th>'
            f'<th>Amount (EGP)</th></tr></thead><tbody>{body}</tbody></table>')


def render(t, b, m=None):
    today = cairo_now()
    if not t or not t.get("clients_detail"):
        return "<html><body><p>Follow-up unavailable — could not read the sales tracker.</p></body></html>"

    kpis = f"""
    <div class="kpis">
      <div class="kpi"><div class="lbl">Open — need follow-up</div>
        <div class="val" style="color:#b06a12">{b['open_count']}</div>
        <div class="sm">clients with a next step due</div></div>
      <div class="kpi"><div class="lbl">Pipeline at risk</div>
        <div class="val">{fmt_m(b['open_value'])}</div><div class="sm">EGP across open clients</div></div>
      <div class="kpi"><div class="lbl">Signed value</div>
        <div class="val" style="color:#1f7a44">{fmt_m(t.get('signed_value',0))}</div>
        <div class="sm">EGP contracted</div></div>
      <div class="kpi"><div class="lbl">In production</div>
        <div class="val">{b['won_orders']}</div><div class="sm">orders placed</div></div>
    </div>"""

    return f"""<!doctype html><html><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1"><style>{CSS}</style></head>
<body><div class="wrap">
  <div class="head">
    <h1>AHD — Daily Client Follow-up</h1>
    <div class="sub">{esc(f'{today:%A, %d %B %Y}')} · live from the sales tracker</div>
  </div>

  <div class="card">
    <h3 style="margin-top:0">Where the pipeline stands</h3>
    {kpis}
  </div>

  <div class="card">
    <h3 style="margin-top:0">🗓️ Clients this month <span style="font-size:10.5px;color:#9a9a9a;font-weight:400">— {esc((m or {}).get('label',''))} · {(m or {}).get('month_count',0)} clients · {fmt((m or {}).get('value',0))} EGP</span></h3>
    {_month_list(m or {})}
  </div>

  <div class="card">
    <h3 style="margin-top:0">💎 Clients to chase — 1M and up <span style="font-size:10.5px;color:#9a9a9a;font-weight:400">— open deals ≥ {fmt(MID_TICKET_EGP)} EGP · {len(b['chase'])} clients · {fmt(b['chase_value'])} EGP · 💎 = ≥ {fmt(HIGH_TICKET_EGP)}</span></h3>
    {_chase_all(b)}
  </div>

  <div class="card">
    <h3 style="margin-top:0">⚠️ Clients missing a value <span style="font-size:10.5px;color:#9a9a9a;font-weight:400">— add their amount in the tracker so they rank above</span></h3>
    {_missing(b)}
  </div>

  <div class="foot">Auto-generated daily from the live tracker · LOST deals excluded · <span class="usd">$→EGP</span> converted at {esc(f"{t.get('usd_egp_rate', 0):.2f}")} EGP/USD · update stage/status in the tracker to change this.</div>
</div></body></html>"""


# ---------- send guard ----------

def _already_sent_today():
    try:
        with open(SENT_MARKER, encoding="utf-8") as fh:
            return fh.read().strip() == cairo_now().date().isoformat()
    except OSError:
        return False


def _mark_sent_today():
    try:
        with open(SENT_MARKER, "w", encoding="utf-8") as fh:
            fh.write(cairo_now().date().isoformat())
    except OSError:
        pass


def main():
    dry = "--dry-run" in sys.argv
    force = "--force" in sys.argv
    recipients = [r.strip() for r in os.environ.get("FOLLOWUP_RECIPIENTS", "").split(",") if r.strip()] \
        or DEFAULT_RECIPIENTS

    print("Pulling tracker…")
    t = analyze_tracker(fetch(SHEETS["ahd_tracker"], TRACKER_GID))
    b = build(t)
    m = this_month(t)
    html_body = render(t, b, m)
    out = os.path.join(HERE, "followup_report.html")
    with open(out, "w", encoding="utf-8") as fh:
        fh.write(html_body)
    print(f"Wrote {out}")
    print(json.dumps({"open_count": b["open_count"], "open_value": round(b["open_value"]),
                      "won_orders": b["won_orders"],
                      "this_month": m["month_count"],
                      "buckets": [(x["key"], x["count"], round(x["value"])) for x in b["buckets"]]},
                     ensure_ascii=False, indent=1))

    subject = f"AHD Client Follow-up — {cairo_now():%a %d %b}"
    if dry:
        print("DRY RUN — not sending.")
        return 0
    if not force:
        if cairo_now().weekday() == 4:  # Friday — weekend, no send
            print("Cairo day is Friday — skipping.")
            return 0
        if _already_sent_today():
            print("Already sent today — skipping (dedupe).")
            return 0
        hour = cairo_now().hour
        if not (SEND_HOUR_CAIRO <= hour < SEND_WINDOW_END_CAIRO):
            print(f"Cairo hour is {hour}, outside send window — skipping.")
            return 0
    # Exit non-zero on a failed send so the GitHub job goes red and the
    # failure-alert step fires. Silently returning 0 made a failed send look
    # exactly like a successful one.
    if not send_email(subject, html_body, recipients):
        print("Send failed — NOT marking today as sent, so a later run can retry.")
        return 1
    _mark_sent_today()
    return 0


if __name__ == "__main__":
    sys.exit(main())
