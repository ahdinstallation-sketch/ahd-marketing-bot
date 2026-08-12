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
import os, sys, json, html

from sheets_pull import fetch, analyze_tracker, SHEETS, TRACKER_GID
from daily_report import cairo_now, send_email

HERE = os.path.dirname(os.path.abspath(__file__))

DEFAULT_RECIPIENTS = [
    "ahmed.helmy@amrhelmydesigns.com",     # Ahmed (me)
    "ahdh@amrhelmydesigns.com",            # Ezz (AHDH)
    "helmymalak@gmail.com",                # Malak Helmy
    "orders@amrhelmydesigns.com",          # Orders
    "crm@amrhelmydesigns.com",             # CRM
]

SEND_HOUR_CAIRO = 8
SEND_WINDOW_END_CAIRO = 21
SENT_MARKER = os.environ.get("FOLLOWUP_SENT_MARKER") or os.path.join(HERE, ".followup_sent_marker")

STALE_DAYS = 90    # a deal untouched longer than this is flagged as ageing
RECENT_DAYS = 120  # "recent" cutoff for the focus list


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
        return ("contracted_no_order", "Contracted — no production order yet",
                "Signed but not yet released to the factory. Collect deposit & place the order.")
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
    ("contracted_no_order", "🏭 Contracted — no production order"),
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

    return {
        "buckets": ranked, "ranked_by_value": ranked_by_value,
        "focus": focus, "focus_recent": bool(recent),
        "won_orders": won_orders,
        "open_count": sum(b["count"] for b in ranked),
        "open_value": sum(b["value"] for b in ranked),
    }


# ---------- render ----------

CSS = """
@page{size:A4;margin:8mm}
body{margin:0;background:#f4f1ea;font-family:-apple-system,Segoe UI,Roboto,Helvetica,Arial,sans-serif;color:#2b2b2b}
.wrap{max-width:700px;margin:0 auto;padding:14px 16px 20px}
.head{text-align:center;padding:2px 0 8px}
.head h1{margin:0;font-size:19px;letter-spacing:.5px;color:#1c1c1c}
.head .sub{font-size:12px;color:#8a7a52;margin-top:3px}
.card{background:#fff;border:1px solid #e7e0cf;border-radius:10px;padding:11px 14px;margin:9px 0;box-shadow:0 1px 2px rgba(0,0,0,.03)}
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
.stale{color:#b06a12;font-size:10px;font-weight:700}
.usd{color:#2f6f8f;font-size:9.5px;font-weight:700;background:#eaf2f6;border:1px solid #cfe0e8;border-radius:4px;padding:0 4px;margin-left:4px}
.rec{background:#fbf7ee;border:1px solid #ece4d2;border-left:3px solid #c8a24a;border-radius:7px;padding:6px 11px;margin:5px 0;font-size:12.5px}
.rec b{color:#8a6a12}
.note{font-size:11px;color:#9a9a9a;margin-top:6px}
.more{font-size:11px;color:#9a8c66;padding:5px 12px}
.chips{font-size:12px;color:#555;margin-top:8px}
.chips b{color:#1c1c1c}
.foot{text-align:center;font-size:10.5px;color:#a9a190;margin-top:14px}
"""

_MAXROWS = 4           # top N clients per money bucket
_PRIMARY = {"presented_no_contract", "contracted_no_order", "needs_presentation"}


def _days_tag(da):
    if da is None:
        return ""
    if da > STALE_DAYS:
        return f' <span class="stale">{da}d</span>'
    return f' <span style="color:#9a9a9a;font-size:10.5px">{da}d</span>'


def _usd_tag(is_usd):
    return ' <span class="usd">$→EGP</span>' if is_usd else ""


def _client_table(rows):
    body = ""
    for r in rows[:_MAXROWS]:
        amt = fmt(r["amount"]) if r["amount"] else "—"
        body += (f"<tr><td>{esc(r['client']) or '—'}{_days_tag(r['days_ago'])}</td>"
                 f"<td>{amt}{_usd_tag(r.get('usd'))}</td></tr>")
    return (f'<table><thead><tr><th>Client</th><th>Amount (EGP)</th></tr></thead>'
            f'<tbody>{body}</tbody></table>')


def _funnel(t):
    f = t.get("funnel", {})
    stages = [("Session", "SESSION"), ("Offer / quote", "OFFER"),
              ("Initial presentation", "INITIAL PRESENTATION"),
              ("Final presentation", "FINAL PRESENTATION AFTER CLIENT COMMENTS"),
              ("Contracted", "CONTRACTED"), ("Order (production)", "ORDER")]
    top = f.get("SESSION", 0) or 1
    rows = ""
    prev = None
    for label, key in stages:
        n = f.get(key, 0)
        w = max(2, round(100 * n / top))
        drop = ""
        if prev is not None and prev > 0:
            lost = prev - n
            if lost > 0:
                drop = f' <span style="color:#b3261e;font-size:10.5px">−{lost}</span>'
        rows += (f'<div class="frow"><div class="cap">{esc(label)}</div>'
                 f'<div class="track"><div class="fill" style="width:{w}%"></div></div>'
                 f'<div class="n">{n}{drop}</div></div>')
        prev = n
    return f'<div class="funnel">{rows}</div>'


def _focus(b):
    """Named, customer-oriented focus list: the biggest recent tickets to chase."""
    rows = ""
    for i, c in enumerate(b["focus"], 1):
        amt = fmt(c["amount"]) if c["amount"] else "—"
        da = c.get("days_ago")
        recency = f'{da}d ago' if da is not None else 'no date'
        rows += (
            f'<div class="rec"><b>{i}. {esc(c["client"]) or "—"}</b> · '
            f'<b>{amt} EGP</b>{_usd_tag(c.get("usd"))} · <span style="color:#8a7a52">{esc(c["label"])}</span>'
            f' <span style="color:#9a9a9a;font-size:11px">· {recency}</span></div>')
    return rows


def render(t, b):
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

    sections, chips = "", []
    for bk in b["buckets"]:
        rows = bk["rows"]
        if bk["key"] in _PRIMARY:
            sections += f"""
        <div class="bkt">
          <div class="bh"><div class="bt">{bk['title']}</div>
            <div class="bv">{bk['count']} clients · {fmt(bk['value'])} EGP</div></div>
          {_client_table(rows)}
        </div>"""
        else:
            val = f" · {fmt(bk['value'])} EGP" if bk["value"] else ""
            chips.append(f"{bk['title']}: <b>{bk['count']}</b>{val}")
    if chips:
        sections += f'<div class="chips">Earlier stages — {" &nbsp;·&nbsp; ".join(chips)}</div>'

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
    <h3 style="margin-top:0">🔎 Focus on these clients today <span style="font-size:11px;color:#9a9a9a;font-weight:400">— biggest {"recent " if b['focus_recent'] else ""}tickets</span></h3>
    {_focus(b)}
  </div>

  <div class="card">
    <h3 style="margin-top:0">Who's missing what <span style="font-size:11px;color:#9a9a9a;font-weight:400">— top {_MAXROWS} by value · <span class="stale">amber</span> = &gt;{STALE_DAYS}d untouched</span></h3>
    {sections}
  </div>

  <div class="foot">AHD Group sales · auto-generated daily from the live tracker.<br>
    LOST deals are excluded. <span class="usd">$→EGP</span> = amount looked like unmarked USD, converted at {esc(f"{t.get('usd_egp_rate', 0):.2f}")} EGP/USD.<br>
    Update a client's stage/status in the tracker and it reflects here next run.</div>
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
    html_body = render(t, b)
    out = os.path.join(HERE, "followup_report.html")
    with open(out, "w", encoding="utf-8") as fh:
        fh.write(html_body)
    print(f"Wrote {out}")
    print(json.dumps({"open_count": b["open_count"], "open_value": round(b["open_value"]),
                      "won_orders": b["won_orders"],
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
    if send_email(subject, html_body, recipients):
        _mark_sent_today()
    return 0


if __name__ == "__main__":
    sys.exit(main())
