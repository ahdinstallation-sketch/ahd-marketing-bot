#!/usr/bin/env python3
"""AHD Group — Daily Cash Position report.

A standalone daily email (separate from the marketing pulse) summarising the
group's live treasury position from the Cash / Treasury workbook that backs the
Looker Studio dashboard. Pulls the public Google Sheet each run (no auth), builds
a clean HTML email, and sends it via the same SMTP path as the marketing bot.

Source workbook (public, gviz CSV export):
  CASH_SHEET_ID   default = the treasury workbook id below. Tabs used:
    - "Monthly Profit Report"  -> cash in/out/net, bank balances, debt, ratios
    - "Cash In by Company"     -> per-company collections (latest month + YTD)
    - "Cash Out Details"       -> spend breakdown (latest month)
    - "Bank Balance by Bank"   -> per-bank EGP + FX balances

Env vars:
  MAIL_USER / MAIL_PASSWORD   sender + app password (reused from marketing bot;
                              GMAIL_* still accepted as fallback)
  CASH_RECIPIENTS             comma-separated recipient list (overrides default)
  CASH_SHEET_ID               (optional) override the workbook id

Usage:
  python3 cash_report.py --dry-run   # build cash_report.html, print summary, DO NOT send
  python3 cash_report.py --force     # build + send now regardless of the hour
  python3 cash_report.py             # build + send (guarded to the morning window)
"""
import os, sys, json, html, datetime

from sheets_pull import fetch, num, _find_col
from daily_report import cairo_now, send_email

HERE = os.path.dirname(os.path.abspath(__file__))
CASH_SHEET_ID = os.environ.get("CASH_SHEET_ID", "").strip() or \
    "1_2HtZU0PzQ8b1ZPPg-4-kChAMC72E_DOD8495X8uHyM"

# Default recipients for the cash-position report (override with CASH_RECIPIENTS).
DEFAULT_RECIPIENTS = [
    "ahmed.helmy@amrhelmydesigns.com",   # Ahmed Amr Helmy
    "helmymalak@gmail.com",              # Malak Helmy
    # Mohamed Fahmy, Mohamed Abdelrahman, Hisham — add their addresses here /
    # via the CASH_RECIPIENTS secret before enabling the cron.
]

SEND_HOUR_CAIRO = 8
# Widened 21→23: GitHub cron can fire hours late, so accept any run 08:00–23:00 Cairo
# (dedupe keeps it once/day) rather than dropping a late fire and missing the day.
SEND_WINDOW_END_CAIRO = 23
SENT_MARKER = os.environ.get("CASH_SENT_MARKER") or os.path.join(HERE, ".cash_sent_marker")

_MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
           "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]


# ---------- formatting helpers ----------

def esc(s):
    return html.escape(str(s if s is not None else ""))


def fmt(n):
    """Thousands-separated integer EGP, e.g. 5522350.56 -> '5,522,351'."""
    try:
        return f"{round(float(n)):,}"
    except (TypeError, ValueError):
        return "0"


def fmt_m(n):
    """Compact millions, e.g. 5522351 -> '5.52M'."""
    try:
        return f"{float(n) / 1e6:.2f}M"
    except (TypeError, ValueError):
        return "0"


def _cell(r, i):
    return r[i] if (i is not None and len(r) > i) else ""


# ---------- data pull ----------

def _rows(tab):
    return fetch(CASH_SHEET_ID, tab=tab)


def pull_cash_position():
    """Assemble the treasury snapshot into a structured dict."""
    out = {"ok": False}

    # --- Monthly Profit Report: the spine (latest month with real cash) ---
    mp = _rows("Monthly Profit Report")
    if not mp:
        return out
    h = mp[0]
    ci = {n: _find_col(h, n) for n in [
        "Year", "Month", "Cash In", "Cash Out", "Balance", "Bank Balance EGP",
        "Bank Balance including FX", "Bank Balance EGP Only", "Revenue",
        "Outstanding Checks", "Hafele Debt EUR", "Hafele Debt EGP",
        "Total Debt EGP", "Debt to Revenue Ratio", "Euro Rate"]}

    def mrow(r, key):
        return _cell(r, ci.get(key))

    series = []  # (year, month, cash_in, cash_out, net) for months with cash
    latest = None
    for r in mp[1:]:
        cin = num(mrow(r, "Cash In"))
        if cin > 0:
            latest = r
            cout = num(mrow(r, "Cash Out"))
            series.append((
                mrow(r, "Year").strip(), mrow(r, "Month").strip(),
                cin, cout, cin - cout,  # net computed (num() drops '-' signs)
            ))
    if latest is None:
        return out

    out["ok"] = True
    out["month_label"] = f"{mrow(latest,'Month').strip()} {mrow(latest,'Year').strip()}".strip()
    out["cash_in"] = num(mrow(latest, "Cash In"))
    out["cash_out"] = num(mrow(latest, "Cash Out"))
    out["net"] = out["cash_in"] - out["cash_out"]  # signed; num() strips '-'
    out["bank_egp"] = num(mrow(latest, "Bank Balance EGP Only")) or num(mrow(latest, "Bank Balance EGP"))
    out["bank_fx"] = num(mrow(latest, "Bank Balance including FX"))
    out["outstanding_checks"] = num(mrow(latest, "Outstanding Checks"))
    out["hafele_eur"] = num(mrow(latest, "Hafele Debt EUR"))
    out["hafele_egp"] = num(mrow(latest, "Hafele Debt EGP"))
    out["total_debt"] = num(mrow(latest, "Total Debt EGP"))
    out["debt_to_rev"] = (mrow(latest, "Debt to Revenue Ratio") or "").strip()
    out["euro_rate"] = (mrow(latest, "Euro Rate") or "").strip()
    out["trend"] = series[-6:]  # last 6 months for the MoM strip

    # --- Cash In by Company (latest month + YTD) ---
    out["company"] = _pull_company()
    # --- Cash Out Details (latest month breakdown) ---
    out["cashout"] = _pull_cashout()
    # --- Bank Balance by Bank ---
    out["banks"] = _pull_banks()
    return out


def _pull_company():
    rows = _rows("Cash In by Company")
    if not rows:
        return None
    h = rows[0]
    cols = ["AHD", "Designy", "Projects", "Helmy", "YNG", "Countertops",
            "Appliances", "Esspressonalitea"]
    idx = {c: _find_col(h, c) for c in cols}
    i_total = _find_col(h, "Total EGP", "Total")
    i_yr, i_mo = _find_col(h, "Year"), _find_col(h, "Month")

    latest, ytd = None, {c: 0.0 for c in cols}
    ytd["total"] = 0.0
    for r in rows[1:]:
        if num(_cell(r, i_total)) > 0:
            latest = r
    if latest is None:
        return None
    yr = (_cell(latest, i_yr) or "").strip()
    for r in rows[1:]:
        if (_cell(r, i_yr) or "").strip() != yr or num(_cell(r, i_total)) <= 0:
            continue
        for c in cols:
            ytd[c] += num(_cell(r, idx[c]))
        ytd["total"] += num(_cell(r, i_total))
    month = {c: num(_cell(latest, idx[c])) for c in cols}
    month["total"] = num(_cell(latest, i_total))
    return {
        "month_label": f"{_cell(latest,i_mo).strip()} {yr}".strip(),
        "year": yr, "month": month, "ytd": ytd,
    }


def _pull_cashout():
    rows = _rows("Cash Out Details")
    if not rows:
        return None
    h = rows[0]
    cols = ["Material", "Holding", "Badr", "Café", "AHD", "helmy",
            "Road 9 Sofa", "Designy"]
    idx = {c: _find_col(h, c) for c in cols}
    i_total = _find_col(h, "Total EGP", "Total")
    i_mo, i_yr = _find_col(h, "Month"), _find_col(h, "Year")
    latest = None
    for r in rows[1:]:
        if num(_cell(r, i_total)) > 0:
            latest = r
    if latest is None:
        return None
    breakdown = [(c, num(_cell(latest, idx[c]))) for c in cols]
    breakdown = [(c, v) for c, v in breakdown if v > 0]
    breakdown.sort(key=lambda x: -x[1])
    return {
        "month_label": f"{_cell(latest,i_mo).strip()} {_cell(latest,i_yr).strip()}".strip(),
        "total": num(_cell(latest, i_total)),
        "breakdown": breakdown,
    }


def _pull_banks():
    rows = _rows("Bank Balance by Bank")
    if not rows:
        return None
    banks = []
    for r in rows[1:]:
        name = (_cell(r, 0) or "").strip()
        if not name:
            continue
        egp = num(_cell(r, 1))
        fx = num(_cell(r, 2))
        if egp == 0 and fx == 0:
            continue
        banks.append((name, egp, fx))
    if not banks:
        return None
    return {
        "rows": banks,
        "egp_total": sum(b[1] for b in banks),
        "fx_total": sum(b[2] for b in banks),
    }


# ---------- render ----------

CSS = """
body{margin:0;background:#f4f1ea;font-family:-apple-system,Segoe UI,Roboto,Helvetica,Arial,sans-serif;color:#2b2b2b}
.wrap{max-width:680px;margin:0 auto;padding:22px 18px 40px}
.head{text-align:center;padding:8px 0 14px}
.head h1{margin:0;font-size:21px;letter-spacing:.5px;color:#1c1c1c}
.head .sub{font-size:12.5px;color:#8a7a52;margin-top:4px}
.card{background:#fff;border:1px solid #e7e0cf;border-radius:12px;padding:16px 18px;margin:14px 0;box-shadow:0 1px 2px rgba(0,0,0,.03)}
.kpis{display:flex;gap:10px;flex-wrap:wrap}
.kpi{flex:1;min-width:150px;background:#fbf9f3;border:1px solid #ece4d2;border-radius:10px;padding:12px 14px}
.kpi .lbl{font-size:11px;text-transform:uppercase;letter-spacing:.6px;color:#9a8c66}
.kpi .val{font-size:22px;font-weight:700;margin-top:3px}
.kpi .sm{font-size:11.5px;color:#8a8a8a;margin-top:2px}
.pos{color:#1f7a44}.neg{color:#b3261e}
h3{margin:18px 0 8px;font-size:14.5px;color:#1c1c1c}
table{width:100%;border-collapse:collapse;font-size:13px}
th,td{padding:7px 8px;text-align:right;border-bottom:1px solid #f0eae0}
th:first-child,td:first-child{text-align:left}
thead th{font-size:10.5px;text-transform:uppercase;letter-spacing:.5px;color:#9a8c66;border-bottom:1px solid #e7e0cf}
tfoot td{font-weight:700;border-top:2px solid #e7e0cf;border-bottom:none}
.bar{height:8px;background:#efe8d6;border-radius:5px;overflow:hidden;margin-top:4px}
.bar>i{display:block;height:100%;background:#c8a24a}
.note{font-size:11.5px;color:#9a9a9a;margin-top:8px}
.pending{font-size:11.5px;color:#b06a12;background:#fdf3e2;border:1px solid #f0d9a8;border-radius:8px;padding:8px 12px;margin:0 0 6px}
.foot{text-align:center;font-size:11px;color:#a9a190;margin-top:22px}
.legend{margin-top:6px;line-height:1.9}
.lg{display:inline-block;margin:0 14px 0 0;font-size:11.5px;color:#555}
.lg i{display:inline-block;width:11px;height:11px;border-radius:2px;vertical-align:middle;margin-right:5px}
.hb{margin:8px 0}
.hb .row{display:flex;align-items:center;gap:8px;margin:5px 0}
.hb .cap{width:70px;font-size:11.5px;color:#8a8a8a}
.hb .track{flex:1;background:#efe8d6;border-radius:6px;height:20px;overflow:hidden}
.hb .fill{height:100%;border-radius:6px}
.hb .amt{width:96px;text-align:right;font-size:13px;font-weight:700}
"""

# Infographic colour palette (warm/cool mix that prints well on the cream theme).
_PALETTE = ["#c8a24a", "#4f9d69", "#5a9aa8", "#a9743f", "#7d7f9a",
            "#9a6a8a", "#8a9a5a", "#6a86b0", "#c98b7a", "#b0a06a"]
_GREEN = "#4f9d69"
_RED = "#c98b7a"


def _hbars(rows, peak=None):
    """Horizontal comparison bars. rows = [(label, value, color)]."""
    peak = peak or max((v for _, v, _ in rows), default=0) or 1
    out = ['<div class="hb">']
    for label, v, color in rows:
        w = max(2, round(100 * v / peak))
        out.append(
            f'<div class="row"><div class="cap">{esc(label)}</div>'
            f'<div class="track"><div class="fill" style="width:{w}%;background:{color}"></div></div>'
            f'<div class="amt">{fmt(v)}</div></div>')
    out.append('</div>')
    return "".join(out)


def _stacked(items, height=24):
    """100%-composition bar + legend. items = [(label, value)]."""
    items = [(l, v) for l, v in items if v > 0]
    total = sum(v for _, v in items) or 1
    segs, legend = "", ""
    for i, (label, v) in enumerate(items):
        c = _PALETTE[i % len(_PALETTE)]
        pct = 100 * v / total
        segs += (f'<td style="background:{c};height:{height}px;width:{pct:.3f}%;'
                 f'font-size:0;line-height:0">&nbsp;</td>')
        legend += (f'<span class="lg"><i style="background:{c}"></i>'
                   f'{esc(label)} <b>{pct:.0f}%</b></span>')
    return (f'<table style="width:100%;border-collapse:collapse;border-radius:6px;'
            f'overflow:hidden;table-layout:fixed"><tr>{segs}</tr></table>'
            f'<div class="legend">{legend}</div>')


def _column_chart(trend, maxh=98):
    """Grouped in/out column chart across months (inline-block bars, email-safe)."""
    peak = max((max(cin, cout) for _, _, cin, cout, _ in trend), default=0) or 1
    cells = ""
    for yr, mo, cin, cout, net in trend:
        hin = max(3, round(maxh * cin / peak))
        hout = max(3, round(maxh * cout / peak))
        ncls = "pos" if net >= 0 else "neg"
        sign = "+" if net >= 0 else "−"
        cells += (
            f'<td style="text-align:center;vertical-align:bottom;padding:0 3px">'
            f'<div style="height:{maxh}px;line-height:{maxh}px;white-space:nowrap">'
            f'<span style="display:inline-block;width:13px;height:{hin}px;background:{_GREEN};'
            f'vertical-align:bottom;border-radius:3px 3px 0 0"></span>'
            f'<span style="display:inline-block;width:2px"></span>'
            f'<span style="display:inline-block;width:13px;height:{hout}px;background:{_RED};'
            f'vertical-align:bottom;border-radius:3px 3px 0 0"></span>'
            f'</div>'
            f'<div style="font-size:10.5px;color:#8a8a8a;margin-top:5px">{esc(mo)} {esc(yr[2:])}</div>'
            f'<div class="{ncls}" style="font-size:10px;font-weight:700">{sign}{fmt_m(abs(net))}</div>'
            f'</td>')
    return (f'<table style="width:100%;border-collapse:collapse"><tr valign="bottom">{cells}</tr></table>'
            f'<div class="legend" style="text-align:center;margin-top:8px">'
            f'<span class="lg"><i style="background:{_GREEN}"></i>Cash In</span>'
            f'<span class="lg"><i style="background:{_RED}"></i>Cash Out</span>'
            f'<span class="lg">Net shown per column (M EGP)</span></div>')


def render(d):
    if not d.get("ok"):
        return "<html><body><p>Cash position unavailable — could not read the treasury sheet.</p></body></html>"

    net = d["net"]
    net_cls = "pos" if net >= 0 else "neg"
    net_sign = "+" if net >= 0 else "−"
    today = cairo_now()
    cur_label = f"{today:%B %Y}"

    pending = ""
    # Latest month with cash vs current calendar month.
    ml = d["month_label"]
    if ml and ml.split()[0].lower() != today.strftime("%b").lower():
        pending = (f'<div class="pending">ℹ️ {esc(cur_label)} has no cash posted yet — '
                   f'showing the last closed month (<b>{esc(ml)}</b>). Updates automatically '
                   f'once {esc(cur_label)} collections land in the sheet.</div>')

    # Hero KPIs
    hero = f"""
    <div class="kpis">
      <div class="kpi"><div class="lbl">Cash In</div>
        <div class="val pos">{fmt(d['cash_in'])}</div><div class="sm">EGP · {esc(ml)}</div></div>
      <div class="kpi"><div class="lbl">Cash Out</div>
        <div class="val neg">{fmt(d['cash_out'])}</div><div class="sm">EGP · {esc(ml)}</div></div>
      <div class="kpi"><div class="lbl">Net Cash Movement</div>
        <div class="val {net_cls}">{net_sign}{fmt(abs(net))}</div><div class="sm">EGP · in − out</div></div>
    </div>
    {_hbars([("Cash In", d['cash_in'], _GREEN), ("Cash Out", d['cash_out'], _RED)])}"""

    # Bank position
    banks = d.get("banks")
    bank_rows = ""
    bank_stack = ""
    if banks:
        for name, egp, fx in banks["rows"]:
            bank_rows += (f"<tr><td>{esc(name)}</td><td>{fmt(egp)}</td>"
                          f"<td>{fmt(fx) if fx else '—'}</td></tr>")
        bank_stack = ('<div style="font-size:11px;color:#9a8c66;text-transform:uppercase;'
                      'letter-spacing:.5px;margin:2px 0 5px">EGP cash by bank</div>'
                      + _stacked([(n, e) for n, e, _ in banks["rows"] if e > 0]))
        bank_tbl = f"""
        <table>
          <thead><tr><th>Bank</th><th>EGP balance</th><th>FX (USD→EGP)</th></tr></thead>
          <tbody>{bank_rows}</tbody>
          <tfoot><tr><td>Total</td><td>{fmt(banks['egp_total'])}</td><td>{fmt(banks['fx_total'])}</td></tr></tfoot>
        </table>"""
    else:
        bank_tbl = ""

    bank_kpi = f"""
    <div class="kpis" style="margin-bottom:10px">
      <div class="kpi"><div class="lbl">Bank Balance (EGP)</div>
        <div class="val">{fmt(d['bank_egp'])}</div><div class="sm">EGP cash across banks</div></div>
      <div class="kpi"><div class="lbl">Bank incl. FX</div>
        <div class="val">{fmt(d['bank_fx'])}</div><div class="sm">EGP + foreign currency</div></div>
    </div>"""

    # Debt & obligations
    dr = d.get("debt_to_rev") or "—"
    haf = ""
    if d.get("hafele_egp"):
        haf = (f'<div class="kpi"><div class="lbl">Hafele Debt</div>'
               f'<div class="val neg">{fmt(d["hafele_egp"])}</div>'
               f'<div class="sm">EGP{" · €"+fmt(d["hafele_eur"]) if d.get("hafele_eur") else ""}</div></div>')
    debt = f"""
    <div class="kpis">
      <div class="kpi"><div class="lbl">Total Debt</div>
        <div class="val neg">{fmt(d['total_debt'])}</div><div class="sm">EGP</div></div>
      <div class="kpi"><div class="lbl">Outstanding Cheques</div>
        <div class="val">{fmt(d['outstanding_checks'])}</div><div class="sm">EGP not yet cleared</div></div>
      <div class="kpi"><div class="lbl">Debt / Revenue</div>
        <div class="val">{esc(dr)}</div><div class="sm">Euro rate {esc(d.get('euro_rate') or '—')}</div></div>
      {haf}
    </div>"""

    # Cash in by company
    comp = d.get("company")
    comp_html = ""
    if comp:
        m = comp["month"]
        y = comp["ytd"]
        order = ["AHD", "Designy", "Countertops", "Appliances", "Projects",
                 "Helmy", "YNG", "Esspressonalitea"]
        peak = max((m.get(c, 0) for c in order), default=0) or 1
        rows = ""
        for c in order:
            v = m.get(c, 0)
            if v <= 0 and y.get(c, 0) <= 0:
                continue
            w = max(2, round(100 * v / peak))
            rows += (f"<tr><td>{esc(c)}</td>"
                     f"<td>{fmt(v)}<div class='bar'><i style='width:{w}%'></i></div></td>"
                     f"<td>{fmt(y.get(c,0))}</td></tr>")
        comp_stack = _stacked([(c, m.get(c, 0)) for c in order])
        comp_html = f"""
        <div class="card">
          <h3 style="margin-top:0">Cash in by company — {esc(comp['month_label'])}</h3>
          {comp_stack}
          <table style="margin-top:12px">
            <thead><tr><th>Company</th><th>This month</th><th>YTD {esc(comp['year'])}</th></tr></thead>
            <tbody>{rows}</tbody>
            <tfoot><tr><td>Total</td><td>{fmt(m.get('total',0))}</td><td>{fmt(y.get('total',0))}</td></tr></tfoot>
          </table>
        </div>"""

    # Cash out breakdown
    co = d.get("cashout")
    co_html = ""
    if co and co["breakdown"]:
        peak = co["breakdown"][0][1] or 1
        rows = ""
        for name, v in co["breakdown"]:
            w = max(2, round(100 * v / peak))
            pct = 100 * v / co["total"] if co["total"] else 0
            rows += (f"<tr><td>{esc(name)}</td>"
                     f"<td>{fmt(v)}<div class='bar'><i style='width:{w}%'></i></div></td>"
                     f"<td>{pct:.0f}%</td></tr>")
        co_stack = _stacked(co["breakdown"])
        co_html = f"""
        <div class="card">
          <h3 style="margin-top:0">Cash out breakdown — {esc(co['month_label'])}</h3>
          {co_stack}
          <table style="margin-top:12px">
            <thead><tr><th>Category</th><th>Amount (EGP)</th><th>% of out</th></tr></thead>
            <tbody>{rows}</tbody>
            <tfoot><tr><td>Total</td><td>{fmt(co['total'])}</td><td>100%</td></tr></tfoot>
          </table>
        </div>"""

    # MoM trend
    tr = d.get("trend") or []
    tr_html = ""
    if tr:
        rows = ""
        for yr, mo, cin, cout, nt in tr:
            ncls = "pos" if nt >= 0 else "neg"
            sign = "+" if nt >= 0 else "−"
            rows += (f"<tr><td>{esc(mo)} {esc(yr[2:])}</td><td>{fmt(cin)}</td>"
                     f"<td>{fmt(cout)}</td><td class='{ncls}'>{sign}{fmt(abs(nt))}</td></tr>")
        tr_html = f"""
        <div class="card">
          <h3 style="margin-top:0">Monthly cash flow — last {len(tr)} months</h3>
          {_column_chart(tr)}
          <table style="margin-top:14px">
            <thead><tr><th>Month</th><th>Cash In</th><th>Cash Out</th><th>Net</th></tr></thead>
            <tbody>{rows}</tbody>
          </table>
          <div class="note">Net = collections − payments. A negative month means cheque clearances /
            debt service outran collections that month (timing, not necessarily a loss).</div>
        </div>"""

    return f"""<!doctype html><html><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1"><style>{CSS}</style></head>
<body><div class="wrap">
  <div class="head">
    <h1>AHD Group — Daily Cash Position</h1>
    <div class="sub">{esc(f'{today:%A, %d %B %Y}')} · as of {esc(ml)}</div>
  </div>
  {pending}
  <div class="card">
    <h3 style="margin-top:0">Cash movement — {esc(ml)}</h3>
    {hero}
    <div class="note">Cash flow is collection-timing, not profit. Positive months reflect
      collections landing; negative months reflect payments/cheque clearances outrunning them.</div>
  </div>
  <div class="card">
    <h3 style="margin-top:0">Bank position</h3>
    {bank_kpi}
    {bank_stack}
    {bank_tbl}
  </div>
  <div class="card">
    <h3 style="margin-top:0">Debt &amp; obligations</h3>
    {debt}
  </div>
  {comp_html}
  {co_html}
  {tr_html}
  <div class="foot">AHD Group treasury · auto-generated daily from the live cash workbook.<br>
    Figures update automatically as the sheet is filled.</div>
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
    recipients = [r.strip() for r in os.environ.get("CASH_RECIPIENTS", "").split(",") if r.strip()] \
        or DEFAULT_RECIPIENTS

    print("Pulling cash position…")
    d = pull_cash_position()
    html_body = render(d)
    out = os.path.join(HERE, "cash_report.html")
    with open(out, "w", encoding="utf-8") as fh:
        fh.write(html_body)
    print(f"Wrote {out}")
    print(json.dumps({k: d.get(k) for k in
                      ("ok", "month_label", "cash_in", "cash_out", "net",
                       "bank_egp", "bank_fx", "total_debt", "debt_to_rev")},
                     ensure_ascii=False, indent=1))

    subject = f"AHD Cash Position — {cairo_now():%a %d %b}"
    if dry:
        print("DRY RUN — not sending.")
        return 0
    if not force:
        if cairo_now().weekday() == 4:  # Friday (Mon=0 … Fri=4) — weekend, no send
            print("Cairo day is Friday — skipping (no cash report on Fridays).")
            return 0
        if _already_sent_today():
            print("Already sent today — skipping (dedupe).")
            return 0
        hour = cairo_now().hour
        if not (SEND_HOUR_CAIRO <= hour < SEND_WINDOW_END_CAIRO):
            print(f"Cairo hour is {hour}, outside send window "
                  f"{SEND_HOUR_CAIRO}:00–{SEND_WINDOW_END_CAIRO}:00 — skipping.")
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
