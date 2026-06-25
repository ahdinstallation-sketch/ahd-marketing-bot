#!/usr/bin/env python3
"""AHD Group — Daily Marketing Pulse.

Merges Meta ads (meta_pull) + Google Sheets pipeline (sheets_pull) into a single
action-oriented HTML email and sends it via SMTP (Outlook/Microsoft 365 or Gmail;
host auto-picked from the sender domain).

Env vars:
  META_TOKEN       Meta Graph API token (System User / long-lived)
  MAIL_USER        sender address (e.g. projects@designy-egypt.com)
  MAIL_PASSWORD    app password for that mailbox
  SMTP_HOST/PORT   (optional) override SMTP host/port; inferred from domain if unset
  RECIPIENTS       comma-separated list (overrides the default below)
  CASHIN_SHEET_ID  (optional) Looker-backed sheet for live cash-in to date
  (legacy GMAIL_USER / GMAIL_APP_PASSWORD still accepted as fallbacks)

Usage:
  python3 daily_report.py --dry-run   # build report.html, print summary, DO NOT send
  python3 daily_report.py             # build + send (guarded to 9 AM Cairo unless --force)
  python3 daily_report.py --force     # build + send now regardless of the hour
"""
import os, sys, json, html, datetime, smtplib, ssl
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

from meta_pull import pull_meta
from sheets_pull import pull_sheets

HERE = os.path.dirname(os.path.abspath(__file__))
DEFAULT_RECIPIENTS = [
    "ahmed.helmy@amrhelmydesigns.com",
    "helmymalak@gmail.com",
    "nourannoor4@gmail.com",
]
SEND_HOUR_CAIRO = 9
# Pipeline-leak recency window. Default ~3 months; override with PIPELINE_LEAK_DAYS.
LEAK_WINDOW_DAYS = int(os.environ.get("PIPELINE_LEAK_DAYS", "90") or 90)


def _recent(deals, days=LEAK_WINDOW_DAYS):
    """Keep deals dated within the window (undated deals are excluded)."""
    return [d for d in deals
            if d.get("days_ago") is not None and 0 <= d["days_ago"] <= days]


def cairo_now():
    return datetime.datetime.utcnow() + datetime.timedelta(hours=3)


def fmt(n, prefix="", suffix=""):
    if n is None:
        return "—"
    if isinstance(n, float):
        n = round(n)
    return f"{prefix}{n:,}{suffix}"


def arrow(cur, ref):
    if cur is None or ref is None or ref == 0:
        return ""
    if cur > ref * 1.02:
        return ' <span style="color:#c0392b">▲</span>'  # spend up = red
    if cur < ref * 0.98:
        return ' <span style="color:#27ae60">▼</span>'
    return ""


def esc(s):
    return html.escape(str(s or ""))


# ---------- linkage math ----------

def build_context(meta, sheets):
    ctx = {"meta": meta, "sheets": sheets}
    # group spend
    y_spend = sum((a.get("yesterday", {}) or {}).get("spend", 0) or 0
                  for a in meta.get("accounts", []))
    m_spend = sum((a.get("mtd", {}) or {}).get("spend", 0) or 0
                  for a in meta.get("accounts", []))
    y_leads = sum((a.get("yesterday", {}) or {}).get("leads", 0) or 0
                  for a in meta.get("accounts", []))
    m_leads = sum((a.get("mtd", {}) or {}).get("leads", 0) or 0
                  for a in meta.get("accounts", []))
    w_spend = sum((a.get("last7", {}) or {}).get("spend", 0) or 0
                  for a in meta.get("accounts", []))
    w_leads = sum((a.get("last7", {}) or {}).get("leads", 0) or 0
                  for a in meta.get("accounts", []))
    ctx["group"] = {
        "spend_yest": y_spend, "spend_mtd": m_spend,
        "leads_yest": y_leads, "leads_mtd": m_leads,
        "spend_7d": w_spend, "leads_7d": w_leads,
        "cpl_yest": round(y_spend / y_leads, 1) if y_leads else None,
        "cpl_7d": round(w_spend / w_leads, 1) if w_leads else None,
        "cpl_mtd": round(m_spend / m_leads, 1) if m_leads else None,
    }
    tr = sheets.get("ahd_tracker", {}) or {}
    f = tr.get("funnel", {}) or {}
    signed = tr.get("signed_value", 0) or 0
    ctx["sales"] = {
        "sessions": f.get("SESSION", 0),
        "contracted": f.get("CONTRACTED", 0),
        "orders": f.get("ORDER", 0),
        "signed_value": signed,
        "open_pipeline": tr.get("open_pipeline_value", 0),
        # rough efficiency: signed EGP per MTD ad pound (EGP accounts only)
        "egp_per_spend": round(signed / m_spend, 1) if m_spend else None,
    }
    return ctx


# ---------- HTML ----------

def render(ctx):
    meta = ctx["meta"]
    sheets = ctx["sheets"]
    g = ctx["group"]
    s = ctx["sales"]
    today = cairo_now().date()
    P = []
    P.append(f"""<!doctype html><html><head><meta charset="utf-8">
      <meta name="viewport" content="width=device-width,initial-scale=1"></head>
      <body style="margin:0;background:#f0ece1;padding:14px">
      <div style="font-family:-apple-system,Segoe UI,Roboto,Arial,sans-serif;
      max-width:720px;margin:auto;color:#1a1a1a">
      <div style="background:#111;color:#e8d9a0;padding:18px 22px;border-radius:8px 8px 0 0">
        <div style="font-size:20px;font-weight:700">AHD Group — Daily Marketing Pulse</div>
        <div style="font-size:13px;opacity:.8">{today:%A, %d %B %Y} · Cairo</div>
      </div>
      <div style="padding:18px 22px;border:1px solid #eee;border-top:none;background:#fff">""")

    # 1) headline numbers
    P.append(f"""<div style="display:flex;gap:10px;flex-wrap:wrap;margin-bottom:6px">
      {_kpi('Ad spend yesterday', fmt(g['spend_yest']))}
      {_kpi('Spend month-to-date', fmt(g['spend_mtd']))}
      {_kpi('Leads yesterday', fmt(g['leads_yest']))}
      {_kpi('CPL (MTD)', fmt(g['cpl_mtd']))}
    </div>
    <div style="font-size:11px;color:#888;margin-bottom:14px">
      Spend in each account's own currency; EGP accounts dominate. CPL = ad spend ÷ leads.</div>""")

    # 1b) COST PER LEAD — headline metric for the profitability calculator
    P.append(_cpl_hero(ctx))

    # 2) per-account performance table
    P.append('<h3 style="margin:14px 0 6px;font-size:15px">Spend & performance by account</h3>')
    P.append("""<table style="width:100%;border-collapse:collapse;font-size:13px">
      <tr style="background:#f5f1e6;text-align:right">
        <th style="text-align:left;padding:6px">Account</th>
        <th style="padding:6px">Spend y'day</th><th style="padding:6px">Spend MTD</th>
        <th style="padding:6px">Leads y'day</th><th style="padding:6px">CPL y'day</th>
        <th style="padding:6px">CTR y'day</th></tr>""")
    for a in meta.get("accounts", []):
        y = a.get("yesterday", {}) or {}
        m = a.get("mtd", {}) or {}
        if "error" in y:
            P.append(f"""<tr><td style="padding:6px">{esc(a['name'])}</td>
              <td colspan="5" style="padding:6px;color:#b00">{esc(y['error'])[:80]}</td></tr>""")
            continue
        P.append(f"""<tr style="text-align:right;border-bottom:1px solid #eee">
          <td style="text-align:left;padding:6px">{esc(a['name'])}
            <span style="color:#aaa">{esc(a['currency'])}</span></td>
          <td style="padding:6px">{fmt(y.get('spend'))}</td>
          <td style="padding:6px">{fmt(m.get('spend'))}</td>
          <td style="padding:6px">{fmt(y.get('leads'))}</td>
          <td style="padding:6px">{fmt(y.get('cpl'))}</td>
          <td style="padding:6px">{fmt(y.get('ctr'))}%</td></tr>""")
    P.append("</table>")

    # 3) marketing -> sales linkage
    P.append('<h3 style="margin:18px 0 6px;font-size:15px">Marketing → Sales effect</h3>')
    P.append(f"""<div style="background:#faf7ee;border:1px solid #eee;border-radius:6px;
      padding:10px 14px;font-size:13px;line-height:1.7">
      <b>Funnel (live tracker):</b> Showroom sessions <b>{fmt(s['sessions'])}</b>
      → Contracted <b>{fmt(s['contracted'])}</b> → Orders <b>{fmt(s['orders'])}</b><br>
      <b>Signed / contracted value:</b> {fmt(s['signed_value'], 'EGP ')} ·
      <b>Open pipeline:</b> {fmt(s['open_pipeline'], 'EGP ')}<br>
      <b>Spend MTD vs signed:</b> {fmt(g['spend_mtd'], 'EGP ~')} ad spend →
      {('EGP ' + fmt(s['egp_per_spend']) + ' signed per EGP spent') if s['egp_per_spend'] else 'signed/ spend ratio n/a'}
    </div>""")

    # 3b) live cash-in to date (Looker-Studio-backed sheet, auto-updates)
    P.append(_cash_in_block(sheets.get("cash_in")))

    # 4) boost winners (organic)
    P.append('<h3 style="margin:18px 0 6px;font-size:15px">⬆ Boost these — top organic posts</h3>')
    any_posts = False
    for pg in meta.get("pages", []):
        if pg.get("error"):
            P.append(f"""<div style="font-size:12px;color:#b00">{esc(pg['name'])}: {esc(pg['error'])[:90]}</div>""")
            continue
        for post in pg.get("posts", [])[:3]:
            any_posts = True
            link = esc(post["link"])
            metrics = [f"{fmt(post['shares'])} shares"]
            if post.get("reactions") is not None:
                metrics.append(f"{fmt(post['reactions'])} reactions")
            if post.get("comments") is not None:
                metrics.append(f"{fmt(post['comments'])} comments")
            P.append(f"""<div style="border-left:3px solid #27ae60;padding:6px 10px;margin:6px 0;
              background:#f6fbf7;font-size:13px">
              <b>{esc(pg['name'])}</b> · {esc(post['created'])} ·
              {esc(' · '.join(metrics))}<br>
              <span style="color:#444">{esc(post['msg'])}</span>
              {(' · <a href="'+link+'">view</a>') if link else ''}</div>""")
    if not any_posts:
        P.append('<div style="font-size:12px;color:#888">No organic post data yet (token/page scope).</div>')

    # 5) kill/fix losers (paid)
    P.append('<h3 style="margin:18px 0 6px;font-size:15px">⛔ Fix or pause — paid ads leaking budget</h3>')
    any_loser = False
    for a in meta.get("accounts", []):
        for ad in a.get("ad_losers", [])[:3]:
            any_loser = True
            why = "no leads" if not ad["leads"] else (f"CPL {fmt(ad['cpl'])}" )
            P.append(f"""<div style="border-left:3px solid #c0392b;padding:6px 10px;margin:6px 0;
              background:#fdf6f5;font-size:13px">
              <b>{esc(ad['ad_name'] or ad['campaign'])}</b>
              <span style="color:#888">({esc(a['name'])})</span><br>
              {fmt(ad['spend'])} spent / 7d · CTR {fmt(ad['ctr'])}% · {esc(why)} — review creative/audience or pause.</div>""")
    if not any_loser:
        P.append('<div style="font-size:12px;color:#888">No clear paid losers flagged (good, or token/scope).</div>')

    # 5b) media-buyer view: delivery health + per-campaign 7d trend + fatigue
    P.append(_media_buyer_block(meta))

    # 5c) lead intelligence: brand-fit scoring + best-fit leads to call today
    P.append(_lead_intel_block(sheets.get("ahd_leads")))

    # 6) pipeline leaks (from sheets) — limited to the recency window
    tr = sheets.get("ahd_tracker", {}) or {}
    win = LEAK_WINDOW_DAYS
    months = max(1, round(win / 30))
    P.append(f"""<h3 style="margin:18px 0 6px;font-size:15px">Pipeline leaks — money at risk
      <span style="font-size:11px;color:#999;font-weight:400">— last {months} month{'s' if months!=1 else ''} (by offer/session date)</span></h3>""")
    leaks = [
        ("OVER BUDGET — re-spec to budget / cheque plan, don't blanket-discount",
         _recent(tr.get("over_budget", []), win), "#e67e22"),
        ("NO ANSWER AFTER OFFER — re-contact within 48h",
         _recent(tr.get("no_answer_after_offer", []), win), "#8e44ad"),
        ("CONTRACTED, NO ORDER — chase to deposit/production",
         _recent(tr.get("contracted_no_order", []), win), "#2980b9"),
    ]
    if any(d for _, d, _ in leaks):
        for title, deals, color in leaks:
            P.append(_leak_block(title, deals, color))
    else:
        P.append(f"""<div style="font-size:12px;color:#888">No pipeline leaks dated within
          the last {win} days. (Older flagged deals exist but fall outside the window — widen
          it with PIPELINE_LEAK_DAYS if you want them surfaced.)</div>""")

    # leads snapshot
    dl = sheets.get("designy_leads", {}) or {}
    ol = sheets.get("outdoor_leads", {}) or {}
    P.append(f"""<div style="font-size:12px;color:#666;margin-top:10px">
      Leads in sheets — Designy: {fmt(dl.get('total'))} ({fmt(dl.get('last7'))} last 7d) ·
      Outdoor: {fmt(ol.get('total'))} ({fmt(ol.get('last7'))} last 7d)</div>""")

    # 7) today's focus
    focus = pick_focus(ctx)
    P.append(f"""<div style="margin-top:16px;background:#111;color:#e8d9a0;padding:12px 16px;
      border-radius:6px;font-size:14px"><b>Today's one focus:</b> {esc(focus)}</div>""")

    P.append("""<div style="font-size:11px;color:#aaa;margin-top:14px">
      Auto-generated daily at 9 AM Cairo · Meta Graph API + live sales sheets ·
      Spend in account currency · not financial advice.</div></div></div></body></html>""")
    return "".join(P)


def _kpi(label, value):
    return (f'<div style="flex:1;min-width:150px;background:#f5f1e6;border-radius:6px;'
            f'padding:10px 12px"><div style="font-size:11px;color:#888">{esc(label)}</div>'
            f'<div style="font-size:20px;font-weight:700">{value}</div></div>')


def _cpl_big(label, value, sub=""):
    sub_html = f'<div style="font-size:11px;color:#bda86a;margin-top:2px">{esc(sub)}</div>' if sub else ""
    return (f'<div style="flex:1;min-width:140px;background:#1b1b1b;border:1px solid #3a3a2c;'
            f'border-radius:8px;padding:12px 14px">'
            f'<div style="font-size:11px;color:#cdbf8e;text-transform:uppercase;letter-spacing:.5px">{esc(label)}</div>'
            f'<div style="font-size:26px;font-weight:800;color:#f3e6b0">{value}</div>{sub_html}</div>')


def _cpl_hero(ctx):
    """Cost per lead — the headline number the team feeds into the profitability
    calculator. Shows yesterday / last-7d / month-to-date so they can pick a stable
    figure (7d or MTD), plus the per-account CPL underneath."""
    g = ctx["group"]
    meta = ctx["meta"]
    cards = (
        _cpl_big("CPL — last 7 days", fmt(g.get("cpl_7d"), "EGP "),
                 f"{fmt(g.get('leads_7d'))} leads · {fmt(g.get('spend_7d'),'EGP ')} spend")
        + _cpl_big("CPL — month to date", fmt(g.get("cpl_mtd"), "EGP "),
                   f"{fmt(g.get('leads_mtd'))} leads · {fmt(g.get('spend_mtd'),'EGP ')} spend")
        + _cpl_big("CPL — yesterday", fmt(g.get("cpl_yest"), "EGP "),
                   f"{fmt(g.get('leads_yest'))} leads")
    )
    # per-account CPL line (only accounts with leads)
    parts = []
    for a in meta.get("accounts", []):
        m = a.get("mtd", {}) or {}
        if "error" in m or not m.get("leads"):
            continue
        parts.append(f"{esc(a['name'])} <b style='color:#f3e6b0'>"
                     f"{fmt(m.get('cpl'))} {esc(a['currency'])}</b> ({fmt(m.get('leads'))} leads)")
    by_acct = (' · '.join(parts)) or "no per-account leads yet"
    return (f"""<div style="background:#111;border-radius:8px;padding:14px 16px;margin:4px 0 16px">
      <div style="color:#f3e6b0;font-size:15px;font-weight:700;margin-bottom:10px">
        💰 Cost per lead (CPL) — feed this into the profitability calculator</div>
      <div style="display:flex;gap:10px;flex-wrap:wrap">{cards}</div>
      <div style="color:#bda86a;font-size:11.5px;margin-top:10px">
        Per-account (MTD): {by_acct}</div>
      <div style="color:#7d7458;font-size:10.5px;margin-top:4px">
        Use the 7-day or MTD figure (more stable than a single day). CPL = ad spend ÷ Meta leads.</div>
    </div>""")


def _cash_in_block(ci):
    """Live group treasury snapshot from the auto-updating Looker-Studio-backed
    sheet: latest month's cash-in/out + net, cumulative cash-in YTD, bank balance
    and total debt. Group-level (no per-company split in that sheet). Silent when
    the source isn't configured / unreadable."""
    if not ci or not ci.get("configured") or not ci.get("ok"):
        return ""
    mlabel = esc(ci.get("month_label") or "")
    cards = [_kpi(f"Cash in — {mlabel}", fmt(ci.get("cash_in_month"), "EGP "))]
    if ci.get("cash_in_ytd"):
        cards.append(_kpi(f"Cash in YTD {esc(ci.get('ytd_year') or '')}",
                          fmt(ci["cash_in_ytd"], "EGP ")))
    if ci.get("bank_balance"):
        cards.append(_kpi("Bank balance (incl FX)", fmt(ci["bank_balance"], "EGP ")))
    if ci.get("total_debt"):
        cards.append(_kpi("Total debt", fmt(ci["total_debt"], "EGP ")))
    # second line: cash out + net for the latest month
    line2 = []
    if ci.get("cash_out_month"):
        line2.append(f"Cash out {fmt(ci['cash_out_month'], 'EGP ')}")
    if ci.get("net_month") is not None:
        net = ci["net_month"]
        col = "#27ae60" if net >= 0 else "#c0392b"
        line2.append(f'Net <span style="color:{col};font-weight:700">'
                     f'{fmt(net, "EGP ")}</span>')
    line2_html = (f'<div style="font-size:12px;color:#666;margin-top:6px">'
                  f'{mlabel}: {" · ".join(line2)}</div>') if line2 else ""
    asof = f" · as of {esc(ci['as_of'])}" if ci.get("as_of") else ""
    return (f"""<h3 style="margin:18px 0 6px;font-size:15px">Cash &amp; treasury
      <span style="font-size:11px;color:#999;font-weight:400">— live from treasury dashboard{asof}</span></h3>
      <div style="display:flex;gap:10px;flex-wrap:wrap">{''.join(cards)}</div>
      {line2_html}
      <div style="font-size:11px;color:#999;margin-top:5px">
        Group-level cash position (auto-updates with the Looker Studio sheet). Cash flow is
        collection-timing, not profit.</div>""")


def _cpl_trend(cur, prev):
    """Arrow for CPL change (lower CPL = green/good, higher = red/bad)."""
    if cur is None or prev is None or prev == 0:
        return ""
    if cur > prev * 1.05:
        return f' <span style="color:#c0392b">▲{round((cur/prev-1)*100)}%</span>'
    if cur < prev * 0.95:
        return f' <span style="color:#27ae60">▼{round((1-cur/prev)*100)}%</span>'
    return ' <span style="color:#999">≈</span>'


def _media_buyer_block(meta):
    """A media buyer's daily read: what's actually delivering vs paused/blocked,
    which campaigns are improving vs decaying (7d vs prior 7d), and ad fatigue."""
    P = ['<h3 style="margin:18px 0 6px;font-size:15px">📊 Media-buyer view — delivery, '
         'campaign trend & fatigue</h3>']
    any_data = False
    for a in meta.get("accounts", []):
        dh = a.get("delivery", {}) or {}
        ct = a.get("campaign_trend", {}) or {}
        if "error" in a.get("yesterday", {}):
            continue
        any_data = True
        camps = dh.get("campaigns", {}) or {}
        ads = dh.get("ads", {}) or {}
        active_c = camps.get("ACTIVE", 0)
        active_a = ads.get("ACTIVE", 0)
        issues = dh.get("issue_ads", []) or []
        P.append(f"""<div style="border:1px solid #eee;border-radius:6px;padding:8px 12px;
          margin:6px 0;background:#fafafa;font-size:12.5px">
          <b>{esc(a['name'])}</b> <span style="color:#aaa">{esc(a.get('currency',''))}</span><br>
          Campaigns active: <b>{active_c}</b> · Ads active: <b>{active_a}</b>
          {(' · <span style="color:#c0392b">'+str(len(issues))+' ad(s) blocked in active campaigns</span>') if issues else ''}""")
        # flag blocked ads explicitly (these silently kill live delivery)
        for it in issues[:4]:
            P.append(f"""<div style="color:#c0392b;padding-left:8px">⚠ {esc(it['name'] or it['campaign'])}
              — {esc(it['status'])}</div>""")
        # per-campaign trend table (top spenders, 7d vs prior 7d)
        rows = ct.get("campaigns", []) if "error" not in ct else []
        if rows:
            P.append("""<table style="width:100%;border-collapse:collapse;font-size:12px;margin-top:6px">
              <tr style="background:#f0ece1;text-align:right">
              <th style="text-align:left;padding:4px">Campaign (7d)</th>
              <th style="padding:4px">Spend</th><th style="padding:4px">Leads</th>
              <th style="padding:4px">CPL</th><th style="padding:4px">vs prior</th>
              <th style="padding:4px">Freq</th></tr>""")
            for r in rows[:5]:
                freq = r.get("frequency", 0) or 0
                fatig = ' <span style="color:#c0392b">🔥</span>' if freq >= 3.5 else ""
                P.append(f"""<tr style="text-align:right;border-bottom:1px solid #eee">
                  <td style="text-align:left;padding:4px">{esc(r['name'][:34])}</td>
                  <td style="padding:4px">{fmt(r['spend'])}</td>
                  <td style="padding:4px">{fmt(r['leads'])}</td>
                  <td style="padding:4px">{fmt(r['cpl'])}</td>
                  <td style="padding:4px">{_cpl_trend(r['cpl'], r['prev_cpl'])}</td>
                  <td style="padding:4px">{fmt(freq)}{fatig}</td></tr>""")
            P.append("</table>")
        if "error" in ct:
            P.append(f'<div style="color:#b00;font-size:11px">trend: {esc(ct["error"])[:70]}</div>')
        P.append("</div>")
    if not any_data:
        P.append('<div style="font-size:12px;color:#888">No delivery data (token/scope).</div>')
    P.append("""<div style="font-size:11px;color:#999;margin-top:2px">
      CPL arrow = change vs the previous 7 days (▼ green = improving, ▲ red = worse).
      Freq ≥ 3.5 (🔥) = audience seeing the ad too often — refresh creative or widen targeting.</div>""")
    return "".join(P)


# Recommended next action per lead tier (first-party-data driven, privacy-safe).
_TIER_ACTION = {
    "A": "Call within the hour — premium fit + ready to buy. Book a showroom session.",
    "B": "Call today — qualify budget & timeline, nurture toward a session.",
    "C": "WhatsApp a catalogue + offer; low priority, batch these.",
}


def _lead_intel_block(li):
    """Brand-fit lead intelligence: where leads come from, what they want, and the
    specific best-fit people to call first (name + phone + why), scored from the
    Facebook form answers (urgency, project size, premium compound)."""
    if not li or li.get("error") or not li.get("total"):
        return ('<h3 style="margin:18px 0 6px;font-size:15px">🎯 Lead intelligence</h3>'
                '<div style="font-size:12px;color:#888">No lead-form data available '
                f'({esc((li or {}).get("error","")) or "empty feed"}).</div>')
    t = li["tiers"]
    P = [f"""<h3 style="margin:18px 0 6px;font-size:15px">🎯 Lead intelligence — brand fit & who to call
      <span style="font-size:11px;color:#999;font-weight:400">— {fmt(li['total'])} leads scored ·
      {fmt(li['last7'])} new last 7d</span></h3>"""]
    # tier summary + unworked warning
    P.append(f"""<div style="display:flex;gap:10px;flex-wrap:wrap;margin-bottom:8px">
      {_kpi('🟢 Tier A (call now)', fmt(t['A']))}
      {_kpi('🟡 Tier B (today)', fmt(t['B']))}
      {_kpi('⚪ Tier C (nurture)', fmt(t['C']))}
    </div>""")
    if li.get("unworked"):
        P.append(f"""<div style="background:#fdf6f5;border:1px solid #f0d4cf;border-radius:6px;
          padding:8px 12px;font-size:12.5px;color:#a33;margin-bottom:8px">
          ⚠ <b>{fmt(li['unworked'])}</b> of {fmt(li['total'])} leads still show status CREATED
          (un-worked in the feed) — speed-to-lead is the #1 fixable leak.</div>""")
    # breakdowns
    def chips(d):
        return " · ".join(f"{esc(k)} <b>{v}</b>" for k, v in d.items())
    P.append(f"""<div style="font-size:12.5px;line-height:1.8;background:#faf7ee;
      border:1px solid #eee;border-radius:6px;padding:8px 12px;margin-bottom:8px">
      <b>Interest:</b> {chips(li['by_interest'])}<br>
      <b>Timeline:</b> {chips(li['by_urgency'])}<br>
      <b>Platform:</b> {chips(li['by_platform'])}<br>
      <b>Top compounds:</b> {chips(li['by_compound'])}</div>""")
    # best-fit leads to call, with action
    P.append('<div style="font-weight:600;font-size:13px;margin:8px 0 4px">'
             'Best-fit leads to action first:</div>')
    shown_tiers = set()
    for l in li.get("top_leads", [])[:10]:
        color = {"A": "#27ae60", "B": "#e67e22", "C": "#999"}.get(l["tier"], "#999")
        why = (" · " + ", ".join(l["why"])) if l.get("why") else ""
        ago = f"{l['days_ago']}d ago" if isinstance(l.get("days_ago"), int) else ""
        action = _TIER_ACTION.get(l["tier"], "")
        act_line = (f'<div style="color:#555;font-size:11.5px;margin-top:2px">→ {esc(action)}</div>'
                    if l["tier"] not in shown_tiers else "")
        shown_tiers.add(l["tier"])
        P.append(f"""<div style="border-left:3px solid {color};padding:5px 10px;margin:4px 0;
          background:#fcfcfa;font-size:12.5px">
          <b>{esc(l['name'] or '(no name)')}</b>
          <span style="background:{color};color:#fff;border-radius:3px;padding:0 5px;font-size:10px">
          {esc(l['tier'])} · {l['score']}</span>
          {('· <a href="tel:'+esc(l['phone'])+'">'+esc(l['phone'])+'</a>') if l.get('phone') else '· <span style="color:#c0392b">no phone</span>'}
          <span style="color:#aaa">{esc(ago)}</span><br>
          <span style="color:#444">{esc(l['interest'])} · {esc(l['when'])} ·
          {esc(l['compound'] or 'no compound')} · {esc(l['platform'])}{esc(why)}</span>
          {act_line}</div>""")
    P.append("""<div style="font-size:11px;color:#999;margin-top:4px">
      Score = urgency + project size + premium compound + contactability (first-party form
      answers only — no external profiling). Tier A ≥ 6, B 4–5, C &lt; 4.</div>""")
    return "".join(P)


def _leak_block(title, deals, color):
    if not deals:
        return ""
    rows = []
    for d in deals[:6]:
        amt = d.get("amount") or ""
        ago = d.get("days_ago")
        age = f"· {ago}d ago" if isinstance(ago, int) else ""
        rows.append(f"""<div style="font-size:12.5px;padding:2px 0">
          • <b>{esc(d.get('client'))}</b> {('· '+esc(amt)) if amt else ''}
          {('· '+esc(d.get('rep'))) if d.get('rep') else ''}
          <span style="color:#aaa">{age}</span></div>""")
    return (f'<div style="border-left:3px solid {color};padding:6px 10px;margin:6px 0">'
            f'<div style="font-weight:600;font-size:13px">{esc(title)} '
            f'<span style="color:#888">({len(deals)})</span></div>{"".join(rows)}</div>')


def pick_focus(ctx):
    tr = ctx["sheets"].get("ahd_tracker", {}) or {}
    ob = _recent(tr.get("over_budget", []))
    na = _recent(tr.get("no_answer_after_offer", []))
    co = _recent(tr.get("contracted_no_order", []))
    # biggest single money-at-risk item
    cands = []
    if ob:
        cands.append((ob[0].get("amount_num", 0),
                      f"Re-engage the OVER-BUDGET cluster ({len(ob)} deals) — start with {ob[0].get('client')}."))
    if na:
        cands.append((na[0].get("amount_num", 0),
                      f"Chase NO-ANSWER-AFTER-OFFER ({len(na)}) — {na[0].get('client')} is the biggest at risk."))
    if co:
        cands.append((co[0].get("amount_num", 0),
                      f"Convert CONTRACTED-NO-ORDER ({len(co)}) to production — {co[0].get('client')} first."))
    if cands:
        return max(cands, key=lambda x: x[0])[1]
    return "Review yesterday's spend vs leads and keep follow-up speed tight."


# ---------- send ----------

def _smtp_settings(user):
    """Pick SMTP host/port. Explicit SMTP_HOST/SMTP_PORT win; otherwise infer
    from the sender domain (Outlook/Microsoft 365 vs Gmail)."""
    host = os.environ.get("SMTP_HOST", "").strip()
    port = os.environ.get("SMTP_PORT", "").strip()
    if host:
        return host, int(port or 587)
    dom = user.split("@")[-1].lower()
    if dom in ("gmail.com", "googlemail.com"):
        return "smtp.gmail.com", 465            # SSL
    if dom in ("outlook.com", "hotmail.com", "live.com", "msn.com"):
        return "smtp-mail.outlook.com", 587     # STARTTLS
    # Custom domain on Microsoft 365 (e.g. designy-egypt.com) -> 365 relay.
    return "smtp.office365.com", 587            # STARTTLS


def send_email(subject, html_body, recipients):
    # New generic names, with backward-compatible GMAIL_* fallback.
    user = (os.environ.get("MAIL_USER") or os.environ.get("GMAIL_USER") or "").strip()
    pw = (os.environ.get("MAIL_PASSWORD") or os.environ.get("GMAIL_APP_PASSWORD") or "")
    pw = pw.replace(" ", "").strip()
    if not user or not pw:
        print("MAIL_USER / MAIL_PASSWORD (or GMAIL_*) not set — cannot send.")
        return False
    host, port = _smtp_settings(user)
    msg = MIMEMultipart("alternative")
    msg["Subject"] = subject
    msg["From"] = f"AHD Marketing Bot <{user}>"
    msg["To"] = ", ".join(recipients)
    msg.attach(MIMEText("Your client doesn't support HTML. Open in an HTML-capable client.", "plain", "utf-8"))
    msg.attach(MIMEText(html_body, "html", "utf-8"))
    ctx = ssl.create_default_context()
    if port == 465:
        with smtplib.SMTP_SSL(host, port, context=ctx) as server:
            server.login(user, pw)
            server.sendmail(user, recipients, msg.as_string())
    else:
        with smtplib.SMTP(host, port) as server:
            server.ehlo()
            server.starttls(context=ctx)
            server.login(user, pw)
            server.sendmail(user, recipients, msg.as_string())
    print(f"Sent to: {', '.join(recipients)} via {host}:{port}")
    return True


def main():
    dry = "--dry-run" in sys.argv
    force = "--force" in sys.argv
    recipients = [r.strip() for r in os.environ.get("RECIPIENTS", "").split(",") if r.strip()] \
        or DEFAULT_RECIPIENTS

    print("Pulling Meta…")
    meta = pull_meta()
    print("Pulling Sheets…")
    sheets = pull_sheets()
    ctx = build_context(meta, sheets)
    html_body = render(ctx)

    out = os.path.join(HERE, "report.html")
    with open(out, "w", encoding="utf-8") as fh:
        fh.write(html_body)
    print(f"Wrote {out}")
    print(json.dumps({"group": ctx["group"], "sales": ctx["sales"],
                      "meta_token": meta.get("token_present"),
                      "meta_error": meta.get("error")}, ensure_ascii=False, indent=1))

    subject = f"AHD Marketing Pulse — {cairo_now():%a %d %b}"
    if dry:
        print("DRY RUN — not sending.")
        return 0
    if not force and cairo_now().hour != SEND_HOUR_CAIRO:
        print(f"Cairo hour is {cairo_now().hour}, not {SEND_HOUR_CAIRO} — skipping send "
              f"(use --force to override).")
        return 0
    send_email(subject, html_body, recipients)
    return 0


if __name__ == "__main__":
    sys.exit(main())
