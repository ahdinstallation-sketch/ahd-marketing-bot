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
try:
    from zoneinfo import ZoneInfo
except ImportError:  # pragma: no cover (Python < 3.9)
    ZoneInfo = None

from meta_pull import pull_meta
from sheets_pull import pull_sheets

HERE = os.path.dirname(os.path.abspath(__file__))
DEFAULT_RECIPIENTS = [
    "ahmed.helmy@amrhelmydesigns.com",
    "helmymalak@gmail.com",
    "nourannoor4@gmail.com",
]
SEND_HOUR_CAIRO = 9
# GitHub's free cron scheduler is best-effort and can fire hours late, so we accept
# any run from 09:00 up to (but not including) this hour, and dedupe to once/day.
SEND_WINDOW_END_CAIRO = 21
# Once-per-Cairo-day send marker (persisted across GitHub Actions runs via actions/cache).
SENT_MARKER = os.environ.get("SENT_MARKER") or os.path.join(HERE, ".sent_marker")
# Pipeline-leak recency window. Default ~6 months; override with PIPELINE_LEAK_DAYS.
LEAK_WINDOW_DAYS = int(os.environ.get("PIPELINE_LEAK_DAYS", "180") or 180)


def _recent(deals, days=LEAK_WINDOW_DAYS):
    """Keep deals dated within the window (undated deals are excluded)."""
    return [d for d in deals
            if d.get("days_ago") is not None and 0 <= d["days_ago"] <= days]


def cairo_now():
    """Current local time in Cairo as a naive datetime (DST-correct via zoneinfo;
    falls back to the old fixed UTC+3 if the tz database is unavailable)."""
    if ZoneInfo is not None:
        try:
            return datetime.datetime.now(ZoneInfo("Africa/Cairo")).replace(tzinfo=None)
        except Exception:
            pass
    return datetime.datetime.utcnow() + datetime.timedelta(hours=3)


def already_sent_today():
    """True if we already sent today's email (marker holds today's Cairo date)."""
    try:
        with open(SENT_MARKER, encoding="utf-8") as fh:
            return fh.read().strip() == cairo_now().date().isoformat()
    except OSError:
        return False


def mark_sent_today():
    """Record that today's email went out, so other same-day runs skip it."""
    try:
        with open(SENT_MARKER, "w", encoding="utf-8") as fh:
            fh.write(cairo_now().date().isoformat())
    except OSError as exc:
        print(f"Warning: could not write sent-marker {SENT_MARKER}: {exc}")


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


def _egp_accounts(meta):
    """Display accounts scoped to EGP only (drops the Designy USD account so its
    spend never shows). Falls back to all accounts if none are EGP."""
    accts = [a for a in meta.get("accounts", []) if a.get("currency") == "EGP"]
    return accts or meta.get("accounts", [])


# ---------- linkage math ----------

def build_context(meta, sheets):
    ctx = {"meta": meta, "sheets": sheets}
    # Group headline spend/CPL is monetary, so scope it to EGP accounts only — summing
    # spend across currencies (EGP + Designy USD) would be meaningless. Leads are unitless
    # but we keep them on the same EGP basis so the headline CPL is internally consistent.
    accts = [a for a in meta.get("accounts", []) if a.get("currency") == "EGP"]
    if not accts:  # safety: never blank the headline if no EGP account is present
        accts = meta.get("accounts", [])
    y_spend = sum((a.get("yesterday", {}) or {}).get("spend", 0) or 0 for a in accts)
    m_spend = sum((a.get("mtd", {}) or {}).get("spend", 0) or 0 for a in accts)
    y_leads = sum((a.get("yesterday", {}) or {}).get("leads", 0) or 0 for a in accts)
    m_leads = sum((a.get("mtd", {}) or {}).get("leads", 0) or 0 for a in accts)
    w_spend = sum((a.get("last7", {}) or {}).get("spend", 0) or 0 for a in accts)
    w_leads = sum((a.get("last7", {}) or {}).get("leads", 0) or 0 for a in accts)
    # CPL numerator = spend on LEAD campaigns only (falls back to total spend), so
    # traffic/awareness campaign spend doesn't inflate cost-per-lead.
    y_lead_spend = sum((a.get("yesterday", {}) or {}).get("lead_spend",
                       (a.get("yesterday", {}) or {}).get("spend", 0)) or 0 for a in accts)
    m_lead_spend = sum((a.get("mtd", {}) or {}).get("lead_spend",
                       (a.get("mtd", {}) or {}).get("spend", 0)) or 0 for a in accts)
    w_lead_spend = sum((a.get("last7", {}) or {}).get("lead_spend",
                       (a.get("last7", {}) or {}).get("spend", 0)) or 0 for a in accts)
    ctx["group"] = {
        "spend_yest": y_spend, "spend_mtd": m_spend,
        "leads_yest": y_leads, "leads_mtd": m_leads,
        "spend_7d": w_spend, "leads_7d": w_leads,
        "cpl_yest": round(y_lead_spend / y_leads, 1) if y_leads else None,
        "cpl_7d": round(w_lead_spend / w_leads, 1) if w_leads else None,
        "cpl_mtd": round(m_lead_spend / m_leads, 1) if m_leads else None,
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
      Spend in each account's own currency; EGP accounts dominate. CPL = spend on lead
      campaigns ÷ leads (traffic / awareness spend is excluded so CPL isn't inflated).</div>""")

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
    for a in _egp_accounts(meta):
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

    # 3b) live cash-in to date (Looker-Studio-backed sheet, auto-updates)
    P.append(_cash_in_block(sheets.get("cash_in")))

    # 4) boost winners (organic) — the single top 3 across ALL pages, ranked by
    # engagement (shares weigh most for organic reach, then reactions, then comments)
    P.append('<h3 style="margin:18px 0 6px;font-size:15px">⬆ Boost these — top 3 organic posts</h3>')
    all_posts = []
    for pg in meta.get("pages", []):
        if pg.get("error"):
            P.append(f"""<div style="font-size:12px;color:#b00">{esc(pg['name'])}: {esc(pg['error'])[:90]}</div>""")
            continue
        for post in pg.get("posts", []):
            score = ((post.get("shares") or 0) * 3
                     + (post.get("reactions") or 0)
                     + (post.get("comments") or 0))
            all_posts.append((score, pg.get("name", ""), post))
    all_posts.sort(key=lambda t: -t[0])
    for score, pg_name, post in all_posts[:3]:
        link = esc(post["link"])
        metrics = [f"{fmt(post['shares'])} shares"]
        if post.get("reactions") is not None:
            metrics.append(f"{fmt(post['reactions'])} reactions")
        if post.get("comments") is not None:
            metrics.append(f"{fmt(post['comments'])} comments")
        P.append(f"""<div style="border-left:3px solid #27ae60;padding:6px 10px;margin:6px 0;
          background:#f6fbf7;font-size:13px">
          <b>{esc(pg_name)}</b> · {esc(post['created'])} ·
          {esc(' · '.join(metrics))}<br>
          <span style="color:#444">{esc(post['msg'])}</span>
          {(' · <a href="'+link+'">view</a>') if link else ''}</div>""")
    if not all_posts:
        P.append('<div style="font-size:12px;color:#888">No organic post data yet (token/page scope).</div>')

    # 5) kill/fix losers (paid)
    P.append('<h3 style="margin:18px 0 6px;font-size:15px">⛔ Fix or pause — paid ads leaking budget</h3>')
    any_loser = False
    for a in _egp_accounts(meta):
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

    # 5c-i) lead → contract conversion, from the reception team's Stage column
    P.append(_lead_conversion_block(sheets.get("ahd_leads")))

    # 5c) lead intelligence: brand-fit scoring + best-fit leads to call today
    P.append(_lead_intel_block(sheets.get("ahd_leads")))

    # 5d) lead → sales journey: leads matched by name to tracker clients
    P.append(_lead_journey_block(sheets.get("lead_journey")))

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
    """Cost per lead — the two numbers that matter: last-7d (recent trend) and
    month-to-date (stable). This is the figure the team feeds into the
    profitability calculator."""
    g = ctx["group"]
    cards = (
        _cpl_big("CPL — last 7 days", fmt(g.get("cpl_7d"), "EGP "),
                 f"{fmt(g.get('leads_7d'))} leads")
        + _cpl_big("CPL — month to date", fmt(g.get("cpl_mtd"), "EGP "),
                   f"{fmt(g.get('leads_mtd'))} leads")
    )
    return (f"""<div style="background:#111;border-radius:8px;padding:14px 16px;margin:4px 0 16px">
      <div style="color:#f3e6b0;font-size:15px;font-weight:700;margin-bottom:10px">
        💰 Cost per lead (CPL) — feed this into the profitability calculator</div>
      <div style="display:flex;gap:10px;flex-wrap:wrap">{cards}</div>
      <div style="color:#7d7458;font-size:10.5px;margin-top:8px">
        CPL = lead-campaign spend ÷ Meta leads (EGP accounts). Use MTD for a stable figure.</div>
    </div>""")


def _cash_in_block(ci):
    """Live group treasury snapshot from the auto-updating Looker-Studio-backed
    sheet: latest month's cash-in/out + net, cumulative cash-in YTD, bank balance
    and total debt. Group-level (no per-company split in that sheet). Silent when
    the source isn't configured / unreadable."""
    if not ci or not ci.get("configured") or not ci.get("ok"):
        return ""
    mlabel = esc(ci.get("month_label") or "")
    m = ci.get("month", {}) or {}
    y = ci.get("ytd", {}) or {}
    cards = []
    if m.get("total") is not None:
        cards.append(_kpi(f"Total cash in — {mlabel}", fmt(m["total"], "EGP ")))
    if m.get("ahd") is not None:
        cards.append(_kpi(f"AHD — {mlabel}", fmt(m["ahd"], "EGP ")))
    if m.get("designy") is not None:
        cards.append(_kpi(f"Designy — {mlabel}", fmt(m["designy"], "EGP ")))
    # YTD-per-company subline
    yparts = []
    if y.get("total") is not None:
        yparts.append(f"Total <b>{fmt(y['total'], 'EGP ')}</b>")
    if y.get("ahd") is not None:
        yparts.append(f"AHD <b>{fmt(y['ahd'], 'EGP ')}</b>")
    if y.get("designy") is not None:
        yparts.append(f"Designy <b>{fmt(y['designy'], 'EGP ')}</b>")
    yline = (f'<div style="font-size:12px;color:#666;margin-top:6px">'
             f'YTD {esc(ci.get("ytd_year") or "")}: {" · ".join(yparts)}</div>') if yparts else ""
    asof = f" · as of {esc(ci['as_of'])}" if ci.get("as_of") else ""
    return (f"""<h3 style="margin:18px 0 6px;font-size:15px">Cash in by company
      <span style="font-size:11px;color:#999;font-weight:400">— live from treasury sheet{asof}</span></h3>
      <div style="display:flex;gap:10px;flex-wrap:wrap">{''.join(cards)}</div>
      {yline}
      <div style="font-size:11px;color:#999;margin-top:5px">
        Collections by company (auto-updates with the sheet). Cash flow is collection-timing,
        not profit.</div>""")


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
    for a in _egp_accounts(meta):
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
    # tier summary
    P.append(f"""<div style="display:flex;gap:10px;flex-wrap:wrap;margin-bottom:8px">
      {_kpi('🟢 Tier A (call now)', fmt(t['A']))}
      {_kpi('🟡 Tier B (today)', fmt(t['B']))}
      {_kpi('⚪ Tier C (nurture)', fmt(t['C']))}
    </div>""")
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
        stg = (l.get("sale_stage") or "").strip()
        stg_badge = (f'<span style="background:#444;color:#fff;border-radius:3px;'
                     f'padding:0 5px;font-size:10px">{esc(stg)}</span>') if stg else ''
        P.append(f"""<div style="border-left:3px solid {color};padding:5px 10px;margin:4px 0;
          background:#fcfcfa;font-size:12.5px">
          <b>{esc(l['name'] or '(no name)')}</b>
          <span style="background:{color};color:#fff;border-radius:3px;padding:0 5px;font-size:10px">
          {esc(l['tier'])} · {l['score']}</span>
          {stg_badge}
          {('· <a href="tel:'+esc(l['phone'])+'">'+esc(l['phone'])+'</a>') if l.get('phone') else '· <span style="color:#c0392b">no phone</span>'}
          <span style="color:#aaa">{esc(ago)}</span><br>
          <span style="color:#444">{esc(l['interest'])} · {esc(l['when'])} ·
          {esc(l['compound'] or 'no compound')} · {esc(l['platform'])}{esc(why)}</span>
          {act_line}</div>""")
    P.append("""<div style="font-size:11px;color:#999;margin-top:4px">
      Score = urgency + project size + premium compound + contactability (first-party form
      answers only — no external profiling). Tier A ≥ 6, B 4–5, C &lt; 4.</div>""")
    return "".join(P)


# Reception funnel groups → colour (matches the Stage groups in sheets_pull).
_STAGE_GROUP_COLOUR = {
    "Won / paid": "#1e8e4e", "Quote made": "#27ae60", "Engaged": "#c9871f",
    "Contacted": "#2980b9", "Lost / out": "#c0392b", "Unqualified": "#9a9a9a",
    "Other": "#777777",
}


def _lead_conversion_block(al):
    """Lead → contract conversion straight from the reception team's own 'Stage'
    column in the AHD leads sheet: the % of qualified leads that reached a paid
    contract, the status funnel, and who actually converted."""
    if not al or not al.get("conversion"):
        return ""
    c = al["conversion"]
    funnel = al.get("stage_funnel", []) or []
    won_leads = al.get("won_leads", []) or []
    pct = c.get("pct")
    pct_str = f"{pct}%" if pct is not None else "—"

    # headline scoreboard
    scoreboard = f"""
    <table role="presentation" width="100%" cellpadding="0" cellspacing="0"
      style="margin:6px 0 10px;border-collapse:separate;border-spacing:6px 0">
      <tr>
        <td width="33%" style="background:#e9f7ef;border:1px solid #b7e2c8;border-radius:8px;
          padding:10px 8px;text-align:center">
          <div style="font-size:24px;font-weight:800;color:#1e8e4e;line-height:1">{pct_str}</div>
          <div style="font-size:10.5px;color:#3a6b4e;text-transform:uppercase;letter-spacing:.4px">
            Lead→contract</div></td>
        <td width="33%" style="background:#fdf3e2;border:1px solid #f0d9a8;border-radius:8px;
          padding:10px 8px;text-align:center">
          <div style="font-size:24px;font-weight:800;color:#c9871f;line-height:1">{fmt(c.get('won'))} <span
            style="font-size:13px;color:#8a6a2a">/ {fmt(c.get('qualified'))}</span></div>
          <div style="font-size:10.5px;color:#8a6a2a;text-transform:uppercase;letter-spacing:.4px">
            Won / qualified</div></td>
        <td width="33%" style="background:#eef4fb;border:1px solid #c3ddf3;border-radius:8px;
          padding:10px 8px;text-align:center">
          <div style="font-size:24px;font-weight:800;color:#2471b3;line-height:1">{fmt(c.get('groups_made'))}</div>
          <div style="font-size:10.5px;color:#2c5a83;text-transform:uppercase;letter-spacing:.4px">
            Quotes made</div></td>
      </tr>
    </table>"""

    # status funnel — one labelled bar per group, sized to the largest group
    maxn = max((n for _, n in funnel), default=1) or 1
    bars = []
    for label, n in funnel:
        col = _STAGE_GROUP_COLOUR.get(label, "#777")
        w = max(3, round(100 * n / maxn))
        bars.append(f"""<tr>
          <td width="90" style="font-size:11.5px;color:#444;padding:2px 6px 2px 0">{esc(label)}</td>
          <td style="padding:2px 0"><table role="presentation" cellpadding="0" cellspacing="0"
            style="width:100%"><tr>
            <td style="background:{col};height:14px;width:{w}%;border-radius:3px"></td>
            <td width="34" style="font-size:11px;font-weight:700;color:{col};padding-left:6px">{fmt(n)}</td>
            </tr></table></td></tr>""")
    funnel_html = (f'<table role="presentation" width="100%" cellpadding="0" cellspacing="0" '
                   f'style="margin:2px 0 8px">{"".join(bars)}</table>') if bars else ""

    # who converted (money down)
    won_rows = []
    for w in won_leads[:8]:
        ago = f" · {w['days_ago']}d ago" if isinstance(w.get("days_ago"), int) else ""
        comp = f" · {esc(w['compound'])}" if w.get("compound") else ""
        tag = esc(w.get("stage") or "")
        won_rows.append(f"""<div style="border-left:3px solid #1e8e4e;padding:4px 10px;margin:4px 0;
          background:#f6fbf7;font-size:12.5px">
          <b>{esc(w.get('name'))}</b>
          <span style="background:#1e8e4e;color:#fff;border-radius:3px;padding:1px 6px;font-size:10px">{tag}</span>
          {('· <a href="tel:'+esc(w['phone'])+'" style="color:#2980b9;text-decoration:none">'+esc(w['phone'])+'</a>') if w.get('phone') else ''}
          <span style="color:#666">{comp}{ago}</span></div>""")
    won_html = ('<div style="font-size:12px;color:#1e8e4e;font-weight:700;margin:6px 0 2px">'
                '✅ Converted (paid / signed):</div>' + "".join(won_rows)) if won_rows else ""

    return (f"""<h3 style="margin:18px 0 6px;font-size:15px">🎯 Lead → contract conversion
      <span style="font-size:11px;color:#999;font-weight:400">— from the sales team's Stage
      column ({fmt(c.get('total'))} leads)</span></h3>{scoreboard}
      <div style="font-size:12px;color:#555;font-weight:600;margin:4px 0 2px">Reception funnel:</div>
      {funnel_html}{won_html}
      <div style="font-size:11px;color:#999;margin-top:5px">
        Conversion = leads at Stage <b>Converted</b> or <b>Paid for Measurement</b> ÷ qualified
        leads (junk / N-A / international / Designy excluded). Status is whatever the team logs
        in the AHD leads sheet.</div>""")


# Funnel ladder (ascending) → (progress %, short label, emoji). The furthest TRUE
# checkbox on a client row is their real position, so we read `stage` (which the
# roll-up fills for grouped clients like Farida) as the PRIMARY state — not the
# free-text STATUS column, which is often blank on a rolled-up parent row.
_FUNNEL_STEPS = [
    ("SESSION", 17, "Session booked", "🗓️"),
    ("OFFER", 33, "Offer sent", "💬"),
    ("INITIAL PRESENTATION", 50, "Initial design", "🎨"),
    ("FINAL PRESENTATION AFTER CLIENT COMMENTS", 67, "Final design", "✏️"),
    ("CONTRACTED", 83, "CONTRACTED", "✍️"),
    ("ORDER", 100, "WON — in production", "🏆"),
]
_FUNNEL_PCT = {name: (pct, label, emoji) for name, pct, label, emoji in _FUNNEL_STEPS}


def _lead_journey_block(journey):
    """Leads matched by name to clients in the sales tracker — the closed loop:
    which Facebook leads actually became tracked clients, and where each stands
    now (the tracker is what sales keeps after a measurement)."""
    if not journey:
        return ('<h3 style="margin:18px 0 6px;font-size:15px">🔗 Lead → sales journey</h3>'
                '<div style="font-size:12px;color:#888">No lead names matched tracker clients '
                'yet (names must match closely to link safely).</div>')

    def _resolve(j):
        """Return (pct, label, emoji, colour, at_risk) for a matched lead.
        Primary state = furthest funnel stage reached; STATUS only flags risk."""
        stage = (j.get("stage") or "").upper().strip()
        status = (j.get("status") or "").upper()
        at_risk = "OVER BUDGET" in status or "NO ANSWER" in status
        pct, label, emoji = _FUNNEL_PCT.get(stage, (0, "", ""))
        if not label:                       # in tracker but no checkbox yet
            pct, label, emoji = 8, "New — in tracker", "🌱"
        # colour band by progress / risk
        if pct >= 83:
            colour = "#1e8e4e"              # green — contracted / won
        elif pct >= 50:
            colour = "#c9871f"              # amber — mid-funnel, designing
        else:
            colour = "#2980b9"             # blue — early
        if at_risk:
            colour = "#c0392b"             # red overrides — needs rescue
        return pct, label, emoji, colour, at_risk

    # ---- headline conversion scoreboard (motivational) ----
    n = len(journey)
    won = sum(1 for j in journey if (j.get("stage") or "").upper() in ("CONTRACTED", "ORDER"))
    in_play = sum(1 for j in journey
                  if (j.get("stage") or "").upper() in
                  ("OFFER", "INITIAL PRESENTATION",
                   "FINAL PRESENTATION AFTER CLIENT COMMENTS", "SESSION"))
    conv = round(100 * won / n) if n else 0

    def _amt_num(j):
        try:
            return float(str(j.get("amount") or "0").replace(",", "") or 0)
        except ValueError:
            return 0.0
    won_value = sum(_amt_num(j) for j in journey
                    if (j.get("stage") or "").upper() in ("CONTRACTED", "ORDER"))
    won_value_str = f"{won_value:,.0f}" if won_value else ""

    scoreboard = f"""
    <table role="presentation" width="100%" cellpadding="0" cellspacing="0"
      style="margin:6px 0 10px;border-collapse:separate;border-spacing:6px 0">
      <tr>
        <td width="33%" style="background:#e9f7ef;border:1px solid #b7e2c8;border-radius:8px;
          padding:10px 8px;text-align:center">
          <div style="font-size:24px;font-weight:800;color:#1e8e4e;line-height:1">{conv}%</div>
          <div style="font-size:10.5px;color:#3a6b4e;text-transform:uppercase;letter-spacing:.4px">
            Lead→contract</div></td>
        <td width="33%" style="background:#fdf3e2;border:1px solid #f0d9a8;border-radius:8px;
          padding:10px 8px;text-align:center">
          <div style="font-size:24px;font-weight:800;color:#c9871f;line-height:1">{won} <span
            style="font-size:13px;color:#8a6a2a">/ {n}</span></div>
          <div style="font-size:10.5px;color:#8a6a2a;text-transform:uppercase;letter-spacing:.4px">
            Leads won</div></td>
        <td width="33%" style="background:#eef4fb;border:1px solid #c3ddf3;border-radius:8px;
          padding:10px 8px;text-align:center">
          <div style="font-size:24px;font-weight:800;color:#2471b3;line-height:1">{in_play}</div>
          <div style="font-size:10.5px;color:#2c5a83;text-transform:uppercase;letter-spacing:.4px">
            Still in play</div></td>
      </tr>
    </table>
    {(f'<div style="font-size:12px;color:#1e8e4e;font-weight:600;margin:-4px 0 8px">'
      f'💰 {won_value_str} EGP already contracted from Facebook leads.</div>') if won_value_str else ''}"""

    rows = []
    for j in sorted(journey[:12], key=lambda x: -_resolve(x)[0]):
        pct, label, emoji, colour, at_risk = _resolve(j)
        amt = f" · <b>{esc(j['amount'])}</b>" if j.get("amount") else ""
        rep = f" · rep {esc(j['rep'])}" if j.get("rep") else ""
        ago = f"{j['lead_days_ago']}d ago" if isinstance(j.get("lead_days_ago"), int) else ""
        tier = j.get("tier") or ""
        risk_badge = ('<span style="background:#c0392b;color:#fff;border-radius:3px;'
                      'padding:1px 6px;font-size:10px;font-weight:700">⚠ '
                      + esc((j.get("status") or "").title()) + '</span>') if at_risk else ''
        # a slim progress bar so the funnel position reads at a glance
        bar = (f'<table role="presentation" width="100%" cellpadding="0" cellspacing="0" '
               f'style="margin:5px 0 2px"><tr>'
               f'<td style="background:#ececec;border-radius:5px;height:8px;padding:0">'
               f'<div style="background:{colour};width:{pct}%;height:8px;border-radius:5px">'
               f'</div></td>'
               f'<td width="42" style="text-align:right;font-size:11px;font-weight:700;'
               f'color:{colour};padding-left:8px">{pct}%</td></tr></table>')
        rows.append(f"""<div style="border-left:4px solid {colour};padding:7px 11px;
          margin:6px 0;background:#fcfcfa;border-radius:0 6px 6px 0;font-size:12.5px">
          <b style="font-size:13.5px">{esc(j.get('name') or j.get('client'))}</b>
          {('<span style="background:#666;color:#fff;border-radius:3px;padding:1px 6px;font-size:10px">'+esc(tier)+'</span>') if tier else ''}
          {risk_badge}
          {('· <a href="tel:'+esc(j['phone'])+'" style="color:#2980b9;text-decoration:none">'+esc(j['phone'])+'</a>') if j.get('phone') else ''}
          <span style="float:right;color:{colour};font-weight:700">{emoji} {esc(label)}</span>
          {bar}
          <span style="color:#555">Lead: {esc(j.get('interest') or '?')} ·
          {esc(j.get('when') or '?')} · {esc(j.get('compound') or 'no compound')}
          <span style="color:#aaa">{esc(ago)}</span></span>{amt}{rep}</div>""")
    return (f"""<h3 style="margin:18px 0 6px;font-size:15px">🔗 Lead → sales journey
      <span style="font-size:11px;color:#999;font-weight:400">— {n} Facebook lead(s)
      now tracked as clients</span></h3>{scoreboard}{''.join(rows)}
      <div style="font-size:11px;color:#999;margin-top:6px">
        Each bar shows how far that Facebook lead has moved through the sales funnel
        (Session → Offer → Design → <b>Contracted</b> → Won). Names matched between the
        lead-ads feed and the sales tracker; conservative matching — only confident links shown.</div>""")


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
    if not force:
        if already_sent_today():
            print("Already sent today — skipping (dedupe).")
            return 0
        hour = cairo_now().hour
        if not (SEND_HOUR_CAIRO <= hour < SEND_WINDOW_END_CAIRO):
            print(f"Cairo hour is {hour}, outside the send window "
                  f"{SEND_HOUR_CAIRO}:00–{SEND_WINDOW_END_CAIRO}:00 — skipping "
                  f"(use --force to override).")
            return 0
    send_email(subject, html_body, recipients)
    mark_sent_today()
    return 0


if __name__ == "__main__":
    sys.exit(main())
