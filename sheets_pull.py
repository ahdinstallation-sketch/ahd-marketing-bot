#!/usr/bin/env python3
"""Pull the 3 AHD Google Sheets (published CSV) and return a structured snapshot:
lead funnel, dispositions, pipeline value, showroom sessions/visits, and the
stuck/over-budget/no-answer deal lists. No external deps (stdlib only).

Importable: `from sheets_pull import pull_sheets` -> dict.
CLI: `python3 sheets_pull.py` prints the snapshot JSON.
"""
import os, csv, io, json, collections, datetime, re, subprocess, unicodedata

SHEETS = {
    "designy_leads": "1VwE90ugoXTI4_90tKXzo7PT2TXpXYxyci0Ml3g9rrqs",
    "outdoor_leads": "1dcpQoXDntz7qfL2rMPOQ8Ba2xfLDwVxg6Tr-vT4H3BI",
    "ahd_tracker":   "1DuEomzuYrvveXzvwg4hbj0tsGBQpjRk6iUpbOFm2FvQ",
}
TRACKER_GID = "57990844"
# The AHD Facebook Lead-Ads feed (same workbook as outdoor_leads, specific tab).
AHD_LEADS_GID = "1977817316"

# Live "cash in to date" source — the Google Sheet that backs the Looker Studio
# treasury dashboard (auto-updates). Set CASHIN_SHEET_ID (and optionally
# CASHIN_GID for a specific tab) as an env var / GitHub secret. Expected layout:
# a simple label/value table, one row per company, e.g.
#     AHD cash in to date , 5,430,000
#     Designy cash in to date , 1,330,000
#     As of , 2026-06-25
# Label matching is fuzzy (looks for "ahd"/"designy" + the largest number in the
# row), so column order is flexible. Degrades gracefully if unset or unreachable.
CASHIN_SHEET_ID = os.environ.get("CASHIN_SHEET_ID", "").strip()
CASHIN_GID = os.environ.get("CASHIN_GID", "").strip() or None
DISPO_KEYS = ("bad lead", "good lead", "paid measurement", "no answer",
              "over budget", "signed", "visit", "low budget", "early stage",
              "converted", "contacted")


def fetch(idn, gid=None, retries=4):
    import time
    url = f"https://docs.google.com/spreadsheets/d/{idn}/export?format=csv"
    if gid:
        url += f"&gid={gid}"
    last = ""
    for _ in range(retries):
        raw = subprocess.run(["curl", "-sL", "--max-time", "40", url],
                             capture_output=True, text=True).stdout
        if raw and len(raw) > 50 and "<html" not in raw[:200].lower():
            return list(csv.reader(io.StringIO(raw)))
        last = raw
        time.sleep(2)
    return list(csv.reader(io.StringIO(last)))


def num(s):
    s = re.sub(r"[^0-9.]", "", s or "")
    try:
        return float(s) if s else 0.0
    except ValueError:
        return 0.0


def analyze_leads(rows):
    if not rows:
        return {}
    hdr = rows[0]
    data = [r for r in rows[1:] if any(r) and r and r[0].strip()]
    out = {"total": len(data)}
    if "created_time" in hdr:
        ci = hdr.index("created_time")
        ds = []
        for r in data:
            if len(r) > ci and r[ci]:
                try:
                    ds.append(datetime.date.fromisoformat(r[ci][:10]))
                except ValueError:
                    pass
        if ds:
            today = datetime.date.today()
            out["date_min"] = str(min(ds))
            out["date_max"] = str(max(ds))
            out["yesterday"] = sum(1 for d in ds if (today - d).days == 1)
            out["last7"] = sum(1 for d in ds if 0 <= (today - d).days <= 7)
            out["last30"] = sum(1 for d in ds if 0 <= (today - d).days <= 30)
    # disposition: column with the most known status keywords
    best_ci, best_hits = None, 0
    for ci in range(len(hdr)):
        hits = sum(1 for r in data if len(r) > ci and r[ci]
                   and any(k in r[ci].strip().lower() for k in DISPO_KEYS))
        if hits > best_hits:
            best_ci, best_hits = ci, hits
    if best_ci is not None:
        c = collections.Counter()
        for r in data:
            v = (r[best_ci].strip() if len(r) > best_ci and r[best_ci].strip()
                 else "(no disposition)")
            c[v] += 1
        out["disposition"] = dict(c.most_common(12))
    return out


def _tracker_date(s):
    """Parse the tracker's day/month/year date cells; None if unparseable."""
    s = (s or "").strip()
    for f in ("%d/%m/%Y", "%d/%m/%y", "%Y-%m-%d"):
        try:
            return datetime.datetime.strptime(s, f).date()
        except ValueError:
            pass
    return None


def analyze_tracker(rows):
    if not rows or len(rows) < 3:
        return {}
    hdr = rows[1]
    data = [r for r in rows[2:] if any(r) and r and r[0].strip()]

    def idx(n):
        for i, h in enumerate(hdr):
            if h.strip() == n:
                return i
        return None

    def tcount(name):
        i = idx(name)
        return sum(1 for r in data if i is not None and len(r) > i
                   and r[i].strip().upper() == "TRUE")

    out = {"clients": len(data), "funnel": {}}
    for stg in ["SESSION", "OFFER", "RHINO PRESENTATION",
                "PRESENTATION SENT ON EXT", "CONTRACTED", "ORDER"]:
        out["funnel"][stg] = tcount(stg)
    si = idx("STATUS")
    if si is not None:
        c = collections.Counter()
        for r in data:
            v = (r[si].strip() if len(r) > si and r[si].strip() else "(blank)")
            c[v] += 1
        out["status"] = dict(c.most_common(20))
    ai = idx("Amount in EGP")
    ci = idx("CONTRACTED")
    oi = idx("ORDER")
    signed = openval = 0.0
    for r in data:
        if ai is None or len(r) <= ai:
            continue
        v = num(r[ai])
        st = (r[si].strip().upper() if si is not None and len(r) > si else "")
        contr = (len(r) > ci and r[ci].strip().upper() == "TRUE") if ci is not None else False
        if st == "SIGNED CONTRACT" or contr:
            signed += v
        elif v > 0 and st in ("DESIGN DISCUSSION", "AWAITING PRESENTATION", ""):
            openval += v
    out["signed_value"] = round(signed)
    out["open_pipeline_value"] = round(openval)

    cl = idx("CLIENT")
    sp = idx("SALES PERSON")
    nt = idx("Notes")
    # Date columns used to age a deal (most recent available wins).
    date_idx = [idx(n) for n in ("OFFER DATE", "SESSION DATE", "Call Date",
                                 "Finalize Date")]
    date_idx = [i for i in date_idx if i is not None]
    today = datetime.date.today()

    def deal(r):
        ds = [_tracker_date(r[i]) for i in date_idx if len(r) > i]
        ds = [d for d in ds if d]
        bd = max(ds) if ds else None
        return {
            "client": r[cl].strip() if cl is not None and len(r) > cl else "",
            "rep": r[sp].strip() if sp is not None and len(r) > sp else "",
            "amount": r[ai].strip() if ai is not None and len(r) > ai else "",
            "amount_num": num(r[ai]) if ai is not None and len(r) > ai else 0.0,
            "note": r[nt].strip() if nt is not None and len(r) > nt else "",
            "date": bd.isoformat() if bd else None,
            "days_ago": (today - bd).days if bd else None,
        }

    stuck = []
    for r in data:
        if ci is None or oi is None:
            break
        contr = len(r) > ci and r[ci].strip().upper() == "TRUE"
        order = len(r) > oi and r[oi].strip().upper() == "TRUE"
        if contr and not order:
            stuck.append(deal(r))
    out["contracted_no_order"] = sorted(stuck, key=lambda d: -d["amount_num"])

    def status_list(want):
        L = [deal(r) for r in data
             if si is not None and len(r) > si and r[si].strip().upper() == want]
        return sorted(L, key=lambda d: -d["amount_num"])

    out["no_answer_after_offer"] = status_list("NO ANSWER AFTER OFFER")
    out["over_budget"] = status_list("OVER BUDGET")

    # Per-client detail (with the furthest funnel stage reached) for correlating
    # lead names against the tracker the sales team keeps post-measurement.
    stage_order = [("ORDER", oi), ("CONTRACTED", ci),
                   ("PRESENTATION SENT ON EXT", idx("PRESENTATION SENT ON EXT")),
                   ("RHINO PRESENTATION", idx("RHINO PRESENTATION")),
                   ("OFFER", idx("OFFER")), ("SESSION", idx("SESSION"))]

    def stage_of(r):
        for name, i in stage_order:
            if i is not None and len(r) > i and r[i].strip().upper() == "TRUE":
                return name
        return ""

    clients_detail = []
    for r in data:
        if cl is None or len(r) <= cl or not r[cl].strip():
            continue
        rec = deal(r)
        rec["status"] = (r[si].strip() if si is not None and len(r) > si else "")
        rec["stage"] = stage_of(r)
        clients_detail.append(rec)
    out["clients_detail"] = clients_detail
    return out


# ---- AHD lead-ads feed (positional parse + brand-fit scoring) ----
# The Facebook Lead-Ads export sheet has NO header on its data columns; the
# column legend lives off to the right (cols 19-37). Data is positional, cols 0-18,
# newest lead first. Map (validated on live data, 304 leads):
_LEAD_COLS = {
    "id": 0, "created_time": 1, "ad_name": 3, "adset_name": 5,
    "campaign_name": 7, "form_name": 9, "is_organic": 10, "platform": 11,
    "interest": 12, "when_planning": 13, "compound": 14, "full_name": 15,
    "phone": 16, "lead_status": 18,
}
# Premium Cairo compounds (brand fit for AHD's high-end kitchens/wardrobes).
_PREMIUM_COMPOUNDS = ("palm hills", "new giza", "mivida", "mountain view",
                      "madinaty", "maadi", "new cairo", "zayed", "الشيخ زايد",
                      "zed", "katameya", "uptown", "hyde park", "swan lake")
_URGENCY_SCORE = {"immediately": 3, "within_3_months": 2, "within_3_month": 2,
                  "3-6_months": 1, "3–6_months": 1, "3_6_months": 1,
                  "exploring_options": 0, "exploring": 0}
_BIG_PROJECT = ("full_home", "full home", "multiple_spaces", "multiple spaces")
_SMALL_PROJECT = ("kitchen", "wardrobe", "dressing", "closet")


def _norm(s):
    return (s or "").strip().lower()


def _lead_date(s):
    """Lead-ads created_time is ISO with offset, e.g. 2026-06-23T05:01:13-05:00."""
    s = (s or "").strip()
    if not s:
        return None
    try:
        return datetime.date.fromisoformat(s[:10])
    except ValueError:
        return None


def _score_lead(lead):
    """Brand-fit score from first-party form answers (no online profiling):
    urgency + project size + premium compound + contactability. Tier A/B/C."""
    sc, why = 0, []
    u = _URGENCY_SCORE.get(_norm(lead["when_planning"]))
    if u is not None:
        sc += u
        if u >= 2:
            why.append("ready to buy soon")
    interest = _norm(lead["interest"])
    if any(k in interest for k in _BIG_PROJECT):
        sc += 2
        why.append("whole-home / multi-room")
    elif any(k in interest for k in _SMALL_PROJECT):
        sc += 1
    comp = _norm(lead["compound"])
    if any(k in comp for k in _PREMIUM_COMPOUNDS):
        sc += 2
        why.append("premium compound")
    elif comp:
        sc += 1
    # contactability: a real phone present
    if re.search(r"\d{7,}", lead.get("phone", "") or ""):
        sc += 1
    else:
        why.append("no phone")
    tier = "A" if sc >= 6 else ("B" if sc >= 4 else "C")
    return sc, tier, why


def analyze_ahd_leads(rows, gid_note=""):
    """Parse the AHD Facebook Lead-Ads sheet (positional) and score each lead for
    brand fit. Returns aggregate breakdowns + the best-fit (Tier A/B) leads with
    name+phone so the team can act per lead."""
    if not rows:
        return {"error": "no rows", "leads": []}
    out = {"total": 0, "by_platform": {}, "by_interest": {}, "by_urgency": {},
           "by_campaign": {}, "by_compound": {}, "tiers": {"A": 0, "B": 0, "C": 0},
           "yesterday": 0, "last7": 0, "last30": 0, "top_leads": [],
           "unworked": 0}
    today = datetime.date.today()
    leads = []
    for r in rows:
        # A valid lead row starts with a Facebook lead id (l:...) or a long digit id.
        idv = r[_LEAD_COLS["id"]].strip() if len(r) > _LEAD_COLS["id"] else ""
        if not idv or not (idv.lower().startswith("l:") or idv.isdigit()):
            continue

        def cell(key):
            i = _LEAD_COLS[key]
            return r[i].strip() if len(r) > i and r[i] is not None else ""

        lead = {k: cell(k) for k in _LEAD_COLS}
        # Strip the Lead-Ads field prefixes (phone "p:+201…", id "l:…").
        lead["phone"] = re.sub(r"^p:\s*", "", lead["phone"]).strip()
        d = _lead_date(lead["created_time"])
        sc, tier, why = _score_lead(lead)
        lead["score"], lead["tier"], lead["why"] = sc, tier, why
        lead["date"] = d.isoformat() if d else None
        lead["days_ago"] = (today - d).days if d else None
        leads.append(lead)

        out["total"] += 1
        out["tiers"][tier] += 1
        plat = _norm(lead["platform"]) or "(unknown)"
        out["by_platform"][plat] = out["by_platform"].get(plat, 0) + 1
        intr = lead["interest"].strip() or "(blank)"
        out["by_interest"][intr] = out["by_interest"].get(intr, 0) + 1
        urg = lead["when_planning"].strip() or "(blank)"
        out["by_urgency"][urg] = out["by_urgency"].get(urg, 0) + 1
        camp = lead["campaign_name"].strip() or "(blank)"
        out["by_campaign"][camp] = out["by_campaign"].get(camp, 0) + 1
        comp = lead["compound"].strip() or "(blank)"
        out["by_compound"][comp] = out["by_compound"].get(comp, 0) + 1
        if _norm(lead["lead_status"]) in ("", "created"):
            out["unworked"] += 1
        if d:
            ago = (today - d).days
            if ago == 1:
                out["yesterday"] += 1
            if 0 <= ago <= 7:
                out["last7"] += 1
            if 0 <= ago <= 30:
                out["last30"] += 1

    # Sort helper for the breakdown dicts (most common first, trimmed).
    def topn(d, n=8):
        return dict(sorted(d.items(), key=lambda kv: -kv[1])[:n])
    out["by_platform"] = topn(out["by_platform"])
    out["by_interest"] = topn(out["by_interest"])
    out["by_urgency"] = topn(out["by_urgency"])
    out["by_campaign"] = topn(out["by_campaign"], 6)
    out["by_compound"] = topn(out["by_compound"], 10)

    # Best-fit leads to act on first: highest score, then most recent. Prefer
    # the last 14 days (actionable now) but fall back to overall if the feed is old.
    recent = [l for l in leads if l["days_ago"] is not None and l["days_ago"] <= 14]
    pool = recent if recent else leads
    pool = sorted(pool, key=lambda l: (l["score"],
                  -(l["days_ago"] if l["days_ago"] is not None else 9999)),
                  reverse=True)
    out["top_leads"] = [{
        "name": l["full_name"], "phone": l["phone"], "tier": l["tier"],
        "score": l["score"], "interest": l["interest"],
        "when": l["when_planning"], "compound": l["compound"],
        "platform": l["platform"], "campaign": l["campaign_name"],
        "days_ago": l["days_ago"], "why": l["why"],
    } for l in pool[:12]]
    out["recent_window"] = bool(recent)
    # Compact full list, used to correlate lead names against the sales tracker.
    out["leads_detail"] = [{
        "name": l["full_name"], "phone": l["phone"], "date": l["date"],
        "days_ago": l["days_ago"], "campaign": l["campaign_name"],
        "interest": l["interest"], "when": l["when_planning"],
        "compound": l["compound"], "platform": l["platform"],
        "tier": l["tier"], "score": l["score"],
    } for l in leads]
    return out


def _find_col(hdr, *names):
    for i, h in enumerate(hdr):
        hl = (h or "").strip().lower()
        if any(hl == n.lower() for n in names):
            return i
    return None


def pull_cash_in():
    """Live treasury snapshot from the Looker-Studio-backed sheet (the same data
    feeding the cash dashboard). The sheet is a GROUP-LEVEL monthly time series
    (no per-company split), so we surface: latest month's cash-in / cash-out /
    net, the bank balance and total debt, and the cumulative cash-in YTD.

    Returns {"configured", "ok", "month_label", "cash_in_month", "cash_out_month",
    "net_month", "bank_balance", "total_debt", "cash_in_ytd", "ytd_year", "as_of"}.
    Degrades gracefully (configured=False or ok=False) so the section just hides."""
    out = {"configured": bool(CASHIN_SHEET_ID), "ok": False}
    if not CASHIN_SHEET_ID:
        return out
    rows = fetch(CASHIN_SHEET_ID, CASHIN_GID)
    if not rows:
        return out
    # Locate the header row (the one that names the Cash In column).
    hdr_i = next((i for i, r in enumerate(rows[:5])
                  if any((c or "").strip().lower() == "cash in" for c in r)), 0)
    hdr = rows[hdr_i]
    c_ci = _find_col(hdr, "Cash In")
    if c_ci is None:
        return out
    c_co = _find_col(hdr, "Cash Out")
    c_bal = _find_col(hdr, "Balance")
    c_bank = _find_col(hdr, "Bank Balance including FX", "Bank Balance EGP")
    c_debt = _find_col(hdr, "Total Debt EGP", "Total Debt")
    c_yr = _find_col(hdr, "Year")
    c_mo = _find_col(hdr, "Month")
    c_dt = _find_col(hdr, "Date")

    def cell(r, i):
        return r[i] if (i is not None and len(r) > i) else ""

    latest, ytd = None, {}
    for r in rows[hdr_i + 1:]:
        v = num(cell(r, c_ci))
        if v <= 0:
            continue
        latest = r
        yr = (cell(r, c_yr) or "").strip()
        ytd[yr] = ytd.get(yr, 0.0) + v
    if latest is None:
        return out
    yr = (cell(latest, c_yr) or "").strip()
    mo = (cell(latest, c_mo) or "").strip()
    out.update({
        "ok": True,
        "month_label": f"{mo} {yr}".strip(),
        "as_of": (cell(latest, c_dt) or "").strip() or f"{mo} {yr}".strip(),
        "cash_in_month": round(num(cell(latest, c_ci))),
        "cash_out_month": round(num(cell(latest, c_co))) or None,
        "net_month": round(num(cell(latest, c_bal))) if cell(latest, c_bal) else None,
        "bank_balance": round(num(cell(latest, c_bank))) or None,
        "total_debt": round(num(cell(latest, c_debt))) or None,
        "cash_in_ytd": round(ytd.get(yr, 0.0)) or None,
        "ytd_year": yr or None,
    })
    return out


# ---- Lead ↔ tracker name correlation (close the marketing→sales loop) ----
# The sales team logs every client in the tracker AFTER a measurement/session, so
# matching a Facebook lead's name to a tracker client tells us which ad leads
# actually progressed — and where each one now stands.
_NAME_TITLES = {"dr", "eng", "mr", "mrs", "ms", "prof", "engineer", "arch", "m"}
# Ultra-common Egyptian name particles/first-names — too generic to confirm a
# match on their own, so they don't count as the "distinctive" shared token.
_NAME_COMMON = {"el", "al", "abd", "abdel", "abdul", "abo", "abou", "mohamed",
                "mohammed", "mohamad", "ahmed", "mahmoud", "ali"}


def _name_tokens(s):
    s = (s or "").replace("\xa0", " ")
    s = unicodedata.normalize("NFKD", s).lower()
    s = re.sub(r"[^a-z؀-ۿ\s]", " ", s)  # keep latin + arabic letters
    return [t for t in s.split() if t and t not in _NAME_TITLES]


def _name_match(a, b):
    """Conservative full-name match: identical token sets, OR same first name
    plus at least one distinctive (non-common) shared token (surname). Tuned to
    avoid false positives from shared common first names / family surnames."""
    ta, tb = _name_tokens(a), _name_tokens(b)
    if not ta or not tb:
        return False
    if ta == tb:
        return True
    if ta[0] != tb[0]:
        return False
    distinct = (set(ta) & set(tb)) - _NAME_COMMON - {ta[0]}
    return len(distinct) >= 1


def correlate_leads_tracker(leads_detail, clients_detail):
    """Match scored leads to tracker clients by name. Returns one record per
    matched lead with both the marketing side (campaign, tier, interest, when)
    and the current sales side (status, stage, amount, rep)."""
    if not leads_detail or not clients_detail:
        return []
    # Pre-tokenize clients once (the inner loop runs leads × clients).
    ct = [(c, _name_tokens(c.get("client", ""))) for c in clients_detail]
    out = []
    for l in leads_detail:
        lt = _name_tokens(l.get("name", ""))
        if len(lt) < 2:            # need at least first+last to match safely
            continue
        for c, ctok in ct:
            if not ctok:
                continue
            if lt == ctok or (lt[0] == ctok[0]
                              and (set(lt) & set(ctok)) - _NAME_COMMON - {lt[0]}):
                out.append({
                    "name": l.get("name"), "phone": l.get("phone"),
                    "tier": l.get("tier"), "interest": l.get("interest"),
                    "when": l.get("when"), "compound": l.get("compound"),
                    "campaign": l.get("campaign"), "lead_date": l.get("date"),
                    "lead_days_ago": l.get("days_ago"),
                    "client": c.get("client"), "status": c.get("status") or "",
                    "stage": c.get("stage") or "", "amount": c.get("amount") or "",
                    "rep": c.get("rep") or "",
                })
                break
    # Most recently-generated leads first.
    out.sort(key=lambda d: (d["lead_days_ago"] is None,
                            d["lead_days_ago"] if d["lead_days_ago"] is not None else 0))
    return out


def pull_sheets():
    snap = {"pulled_at": datetime.datetime.now().isoformat(timespec="seconds")}
    snap["designy_leads"] = analyze_leads(fetch(SHEETS["designy_leads"]))
    snap["outdoor_leads"] = analyze_leads(fetch(SHEETS["outdoor_leads"]))
    # Same workbook, the Facebook Lead-Ads tab — parsed positionally + brand-scored.
    snap["ahd_leads"] = analyze_ahd_leads(fetch(SHEETS["outdoor_leads"], AHD_LEADS_GID))
    # Default sheet matches the tracker tab the prior session validated (header on row 2).
    snap["ahd_tracker"] = analyze_tracker(fetch(SHEETS["ahd_tracker"]))
    # Close the loop: which Facebook leads are now clients in the sales tracker.
    snap["lead_journey"] = correlate_leads_tracker(
        snap["ahd_leads"].get("leads_detail", []),
        snap["ahd_tracker"].get("clients_detail", []))
    # Drop the bulky working lists (full names/phones) now that matching is done.
    snap["ahd_leads"].pop("leads_detail", None)
    snap["ahd_tracker"].pop("clients_detail", None)
    snap["cash_in"] = pull_cash_in()
    return snap


if __name__ == "__main__":
    print(json.dumps(pull_sheets(), ensure_ascii=False, indent=1))
