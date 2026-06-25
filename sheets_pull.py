#!/usr/bin/env python3
"""Pull the 3 AHD Google Sheets (published CSV) and return a structured snapshot:
lead funnel, dispositions, pipeline value, showroom sessions/visits, and the
stuck/over-budget/no-answer deal lists. No external deps (stdlib only).

Importable: `from sheets_pull import pull_sheets` -> dict.
CLI: `python3 sheets_pull.py` prints the snapshot JSON.
"""
import os, csv, io, json, collections, datetime, re, subprocess

SHEETS = {
    "designy_leads": "1VwE90ugoXTI4_90tKXzo7PT2TXpXYxyci0Ml3g9rrqs",
    "outdoor_leads": "1dcpQoXDntz7qfL2rMPOQ8Ba2xfLDwVxg6Tr-vT4H3BI",
    "ahd_tracker":   "1DuEomzuYrvveXzvwg4hbj0tsGBQpjRk6iUpbOFm2FvQ",
}
TRACKER_GID = "57990844"

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

    def deal(r):
        return {
            "client": r[cl].strip() if cl is not None and len(r) > cl else "",
            "rep": r[sp].strip() if sp is not None and len(r) > sp else "",
            "amount": r[ai].strip() if ai is not None and len(r) > ai else "",
            "amount_num": num(r[ai]) if ai is not None and len(r) > ai else 0.0,
            "note": r[nt].strip() if nt is not None and len(r) > nt else "",
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
    return out


def _row_max_num(cells):
    vals = [num(c) for c in cells if num(c) > 0]
    return max(vals) if vals else None


def pull_cash_in():
    """Live cash-in-to-date per company from the Looker-Studio-backed sheet.
    Returns {"configured", "ahd", "designy", "as_of"}; degrades to unconfigured."""
    out = {"configured": bool(CASHIN_SHEET_ID), "ahd": None,
           "designy": None, "as_of": None}
    if not CASHIN_SHEET_ID:
        return out
    rows = fetch(CASHIN_SHEET_ID, CASHIN_GID)
    for r in rows:
        if not r:
            continue
        label = " ".join(c for c in r[:2]).strip().lower()
        rowmax = _row_max_num(r)
        if "ahd" in label and out["ahd"] is None and rowmax:
            out["ahd"] = rowmax
        elif "designy" in label and out["designy"] is None and rowmax:
            out["designy"] = rowmax
        elif out["as_of"] is None and any(
                k in label for k in ("as of", "as at", "updated", "refresh")):
            for c in r[1:]:
                # a date-like value: has a digit and a separator
                if c and any(ch.isdigit() for ch in c) and re.search(r"[-/]", c):
                    out["as_of"] = c.strip()
                    break
    return out


def pull_sheets():
    snap = {"pulled_at": datetime.datetime.now().isoformat(timespec="seconds")}
    snap["designy_leads"] = analyze_leads(fetch(SHEETS["designy_leads"]))
    snap["outdoor_leads"] = analyze_leads(fetch(SHEETS["outdoor_leads"]))
    # Default sheet matches the tracker tab the prior session validated (header on row 2).
    snap["ahd_tracker"] = analyze_tracker(fetch(SHEETS["ahd_tracker"]))
    snap["cash_in"] = pull_cash_in()
    return snap


if __name__ == "__main__":
    print(json.dumps(pull_sheets(), ensure_ascii=False, indent=1))
