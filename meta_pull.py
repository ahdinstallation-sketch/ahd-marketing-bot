#!/usr/bin/env python3
"""Pull Meta (Facebook/Instagram) ads + organic data via the Graph API for the AHD
group. Needs env var META_TOKEN (a long-lived / System User token with ads_read,
read_insights, pages_read_engagement, pages_read_user_content, business_management).

Returns a snapshot dict with, per ad account:
  - spend / impressions / clicks / CTR / CPC / leads / CPL for YESTERDAY and MTD
  - ad-level winners (low CPL, high CTR) and losers (high spend, poor CPL/CTR), last 7d
and per page:
  - top organic posts in the last 14d (to recommend boosting), with permalink

Importable: `from meta_pull import pull_meta` -> dict. Degrades gracefully: if a
call fails or the token is missing, the affected section is returned with an "error".
"""
import os, json, datetime, urllib.parse, urllib.request

GRAPH = "https://graph.facebook.com/v21.0"
TOKEN = os.environ.get("META_TOKEN", "").strip()

ACCOUNTS = [
    {"id": "211980975881105", "name": "Amr Helmy Designs", "currency": "EGP"},
    {"id": "2634184000109460", "name": "KORSAGY", "currency": "EGP"},
    {"id": "1009314990138389", "name": "AHD (USD)", "currency": "USD"},
    {"id": "1535046841131907", "name": "Designy (USD)", "currency": "USD"},
]
PAGES = [
    {"id": "133848720036614", "name": "Amr Helmy Designs"},
    {"id": "372039879482122", "name": "Designy"},
    {"id": "2690515210977549", "name": "Korsagy"},
]

# Lead-type actions we count as "leads" across objectives.
LEAD_ACTIONS = {"lead", "onsite_conversion.lead_grouped", "leadgen_grouped",
                "offsite_conversion.fb_pixel_lead", "onsite_conversion.messaging_conversation_started_7d"}


def _get(path, params):
    params = dict(params)
    params["access_token"] = TOKEN
    url = f"{GRAPH}/{path}?{urllib.parse.urlencode(params)}"
    try:
        with urllib.request.urlopen(url, timeout=60) as r:
            return json.loads(r.read().decode())
    except urllib.error.HTTPError as e:
        try:
            body = json.loads(e.read().decode())
            msg = body.get("error", {}).get("message", str(e))
        except Exception:
            msg = str(e)
        return {"error": msg}
    except Exception as e:
        return {"error": str(e)}


def _today_cairo():
    # Cairo is UTC+2 (or +3 under DST); date-only, so +3 is safe enough for "yesterday".
    return (datetime.datetime.utcnow() + datetime.timedelta(hours=3)).date()


def _dates():
    today = _today_cairo()
    yest = today - datetime.timedelta(days=1)
    first = today.replace(day=1)
    # MTD = 1st of month through yesterday (today is usually still accumulating).
    mtd_end = yest if yest >= first else first
    return {
        "yesterday": {"since": str(yest), "until": str(yest)},
        "mtd": {"since": str(first), "until": str(mtd_end)},
        "last7": {"since": str(today - datetime.timedelta(days=7)),
                  "until": str(yest)},
    }


def _f(d, k, default=0.0):
    try:
        return float(d.get(k, default))
    except (TypeError, ValueError):
        return default


def _leads_from_actions(row):
    n = 0.0
    for a in row.get("actions", []) or []:
        if a.get("action_type") in LEAD_ACTIONS:
            n += _f(a, "value")
    return n


def _insights(account_id, time_range, level="account"):
    fields = "spend,impressions,clicks,ctr,cpc,reach,actions,date_start,date_stop"
    if level != "account":
        fields = "ad_id,ad_name,campaign_name," + fields
    params = {
        "level": level,
        "fields": fields,
        "time_range": json.dumps(time_range),
        "limit": 200,
    }
    res = _get(f"act_{account_id}/insights", params)
    return res


def _account_block(acc):
    dates = _dates()
    out = {"name": acc["name"], "currency": acc["currency"], "id": acc["id"]}
    for key in ("yesterday", "mtd"):
        res = _insights(acc["id"], dates[key])
        if "error" in res:
            out[key] = {"error": res["error"]}
            continue
        rows = res.get("data", [])
        row = rows[0] if rows else {}
        spend = _f(row, "spend")
        clicks = _f(row, "clicks")
        leads = _leads_from_actions(row)
        out[key] = {
            "spend": round(spend, 2),
            "impressions": int(_f(row, "impressions")),
            "clicks": int(clicks),
            "ctr": round(_f(row, "ctr"), 2),
            "cpc": round(_f(row, "cpc"), 2),
            "reach": int(_f(row, "reach")),
            "leads": int(leads),
            "cpl": round(spend / leads, 2) if leads else None,
        }
    # Ad-level winners/losers over last 7d
    res = _insights(acc["id"], dates["last7"], level="ad")
    ads = []
    if "error" not in res:
        for r in res.get("data", []):
            spend = _f(r, "spend")
            leads = _leads_from_actions(r)
            ads.append({
                "ad_name": r.get("ad_name", ""),
                "campaign": r.get("campaign_name", ""),
                "spend": round(spend, 2),
                "ctr": round(_f(r, "ctr"), 2),
                "leads": int(leads),
                "cpl": round(spend / leads, 2) if leads else None,
            })
    # Losers: meaningful spend, no/expensive leads or weak CTR
    spends = sorted((a["spend"] for a in ads), reverse=True)
    spend_thresh = spends[min(len(spends) - 1, 4)] if spends else 0  # ~top-5 spenders
    losers = [a for a in ads if a["spend"] >= max(spend_thresh, 1)
              and (a["cpl"] is None or a["leads"] == 0 or a["ctr"] < 0.7)]
    losers = sorted(losers, key=lambda a: -a["spend"])[:5]
    winners = [a for a in ads if a["leads"] > 0 and a["cpl"] is not None]
    winners = sorted(winners, key=lambda a: a["cpl"])[:5]
    out["ad_losers"] = losers
    out["ad_winners"] = winners
    return out


def _page_top_posts(page):
    """Top organic posts in the last 14d, ranked by engaged users / reactions."""
    since = (_today_cairo() - datetime.timedelta(days=14)).isoformat()
    params = {
        "fields": ("message,created_time,permalink_url,"
                   "insights.metric(post_impressions,post_impressions_unique,"
                   "post_engaged_users,post_clicks,post_reactions_by_type_total)"),
        "since": since,
        "limit": 50,
    }
    res = _get(f"{page['id']}/posts", params)
    if "error" in res:
        return {"name": page["name"], "error": res["error"], "posts": []}
    posts = []
    for p in res.get("data", []):
        ins = {}
        for m in (p.get("insights", {}) or {}).get("data", []):
            vals = m.get("values", [{}])
            ins[m["name"]] = vals[0].get("value", 0) if vals else 0
        reactions = ins.get("post_reactions_by_type_total", {})
        react_total = sum(reactions.values()) if isinstance(reactions, dict) else 0
        engaged = ins.get("post_engaged_users", 0) or 0
        impr = ins.get("post_impressions", 0) or 0
        eng_rate = round(100.0 * engaged / impr, 1) if impr else 0.0
        posts.append({
            "msg": (p.get("message", "") or "")[:90],
            "created": p.get("created_time", "")[:10],
            "link": p.get("permalink_url", ""),
            "impressions": impr,
            "engaged": engaged,
            "reactions": react_total,
            "clicks": ins.get("post_clicks", 0) or 0,
            "eng_rate": eng_rate,
        })
    # Rank by engaged users then engagement rate; surface the strongest organic content.
    posts = sorted(posts, key=lambda x: (x["engaged"], x["eng_rate"]), reverse=True)[:5]
    return {"name": page["name"], "posts": posts}


def pull_meta():
    snap = {
        "pulled_at": datetime.datetime.now().isoformat(timespec="seconds"),
        "dates": _dates(),
        "accounts": [],
        "pages": [],
        "token_present": bool(TOKEN),
    }
    if not TOKEN:
        snap["error"] = "META_TOKEN not set — Meta sections skipped."
        return snap
    for acc in ACCOUNTS:
        snap["accounts"].append(_account_block(acc))
    for page in PAGES:
        snap["pages"].append(_page_top_posts(page))
    return snap


if __name__ == "__main__":
    print(json.dumps(pull_meta(), ensure_ascii=False, indent=1))
