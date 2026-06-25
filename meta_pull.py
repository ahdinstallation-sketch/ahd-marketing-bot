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
    {"id": "1535046841131907", "name": "Designy (USD)", "currency": "USD"},
    # Not yet assigned to the AHDSYSTEM system user (Graph API returns
    # "ad account owner has NOT granted ads_read"). Re-enable once added in
    # Business Settings → System Users → AHDSYSTEM → Add Assets → Ad Accounts:
    # {"id": "2634184000109460", "name": "KORSAGY", "currency": "EGP"},
    # {"id": "1009314990138389", "name": "AHD (USD)", "currency": "USD"},
]
PAGES = [
    {"id": "133848720036614", "name": "Amr Helmy Designs"},
    {"id": "372039879482122", "name": "Designy"},
    {"id": "2690515210977549", "name": "Korsagy"},
]

# Lead-type actions we count as "leads" across objectives.
LEAD_ACTIONS = {"lead", "onsite_conversion.lead_grouped", "leadgen_grouped",
                "offsite_conversion.fb_pixel_lead", "onsite_conversion.messaging_conversation_started_7d"}


def _get(path, params, token=None):
    params = dict(params)
    params["access_token"] = token or TOKEN
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
        # the 7 days BEFORE last7, for trend comparison
        "prev7": {"since": str(today - datetime.timedelta(days=14)),
                  "until": str(today - datetime.timedelta(days=8))},
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
    fields = "spend,impressions,clicks,ctr,cpc,reach,frequency,actions,date_start,date_stop"
    if level == "ad":
        fields = "ad_id,ad_name,campaign_name," + fields
    elif level == "campaign":
        fields = "campaign_id,campaign_name," + fields
    params = {
        "level": level,
        "fields": fields,
        "time_range": json.dumps(time_range),
        "limit": 200,
    }
    res = _get(f"act_{account_id}/insights", params)
    return res


def _delivery_health(account_id):
    """Count campaigns and ads by delivery status, so the report can flag what's
    actually running vs paused vs blocked (WITH_ISSUES / DISAPPROVED / etc.).
    Uses effective_status, the status Meta actually delivers on."""
    out = {"campaigns": {}, "ads": {}, "active_campaign_names": [],
           "issue_ads": []}
    # Campaigns
    res = _get(f"act_{account_id}/campaigns",
               {"fields": "name,effective_status", "limit": 500})
    if "error" not in res:
        for c in res.get("data", []):
            st = c.get("effective_status", "UNKNOWN")
            out["campaigns"][st] = out["campaigns"].get(st, 0) + 1
            if st == "ACTIVE":
                out["active_campaign_names"].append(c.get("name", ""))
    else:
        out["campaigns_error"] = res["error"]
    # Ads
    active_set = set(out["active_campaign_names"])
    res = _get(f"act_{account_id}/ads",
               {"fields": "name,effective_status,campaign{name}", "limit": 500})
    if "error" not in res:
        for a in res.get("data", []):
            st = a.get("effective_status", "UNKNOWN")
            out["ads"][st] = out["ads"].get(st, 0) + 1
            # Only flag blocked ads that sit INSIDE an active campaign — those are
            # the ones actually losing you live delivery. (Archived/disapproved ads
            # in long-paused campaigns are noise.)
            cname = (a.get("campaign") or {}).get("name", "")
            if st in ("WITH_ISSUES", "DISAPPROVED", "PENDING_REVIEW") \
                    and cname in active_set:
                out["issue_ads"].append({
                    "name": a.get("name", ""), "status": st, "campaign": cname,
                })
    else:
        out["ads_error"] = res["error"]
    return out


def _campaign_trend(account_id):
    """Per-campaign performance for last7, with the prior 7 days as a trend
    baseline (spend / leads / CPL direction). Media-buyer view of which
    campaigns are improving vs decaying."""
    dates = _dates()

    def by_campaign(time_range):
        res = _insights(account_id, time_range, level="campaign")
        out = {}
        if "error" in res:
            return out, res["error"]
        for r in res.get("data", []):
            cid = r.get("campaign_id") or r.get("campaign_name")
            spend = _f(r, "spend")
            leads = _leads_from_actions(r)
            out[cid] = {
                "name": r.get("campaign_name", ""),
                "spend": spend,
                "leads": leads,
                "ctr": _f(r, "ctr"),
                "frequency": _f(r, "frequency"),
                "cpl": (spend / leads) if leads else None,
            }
        return out, None

    cur, err = by_campaign(dates["last7"])
    if err:
        return {"error": err, "campaigns": []}
    prev, _ = by_campaign(dates["prev7"])
    rows = []
    for cid, c in cur.items():
        p = prev.get(cid, {})
        rows.append({
            "name": c["name"],
            "spend": round(c["spend"], 2),
            "leads": int(c["leads"]),
            "cpl": round(c["cpl"], 2) if c["cpl"] is not None else None,
            "ctr": round(c["ctr"], 2),
            "frequency": round(c["frequency"], 2),
            "prev_spend": round(p.get("spend", 0), 2),
            "prev_leads": int(p.get("leads", 0)),
            "prev_cpl": round(p["cpl"], 2) if p.get("cpl") is not None else None,
        })
    rows = sorted(rows, key=lambda r: -r["spend"])
    return {"campaigns": rows}


def _account_block(acc):
    dates = _dates()
    out = {"name": acc["name"], "currency": acc["currency"], "id": acc["id"]}
    for key in ("yesterday", "mtd", "last7"):
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
            "frequency": round(_f(row, "frequency"), 2),
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
    # Media-buyer view: delivery health + per-campaign 7d-vs-prior-7d trend.
    out["delivery"] = _delivery_health(acc["id"])
    out["campaign_trend"] = _campaign_trend(acc["id"])
    return out


def _page_token(page_id):
    """A System-User token can read a Page's public fields but NOT its /posts edge.
    Exchange it for a per-Page token (works with pages_read_engagement)."""
    res = _get(page_id, {"fields": "access_token"})
    return res.get("access_token")


# Rich set needs pages_read_user_content; shares-only works with pages_read_engagement.
_POST_FIELDS_RICH = ("message,created_time,permalink_url,shares,"
                     "reactions.summary(true).limit(0),"
                     "comments.summary(true).limit(0)")
_POST_FIELDS_SAFE = "message,created_time,permalink_url,shares"


def _page_top_posts(page):
    """Top organic posts in the last 14d, ranked by engagement.

    v21 deprecated the old post_impressions/post_engaged_users insight metrics, so
    we rank on the engagement signals the token can actually read: shares always
    (pages_read_engagement), plus reactions+comments when the token also has
    pages_read_user_content. Degrades gracefully to shares-only otherwise.
    """
    ptok = _page_token(page["id"])
    if not ptok:
        return {"name": page["name"],
                "error": "no page access token (assign page to system user)",
                "posts": []}
    since = (_today_cairo() - datetime.timedelta(days=14)).isoformat()
    base = {"since": since, "limit": 50}
    res = _get(f"{page['id']}/posts", dict(base, fields=_POST_FIELDS_RICH), token=ptok)
    rich = "error" not in res
    if not rich:  # usually missing pages_read_user_content -> fall back to shares
        res = _get(f"{page['id']}/posts", dict(base, fields=_POST_FIELDS_SAFE), token=ptok)
    if "error" in res:
        return {"name": page["name"], "error": res["error"], "posts": []}

    posts = []
    for p in res.get("data", []):
        shares = (p.get("shares") or {}).get("count", 0) or 0
        react = ((p.get("reactions") or {}).get("summary") or {}).get("total_count")
        comments = ((p.get("comments") or {}).get("summary") or {}).get("total_count")
        engagement = shares + (react or 0) + (comments or 0)
        posts.append({
            "msg": (p.get("message", "") or "")[:90],
            "created": p.get("created_time", "")[:10],
            "link": p.get("permalink_url", ""),
            "shares": shares,
            "reactions": react,        # None when scope absent
            "comments": comments,      # None when scope absent
            "engagement": engagement,
        })
    # Strongest organic content first; ties broken by recency.
    posts = sorted(posts, key=lambda x: (x["engagement"], x["created"]), reverse=True)[:5]
    return {"name": page["name"], "rich": rich, "posts": posts}


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
