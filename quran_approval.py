#!/usr/bin/env python3
"""Quran With My Child — publish-by-email-reply.

The episode pipeline (on Ahmed's Mac) uploads every video PRIVATE. This script
is the only thing that ever makes one public, and it only does so when Ahmed
replies APPROVE to the approval email. It runs on GitHub Actions so it works
while the Mac is closed.

    python3 quran_approval.py request --video-id ID --title "Az-Zalzalah" [--notes "..."]
    python3 quran_approval.py poll        # read replies, publish approved videos
    python3 quran_approval.py poll --dry-run

A reply is acted on ONLY if all of these hold:
  1. it is unread in the bot mailbox and its subject carries a QWMC-REF tag,
  2. the tag's HMAC matches the video id (so only someone who received the
     approval email can produce it — a guessed video id is not enough),
  3. the sender is in QWMC_APPROVERS,
  4. Gmail's Authentication-Results show dkim=pass or dmarc=pass (not spoofed),
  5. the first line of the reply (above the quoted text) says approve/yes.

stdlib only. Env: MAIL_USER, MAIL_PASSWORD, YOUTUBE_TOKEN_JSON,
QWMC_APPROVAL_KEY, QWMC_APPROVERS, QWMC_NOTIFY.
"""
from __future__ import annotations

import argparse
import email
import email.utils
import hashlib
import hmac
import imaplib
import json
import os
import re
import smtplib
import sys
import urllib.parse
import urllib.request
from email.header import decode_header, make_header
from email.message import EmailMessage

TAG = "QWMC-REF"
REF_RE = re.compile(TAG + r":([A-Za-z0-9_-]{11})\.([0-9a-f]{12})")
APPROVE = ("approve", "approved", "yes", "publish", "ok", "okay", "go",
           "موافق", "نعم", "انشر", "تمام")
REJECT = ("reject", "rejected", "no", "don't", "dont", "stop", "لا", "ارفض")


def env(name: str) -> str:
    v = os.environ.get(name, "").strip()
    if not v:
        sys.exit(f"missing env {name}")
    return v


def ref_for(video_id: str) -> str:
    sig = hmac.new(env("QWMC_APPROVAL_KEY").encode(), video_id.encode(),
                   hashlib.sha256).hexdigest()[:12]
    return f"{TAG}:{video_id}.{sig}"


def approvers() -> set[str]:
    return {a.strip().lower() for a in env("QWMC_APPROVERS").split(",") if a.strip()}


# ---------------------------------------------------------------- mail
def _hdr(v: str) -> str:
    """Flatten a header value onto one line.

    A Message-ID pulled off an incoming mail can arrive folded across lines.
    Assigning that straight into In-Reply-To raises
    "Header values may not contain linefeed or carriage return characters" —
    which crashed the poll AFTER it had already published the video, so the
    run reported failure for an episode that had actually gone public.
    """
    return " ".join(str(v).split())


def send(to: str, subject: str, body: str, in_reply_to: str | None = None):
    msg = EmailMessage()
    msg["From"] = f"Quran With My Child bot <{env('MAIL_USER')}>"
    msg["To"] = _hdr(to)
    msg["Subject"] = _hdr(subject)
    if in_reply_to:
        ref = _hdr(in_reply_to)
        if ref:
            msg["In-Reply-To"] = ref
            msg["References"] = ref
    msg.set_content(body)
    with smtplib.SMTP_SSL("smtp.gmail.com", 465, timeout=60) as s:
        s.login(env("MAIL_USER"), env("MAIL_PASSWORD"))
        s.send_message(msg)


def first_line(msg) -> str:
    part = None
    if msg.is_multipart():
        for p in msg.walk():
            if p.get_content_type() == "text/plain" and not p.get_filename():
                part = p
                break
    else:
        part = msg
    if part is None:
        return ""
    text = part.get_payload(decode=True) or b""
    text = text.decode(part.get_content_charset() or "utf-8", "replace")
    for line in text.splitlines():
        s = line.strip()
        if not s:
            continue
        if s.startswith(">") or re.match(r"^(On .+ wrote:|From:|-----)", s):
            return ""
        return s.lower()
    return ""


def reply_text(msg) -> str:
    """Everything Ahmed typed, stopping at the quoted original."""
    part = None
    for p in (msg.walk() if msg.is_multipart() else [msg]):
        if p.get_content_type() == "text/plain" and not p.get_filename():
            part = p
            break
    if part is None:
        return ""
    raw = part.get_payload(decode=True) or b""
    text = raw.decode(part.get_content_charset() or "utf-8", "replace")
    out = []
    for line in text.splitlines():
        s = line.strip()
        if s.startswith(">") or re.match(r"^(On .+ wrote:|From:|-----|Sent from)", s):
            break
        out.append(line.rstrip())
    return "\n".join(out).strip()[:4000]


def verdict(line: str) -> str | None:
    words = re.findall(r"[\w']+", line)
    if not words:
        return None
    w = words[0]
    if w in REJECT:
        return "reject"
    if w in APPROVE:
        return "approve"
    return None


def authenticated(msg) -> bool:
    ar = " ".join(msg.get_all("Authentication-Results", []) or []).lower()
    return "dkim=pass" in ar or "dmarc=pass" in ar


# ---------------------------------------------------------------- quran-kids repo
QK = "/tmp/qk"


def qk_repo() -> str | None:
    """Shallow clone of the private quran-kids repo, via a deploy key scoped to
    that one repo (secret QK_DEPLOY_KEY). Only costs.jsonl and amendments/ are
    fetched. Returns the path, or None if the key is missing."""
    import subprocess
    key = os.environ.get("QK_DEPLOY_KEY", "")
    if not key:
        return None
    if os.path.isdir(os.path.join(QK, ".git")):
        return QK
    kf = "/tmp/qk_key"
    with open(kf, "w") as fh:
        fh.write(key if key.endswith("\n") else key + "\n")
    os.chmod(kf, 0o600)
    os.environ["GIT_SSH_COMMAND"] = f"ssh -i {kf} -o StrictHostKeyChecking=accept-new"
    subprocess.run(["git", "clone", "-q", "--depth", "1", "--filter=blob:none", "--sparse",
                    "git@github.com:ahdinstallation-sketch/quran-kids.git", QK], check=True)
    subprocess.run(["git", "-C", QK, "sparse-checkout", "set", "--no-cone",
                    "/costs.jsonl", "/published.json", "/amendments/"], check=True)
    return QK


def _qk_commit(repo: str, paths: list[str], msg: str) -> bool:
    """Commit and push inside the quran-kids clone, rebasing on contention."""
    import subprocess
    ident = ["-c", "user.name=quran-approval-bot",
             "-c", "user.email=actions@users.noreply.github.com"]
    g = ["git", "-C", repo]
    subprocess.run(g + ident + ["add"] + paths, check=True)
    if subprocess.run(g + ["diff", "--cached", "--quiet"]).returncode == 0:
        return True                                  # nothing changed
    subprocess.run(g + ident + ["commit", "-q", "-m", msg], check=True)
    for _ in range(3):
        if subprocess.run(g + ["push", "-q", "origin", "HEAD:main"]).returncode == 0:
            return True
        subprocess.run(g + ["pull", "-q", "--rebase", "origin", "main"])
    return False


def mark_published(video_id: str, dry_run: bool) -> None:
    """Record in the quran-kids ledger that this video is now public.

    Without this the ledger keeps saying "private" forever: publishing changed
    YouTube and nothing wrote back, so published.json disagreed with the
    channel and anyone reading it got a wrong picture of what was live.
    Best-effort — the video is already public by the time this runs, so a
    failure here must never raise.
    """
    import json as _json
    try:
        repo = qk_repo()
        if not repo:
            return
        lp = os.path.join(repo, "published.json")
        if not os.path.exists(lp):
            return
        led = _json.load(open(lp))
        hit = False
        for e in led.get("published", []):
            if e.get("video_id") == video_id and e.get("privacy") != "public":
                e["privacy"] = "public"
                hit = True
        if not hit:
            return
        if dry_run:
            print(f"DRY RUN: would mark {video_id} public in the ledger")
            return
        with open(lp, "w") as fh:
            _json.dump(led, fh, indent=2)
            fh.write("\n")
        ok = _qk_commit(repo, ["published.json"], f"ledger: {video_id} is public")
        print(f"ledger updated for {video_id}" if ok else
              f"ledger update for {video_id} could not be pushed")
    except Exception as e:                              # noqa: BLE001
        print(f"ledger update for {video_id} failed (video is public regardless): {e}")


def file_amendment(video_id: str, title: str, sender: str, text: str, dry_run: bool) -> bool:
    import subprocess
    from datetime import datetime, timezone
    repo = qk_repo()
    if not repo:
        return False
    d = os.path.join(repo, "amendments")
    os.makedirs(d, exist_ok=True)
    stamp = datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%SZ")
    path = os.path.join(d, f"{stamp}_{video_id}.md")
    with open(path, "w") as fh:
        fh.write(f"video_id: {video_id}\ntitle: {title}\nfrom: {sender}\n"
                 f"received: {stamp}\nstatus: open\n---\n{text}\n")
    if dry_run:
        return True
    g = ["git", "-C", repo]
    subprocess.run(g + ["-c", "user.name=quran-approval-bot", "-c",
                        "user.email=actions@users.noreply.github.com", "add", "amendments"], check=True)
    subprocess.run(g + ["-c", "user.name=quran-approval-bot", "-c",
                        "user.email=actions@users.noreply.github.com", "commit", "-q", "-m",
                        f"amendment requested for {video_id}"], check=True)
    for _ in range(3):
        if subprocess.run(g + ["push", "-q", "origin", "HEAD:main"]).returncode == 0:
            return True
        subprocess.run(g + ["pull", "-q", "--rebase", "origin", "main"])
    return False


# ---------------------------------------------------------------- report
YPP_SUBS, YPP_SHORTS_VIEWS = 1000, 10_000_000


def report() -> str:
    """Views + monetisation + money spent, appended to every approval email."""
    from datetime import date
    lines = ["", "=" * 46, "CHANNEL REPORT", "=" * 46]
    try:
        tok = yt_token()
        ch = yt("GET", "channels?part=statistics,contentDetails&mine=true", tok)["items"][0]
        st = ch["statistics"]
        subs = int(st.get("subscriberCount", 0))
        total = int(st.get("viewCount", 0))
        lines += [f"Subscribers: {subs:,}    Total views: {total:,}    Videos: {st.get('videoCount')}"]
        up = ch["contentDetails"]["relatedPlaylists"]["uploads"]
        items = yt("GET", f"playlistItems?part=contentDetails&maxResults=25&playlistId={up}",
                   tok).get("items", [])
        ids = ",".join(i["contentDetails"]["videoId"] for i in items)
        if ids:
            vids = yt("GET", f"videos?part=snippet,statistics,status&id={ids}", tok).get("items", [])
            lines += ["", "Latest videos (views / likes / comments, visibility):"]
            for v in vids:
                if v["snippet"]["title"].startswith("LICENCE TEST"):
                    continue                            # old reciter Content-ID probes
                s2 = v.get("statistics", {})
                lines.append(f"  {int(s2.get('viewCount', 0)):>7,} / {int(s2.get('likeCount', 0)):>4,} / "
                             f"{int(s2.get('commentCount', 0)):>3,}  {v['status']['privacyStatus']:<8} "
                             f"{v['snippet']['title'][:48]}")
        from datetime import date as _d, timedelta as _td
        end = _d.today().isoformat()
        start = (_d.today() - _td(days=90)).isoformat()

        def ya(metrics, extra=""):
            q = (f"https://youtubeanalytics.googleapis.com/v2/reports?ids=channel%3D%3DMINE"
                 f"&startDate={start}&endDate={end}&metrics={metrics}{extra}")
            req = urllib.request.Request(q, headers={"Authorization": f"Bearer {tok}"})
            with urllib.request.urlopen(req, timeout=60) as r:
                rows = json.load(r).get("rows") or [[0]]
            return rows[0][0]

        try:
            shorts90 = int(ya("views", "&filters=creatorContentType%3D%3DSHORTS"))
            s90 = f"{shorts90:,} = {100*shorts90/YPP_SHORTS_VIEWS:.3f}% of 10M"
        except Exception as e:                          # noqa: BLE001
            s90 = f"unavailable ({str(e)[:60]})"
        try:
            rev = f"${float(ya('estimatedRevenue')):,.2f} (last 90 days)"
        except Exception:                               # noqa: BLE001
            rev = "$0 - channel not monetised yet (YouTube returns no revenue data)"
        lines += ["", "Monetisation:",
                  f"  Subscribers: {subs:,} of {YPP_SUBS:,} needed ({100*subs/YPP_SUBS:.1f}%)",
                  f"  Shorts views, last 90 days: {s90}",
                  f"  Ad revenue: {rev}"]
    except Exception as e:                              # noqa: BLE001
        lines.append(f"(YouTube stats unavailable: {e})")

    # money spent — written by the daily-art routine into quran-kids/costs.jsonl
    month = date.today().strftime("%Y-%m")
    imgs = usd = credits = 0.0
    eps = set()
    try:
        repo = qk_repo()
        p = os.path.join(repo, "costs.jsonl") if repo else ""
        if p and os.path.exists(p):
            for ln in open(p):
                try:
                    r = json.loads(ln)
                except ValueError:
                    continue
                if str(r.get("date", "")).startswith(month):
                    imgs += r.get("images", 0)
                    usd += float(r.get("usd", 0) or 0)
                    credits += float(r.get("krea_credits", 0) or 0)
                    eps.add(r.get("surah"))
    except Exception as e:                              # noqa: BLE001
        lines.append(f"(cost log unavailable: {e})")
    lines += ["", f"Money spent this month ({month}):",
              f"  Krea images: {int(imgs)} generated for {len(eps)} episode(s), "
              f"{int(credits)} Krea credits from your plan, extra cash ${usd:.2f}",
              "  GitHub (render + email bot): $0 - free tier",
              "  YouTube API / quran.com text + audio: $0",
              "  Claude (daily art + review): included in your Claude plan, $0 extra",
              f"  TOTAL: ${usd:.2f}"]
    return "\n".join(lines) + "\n"


# ---------------------------------------------------------------- youtube
def yt_token() -> str:
    t = json.loads(env("YOUTUBE_TOKEN_JSON"))
    data = urllib.parse.urlencode({
        "client_id": t["client_id"], "client_secret": t["client_secret"],
        "refresh_token": t["refresh_token"], "grant_type": "refresh_token"}).encode()
    with urllib.request.urlopen(t.get("token_uri", "https://oauth2.googleapis.com/token"),
                                data, timeout=30) as r:
        return json.load(r)["access_token"]


def yt(method: str, path: str, token: str, body: dict | None = None) -> dict:
    req = urllib.request.Request(
        "https://www.googleapis.com/youtube/v3/" + path, method=method,
        data=json.dumps(body).encode() if body is not None else None,
        headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"})
    with urllib.request.urlopen(req, timeout=60) as r:
        return json.load(r)


def make_public(video_id: str, dry_run: bool) -> str:
    tok = yt_token()
    items = yt("GET", f"videos?part=status,snippet&id={video_id}", tok).get("items", [])
    if not items:
        return f"video {video_id} not found on the channel — nothing changed"
    st = items[0]["status"]
    title = items[0]["snippet"]["title"]
    if st.get("privacyStatus") == "public":
        return f"'{title}' was already public"
    keep = ("embeddable", "license", "publicStatsViewable",
            "selfDeclaredMadeForKids", "containsSyntheticMedia")
    status = {k: st[k] for k in keep if k in st}
    status["privacyStatus"] = "public"
    if dry_run:
        return f"DRY RUN: would make '{title}' public"
    yt("PUT", "videos?part=status", tok, {"id": video_id, "status": status})
    return f"'{title}' is now PUBLIC"


# ---------------------------------------------------------------- modes
def request(video_id: str, title: str, notes: str):
    url = f"https://youtube.com/watch?v={video_id}"
    body = (f"New episode ready for your approval: {title}\n\n"
            f"Watch it (private, only you can see it):\n{url}\n\n")
    if notes.strip():
        body += notes.strip() + "\n\n"
    body += (
        "TO PUBLISH: reply to this email with the single word  APPROVE\n"
        "TO KEEP IT PRIVATE: reply  REJECT  (or just ignore this email)\n\n"
        "The bot checks for replies about every 30 minutes and will email you "
        "back when the video is public. Only replies from your own address are "
        "accepted, and only to this email.\n"
        "TO ASK FOR CHANGES: reply with what to change (e.g. \"make the birds bigger\" or\n"
        "\"redo the ayah 3 picture\"). A corrected version comes back to you for approval.\n")
    body += report()
    subject = f"[Quran With My Child] Approve to publish: {title}  ({ref_for(video_id)})"
    send(env("QWMC_NOTIFY"), subject, body)
    print("approval request sent for", video_id)


def announce(M, dry_run: bool):
    """Email Ahmed about every new PRIVATE upload that has not been announced.

    The cloud renderer (quran-kids repo) cannot send mail, so this finds its
    uploads on the channel instead. "Already announced" = a request carrying
    that video's tag sits in the bot's Sent folder — no state file to rot.
    Only uploads after QWMC_ANNOUNCE_AFTER (ISO date) count, so old test
    clips are never announced.
    """
    after = os.environ.get("QWMC_ANNOUNCE_AFTER", "2026-10-07T12:00")
    tok = yt_token()
    ch = yt("GET", "channels?part=contentDetails&mine=true", tok)["items"][0]
    up = ch["contentDetails"]["relatedPlaylists"]["uploads"]
    items = yt("GET", f"playlistItems?part=contentDetails&maxResults=15&playlistId={up}",
               tok).get("items", [])
    ids = [i["contentDetails"]["videoId"] for i in items
           if i["contentDetails"].get("videoPublishedAt", "9") >= after]
    if not ids:
        return
    vids = yt("GET", "videos?part=status,snippet&id=" + ",".join(ids), tok).get("items", [])
    M.select('"[Gmail]/Sent Mail"', readonly=True)
    for v in vids:
        if v["status"].get("privacyStatus") != "private":
            continue
        vid = v["id"]
        typ, data = M.search(None, "SUBJECT", f'"{ref_for(vid)}"')
        if typ == "OK" and data[0].split():
            continue
        title = v["snippet"]["title"]
        print(f"announcing new private upload {vid}: {title}")
        if not dry_run:
            request(vid, title, "Rendered and uploaded automatically by the pipeline.")
    M.select("INBOX")


def poll(dry_run: bool):
    allowed = approvers()
    M = imaplib.IMAP4_SSL("imap.gmail.com")
    M.login(env("MAIL_USER"), env("MAIL_PASSWORD"))
    try:
        announce(M, dry_run)
    except Exception as e:                              # noqa: BLE001
        print("announce failed:", e)
    M.select("INBOX")
    typ, data = M.search(None, "UNSEEN", "SUBJECT", f'"{TAG}"')
    ids = data[0].split() if typ == "OK" else []
    print(f"{len(ids)} unread approval repl(y/ies)")
    for num in ids:
        # BODY.PEEK so a dry run leaves the message unread
        _, raw = M.fetch(num, "(BODY.PEEK[])")
        msg = email.message_from_bytes(raw[0][1])
        subject = str(make_header(decode_header(msg.get("Subject", ""))))
        sender = email.utils.parseaddr(msg.get("From", ""))[1].lower()
        m = REF_RE.search(subject)
        mid = msg.get("Message-ID")
        print(f"- from {sender}: {subject[:90]}")

        if sender == env("MAIL_USER").lower():
            continue                                   # our own outgoing copy
        problems = []
        if not m or not hmac.compare_digest(ref_for(m.group(1)), m.group(0)):
            problems.append("reference tag does not verify")
        if sender not in allowed:
            problems.append(f"sender {sender} is not an approver")
        if not authenticated(msg):
            problems.append("sender authentication (DKIM/DMARC) did not pass")
        if problems:
            print("  ignored:", "; ".join(problems))
            if not dry_run:
                M.store(num, "+FLAGS", "\\Seen")
            continue

        vid = m.group(1)
        v = verdict(first_line(msg))
        if v == "approve":
            try:
                result = make_public(vid, dry_run)
            except Exception as e:                      # noqa: BLE001
                result = f"FAILED to publish ({e}). The video is still private."
            else:
                # Outside the try on purpose: the video is already public here,
                # so nothing this does may ever report the publish as failed.
                mark_published(vid, dry_run)
            reply = (f"{result}.\n\nhttps://youtube.com/watch?v={vid}\n")
        elif v == "reject":
            reply = ("Got it — the video stays PRIVATE. Nothing was published.\n\n"
                     f"https://youtube.com/watch?v={vid}\n")
        else:
            text = reply_text(msg)
            if len(text) < 3:
                reply = ("That reply was empty, so nothing changed. Reply APPROVE, REJECT, "
                         "or describe what to change.\n")
            else:
                title = re.sub(r"^.*Approve to publish:\s*", "", subject)
                title = re.sub(r"\s*\(" + TAG + r".*$", "", title)
                ok = file_amendment(vid, title, sender, text, dry_run)
                reply = (("Got it - change request logged:\n\n" + text + "\n\nThe video stays "
                          "PRIVATE. The cloud will apply it on its next daily run and send you a "
                          "corrected version to approve.\n") if ok else
                         ("I could not log that change request (repo access failed), so nothing "
                          "changed and the video stays private. Please try again later.\n"))
        print("  ->", reply.splitlines()[0])
        if not dry_run:
            # The video is already published by this point. A failure sending
            # the courtesy confirmation must NOT fail the run — otherwise the
            # workflow reports failure for an episode that actually went
            # public, and Ahmed gets an alarming email about a success.
            try:
                send(sender, "Re: " + re.sub(r"^(re:\s*)+", "", subject, flags=re.I),
                     reply, in_reply_to=mid)
            except Exception as e:                      # noqa: BLE001
                print(f"  (confirmation email failed, action already applied: {e})")
            # Mark read regardless, so one bad reply cannot be reprocessed in a
            # loop and re-publish or re-log on every poll.
            try:
                M.store(num, "+FLAGS", "\\Seen")
            except Exception as e:                      # noqa: BLE001
                print(f"  (could not flag as read: {e})")
    M.logout()


def watch(minutes: int, every: int, dry_run: bool):
    """Poll on a loop inside ONE Actions run.

    GitHub throttles frequent schedules hard: this workflow asked for every
    30 minutes and actually fired roughly every four hours, so an APPROVE
    reply could sit most of a day. Long-running jobs are not throttled the
    same way, so one surviving trigger now covers hours instead of a single
    instant. A transient IMAP/SMTP blip must not end the watch — log it and
    keep going.
    """
    import time
    deadline = time.time() + minutes * 60
    n, failures = 0, 0
    while time.time() < deadline:
        n += 1
        try:
            poll(dry_run)
            failures = 0
        except Exception as e:                          # noqa: BLE001
            failures += 1
            print(f"[watch] poll {n} failed ({failures} in a row): {e}", flush=True)
            # Only give up if it is clearly not transient.
            if failures >= 10:
                raise
        remaining = deadline - time.time()
        if remaining <= 0:
            break
        time.sleep(min(every, max(1, remaining)))
    print(f"[watch] finished after {n} poll(s)", flush=True)


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    sub = ap.add_subparsers(dest="mode", required=True)
    r = sub.add_parser("request")
    r.add_argument("--video-id", required=True)
    r.add_argument("--title", required=True)
    r.add_argument("--notes", default="")
    p = sub.add_parser("poll")
    p.add_argument("--dry-run", action="store_true")
    w = sub.add_parser("watch")
    w.add_argument("--minutes", type=int, default=330)
    w.add_argument("--every", type=int, default=120)
    w.add_argument("--dry-run", action="store_true")
    a = ap.parse_args()
    if a.mode == "request":
        request(a.video_id, a.title, a.notes)
    elif a.mode == "watch":
        watch(a.minutes, a.every, a.dry_run)
    else:
        poll(a.dry_run)
