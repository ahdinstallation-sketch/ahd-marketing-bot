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
def send(to: str, subject: str, body: str, in_reply_to: str | None = None):
    msg = EmailMessage()
    msg["From"] = f"Quran With My Child bot <{env('MAIL_USER')}>"
    msg["To"] = to
    msg["Subject"] = subject
    if in_reply_to:
        msg["In-Reply-To"] = in_reply_to
        msg["References"] = in_reply_to
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
        "accepted, and only to this email.\n")
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
    after = os.environ.get("QWMC_ANNOUNCE_AFTER", "2026-10-08")
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
            request(vid, title, "Rendered and uploaded automatically in the cloud.")
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
            reply = (f"{result}.\n\nhttps://youtube.com/watch?v={vid}\n")
        elif v == "reject":
            reply = ("Got it — the video stays PRIVATE. Nothing was published.\n\n"
                     f"https://youtube.com/watch?v={vid}\n")
        else:
            reply = ("I couldn't tell whether that was an approval, so nothing was "
                     "published. Reply with just APPROVE or REJECT on the first line.\n")
        print("  ->", reply.splitlines()[0])
        if not dry_run:
            send(sender, "Re: " + re.sub(r"^(re:\s*)+", "", subject, flags=re.I),
                 reply, in_reply_to=mid)
            M.store(num, "+FLAGS", "\\Seen")
    M.logout()


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    sub = ap.add_subparsers(dest="mode", required=True)
    r = sub.add_parser("request")
    r.add_argument("--video-id", required=True)
    r.add_argument("--title", required=True)
    r.add_argument("--notes", default="")
    p = sub.add_parser("poll")
    p.add_argument("--dry-run", action="store_true")
    a = ap.parse_args()
    if a.mode == "request":
        request(a.video_id, a.title, a.notes)
    else:
        poll(a.dry_run)
