#!/usr/bin/env python3
"""Send Ahmed a plain, unmissable alert when a daily report fails to go out.

Why this exists
---------------
On 1 Sep 2026 the Cash Position email arrived and the Client Follow-up and
Marketing Pulse did not. Nobody found out until Ahmed noticed the gap himself,
because a failed GitHub Actions run is completely silent — no email, no warning,
nothing in the inbox to react to. The absence of a message is not a message.

This is called from each workflow's `if: failure()` step, and by the watchdog.

  python3 alert.py --report "AHD Client Follow-up" --reason "SMTP refused" \
                   --run-url "https://github.com/.../actions/runs/123"

Env vars (same secrets the reports already use):
  MAIL_USER / MAIL_PASSWORD   sender + app password
  ALERT_RECIPIENTS            comma-separated; defaults to Ahmed only
"""
import os
import sys
import html

from daily_report import cairo_now, send_email

# Alerts go to Ahmed alone by default. An alert is an operational signal, not a
# business report — the rest of the recipient list does not need it and would
# quickly learn to ignore it.
DEFAULT_ALERT_RECIPIENTS = ["ahmed.helmy@amrhelmydesigns.com"]


def _arg(flag, default=""):
    return sys.argv[sys.argv.index(flag) + 1] if flag in sys.argv else default


def build_body(report, reason, run_url):
    link = (f'<p style="margin:18px 0 0"><a href="{html.escape(run_url)}" '
            f'style="background:#b8933f;color:#1c1c1c;text-decoration:none;'
            f'padding:11px 22px;border-radius:6px;font-weight:bold;'
            f'display:inline-block">Open the run log</a></p>') if run_url else ""
    return (
        '<div style="font-family:Arial,Helvetica,sans-serif;background:#faf8f4;'
        'padding:24px;color:#2b2b2b">'
        '<div style="max-width:560px;margin:0 auto;background:#fff;'
        'border:1px solid #e6ded0;border-left:5px solid #c0392b;border-radius:8px;'
        'overflow:hidden">'
        '<div style="background:#1c1c1c;color:#e8b84b;padding:16px 20px;'
        'font-size:17px;font-weight:bold">Daily email did NOT go out</div>'
        '<div style="padding:20px">'
        f'<p style="font-size:16px;margin:0 0 14px"><b>{html.escape(report)}</b> '
        f'failed on {cairo_now():%a %d %b %Y at %H:%M} Cairo.</p>'
        '<p style="font-size:14px;line-height:1.7;margin:0 0 6px;color:#555">Reason reported:</p>'
        f'<pre style="font-size:13px;background:#f4f1ea;border:1px solid #e6ded0;'
        f'border-radius:5px;padding:12px;margin:0;white-space:pre-wrap;'
        f'word-break:break-word">{html.escape(reason or "no detail captured")}</pre>'
        f'{link}'
        '<p style="font-size:12px;color:#7a7a7a;margin:18px 0 0">'
        'Nobody on the recipient list received this report. It will be retried at the '
        'next scheduled slot today; if none succeed, the 20:00 Cairo watchdog will try '
        'once more.</p>'
        '</div></div></div>'
    )


def main():
    report = _arg("--report", "A daily AHD report")
    reason = _arg("--reason", "")
    run_url = _arg("--run-url", "")

    recipients = [r.strip() for r in os.environ.get("ALERT_RECIPIENTS", "").split(",")
                  if r.strip()] or DEFAULT_ALERT_RECIPIENTS
    subject = f"[FAILED] {report} — {cairo_now():%a %d %b}"

    if not send_email(subject, build_body(report, reason, run_url), recipients):
        # Deliberately exit 0. If the alert itself can't send, the underlying
        # problem is almost certainly SMTP — the very thing being reported. Going
        # red here would only add a second confusing failure on top of the real one.
        print("Could not send the failure alert (SMTP likely down — that IS the failure).")
        return 0
    print(f"Failure alert sent to: {', '.join(recipients)}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
