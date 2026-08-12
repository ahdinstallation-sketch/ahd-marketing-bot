#!/bin/bash
# Build the daily AHD client follow-up report and PRINT it on the local/WiFi
# printer. Runs on Ahmed's Mac (the GitHub Actions cloud job can't reach the
# printer). Prints only when the Mac is on/awake and on the same network as the
# printer. Skips Friday to match the emailed report.
set -euo pipefail
cd "$(dirname "$0")"

LOG="print_followup.log"
exec >>"$LOG" 2>&1
echo "=== $(date) : print_followup start ==="

# Skip Friday (Cairo) — same weekend rule as the email (date +%u: 5 = Friday).
if [ "$(TZ=Africa/Cairo date +%u)" = "5" ]; then
  echo "Cairo Friday — skipping print."
  exit 0
fi

# Target printer. Override with FOLLOWUP_PRINTER if it ever changes / you add one.
PRINTER="${FOLLOWUP_PRINTER:-HP_LaserJet_M1536dnf_MFP__6DD913_}"

# 1) Build the report HTML (pulls the live tracker, sends NO email).
/usr/bin/python3 followup_report.py --dry-run

# 2) Render to PDF (headless Chrome, single page).
CHROME="/Applications/Google Chrome.app/Contents/MacOS/Google Chrome"
"$CHROME" --headless --disable-gpu --no-pdf-header-footer \
  --print-to-pdf="followup_print.pdf" "followup_report.html" 2>/dev/null

# 3) Print it. If the default printer is unreachable, CUPS queues it and prints
#    once the printer is back on the network.
if [ ! -s followup_print.pdf ]; then
  echo "ERROR: followup_print.pdf not produced — nothing to print."
  exit 1
fi
lpr -P "$PRINTER" -o media=A4 followup_print.pdf
echo "Queued followup_print.pdf -> $PRINTER"
lpstat -o "$PRINTER" 2>/dev/null || true
echo "=== done ==="
