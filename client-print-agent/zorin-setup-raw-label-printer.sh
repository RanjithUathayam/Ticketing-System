#!/usr/bin/env bash
# Converts the TSC label printer's CUPS queue on Zorin OS / Ubuntu into a RAW queue.
#
# Why: with a normal CUPS driver, TSPL sent by QZ Tray is detected as plain text and
# rendered as a text page (the label prints "TEXT 20,252,..." instead of the label),
# and banner/cover pages get printed before each job. A raw queue passes the bytes
# straight to the printer so it executes the TSPL commands.
#
# Usage:  sudo bash zorin-setup-raw-label-printer.sh ["Existing CUPS printer name"]
set -euo pipefail

if [[ $EUID -ne 0 ]]; then
  echo "Run with sudo: sudo bash $0 [printer-name]" >&2
  exit 1
fi

OLD_NAME="${1:-}"
NEW_NAME="TSC_TTP-244_Pro"

if [[ -z "$OLD_NAME" ]]; then
  echo "Installed printers:"
  lpstat -p 2>/dev/null | awk '{print "  " $2}'
  read -rp "Enter the label printer name from the list above: " OLD_NAME
fi

DEVICE_URI="$(lpstat -v "$OLD_NAME" 2>/dev/null | sed -E 's/^device for [^:]+: //')"
if [[ -z "$DEVICE_URI" ]]; then
  echo "Could not find printer \"$OLD_NAME\". Detected USB devices:" >&2
  lpinfo -v | grep -i usb >&2 || true
  exit 1
fi
echo "Device URI: $DEVICE_URI"

# Re-create the queue without any driver (raw) and without banner pages.
lpadmin -x "$OLD_NAME" || true
lpadmin -p "$NEW_NAME" -E -v "$DEVICE_URI" -m raw \
  -o job-sheets-default=none,none \
  -o printer-error-policy=abort-job
cupsenable "$NEW_NAME"
cupsaccept "$NEW_NAME"

# Belt and braces: make sure raw printing is allowed by the CUPS mime config.
for f in /etc/cups/raw.types /usr/share/cups/mime/raw.types; do
  [[ -f "$f" ]] && sed -i 's/^#\s*application\/octet-stream/application\/octet-stream/' "$f"
done
for f in /etc/cups/raw.convs /usr/share/cups/mime/raw.convs; do
  [[ -f "$f" ]] && sed -i 's/^#\s*application\/octet-stream/application\/octet-stream/' "$f"
done
systemctl restart cups

# Test label: should print a real barcode and text, not command text.
printf 'SIZE 90 mm,44 mm\r\nGAP 3 mm,0 mm\r\nCLS\r\nBARCODE 20,20,"128",80,1,0,2,2,"TEST123"\r\nTEXT 20,140,"3",0,1,1,"RAW QUEUE OK"\r\nPRINT 1,1\r\n' \
  | lp -d "$NEW_NAME" -o raw

echo
echo "Done. Printer queue \"$NEW_NAME\" is now RAW."
echo "In the web app label settings, set Printer Name to: $NEW_NAME"
