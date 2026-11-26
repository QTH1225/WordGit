#!/bin/sh
# Runs one probe pass: (re)starts WPS Writer in the requested window mode, confirms WPS's own
# start-up and trust prompts, and waits until the probe finishes or WPS crashes.
#   WPS_WINDOW_MODE=fusion       one window with tabs for all components (WPS default)
#   WPS_WINDOW_MODE=independent  separate Writer window
# Usage: probe-run.sh <report-file>; prints "result=done|crashed|timeout".
set -u

REPORT="$1"
OUT=/out
WORK="$HOME/probe"
CONF="$HOME/.config/Kingsoft/Office.conf"
WINDOW_MODE="${WPS_WINDOW_MODE:-fusion}"
export DISPLAY=:99

pkill -9 -f '[/]opt/kingsoft' 2>/dev/null
sleep 1
rm -f "$REPORT"
find "$WORK" -maxdepth 1 -name 'probe-*' ! -name 'probe-options.txt' -delete

case "$WINDOW_MODE" in
    fusion) component=prome_fushion ;;
    independent) component=prome_independ ;;
    *) echo "unknown WPS_WINDOW_MODE $WINDOW_MODE" >&2; exit 2 ;;
esac
python3 - "$CONF" "$component" <<'EOF'
import re, sys
path, component = sys.argv[1], sys.argv[2]
key = "wpsoffice\\Application%20Settings\\AppComponentMode="
try:
    text = open(path, encoding="utf-8").read()
except FileNotFoundError:
    text = "[6.0]\n"
lines = [line for line in text.splitlines() if not line.startswith(key)]
section = lines.index("[6.0]") if "[6.0]" in lines else None
if section is None:
    lines[:0] = ["[6.0]"]
    section = 0
lines.insert(section + 1, key + component)
if not any(line.startswith("common\\AcceptedEULA=") for line in lines):
    lines.insert(section + 1, "common\\AcceptedEULA=true")
open(path, "w", encoding="utf-8").write("\n".join(lines) + "\n")
EOF

crash_marker="$OUT/.probe-start"
touch "$crash_marker"
wps "$WORK/start.docx" >>"$OUT/wps.log" 2>&1 &

# WPS opens modal prompts on every start ("set as default office suite") and whenever the add-in
# is new or changed ("trust this add-in?"). Their default button confirms, so press Return on a
# window titled exactly "WPS Office" or "WPS 文字" — but only until the probe has started, so the
# probe's own document operations are never answered on its behalf.
i=0
result=timeout
while [ "$i" -lt 90 ]; do
    sleep 2
    i=$((i + 1))
    if grep -q '"step":"summary.save-text"' "$REPORT" 2>/dev/null; then
        result=done
        break
    fi
    if [ -n "$(find "$HOME/.local/share/Kingsoft/tmp" -name '*.dmp' -newer "$crash_marker" 2>/dev/null)" ]; then
        result=crashed
        break
    fi
    if ! grep -q '"step":"start"' "$REPORT" 2>/dev/null; then
        for window in $(xdotool search --onlyvisible --name '^WPS (Office|文字)$' 2>/dev/null); do
            xdotool windowactivate --sync "$window" key Return 2>/dev/null || true
            echo "$(date +%T) confirmed dialog $window" >>"$OUT/dialogs.log"
        done
    fi
done
import -window root "$OUT/screen-$WINDOW_MODE-$result.png" 2>/dev/null || true
echo "result=$result window=$WINDOW_MODE seconds=$((i * 2))"
