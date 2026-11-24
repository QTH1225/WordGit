#!/bin/sh
# End-to-end run of the real WordGit add-in and service inside WPS:
# fresh data, a local bare "remote", the service, `wordgit install`, then WPS with the e2e script
# injected into the deployed add-in. Waits for the test to finish or WPS to crash.
#   WPS_WINDOW_MODE=fusion|independent (see probe-run.sh)
# Prints "result=done|failed|crashed|timeout".
set -u

OUT=/out
E2E="$HOME/e2e"
REPORT="$OUT/e2e-report.jsonl"
CONF="$HOME/.config/Kingsoft/Office.conf"
WINDOW_MODE="${WPS_WINDOW_MODE:-fusion}"
export DISPLAY=:99

pkill -9 -f '[/]opt/kingsoft' 2>/dev/null
pkill -f '[w]ordgit serve' 2>/dev/null
sleep 1
rm -rf "$E2E" "$HOME/.local/share/WordGit" "$REPORT"
mkdir -p "$E2E" "$HOME/.config/Kingsoft"
git init -q --bare --initial-branch=main "$E2E/shared.git"

WORDGIT_ALLOW_LOCAL_REMOTES=1 nohup wordgit serve >"$OUT/service.log" 2>&1 &
i=0
until curl -fs http://127.0.0.1:47821/api/v1/ping >/dev/null 2>&1; do
    i=$((i + 1))
    if [ "$i" -gt 50 ]; then echo "result=service-failed"; cat "$OUT/service.log"; exit 1; fi
    sleep 0.2
done

wordgit install --addin /opt/wordgit/addin-wps --no-autostart >"$OUT/install.log" 2>&1 || { echo "result=install-failed"; cat "$OUT/install.log"; exit 1; }
DEPLOYED=$(ls -d "$HOME"/.local/share/Kingsoft/wps/jsaddons/WordGit_* | head -1)
mkdir -p "$DEPLOYED/test"
cp /opt/wordgit/addin-wps-test/e2e.js "$DEPLOYED/test/e2e.js"
sed -i 's|<script src="js/ribbon.js"></script>|<script src="js/ribbon.js"></script>\n    <script src="test/e2e.js"></script>|' "$DEPLOYED/index.html"
cp "$HOME/.local/share/Kingsoft/wps/jsaddons/publish.xml" "$OUT/publish-e2e.xml"

case "$WINDOW_MODE" in
    fusion) component=prome_fushion ;;
    independent) component=prome_independ ;;
    *) echo "unknown WPS_WINDOW_MODE $WINDOW_MODE" >&2; exit 2 ;;
esac
python3 - "$CONF" "$component" <<'EOF'
import sys
path, component = sys.argv[1], sys.argv[2]
key = "wpsoffice\\Application%20Settings\\AppComponentMode="
try:
    text = open(path, encoding="utf-8").read()
except FileNotFoundError:
    text = "[6.0]\n"
lines = [line for line in text.splitlines() if not line.startswith(key)]
if "[6.0]" not in lines:
    lines.insert(0, "[6.0]")
section = lines.index("[6.0]")
lines.insert(section + 1, key + component)
if not any(line.startswith("common\\AcceptedEULA=") for line in lines):
    lines.insert(section + 1, "common\\AcceptedEULA=true")
open(path, "w", encoding="utf-8").write("\n".join(lines) + "\n")
EOF

python3 /opt/wordgit/make-docx.py "$E2E/start.docx"
touch "$E2E/run-e2e"
crash_marker="$OUT/.e2e-start"
touch "$crash_marker"
wps "$E2E/start.docx" >>"$OUT/wps.log" 2>&1 &

# Confirm WPS's own start-up prompts ("default office suite", "trust this add-in") only until the
# test has started, so nothing the test triggers is answered on its behalf.
i=0
result=timeout
while [ "$i" -lt 150 ]; do
    sleep 2
    i=$((i + 1))
    if grep -q '"step":"done"' "$REPORT" 2>/dev/null; then
        if grep -q '"step":"done","ok":true' "$REPORT"; then result=done; else result=failed; fi
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
sleep 2
import -window root "$OUT/screen-e2e-$WINDOW_MODE-$result.png" 2>/dev/null || true
echo "result=$result window=$WINDOW_MODE seconds=$((i * 2))"
