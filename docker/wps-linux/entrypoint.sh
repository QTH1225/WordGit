#!/bin/sh
# Starts a virtual display reachable through noVNC and the probe service, deploys the probe
# add-in in PROBE_MODE, then runs one probe pass with probe-run.sh.
#   PROBE_MODE=offline  files in jsaddons/WordGitProbe_<version>, registered like the Zotero add-ins
#   PROBE_MODE=online   pages served by the probe service (jspluginonline)
# Rerun a pass later with: docker exec <container> /opt/wordgit/probe-run.sh /out/report-<mode>.jsonl
set -eu

MODE="${PROBE_MODE:-offline}"
OUT=/out
WORK="$HOME/probe"
JSADDONS="$HOME/.local/share/Kingsoft/wps/jsaddons"
PORT=47820
export DISPLAY=:99

mkdir -p "$OUT/shots" "$WORK" "$JSADDONS" "$HOME/.config/Kingsoft"
rm -f "$OUT"/report-"$MODE".jsonl "$OUT"/shots/"$MODE"-*.png
# A restarted container keeps the previous display's lock; Xvfb refuses to start while it exists.
rm -f /tmp/.X99-lock /tmp/.X11-unix/X99

Xvfb :99 -screen 0 1600x1000x24 -nolisten tcp >"$OUT/xvfb.log" 2>&1 &
sleep 1
fluxbox >"$OUT/fluxbox.log" 2>&1 &
x11vnc -display :99 -forever -shared -nopw -rfbport 5900 -quiet >"$OUT/x11vnc.log" 2>&1 &
websockify --web /usr/share/novnc 6080 localhost:5900 >"$OUT/novnc.log" 2>&1 &

wordgit-probe --port "$PORT" --addin-dir /opt/wordgit/addin \
    --report "$OUT/report-$MODE.jsonl" --work-dir "$WORK" >"$OUT/server-$MODE.log" 2>&1 &

rm -rf "$JSADDONS"/WordGitProbe_*
case "$MODE" in
    online)
        printf '%s\n' '<jsplugins>' \
            "  <jspluginonline name=\"WordGitProbe\" type=\"wps\" url=\"http://127.0.0.1:$PORT/addin/\" debug=\"\" enable=\"enable_dev\" install=\"null\" customDomain=\"\"/>" \
            '</jsplugins>' >"$JSADDONS/publish.xml"
        ;;
    offline)
        # Same registration as the widely used Zotero WPS add-ins: files load from the
        # WordGitProbe_<version> folder and the url only identifies the add-in.
        cp -r /opt/wordgit/addin "$JSADDONS/WordGitProbe_0.1.0"
        printf '%s\n' '<?xml version="1.0" encoding="UTF-8"?>' '<jsplugins>' \
            "  <jsplugin url=\"http://127.0.0.1:$PORT/\" type=\"wps\" enable=\"enable_dev\" install=\"null\" version=\"0.1.0\" name=\"WordGitProbe\"/>" \
            '</jsplugins>' >"$JSADDONS/publish.xml"
        ;;
    *)
        echo "unknown PROBE_MODE $MODE" >&2
        exit 2
        ;;
esac
cp "$JSADDONS/publish.xml" "$OUT/publish-$MODE.xml"
python3 /opt/wordgit/make-docx.py "$WORK/start.docx"

/opt/wordgit/probe-run.sh "$OUT/report-$MODE.jsonl" | tee -a "$OUT/runs.log" || true

# Keep a rolling screenshot plus a short history so later manual runs can still be inspected.
n=0
while true; do
    sleep 5
    import -window root "$OUT/screen.png" 2>/dev/null || true
    n=$((n + 1))
    if [ $((n % 3)) -eq 0 ] && [ "$n" -le 360 ]; then
        cp "$OUT/screen.png" "$OUT/shots/$MODE-$(printf '%03d' $((n / 3))).png" 2>/dev/null || true
    fi
    cp "$WORK"/probe-summary-*.txt "$OUT"/ 2>/dev/null || true
done
