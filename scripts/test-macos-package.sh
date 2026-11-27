#!/usr/bin/env bash
set -euo pipefail

abs_path() {
    python3 -c 'import os, sys; print(os.path.realpath(sys.argv[1]))' "$1"
}

archive=$(abs_path "$1")
legacy=${2:-}
if [[ -n $legacy ]]; then legacy=$(abs_path "$legacy"); fi
scratch=$(abs_path "$(mktemp -d)")
export HOME="$scratch/home with spaces"
export WORDGIT_NO_LAUNCHCTL=1
unset WORDGIT_HOME
root="$HOME/Library/Containers/com.kingsoft.wpsoffice.mac/Data/WordGit"
service="$root/service.json"
installed="$root/app/wordgit"
plist="$HOME/Library/LaunchAgents/com.qth1225.wordgit.plist"
mkdir -p "$HOME" "$scratch/first" "$scratch/second" "$scratch/legacy"

service_pid() {
    python3 -c 'import json, sys; print(json.load(open(sys.argv[1]))["pid"])' "$service"
}

process_exe() {
    ps -p "$1" -o comm= 2>/dev/null | sed 's/^[[:space:]]*//;s/[[:space:]]*$//'
}

sha256_of() {
    shasum -a 256 "$1" | awk '{print $1}'
}

cleanup() {
    if [[ -f $service ]]; then
        local pid program
        pid=$(service_pid 2>/dev/null) || pid=
        if [[ $pid =~ ^[0-9]+$ ]]; then
            program=$(process_exe "$pid" || true)
            if [[ $program == "$scratch/"* ]]; then kill -TERM "$pid" 2>/dev/null || true; fi
        fi
    fi
    rm -rf -- "$scratch"
}
trap cleanup EXIT

wait_for_pid() {
    local expected=$1
    for _ in $(seq 1 100); do
        if [[ -f $service ]] && [[ $(service_pid) == "$expected" ]] && curl -fsS --max-time 1 http://127.0.0.1:47821/api/v1/ping >/dev/null; then
            return
        fi
        sleep 0.1
    done
    echo "WordGit service $expected did not start" >&2
    return 1
}

tar -xzf "$archive" -C "$scratch/first"
tar -xzf "$archive" -C "$scratch/second"
first=("$scratch/first"/WordGit-*)
second=("$scratch/second"/WordGit-*)
storage="$HOME/custom-history"
wps="$HOME/Library/Containers/com.kingsoft.wpsoffice.mac/Data/.kingsoft/wps/jsaddons"
mkdir -p -- "$wps"
cat > "$wps/publish.xml" <<'EOF'
<jsplugins>
<jsplugin url="file:///example/" name="OtherAddin"/>
</jsplugins>
EOF

"$first/install.sh" --data-directory "$storage"
test -x "$installed"
test -f "$plist"
plutil -lint "$plist" >/dev/null
grep -F "<string>$installed</string>" "$plist" >/dev/null
grep -F "<string>com.qth1225.wordgit</string>" "$plist" >/dev/null
grep -F "<key>EnvironmentVariables</key>" "$plist" >/dev/null
first_pid=$(service_pid)
wait_for_pid "$first_pid"
cp -- "${legacy:-$first/wordgit}" "$scratch/legacy/wordgit"
chmod +x "$scratch/legacy/wordgit"
mv -- "$first" "$scratch/retired"

document="$HOME/test.docx"
python3 - "$document" <<'PY'
import sys
from zipfile import ZipFile
with ZipFile(sys.argv[1], "w") as doc:
    doc.writestr("[Content_Types].xml", "<Types/>")
    doc.writestr("word/document.xml", "<w:document><w:body><w:p><w:t>kept across upgrade</w:t></w:p></w:body></w:document>")
PY
original_hash=$(sha256_of "$document")
commit=$("$scratch/legacy/wordgit" commit "$document" -m first --author "WordGit test" | python3 -c 'import json,sys; print(json.load(sys.stdin)["commit"])')
test "$commit" != None
kill -TERM "$first_pid"
for _ in $(seq 1 100); do
    if [[ ! -e $service ]]; then break; fi
    sleep 0.1
done
test ! -e "$service"

"$scratch/legacy/wordgit" serve >"$scratch/legacy.log" 2>&1 &
legacy_pid=$!
wait_for_pid "$legacy_pid"
"$second/install.sh"
updated_pid=$(service_pid)
test "$updated_pid" != "$legacy_pid"
wait "$legacy_pid" || test "$?" -eq 143
wait_for_pid "$updated_pid"
test "$(process_exe "$updated_pid")" = "$installed"
test "$("$installed" data-directory)" = "$storage"
revision=$(python3 -c 'import json,sys; print(json.dumps({"document":sys.argv[1],"revision":sys.argv[2]}))' "$document" "$commit")
file=$("$installed" call version/preview "$revision" | python3 -c 'import json,sys; print(json.load(sys.stdin)["file"])')
test "$(sha256_of "$file")" = "$original_hash"

"$second/uninstall.sh"
test ! -e "$installed"
test ! -e "$service"
test ! -e "$plist"
grep -F 'name="OtherAddin"' "$wps/publish.xml" >/dev/null
! grep -F 'name="WordGit"' "$wps/publish.xml"
test -f "$root/data-directory.json" && test -f "$file" && test -d "$storage/repositories"
"$scratch/legacy/wordgit" log "$document" | grep -F "${commit:0:8}" >/dev/null
echo 'PASS: macOS install, LaunchAgent plist, SIGTERM, upgrade and uninstall preserve history.'
