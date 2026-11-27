#!/usr/bin/env bash
set -euo pipefail

archive=$(realpath "$1")
legacy=${2:-}
if [[ -n $legacy ]]; then legacy=$(realpath "$legacy"); fi
scratch=$(mktemp -d)
export HOME="$scratch/home with spaces"
unset WORDGIT_HOME
root="$HOME/.local/share/WordGit"
service="$root/service.json"
installed="$root/app/wordgit"
mkdir -p "$HOME" "$scratch/first" "$scratch/second" "$scratch/legacy"

service_pid() {
    python3 -c 'import json, sys; print(json.load(open(sys.argv[1]))["pid"])' "$service"
}

cleanup() {
    if [[ -f $service ]]; then
        local pid program
        pid=$(service_pid 2>/dev/null) || pid=
        if [[ $pid =~ ^[0-9]+$ ]]; then
            program=$(readlink "/proc/$pid/exe" || true)
            if [[ $program == "$scratch/"* ]]; then kill -TERM "$pid" 2>/dev/null || true; fi
        fi
    fi
    rm -rf -- "$scratch"
}
trap cleanup EXIT

wait_for_pid() {
    local expected=$1
    for _ in {1..100}; do
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
wps="$HOME/.local/share/Kingsoft/wps/jsaddons"
mkdir -p -- "$wps"
cat > "$wps/publish.xml" <<'EOF'
<jsplugins>
<jsplugin url="file:///example/" name="OtherAddin"/>
</jsplugins>
EOF

"$first/install.sh" --data-directory "$storage"
test -x "$installed"
grep -F "Exec=\"$installed\" serve" "$HOME/.config/autostart/wordgit.desktop"
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
original_hash=$(sha256sum "$document" | cut -d ' ' -f 1)
commit=$("$scratch/legacy/wordgit" commit "$document" -m first --author "WordGit test" | python3 -c 'import json,sys; print(json.load(sys.stdin)["commit"])')
test "$commit" != None
kill -TERM "$first_pid"
for _ in {1..100}; do
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
test "$(readlink "/proc/$updated_pid/exe")" = "$installed"
test "$("$installed" data-directory)" = "$storage"
revision=$(python3 -c 'import json,sys; print(json.dumps({"document":sys.argv[1],"revision":sys.argv[2]}))' "$document" "$commit")
file=$("$installed" call version/preview "$revision" | python3 -c 'import json,sys; print(json.load(sys.stdin)["file"])')
test "$(sha256sum "$file" | cut -d ' ' -f 1)" = "$original_hash"

"$second/uninstall.sh"
test ! -e "$installed"
test ! -e "$service"
test ! -e "$HOME/.config/autostart/wordgit.desktop"
grep -F 'name="OtherAddin"' "$wps/publish.xml"
! grep -F 'name="WordGit"' "$wps/publish.xml"
test -f "$root/data-directory.json" && test -f "$file" && test -d "$storage/repositories"
"$scratch/legacy/wordgit" log "$document" | grep -F "${commit:0:8}"
echo 'PASS: Linux install, autostart, SIGTERM, upgrade and uninstall preserve history.'
