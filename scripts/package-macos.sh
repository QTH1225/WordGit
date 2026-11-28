#!/usr/bin/env bash
set -euo pipefail

if [[ $# -lt 2 || $# -gt 3 ]]; then
    echo "Usage: $0 <wordgit-binary> <output-dir> [arch]" >&2
    exit 1
fi

abs_path() {
    python3 -c 'import os, sys; print(os.path.realpath(sys.argv[1]))' "$1"
}

kernel=$(abs_path "$1")
mkdir -p -- "$2"
output=$(abs_path "$2")
version=$("$kernel" version)
if [[ ! $version =~ ^wordgit\ ([0-9]+\.[0-9]+\.[0-9]+(-beta\.[0-9]+)?)$ ]]; then
    echo "Unexpected WordGit version: $version" >&2
    exit 1
fi
ver="${BASH_REMATCH[1]}"

if [[ $# -eq 3 ]]; then
    arch="$3"
else
    archs=$(lipo -archs "$kernel" 2>/dev/null || uname -m)
    if [[ " $archs " == *" arm64 "* && " $archs " == *" x86_64 "* ]]; then
        arch="universal"
    elif [[ " $archs " == *" arm64 "* ]]; then
        arch="arm64"
    elif [[ " $archs " == *" x86_64 "* ]]; then
        arch="x86_64"
    else
        arch="$(uname -m)"
    fi
fi

name="WordGit-${ver}-macos-${arch}"
root=$(cd "$(dirname "$0")/.." && pwd)
scratch=$(mktemp -d)
trap 'rm -rf -- "$scratch"' EXIT
payload="$scratch/$name"
mkdir -p -- "$payload/docs" "$output"
install -m 755 -- "$kernel" "$payload/wordgit"
cp -R -- "$root/addin-wps" "$payload/addin-wps"
cp -- "$root/LICENSE" "$root/README.md" "$payload/"
cp -- "$root/docs/WPS.md" "$root/docs/MACOS.md" "$root/docs/USAGE.md" \
      "$root/docs/DEVELOPMENT.md" "$root/docs/RELEASING.md" "$root/docs/INSTALL-DISCLAIMER.md" "$payload/docs/"
cp -R -- "$root/docs/assets" "$payload/docs/"

cat > "$payload/install.sh" <<'EOF'
#!/bin/sh
set -eu
package_dir=$(CDPATH= cd -- "$(dirname -- "$0")" && pwd)
exec "$package_dir/wordgit" install --addin "$package_dir/addin-wps" "$@"
EOF

cat > "$payload/uninstall.sh" <<'EOF'
#!/bin/sh
set -eu
package_dir=$(CDPATH= cd -- "$(dirname -- "$0")" && pwd)
installed="${WORDGIT_HOME:-$HOME/Library/Containers/com.kingsoft.wpsoffice.mac/Data/WordGit}/app/wordgit"
if [ -x "$installed" ]; then
    exec "$installed" uninstall
fi
exec "$package_dir/wordgit" uninstall
EOF

cat > "$payload/Install.command" <<'EOF'
#!/bin/sh
set -eu
package_dir=$(CDPATH= cd -- "$(dirname -- "$0")" && pwd)
echo "=========================================="
echo "  WordGit for macOS (WPS 文字) 安装程序"
echo "=========================================="
echo ""
if command -v xattr >/dev/null 2>&1; then
    xattr -dr com.apple.quarantine "$package_dir" 2>/dev/null || true
fi
"$package_dir/install.sh" "$@"
status=$?
echo ""
if [ "$status" -eq 0 ]; then
    echo "安装完成！请完全退出并重新打开 WPS Office。"
else
    echo "安装未完成（退出码：$status）。请检查上方提示。"
fi
printf "按回车键关闭窗口..."
read -r _
exit "$status"
EOF

cat > "$payload/Uninstall.command" <<'EOF'
#!/bin/sh
set -eu
package_dir=$(CDPATH= cd -- "$(dirname -- "$0")" && pwd)
echo "=========================================="
echo "  WordGit for macOS (WPS 文字) 卸载程序"
echo "=========================================="
echo ""
"$package_dir/uninstall.sh" "$@"
status=$?
echo ""
printf "按回车键关闭窗口..."
read -r _
exit "$status"
EOF

chmod 755 "$payload/install.sh" "$payload/uninstall.sh" "$payload/Install.command" "$payload/Uninstall.command"
COPYFILE_DISABLE=1 tar -C "$scratch" -czf "$output/$name.tar.gz" "$name"
(
    cd "$output"
    shasum -a 256 "$name.tar.gz" > "$name.tar.gz.sha256"
    shasum -a 256 -c "$name.tar.gz.sha256"
)
