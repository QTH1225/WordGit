#!/usr/bin/env bash
set -euo pipefail

kernel=$(realpath "$1")
output=$(realpath -m "$2")
version=$("$kernel" version)
if [[ ! $version =~ ^wordgit\ ([0-9]+\.[0-9]+\.[0-9]+(-beta\.[0-9]+)?)$ ]]; then
    echo "Unexpected WordGit version: $version" >&2
    exit 1
fi
name="WordGit-${BASH_REMATCH[1]}-linux-x86_64"
root=$(cd "$(dirname "$0")/.." && pwd)
scratch=$(mktemp -d)
trap 'rm -rf -- "$scratch"' EXIT
payload="$scratch/$name"
mkdir -p -- "$payload/docs" "$output"
install -m 755 -- "$kernel" "$payload/wordgit"
cp -R -- "$root/addin-wps" "$payload/addin-wps"
cp -- "$root/LICENSE" "$root/README.md" "$payload/"
cp -- "$root/docs/"{WPS,USAGE,DEVELOPMENT,RELEASING,INSTALL-DISCLAIMER}.md "$payload/docs/"
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
installed="${WORDGIT_HOME:-$HOME/.local/share/WordGit}/app/wordgit"
if [ -x "$installed" ]; then
    exec "$installed" uninstall
fi
exec "$package_dir/wordgit" uninstall
EOF
chmod 755 "$payload/install.sh" "$payload/uninstall.sh"
tar -C "$scratch" -czf "$output/$name.tar.gz" "$name"
(
    cd "$output"
    sha256sum "$name.tar.gz" > "$name.tar.gz.sha256"
    sha256sum -c "$name.tar.gz.sha256"
)
