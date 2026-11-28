#!/usr/bin/env bash
# Linux tasks run inside the development container (scripts/dev.ps1, scripts/dev.sh) and by CI.
#   check    encoding check, WPS add-in checks, clippy, tests and the Linux package test
#   package  static x86_64 release package into dist/
set -euo pipefail
cd "$(dirname "$0")/.."
target=${CARGO_TARGET_DIR:-target}

step() { printf '\n==> %s\n' "$*"; }

check_powershell_encoding() {
    # Windows PowerShell 5.1 reads BOM-less scripts as the ANSI code page (GBK on Chinese Windows),
    # which breaks any non-ASCII string literal.
    local bad=0 f
    while IFS= read -r -d '' f; do
        if [[ $(head -c3 -- "$f" | od -An -tx1 | tr -d ' \n') != efbbbf ]] && LC_ALL=C grep -q $'[\x80-\xff]' -- "$f"; then
            echo "::error file=$f::含非 ASCII 字符但缺少 UTF-8 BOM，Windows PowerShell 5.1 会按 GBK 解析。"
            bad=1
        fi
    done < <(git ls-files -z -- '*.ps1' '*.psm1' '*.psd1')
    return "$bad"
}

case "${1:-}" in
    check)
        step 'PowerShell scripts are UTF-8 with BOM'
        check_powershell_encoding
        step 'WPS add-in scripts parse'
        for f in addin-wps/js/*.js addin-wps/test/*.js; do node --check "$f"; done
        step 'WPS update tests'
        node addin-wps/test/updates.js
        step 'Clippy'
        cargo clippy --locked --workspace --all-targets -- -D warnings
        step 'Tests'
        cargo test --locked --workspace
        step 'Linux package install, upgrade and uninstall'
        cargo build --locked -p wordgit
        package_dir=$(mktemp -d)
        trap 'rm -rf -- "$package_dir"' EXIT
        bash scripts/package-linux.sh "$target/debug/wordgit" "$package_dir"
        bash scripts/test-linux-package.sh "$package_dir"/WordGit-*-linux-x86_64.tar.gz
        ;;
    package)
        if [[ $(uname -m) != x86_64 ]]; then
            echo 'The static Linux release is built on x86_64 hosts only.' >&2
            exit 1
        fi
        cargo build --locked --release --target x86_64-unknown-linux-musl -p wordgit
        bash scripts/package-linux.sh "$target/x86_64-unknown-linux-musl/release/wordgit" dist
        ;;
    *)
        echo "usage: $0 check|package" >&2
        exit 2
        ;;
esac
