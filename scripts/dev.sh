#!/usr/bin/env bash
# Run WordGit's Linux development tasks in the container defined by docker/dev/Dockerfile.
# Windows users run scripts/dev.ps1 instead; the commands are the same.
#   scripts/dev.sh check       encoding, WPS add-in, clippy, tests and Linux package checks (what CI runs)
#   scripts/dev.sh package     static x86_64 Linux release package into dist/
#   scripts/dev.sh shell       interactive shell in the container (the default)
#   scripts/dev.sh image       only build the image
#   scripts/dev.sh <command>   any command, e.g. scripts/dev.sh cargo test -p wordgit-core
# Optional download mirrors: WORDGIT_DOCKER_REGISTRY, WORDGIT_APT_MIRROR, WORDGIT_NODE_MIRROR
# (see docs/DEVELOPMENT.md).
set -euo pipefail
root=$(cd "$(dirname "$0")/.." && pwd)
cd "$root"

inputs=(docker/dev/Dockerfile rust-toolchain.toml .node-version .python-version)
rust=$(sed -n 's/^channel *= *"\(.*\)".*/\1/p' rust-toolchain.toml | tr -d '\r')
node=$(tr -d ' \r\n' < .node-version)
python=$(tr -d ' \r\n' < .python-version)
if command -v sha256sum >/dev/null; then hash=sha256sum; else hash='shasum -a 256'; fi
tag=$(cat "${inputs[@]}" | tr -d '\r' | $hash | cut -c1-12)
image="wordgit-dev:$tag"

if [[ -z $(docker images -q "$image") ]]; then
    build=(build -f docker/dev/Dockerfile -t "$image"
        --build-arg "RUST_VERSION=$rust" --build-arg "NODE_VERSION=$node" --build-arg "PYTHON_VERSION=$python")
    if [[ -n ${WORDGIT_DOCKER_REGISTRY:-} ]]; then build+=(--build-arg "REGISTRY=$WORDGIT_DOCKER_REGISTRY"); fi
    if [[ -n ${WORDGIT_APT_MIRROR:-} ]]; then build+=(--build-arg "APT_MIRROR=$WORDGIT_APT_MIRROR"); fi
    if [[ -n ${WORDGIT_NODE_MIRROR:-} ]]; then build+=(--build-arg "NODE_MIRROR=$WORDGIT_NODE_MIRROR"); fi
    docker "${build[@]}" .
    # Images from older version files are no longer used.
    docker images wordgit-dev --format '{{.Tag}}' | while read -r old; do
        if [[ $old != "$tag" ]]; then docker rmi "wordgit-dev:$old" >/dev/null 2>&1 || true; fi
    done
fi

command=("$@")
case "${1:-shell}" in
    image) exit 0 ;;
    check | package) command=(bash scripts/linux-tasks.sh "$@") ;;
    shell) command=(bash) ;;
esac

run=(run --rm --init -v "$root:/work" -v wordgit-dev-cache:/cache -w /work)
if [[ -t 0 && -t 1 ]]; then run+=(-it); fi
# On Linux the bind mount keeps numeric owners, so run as the host user to keep dist/ writable.
if [[ $(uname -s) == Linux ]]; then run+=(--user "$(id -u):$(id -g)"); fi
exec docker "${run[@]}" "$image" "${command[@]}"
