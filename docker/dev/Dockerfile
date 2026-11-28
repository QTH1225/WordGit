# syntax=docker/dockerfile:1
# WordGit Linux development environment, shared by local work and the Linux CI jobs.
# Do not build or run this file directly: use scripts/dev.ps1 (Windows) or scripts/dev.sh (Linux/macOS).
# Tool versions are build arguments read from rust-toolchain.toml, .node-version and .python-version,
# so the repository's version files stay the only place to change them.

ARG REGISTRY=
ARG RUST_VERSION
FROM ${REGISTRY}rust:${RUST_VERSION}-slim-trixie

ARG APT_MIRROR=
ARG NODE_VERSION
ARG NODE_MIRROR=https://nodejs.org/dist
ARG PYTHON_VERSION

# git: the kernel's real-git integration tests. curl/python3: the Linux package install tests.
# musl-tools: the static x86_64 release build.
RUN set -eux; \
    if [ -n "$APT_MIRROR" ]; then sed -i "s|http://deb.debian.org|$APT_MIRROR|g" /etc/apt/sources.list.d/debian.sources; fi; \
    apt-get update; \
    apt-get install -y --no-install-recommends ca-certificates curl git python3 xz-utils; \
    if [ "$(dpkg --print-architecture)" = amd64 ]; then apt-get install -y --no-install-recommends musl-tools; fi; \
    rm -rf /var/lib/apt/lists/*; \
    actual=$(python3 -c 'import sys; print("%d.%d" % sys.version_info[:2])'); \
    if [ "$actual" != "$PYTHON_VERSION" ]; then echo "Python $actual in the base image, .python-version wants $PYTHON_VERSION" >&2; exit 1; fi

RUN set -eux; \
    case "$(dpkg --print-architecture)" in amd64) arch=x64 ;; arm64) arch=arm64 ;; *) echo "unsupported architecture" >&2; exit 1 ;; esac; \
    file="node-v$NODE_VERSION-linux-$arch.tar.xz"; \
    cd /tmp; \
    curl -fsSLO "$NODE_MIRROR/v$NODE_VERSION/$file"; \
    curl -fsSL "$NODE_MIRROR/v$NODE_VERSION/SHASUMS256.txt" | grep " $file\$" | sha256sum -c -; \
    tar -xJf "$file" -C /usr/local --strip-components=1 --no-same-owner; \
    rm "$file"; \
    test "$(node --version)" = "v$NODE_VERSION"

# Install exactly the channel and components listed in rust-toolchain.toml.
RUN --mount=type=bind,source=rust-toolchain.toml,target=/tmp/toolchain/rust-toolchain.toml set -eux; \
    cd /tmp/toolchain; \
    rustup install; \
    if [ "$(dpkg --print-architecture)" = amd64 ]; then rustup target add x86_64-unknown-linux-musl; fi; \
    rustc --version; cargo clippy --version; \
    chmod -R a+rwX "$RUSTUP_HOME" "$CARGO_HOME"

# Build output, the crate registry and HOME live in one named volume (/cache) so they survive
# between runs and stay off the bind-mounted source tree. The scripts may run as the host user.
ENV CARGO_HOME=/cache/cargo \
    CARGO_TARGET_DIR=/cache/target \
    HOME=/cache/home \
    GIT_CONFIG_COUNT=1 \
    GIT_CONFIG_KEY_0=safe.directory \
    GIT_CONFIG_VALUE_0=*
RUN mkdir -p /cache/cargo /cache/target /cache/home && chmod -R 1777 /cache
WORKDIR /work
