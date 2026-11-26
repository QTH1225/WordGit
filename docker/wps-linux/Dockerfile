# syntax=docker/dockerfile:1
# WordGit on Linux WPS (personal edition): the WordGit service and add-in, plus the P0 probe.
# Build from the repository root:
#   docker compose -f docker/wps-linux/compose.yaml build
# The WPS package must already be in docker/wps-linux/cache/ (see docker/wps-linux/README.md).

FROM rust:1.96.1-bookworm AS rust-build
WORKDIR /src
COPY Cargo.toml Cargo.lock ./
COPY crates crates
COPY poc/probe-server poc/probe-server
RUN --mount=type=cache,target=/usr/local/cargo/registry \
    --mount=type=cache,target=/src/target \
    cargo build --release --locked -p wordgit -p wordgit-probe \
 && mkdir -p /out && cp target/release/wordgit target/release/wordgit-probe /out/

FROM ubuntu:24.04
ARG WPS_DEB=wps-office_12.1.2.29183_amd64.deb
ARG WPS_SHA256=77ffbdc1be76c834e1853c27ff9063e49d5d3b1f30b206d785761b5bf560c83b
ENV DEBIAN_FRONTEND=noninteractive \
    LANG=zh_CN.UTF-8 \
    LC_ALL=zh_CN.UTF-8 \
    TZ=Asia/Shanghai

# The last two package lines are CEF runtime libraries. WPS embeds CEF for JS add-ins and its own
# web pages but its package does not declare them; without them add-ins are trusted yet never run.
RUN sed -i -e 's|http://archive.ubuntu.com/ubuntu|http://mirrors.huaweicloud.com/ubuntu|g' \
           -e 's|http://security.ubuntu.com/ubuntu|http://mirrors.huaweicloud.com/ubuntu|g' \
           /etc/apt/sources.list.d/ubuntu.sources \
 && apt-get update \
 && apt-get install -y --no-install-recommends \
        ca-certificates locales tzdata dbus-x11 \
        xvfb x11vnc fluxbox novnc websockify xdotool x11-utils imagemagick \
        fonts-noto-cjk fonts-dejavu-core python3 procps \
        bsdextrautils xdg-utils desktop-file-utils \
        libnspr4 libnss3 libatk1.0-0t64 libatk-bridge2.0-0t64 libasound2t64 libatspi2.0-0t64 \
        libgbm1 libxcomposite1 libxdamage1 libxrandr2 libcups2t64 libxkbcommon0 libdrm2 \
 && sed -i 's/^# *zh_CN.UTF-8/zh_CN.UTF-8/' /etc/locale.gen && locale-gen \
 && rm -rf /var/lib/apt/lists/*

RUN --mount=type=bind,source=docker/wps-linux/cache,target=/cache \
    echo "${WPS_SHA256}  /cache/${WPS_DEB}" | sha256sum -c - \
 && apt-get update \
 && apt-get install -y --no-install-recommends /cache/${WPS_DEB} \
 && rm -rf /var/lib/apt/lists/*

# WordGit's runtime dependency (git) and curl for the test scripts.
RUN apt-get update \
 && apt-get install -y --no-install-recommends git curl openssh-client \
 && rm -rf /var/lib/apt/lists/*

COPY --from=rust-build /out/wordgit /out/wordgit-probe /usr/local/bin/
COPY poc/wps-addin /opt/wordgit/addin
COPY addin-wps /opt/wordgit/addin-wps-src
RUN cp -r /opt/wordgit/addin-wps-src /opt/wordgit/addin-wps \
 && rm -rf /opt/wordgit/addin-wps/test \
 && mkdir -p /opt/wordgit/addin-wps-test \
 && cp /opt/wordgit/addin-wps-src/test/e2e.js /opt/wordgit/addin-wps-test/ \
 && rm -rf /opt/wordgit/addin-wps-src
COPY docker/wps-linux/entrypoint.sh docker/wps-linux/probe-run.sh docker/wps-linux/e2e-run.sh docker/wps-linux/make-docx.py /opt/wordgit/
RUN chmod 755 /opt/wordgit/entrypoint.sh /opt/wordgit/probe-run.sh /opt/wordgit/e2e-run.sh

USER ubuntu
WORKDIR /home/ubuntu
EXPOSE 6080
ENTRYPOINT ["/opt/wordgit/entrypoint.sh"]
