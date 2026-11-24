# syntax=docker/dockerfile:1
# WordGit P0: Linux WPS (personal edition) with the probe add-in and probe service.
# Build from the repository root:
#   docker compose -f docker/wps-linux/compose.yaml build
# The WPS package must already be in docker/wps-linux/cache/ (see docker/wps-linux/README.md).

FROM rust:1.96.1-bookworm AS probe-build
WORKDIR /src
COPY Cargo.toml Cargo.lock ./
COPY poc/probe-server poc/probe-server
RUN cargo build --release --locked -p wordgit-probe

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

COPY --from=probe-build /src/target/release/wordgit-probe /usr/local/bin/wordgit-probe
COPY poc/wps-addin /opt/wordgit/addin
COPY docker/wps-linux/entrypoint.sh docker/wps-linux/probe-run.sh docker/wps-linux/make-docx.py /opt/wordgit/
RUN chmod 755 /opt/wordgit/entrypoint.sh /opt/wordgit/probe-run.sh

USER ubuntu
WORKDIR /home/ubuntu
EXPOSE 6080
ENTRYPOINT ["/opt/wordgit/entrypoint.sh"]
