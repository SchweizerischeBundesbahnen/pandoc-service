# Copy uv from official image (version matches weasyprint-service)
FROM ghcr.io/astral-sh/uv:0.11.26@sha256:3d868e555f8f1dbc324afa005066cd11e1053fc4743b9808ca8025283e65efa5 AS uv-source

# Use debian:trixie-slim as base (same base as weasyprint-service / python:3.14-slim).
# A glibc base is required because Playwright publishes no musllinux wheel, so the
# Chromium used to rasterize embedded SVGs comes from Playwright's bundled browser.
FROM debian:trixie-slim@sha256:28de0877c2189802884ccd20f15ee41c203573bd87bb6b883f5f46362d24c5c2
LABEL maintainer="SBB Polarion Team <polarion-opensource@sbb.ch>"

# Copy uv binary from source stage
COPY --from=uv-source /uv /usr/local/bin/uv

ARG APP_IMAGE_VERSION=0.0.0
ARG PANDOC_VERSION=3.10
ARG TECTONIC_VERSION=0.16.9
ARG TARGETARCH
ENV ARCH=${TARGETARCH:-amd64}

# Install system dependencies:
# - Chromium runtime libraries + fonts (for Playwright's bundled Chromium, used
#   to rasterize embedded SVGs to PNG so renderers without full SVG support get a
#   usable image instead of a fallback warning)
# - procps provides pgrep, used for psutil-based Chromium process monitoring
# - curl / ca-certificates to fetch the pandoc and tectonic binaries
# hadolint ignore=DL3008
RUN apt-get update && \
    apt-get upgrade -y && \
    apt-get --yes --no-install-recommends install \
    curl \
    ca-certificates \
    fonts-dejavu \
    fonts-liberation \
    fonts-noto-cjk \
    fonts-noto-cjk-extra \
    fonts-noto-color-emoji \
    libasound2 \
    libatk-bridge2.0-0 \
    libatk1.0-0 \
    libcups2 \
    libdrm2 \
    libgbm1 \
    libnspr4 \
    libnss3 \
    libpango-1.0-0 \
    libpangoft2-1.0-0 \
    libxcomposite1 \
    libxdamage1 \
    libxfixes3 \
    libxkbcommon0 \
    libxrandr2 \
    procps && \
    apt-get clean autoclean && \
    apt-get --yes autoremove && \
    rm -rf /var/lib/apt/lists/*

# Install pandoc (downloaded binary) and the upstream pagebreak Lua filter
RUN curl -fsSL "https://github.com/jgm/pandoc/releases/download/${PANDOC_VERSION}/pandoc-${PANDOC_VERSION}-linux-${ARCH}.tar.gz" -o /tmp/pandoc.tar.gz && \
    tar -xzf /tmp/pandoc.tar.gz -C /tmp && \
    mv "/tmp/pandoc-${PANDOC_VERSION}/bin/pandoc" /usr/local/bin/ && \
    mkdir -p /usr/local/share/pandoc/filters/ && \
    curl -fsSL https://raw.githubusercontent.com/pandoc/lua-filters/master/pagebreak/pagebreak.lua -o /usr/local/share/pandoc/filters/pagebreak.lua && \
    rm -rf /tmp/pandoc*

# Install tectonic (downloaded static binary) for the PDF/LaTeX target.
# Installed at /usr/bin/tectonic to match the path the health check probes.
# The musl static build is fetched directly (instead of the drop-sh installer)
# because upstream ships only an aarch64-unknown-linux-musl build for arm64 —
# there is no aarch64-unknown-linux-gnu asset, so the glibc auto-detecting
# installer 404s on arm64. The musl binary is fully static and runs on Debian.
RUN case "${ARCH}" in \
        amd64) TECTONIC_ARCH=x86_64 ;; \
        arm64) TECTONIC_ARCH=aarch64 ;; \
        *) echo "Unsupported architecture: ${ARCH}" >&2; exit 1 ;; \
    esac && \
    curl --proto '=https' --tlsv1.2 -fsSL \
        "https://github.com/tectonic-typesetting/tectonic/releases/download/tectonic%40${TECTONIC_VERSION}/tectonic-${TECTONIC_VERSION}-${TECTONIC_ARCH}-unknown-linux-musl.tar.gz" \
        -o /tmp/tectonic.tar.gz && \
    tar -xzf /tmp/tectonic.tar.gz -C /usr/bin tectonic && \
    rm -f /tmp/tectonic.tar.gz

ENV WORKING_DIR="/opt/pandoc"
ENV PANDOC_SERVICE_VERSION="${APP_IMAGE_VERSION}"

# Create and configure logging directory
RUN mkdir -p ${WORKING_DIR}/logs && \
    chmod 777 ${WORKING_DIR}/logs

WORKDIR "${WORKING_DIR}"

# Copy Python version file and dependency files
COPY .tool-versions pyproject.toml uv.lock ./

# Install Python via uv to /opt/python (version from .tool-versions file)
ENV UV_PYTHON_INSTALL_DIR=/opt/python
RUN PYTHON_VERSION=$(awk '/^python / {print $2}' .tool-versions) && \
    uv python install "${PYTHON_VERSION}"

# Install dependencies and Playwright's bundled Chromium (+ its OS deps).
# PLAYWRIGHT_BROWSERS_PATH keeps the browser in a fixed, predictable location.
ENV PLAYWRIGHT_BROWSERS_PATH=/opt/playwright
RUN --mount=type=cache,target=/root/.cache/uv \
    uv sync --frozen --no-dev && \
    uv run playwright install chromium --with-deps

# Put the venv on PATH (so python/pandoc tooling resolves without activation) and
# set PYTHONPATH to the working dir.
ENV PATH="${WORKING_DIR}/.venv/bin:$PATH" \
    PYTHONPATH=${WORKING_DIR}

RUN BUILD_TIMESTAMP="$(date -u +"%Y-%m-%dT%H:%M:%SZ")" && \
    echo "${BUILD_TIMESTAMP}" > "${WORKING_DIR}/.build_timestamp"

COPY ./app/*.py "${WORKING_DIR}/app/"

COPY entrypoint.sh "${WORKING_DIR}/entrypoint.sh"
RUN chmod +x "${WORKING_DIR}/entrypoint.sh"

COPY filters/page_orientation.lua "/usr/local/share/pandoc/filters/page_orientation.lua"
COPY filters/heading_levels.lua "/usr/local/share/pandoc/filters/heading_levels.lua"
COPY filters/inline_styles.lua "/usr/local/share/pandoc/filters/inline_styles.lua"
COPY filters/docx_text_decorations.lua "/usr/local/share/pandoc/filters/docx_text_decorations.lua"
COPY filters/docx_colors_to_latex.lua "/usr/local/share/pandoc/filters/docx_colors_to_latex.lua"
COPY filters/docx_math_colors_to_latex.lua "/usr/local/share/pandoc/filters/docx_math_colors_to_latex.lua"
COPY filters/docx_paragraphs_to_latex.lua "/usr/local/share/pandoc/filters/docx_paragraphs_to_latex.lua"
COPY filters/docx_lists_to_latex.lua "/usr/local/share/pandoc/filters/docx_lists_to_latex.lua"
COPY filters/docx_tables_to_latex.lua "/usr/local/share/pandoc/filters/docx_tables_to_latex.lua"
COPY filters/html_lists.lua "/usr/local/share/pandoc/filters/html_lists.lua"

HEALTHCHECK --interval=30s --timeout=5s --start-period=10s --retries=3 \
  CMD curl -fsS http://localhost:9082/health || exit 1

ENTRYPOINT ["./entrypoint.sh"]
