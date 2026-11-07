# Copy uv from official image (version matches weasyprint-service)
FROM ghcr.io/astral-sh/uv:0.13.7@sha256:96f40d3904aad8d0bd010539001d36c8e2188b981e1334471a90c4a1d93bb490 AS uv-source

# Use debian:trixie-slim as base (same base as weasyprint-service).
FROM debian:trixie-slim@sha256:a72b4433de1cf7d6bae0c4cee74d0547f58601a643b3375f070c51600d63b6b2
LABEL maintainer="SBB Polarion Team <polarion-opensource@sbb.ch>"

# Copy uv binary from source stage
COPY --from=uv-source /uv /usr/local/bin/uv

ARG APP_IMAGE_VERSION=0.0.0
ARG PANDOC_VERSION=3.13
ARG TECTONIC_VERSION=0.16.9
ARG TARGETARCH
ENV ARCH=${TARGETARCH:-amd64}

# Install system dependencies:
# - Pango and fonts, which WeasyPrint draws the SVG of a document with (weasysvg)
# - poppler-utils provides pdftocairo, which weasysvg-convert flattens SVG for Word with;
#   libharfbuzz-subset0 lets WeasyPrint subset fonts fast
# - curl / ca-certificates to fetch the pandoc and tectonic binaries
# hadolint ignore=DL3008
RUN apt-get update && \
    apt-get upgrade -y && \
    apt-get --yes --no-install-recommends install \
    curl \
    ca-certificates \
    fonts-crosextra-carlito \
    fonts-dejavu \
    fonts-liberation \
    fonts-noto-cjk \
    fonts-noto-cjk-extra \
    fonts-noto-color-emoji \
    libharfbuzz-subset0 \
    libpango-1.0-0 \
    libpangoft2-1.0-0 \
    poppler-utils && \
    apt-get clean autoclean && \
    apt-get --yes autoremove && \
    rm -rf /var/lib/apt/lists/*

# Install pandoc (downloaded binary) and the upstream pagebreak Lua filter
RUN curl -fsSL "https://github.com/jgm/pandoc/releases/download/${PANDOC_VERSION}/pandoc-${PANDOC_VERSION}-linux-${ARCH}.tar.gz" -o /tmp/pandoc.tar.gz && \
    tar -xzf /tmp/pandoc.tar.gz -C /tmp && \
    mv "/tmp/pandoc-${PANDOC_VERSION}/bin/pandoc" /usr/local/bin/ && \
    mkdir -p /usr/local/share/pandoc/filters/common/ && \
    curl -fsSL https://raw.githubusercontent.com/pandoc/lua-filters/master/pagebreak/pagebreak.lua -o /usr/local/share/pandoc/filters/common/pagebreak.lua && \
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

# Warm the tectonic bundle into the image. Tectonic fetches the TeX support files it needs on its first
# run and caches them, some 45 MB. Without this step that fetch happens on the first pdf conversion of a
# fresh container, which makes that conversion depend on the network and fail intermittently.
#
# TECTONIC_CACHE_DIR pins the cache to a fixed path. The default is under $HOME, so a container started
# with --user would look elsewhere and download the bundle again.
ENV TECTONIC_CACHE_DIR=/opt/tectonic-cache

# The warm-up document carries one instance of every construct the service converts, because tectonic
# fetches files on demand: pandoc's latex template requests graphicx, longtable, booktabs, array, ulem,
# fancyvrb, geometry and the math fonts only for the matching content, the docx filters add colortbl and
# soul through header-includes, and each heading size and each highlighting style needs its own font file.
#
# Each file is a separate request, and the bundle host answers 429 to a build machine that asks too
# fast. Tectonic gives up after three retries roughly half a second apart, which such a limit outlives,
# so the compile is repeated with a growing pause. Every attempt keeps what it already cached, which
# walks the cache to completion. A build that never gets there fails on the missing pdf, on purpose.
COPY warmup/ /tmp/warmup/
RUN for delay in 0 20 40 60 90 120; do \
        [ "${delay}" -eq 0 ] || { echo "warm-up failed, retrying in ${delay}s"; sleep "${delay}"; }; \
        pandoc /tmp/warmup/warmup.md -o /tmp/warmup.pdf --pdf-engine=tectonic --resource-path=/tmp/warmup && break; \
    done; \
    test -f /tmp/warmup.pdf && \
    chmod -R a+rX "${TECTONIC_CACHE_DIR}" && \
    rm -rf /tmp/warmup /tmp/warmup.pdf

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

# Install dependencies
RUN --mount=type=cache,target=/root/.cache/uv \
    uv sync --frozen --no-dev

# Put the venv on PATH (so python/pandoc tooling resolves without activation) and
# set PYTHONPATH to the working dir.
ENV PATH="${WORKING_DIR}/.venv/bin:$PATH" \
    PYTHONPATH=${WORKING_DIR}

# Open fonts as wide and as high as the fonts of Word they stand for, which app/docx_document_styles.py maps
# the fonts of Word to: fonts/metric_fonts.py makes them from the outlines of open fonts and the metrics of Word.
# The open fonts and their licenses come from fixed commits, checked against fonts/sources.sha256.
ARG GOOGLE_FONTS_COMMIT=86aced91de05cd4f96a06838124cca45be337b2b
ARG SELAWIK_COMMIT=89362e84731d1f5777fa078fd4d5ebbd339e4378
ARG SELAWIK_VERSION=1.01
COPY fonts/metric_fonts.py fonts/word_metrics.json fonts/sources.sha256 /tmp/fonts/
RUN google="https://raw.githubusercontent.com/google/fonts/${GOOGLE_FONTS_COMMIT}/ofl" && \
    sources=/tmp/fonts/sources && \
    mkdir -p "${sources}" && \
    curl -fsSL -o "${sources}/Archivo[wdth,wght].ttf" "${google}/archivo/Archivo%5Bwdth,wght%5D.ttf" && \
    curl -fsSL -o "${sources}/Archivo-Italic[wdth,wght].ttf" "${google}/archivo/Archivo-Italic%5Bwdth,wght%5D.ttf" && \
    curl -fsSL -o "${sources}/Archivo-OFL.txt" "${google}/archivo/OFL.txt" && \
    curl -fsSL -o "${sources}/Gelasio[wght].ttf" "${google}/gelasio/Gelasio%5Bwght%5D.ttf" && \
    curl -fsSL -o "${sources}/Gelasio-Italic[wght].ttf" "${google}/gelasio/Gelasio-Italic%5Bwght%5D.ttf" && \
    curl -fsSL -o "${sources}/Gelasio-OFL.txt" "${google}/gelasio/OFL.txt" && \
    curl -fsSL -o "${sources}/Selawik_Release.zip" "https://github.com/microsoft/Selawik/releases/download/${SELAWIK_VERSION}/Selawik_Release.zip" && \
    curl -fsSL -o "${sources}/Selawik-LICENSE.txt" "https://raw.githubusercontent.com/microsoft/Selawik/${SELAWIK_COMMIT}/LICENSE.txt" && \
    sed "s|  |  ${sources}/|" /tmp/fonts/sources.sha256 > /tmp/fonts/sources.paths && \
    sha256sum -c /tmp/fonts/sources.paths && \
    python /tmp/fonts/metric_fonts.py build "${sources}" /tmp/fonts/word_metrics.json /usr/local/share/fonts/metric && \
    cp "${sources}"/*.txt /usr/local/share/fonts/metric/ && \
    fc-cache -f && \
    rm -rf /tmp/fonts

# rsvg-convert is weasysvg-convert, in a directory of its own, which pandoc gets first in PATH (app/svg_renderer.py)
RUN mkdir -p /opt/weasysvg/bin && \
    ln -s "${WORKING_DIR}/.venv/bin/weasysvg-convert" /opt/weasysvg/bin/rsvg-convert

RUN BUILD_TIMESTAMP="$(date -u +"%Y-%m-%dT%H:%M:%SZ")" && \
    echo "${BUILD_TIMESTAMP}" > "${WORKING_DIR}/.build_timestamp"

COPY ./app/*.py "${WORKING_DIR}/app/"

COPY entrypoint.sh "${WORKING_DIR}/entrypoint.sh"
COPY healthcheck.sh "${WORKING_DIR}/healthcheck.sh"
RUN chmod +x "${WORKING_DIR}/entrypoint.sh" "${WORKING_DIR}/healthcheck.sh"

# The filters by the format they read, and those for any format
COPY filters/html/ "/usr/local/share/pandoc/filters/html/"
COPY filters/docx/ "/usr/local/share/pandoc/filters/docx/"
COPY filters/common/ "/usr/local/share/pandoc/filters/common/"

HEALTHCHECK --interval=30s --timeout=5s --start-period=10s --retries=3 \
  CMD ["/bin/sh", "-c", "./healthcheck.sh"]

ENTRYPOINT ["./entrypoint.sh"]
