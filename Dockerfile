FROM ghcr.io/astral-sh/uv:alpine3.23@sha256:38729bf4f24e6d8c4ad454f8162a1dadeb1dbf68348622e5b4ae9ebd861debc5
LABEL maintainer="SBB Polarion Team <polarion-opensource@sbb.ch>"

ARG APP_IMAGE_VERSION=0.0.0
ARG PANDOC_VERSION=3.9
ARG TARGETARCH
ENV ARCH=${TARGETARCH:-amd64}

# Install pandoc and other dependencies
# hadolint ignore=DL3018
RUN apk add --no-cache \
    bash \
    wget \
    ca-certificates \
    tar \
    gzip \
    lua \
    tectonic \
    && wget -q https://github.com/jgm/pandoc/releases/download/${PANDOC_VERSION}/pandoc-${PANDOC_VERSION}-linux-${ARCH}.tar.gz -O /tmp/pandoc.tar.gz \
    && tar -xzf /tmp/pandoc.tar.gz -C /tmp \
    && mv /tmp/pandoc-${PANDOC_VERSION}/bin/pandoc /usr/local/bin/ \
    && mkdir -p /usr/local/share/pandoc/filters/ \
    && wget -q https://raw.githubusercontent.com/pandoc/lua-filters/master/pagebreak/pagebreak.lua -O /usr/local/share/pandoc/filters/pagebreak.lua \
    && rm -rf /tmp/pandoc* \
    && apk del wget tar gzip

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
SHELL ["/bin/bash", "-o", "pipefail", "-c"]
RUN PYTHON_VERSION=$(awk '/^python / {print $2}' .tool-versions) && \
    uv python install "${PYTHON_VERSION}"

# Install dependencies with cache mount
RUN --mount=type=cache,target=/root/.cache/uv \
    uv sync --frozen --no-dev

RUN BUILD_TIMESTAMP="$(date -u +"%Y-%m-%dT%H:%M:%SZ")" && \
    echo "${BUILD_TIMESTAMP}" > "${WORKING_DIR}/.build_timestamp"

COPY ./app/*.py "${WORKING_DIR}/app/"

COPY entrypoint.sh "${WORKING_DIR}/entrypoint.sh"
RUN chmod +x "${WORKING_DIR}/entrypoint.sh"

COPY page_orientation.lua "/usr/local/share/pandoc/filters/page_orientation.lua"

# Use Tini as entrypoint with security options
ENTRYPOINT ["./entrypoint.sh"]
