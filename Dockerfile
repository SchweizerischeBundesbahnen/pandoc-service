FROM python:3.13.3-alpine@sha256:452682e4648deafe431ad2f2391d726d7c52f0ff291be8bd4074b10379bb89ff
LABEL maintainer="SBB Polarion Team <polarion-opensource@sbb.ch>"

ARG APP_IMAGE_VERSION=0.0.0
ARG PANDOC_VERSION=3.7

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
    && wget -q https://github.com/jgm/pandoc/releases/download/${PANDOC_VERSION}/pandoc-${PANDOC_VERSION}-linux-amd64.tar.gz -O /tmp/pandoc.tar.gz \
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

RUN BUILD_TIMESTAMP="$(date -u +"%Y-%m-%dT%H:%M:%SZ")" && \
    echo "${BUILD_TIMESTAMP}" > "${WORKING_DIR}/.build_timestamp"

COPY requirements.txt "${WORKING_DIR}/requirements.txt"

COPY ./app/*.py "${WORKING_DIR}/app/"
COPY ./pyproject.toml ${WORKING_DIR}/pyproject.toml
COPY ./poetry.lock ${WORKING_DIR}/poetry.lock

RUN pip3 install --no-cache-dir -r "${WORKING_DIR}/requirements.txt" && poetry install --no-root --only main

COPY entrypoint.sh "${WORKING_DIR}/entrypoint.sh"
RUN chmod +x "${WORKING_DIR}/entrypoint.sh"

# Use Tini as entrypoint with security options
ENTRYPOINT ["./entrypoint.sh"]
