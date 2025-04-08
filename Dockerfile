FROM pandoc/minimal:3.6.4-alpine@sha256:6de776089b7204840084cd9f6267a96162742f45db493b343bf8464f10044810
LABEL maintainer="SBB Polarion Team <polarion-opensource@sbb.ch>"

ARG APP_IMAGE_VERSION=0.0.0

RUN apk add --no-cache  \
    python3=~3.12  \
    py3-pip=~24.3  \
    bash=~5.2 &&  \
    mkdir -p /usr/local/share/pandoc/filters/ &&  \
    wget -q https://raw.githubusercontent.com/pandoc/lua-filters/master/pagebreak/pagebreak.lua -O /usr/local/share/pandoc/filters/pagebreak.lua

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

RUN pip3 install --no-cache-dir --break-system-packages -r "${WORKING_DIR}/requirements.txt" && poetry install --no-root --only main

COPY entrypoint.sh "${WORKING_DIR}/entrypoint.sh"
RUN chmod +x "${WORKING_DIR}/entrypoint.sh"

ENTRYPOINT [ "./entrypoint.sh" ]
