FROM pandoc/extra:3.6.4-ubuntu@sha256:858fbccef809a01d2db30f03cb881d22a223907da303e4d66023c2b549f4f99a
LABEL maintainer="SBB Polarion Team <polarion-opensource@sbb.ch>"

ARG APP_IMAGE_VERSION=0.0.0-dev

ENV WORKING_DIR="/opt/pandoc"
ENV PANDOC_SERVICE_VERSION="${APP_IMAGE_VERSION}"

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
