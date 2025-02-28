FROM python:3.13.2-slim@sha256:f3614d98f38b0525d670f287b0474385952e28eb43016655dd003d0e28cf8652
LABEL maintainer="SBB Polarion Team <polarion-opensource@sbb.ch>"

ARG APP_IMAGE_VERSION=0.0.0-dev
WORKDIR ${WORKING_DIR}

COPY requirements.txt ${WORKING_DIR}/requirements.txt
COPY ./app/ ${WORKING_DIR}/app/
COPY ./poetry.lock ${WORKING_DIR}
COPY ./pyproject.toml ${WORKING_DIR}

RUN pip install --no-cache-dir -r "${WORKING_DIR}"/requirements.txt && poetry install --no-root

ENTRYPOINT [ "poetry", "run", "python", "-m", "app.requirements_inspector_service" ]
