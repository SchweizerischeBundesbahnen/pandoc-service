#!/bin/bash

BUILD_TIMESTAMP="$(cat /opt/pandoc/.build_timestamp)"
export PANDOC_SERVICE_BUILD_TIMESTAMP=${BUILD_TIMESTAMP}

# Also possible with:
# uv run python -m app.pandoc_service_application
# But this will re-install some dependencies

source .venv/bin/activate

# exec replaces this shell with the server process, so "docker stop" delivers SIGTERM to it
# instead of to a shell which would not forward the signal. uvicorn then stops accepting
# requests and runs the shutdown hooks which close the Chromium browser and the metrics
# server.
exec python -m app.pandoc_service_application
