#!/bin/bash

BUILD_TIMESTAMP="$(cat /opt/pandoc/.build_timestamp)"
export PANDOC_SERVICE_BUILD_TIMESTAMP=${BUILD_TIMESTAMP}

# Also possible with:
# uv run python -m app.PandocServiceApplication &
# But this will re-install some dependencies

source .venv/bin/activate
python -m app.PandocServiceApplication &

wait -n

exit $?
